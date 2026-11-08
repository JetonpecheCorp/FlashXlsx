using System.IO.Compression;
using System.Xml.Linq;
using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneTresorerie
{
     [XlsxRequired]
     public string Libelle { get; set; } = string.Empty;

     [XlsxNumberFormat("#,##0.00 €")]
     public decimal Montant { get; set; }

     [XlsxNumberFormat("0.00%")]
     public double Taux { get; set; }
}

public class FlashXlsxFileAndNumberFormatTests
{
     [Fact]
     public async Task FilePathOverloads_ShouldWriteAndReadBack_IncludingWhileFileIsOpened()
     {
          string tempFile = Path.Combine(Path.GetTempPath(), $"flashxlsx_io_{Guid.NewGuid():N}.xlsx");
          try
          {
               var data = new List<LigneTresorerie>
            {
                new() { Libelle = "Abonnement SaaS", Montant = 149.90m, Taux = 0.20 },
                new() { Libelle = "Prestation Conseil", Montant = 2500.00m, Taux = 0.20 }
            };

               await FlashXlsxSerializer.WriteAsync(tempFile, data);
               Assert.True(File.Exists(tempFile));

               await using (var lockStream = new FileStream(tempFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
               {
                    var readAll = await FlashXlsxSerializer.ReadAllAsync<LigneTresorerie>(tempFile);
                    Assert.Equal(2, readAll.Count);
                    Assert.Equal("Abonnement SaaS", readAll[0].Libelle);
                    Assert.Equal(149.90m, readAll[0].Montant);
                    Assert.Equal("Prestation Conseil", readAll[1].Libelle);
                    Assert.Equal(2500.00m, readAll[1].Montant);
               }

               string dictFile = Path.Combine(Path.GetTempPath(), $"flashxlsx_dict_{Guid.NewGuid():N}.xlsx");
               try
               {
                    var dictRows = new List<Dictionary<string, XlsxValue>>
                {
                    new() { ["Produit"] = "Clavier", ["Prix"] = 89.50m },
                    new() { ["Produit"] = "Souris", ["Prix"] = 45.00m }
                };

                    await FlashXlsxSerializer.WriteDictionaryAsync(dictFile, dictRows);

                    var loadedDicts = new List<Dictionary<string, XlsxValue>>();
                    await foreach (var d in FlashXlsxSerializer.ReadDictionaryAsync(dictFile))
                         loadedDicts.Add(d);

                    Assert.Equal(2, loadedDicts.Count);
                    Assert.Equal("Clavier", loadedDicts[0]["Produit"].AsString());
                    Assert.Equal(89.50, loadedDicts[0]["Prix"].AsDouble(), 2);
               }
               finally
               {
                    if (File.Exists(dictFile)) File.Delete(dictFile);
               }
          }
          finally
          {
               if (File.Exists(tempFile)) File.Delete(tempFile);
          }
     }

     [Fact]
     public async Task NumberFormatAttribute_ShouldRegisterNativeNumFmts_AndKeepCellsNumeric()
     {
          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, new[]
          {
            new LigneTresorerie { Libelle = "Facture A", Montant = 1250.50m, Taux = 0.195 }
        });

          ms.Position = 0;
          using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);

          var stylesDoc = XDocument.Load(zip.GetEntry("xl/styles.xml")!.Open());
          XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
          var numFmts = stylesDoc.Descendants(ns + "numFmt").ToList();

          Assert.Equal(2, numFmts.Count);
          Assert.Equal("#,##0.00 €", numFmts[0].Attribute("formatCode")?.Value);
          Assert.Equal("0.00%", numFmts[1].Attribute("formatCode")?.Value);

          var sheetDoc = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
          var cells = sheetDoc.Descendants(ns + "c").ToDictionary(c => c.Attribute("r")!.Value);

          Assert.Equal("5", cells["B2"].Attribute("s")?.Value);
          Assert.Null(cells["B2"].Attribute("t"));
          Assert.Equal("1250.50", cells["B2"].Element(ns + "v")?.Value);

          Assert.Equal("6", cells["C2"].Attribute("s")?.Value);
          Assert.Null(cells["C2"].Attribute("t"));
          Assert.Equal("0.195", cells["C2"].Element(ns + "v")?.Value);
     }

     [Fact]
     public async Task ReadAllSheetsAsync_ShouldMergeMatchingSheets()
     {
          string tempFile = Path.Combine(Path.GetTempPath(), $"flashxlsx_multisheet_{Guid.NewGuid():N}.xlsx");
          try
          {
               await using (var fs = File.Create(tempFile))
               await using (var writer = new XlsxWorkbookWriter(fs))
               {
                    await writer.WriteSheetAsync(
                        [new LigneTresorerie { Libelle = "Jan-1", Montant = 100m, Taux = 0.2 }],
                        new XlsxWriteOptions<LigneTresorerie>().WithSheetName("2026-01"));

                    await writer.WriteSheetAsync(
                        [new LigneTresorerie { Libelle = "Fev-1", Montant = 200m, Taux = 0.2 }],
                        new XlsxWriteOptions<LigneTresorerie>().WithSheetName("2026-02"));

                    await writer.WriteSheetAsync(
                        [new LigneTresorerie { Libelle = "TOTAL", Montant = 300m, Taux = 0.2 }],
                        new XlsxWriteOptions<LigneTresorerie>().WithSheetName("Recap"));
               }

               var merged = new List<LigneTresorerie>();
               await foreach (var item in FlashXlsxSerializer.ReadAllSheetsAsync<LigneTresorerie>(
                   tempFile,
                   sheetFilter: s => s.Name.StartsWith("2026-", StringComparison.Ordinal)))
               {
                    merged.Add(item);
               }

               Assert.Equal(2, merged.Count);
               Assert.Equal("Jan-1", merged[0].Libelle);
               Assert.Equal("Fev-1", merged[1].Libelle);
          }
          finally
          {
               if (File.Exists(tempFile)) File.Delete(tempFile);
          }
     }

     [Fact]
     public async Task ReadAsync_MissingFilePath_ShouldNotifyOnError_AndThrowFileNotFound()
     {
          string missingPath = Path.Combine(Path.GetTempPath(), $"inexistant_{Guid.NewGuid():N}.xlsx");
          var errors = new List<XlsxErrorContext>();

          var opts = new XlsxReadOptions<LigneTresorerie>()
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          await Assert.ThrowsAsync<FileNotFoundException>(async () =>
          {
               await FlashXlsxSerializer.ReadAllAsync(missingPath, opts);
          });

          Assert.Single(errors);
          Assert.Equal(XlsxCellStatus.CorruptedFile, errors[0].Status);
          Assert.Equal(missingPath, errors[0].ColumnName);
     }
}