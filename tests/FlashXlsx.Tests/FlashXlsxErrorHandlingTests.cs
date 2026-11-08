using FlashXlsx;
using FlashXlsx.Abstractions;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneAvecConvertisseurFautif
{
     public string Nom { get; set; } = string.Empty;

     [XlsxConverter(typeof(ConvertisseurExplosif))]
     public int Valeur { get; set; }
}

[XlsxSerializable]
public class LigneFusionAvecValidation
{
     public string Groupe { get; set; } = string.Empty;

     [XlsxRange(1000, 10000)]
     public decimal Salaire { get; set; }
}

public sealed class ConvertisseurExplosif : IXlsxConverter<int>
{
     public bool TryRead(ReadOnlySpan<char> rawValue, XlsxReadOptions options, out int value)
     {
          if (rawValue.Equals("BOOM", StringComparison.OrdinalIgnoreCase))
               throw new InvalidOperationException("Erreur critique de conversion");

          return int.TryParse(rawValue, out value);
     }

     public string? Write(in int value, XlsxWriteOptions options)
     {
          if (value < 0)
               throw new InvalidOperationException("Valeur négative interdite à l'export");

          return value.ToString();
     }
}

public class FlashXlsxErrorHandlingTests
{
     [Fact]
     public async Task Write_WithConverterException_And_ReportAndContinue_ShouldLogErrorAndContinue()
     {
          var items = new List<LigneAvecConvertisseurFautif>
        {
            new() { Nom = "OK1", Valeur = 10 },
            new() { Nom = "ERREUR", Valeur = -5 },
            new() { Nom = "OK2", Valeur = 20 }
        };

          var errors = new List<XlsxErrorContext>();
          var writeOpts = new XlsxWriteOptions<LigneAvecConvertisseurFautif>()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, items, writeOpts);

          Assert.Single(errors);
          Assert.Equal(3, errors[0].RowNumber); // Ligne 1 = Header, Ligne 2 = OK1, Ligne 3 = ERREUR
          Assert.Equal(XlsxCellStatus.ConverterError, errors[0].Status);

          ms.Position = 0;
          var readBack = await FlashXlsxSerializer.ReadAllAsync<LigneAvecConvertisseurFautif>(ms);
          Assert.Equal(3, readBack.Count);
          Assert.Equal(10, readBack[0].Valeur);
          Assert.Equal(0, readBack[1].Valeur); // Cellule laissée vide suite à l'erreur ignorée
          Assert.Equal(20, readBack[2].Valeur);
     }

     [Fact]
     public async Task Read_WithCorruptedZipStream_ShouldThrowDescriptiveInvalidDataException()
     {
          using var ms = new MemoryStream(Encoding.UTF8.GetBytes("Ceci n'est pas un fichier ZIP/XLSX"));

          await Assert.ThrowsAsync<InvalidDataException>(async () =>
          {
               await FlashXlsxSerializer.ReadAllAsync<LigneAvecConvertisseurFautif>(ms);
          });
     }

     [Fact]
     public async Task FillMergedCells_WithInvalidPropagatedValue_ShouldReportErrorPerImpactedRow()
     {
          // Le salaire fusionné en C2:C4 vaut "500" (en dessous du [XlsxRange(1000, 10000)])
          using var ms = new MemoryStream();
          using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
          {
               WriteZipEntry(zip, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """);

               WriteZipEntry(zip, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

               WriteZipEntry(zip, "xl/workbook.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);

               WriteZipEntry(zip, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """);

               WriteZipEntry(zip, "xl/worksheets/sheet1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>Groupe</t></is></c>
                      <c r="B1" t="inlineStr"><is><t>Salaire</t></is></c>
                    </row>
                    <row r="2">
                      <c r="A2" t="inlineStr"><is><t>Equipe A</t></is></c>
                      <c r="B2"><v>500</v></c>
                    </row>
                    <row r="3">
                      <c r="A3" t="inlineStr"><is><t>Equipe B</t></is></c>
                    </row>
                    <row r="4">
                      <c r="A4" t="inlineStr"><is><t>Equipe C</t></is></c>
                    </row>
                  </sheetData>
                  <mergeCells count="1">
                    <mergeCell ref="B2:B4"/>
                  </mergeCells>
                </worksheet>
                """);
          }

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();

          var opts = new XlsxReadOptions<LigneFusionAvecValidation>()
              .WithFillMergedCells()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          var validRows = await FlashXlsxSerializer.ReadAllAsync(ms, opts);

          // Les 3 lignes (2, 3 et 4) reçoivent "500" propagé et échouent sur [XlsxRange(1000, 10000)]
          Assert.Equal(3, validRows.Count);
          Assert.Equal(3, errors.Count);
          Assert.Equal(2, errors[0].RowNumber);
          Assert.Equal(3, errors[1].RowNumber);
          Assert.Equal(4, errors[2].RowNumber);
          Assert.All(errors, e => Assert.Equal(XlsxCellStatus.OutOfRange, e.Status));
     }

     [Fact]
     public async Task CorruptedWorksheetXml_ShouldNotifyOnError_AndThrowInvalidDataException()
     {
          using var ms = new MemoryStream();
          using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
          {
               WriteZipEntry(zip, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                </Types>
                """);

               WriteZipEntry(zip, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);

               WriteZipEntry(zip, "xl/workbook.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Corrupted" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);

               WriteZipEntry(zip, "xl/_rels/workbook.xml.rels", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                </Relationships>
                """);

               // XML volontairement tronqué à la ligne 2
               WriteZipEntry(zip, "xl/worksheets/sheet1.xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>Groupe</t></is></c></row>
                    <row r="2"><c r="A2" t="inlineStr"><is><t>Tronque
                """);
          }

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();

          var opts = new XlsxReadOptions<LigneFusionAvecValidation>()
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          await Assert.ThrowsAsync<InvalidDataException>(async () =>
          {
               await FlashXlsxSerializer.ReadAllAsync(ms, opts);
          });

          Assert.Single(errors);
          Assert.Equal(XlsxCellStatus.CorruptedFile, errors[0].Status);
          Assert.Equal(2, errors[0].RowNumber);
     }

     [Fact]
     public async Task WriteDictionaryAsync_WithWriteError_ShouldReportAndContinue_WhenConfigured()
     {
          // Double.NaN n'est pas valide en XML OpenXML ou peut déclencher une validation;
          // testons avec un chemin de fichier invalide pour vérifier la notification d'erreur I/O en écriture
          string invalidPath = Path.Combine(Path.GetTempPath(), $"dossier_inexistant_{Guid.NewGuid():N}", "export.xlsx");
          var errors = new List<XlsxErrorContext>();

          var opts = new XlsxWriteOptions
          {
               OnError = (in XlsxErrorContext err) => errors.Add(err)
          };

          await Assert.ThrowsAnyAsync<IOException>(async () =>
          {
               await FlashXlsxSerializer.WriteDictionaryAsync(
                   invalidPath,
                   [new Dictionary<string, XlsxValue> { ["Col1"] = "Test" }],
                   opts);
          });

          Assert.Single(errors);
          Assert.Equal(XlsxCellStatus.WriteError, errors[0].Status);
          Assert.Equal(invalidPath, errors[0].ColumnName);
     }

     private static void WriteZipEntry(ZipArchive zip, string path, string xml)
     {
          var entry = zip.CreateEntry(path);
          using var s = entry.Open();
          s.Write(Encoding.UTF8.GetBytes(xml.Trim()));
     }
}