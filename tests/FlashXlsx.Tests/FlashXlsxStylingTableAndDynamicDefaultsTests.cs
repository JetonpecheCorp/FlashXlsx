using FlashXlsx.Abstractions;
using System.IO.Compression;
using System.Xml.Linq;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneRapportStyle
{
     [XlsxDefault(XlsxDefaultKind.NewGuid)]
     public Guid Identifiant { get; set; }

     public string LibelleLong { get; set; } = string.Empty;

     [XlsxDefault(XlsxDefaultKind.Today)]
     public DateOnly DateCreation { get; set; }
}

public class FlashXlsxStylingTableAndDynamicDefaultsTests
{
     [Fact]
     public async Task AutoFitColumns_HeaderStyle_And_ExcelTable_ShouldGenerateValidOpenXml()
     {
          var lignes = new List<LigneRapportStyle>
        {
            new() { Identifiant = Guid.NewGuid(), LibelleLong = "Description très détaillée d'une ligne de rapport", DateCreation = new DateOnly(2026, 11, 7) },
            new() { Identifiant = Guid.NewGuid(), LibelleLong = "Court", DateCreation = new DateOnly(2026, 11, 8) }
        };

          var options = new XlsxWriteOptions<LigneRapportStyle>()
              .WithSheetName("Rapport")
              .WithAutoFitColumns()
              .WithHeaderStyle(bold: true, backgroundHex: "1F4E78", textHex: "FFFFFF")
              .WithExcelTable(tableName: "TableauRapport", styleName: "TableStyleMedium2", enableAutoFilter: true);

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, lignes, options);

          ms.Position = 0;
          using (var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true))
          {
               // 1. Vérification de sheet1.xml (<cols> + style s="4" sur header + <tableParts>)
               var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
               Assert.NotNull(sheetEntry);
               using (var s = sheetEntry!.Open())
               {
                    var doc = XDocument.Load(s);
                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                    var cols = doc.Descendants(ns + "col").ToList();
                    Assert.Equal(3, cols.Count);

                    var headerCells = doc.Descendants(ns + "row").First().Elements(ns + "c").ToList();
                    Assert.All(headerCells, c => Assert.Equal("4", c.Attribute("s")?.Value));

                    var tablePart = doc.Descendants(ns + "tablePart").SingleOrDefault();
                    Assert.NotNull(tablePart);
               }

               // 2. Vérification de xl/tables/table1.xml et de l'autoFilter
               var tableEntry = zip.GetEntry("xl/tables/table1.xml");
               Assert.NotNull(tableEntry);
               using (var ts = tableEntry!.Open())
               {
                    var tDoc = XDocument.Load(ts);
                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                    var tableEl = tDoc.Root!;
                    Assert.Equal("TableauRapport", tableEl.Attribute("name")?.Value);
                    Assert.Equal("A1:C3", tableEl.Attribute("ref")?.Value);

                    var autoFilter = tableEl.Element(ns + "autoFilter");
                    Assert.NotNull(autoFilter);
                    Assert.Equal("A1:C3", autoFilter!.Attribute("ref")?.Value);
               }
          }
     }

     [Fact]
     public async Task DynamicDefaults_NewGuid_And_Today_ShouldPopulateEmptyCellsOnImport()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["Identifiant"] = XlsxValue.Empty,
                ["LibelleLong"] = "Ligne sans ID ni date",
                ["DateCreation"] = XlsxValue.Empty
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<LigneRapportStyle>(ms);

          Assert.Single(imported);
          Assert.NotEqual(Guid.Empty, imported[0].Identifiant);
          Assert.Equal(DateOnly.FromDateTime(DateTime.Today), imported[0].DateCreation);
     }
}