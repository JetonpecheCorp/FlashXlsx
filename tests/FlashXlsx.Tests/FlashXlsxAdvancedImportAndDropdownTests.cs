using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

public class FlashXlsxAdvancedImportAndDropdownTests
{
     [Fact]
     public async Task FreezeHeader_And_DropdownValidation_ShouldGenerateValidOpenXml()
     {
          var articles = new List<ArticleCatalogue>
        {
            new()
            {
                Code = "PRD-0001",
                Designation = "Souris Sans Fil",
                Quantite = 10,
                PrixUnitaire = 29.90m,
                Statut = StatutStock.Disponible
            }
        };

          var options = new XlsxWriteOptions<ArticleCatalogue>()
              .WithFreezeHeader()
              .WithDropdownValidation();

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, articles, options);

          ms.Position = 0;
          using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
          var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
          Assert.NotNull(sheetEntry);

          using var s = sheetEntry!.Open();
          var doc = XDocument.Load(s);
          XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

          // 1. Vérification du gel des volets (<pane ySplit="1" ... state="frozen"/>)
          var pane = doc.Descendants(ns + "pane").SingleOrDefault();
          Assert.NotNull(pane);
          Assert.Equal("1", pane!.Attribute("ySplit")?.Value);
          Assert.Equal("frozen", pane.Attribute("state")?.Value);

          // 2. Vérification de la liste déroulante sur la colonne E (Statut)
          var dv = doc.Descendants(ns + "dataValidation").SingleOrDefault();
          Assert.NotNull(dv);
          Assert.Equal("list", dv!.Attribute("type")?.Value);
          Assert.Equal("E2:E2", dv.Attribute("sqref")?.Value);
          Assert.Contains("En stock,Réapprovisionnement,Épuisé", dv.Element(ns + "formula1")?.Value);
     }

     [Fact]
     public async Task Import_ShouldReadCachedFormulaValue_And_RichTextCells()
     {
          // Création d'un fichier XLSX contenant une formule avec valeur en cache <v> et du Rich Text <is><r><t>
          using var ms = new MemoryStream();
          using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
          {
               static void AddXml(ZipArchive z, string path, string xml)
               {
                    var e = z.CreateEntry(path);
                    using var w = new StreamWriter(e.Open(), Encoding.UTF8);
                    w.Write(xml);
               }

               AddXml(zip, "[Content_Types].xml",
                   """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");

               AddXml(zip, "_rels/.rels",
                   """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");

               AddXml(zip, "xl/workbook.xml",
                   """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");

               AddXml(zip, "xl/_rels/workbook.xml.rels",
                   """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");

               AddXml(zip, "xl/worksheets/sheet1.xml",
                   """
                <?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr"><is><t>Code</t></is></c>
                      <c r="B1" t="inlineStr"><is><t>Designation</t></is></c>
                      <c r="C1" t="inlineStr"><is><t>Quantite</t></is></c>
                      <c r="D1" t="inlineStr"><is><t>PrixUnitaire</t></is></c>
                      <c r="E1" t="inlineStr"><is><t>Statut</t></is></c>
                      <c r="F1" t="inlineStr"><is><t>TotalLigne</t></is></c>
                    </row>
                    <row r="2">
                      <c r="A2" t="inlineStr"><is><t>PRD-9999</t></is></c>
                      <c r="B2" t="inlineStr"><is><r><t>Écran </t></r><r><t>4K OLED</t></r></is></c>
                      <c r="C2"><v>2</v></c>
                      <c r="D2"><v>450.00</v></c>
                      <c r="E2" t="inlineStr"><is><t>En stock</t></is></c>
                      <c r="F2"><f>C2*D2</f><v>900.00</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
          }

          ms.Position = 0;
          var items = await FlashXlsxSerializer.ReadAllAsync<ArticleCatalogue>(ms);

          Assert.Single(items);
          Assert.Equal("PRD-9999", items[0].Code);
          Assert.Equal("Écran 4K OLED", items[0].Designation); // Rich Text concaténé
          Assert.Equal(900.00m, items[0].TotalLigne);          // Valeur mise en cache <v> de la formule <f>
     }
}