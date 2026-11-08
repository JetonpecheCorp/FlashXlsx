using FlashXlsx.Abstractions;
using System.IO.Compression;
using System.Text;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneExportErp
{
     public string Reference { get; set; } = string.Empty;
     public decimal Montant { get; set; }
}

public class FlashXlsxAutoDetectHeaderAndProgressTests
{
     [Fact]
     public async Task AutoDetectHeaderRow_ShouldSkipBannerRows_And_FindTableHeader()
     {
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
                   """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Export" sheetId="1" r:id="rId1"/></sheets></workbook>""");
               AddXml(zip, "xl/_rels/workbook.xml.rels",
                   """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");

               // Lignes 1 à 3 = Cartouche ERP, Ligne 4 = Vrais en-têtes, Lignes 5-6 = Données
               AddXml(zip, "xl/worksheets/sheet1.xml",
                   """
                <?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>SOCIÉTÉ ACME - EXPORT COMPTABLE</t></is></c></row>
                    <row r="2"><c r="A2" t="inlineStr"><is><t>Période : Novembre 2026</t></is></c></row>
                    <row r="4">
                      <c r="A4" t="inlineStr"><is><t>Reference</t></is></c>
                      <c r="B4" t="inlineStr"><is><t>Montant</t></is></c>
                    </row>
                    <row r="5">
                      <c r="A5" t="inlineStr"><is><t>FAC-001</t></is></c>
                      <c r="B5"><v>150.50</v></c>
                    </row>
                    <row r="6">
                      <c r="A6" t="inlineStr"><is><t>FAC-002</t></is></c>
                      <c r="B6"><v>320.00</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """);
          }

          ms.Position = 0;
          var readReports = new List<int>();
          var progress = new SyncProgress<int>(count => readReports.Add(count));

          var options = new XlsxReadOptions<LigneExportErp>()
              .WithAutoDetectHeaderRow()
              .WithProgress(progress, intervalRows: 1);

          var rows = await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Equal(2, rows.Count);
          Assert.Equal("FAC-001", rows[0].Reference);
          Assert.Equal(150.50m, rows[0].Montant);
          Assert.Equal("FAC-002", rows[1].Reference);
          Assert.Equal([1, 2], readReports);
     }

     private sealed class SyncProgress<T>(Action<T> onReport) : IProgress<T>
     {
          public void Report(T value) => onReport(value);
     }
}