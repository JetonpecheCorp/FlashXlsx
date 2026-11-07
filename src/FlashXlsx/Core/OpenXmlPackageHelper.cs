using FlashXlsx.Abstractions;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx.Core;

internal sealed record XlsxTableArtifact(
    int TableId,
    int SheetNumber,
    string TableName,
    string RefRange,
    IReadOnlyList<string> Headers,
    XlsxTableOptions Options
);

internal static class OpenXmlPackageHelper
{
     public static async ValueTask WritePackageMetadataAsync(
         ZipArchive archive,
         IReadOnlyList<string> sheetNames,
         IReadOnlyList<XlsxTableArtifact>? tables = null,
         XlsxHeaderStyle? headerStyle = null,
         CancellationToken ct = default)
     {
          tables ??= [];

          // 1. _rels/.rels
          await WriteUtf8EntryAsync(archive, "_rels/.rels",
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
              "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
              "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
              "</Relationships>", ct);

          // 2. xl/styles.xml (Index 0=Standard, 1=Date, 2=DateTime, 3=Time, 4=HeaderStyle)
          string boldXml = (headerStyle?.Bold ?? true) ? "<b/>" : string.Empty;
          string fgRgb = "FF" + (headerStyle?.TextHex ?? "000000").TrimStart('#');
          string bgRgb = "FF" + (headerStyle?.BackgroundHex ?? "F2F2F2").TrimStart('#');

          await WriteUtf8EntryAsync(archive, "xl/styles.xml",
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
              "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
              "<numFmts count=\"3\">" +
              "<numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>" +
              "<numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>" +
              "<numFmt numFmtId=\"166\" formatCode=\"hh:mm:ss\"/>" +
              "</numFmts>" +
              "<fonts count=\"2\">" +
              "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
              $"<font>{boldXml}<sz val=\"11\"/><color rgb=\"{fgRgb}\"/><name val=\"Calibri\"/></font>" +
              "</fonts>" +
              "<fills count=\"3\">" +
              "<fill><patternFill patternType=\"none\"/></fill>" +
              "<fill><patternFill patternType=\"gray125\"/></fill>" +
              $"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{bgRgb}\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
              "</fills>" +
              "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
              "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
              "<cellXfs count=\"5\">" +
              "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
              "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
              "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
              "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
              "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +
              "</cellXfs>" +
              "</styleSheet>", ct);

          // 3. [Content_Types].xml
          var ctSb = new StringBuilder();
          ctSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
          ctSb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
          ctSb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
          ctSb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
          ctSb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
          ctSb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");

          for (int i = 0; i < sheetNames.Count; i++)
          {
               ctSb.Append($"<Override PartName=\"/xl/worksheets/sheet{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
          }

          for (int i = 0; i < tables.Count; i++)
          {
               ctSb.Append($"<Override PartName=\"/xl/tables/table{tables[i].TableId}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
          }

          ctSb.Append("</Types>");
          await WriteUtf8EntryAsync(archive, "[Content_Types].xml", ctSb.ToString(), ct);

          // 4. xl/workbook.xml & xl/_rels/workbook.xml.rels
          var wbSb = new StringBuilder();
          wbSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
          wbSb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");

          var relsSb = new StringBuilder();
          relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
          relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
          relsSb.Append("<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");

          for (int i = 0; i < sheetNames.Count; i++)
          {
               string escapedName = EscapeXmlAttribute(sheetNames[i]);
               wbSb.Append($"<sheet name=\"{escapedName}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
               relsSb.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
          }

          wbSb.Append("</sheets></workbook>");
          relsSb.Append("</Relationships>");

          await WriteUtf8EntryAsync(archive, "xl/workbook.xml", wbSb.ToString(), ct);
          await WriteUtf8EntryAsync(archive, "xl/_rels/workbook.xml.rels", relsSb.ToString(), ct);

          // 5. Véritables tableaux Excel (xl/tables/table{id}.xml + relations de feuille)
          for (int i = 0; i < tables.Count; i++)
          {
               var t = tables[i];
               string escapedTableName = EscapeXmlAttribute(t.TableName);
               string autoFilterXml = t.Options.EnableAutoFilter ? $"<autoFilter ref=\"{t.RefRange}\"/>" : string.Empty;
               string stripes = t.Options.ShowRowStripes ? "1" : "0";

               var tblSb = new StringBuilder();
               tblSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
               tblSb.Append($"<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{t.TableId}\" name=\"{escapedTableName}\" displayName=\"{escapedTableName}\" ref=\"{t.RefRange}\" totalsRowShown=\"0\">");
               tblSb.Append(autoFilterXml);
               tblSb.Append($"<tableColumns count=\"{t.Headers.Count}\">");
               for (int c = 0; c < t.Headers.Count; c++)
               {
                    string colName = EscapeXmlAttribute(t.Headers[c]);
                    tblSb.Append($"<tableColumn id=\"{c + 1}\" name=\"{colName}\"/>");
               }
               tblSb.Append("</tableColumns>");
               tblSb.Append($"<tableStyleInfo name=\"{EscapeXmlAttribute(t.Options.StyleName)}\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"{stripes}\" showColumnStripes=\"0\"/>");
               tblSb.Append("</table>");

               await WriteUtf8EntryAsync(archive, $"xl/tables/table{t.TableId}.xml", tblSb.ToString(), ct);

               string sheetRelsXml =
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                   $"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{t.TableId}.xml\"/>" +
                   "</Relationships>";

               await WriteUtf8EntryAsync(archive, $"xl/worksheets/_rels/sheet{t.SheetNumber}.xml.rels", sheetRelsXml, ct);
          }
     }

     private static async ValueTask WriteUtf8EntryAsync(ZipArchive archive, string entryName, string content, CancellationToken ct)
     {
          var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
          await using var stream = entry.Open();
          byte[] bytes = Encoding.UTF8.GetBytes(content);
          await stream.WriteAsync(bytes, ct);
     }

     private static string EscapeXmlAttribute(string value)
         => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}