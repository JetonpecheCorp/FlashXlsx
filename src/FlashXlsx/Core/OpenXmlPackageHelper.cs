using FlashXlsx.Abstractions;
using System.IO.Compression;
using System.Text;

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
         IReadOnlyList<string>? customNumberFormats = null,
         CancellationToken ct = default)
     {
          tables ??= [];

          // 1. _rels/.rels
          await WriteUtf8EntryAsync(archive, "_rels/.rels",
              "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
              "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
              "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
              "</Relationships>", ct);

          // 2. xl/styles.xml (Index 0=Standard, 1=Date, 2=DateTime, 3=Time, 4=HeaderStyle, 5+=CustomNumberFormats)
          int customCount = customNumberFormats?.Count ?? 0;
          var sbStyles = new StringBuilder();
          sbStyles.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
          sbStyles.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

          if (customCount > 0)
          {
               sbStyles.Append($"<numFmts count=\"{customCount}\">");
               for (int i = 0; i < customCount; i++)
               {
                    string escapedFmt = EscapeXmlAttribute(customNumberFormats![i]);
                    sbStyles.Append($"<numFmt numFmtId=\"{164 + i}\" formatCode=\"{escapedFmt}\"/>");
               }
               sbStyles.Append("</numFmts>");
          }

          // Fonts (fontId=0 normal, fontId=1 gras pour en-tête)
          bool headerBold = headerStyle?.Bold ?? false;
          string headerColor = NormalizeHexColor(headerStyle?.TextHex, "000000");
          sbStyles.Append("<fonts count=\"2\">");
          sbStyles.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/></font>");
          sbStyles.Append($"<font>{(headerBold ? "<b/>" : "")}<sz val=\"11\"/><color rgb=\"FF{headerColor}\"/><name val=\"Calibri\"/></font>");
          sbStyles.Append("</fonts>");

          // Fills (fillId=0 none, fillId=1 gray125 requis par Excel, fillId=2 fond d'en-tête)
          string bgHex = NormalizeHexColor(headerStyle?.BackgroundHex, "F2F2F2");
          sbStyles.Append("<fills count=\"3\">");
          sbStyles.Append("<fill><patternFill patternType=\"none\"/></fill>");
          sbStyles.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
          sbStyles.Append($"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{bgHex}\"/><bgColor indexed=\"64\"/></patternFill></fill>");
          sbStyles.Append("</fills>");

          // Borders
          sbStyles.Append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
          sbStyles.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

          // CellXfs : 0=Défaut, 1=Date (14), 2=DateTime (22), 3=Time (21), 4=HeaderStyle, 5+=CustomNumberFormats
          sbStyles.Append($"<cellXfs count=\"{5 + customCount}\">");
          sbStyles.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
          sbStyles.Append("<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
          sbStyles.Append("<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
          sbStyles.Append("<xf numFmtId=\"21\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");

          if (headerStyle is not null)
          {
               string fillAttr = headerStyle.BackgroundHex is not null ? " fillId=\"2\" applyFill=\"1\"" : " fillId=\"0\"";
               sbStyles.Append($"<xf numFmtId=\"0\" fontId=\"1\"{fillAttr} borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>");
          }
          else
          {
               sbStyles.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
          }

          for (int i = 0; i < customCount; i++)
          {
               sbStyles.Append($"<xf numFmtId=\"{164 + i}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
          }

          sbStyles.Append("</cellXfs>");
          sbStyles.Append("</styleSheet>");

          await WriteUtf8EntryAsync(archive, "xl/styles.xml", sbStyles.ToString(), ct);

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

     private static string NormalizeHexColor(string? hex, string fallback)
     {
          if (string.IsNullOrWhiteSpace(hex))
               return fallback;
          string trimmed = hex.Trim().TrimStart('#');
          return trimmed.Length == 6 ? trimmed : fallback;
     }

     private static string EscapeXmlAttribute(string value)
         => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}