using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace FlashXlsx.Core;

public readonly record struct XlsxSheetInfo(
    int Index,
    string Name,
    string ZipEntryPath,
    int? EstimatedRowCount = null
);

internal static class WorkbookMetadataReader
{
     public static (IReadOnlyList<XlsxSheetInfo> Sheets, bool Use1904DateSystem) ReadWorkbookInfo(ZipArchive archive)
     {
          var wbEntry = archive.GetEntry("xl/workbook.xml")
              ?? throw new InvalidOperationException("Archive XLSX invalide : xl/workbook.xml introuvable.");

          var relsMap = new Dictionary<string, string>(StringComparer.Ordinal);
          var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
          if (relsEntry is not null)
          {
               using var relsStream = relsEntry.Open();
               using var relsReader = XmlReader.Create(relsStream, new XmlReaderSettings { IgnoreWhitespace = true });
               while (relsReader.Read())
               {
                    if (relsReader.NodeType == XmlNodeType.Element && relsReader.LocalName == "Relationship")
                    {
                         string? id = relsReader.GetAttribute("Id");
                         string? target = relsReader.GetAttribute("Target");
                         if (id is not null && target is not null)
                         {
                              string normalized = target.TrimStart('/');
                              if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                                   normalized = "xl/" + normalized;
                              relsMap[id] = normalized;
                         }
                    }
               }
          }

          var sheets = new List<XlsxSheetInfo>();
          bool use1904 = false;

          using var wbStream = wbEntry.Open();
          using var wbReader = XmlReader.Create(wbStream, new XmlReaderSettings { IgnoreWhitespace = true });
          int idx = 0;
          while (wbReader.Read())
          {
               if (wbReader.NodeType == XmlNodeType.Element)
               {
                    if (wbReader.LocalName == "workbookPr")
                    {
                         string? date1904 = wbReader.GetAttribute("date1904");
                         if (date1904 is "1" or "true" or "TRUE")
                              use1904 = true;
                    }
                    else if (wbReader.LocalName == "sheet")
                    {
                         string name = wbReader.GetAttribute("name") ?? $"Sheet{idx + 1}";
                         string? rId = wbReader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                       ?? wbReader.GetAttribute("r:id");

                         string path = (rId is not null && relsMap.TryGetValue(rId, out var targetPath))
                             ? targetPath
                             : $"xl/worksheets/sheet{idx + 1}.xml";

                         int? estRows = ReadEstimatedRowCount(archive, path);
                         sheets.Add(new XlsxSheetInfo(idx++, name, path, estRows));
                    }
               }
          }

          return (sheets, use1904);
     }

     private static int? ReadEstimatedRowCount(ZipArchive archive, string entryPath)
     {
          var entry = archive.GetEntry(entryPath);
          if (entry is null) return null;

          try
          {
               using var stream = entry.Open();
               using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });
               while (reader.Read())
               {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                         if (reader.LocalName == "dimension")
                         {
                              string? refAttr = reader.GetAttribute("ref");
                              if (!string.IsNullOrEmpty(refAttr))
                              {
                                   int colonIdx = refAttr.IndexOf(':');
                                   string endCell = colonIdx >= 0 ? refAttr[(colonIdx + 1)..] : refAttr;
                                   int rowNum = ExtractRowNumber(endCell);
                                   return rowNum > 0 ? rowNum : null;
                              }
                              return null;
                         }

                         // <dimension> apparaît toujours avant <sheetData> en OpenXML
                         if (reader.LocalName == "sheetData")
                              return null;
                    }
               }
          }
          catch
          {
               // Ignore si <dimension> est absent ou malformé
          }

          return null;
     }

     private static int ExtractRowNumber(string cellRef)
     {
          int i = 0;
          while (i < cellRef.Length && char.IsLetter(cellRef[i]))
               i++;
          return int.TryParse(cellRef.AsSpan(i), out int row) ? row : 0;
     }
}