using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace FlashXlsx.Core;

public readonly record struct XlsxSheetInfo(
    int Index,
    string Name,
    string SheetId,
    bool IsVisible,
    string ZipEntryPath
);

internal static class WorkbookMetadataReader
{
     private static readonly XmlReaderSettings Settings = new()
     {
          IgnoreWhitespace = true,
          IgnoreComments = true,
          IgnoreProcessingInstructions = true,
          DtdProcessing = DtdProcessing.Prohibit
     };

     public static IReadOnlyList<XlsxSheetInfo> ReadSheets(ZipArchive archive)
         => ReadWorkbookInfo(archive).Sheets;

     public static (IReadOnlyList<XlsxSheetInfo> Sheets, bool Use1904DateSystem) ReadWorkbookInfo(ZipArchive archive)
     {
          // 1. Lire xl/_rels/workbook.xml.rels pour mapper rId -> chemin cible
          var relsMap = new Dictionary<string, string>(StringComparer.Ordinal);
          var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
          if (relsEntry is not null)
          {
               using var relsStream = relsEntry.Open();
               using var relsReader = XmlReader.Create(relsStream, Settings);
               while (relsReader.Read())
               {
                    if (relsReader.NodeType == XmlNodeType.Element && relsReader.LocalName == "Relationship")
                    {
                         var id = relsReader.GetAttribute("Id");
                         var target = relsReader.GetAttribute("Target");
                         if (id is not null && target is not null)
                         {
                              var normalized = target.StartsWith('/')
                                  ? target.TrimStart('/')
                                  : target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                                      ? target
                                      : $"xl/{target}";
                              relsMap[id] = normalized;
                         }
                    }
               }
          }

          // 2. Lire xl/workbook.xml pour récupérer la liste ordonnée des onglets et la propriété date1904
          var sheets = new List<XlsxSheetInfo>();
          bool use1904 = false;

          var wbEntry = archive.GetEntry("xl/workbook.xml");
          if (wbEntry is null)
               return (sheets, false);

          using var wbStream = wbEntry.Open();
          using var wbReader = XmlReader.Create(wbStream, Settings);
          int index = 0;

          while (wbReader.Read())
          {
               if (wbReader.NodeType == XmlNodeType.Element)
               {
                    if (wbReader.LocalName == "workbookPr")
                    {
                         var date1904Attr = wbReader.GetAttribute("date1904");
                         if (date1904Attr is "1" or "true" or "TRUE")
                              use1904 = true;
                    }
                    else if (wbReader.LocalName == "sheet")
                    {
                         var name = wbReader.GetAttribute("name") ?? $"Sheet{index + 1}";
                         var sheetId = wbReader.GetAttribute("sheetId") ?? (index + 1).ToString();
                         var state = wbReader.GetAttribute("state");
                         var rId = wbReader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                   ?? wbReader.GetAttribute("r:id");

                         string path = (rId is not null && relsMap.TryGetValue(rId, out var resolved))
                             ? resolved
                             : $"xl/worksheets/sheet{index + 1}.xml";

                         bool isVisible = !string.Equals(state, "hidden", StringComparison.OrdinalIgnoreCase)
                                       && !string.Equals(state, "veryHidden", StringComparison.OrdinalIgnoreCase);

                         sheets.Add(new XlsxSheetInfo(index++, name, sheetId, isVisible, path));
                    }
               }
          }

          return (sheets, use1904);
     }
}