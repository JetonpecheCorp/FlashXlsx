using System.Collections.Generic;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace FlashXlsx.Core;

internal static class SharedStringsReader
{
     private static readonly XmlReaderSettings Settings = new()
     {
          IgnoreWhitespace = true,
          IgnoreComments = true,
          IgnoreProcessingInstructions = true,
          DtdProcessing = DtdProcessing.Prohibit
     };

     public static string[] Read(ZipArchive archive)
     {
          var entry = archive.GetEntry("xl/sharedStrings.xml");
          if (entry is null)
               return [];

          using var stream = entry.Open();
          using var reader = XmlReader.Create(stream, Settings);

          var strings = new List<string>();
          while (reader.Read())
          {
               if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
               {
                    if (reader.IsEmptyElement)
                    {
                         strings.Add(string.Empty);
                         continue;
                    }

                    string singleText = string.Empty;
                    StringBuilder? richSb = null;

                    while (reader.Read())
                    {
                         if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si")
                              break;

                         if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
                         {
                              string part = reader.ReadElementContentAsString();
                              if (richSb is not null)
                              {
                                   richSb.Append(part);
                              }
                              else if (singleText.Length > 0)
                              {
                                   richSb = new StringBuilder(singleText.Length + part.Length);
                                   richSb.Append(singleText).Append(part);
                              }
                              else
                              {
                                   singleText = part;
                              }

                              if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si")
                                   break;
                         }
                    }

                    strings.Add(richSb is not null ? richSb.ToString() : singleText);
               }
          }

          return [.. strings];
     }
}