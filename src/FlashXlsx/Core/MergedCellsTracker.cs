using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace FlashXlsx.Core;

internal sealed class MergedCellsTracker
{
     private readonly record struct MergeRegion(int StartRow, int EndRow, int StartCol, int EndCol);

     private readonly List<MergeRegion> _regions;
     private readonly Dictionary<(int StartRow, int StartCol), string> _anchorValues = new();

     private MergedCellsTracker(List<MergeRegion> regions)
     {
          _regions = regions;
     }

     public static MergedCellsTracker? LoadIfNeeded(ZipArchiveEntry entry, bool enabled)
     {
          if (!enabled)
               return null;

          var regions = new List<MergeRegion>();
          using var stream = entry.Open();
          using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

          while (reader.Read())
          {
               if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "mergeCell")
               {
                    string? refAttr = reader.GetAttribute("ref");
                    if (!string.IsNullOrEmpty(refAttr) && TryParseRange(refAttr, out var region))
                    {
                         regions.Add(region);
                    }
               }
          }

          return regions.Count > 0 ? new MergedCellsTracker(regions) : null;
     }

     public void RecordCellValue(int row, int col, string? value)
     {
          if (string.IsNullOrEmpty(value))
               return;

          for (int i = 0; i < _regions.Count; i++)
          {
               var r = _regions[i];
               if (r.StartRow == row && r.StartCol == col)
               {
                    _anchorValues[(row, col)] = value;
                    return;
               }
          }
     }

     public IEnumerable<(int Col, string Value)> GetPropagatedValuesForRow(int row)
     {
          for (int i = 0; i < _regions.Count; i++)
          {
               var r = _regions[i];
               if (row >= r.StartRow && row <= r.EndRow && _anchorValues.TryGetValue((r.StartRow, r.StartCol), out var val))
               {
                    for (int c = r.StartCol; c <= r.EndCol; c++)
                    {
                         if (row == r.StartRow && c == r.StartCol)
                              continue;

                         yield return (c, val);
                    }
               }
          }
     }

     private static bool TryParseRange(string refRange, out MergeRegion region)
     {
          region = default;
          int colonIdx = refRange.IndexOf(':');
          if (colonIdx <= 0)
               return false;

          if (!TryParseCellRef(refRange.AsSpan(0, colonIdx), out int r1, out int c1) ||
              !TryParseCellRef(refRange.AsSpan(colonIdx + 1), out int r2, out int c2))
               return false;

          region = new MergeRegion(Math.Min(r1, r2), Math.Max(r1, r2), Math.Min(c1, c2), Math.Max(c1, c2));
          return true;
     }

     private static bool TryParseCellRef(ReadOnlySpan<char> cellRef, out int row, out int col)
     {
          row = 0;
          col = 0;
          int i = 0;
          while (i < cellRef.Length && char.IsLetter(cellRef[i]))
          {
               char upper = char.ToUpperInvariant(cellRef[i]);
               if (upper is < 'A' or > 'Z') return false;
               col = (col * 26) + (upper - 'A' + 1);
               i++;
          }
          col -= 1;
          if (col < 0 || i >= cellRef.Length) return false;
          return int.TryParse(cellRef[i..], out row) && row > 0;
     }
}