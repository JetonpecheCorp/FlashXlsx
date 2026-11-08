using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace FlashXlsx.Core;

internal static class SheetRowReader
{
     private readonly record struct RawCell(int ColIndex, string Value, bool IsNumeric, bool IsBoolean);

     private sealed class AtomizedXmlNames
     {
          public readonly NameTable Table = new();
          public readonly string Row;
          public readonly string Cell;
          public readonly string Value;
          public readonly string InlineString;
          public readonly string Text;
          public readonly string AttrRef;
          public readonly string AttrType;
          public readonly string TypeSharedString;
          public readonly string TypeInlineString;
          public readonly string TypeBoolean;
          public readonly string TypeNumber;

          public AtomizedXmlNames()
          {
               Row = Table.Add("row");
               Cell = Table.Add("c");
               Value = Table.Add("v");
               InlineString = Table.Add("is");
               Text = Table.Add("t");
               AttrRef = Table.Add("r");
               AttrType = Table.Add("t");
               TypeSharedString = Table.Add("s");
               TypeInlineString = Table.Add("inlineStr");
               TypeBoolean = Table.Add("b");
               TypeNumber = Table.Add("n");
          }
     }

     public static async IAsyncEnumerable<T> ReadTypedAsync<T>(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          var mergedTracker = LoadMergedCellsSafe(entry, options);
          var names = new AtomizedXmlNames();

          await using var rawStream = entry.Open();
          await using var bufferedStream = new BufferedStream(rawStream, 65536);
          using var reader = XmlReader.Create(bufferedStream, new XmlReaderSettings
          {
               NameTable = names.Table,
               IgnoreWhitespace = true,
               IgnoreComments = true,
               CheckCharacters = false
          });

          int colCount = map.ColumnCount;
          int[] colToProp = ArrayPool<int>.Shared.Rent(Math.Max(256, colCount * 4));
          Array.Fill(colToProp, -1);

          RawCell[] rowBuffer = ArrayPool<RawCell>.Shared.Rent(Math.Max(64, colCount * 2));
          int[] propToExcelCol = new int[colCount];
          Array.Fill(propToExcelCol, -1);
          string?[] excelHeaders = new string?[Math.Max(64, colCount * 2)];

          bool hasFixedColumns = false;
          for (int p = 0; p < colCount; p++)
          {
               int fixedCol = map.GetFixedColumnIndex(p);
               if (fixedCol >= 0)
               {
                    hasFixedColumns = true;
                    if (fixedCol >= colToProp.Length)
                         ResizeIntArray(ref colToProp, fixedCol + 64, -1);
                    colToProp[fixedCol] = p;
                    propToExcelCol[p] = fixedCol;
               }
          }

          int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          bool headerResolved = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
          int currentRow = 0;
          int rowsYielded = 0;
          int rowsProcessed = 0;
          const int maxHeaderSearchRows = 50;
          var inlineSb = new StringBuilder(128);

          try
          {
               while (TryReadNextNode(reader, currentRow, options))
               {
                    if ((rowsProcessed & 1023) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, names.Row))
                    {
                         string? rAttr = reader.GetAttribute(names.AttrRef);
                         currentRow = int.TryParse(rAttr, out int parsedRow) ? parsedRow : currentRow + 1;
                         bool isEmptyRowTag = reader.IsEmptyElement;

                         if (isEmptyRowTag && options.SkipEmptyRows)
                              continue;

                         if (!headerResolved && options.AutoDetectHeaderRow && currentRow <= maxHeaderSearchRows)
                         {
                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);
                              if (IsHeaderCandidateRow(rowBuffer.AsSpan(0, cellCount), map, options))
                              {
                                   resolvedHeaderRowNumber = currentRow;
                                   ApplyHeaderMapping(rowBuffer.AsSpan(0, cellCount), map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                                   headerResolved = true;
                              }
                              continue;
                         }

                         if (options.HasHeaderRow && !options.AutoDetectHeaderRow && currentRow == resolvedHeaderRowNumber)
                         {
                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);
                              ApplyHeaderMapping(rowBuffer.AsSpan(0, cellCount), map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                              headerResolved = true;
                              continue;
                         }

                         if (currentRow > resolvedHeaderRowNumber)
                         {
                              if (!headerResolved)
                              {
                                   if (!hasFixedColumns)
                                   {
                                        for (int p = 0; p < colCount; p++)
                                        {
                                             colToProp[p] = p;
                                             propToExcelCol[p] = p;
                                        }
                                   }
                                   headerResolved = true;
                              }

                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);

                              if (mergedTracker is not null)
                              {
                                   PropagateMergedCells(mergedTracker, currentRow, ref rowBuffer, ref cellCount);
                              }

                              if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                                   continue;

                              if (options.SkipEmptyRows && IsRowEffectivelyEmpty(rowBuffer.AsSpan(0, cellCount)))
                                   continue;

                              T instance = map.CreateInstance();

                              for (int i = 0; i < cellCount; i++)
                              {
                                   ref readonly RawCell cell = ref rowBuffer[i];
                                   if ((uint)cell.ColIndex >= (uint)colToProp.Length)
                                        continue;

                                   int propIdx = colToProp[cell.ColIndex];
                                   if (propIdx < 0)
                                        continue;

                                   var status = map.TrySetCellValue(ref instance, propIdx, cell.Value.AsSpan(), cell.IsNumeric, options);
                                   if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                                   {
                                        string? colHeader = (cell.ColIndex < excelHeaders.Length) ? excelHeaders[cell.ColIndex] : map.GetDefaultHeaderName(propIdx);
                                        var err = new XlsxErrorContext(currentRow, cell.ColIndex, XlsxCoordinateHelper.GetColumnLetter(cell.ColIndex), colHeader, cell.Value, status);
                                        options.OnError?.Invoke(in err);
                                        if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                             yield break;
                                   }
                              }

                              if (!map.ValidateRow(ref instance, currentRow, propToExcelCol, excelHeaders, options))
                              {
                                   if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                        yield break;
                              }

                              rowsProcessed++;
                              if (options.Progress is not null && rowsProcessed % options.ProgressInterval == 0)
                                   options.Progress.Report(rowsProcessed);

                              yield return instance;
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                    }
               }

               if (options.Progress is not null && rowsProcessed > 0 && rowsProcessed % options.ProgressInterval != 0)
                    options.Progress.Report(rowsProcessed);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(colToProp);
               ArrayPool<RawCell>.Shared.Return(rowBuffer, clearArray: true);
          }
     }

     public static async IAsyncEnumerable<T> ReadTypedAsync<T, TContext>(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         IXlsxTypeMap<T, TContext> map,
         XlsxReadOptions<T, TContext> options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          var mergedTracker = LoadMergedCellsSafe(entry, options);
          var names = new AtomizedXmlNames();

          await using var rawStream = entry.Open();
          await using var bufferedStream = new BufferedStream(rawStream, 65536);
          using var reader = XmlReader.Create(bufferedStream, new XmlReaderSettings
          {
               NameTable = names.Table,
               IgnoreWhitespace = true,
               IgnoreComments = true,
               CheckCharacters = false
          });

          int colCount = map.ColumnCount;
          int[] colToProp = ArrayPool<int>.Shared.Rent(Math.Max(256, colCount * 4));
          Array.Fill(colToProp, -1);

          RawCell[] rowBuffer = ArrayPool<RawCell>.Shared.Rent(Math.Max(64, colCount * 2));
          int[] propToExcelCol = new int[colCount];
          Array.Fill(propToExcelCol, -1);
          string?[] excelHeaders = new string?[Math.Max(64, colCount * 2)];

          bool hasFixedColumns = false;
          for (int p = 0; p < colCount; p++)
          {
               int fixedCol = map.GetFixedColumnIndex(p);
               if (fixedCol >= 0)
               {
                    hasFixedColumns = true;
                    if (fixedCol >= colToProp.Length)
                         ResizeIntArray(ref colToProp, fixedCol + 64, -1);
                    colToProp[fixedCol] = p;
                    propToExcelCol[p] = fixedCol;
               }
          }

          int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          bool headerResolved = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
          int currentRow = 0;
          int rowsYielded = 0;
          int rowsProcessed = 0;
          TContext context = options.Context;
          const int maxHeaderSearchRows = 50;
          var inlineSb = new StringBuilder(128);

          try
          {
               while (TryReadNextNode(reader, currentRow, options))
               {
                    if ((rowsProcessed & 1023) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, names.Row))
                    {
                         string? rAttr = reader.GetAttribute(names.AttrRef);
                         currentRow = int.TryParse(rAttr, out int parsedRow) ? parsedRow : currentRow + 1;
                         bool isEmptyRowTag = reader.IsEmptyElement;

                         if (isEmptyRowTag && options.SkipEmptyRows)
                              continue;

                         if (!headerResolved && options.AutoDetectHeaderRow && currentRow <= maxHeaderSearchRows)
                         {
                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);
                              if (IsHeaderCandidateRow(rowBuffer.AsSpan(0, cellCount), map, options))
                              {
                                   resolvedHeaderRowNumber = currentRow;
                                   ApplyHeaderMapping(rowBuffer.AsSpan(0, cellCount), map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                                   headerResolved = true;
                              }
                              continue;
                         }

                         if (options.HasHeaderRow && !options.AutoDetectHeaderRow && currentRow == resolvedHeaderRowNumber)
                         {
                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);
                              ApplyHeaderMapping(rowBuffer.AsSpan(0, cellCount), map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                              headerResolved = true;
                              continue;
                         }

                         if (currentRow > resolvedHeaderRowNumber)
                         {
                              if (!headerResolved)
                              {
                                   if (!hasFixedColumns)
                                   {
                                        for (int p = 0; p < colCount; p++)
                                        {
                                             colToProp[p] = p;
                                             propToExcelCol[p] = p;
                                        }
                                   }
                                   headerResolved = true;
                              }

                              int cellCount = isEmptyRowTag
                                  ? 0
                                  : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);

                              if (mergedTracker is not null)
                              {
                                   PropagateMergedCells(mergedTracker, currentRow, ref rowBuffer, ref cellCount);
                              }

                              if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                                   continue;

                              if (options.SkipEmptyRows && IsRowEffectivelyEmpty(rowBuffer.AsSpan(0, cellCount)))
                                   continue;

                              T instance = map.CreateInstance();

                              for (int i = 0; i < cellCount; i++)
                              {
                                   ref readonly RawCell cell = ref rowBuffer[i];
                                   if ((uint)cell.ColIndex >= (uint)colToProp.Length)
                                        continue;

                                   int propIdx = colToProp[cell.ColIndex];
                                   if (propIdx < 0)
                                        continue;

                                   var status = map.TrySetCellValue(ref instance, propIdx, cell.Value.AsSpan(), cell.IsNumeric, options);
                                   if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                                   {
                                        string? colHeader = (cell.ColIndex < excelHeaders.Length) ? excelHeaders[cell.ColIndex] : map.GetDefaultHeaderName(propIdx);
                                        var err = new XlsxErrorContext(currentRow, cell.ColIndex, XlsxCoordinateHelper.GetColumnLetter(cell.ColIndex), colHeader, cell.Value, status);
                                        options.OnError?.Invoke(in err);
                                        if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                             yield break;
                                   }
                              }

                              if (!map.ValidateRow(ref instance, in context, currentRow, propToExcelCol, excelHeaders, options))
                              {
                                   if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                        yield break;
                              }

                              rowsProcessed++;
                              if (options.Progress is not null && rowsProcessed % options.ProgressInterval == 0)
                                   options.Progress.Report(rowsProcessed);

                              yield return instance;
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                    }
               }

               if (options.Progress is not null && rowsProcessed > 0 && rowsProcessed % options.ProgressInterval != 0)
                    options.Progress.Report(rowsProcessed);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(colToProp);
               ArrayPool<RawCell>.Shared.Return(rowBuffer, clearArray: true);
          }
     }

     public static async IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadDictionaryAsync(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          var mergedTracker = LoadMergedCellsSafe(entry, options);
          var names = new AtomizedXmlNames();

          await using var rawStream = entry.Open();
          await using var bufferedStream = new BufferedStream(rawStream, 65536);
          using var reader = XmlReader.Create(bufferedStream, new XmlReaderSettings
          {
               NameTable = names.Table,
               IgnoreWhitespace = true,
               IgnoreComments = true,
               CheckCharacters = false
          });

          RawCell[] rowBuffer = ArrayPool<RawCell>.Shared.Rent(64);
          var headersByCol = new Dictionary<int, string>();
          int headerRow = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          int currentRow = 0;
          int rowsYielded = 0;
          var inlineSb = new StringBuilder(128);

          try
          {
               while (TryReadNextNode(reader, currentRow, options))
               {
                    if ((rowsYielded & 1023) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, names.Row))
                    {
                         string? rAttr = reader.GetAttribute(names.AttrRef);
                         currentRow = int.TryParse(rAttr, out int parsedRow) ? parsedRow : currentRow + 1;
                         bool isEmptyRowTag = reader.IsEmptyElement;

                         if (isEmptyRowTag && options.SkipEmptyRows)
                              continue;

                         int cellCount = isEmptyRowTag
                             ? 0
                             : TryReadRowIntoBuffer(reader, names, sharedStrings, ref rowBuffer, inlineSb, currentRow, options);

                         if (options.HasHeaderRow && currentRow == headerRow)
                         {
                              for (int i = 0; i < cellCount; i++)
                              {
                                   ref readonly RawCell c = ref rowBuffer[i];
                                   string colName = !string.IsNullOrWhiteSpace(c.Value)
                                       ? c.Value.Trim()
                                       : XlsxCoordinateHelper.GetColumnLetter(c.ColIndex);
                                   headersByCol[c.ColIndex] = colName;
                              }
                              continue;
                         }

                         if (currentRow > headerRow)
                         {
                              if (mergedTracker is not null)
                              {
                                   PropagateMergedCells(mergedTracker, currentRow, ref rowBuffer, ref cellCount);
                              }

                              if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                                   continue;

                              if (options.SkipEmptyRows && IsRowEffectivelyEmpty(rowBuffer.AsSpan(0, cellCount)))
                                   continue;

                              var dict = new Dictionary<string, XlsxValue>(cellCount, StringComparer.OrdinalIgnoreCase);
                              for (int i = 0; i < cellCount; i++)
                              {
                                   ref readonly RawCell c = ref rowBuffer[i];
                                   string key = headersByCol.TryGetValue(c.ColIndex, out var h)
                                       ? h
                                       : XlsxCoordinateHelper.GetColumnLetter(c.ColIndex);

                                   dict[key] = ParseDynamicValue(c.Value, c.IsNumeric, c.IsBoolean);
                              }

                              yield return dict;
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                    }
               }
          }
          finally
          {
               ArrayPool<RawCell>.Shared.Return(rowBuffer, clearArray: true);
          }
     }

     private static MergedCellsTracker? LoadMergedCellsSafe(ZipArchiveEntry entry, XlsxReadOptions options)
     {
          if (!options.FillMergedCells)
               return null;

          try
          {
               return MergedCellsTracker.LoadIfNeeded(entry, enabled: true);
          }
          catch (XmlException ex)
          {
               var err = new XlsxErrorContext(0, -1, string.Empty, null, ex.Message, XlsxCellStatus.CorruptedFile);
               options.OnError?.Invoke(in err);
               throw new InvalidDataException("Le contenu XML des cellules fusionnées (<mergeCells>) est corrompu.", ex);
          }
     }

     private static void PropagateMergedCells(
         MergedCellsTracker mergedTracker,
         int currentRow,
         ref RawCell[] rowBuffer,
         ref int cellCount)
     {
          for (int i = 0; i < cellCount; i++)
               mergedTracker.RecordCellValue(currentRow, rowBuffer[i].ColIndex, rowBuffer[i].Value);

          foreach (var (mCol, mVal) in mergedTracker.GetPropagatedValuesForRow(currentRow))
          {
               bool exists = false;
               for (int i = 0; i < cellCount; i++)
               {
                    if (rowBuffer[i].ColIndex == mCol && !string.IsNullOrEmpty(rowBuffer[i].Value))
                    {
                         exists = true;
                         break;
                    }
               }
               if (!exists)
               {
                    EnsureBufferCapacity(ref rowBuffer, cellCount + 1);
                    rowBuffer[cellCount++] = new RawCell(mCol, mVal, IsNumeric: false, IsBoolean: false);
               }
          }
     }

     private static bool TryReadNextNode(XmlReader xmlReader, int rowNum, XlsxReadOptions opts)
     {
          try
          {
               return xmlReader.Read();
          }
          catch (XmlException ex)
          {
               var err = new XlsxErrorContext(rowNum, -1, string.Empty, null, ex.Message, XlsxCellStatus.CorruptedFile);
               opts.OnError?.Invoke(in err);
               throw new InvalidDataException($"Le contenu XML de la feuille Excel est corrompu à la ligne {rowNum}.", ex);
          }
     }

     private static int TryReadRowIntoBuffer(
         XmlReader reader,
         AtomizedXmlNames names,
         string[] sharedStrings,
         ref RawCell[] buffer,
         StringBuilder inlineSb,
         int currentRow,
         XlsxReadOptions options)
     {
          try
          {
               return ReadRowIntoBuffer(reader, names, sharedStrings, ref buffer, inlineSb);
          }
          catch (XmlException ex)
          {
               var err = new XlsxErrorContext(currentRow, -1, string.Empty, null, ex.Message, XlsxCellStatus.CorruptedFile);
               options.OnError?.Invoke(in err);
               throw new InvalidDataException($"Le contenu XML de la feuille Excel est corrompu à la ligne {currentRow}.", ex);
          }
     }

     private static int ReadRowIntoBuffer(
         XmlReader reader,
         AtomizedXmlNames names,
         string[] sharedStrings,
         ref RawCell[] buffer,
         StringBuilder inlineSb)
     {
          int count = 0;
          int fallbackColIndex = 0;

          while (reader.Read())
          {
               if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.Row))
                    break;

               if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, names.Cell))
               {
                    bool isEmptyCell = reader.IsEmptyElement;
                    int colIdx = fallbackColIndex;
                    string? cellType = null;

                    if (reader.MoveToFirstAttribute())
                    {
                         do
                         {
                              if (ReferenceEquals(reader.LocalName, names.AttrRef))
                              {
                                   colIdx = ParseColumnIndex(reader.Value.AsSpan());
                              }
                              else if (ReferenceEquals(reader.LocalName, names.AttrType))
                              {
                                   cellType = reader.Value;
                              }
                         } while (reader.MoveToNextAttribute());
                         reader.MoveToElement();
                    }

                    fallbackColIndex = colIdx + 1;
                    if (isEmptyCell)
                         continue;

                    bool isNumeric = cellType is null || cellType == names.TypeNumber;
                    bool isBoolean = cellType == names.TypeBoolean;
                    string val = ReadCellContent(reader, names, cellType, sharedStrings, inlineSb);

                    EnsureBufferCapacity(ref buffer, count + 1);
                    buffer[count++] = new RawCell(colIdx, val, isNumeric, isBoolean);
               }
          }

          return count;
     }

     private static int ParseColumnIndex(ReadOnlySpan<char> cellRef)
     {
          int col = 0;
          for (int i = 0; i < cellRef.Length; i++)
          {
               char c = cellRef[i];
               if (c is >= 'A' and <= 'Z')
                    col = (col * 26) + (c - 'A' + 1);
               else if (c is >= 'a' and <= 'z')
                    col = (col * 26) + (c - 'a' + 1);
               else
                    break;
          }
          return col > 0 ? col - 1 : 0;
     }

     private static string ReadCellContent(
         XmlReader reader,
         AtomizedXmlNames names,
         string? cellType,
         string[] sharedStrings,
         StringBuilder inlineSb)
     {
          string? vText = null;
          string? inlineText = null;

          while (reader.Read())
          {
               if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.Cell))
                    break;

               if (reader.NodeType == XmlNodeType.Element)
               {
                    if (ReferenceEquals(reader.LocalName, names.Value) && !reader.IsEmptyElement)
                    {
                         vText = reader.ReadElementContentAsString();
                         if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.Cell))
                              break;
                    }
                    else if (ReferenceEquals(reader.LocalName, names.InlineString) && !reader.IsEmptyElement)
                    {
                         inlineText = ReadInlineString(reader, names, inlineSb);
                         if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.Cell))
                              break;
                    }
               }
          }

          if (cellType == names.TypeInlineString)
               return inlineText ?? string.Empty;

          if (vText is null)
               return string.Empty;

          if (cellType == names.TypeSharedString && int.TryParse(vText, out int sstIdx) && (uint)sstIdx < (uint)sharedStrings.Length)
               return sharedStrings[sstIdx];

          return vText;
     }

     private static string ReadInlineString(XmlReader reader, AtomizedXmlNames names, StringBuilder sb)
     {
          sb.Clear();
          while (reader.Read())
          {
               if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.InlineString))
                    break;

               if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, names.Text) && !reader.IsEmptyElement)
               {
                    sb.Append(reader.ReadElementContentAsString());
                    if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, names.InlineString))
                         break;
               }
          }
          return sb.ToString();
     }

     private static bool IsHeaderCandidateRow<T>(
         ReadOnlySpan<RawCell> rowCells,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options)
     {
          int matches = 0;
          for (int i = 0; i < rowCells.Length; i++)
          {
               string val = rowCells[i].Value.Trim();
               if (string.IsNullOrEmpty(val))
                    continue;

               int propIdx = ResolvePropertyWithTranslations(val, map, options);
               if (propIdx >= 0)
                    matches++;
          }
          return matches >= 1 && (matches >= 2 || map.ColumnCount == 1);
     }

     private static void ApplyHeaderMapping<T>(
         ReadOnlySpan<RawCell> rowCells,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options,
         ref int[] colToProp,
         int[] propToExcelCol,
         ref string?[] excelHeaders)
     {
          for (int i = 0; i < rowCells.Length; i++)
          {
               ref readonly RawCell cell = ref rowCells[i];
               if (cell.ColIndex >= colToProp.Length)
                    ResizeIntArray(ref colToProp, cell.ColIndex + 64, -1);

               if (cell.ColIndex >= excelHeaders.Length)
                    Array.Resize(ref excelHeaders, cell.ColIndex + 64);

               string trimmedHeader = cell.Value.Trim();
               excelHeaders[cell.ColIndex] = trimmedHeader;

               if (colToProp[cell.ColIndex] >= 0)
                    continue;

               int propIdx = ResolvePropertyWithTranslations(trimmedHeader, map, options);
               if (propIdx >= 0 && propToExcelCol[propIdx] < 0)
               {
                    colToProp[cell.ColIndex] = propIdx;
                    propToExcelCol[propIdx] = cell.ColIndex;
               }
          }
     }

     private static int ResolvePropertyWithTranslations<T>(string header, IXlsxTypeMap<T> map, XlsxReadOptions options)
     {
          foreach (var kvp in options.Translations)
          {
               for (int i = 0; i < kvp.Value.Length; i++)
               {
                    if (string.Equals(kvp.Value[i], header, StringComparison.OrdinalIgnoreCase))
                    {
                         int idx = map.ResolvePropertyIndex(kvp.Key.AsSpan());
                         if (idx >= 0) return idx;
                    }
               }
          }
          return map.ResolvePropertyIndex(header.AsSpan());
     }

     private static bool IsRowEffectivelyEmpty(ReadOnlySpan<RawCell> cells)
     {
          for (int i = 0; i < cells.Length; i++)
          {
               if (!string.IsNullOrWhiteSpace(cells[i].Value))
                    return false;
          }
          return true;
     }

     private static XlsxValue ParseDynamicValue(string raw, bool isNumeric, bool isBoolean)
     {
          if (string.IsNullOrEmpty(raw))
               return XlsxValue.Empty;

          if (isBoolean)
               return raw is "1" or "true" or "TRUE";

          if (isNumeric)
          {
               if (long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out long l))
                    return l;
               if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                    return d;
          }

          return raw;
     }

     private static void EnsureBufferCapacity(ref RawCell[] buffer, int requiredSize)
     {
          if (requiredSize <= buffer.Length)
               return;

          RawCell[] newBuf = ArrayPool<RawCell>.Shared.Rent(buffer.Length * 2);
          Array.Copy(buffer, newBuf, buffer.Length);
          ArrayPool<RawCell>.Shared.Return(buffer, clearArray: true);
          buffer = newBuf;
     }

     private static void ResizeIntArray(ref int[] array, int newSize, int defaultFill)
     {
          int oldLen = array.Length;
          Array.Resize(ref array, newSize);
          Array.Fill(array, defaultFill, oldLen, newSize - oldLen);
     }
}