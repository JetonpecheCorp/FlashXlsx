using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using System.Buffers;
using System.Buffers.Text;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;

namespace FlashXlsx.Core;

internal static class Utf8SheetRowReader
{
     public static bool CanUseFastUtf8Path(XlsxReadOptions options)
         => !options.FillMergedCells;

     private static readonly SearchValues<byte> s_tagDelimiters = SearchValues.Create("<>"u8);

     public static async IAsyncEnumerable<T> ReadTypedFastAsync<T>(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          await using var stream = entry.Open();

          int colCount = map.ColumnCount;
          int[] colToProp = ArrayPool<int>.Shared.Rent(Math.Max(256, colCount * 4));
          Array.Fill(colToProp, -1);

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

          byte[] buffer = ArrayPool<byte>.Shared.Rent(131072);
          char[] charScratch = ArrayPool<char>.Shared.Rent(4096);
          int bufStart = 0;
          int bufEnd = 0;
          bool eof = false;

          int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          bool headerResolved = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
          int currentRow = 0;
          int rowsYielded = 0;
          int rowsProcessed = 0;
          const int maxHeaderSearchRows = 50;

          try
          {
               while (true)
               {
                    if ((rowsProcessed & 2047) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (!eof && (bufEnd - bufStart) < 32768)
                    {
                         if (bufStart > 0)
                         {
                              int remaining = bufEnd - bufStart;
                              if (remaining > 0)
                                   Buffer.BlockCopy(buffer, bufStart, buffer, 0, remaining);
                              bufStart = 0;
                              bufEnd = remaining;
                         }

                         while (bufEnd < buffer.Length)
                         {
                              int read = await stream.ReadAsync(buffer.AsMemory(bufEnd, buffer.Length - bufEnd), ct);
                              if (read == 0)
                              {
                                   eof = true;
                                   break;
                              }
                              bufEnd += read;
                         }
                    }

                    if (bufStart >= bufEnd)
                         break;

                    ReadOnlySpan<byte> span = buffer.AsSpan(bufStart, bufEnd - bufStart);

                    int rowTagIdx = FindTag(span, "<row"u8);
                    if (rowTagIdx < 0)
                    {
                         if (eof)
                              break;
                         int keep = Math.Min(8, span.Length);
                         bufStart = bufEnd - keep;
                         continue;
                    }

                    span = span[rowTagIdx..];
                    bufStart += rowTagIdx;

                    int openTagEnd = span.IndexOf((byte)'>');
                    if (openTagEnd < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);
                         continue;
                    }

                    ReadOnlySpan<byte> openTag = span[..openTagEnd];
                    bool isSelfClosingRow = openTagEnd > 0 && openTag[^1] == (byte)'/';

                    if (TryExtractIntAttribute(openTag, "r"u8, out int parsedRow))
                         currentRow = parsedRow;
                    else
                         currentRow++;

                    if (isSelfClosingRow)
                    {
                         bufStart += openTagEnd + 1;
                         if (options.SkipEmptyRows)
                              continue;

                         if (currentRow > resolvedHeaderRowNumber && (!options.StartRow.HasValue || currentRow >= options.StartRow.Value))
                         {
                              T emptyInstance = map.CreateInstance();
                              if (!map.ValidateRow(ref emptyInstance, currentRow, propToExcelCol, excelHeaders, options) &&
                                  options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                              {
                                   yield break;
                              }

                              rowsProcessed++;
                              yield return emptyInstance;
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                         continue;
                    }

                    int closeRowIdx = FindTag(span[(openTagEnd + 1)..], "</row>"u8);
                    if (closeRowIdx < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);

                         if (bufStart == 0 && bufEnd == buffer.Length)
                         {
                              byte[] larger = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
                              Buffer.BlockCopy(buffer, 0, larger, 0, bufEnd);
                              ArrayPool<byte>.Shared.Return(buffer);
                              buffer = larger;
                         }
                         continue;
                    }

                    ReadOnlySpan<byte> rowInner = span.Slice(openTagEnd + 1, closeRowIdx);
                    bufStart += openTagEnd + 1 + closeRowIdx + 6;

                    if (!headerResolved && options.AutoDetectHeaderRow && currentRow <= maxHeaderSearchRows)
                    {
                         if (IsHeaderRowUtf8(rowInner, sharedStrings, map, options))
                         {
                              resolvedHeaderRowNumber = currentRow;
                              ApplyHeaderMappingUtf8(rowInner, sharedStrings, map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                              headerResolved = true;
                         }
                         continue;
                    }

                    if (options.HasHeaderRow && !options.AutoDetectHeaderRow && currentRow == resolvedHeaderRowNumber)
                    {
                         ApplyHeaderMappingUtf8(rowInner, sharedStrings, map, options, ref colToProp, propToExcelCol, ref excelHeaders);
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

                         if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                              continue;

                         if (options.SkipEmptyRows && IsRowEmptyUtf8(rowInner, sharedStrings))
                              continue;

                         T instance = map.CreateInstance();
                         int fallbackCol = 0;
                         int pos = 0;

                         while (pos < rowInner.Length)
                         {
                              int cIdx = FindTag(rowInner[pos..], "<c"u8);
                              if (cIdx < 0)
                                   break;

                              pos += cIdx;
                              int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                              if (cOpenEnd < 0)
                                   break;

                              ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                              bool isSelfClosingCell = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                              pos += cOpenEnd + 1;

                              int colIdx = TryExtractCellColIndex(cTag, out int parsedCol) ? parsedCol : fallbackCol;
                              fallbackCol = colIdx + 1;

                              if (isSelfClosingCell)
                                   continue;

                              int cClose = FindTag(rowInner[pos..], "</c>"u8);
                              if (cClose < 0)
                                   break;

                              ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                              pos += cClose + 4;

                              if ((uint)colIdx >= (uint)colToProp.Length)
                                   continue;

                              int propIdx = colToProp[colIdx];
                              if (propIdx < 0)
                                   continue;

                              byte cellTypeCode = ExtractCellTypeCode(cTag);
                              bool isNumeric = cellTypeCode == 0;

                              ReadOnlySpan<char> valueChars = DecodeCellValueToSpan(
                                  cellInner,
                                  cellTypeCode,
                                  sharedStrings,
                                  ref charScratch,
                                  out string? rawStringForError);

                              var status = map.TrySetCellValue(ref instance, propIdx, valueChars, isNumeric, options);
                              if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                              {
                                   string rawErr = rawStringForError ?? valueChars.ToString();
                                   string? colHeader = (colIdx < excelHeaders.Length) ? excelHeaders[colIdx] : map.GetDefaultHeaderName(propIdx);
                                   var err = new XlsxErrorContext(currentRow, colIdx, XlsxCoordinateHelper.GetColumnLetter(colIdx), colHeader, rawErr, status);
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

               if (options.Progress is not null && rowsProcessed > 0 && rowsProcessed % options.ProgressInterval != 0)
                    options.Progress.Report(rowsProcessed);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(colToProp);
               ArrayPool<byte>.Shared.Return(buffer);
               ArrayPool<char>.Shared.Return(charScratch);
          }
     }

     public static async IAsyncEnumerable<T> ReadTypedContextFastAsync<T, TContext>(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         IXlsxTypeMap<T, TContext> map,
         XlsxReadOptions<T, TContext> options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          await using var stream = entry.Open();

          int colCount = map.ColumnCount;
          int[] colToProp = ArrayPool<int>.Shared.Rent(Math.Max(256, colCount * 4));
          Array.Fill(colToProp, -1);

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

          byte[] buffer = ArrayPool<byte>.Shared.Rent(131072);
          char[] charScratch = ArrayPool<char>.Shared.Rent(4096);
          int bufStart = 0;
          int bufEnd = 0;
          bool eof = false;

          int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          bool headerResolved = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
          int currentRow = 0;
          int rowsYielded = 0;
          int rowsProcessed = 0;
          TContext context = options.Context;
          const int maxHeaderSearchRows = 50;

          try
          {
               while (true)
               {
                    if ((rowsProcessed & 2047) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (!eof && (bufEnd - bufStart) < 32768)
                    {
                         if (bufStart > 0)
                         {
                              int remaining = bufEnd - bufStart;
                              if (remaining > 0)
                                   Buffer.BlockCopy(buffer, bufStart, buffer, 0, remaining);
                              bufStart = 0;
                              bufEnd = remaining;
                         }

                         while (bufEnd < buffer.Length)
                         {
                              int read = await stream.ReadAsync(buffer.AsMemory(bufEnd, buffer.Length - bufEnd), ct);
                              if (read == 0)
                              {
                                   eof = true;
                                   break;
                              }
                              bufEnd += read;
                         }
                    }

                    if (bufStart >= bufEnd)
                         break;

                    ReadOnlySpan<byte> span = buffer.AsSpan(bufStart, bufEnd - bufStart);

                    int rowTagIdx = FindTag(span, "<row"u8);
                    if (rowTagIdx < 0)
                    {
                         if (eof)
                              break;
                         int keep = Math.Min(8, span.Length);
                         bufStart = bufEnd - keep;
                         continue;
                    }

                    span = span[rowTagIdx..];
                    bufStart += rowTagIdx;

                    int openTagEnd = span.IndexOf((byte)'>');
                    if (openTagEnd < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);
                         continue;
                    }

                    ReadOnlySpan<byte> openTag = span[..openTagEnd];
                    bool isSelfClosingRow = openTagEnd > 0 && openTag[^1] == (byte)'/';

                    if (TryExtractIntAttribute(openTag, "r"u8, out int parsedRow))
                         currentRow = parsedRow;
                    else
                         currentRow++;

                    if (isSelfClosingRow)
                    {
                         bufStart += openTagEnd + 1;
                         if (options.SkipEmptyRows)
                              continue;

                         if (currentRow > resolvedHeaderRowNumber && (!options.StartRow.HasValue || currentRow >= options.StartRow.Value))
                         {
                              T emptyInstance = map.CreateInstance();
                              if (!map.ValidateRow(ref emptyInstance, in context, currentRow, propToExcelCol, excelHeaders, options) &&
                                  options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                              {
                                   yield break;
                              }

                              rowsProcessed++;
                              yield return emptyInstance;
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                         continue;
                    }

                    int closeRowIdx = FindTag(span[(openTagEnd + 1)..], "</row>"u8);
                    if (closeRowIdx < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);

                         if (bufStart == 0 && bufEnd == buffer.Length)
                         {
                              byte[] larger = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
                              Buffer.BlockCopy(buffer, 0, larger, 0, bufEnd);
                              ArrayPool<byte>.Shared.Return(buffer);
                              buffer = larger;
                         }
                         continue;
                    }

                    ReadOnlySpan<byte> rowInner = span.Slice(openTagEnd + 1, closeRowIdx);
                    bufStart += openTagEnd + 1 + closeRowIdx + 6;

                    if (!headerResolved && options.AutoDetectHeaderRow && currentRow <= maxHeaderSearchRows)
                    {
                         if (IsHeaderRowUtf8(rowInner, sharedStrings, map, options))
                         {
                              resolvedHeaderRowNumber = currentRow;
                              ApplyHeaderMappingUtf8(rowInner, sharedStrings, map, options, ref colToProp, propToExcelCol, ref excelHeaders);
                              headerResolved = true;
                         }
                         continue;
                    }

                    if (options.HasHeaderRow && !options.AutoDetectHeaderRow && currentRow == resolvedHeaderRowNumber)
                    {
                         ApplyHeaderMappingUtf8(rowInner, sharedStrings, map, options, ref colToProp, propToExcelCol, ref excelHeaders);
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

                         if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                              continue;

                         if (options.SkipEmptyRows && IsRowEmptyUtf8(rowInner, sharedStrings))
                              continue;

                         T instance = map.CreateInstance();
                         int fallbackCol = 0;
                         int pos = 0;

                         while (pos < rowInner.Length)
                         {
                              int cIdx = FindTag(rowInner[pos..], "<c"u8);
                              if (cIdx < 0)
                                   break;

                              pos += cIdx;
                              int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                              if (cOpenEnd < 0)
                                   break;

                              ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                              bool isSelfClosingCell = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                              pos += cOpenEnd + 1;

                              int colIdx = TryExtractCellColIndex(cTag, out int parsedCol) ? parsedCol : fallbackCol;
                              fallbackCol = colIdx + 1;

                              if (isSelfClosingCell)
                                   continue;

                              int cClose = FindTag(rowInner[pos..], "</c>"u8);
                              if (cClose < 0)
                                   break;

                              ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                              pos += cClose + 4;

                              if ((uint)colIdx >= (uint)colToProp.Length)
                                   continue;

                              int propIdx = colToProp[colIdx];
                              if (propIdx < 0)
                                   continue;

                              byte cellTypeCode = ExtractCellTypeCode(cTag);
                              bool isNumeric = cellTypeCode == 0;

                              ReadOnlySpan<char> valueChars = DecodeCellValueToSpan(
                                  cellInner,
                                  cellTypeCode,
                                  sharedStrings,
                                  ref charScratch,
                                  out string? rawStringForError);

                              var status = map.TrySetCellValue(ref instance, propIdx, valueChars, isNumeric, options);
                              if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                              {
                                   string rawErr = rawStringForError ?? valueChars.ToString();
                                   string? colHeader = (colIdx < excelHeaders.Length) ? excelHeaders[colIdx] : map.GetDefaultHeaderName(propIdx);
                                   var err = new XlsxErrorContext(currentRow, colIdx, XlsxCoordinateHelper.GetColumnLetter(colIdx), colHeader, rawErr, status);
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

               if (options.Progress is not null && rowsProcessed > 0 && rowsProcessed % options.ProgressInterval != 0)
                    options.Progress.Report(rowsProcessed);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(colToProp);
               ArrayPool<byte>.Shared.Return(buffer);
               ArrayPool<char>.Shared.Return(charScratch);
          }
     }

     public static async IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadDictionaryFastAsync(
         ZipArchiveEntry entry,
         string[] sharedStrings,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          if (options.MaxRows == 0)
               yield break;

          await using var stream = entry.Open();

          byte[] buffer = ArrayPool<byte>.Shared.Rent(131072);
          char[] charScratch = ArrayPool<char>.Shared.Rent(4096);
          int bufStart = 0;
          int bufEnd = 0;
          bool eof = false;

          var headersByCol = new Dictionary<int, string>();
          int headerRow = options.HasHeaderRow ? options.HeaderRowNumber : 0;
          int currentRow = 0;
          int rowsYielded = 0;

          try
          {
               while (true)
               {
                    if ((rowsYielded & 2047) == 0)
                         ct.ThrowIfCancellationRequested();

                    if (!eof && (bufEnd - bufStart) < 32768)
                    {
                         if (bufStart > 0)
                         {
                              int remaining = bufEnd - bufStart;
                              if (remaining > 0)
                                   Buffer.BlockCopy(buffer, bufStart, buffer, 0, remaining);
                              bufStart = 0;
                              bufEnd = remaining;
                         }

                         while (bufEnd < buffer.Length)
                         {
                              int read = await stream.ReadAsync(buffer.AsMemory(bufEnd, buffer.Length - bufEnd), ct);
                              if (read == 0)
                              {
                                   eof = true;
                                   break;
                              }
                              bufEnd += read;
                         }
                    }

                    if (bufStart >= bufEnd)
                         break;

                    ReadOnlySpan<byte> span = buffer.AsSpan(bufStart, bufEnd - bufStart);

                    int rowTagIdx = FindTag(span, "<row"u8);
                    if (rowTagIdx < 0)
                    {
                         if (eof)
                              break;
                         int keep = Math.Min(8, span.Length);
                         bufStart = bufEnd - keep;
                         continue;
                    }

                    span = span[rowTagIdx..];
                    bufStart += rowTagIdx;

                    int openTagEnd = span.IndexOf((byte)'>');
                    if (openTagEnd < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);
                         continue;
                    }

                    ReadOnlySpan<byte> openTag = span[..openTagEnd];
                    bool isSelfClosingRow = openTagEnd > 0 && openTag[^1] == (byte)'/';

                    if (TryExtractIntAttribute(openTag, "r"u8, out int parsedRow))
                         currentRow = parsedRow;
                    else
                         currentRow++;

                    if (isSelfClosingRow)
                    {
                         bufStart += openTagEnd + 1;
                         if (options.SkipEmptyRows)
                              continue;

                         if (currentRow > headerRow && (!options.StartRow.HasValue || currentRow >= options.StartRow.Value))
                         {
                              yield return new Dictionary<string, XlsxValue>(StringComparer.OrdinalIgnoreCase);
                              rowsYielded++;
                              if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                                   yield break;
                         }
                         continue;
                    }

                    int closeRowIdx = FindTag(span[(openTagEnd + 1)..], "</row>"u8);
                    if (closeRowIdx < 0)
                    {
                         if (eof)
                              NotifyCorruptedAndThrow(currentRow, options);

                         if (bufStart == 0 && bufEnd == buffer.Length)
                         {
                              byte[] larger = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
                              Buffer.BlockCopy(buffer, 0, larger, 0, bufEnd);
                              ArrayPool<byte>.Shared.Return(buffer);
                              buffer = larger;
                         }
                         continue;
                    }

                    ReadOnlySpan<byte> rowInner = span.Slice(openTagEnd + 1, closeRowIdx);
                    bufStart += openTagEnd + 1 + closeRowIdx + 6;

                    if (options.HasHeaderRow && currentRow == headerRow)
                    {
                         int fallbackCol = 0;
                         int pos = 0;
                         while (pos < rowInner.Length)
                         {
                              int cIdx = FindTag(rowInner[pos..], "<c"u8);
                              if (cIdx < 0) break;
                              pos += cIdx;
                              int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                              if (cOpenEnd < 0) break;
                              ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                              bool selfClosing = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                              pos += cOpenEnd + 1;

                              int colIdx = TryExtractCellColIndex(cTag, out int parsedCol) ? parsedCol : fallbackCol;
                              fallbackCol = colIdx + 1;
                              if (selfClosing) continue;

                              int cClose = FindTag(rowInner[pos..], "</c>"u8);
                              if (cClose < 0) break;
                              ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                              pos += cClose + 4;

                              byte typeCode = ExtractCellTypeCode(cTag);
                              var chars = DecodeCellValueToSpan(cellInner, typeCode, sharedStrings, ref charScratch, out _).Trim();
                              string colName = !chars.IsEmpty
                                  ? chars.ToString()
                                  : XlsxCoordinateHelper.GetColumnLetter(colIdx);
                              headersByCol[colIdx] = colName;
                         }
                         continue;
                    }

                    if (currentRow > headerRow)
                    {
                         if (options.StartRow.HasValue && currentRow < options.StartRow.Value)
                              continue;

                         if (options.SkipEmptyRows && IsRowEmptyUtf8(rowInner, sharedStrings))
                              continue;

                         var dict = new Dictionary<string, XlsxValue>(Math.Max(4, headersByCol.Count), StringComparer.OrdinalIgnoreCase);
                         int fallbackCol = 0;
                         int pos = 0;

                         while (pos < rowInner.Length)
                         {
                              int cIdx = FindTag(rowInner[pos..], "<c"u8);
                              if (cIdx < 0) break;
                              pos += cIdx;
                              int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                              if (cOpenEnd < 0) break;
                              ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                              bool selfClosing = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                              pos += cOpenEnd + 1;

                              int colIdx = TryExtractCellColIndex(cTag, out int parsedCol) ? parsedCol : fallbackCol;
                              fallbackCol = colIdx + 1;
                              if (selfClosing) continue;

                              int cClose = FindTag(rowInner[pos..], "</c>"u8);
                              if (cClose < 0) break;
                              ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                              pos += cClose + 4;

                              string key = headersByCol.TryGetValue(colIdx, out var h)
                                  ? h
                                  : XlsxCoordinateHelper.GetColumnLetter(colIdx);

                              byte typeCode = ExtractCellTypeCode(cTag);
                              dict[key] = ParseDynamicValueUtf8(cellInner, typeCode, sharedStrings, ref charScratch);
                         }

                         yield return dict;
                         rowsYielded++;
                         if (options.MaxRows.HasValue && rowsYielded >= options.MaxRows.Value)
                              yield break;
                    }
               }
          }
          finally
          {
               ArrayPool<byte>.Shared.Return(buffer);
               ArrayPool<char>.Shared.Return(charScratch);
          }
     }

     private static XlsxValue ParseDynamicValueUtf8(
         ReadOnlySpan<byte> cellInner,
         byte cellTypeCode,
         string[] sharedStrings,
         ref char[] charScratch)
     {
          if (cellTypeCode == 1) // sharedString ('s')
          {
               var chars = DecodeCellValueToSpan(cellInner, 1, sharedStrings, ref charScratch, out string? cached);
               if (chars.IsEmpty) return XlsxValue.Empty;
               return cached ?? chars.ToString();
          }

          if (cellTypeCode == 2) // inlineStr
          {
               var chars = DecodeCellValueToSpan(cellInner, 2, sharedStrings, ref charScratch, out string? cached);
               if (chars.IsEmpty) return XlsxValue.Empty;
               return cached ?? chars.ToString();
          }

          // Extraction directe de <v>...</v> en UTF-8 (zéro allocation pour nombres et booléens !)
          int vStart = FindTag(cellInner, "<v"u8);
          if (vStart < 0) return XlsxValue.Empty;
          int vOpenEnd = cellInner[vStart..].IndexOf((byte)'>');
          if (vOpenEnd < 0) return XlsxValue.Empty;
          int vContentStart = vStart + vOpenEnd + 1;
          int vClose = FindTag(cellInner[vContentStart..], "</v>"u8);
          if (vClose <= 0) return XlsxValue.Empty;

          ReadOnlySpan<byte> vUtf8 = cellInner.Slice(vContentStart, vClose);

          if (cellTypeCode == 3) // boolean ('b')
               return vUtf8.SequenceEqual("1"u8) || vUtf8.SequenceEqual("true"u8) || vUtf8.SequenceEqual("TRUE"u8);

          if (cellTypeCode == 0) // numeric
          {
               if (Utf8Parser.TryParse(vUtf8, out long l, out int bytesConsumed) && bytesConsumed == vUtf8.Length)
                    return l;
               if (Utf8Parser.TryParse(vUtf8, out double d, out bytesConsumed) && bytesConsumed == vUtf8.Length)
                    return d;
          }

          return Encoding.UTF8.GetString(vUtf8);
     }

     [MethodImpl(MethodImplOptions.AggressiveInlining)]
     private static int FindTag(ReadOnlySpan<byte> span, ReadOnlySpan<byte> tag)
     {
          int offset = 0;
          while (offset < span.Length)
          {
               int idx = span[offset..].IndexOf(tag);
               if (idx < 0)
                    return -1;

               int abs = offset + idx;
               int after = abs + tag.Length;
               if (tag[^1] == (byte)'>' || after >= span.Length ||
                   span[after] is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n')
               {
                    return abs;
               }
               offset = abs + 1;
          }
          return -1;
     }

     private static bool TryExtractIntAttribute(ReadOnlySpan<byte> tag, ReadOnlySpan<byte> attrName, out int result)
     {
          result = 0;
          int pos = 0;
          while (pos < tag.Length)
          {
               int eq = tag[pos..].IndexOf((byte)'=');
               if (eq < 0) break;
               int absEq = pos + eq;

               int nameStart = absEq - 1;
               while (nameStart >= 0 && tag[nameStart] != (byte)' ' && tag[nameStart] != (byte)'\t')
                    nameStart--;
               nameStart++;

               ReadOnlySpan<byte> name = tag.Slice(nameStart, absEq - nameStart);
               int valStart = absEq + 2;
               if (valStart >= tag.Length) break;

               byte quote = tag[absEq + 1];
               int valEnd = tag[valStart..].IndexOf(quote);
               if (valEnd < 0) break;

               if (name.SequenceEqual(attrName))
                    return Utf8Parser.TryParse(tag.Slice(valStart, valEnd), out result, out _);

               pos = valStart + valEnd + 1;
          }
          return false;
     }

     private static bool TryExtractCellColIndex(ReadOnlySpan<byte> tag, out int colIdx)
     {
          colIdx = 0;
          int pos = 0;
          while (pos < tag.Length)
          {
               int eq = tag[pos..].IndexOf((byte)'=');
               if (eq < 0) break;
               int absEq = pos + eq;

               if (absEq >= 2 && tag[absEq - 1] == (byte)'r' && (tag[absEq - 2] == (byte)' ' || tag[absEq - 2] == (byte)'\t'))
               {
                    int valStart = absEq + 2;
                    if (valStart >= tag.Length) return false;
                    byte quote = tag[absEq + 1];
                    int valEnd = tag[valStart..].IndexOf(quote);
                    if (valEnd <= 0) return false;

                    ReadOnlySpan<byte> cellRef = tag.Slice(valStart, valEnd);
                    int col = 0;
                    for (int i = 0; i < cellRef.Length; i++)
                    {
                         byte b = cellRef[i];
                         if (b is >= (byte)'A' and <= (byte)'Z')
                              col = (col * 26) + (b - (byte)'A' + 1);
                         else if (b is >= (byte)'a' and <= (byte)'z')
                              col = (col * 26) + (b - (byte)'a' + 1);
                         else
                              break;
                    }
                    colIdx = col > 0 ? col - 1 : 0;
                    return true;
               }
               pos = absEq + 1;
          }
          return false;
     }

     private static byte ExtractCellTypeCode(ReadOnlySpan<byte> tag)
     {
          int pos = 0;
          while (pos < tag.Length)
          {
               int eq = tag[pos..].IndexOf((byte)'=');
               if (eq < 0) break;
               int absEq = pos + eq;

               if (absEq >= 2 && tag[absEq - 1] == (byte)'t' && (tag[absEq - 2] == (byte)' ' || tag[absEq - 2] == (byte)'\t'))
               {
                    int valStart = absEq + 2;
                    if (valStart >= tag.Length) return 0;
                    byte quote = tag[absEq + 1];
                    int valEnd = tag[valStart..].IndexOf(quote);
                    if (valEnd <= 0) return 0;

                    ReadOnlySpan<byte> tVal = tag.Slice(valStart, valEnd);
                    if (tVal.SequenceEqual("s"u8)) return 1;
                    if (tVal.SequenceEqual("inlineStr"u8)) return 2;
                    if (tVal.SequenceEqual("b"u8)) return 3;
                    if (tVal.SequenceEqual("str"u8)) return 4;
                    return 0;
               }
               pos = absEq + 1;
          }
          return 0;
     }

     private static ReadOnlySpan<char> DecodeCellValueToSpan(
             ReadOnlySpan<byte> cellInner,
             byte cellTypeCode,
             string[] sharedStrings,
             ref char[] charScratch,
             out string? cachedString)
     {
          cachedString = null;

          if (cellTypeCode == 2) // inlineStr (simple ou Rich Text <is><r><t>...</t></r>...</is>)
          {
               int pos = 0;
               int totalChars = 0;
               bool hasEntities = false;

               while (pos < cellInner.Length)
               {
                    int tStart = FindTag(cellInner[pos..], "<t"u8);
                    if (tStart < 0) break;
                    pos += tStart;

                    int tOpenEnd = cellInner[pos..].IndexOf((byte)'>');
                    if (tOpenEnd < 0) break;
                    bool isEmptyT = tOpenEnd > 0 && cellInner[pos + tOpenEnd - 1] == (byte)'/';
                    pos += tOpenEnd + 1;
                    if (isEmptyT) continue;

                    int tClose = FindTag(cellInner[pos..], "</t>"u8);
                    if (tClose < 0) break;

                    ReadOnlySpan<byte> utf8Text = cellInner.Slice(pos, tClose);
                    pos += tClose + 4;

                    if (utf8Text.IndexOf((byte)'&') >= 0)
                         hasEntities = true;

                    int maxChars = utf8Text.Length;
                    if (totalChars + maxChars > charScratch.Length)
                    {
                         char[] larger = ArrayPool<char>.Shared.Rent((totalChars + maxChars) * 2);
                         charScratch.AsSpan(0, totalChars).CopyTo(larger);
                         ArrayPool<char>.Shared.Return(charScratch);
                         charScratch = larger;
                    }

                    // Voie rapide SIMD ASCII (32 octets/cycle), repli UTF-8 si caractères accentués
                    if (Ascii.ToUtf16(utf8Text, charScratch.AsSpan(totalChars), out int asciiWritten) == OperationStatus.Done)
                    {
                         totalChars += asciiWritten;
                    }
                    else
                    {
                         totalChars += Encoding.UTF8.GetChars(utf8Text, charScratch.AsSpan(totalChars));
                    }
               }

               if (totalChars == 0)
                    return ReadOnlySpan<char>.Empty;

               if (hasEntities)
               {
                    string unescaped = System.Net.WebUtility.HtmlDecode(new string(charScratch, 0, totalChars));
                    cachedString = unescaped;
                    return unescaped.AsSpan();
               }

               return charScratch.AsSpan(0, totalChars);
          }

          int vStart = FindTag(cellInner, "<v"u8);
          if (vStart < 0) return ReadOnlySpan<char>.Empty;
          int vOpenEnd = cellInner[vStart..].IndexOf((byte)'>');
          if (vOpenEnd < 0) return ReadOnlySpan<char>.Empty;
          int vContentStart = vStart + vOpenEnd + 1;
          int vClose = FindTag(cellInner[vContentStart..], "</v>"u8);
          if (vClose < 0) return ReadOnlySpan<char>.Empty;

          ReadOnlySpan<byte> vUtf8 = cellInner.Slice(vContentStart, vClose);

          if (cellTypeCode == 1) // sharedString ('s')
          {
               if (Utf8Parser.TryParse(vUtf8, out int sstIdx, out _) && (uint)sstIdx < (uint)sharedStrings.Length)
               {
                    string s = sharedStrings[sstIdx];
                    cachedString = s;
                    return s.AsSpan();
               }
               return ReadOnlySpan<char>.Empty;
          }

          if (vUtf8.Length > charScratch.Length)
          {
               ArrayPool<char>.Shared.Return(charScratch);
               charScratch = ArrayPool<char>.Shared.Rent(vUtf8.Length * 2);
          }

          if (Ascii.ToUtf16(vUtf8, charScratch, out int written) == OperationStatus.Done)
               return charScratch.AsSpan(0, written);

          int utf8Len = Encoding.UTF8.GetChars(vUtf8, charScratch);
          return charScratch.AsSpan(0, utf8Len);
     }

     private static bool IsHeaderRowUtf8<T>(
         ReadOnlySpan<byte> rowInner,
         string[] sharedStrings,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options)
     {
          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          try
          {
               int matches = 0;
               int pos = 0;
               while (pos < rowInner.Length)
               {
                    int cIdx = FindTag(rowInner[pos..], "<c"u8);
                    if (cIdx < 0) break;
                    pos += cIdx;
                    int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                    if (cOpenEnd < 0) break;
                    ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                    bool selfClosing = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                    pos += cOpenEnd + 1;
                    if (selfClosing) continue;

                    int cClose = FindTag(rowInner[pos..], "</c>"u8);
                    if (cClose < 0) break;
                    ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                    pos += cClose + 4;

                    byte typeCode = ExtractCellTypeCode(cTag);
                    var chars = DecodeCellValueToSpan(cellInner, typeCode, sharedStrings, ref scratch, out _).Trim();
                    if (chars.IsEmpty) continue;

                    string headerStr = chars.ToString();
                    if (ResolveProperty(headerStr, map, options) >= 0)
                         matches++;
               }
               return matches >= 1 && (matches >= 2 || map.ColumnCount == 1);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }
     }

     private static void ApplyHeaderMappingUtf8<T>(
         ReadOnlySpan<byte> rowInner,
         string[] sharedStrings,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options,
         ref int[] colToProp,
         int[] propToExcelCol,
         ref string?[] excelHeaders)
     {
          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          try
          {
               int fallbackCol = 0;
               int pos = 0;
               while (pos < rowInner.Length)
               {
                    int cIdx = FindTag(rowInner[pos..], "<c"u8);
                    if (cIdx < 0) break;
                    pos += cIdx;
                    int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                    if (cOpenEnd < 0) break;
                    ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                    bool selfClosing = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                    pos += cOpenEnd + 1;

                    int colIdx = TryExtractCellColIndex(cTag, out int parsedCol) ? parsedCol : fallbackCol;
                    fallbackCol = colIdx + 1;
                    if (selfClosing) continue;

                    int cClose = FindTag(rowInner[pos..], "</c>"u8);
                    if (cClose < 0) break;
                    ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                    pos += cClose + 4;

                    if (colIdx >= colToProp.Length)
                         ResizeIntArray(ref colToProp, colIdx + 64, -1);
                    if (colIdx >= excelHeaders.Length)
                         Array.Resize(ref excelHeaders, colIdx + 64);

                    byte typeCode = ExtractCellTypeCode(cTag);
                    var chars = DecodeCellValueToSpan(cellInner, typeCode, sharedStrings, ref scratch, out _).Trim();
                    string trimmedHeader = chars.ToString();
                    excelHeaders[colIdx] = trimmedHeader;

                    if (colToProp[colIdx] >= 0)
                         continue;

                    int propIdx = ResolveProperty(trimmedHeader, map, options);
                    if (propIdx >= 0 && propToExcelCol[propIdx] < 0)
                    {
                         colToProp[colIdx] = propIdx;
                         propToExcelCol[propIdx] = colIdx;
                    }
               }
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }
     }

     private static bool IsRowEmptyUtf8(ReadOnlySpan<byte> rowInner, string[] sharedStrings)
     {
          char[] scratch = ArrayPool<char>.Shared.Rent(128);
          try
          {
               int pos = 0;
               while (pos < rowInner.Length)
               {
                    int cIdx = FindTag(rowInner[pos..], "<c"u8);
                    if (cIdx < 0) break;
                    pos += cIdx;
                    int cOpenEnd = rowInner[pos..].IndexOf((byte)'>');
                    if (cOpenEnd < 0) break;
                    ReadOnlySpan<byte> cTag = rowInner.Slice(pos, cOpenEnd);
                    bool selfClosing = cOpenEnd > 0 && cTag[^1] == (byte)'/';
                    pos += cOpenEnd + 1;
                    if (selfClosing) continue;

                    int cClose = FindTag(rowInner[pos..], "</c>"u8);
                    if (cClose < 0) break;
                    ReadOnlySpan<byte> cellInner = rowInner.Slice(pos, cClose);
                    pos += cClose + 4;

                    byte typeCode = ExtractCellTypeCode(cTag);
                    var chars = DecodeCellValueToSpan(cellInner, typeCode, sharedStrings, ref scratch, out _);
                    if (!chars.IsWhiteSpace())
                         return false;
               }
               return true;
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }
     }

     private static int ResolveProperty<T>(string header, IXlsxTypeMap<T> map, XlsxReadOptions options)
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

     private static void NotifyCorruptedAndThrow(int rowNum, XlsxReadOptions options)
     {
          var err = new XlsxErrorContext(rowNum, -1, string.Empty, null, "XML tronqué ou malformé.", XlsxCellStatus.CorruptedFile);
          options.OnError?.Invoke(in err);
          throw new InvalidDataException($"Le contenu XML de la feuille Excel est corrompu à la ligne {rowNum}.");
     }

     private static void ResizeIntArray(ref int[] array, int newSize, int defaultFill)
     {
          int oldLen = array.Length;
          Array.Resize(ref array, newSize);
          Array.Fill(array, defaultFill, oldLen, newSize - oldLen);
     }
}