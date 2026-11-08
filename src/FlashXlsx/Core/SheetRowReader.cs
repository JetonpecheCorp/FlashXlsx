using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace FlashXlsx.Core;

internal static class SheetRowReader
{
     private static readonly XmlReaderSettings Settings = new()
     {
          Async = true,
          IgnoreWhitespace = true,
          IgnoreComments = true,
          IgnoreProcessingInstructions = true,
          DtdProcessing = DtdProcessing.Prohibit
     };

     public static async IAsyncEnumerable<T> ReadTypedAsync<T>(
         ZipArchiveEntry sheetEntry,
         string[] sharedStrings,
         IXlsxTypeMap<T> map,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var stream = sheetEntry.Open();
          using var reader = XmlReader.Create(stream, Settings);

          int propCount = map.ColumnCount;
          int[] excelColToProp = ArrayPool<int>.Shared.Rent(128);
          int[] propToExcelCol = ArrayPool<int>.Shared.Rent(propCount);
          string?[] excelHeaders = ArrayPool<string?>.Shared.Rent(128);

          Array.Fill(excelColToProp, -1);
          Array.Fill(propToExcelCol, -1);
          Array.Fill(excelHeaders, null);

          for (int i = 0; i < propCount; i++)
          {
               int fixedCol = map.GetFixedColumnIndex(i);
               if (fixedCol >= 0)
               {
                    if (fixedCol < excelColToProp.Length)
                         excelColToProp[fixedCol] = i;
                    propToExcelCol[i] = fixedCol;
               }
               else if (!options.HasHeaderRow && !options.AutoDetectHeaderRow)
               {
                    if (i < excelColToProp.Length)
                         excelColToProp[i] = i;
                    propToExcelCol[i] = i;
               }
          }

          Dictionary<string, int>? translationLookup = BuildTranslationLookup(map, options);

          try
          {
               int currentRow = 0;
               bool inRow = false;
               bool hasNonEmptyCell = false;
               bool isHeaderRow = false;
               bool headerFound = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
               int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
               int matchedHeadersInCandidateRow = 0;
               int minHeadersRequired = propCount >= 2 ? 2 : 1;

               T currentInstance = default!;
               int rowsYielded = 0;

               while (await TryReadNextNodeAsync(reader, currentRow, options))
               {
                    ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                         if (reader.LocalName == "row")
                         {
                              var rAttr = reader.GetAttribute("r");
                              currentRow = int.TryParse(rAttr, out int r) ? r : currentRow + 1;
                              inRow = true;
                              hasNonEmptyCell = false;

                              if (!headerFound && options.AutoDetectHeaderRow)
                              {
                                   if (currentRow <= options.AutoDetectMaxRows)
                                   {
                                        isHeaderRow = true;
                                        matchedHeadersInCandidateRow = 0;
                                        Array.Fill(excelColToProp, -1);
                                        Array.Fill(propToExcelCol, -1);
                                        Array.Fill(excelHeaders, null);
                                   }
                                   else
                                   {
                                        // Aucune ligne d'en-tête trouvée dans la plage : repli sur HeaderRowNumber
                                        headerFound = true;
                                        resolvedHeaderRowNumber = options.HeaderRowNumber;
                                        isHeaderRow = false;
                                   }
                              }
                              else if (!headerFound)
                              {
                                   isHeaderRow = currentRow == options.HeaderRowNumber;
                              }
                              else
                              {
                                   isHeaderRow = false;
                              }

                              if (headerFound && !isHeaderRow && currentRow > resolvedHeaderRowNumber)
                              {
                                   currentInstance = map.CreateInstance();
                              }
                         }
                         else if (reader.LocalName == "c" && inRow)
                         {
                              var cellRef = reader.GetAttribute("r");
                              int colIndex = cellRef is not null ? XlsxCoordinateHelper.ParseColumnIndex(cellRef) : 0;
                              var cellType = reader.GetAttribute("t");

                              if (colIndex >= excelColToProp.Length)
                              {
                                   int newSize = Math.Max(colIndex + 32, excelColToProp.Length * 2);
                                   var newColToProp = ArrayPool<int>.Shared.Rent(newSize);
                                   Array.Fill(newColToProp, -1);
                                   excelColToProp.AsSpan().CopyTo(newColToProp);
                                   ArrayPool<int>.Shared.Return(excelColToProp);
                                   excelColToProp = newColToProp;

                                   var newHeaders = ArrayPool<string?>.Shared.Rent(newSize);
                                   Array.Fill(newHeaders, null);
                                   excelHeaders.AsSpan().CopyTo(newHeaders);
                                   ArrayPool<string?>.Shared.Return(excelHeaders);
                                   excelHeaders = newHeaders;
                              }

                              var (rawVal, isNumeric) = await ReadCellContentAsync(reader, cellType, sharedStrings);
                              ReadOnlySpan<char> span = options.TrimStrings ? rawVal.AsSpan().Trim() : rawVal.AsSpan();

                              if (isHeaderRow)
                              {
                                   string headerText = span.ToString();
                                   excelHeaders[colIndex] = headerText;

                                   int propIdx = (translationLookup is not null && translationLookup.TryGetValue(headerText, out int mappedIdx))
                                       ? mappedIdx
                                       : map.ResolvePropertyIndex(span);

                                   if (propIdx >= 0)
                                   {
                                        excelColToProp[colIndex] = propIdx;
                                        propToExcelCol[propIdx] = colIndex;
                                        matchedHeadersInCandidateRow++;
                                   }
                              }
                              else if (headerFound && currentRow > resolvedHeaderRowNumber)
                              {
                                   int propIdx = colIndex < excelColToProp.Length ? excelColToProp[colIndex] : -1;
                                   if (propIdx >= 0)
                                   {
                                        if (!span.IsEmpty)
                                             hasNonEmptyCell = true;

                                        var status = map.TrySetCellValue(ref currentInstance, propIdx, span, isNumeric, options);
                                        if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                                        {
                                             var err = new XlsxErrorContext(
                                                 currentRow,
                                                 colIndex,
                                                 XlsxCoordinateHelper.GetColumnLetter(colIndex),
                                                 excelHeaders[colIndex] ?? map.GetDefaultHeaderName(propIdx),
                                                 rawVal,
                                                 status);

                                             options.OnError?.Invoke(in err);
                                             if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                                  yield break;
                                        }
                                   }
                              }
                         }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row" && inRow)
                    {
                         inRow = false;

                         if (isHeaderRow)
                         {
                              if (!options.AutoDetectHeaderRow || matchedHeadersInCandidateRow >= minHeadersRequired)
                              {
                                   headerFound = true;
                                   resolvedHeaderRowNumber = currentRow;

                                   // Réappliquer les éventuels [XlsxColumnLetter] fixes non écrasés
                                   for (int i = 0; i < propCount; i++)
                                   {
                                        int fixedCol = map.GetFixedColumnIndex(i);
                                        if (fixedCol >= 0)
                                        {
                                             if (fixedCol < excelColToProp.Length)
                                                  excelColToProp[fixedCol] = i;
                                             propToExcelCol[i] = fixedCol;
                                        }
                                   }
                              }
                         }
                         else if (headerFound && currentRow > resolvedHeaderRowNumber)
                         {
                              if (options.SkipEmptyRows && !hasNonEmptyCell)
                                   continue;

                              bool isValid = map.ValidateRow(
                                  ref currentInstance,
                                  currentRow,
                                  propToExcelCol.AsSpan(0, propCount),
                                  excelHeaders,
                                  options);

                              if (!isValid && options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   yield break;

                              if (isValid || options.ErrorMode == XlsxErrorMode.ReportAndContinue)
                              {
                                   rowsYielded++;
                                   if (options.Progress is not null && rowsYielded % options.ProgressInterval == 0)
                                        options.Progress.Report(rowsYielded);

                                   yield return currentInstance;
                              }
                         }
                    }
               }

               if (options.Progress is not null && rowsYielded > 0 && rowsYielded % options.ProgressInterval != 0)
                    options.Progress.Report(rowsYielded);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(excelColToProp);
               ArrayPool<int>.Shared.Return(propToExcelCol);
               ArrayPool<string?>.Shared.Return(excelHeaders);
          }
     }

     public static async IAsyncEnumerable<T> ReadTypedAsync<T, TContext>(
         ZipArchiveEntry sheetEntry,
         string[] sharedStrings,
         IXlsxTypeMap<T, TContext> map,
         XlsxReadOptions<T, TContext> options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var stream = sheetEntry.Open();
          using var reader = XmlReader.Create(stream, Settings);

          int propCount = map.ColumnCount;
          int[] excelColToProp = ArrayPool<int>.Shared.Rent(128);
          int[] propToExcelCol = ArrayPool<int>.Shared.Rent(propCount);
          string?[] excelHeaders = ArrayPool<string?>.Shared.Rent(128);

          Array.Fill(excelColToProp, -1);
          Array.Fill(propToExcelCol, -1);
          Array.Fill(excelHeaders, null);

          for (int i = 0; i < propCount; i++)
          {
               int fixedCol = map.GetFixedColumnIndex(i);
               if (fixedCol >= 0)
               {
                    if (fixedCol < excelColToProp.Length)
                         excelColToProp[fixedCol] = i;
                    propToExcelCol[i] = fixedCol;
               }
               else if (!options.HasHeaderRow && !options.AutoDetectHeaderRow)
               {
                    if (i < excelColToProp.Length)
                         excelColToProp[i] = i;
                    propToExcelCol[i] = i;
               }
          }

          Dictionary<string, int>? translationLookup = BuildTranslationLookup(map, options);

          try
          {
               int currentRow = 0;
               bool inRow = false;
               bool hasNonEmptyCell = false;
               bool isHeaderRow = false;
               bool headerFound = !options.HasHeaderRow && !options.AutoDetectHeaderRow;
               int resolvedHeaderRowNumber = options.HasHeaderRow ? options.HeaderRowNumber : 0;
               int matchedHeadersInCandidateRow = 0;
               int minHeadersRequired = propCount >= 2 ? 2 : 1;

               T currentInstance = default!;
               TContext context = options.Context;
               int rowsYielded = 0;

               while (await TryReadNextNodeAsync(reader, currentRow, options))
               {
                    ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                         if (reader.LocalName == "row")
                         {
                              var rAttr = reader.GetAttribute("r");
                              currentRow = int.TryParse(rAttr, out int r) ? r : currentRow + 1;
                              inRow = true;
                              hasNonEmptyCell = false;

                              if (!headerFound && options.AutoDetectHeaderRow)
                              {
                                   if (currentRow <= options.AutoDetectMaxRows)
                                   {
                                        isHeaderRow = true;
                                        matchedHeadersInCandidateRow = 0;
                                        Array.Fill(excelColToProp, -1);
                                        Array.Fill(propToExcelCol, -1);
                                        Array.Fill(excelHeaders, null);
                                   }
                                   else
                                   {
                                        headerFound = true;
                                        resolvedHeaderRowNumber = options.HeaderRowNumber;
                                        isHeaderRow = false;
                                   }
                              }
                              else if (!headerFound)
                              {
                                   isHeaderRow = currentRow == options.HeaderRowNumber;
                              }
                              else
                              {
                                   isHeaderRow = false;
                              }

                              if (headerFound && !isHeaderRow && currentRow > resolvedHeaderRowNumber)
                              {
                                   currentInstance = map.CreateInstance();
                              }
                         }
                         else if (reader.LocalName == "c" && inRow)
                         {
                              var cellRef = reader.GetAttribute("r");
                              int colIndex = cellRef is not null ? XlsxCoordinateHelper.ParseColumnIndex(cellRef) : 0;
                              var cellType = reader.GetAttribute("t");

                              if (colIndex >= excelColToProp.Length)
                              {
                                   int newSize = Math.Max(colIndex + 32, excelColToProp.Length * 2);
                                   var newColToProp = ArrayPool<int>.Shared.Rent(newSize);
                                   Array.Fill(newColToProp, -1);
                                   excelColToProp.AsSpan().CopyTo(newColToProp);
                                   ArrayPool<int>.Shared.Return(excelColToProp);
                                   excelColToProp = newColToProp;

                                   var newHeaders = ArrayPool<string?>.Shared.Rent(newSize);
                                   Array.Fill(newHeaders, null);
                                   excelHeaders.AsSpan().CopyTo(newHeaders);
                                   ArrayPool<string?>.Shared.Return(excelHeaders);
                                   excelHeaders = newHeaders;
                              }

                              var (rawVal, isNumeric) = await ReadCellContentAsync(reader, cellType, sharedStrings);
                              ReadOnlySpan<char> span = options.TrimStrings ? rawVal.AsSpan().Trim() : rawVal.AsSpan();

                              if (isHeaderRow)
                              {
                                   string headerText = span.ToString();
                                   excelHeaders[colIndex] = headerText;

                                   int propIdx = (translationLookup is not null && translationLookup.TryGetValue(headerText, out int mappedIdx))
                                       ? mappedIdx
                                       : map.ResolvePropertyIndex(span);

                                   if (propIdx >= 0)
                                   {
                                        excelColToProp[colIndex] = propIdx;
                                        propToExcelCol[propIdx] = colIndex;
                                        matchedHeadersInCandidateRow++;
                                   }
                              }
                              else if (headerFound && currentRow > resolvedHeaderRowNumber)
                              {
                                   int propIdx = colIndex < excelColToProp.Length ? excelColToProp[colIndex] : -1;
                                   if (propIdx >= 0)
                                   {
                                        if (!span.IsEmpty)
                                             hasNonEmptyCell = true;

                                        var status = map.TrySetCellValue(ref currentInstance, propIdx, span, isNumeric, options);
                                        if (status != XlsxCellStatus.Success && status != XlsxCellStatus.Empty)
                                        {
                                             var err = new XlsxErrorContext(
                                                 currentRow,
                                                 colIndex,
                                                 XlsxCoordinateHelper.GetColumnLetter(colIndex),
                                                 excelHeaders[colIndex] ?? map.GetDefaultHeaderName(propIdx),
                                                 rawVal,
                                                 status);

                                             options.OnError?.Invoke(in err);
                                             if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                                  yield break;
                                        }
                                   }
                              }
                         }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row" && inRow)
                    {
                         inRow = false;

                         if (isHeaderRow)
                         {
                              if (!options.AutoDetectHeaderRow || matchedHeadersInCandidateRow >= minHeadersRequired)
                              {
                                   headerFound = true;
                                   resolvedHeaderRowNumber = currentRow;

                                   for (int i = 0; i < propCount; i++)
                                   {
                                        int fixedCol = map.GetFixedColumnIndex(i);
                                        if (fixedCol >= 0)
                                        {
                                             if (fixedCol < excelColToProp.Length)
                                                  excelColToProp[fixedCol] = i;
                                             propToExcelCol[i] = fixedCol;
                                        }
                                   }
                              }
                         }
                         else if (headerFound && currentRow > resolvedHeaderRowNumber)
                         {
                              if (options.SkipEmptyRows && !hasNonEmptyCell)
                                   continue;

                              bool isValid = map.ValidateRow(
                                  ref currentInstance,
                                  in context,
                                  currentRow,
                                  propToExcelCol.AsSpan(0, propCount),
                                  excelHeaders,
                                  options);

                              if (!isValid && options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   yield break;

                              if (isValid || options.ErrorMode == XlsxErrorMode.ReportAndContinue)
                              {
                                   rowsYielded++;
                                   if (options.Progress is not null && rowsYielded % options.ProgressInterval == 0)
                                        options.Progress.Report(rowsYielded);

                                   yield return currentInstance;
                              }
                         }
                    }
               }

               if (options.Progress is not null && rowsYielded > 0 && rowsYielded % options.ProgressInterval != 0)
                    options.Progress.Report(rowsYielded);
          }
          finally
          {
               ArrayPool<int>.Shared.Return(excelColToProp);
               ArrayPool<int>.Shared.Return(propToExcelCol);
               ArrayPool<string?>.Shared.Return(excelHeaders);
          }
     }

     public static async IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadDictionaryAsync(
         ZipArchiveEntry sheetEntry,
         string[] sharedStrings,
         XlsxReadOptions options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var stream = sheetEntry.Open();
          using var reader = XmlReader.Create(stream, Settings);

          string?[] colToKey = ArrayPool<string?>.Shared.Rent(128);
          Array.Fill(colToKey, null);

          Dictionary<string, string>? reverseTranslations = null;
          if (options.Translations.Count > 0)
          {
               reverseTranslations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
               foreach (var kvp in options.Translations)
               {
                    foreach (var alias in kvp.Value)
                         reverseTranslations[alias] = kvp.Key;
               }
          }

          try
          {
               int currentRow = 0;
               bool inRow = false;
               bool hasNonEmptyCell = false;
               bool isHeaderRow = false;
               Dictionary<string, XlsxValue>? currentDict = null;
               int rowsYielded = 0;

               while (await TryReadNextNodeAsync(reader, currentRow, options))
               {
                    ct.ThrowIfCancellationRequested();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                         if (reader.LocalName == "row")
                         {
                              var rAttr = reader.GetAttribute("r");
                              currentRow = int.TryParse(rAttr, out int r) ? r : currentRow + 1;
                              inRow = true;
                              hasNonEmptyCell = false;
                              isHeaderRow = options.HasHeaderRow && currentRow == options.HeaderRowNumber;

                              if (!isHeaderRow && (!options.HasHeaderRow || currentRow > options.HeaderRowNumber))
                              {
                                   currentDict = new Dictionary<string, XlsxValue>(StringComparer.OrdinalIgnoreCase);
                              }
                         }
                         else if (reader.LocalName == "c" && inRow)
                         {
                              var cellRef = reader.GetAttribute("r");
                              int colIndex = cellRef is not null ? XlsxCoordinateHelper.ParseColumnIndex(cellRef) : 0;
                              var cellType = reader.GetAttribute("t");

                              if (colIndex >= colToKey.Length)
                              {
                                   int newSize = Math.Max(colIndex + 32, colToKey.Length * 2);
                                   var newArr = ArrayPool<string?>.Shared.Rent(newSize);
                                   Array.Fill(newArr, null);
                                   colToKey.AsSpan().CopyTo(newArr);
                                   ArrayPool<string?>.Shared.Return(colToKey);
                                   colToKey = newArr;
                              }

                              var (rawVal, isNumeric) = await ReadCellContentAsync(reader, cellType, sharedStrings);
                              ReadOnlySpan<char> span = options.TrimStrings ? rawVal.AsSpan().Trim() : rawVal.AsSpan();

                              if (isHeaderRow)
                              {
                                   string header = span.ToString();
                                   colToKey[colIndex] = (reverseTranslations is not null && reverseTranslations.TryGetValue(header, out var pivot))
                                       ? pivot
                                       : header;
                              }
                              else if (currentDict is not null)
                              {
                                   if (!span.IsEmpty)
                                        hasNonEmptyCell = true;

                                   string key = colToKey[colIndex] ?? XlsxCoordinateHelper.GetColumnLetter(colIndex);
                                   currentDict[key] = ParseDynamicValue(span, cellType, isNumeric, options);
                              }
                         }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row" && inRow)
                    {
                         inRow = false;
                         if (currentDict is not null)
                         {
                              if (!options.SkipEmptyRows || hasNonEmptyCell)
                              {
                                   rowsYielded++;
                                   if (options.Progress is not null && rowsYielded % options.ProgressInterval == 0)
                                        options.Progress.Report(rowsYielded);

                                   yield return currentDict;
                              }
                              currentDict = null;
                         }
                    }
               }

               if (options.Progress is not null && rowsYielded > 0 && rowsYielded % options.ProgressInterval != 0)
                    options.Progress.Report(rowsYielded);
          }
          finally
          {
               ArrayPool<string?>.Shared.Return(colToKey);
          }
     }

     static async ValueTask<bool> TryReadNextNodeAsync(XmlReader xmlReader, int rowNum, XlsxReadOptions opts)
     {
          try
          {
               return await xmlReader.ReadAsync();
          }
          catch (XmlException ex)
          {
               var err = new XlsxErrorContext(rowNum, -1, string.Empty, null, ex.Message, XlsxCellStatus.CorruptedFile);
               opts.OnError?.Invoke(in err);
               throw new InvalidDataException($"Le contenu XML de la feuille Excel est corrompu à la ligne {rowNum}.", ex);
          }
     }

     private static XlsxValue ParseDynamicValue(ReadOnlySpan<char> span, string? cellType, bool isNumeric, XlsxReadOptions options)
     {
          if (span.IsEmpty) return XlsxValue.Empty;

          if (cellType == "b")
          {
               return span.Equals("1", StringComparison.Ordinal) || span.Equals("true", StringComparison.OrdinalIgnoreCase);
          }

          if (isNumeric)
          {
               if (long.TryParse(span, NumberStyles.Integer, CultureInfo.InvariantCulture, out long l))
                    return l;
               if (double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                    return d;
          }

          return span.ToString();
     }

     private static async ValueTask<(string Value, bool IsNumeric)> ReadCellContentAsync(
         XmlReader reader,
         string? cellType,
         string[] sharedStrings)
     {
          if (reader.IsEmptyElement)
               return (string.Empty, false);

          string raw = string.Empty;
          StringBuilder? richTextBuilder = null;
          bool advance = true;

          while (!advance || await reader.ReadAsync())
          {
               advance = true;

               if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
                    break;

               if (reader.NodeType == XmlNodeType.Element)
               {
                    if (reader.LocalName == "f")
                    {
                         await reader.SkipAsync();
                         advance = false;
                    }
                    else if (reader.LocalName == "v")
                    {
                         raw = await reader.ReadElementContentAsStringAsync();
                         advance = false;
                    }
                    else if (reader.LocalName == "t")
                    {
                         string part = await reader.ReadElementContentAsStringAsync();
                         advance = false;

                         if (richTextBuilder is not null)
                         {
                              richTextBuilder.Append(part);
                         }
                         else if (raw.Length > 0)
                         {
                              richTextBuilder = new StringBuilder(raw.Length + part.Length);
                              richTextBuilder.Append(raw).Append(part);
                         }
                         else
                         {
                              raw = part;
                         }
                    }
               }
          }

          if (richTextBuilder is not null)
               raw = richTextBuilder.ToString();

          if (cellType == "s")
          {
               if (int.TryParse(raw, out int ssIdx) && (uint)ssIdx < (uint)sharedStrings.Length)
                    return (sharedStrings[ssIdx], false);
               return (string.Empty, false);
          }

          bool isNumeric = cellType is null or "n";
          return (raw, isNumeric);
     }

     private static Dictionary<string, int>? BuildTranslationLookup<T>(IXlsxTypeMap<T> map, XlsxReadOptions options)
     {
          if (options.Translations.Count == 0)
               return null;

          var lookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
          foreach (var kvp in options.Translations)
          {
               int propIdx = map.ResolvePropertyIndex(kvp.Key.AsSpan());
               if (propIdx >= 0)
               {
                    foreach (var excelHeader in kvp.Value)
                    {
                         lookup[excelHeader] = propIdx;
                    }
               }
          }
          return lookup;
     }
}