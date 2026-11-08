using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using FlashXlsx.Core;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx;

/// <summary>
/// Écrivain multi-onglets en streaming permettant d'ajouter séquentiellement plusieurs feuilles dans un même .xlsx.
/// </summary>
public sealed class XlsxWorkbookWriter : IAsyncDisposable, IDisposable
{
     private readonly ZipArchive _archive;
     private readonly List<string> _sheetNames = [];
     private readonly List<XlsxTableArtifact> _tables = [];
     private XlsxHeaderStyle? _headerStyle;
     private bool _completed;

     public XlsxWorkbookWriter(Stream stream, bool leaveOpen = false)
     {
          _archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen);
     }

     public async ValueTask WriteSheetAsync<T>(
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxWriteOptions<T>();
          var map = XlsxTypeRegistry.GetRequiredMap<T>();
          int sheetNumber = _sheetNames.Count + 1;
          _sheetNames.Add(options.SheetName);
          _headerStyle ??= options.HeaderStyle;

          int colCount = map.ColumnCount;
          bool hasTable = options.ExcelTable is not null && options.WriteHeader && colCount > 0;
          var dropdowns = CollectDropdownRules(map, options);

          var entry = _archive.CreateEntry($"xl/worksheets/sheet{sheetNumber}.xml", CompressionLevel.Fastest);
          await using var entryStream = entry.Open();
          using var writer = new XmlSheetWriter(
              entryStream,
              columnCount: colCount,
              autoFitColumns: options.AutoFitColumns,
              hasHeaderStyle: options.HeaderStyle is not null,
              hasExcelTable: hasTable,
              freezeHeader: options.FreezeHeader,
              dropdowns: dropdowns);

          int rowNumber = 1;
          var writtenHeaders = hasTable ? new List<string>(colCount) : null;

          if (options.WriteHeader)
          {
               writer.StartRow(rowNumber++, isHeader: true);
               for (int c = 0; c < colCount; c++)
               {
                    string propName = map.GetPropertyName(c);
                    string defaultHeader = map.GetDefaultHeaderName(c);
                    string header = options.Translations.TryGetValue(propName, out var translated)
                        ? translated
                        : options.Translations.TryGetValue(defaultHeader, out var translatedAttr)
                            ? translatedAttr
                            : defaultHeader;

                    writtenHeaders?.Add(header);
                    writer.WriteCell(c, XlsxCellWriteKind.InlineString, default, header);
               }
               writer.EndRow();
          }

          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          int rowsWritten = 0;
          try
          {
               foreach (var item in items)
               {
                    ct.ThrowIfCancellationRequested();
                    int currentRow = rowNumber++;
                    writer.StartRow(currentRow);

                    for (int c = 0; c < colCount; c++)
                    {
                         try
                         {
                              if (map.TryFormatCellValue(in item, c, currentRow, scratch, out int written, out var strVal, out var kind, options))
                              {
                                   writer.WriteCell(c, kind, scratch.AsSpan(0, written), strVal);
                              }
                         }
                         catch (Exception ex)
                         {
                              var err = new XlsxErrorContext(
                                  currentRow,
                                  c,
                                  XlsxCoordinateHelper.GetColumnLetter(c),
                                  map.GetDefaultHeaderName(c),
                                  ex.Message,
                                  XlsxCellStatus.WriteError);

                              options.OnError?.Invoke(in err);
                              if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   throw;
                         }
                    }

                    writer.EndRow();

                    rowsWritten++;
                    if (options.Progress is not null && rowsWritten % options.ProgressInterval == 0)
                         options.Progress.Report(rowsWritten);
               }

               if (options.Progress is not null && rowsWritten > 0 && rowsWritten % options.ProgressInterval != 0)
                    options.Progress.Report(rowsWritten);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }

          await writer.CompleteAsync(ct);

          if (hasTable && writtenHeaders is not null)
          {
               RegisterExcelTable(sheetNumber, colCount, writer.RowCount, writtenHeaders, options.ExcelTable!);
          }
     }

     public async ValueTask WriteSheetAsync<T>(
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxWriteOptions<T>();
          var map = XlsxTypeRegistry.GetRequiredMap<T>();
          int sheetNumber = _sheetNames.Count + 1;
          _sheetNames.Add(options.SheetName);
          _headerStyle ??= options.HeaderStyle;

          int colCount = map.ColumnCount;
          bool hasTable = options.ExcelTable is not null && options.WriteHeader && colCount > 0;
          var dropdowns = CollectDropdownRules(map, options);

          var entry = _archive.CreateEntry($"xl/worksheets/sheet{sheetNumber}.xml", CompressionLevel.Fastest);
          await using var entryStream = entry.Open();
          using var writer = new XmlSheetWriter(
              entryStream,
              columnCount: colCount,
              autoFitColumns: options.AutoFitColumns,
              hasHeaderStyle: options.HeaderStyle is not null,
              hasExcelTable: hasTable,
              freezeHeader: options.FreezeHeader,
              dropdowns: dropdowns);

          int rowNumber = 1;
          var writtenHeaders = hasTable ? new List<string>(colCount) : null;

          if (options.WriteHeader)
          {
               writer.StartRow(rowNumber++, isHeader: true);
               for (int c = 0; c < colCount; c++)
               {
                    string propName = map.GetPropertyName(c);
                    string defaultHeader = map.GetDefaultHeaderName(c);
                    string header = options.Translations.TryGetValue(propName, out var translated)
                        ? translated
                        : options.Translations.TryGetValue(defaultHeader, out var translatedAttr)
                            ? translatedAttr
                            : defaultHeader;

                    writtenHeaders?.Add(header);
                    writer.WriteCell(c, XlsxCellWriteKind.InlineString, default, header);
               }
               writer.EndRow();
          }

          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          int rowsWritten = 0;
          try
          {
               await foreach (var item in items.WithCancellation(ct))
               {
                    int currentRow = rowNumber++;
                    writer.StartRow(currentRow);

                    for (int c = 0; c < colCount; c++)
                    {
                         try
                         {
                              if (map.TryFormatCellValue(in item, c, currentRow, scratch, out int written, out var strVal, out var kind, options))
                              {
                                   writer.WriteCell(c, kind, scratch.AsSpan(0, written), strVal);
                              }
                         }
                         catch (Exception ex)
                         {
                              var err = new XlsxErrorContext(
                                  currentRow,
                                  c,
                                  XlsxCoordinateHelper.GetColumnLetter(c),
                                  map.GetDefaultHeaderName(c),
                                  ex.Message,
                                  XlsxCellStatus.WriteError);

                              options.OnError?.Invoke(in err);
                              if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   throw;
                         }
                    }

                    writer.EndRow();

                    rowsWritten++;
                    if (options.Progress is not null && rowsWritten % options.ProgressInterval == 0)
                         options.Progress.Report(rowsWritten);
               }

               if (options.Progress is not null && rowsWritten > 0 && rowsWritten % options.ProgressInterval != 0)
                    options.Progress.Report(rowsWritten);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }

          await writer.CompleteAsync(ct);

          if (hasTable && writtenHeaders is not null)
          {
               RegisterExcelTable(sheetNumber, colCount, writer.RowCount, writtenHeaders, options.ExcelTable!);
          }
     }

     public async ValueTask WriteSheetAsync<T, TContext>(
         IEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          var map = XlsxTypeRegistry.GetRequiredContextMap<T, TContext>();
          int sheetNumber = _sheetNames.Count + 1;
          _sheetNames.Add(options.SheetName);
          _headerStyle ??= options.HeaderStyle;

          int colCount = map.ColumnCount;
          bool hasTable = options.ExcelTable is not null && options.WriteHeader && colCount > 0;
          var dropdowns = CollectDropdownRules(map, options);

          var entry = _archive.CreateEntry($"xl/worksheets/sheet{sheetNumber}.xml", CompressionLevel.Fastest);
          await using var entryStream = entry.Open();
          using var writer = new XmlSheetWriter(
              entryStream,
              columnCount: colCount,
              autoFitColumns: options.AutoFitColumns,
              hasHeaderStyle: options.HeaderStyle is not null,
              hasExcelTable: hasTable,
              freezeHeader: options.FreezeHeader,
              dropdowns: dropdowns);

          int rowNumber = 1;
          TContext context = options.Context;
          var writtenHeaders = hasTable ? new List<string>(colCount) : null;

          if (options.WriteHeader)
          {
               writer.StartRow(rowNumber++, isHeader: true);
               for (int c = 0; c < colCount; c++)
               {
                    string propName = map.GetPropertyName(c);
                    string defaultHeader = map.GetDefaultHeaderName(c);
                    string header = options.Translations.TryGetValue(propName, out var translated)
                        ? translated
                        : options.Translations.TryGetValue(defaultHeader, out var translatedAttr)
                            ? translatedAttr
                            : defaultHeader;

                    writtenHeaders?.Add(header);
                    writer.WriteCell(c, XlsxCellWriteKind.InlineString, default, header);
               }
               writer.EndRow();
          }

          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          int rowsWritten = 0;
          try
          {
               foreach (var item in items)
               {
                    ct.ThrowIfCancellationRequested();
                    int currentRow = rowNumber++;
                    writer.StartRow(currentRow);

                    for (int c = 0; c < colCount; c++)
                    {
                         try
                         {
                              if (map.TryFormatCellValue(in item, c, currentRow, scratch, out int written, out var strVal, out var kind, options))
                              {
                                   writer.WriteCell(c, kind, scratch.AsSpan(0, written), strVal);
                              }
                         }
                         catch (Exception ex)
                         {
                              var err = new XlsxErrorContext(
                                  currentRow,
                                  c,
                                  XlsxCoordinateHelper.GetColumnLetter(c),
                                  map.GetDefaultHeaderName(c),
                                  ex.Message,
                                  XlsxCellStatus.WriteError);

                              options.OnError?.Invoke(in err);
                              if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   throw;
                         }
                    }

                    writer.EndRow();

                    rowsWritten++;
                    if (options.Progress is not null && rowsWritten % options.ProgressInterval == 0)
                         options.Progress.Report(rowsWritten);
               }

               if (options.Progress is not null && rowsWritten > 0 && rowsWritten % options.ProgressInterval != 0)
                    options.Progress.Report(rowsWritten);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }

          await writer.CompleteAsync(ct);

          if (hasTable && writtenHeaders is not null)
          {
               RegisterExcelTable(sheetNumber, colCount, writer.RowCount, writtenHeaders, options.ExcelTable!);
          }
     }

     public async ValueTask WriteSheetAsync<T, TContext>(
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          var map = XlsxTypeRegistry.GetRequiredContextMap<T, TContext>();
          int sheetNumber = _sheetNames.Count + 1;
          _sheetNames.Add(options.SheetName);
          _headerStyle ??= options.HeaderStyle;

          int colCount = map.ColumnCount;
          bool hasTable = options.ExcelTable is not null && options.WriteHeader && colCount > 0;
          var dropdowns = CollectDropdownRules(map, options);

          var entry = _archive.CreateEntry($"xl/worksheets/sheet{sheetNumber}.xml", CompressionLevel.Fastest);
          await using var entryStream = entry.Open();
          using var writer = new XmlSheetWriter(
              entryStream,
              columnCount: colCount,
              autoFitColumns: options.AutoFitColumns,
              hasHeaderStyle: options.HeaderStyle is not null,
              hasExcelTable: hasTable,
              freezeHeader: options.FreezeHeader,
              dropdowns: dropdowns);

          int rowNumber = 1;
          TContext context = options.Context;
          var writtenHeaders = hasTable ? new List<string>(colCount) : null;

          if (options.WriteHeader)
          {
               writer.StartRow(rowNumber++, isHeader: true);
               for (int c = 0; c < colCount; c++)
               {
                    string propName = map.GetPropertyName(c);
                    string defaultHeader = map.GetDefaultHeaderName(c);
                    string header = options.Translations.TryGetValue(propName, out var translated)
                        ? translated
                        : options.Translations.TryGetValue(defaultHeader, out var translatedAttr)
                            ? translatedAttr
                            : defaultHeader;

                    writtenHeaders?.Add(header);
                    writer.WriteCell(c, XlsxCellWriteKind.InlineString, default, header);
               }
               writer.EndRow();
          }

          char[] scratch = ArrayPool<char>.Shared.Rent(256);
          int rowsWritten = 0;
          try
          {
               await foreach (var item in items.WithCancellation(ct))
               {
                    int currentRow = rowNumber++;
                    writer.StartRow(currentRow);

                    for (int c = 0; c < colCount; c++)
                    {
                         try
                         {
                              if (map.TryFormatCellValue(in item, c, currentRow, scratch, out int written, out var strVal, out var kind, options))
                              {
                                   writer.WriteCell(c, kind, scratch.AsSpan(0, written), strVal);
                              }
                         }
                         catch (Exception ex)
                         {
                              var err = new XlsxErrorContext(
                                  currentRow,
                                  c,
                                  XlsxCoordinateHelper.GetColumnLetter(c),
                                  map.GetDefaultHeaderName(c),
                                  ex.Message,
                                  XlsxCellStatus.WriteError);

                              options.OnError?.Invoke(in err);
                              if (options.ErrorMode == XlsxErrorMode.StopOnFirstError)
                                   throw;
                         }
                    }

                    writer.EndRow();

                    rowsWritten++;
                    if (options.Progress is not null && rowsWritten % options.ProgressInterval == 0)
                         options.Progress.Report(rowsWritten);
               }

               if (options.Progress is not null && rowsWritten > 0 && rowsWritten % options.ProgressInterval != 0)
                    options.Progress.Report(rowsWritten);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }

          await writer.CompleteAsync(ct);

          if (hasTable && writtenHeaders is not null)
          {
               RegisterExcelTable(sheetNumber, colCount, writer.RowCount, writtenHeaders, options.ExcelTable!);
          }
     }

     public async ValueTask WriteDictionarySheetAsync(
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxWriteOptions();
          int sheetNumber = _sheetNames.Count + 1;
          _sheetNames.Add(options.SheetName);
          _headerStyle ??= options.HeaderStyle;

          var entry = _archive.CreateEntry($"xl/worksheets/sheet{sheetNumber}.xml", CompressionLevel.Fastest);
          await using var entryStream = entry.Open();
          using var writer = new XmlSheetWriter(
              entryStream,
              hasHeaderStyle: options.HeaderStyle is not null,
              freezeHeader: options.FreezeHeader);

          List<string>? columnKeys = null;
          int rowNumber = 1;
          char[] scratch = ArrayPool<char>.Shared.Rent(128);
          int rowsWritten = 0;

          try
          {
               foreach (var row in rows)
               {
                    ct.ThrowIfCancellationRequested();

                    if (columnKeys is null)
                    {
                         columnKeys = [.. row.Keys];
                         if (options.WriteHeader)
                         {
                              writer.StartRow(rowNumber++, isHeader: true);
                              for (int c = 0; c < columnKeys.Count; c++)
                              {
                                   string key = columnKeys[c];
                                   string header = options.Translations.TryGetValue(key, out var tr) ? tr : key;
                                   writer.WriteCell(c, XlsxCellWriteKind.InlineString, default, header);
                              }
                              writer.EndRow();
                         }
                    }

                    writer.StartRow(rowNumber++);
                    for (int c = 0; c < columnKeys.Count; c++)
                    {
                         if (row.TryGetValue(columnKeys[c], out var val) && !val.IsEmpty)
                         {
                              WriteDynamicCell(writer, c, val, scratch, options.SecurityPolicy);
                         }
                    }
                    writer.EndRow();

                    rowsWritten++;
                    if (options.Progress is not null && rowsWritten % options.ProgressInterval == 0)
                         options.Progress.Report(rowsWritten);
               }

               if (options.Progress is not null && rowsWritten > 0 && rowsWritten % options.ProgressInterval != 0)
                    options.Progress.Report(rowsWritten);
          }
          finally
          {
               ArrayPool<char>.Shared.Return(scratch);
          }

          await writer.CompleteAsync(ct);
     }

     private static List<XlsxDropdownRule>? CollectDropdownRules<T>(IXlsxTypeMap<T> map, XlsxWriteOptions options)
     {
          if (!options.DropdownValidation || map.ColumnCount <= 0)
               return null;

          var rules = new List<XlsxDropdownRule>();
          for (int c = 0; c < map.ColumnCount; c++)
          {
               var vals = map.GetDropdownValues(c);
               if (vals is { Length: > 0 })
                    rules.Add(new XlsxDropdownRule(c, vals));
          }
          return rules;
     }

     private void RegisterExcelTable(
         int sheetNumber,
         int colCount,
         int rowCount,
         IReadOnlyList<string> headers,
         XlsxTableOptions tableOptions)
     {
          int tableId = _tables.Count + 1;
          string tableName = tableOptions.TableName ?? $"Table{tableId}";
          string endColLetter = XlsxCoordinateHelper.GetColumnLetter(colCount - 1);
          int lastRow = Math.Max(2, rowCount);
          string refRange = $"A1:{endColLetter}{lastRow}";
          _tables.Add(new XlsxTableArtifact(tableId, sheetNumber, tableName, refRange, headers, tableOptions));
     }

     private static void WriteDynamicCell(XmlSheetWriter writer, int col, XlsxValue val, Span<char> scratch, XlsxSecurityPolicy securityPolicy)
     {
          switch (val.Type)
          {
               case XlsxDataType.String:
                    string? rawStr = val.AsString();
                    string? safeStr = securityPolicy == XlsxSecurityPolicy.Sanitize
                        ? XlsxSecuritySanitizer.SanitizeForExport(rawStr, isExplicitFormula: false)
                        : rawStr;
                    writer.WriteCell(col, XlsxCellWriteKind.InlineString, default, safeStr ?? string.Empty);
                    break;
               case XlsxDataType.Int64:
                    val.AsInt64().TryFormat(scratch, out int iLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.Number, scratch[..iLen], null);
                    break;
               case XlsxDataType.Double:
                    val.AsDouble().TryFormat(scratch, out int dLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.Number, scratch[..dLen], null);
                    break;
               case XlsxDataType.Decimal:
                    val.AsDecimal().TryFormat(scratch, out int mLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.Number, scratch[..mLen], null);
                    break;
               case XlsxDataType.Boolean:
                    writer.WriteCell(col, XlsxCellWriteKind.Boolean, val.AsBoolean() ? "1" : "0", null);
                    break;
               case XlsxDataType.DateTime:
                    val.AsDateTime().ToOADate().TryFormat(scratch, out int dtLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.DateTime, scratch[..dtLen], null);
                    break;
               case XlsxDataType.DateOnly:
                    val.AsDateOnly().ToDateTime(TimeOnly.MinValue).ToOADate().TryFormat(scratch, out int doLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.Date, scratch[..doLen], null);
                    break;
               case XlsxDataType.TimeOnly:
                    val.AsTimeOnly().ToTimeSpan().TotalDays.TryFormat(scratch, out int toLen, provider: CultureInfo.InvariantCulture);
                    writer.WriteCell(col, XlsxCellWriteKind.Time, scratch[..toLen], null);
                    break;
               case XlsxDataType.Guid:
                    val.AsGuid().TryFormat(scratch, out int gLen);
                    writer.WriteCell(col, XlsxCellWriteKind.InlineString, scratch[..gLen], null);
                    break;
          }
     }

     public async ValueTask CompleteAsync(CancellationToken ct = default)
     {
          if (_completed) return;
          _completed = true;

          if (_sheetNames.Count == 0)
               _sheetNames.Add("Sheet1");

          await OpenXmlPackageHelper.WritePackageMetadataAsync(_archive, _sheetNames, _tables, _headerStyle, ct);
     }

     public async ValueTask DisposeAsync()
     {
          await CompleteAsync();
          _archive.Dispose();
     }

     public void Dispose()
     {
          if (!_completed)
          {
               CompleteAsync().AsTask().GetAwaiter().GetResult();
          }
          _archive.Dispose();
     }
}