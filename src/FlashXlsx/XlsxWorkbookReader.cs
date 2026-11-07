using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using FlashXlsx.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;

namespace FlashXlsx;

/// <summary>
/// Lecteur multi-onglets en streaming permettant de découvrir et lire chaque onglet indépendamment.
/// </summary>
public sealed class XlsxWorkbookReader : IDisposable
{
     private readonly ZipArchive _archive;
     private readonly string[] _sharedStrings;

     public IReadOnlyList<XlsxSheetInfo> Sheets { get; }

     public XlsxWorkbookReader(Stream stream, bool leaveOpen = false)
     {
          _archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen);
          _sharedStrings = SharedStringsReader.Read(_archive);
          Sheets = WorkbookMetadataReader.ReadSheets(_archive);
     }

     public IAsyncEnumerable<T> ReadSheetAsync<T>(string sheetName, XlsxReadOptions<T>? options = null, CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions<T>();
          var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
          if (sheet == default && Sheets.All(s => !string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)))
          {
               NotifySheetNotFound(sheetName, options);
               return AsyncEnumerable.Empty<T>();
          }

          return ReadSheetInternalAsync<T>(sheet, options, ct);
     }

     public IAsyncEnumerable<T> ReadSheetAsync<T>(int sheetIndex, XlsxReadOptions<T>? options = null, CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions<T>();
          if ((uint)sheetIndex >= (uint)Sheets.Count)
          {
               NotifySheetNotFound($"Index:{sheetIndex}", options);
               return AsyncEnumerable.Empty<T>();
          }

          return ReadSheetInternalAsync<T>(Sheets[sheetIndex], options, ct);
     }

     public IAsyncEnumerable<T> ReadSheetAsync<T, TContext>(
        string sheetName,
        XlsxReadOptions<T, TContext> options,
        CancellationToken ct = default)
     {
          var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
          if (Sheets.All(s => !string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)))
          {
               NotifySheetNotFound(sheetName, options);
               return AsyncEnumerable.Empty<T>();
          }

          return ReadSheetInternalAsync<T, TContext>(sheet, options, ct);
     }

     public IAsyncEnumerable<T> ReadSheetAsync<T, TContext>(
         int sheetIndex,
         XlsxReadOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          if ((uint)sheetIndex >= (uint)Sheets.Count)
          {
               NotifySheetNotFound($"Index:{sheetIndex}", options);
               return AsyncEnumerable.Empty<T>();
          }

          return ReadSheetInternalAsync<T, TContext>(Sheets[sheetIndex], options, ct);
     }

     public IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadSheetAsDictionaryAsync(
         string sheetName,
         XlsxReadOptions? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions();
          var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
          if (Sheets.All(s => !string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)))
          {
               NotifySheetNotFound(sheetName, options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          var entry = _archive.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();

          return SheetRowReader.ReadDictionaryAsync(entry, _sharedStrings, options, ct);
     }

     public IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadSheetAsDictionaryAsync(
         int sheetIndex,
         XlsxReadOptions? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions();
          if ((uint)sheetIndex >= (uint)Sheets.Count)
          {
               NotifySheetNotFound($"Index:{sheetIndex}", options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          var entry = _archive.GetEntry(Sheets[sheetIndex].ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();

          return SheetRowReader.ReadDictionaryAsync(entry, _sharedStrings, options, ct);
     }

     private IAsyncEnumerable<T> ReadSheetInternalAsync<T>(XlsxSheetInfo sheet, XlsxReadOptions options, CancellationToken ct)
     {
          var map = XlsxTypeRegistry.GetRequiredMap<T>();
          var entry = _archive.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<T>();

          return SheetRowReader.ReadTypedAsync(entry, _sharedStrings, map, options, ct);
     }
     private IAsyncEnumerable<T> ReadSheetInternalAsync<T, TContext>(
        XlsxSheetInfo sheet,
        XlsxReadOptions<T, TContext> options,
        CancellationToken ct)
     {
          var map = XlsxTypeRegistry.GetRequiredContextMap<T, TContext>();
          var entry = _archive.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<T>();

          return SheetRowReader.ReadTypedAsync(entry, _sharedStrings, map, options, ct);
     }

     private static void NotifySheetNotFound(string sheetIdentifier, XlsxReadOptions options)
     {
          var err = new XlsxErrorContext(0, -1, string.Empty, sheetIdentifier, string.Empty, XlsxCellStatus.SheetNotFound);
          options.OnError?.Invoke(in err);
     }

     public void Dispose() => _archive.Dispose();
}
