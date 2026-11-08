using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using FlashXlsx.Core;
using System.IO.Compression;
using System.Xml;

namespace FlashXlsx;

/// <summary>
/// Lecteur multi-onglets en streaming permettant de découvrir et lire chaque onglet indépendamment.
/// </summary>
public sealed class XlsxWorkbookReader : IDisposable
{
     private readonly ZipArchive _archive;
     private readonly string[] _sharedStrings;
     private readonly bool _use1904DateSystem;

     public IReadOnlyList<XlsxSheetInfo> Sheets { get; }

     public XlsxWorkbookReader(Stream stream, bool leaveOpen = false)
     {
          try
          {
               _archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen);
               _sharedStrings = SharedStringsReader.Read(_archive);
               (Sheets, _use1904DateSystem) = WorkbookMetadataReader.ReadWorkbookInfo(_archive);
          }
          catch (Exception ex) when (ex is InvalidDataException or XmlException)
          {
               throw new InvalidDataException("Le fichier fourni n'est pas un classeur Excel (.xlsx) valide ou est corrompu.", ex);
          }
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
          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

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
          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

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
          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

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
          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

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