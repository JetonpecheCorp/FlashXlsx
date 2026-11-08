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
     // Signature binaire standard des conteneurs OLE2 / Compound File Binary (fichiers .xlsx chiffrés par mot de passe)
     private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

     private readonly ZipArchive? _archive;
     private readonly Stream? _ownedStream;
     private readonly string[] _sharedStrings;
     private readonly bool _use1904DateSystem;

     public bool IsEncrypted { get; }
     public IReadOnlyList<XlsxSheetInfo> Sheets { get; }

     public XlsxWorkbookReader(Stream stream, bool leaveOpen = false)
     {
          if (HasOle2EncryptedSignature(stream))
          {
               IsEncrypted = true;
               Sheets = [];
               _sharedStrings = [];
               if (!leaveOpen)
                    _ownedStream = stream;
               return;
          }

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

     private static bool HasOle2EncryptedSignature(Stream stream)
     {
          if (!stream.CanSeek)
               return false;

          long originalPos = stream.Position;
          Span<byte> header = stackalloc byte[8];
          int totalRead = 0;

          while (totalRead < 8)
          {
               int read = stream.Read(header[totalRead..]);
               if (read == 0)
                    break;
               totalRead += read;
          }

          stream.Position = originalPos;
          return totalRead == 8 && header.SequenceEqual(Ole2Signature);
     }

     public IAsyncEnumerable<T> ReadSheetAsync<T>(string sheetName, XlsxReadOptions<T>? options = null, CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions<T>();
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<T>();
          }

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
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<T>();
          }

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
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<T>();
          }

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
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<T>();
          }

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
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

          var sheet = Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
          if (Sheets.All(s => !string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase)))
          {
               NotifySheetNotFound(sheetName, options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          var entry = _archive?.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();

          if (Utf8SheetRowReader.CanUseFastUtf8Path(options))
               return Utf8SheetRowReader.ReadDictionaryFastAsync(entry, _sharedStrings, options, ct);

          return SheetRowReader.ReadDictionaryAsync(entry, _sharedStrings, options, ct);
     }

     public IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadSheetAsDictionaryAsync(
         int sheetIndex,
         XlsxReadOptions? options = null,
         CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions();
          if (IsEncrypted)
          {
               NotifyEncryptedFile(options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

          if ((uint)sheetIndex >= (uint)Sheets.Count)
          {
               NotifySheetNotFound($"Index:{sheetIndex}", options);
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();
          }

          var entry = _archive?.GetEntry(Sheets[sheetIndex].ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<Dictionary<string, XlsxValue>>();

          if (Utf8SheetRowReader.CanUseFastUtf8Path(options))
               return Utf8SheetRowReader.ReadDictionaryFastAsync(entry, _sharedStrings, options, ct);

          return SheetRowReader.ReadDictionaryAsync(entry, _sharedStrings, options, ct);
     }

     private IAsyncEnumerable<T> ReadSheetInternalAsync<T>(XlsxSheetInfo sheet, XlsxReadOptions options, CancellationToken ct)
     {
          if (_use1904DateSystem)
               options.Use1904DateSystem = true;

          var map = XlsxTypeRegistry.GetRequiredMap<T>();
          var entry = _archive?.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<T>();

          if (Utf8SheetRowReader.CanUseFastUtf8Path(options))
               return Utf8SheetRowReader.ReadTypedFastAsync(entry, _sharedStrings, map, options, ct);

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
          var entry = _archive?.GetEntry(sheet.ZipEntryPath);
          if (entry is null)
               return AsyncEnumerable.Empty<T>();

          if (Utf8SheetRowReader.CanUseFastUtf8Path(options))
               return Utf8SheetRowReader.ReadTypedContextFastAsync(entry, _sharedStrings, map, options, ct);

          return SheetRowReader.ReadTypedAsync(entry, _sharedStrings, map, options, ct);
     }

     private static void NotifySheetNotFound(string sheetIdentifier, XlsxReadOptions options)
     {
          var err = new XlsxErrorContext(0, -1, string.Empty, sheetIdentifier, string.Empty, XlsxCellStatus.SheetNotFound);
          options.OnError?.Invoke(in err);
     }

     private static void NotifyEncryptedFile(XlsxReadOptions options)
     {
          var err = new XlsxErrorContext(0, -1, string.Empty, null, "Le classeur Excel est protégé par mot de passe (chiffré).", XlsxCellStatus.EncryptedFile);
          options.OnError?.Invoke(in err);
     }

     public void Dispose()
     {
          _archive?.Dispose();
          _ownedStream?.Dispose();
     }
}