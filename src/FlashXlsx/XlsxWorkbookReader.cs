using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using FlashXlsx.Core;
using System.IO.Compression;
using System.Xml;

namespace FlashXlsx;

/// <summary>
/// Lecteur de classeur XLSX multi-onglets en streaming haute performance (100% Native AOT).
/// Permet d'inspecter les métadonnées (onglets, chiffrement) et de lire chaque feuille indépendamment sans charger tout le classeur en mémoire.
/// </summary>
internal sealed class XlsxWorkbookReader : IDisposable
{
     // Signature binaire standard des conteneurs OLE2 / Compound File Binary (fichiers .xlsx chiffrés par mot de passe)
     private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

     private readonly ZipArchive? _archive;
     private readonly Stream? _ownedStream;
     private readonly string[] _sharedStrings;
     private readonly bool _use1904DateSystem;

     /// <summary>
     /// Indique si le classeur est protégé par un mot de passe ou encapsulé dans un conteneur binaire chiffré (OLE2/CFB).
     /// Détecté instantanément sur les 8 premiers octets du flux sans lever d'exception.
     /// </summary>
     public bool IsEncrypted { get; }

     /// <summary>
     /// Obtient la liste des feuilles (nom, index, visibilité, chemin interne) présentes dans le classeur.
     /// </summary>
     public IReadOnlyList<XlsxSheetInfo> Sheets { get; }

     /// <summary>
     /// Initialise un nouveau lecteur de classeur à partir d'un flux ouvert en lecture.
     /// </summary>
     /// <param name="stream">Le flux contenant l'archive XLSX (doit être positionnable pour la détection OLE2).</param>
     /// <param name="leaveOpen"><see langword="true"/> pour laisser le flux ouvert après la libération du lecteur ; sinon, <see langword="false"/>.</param>
     /// <exception cref="InvalidDataException">Lancé si le flux ne représente pas une archive ZIP/XLSX valide.</exception>
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

     /// <summary>
     /// Lit en streaming continu les lignes d'un onglet identifié par son nom et les désérialise sous forme d'éléments typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles généré par AOT.</typeparam>
     /// <param name="sheetName">Le libellé de l'onglet à lire (insensible à la casse).</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone émettant chaque ligne désérialisée au fur et à mesure.</returns>
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

     /// <summary>
     /// Lit en streaming continu les lignes d'un onglet identifié par son index basé sur zéro.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="sheetIndex">L'index de l'onglet dans le classeur (base 0).</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone émettant chaque ligne désérialisée.</returns>
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

     /// <summary>
     /// Lit un onglet spécifique en injectant un contexte utilisateur partagé.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <typeparam name="TContext">Le type de contexte passé aux règles de validation et conditions.</typeparam>
     /// <param name="sheetName">Le libellé de l'onglet à lire.</param>
     /// <param name="options">Options de lecture typées contenant le contexte métier.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone des lignes typées.</returns>
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

     /// <summary>
     /// Lit un onglet spécifique par son index en injectant un contexte utilisateur partagé.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <typeparam name="TContext">Le type du contexte partagé.</typeparam>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     /// <param name="options">Options de lecture typées contenant le contexte métier.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone des lignes typées.</returns>
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

     /// <summary>
     /// Lit un onglet sous forme de dictionnaires dynamiques (colonnes / valeurs scalaires <see cref="XlsxValue"/>) sans mapping précompilé.
     /// </summary>
     /// <param name="sheetName">Le nom de la feuille à lire.</param>
     /// <param name="options">Options de lecture dynamique.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux de dictionnaires représentant chaque ligne.</returns>
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

     /// <summary>
     /// Lit un onglet par son index sous forme de dictionnaires dynamiques.
     /// </summary>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     /// <param name="options">Options de lecture dynamique.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux de dictionnaires représentant chaque ligne.</returns>
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

     /// <summary>
     /// Notifie l'erreur EncryptedFile sur les options fournies si le classeur est chiffré.
     /// </summary>
     internal void NotifyIfEncrypted(XlsxReadOptions? options)
     {
          if (IsEncrypted && options is not null)
          {
               NotifyEncryptedFile(options);
          }
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

     /// <summary>
     /// Libère les ressources du décompresseur et le flux sous-jacent si approprié.
     /// </summary>
     public void Dispose()
     {
          _archive?.Dispose();
          _ownedStream?.Dispose();
     }
}