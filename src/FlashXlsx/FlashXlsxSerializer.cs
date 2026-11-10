using FlashXlsx.Abstractions;
using FlashXlsx.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx;

/// <summary>
/// Point d'entrée principal de FlashXlsx pour la lecture et l'écriture XLSX en streaming haute performance (100% Native AOT).
/// </summary>
public static class FlashXlsxSerializer
{
     private static FileStream OpenWriteStream(string filePath, XlsxWriteOptions? options)
     {
          try
          {
               return new FileStream(
                   filePath,
                   FileMode.Create,
                   FileAccess.Write,
                   FileShare.None,
                   bufferSize: 65536,
                   FileOptions.Asynchronous | FileOptions.SequentialScan);
          }
          catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
          {
               var err = new XlsxErrorContext(0, -1, string.Empty, filePath, ex.Message, XlsxCellStatus.WriteError);
               options?.OnError?.Invoke(in err);
               throw;
          }
     }

     private static FileStream OpenReadStream(string filePath, XlsxReadOptions? options)
     {
          try
          {
               return new FileStream(
                   filePath,
                   FileMode.Open,
                   FileAccess.Read,
                   FileShare.ReadWrite,
                   bufferSize: 65536,
                   FileOptions.Asynchronous | FileOptions.SequentialScan);
          }
          catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
          {
               var err = new XlsxErrorContext(0, -1, string.Empty, filePath, ex.Message, XlsxCellStatus.CorruptedFile);
               options?.OnError?.Invoke(in err);
               throw;
          }
     }

     // =========================================================================
     // SURCHARGES PAR CHEMIN DE FICHIER (filePath)
     // =========================================================================

     /// <summary>
     /// Écrit une séquence synchrone d'éléments dans un fichier XLSX à l'emplacement spécifié.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX de destination.</param>
     /// <param name="items">La collection d'éléments à exporter.</param>
     /// <param name="options">Options de configuration de l'écriture (nom d'onglet, en-têtes, styles, gestion d'erreurs).</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T>(
         string filePath,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     /// <summary>
     /// Écrit un flux asynchrone d'éléments dans un fichier XLSX à l'emplacement spécifié.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX de destination.</param>
     /// <param name="items">Le flux asynchrone d'éléments à exporter.</param>
     /// <param name="options">Options de configuration de l'écriture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T>(
         string filePath,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     /// <summary>
     /// Écrit une séquence synchrone d'éléments avec un contexte personnalisé dans un fichier XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <typeparam name="TContext">Le type de contexte partagé passé aux accesseurs/formateurs.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX de destination.</param>
     /// <param name="items">La collection d'éléments à exporter.</param>
     /// <param name="options">Options de configuration d'écriture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T, TContext>(
         string filePath,
         IEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     /// <summary>
     /// Écrit un flux asynchrone d'éléments avec un contexte personnalisé dans un fichier XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <typeparam name="TContext">Le type de contexte partagé passé aux accesseurs/formateurs.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX de destination.</param>
     /// <param name="items">Le flux asynchrone d'éléments à exporter.</param>
     /// <param name="options">Options de configuration d'écriture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T, TContext>(
         string filePath,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     /// <summary>
     /// Écrit une séquence de dictionnaires (clé/valeur dynamique) représentant des lignes dans un fichier XLSX.
     /// </summary>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX de destination.</param>
     /// <param name="rows">L'ensemble des lignes représentées sous forme de paires nom de colonne / valeur.</param>
     /// <param name="options">Options de configuration de l'écriture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteDictionaryAsync(
         string filePath,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteDictionaryAsync(fs, rows, options, ct);
     }

     /// <summary>
     /// Lit en streaming continu les lignes d'un fichier XLSX et les désérialise sous forme d'éléments typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles instanciées à partir des lignes.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX à lire.</param>
     /// <param name="options">Options de lecture (onglet sélectionné, mapping personnalisé, politique d'erreurs).</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone <see cref="IAsyncEnumerable{T}"/> d'éléments désérialisés au fil de l'eau.</returns>
     public static async IAsyncEnumerable<T> ReadAsync<T>(
         string filePath,
         XlsxReadOptions<T>? options = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          await foreach (var item in ReadAsync(fs, options, ct))
               yield return item;
     }

     /// <summary>
     /// Lit l'intégralité d'un fichier XLSX et charge toutes les lignes sous forme de liste en mémoire.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX.</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Une <see cref="List{T}"/> contenant tous les éléments du classeur ou de l'onglet ciblé.</returns>
     public static async Task<List<T>> ReadAllAsync<T>(
         string filePath,
         XlsxReadOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          return await ReadAllAsync(fs, options, ct);
     }

     /// <summary>
     /// Lit en streaming continu un fichier XLSX sous forme de dictionnaires dynamiques (nom de colonne / valeur typée).
     /// </summary>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX à lire.</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone de dictionnaires représentant chaque ligne du tableau.</returns>
     public static async IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadDictionaryAsync(
         string filePath,
         XlsxReadOptions? options = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          await foreach (var row in ReadAsDictionaryAsync(fs, options, ct))
               yield return row;
     }

     // =========================================================================
     // LECTURE MULTI-ONGLETS FUSIONNÉE (ReadAllSheetsAsync)
     // =========================================================================

     /// <summary>
     /// Parcourt et lit consécutivement les onglets d'un flux XLSX dans un flux unique d'éléments typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="stream">Le flux contenant le fichier XLSX.</param>
     /// <param name="options">Options de lecture appliquées à chaque onglet parcouru.</param>
     /// <param name="sheetFilter">Prédicat optionnel pour filtrer les onglets à lire (par nom, index, etc.).</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone unifié combinant les enregistrements de tous les onglets filtrés.</returns>
     public static async IAsyncEnumerable<T> ReadAllSheetsAsync<T>(
         Stream stream,
         XlsxReadOptions<T>? options = null,
         Func<XlsxSheetInfo, bool>? sheetFilter = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);
          foreach (var sheet in reader.Sheets)
          {
               if (sheetFilter is not null && !sheetFilter(sheet))
                    continue;

               await foreach (var item in reader.ReadSheetAsync(sheet.Index, options, ct))
                    yield return item;
          }
     }

     /// <summary>
     /// Parcourt et lit consécutivement les onglets d'un fichier XLSX dans un flux unique d'éléments typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX.</param>
     /// <param name="options">Options de lecture appliquées à chaque onglet parcouru.</param>
     /// <param name="sheetFilter">Prédicat optionnel pour inclure ou exclure certains onglets.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone unifié combinant les enregistrements de tous les onglets filtrés.</returns>
     public static async IAsyncEnumerable<T> ReadAllSheetsAsync<T>(
         string filePath,
         XlsxReadOptions<T>? options = null,
         Func<XlsxSheetInfo, bool>? sheetFilter = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          await foreach (var item in ReadAllSheetsAsync(fs, options, sheetFilter, ct))
               yield return item;
     }

     // =========================================================================
     // LECTURE (IMPORT)
     // =========================================================================

     /// <summary>
     /// Lit en streaming continu les lignes d'un flux XLSX et les désérialise sous forme d'éléments typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="stream">Le flux d'entrée contenant le document XLSX.</param>
     /// <param name="options">Options de lecture (onglet ciblé, mapping, etc.).</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone <see cref="IAsyncEnumerable{T}"/> produit en temps réel sans allocation intermédiaire.</returns>
     public static async IAsyncEnumerable<T> ReadAsync<T>(
         Stream stream,
         XlsxReadOptions<T>? options = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions<T>();
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);

          var source = options.SheetName is not null
              ? reader.ReadSheetAsync(options.SheetName, options, ct)
              : reader.ReadSheetAsync(options.SheetIndex, options, ct);

          await foreach (var item in source.WithCancellation(ct))
          {
               yield return item;
          }
     }

     /// <summary>
     /// Lit l'intégralité d'un flux XLSX et charge toutes les lignes sous forme de liste en mémoire.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <param name="stream">Le flux d'entrée contenant le document XLSX.</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Une <see cref="List{T}"/> contenant tous les éléments désérialisés.</returns>
     public static async ValueTask<List<T>> ReadAllAsync<T>(
         Stream stream,
         XlsxReadOptions<T>? options = null,
         CancellationToken ct = default)
     {
          var results = new List<T>();
          await foreach (var item in ReadAsync(stream, options, ct))
          {
               results.Add(item);
          }
          return results;
     }

     /// <summary>
     /// Lit en streaming continu les lignes d'un flux XLSX en injectant un contexte utilisateur spécifique.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <typeparam name="TContext">Le type de contexte partagé injecté lors de la lecture.</typeparam>
     /// <param name="stream">Le flux d'entrée contenant le document XLSX.</param>
     /// <param name="options">Options de lecture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone <see cref="IAsyncEnumerable{T}"/> d'éléments désérialisés.</returns>
     public static async IAsyncEnumerable<T> ReadAsync<T, TContext>(
         Stream stream,
         XlsxReadOptions<T, TContext> options,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);

          var source = options.SheetName is not null
              ? reader.ReadSheetAsync(options.SheetName, options, ct)
              : reader.ReadSheetAsync(options.SheetIndex, options, ct);

          await foreach (var item in source.WithCancellation(ct))
          {
               yield return item;
          }
     }

     /// <summary>
     /// Lit l'intégralité d'un flux XLSX avec un contexte utilisateur et charge toutes les lignes sous forme de liste.
     /// </summary>
     /// <typeparam name="T">Le type des entités cibles.</typeparam>
     /// <typeparam name="TContext">Le type de contexte partagé.</typeparam>
     /// <param name="stream">Le flux d'entrée contenant le document XLSX.</param>
     /// <param name="options">Options de configuration de la lecture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Une <see cref="List{T}"/> contenant tous les éléments désérialisés.</returns>
     public static async ValueTask<List<T>> ReadAllAsync<T, TContext>(
         Stream stream,
         XlsxReadOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          var results = new List<T>();
          await foreach (var item in ReadAsync(stream, options, ct))
          {
               results.Add(item);
          }
          return results;
     }

     /// <summary>
     /// Lit un flux XLSX sous forme de dictionnaires dynamiques (colonnes / valeurs scalaires <see cref="XlsxValue"/>).
     /// </summary>
     /// <param name="stream">Le flux d'entrée contenant le document XLSX.</param>
     /// <param name="options">Options de configuration de la lecture.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     /// <returns>Un flux asynchrone de dictionnaires pour chaque ligne du tableau.</returns>
     public static async IAsyncEnumerable<Dictionary<string, XlsxValue>> ReadAsDictionaryAsync(
         Stream stream,
         XlsxReadOptions? options = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          options ??= new XlsxReadOptions();
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);

          var source = options.SheetName is not null
              ? reader.ReadSheetAsDictionaryAsync(options.SheetName, options, ct)
              : reader.ReadSheetAsDictionaryAsync(options.SheetIndex, options, ct);

          await foreach (var row in source.WithCancellation(ct))
          {
               yield return row;
          }
     }

     /// <summary>
     /// Exécute un plan de lecture multi-onglets structuré (<see cref="XlsxReadWorkbookMap"/>) sur un classeur.
     /// </summary>
     /// <param name="stream">Le flux d'entrée contenant le classeur XLSX.</param>
     /// <param name="map">La cartographie des étapes de lecture configurées pour chaque onglet.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask ReadWorkbookAsync(
         Stream stream,
         XlsxReadWorkbookMap map,
         CancellationToken ct = default)
     {
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);
          foreach (var step in map.Steps)
          {
               await step(reader, ct);
          }
     }

     // =========================================================================
     // ÉCRITURE (EXPORT)
     // =========================================================================

     /// <summary>
     /// Écrit une séquence synchrone d'éléments typés directement dans un flux de sortie XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="items">La collection d'éléments à exporter.</param>
     /// <param name="options">Options de configuration de l'exportation.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T>(
         Stream stream,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     /// <summary>
     /// Écrit un flux asynchrone d'éléments typés directement dans un flux de sortie XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="items">Le flux asynchrone d'éléments à exporter.</param>
     /// <param name="options">Options de configuration de l'exportation.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T>(
         Stream stream,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     /// <summary>
     /// Écrit une séquence synchrone avec un contexte utilisateur personnalisé dans un flux XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <typeparam name="TContext">Le type du contexte partagé.</typeparam>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="items">La collection d'éléments à exporter.</param>
     /// <param name="options">Options de configuration d'écriture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T, TContext>(
         Stream stream,
         IEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     /// <summary>
     /// Écrit un flux asynchrone avec un contexte utilisateur personnalisé dans un flux XLSX.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <typeparam name="TContext">Le type du contexte partagé.</typeparam>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="items">Le flux asynchrone d'éléments à exporter.</param>
     /// <param name="options">Options de configuration d'écriture contextuelle.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteAsync<T, TContext>(
         Stream stream,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     /// <summary>
     /// Écrit une séquence de dictionnaires (lignes dynamiques) dans un flux XLSX.
     /// </summary>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="rows">L'ensemble des lignes sous forme de dictionnaires clé/valeur.</param>
     /// <param name="options">Options d'écriture de l'onglet.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteDictionaryAsync(
         Stream stream,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteDictionarySheetAsync(rows, options, ct);
     }

     /// <summary>
     /// Exécute un plan d'écriture multi-onglets structuré (<see cref="XlsxWriteWorkbookMap"/>) vers un flux de sortie.
     /// </summary>
     /// <param name="stream">Le flux de destination ouvert en écriture.</param>
     /// <param name="map">La cartographie des onglets et données à écrire.</param>
     /// <param name="ct">Jeton d'annulation de l'opération.</param>
     public static async ValueTask WriteWorkbookAsync(
         Stream stream,
         XlsxWriteWorkbookMap map,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          foreach (var step in map.Steps)
          {
               await step(writer, ct);
          }
     }

     // =========================================================================
     // MÉTADONNÉES LÉGÈRES (GetSheets)
     // =========================================================================

     /// <summary>
     /// Extrait rapidement la liste des onglets (nom, index, visibilité) depuis un flux XLSX sans parser les cellules.
     /// </summary>
     /// <param name="stream">Le flux contenant le document XLSX.</param>
     /// <returns>Une liste en lecture seule décrivant chaque onglet trouvé dans le classeur.</returns>
     public static IReadOnlyList<XlsxSheetInfo> GetSheets(Stream stream)
     {
          using var reader = new XlsxWorkbookReader(stream, leaveOpen: true);
          return reader.Sheets;
     }

     /// <summary>
     /// Extrait rapidement la liste des onglets (nom, index, visibilité) depuis un fichier XLSX sans parser les cellules.
     /// </summary>
     /// <param name="filePath">Le chemin d'accès au fichier XLSX.</param>
     /// <returns>Une liste en lecture seule décrivant chaque onglet trouvé dans le classeur.</returns>
     public static IReadOnlyList<XlsxSheetInfo> GetSheets(string filePath)
     {
          using var fs = OpenReadStream(filePath, options: null);
          using var reader = new XlsxWorkbookReader(fs, leaveOpen: false);
          return reader.Sheets;
     }
}