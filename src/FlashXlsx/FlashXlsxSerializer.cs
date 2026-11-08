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
/// Point d'entrée principal de FlashXlsx pour la lecture et l'écriture XLSX en streaming (100% Native AOT).
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

     public static async ValueTask WriteAsync<T>(
         string filePath,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     public static async ValueTask WriteAsync<T>(
         string filePath,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     public static async ValueTask WriteAsync<T, TContext>(
         string filePath,
         IEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     public static async ValueTask WriteAsync<T, TContext>(
         string filePath,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteAsync(fs, items, options, ct);
     }

     public static async ValueTask WriteDictionaryAsync(
         string filePath,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenWriteStream(filePath, options);
          await WriteDictionaryAsync(fs, rows, options, ct);
     }

     public static async IAsyncEnumerable<T> ReadAsync<T>(
         string filePath,
         XlsxReadOptions<T>? options = null,
         [EnumeratorCancellation] CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          await foreach (var item in ReadAsync(fs, options, ct))
               yield return item;
     }

     public static async Task<List<T>> ReadAllAsync<T>(
         string filePath,
         XlsxReadOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var fs = OpenReadStream(filePath, options);
          return await ReadAllAsync(fs, options, ct);
     }

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

     public static async ValueTask WriteAsync<T>(
         Stream stream,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     public static async ValueTask WriteAsync<T>(
         Stream stream,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     public static async ValueTask WriteAsync<T, TContext>(
         Stream stream,
         IEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     public static async ValueTask WriteAsync<T, TContext>(
         Stream stream,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T, TContext> options,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteSheetAsync(items, options, ct);
     }

     public static async ValueTask WriteDictionaryAsync(
         Stream stream,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null,
         CancellationToken ct = default)
     {
          await using var writer = new XlsxWorkbookWriter(stream, leaveOpen: true);
          await writer.WriteDictionarySheetAsync(rows, options, ct);
     }

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
}