using FlashXlsx.Abstractions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx;

/// <summary>
/// Point d'entrée principal de FlashXlsx pour la lecture et l'écriture XLSX en streaming (100% Native AOT).
/// </summary>
public static class FlashXlsxSerializer
{
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