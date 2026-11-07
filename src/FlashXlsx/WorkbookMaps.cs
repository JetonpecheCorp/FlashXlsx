using FlashXlsx.Abstractions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx;

/// <summary>
/// Plan déclaratif d'import multi-onglets (100% Native AOT, zéro boxing).
/// </summary>
public sealed class XlsxReadWorkbookMap
{
     internal List<Func<XlsxWorkbookReader, CancellationToken, ValueTask>> Steps { get; } = [];

     public XlsxReadWorkbookMap MapSheet<T>(
         string sheetName,
         out List<T> results,
         XlsxReadOptions<T>? options = null)
     {
          var list = new List<T>();
          results = list;

          Steps.Add(async (reader, ct) =>
          {
               await foreach (var item in reader.ReadSheetAsync(sheetName, options, ct))
                    list.Add(item);
          });
          return this;
     }

     public XlsxReadWorkbookMap MapSheet<T>(
         string sheetName,
         Func<T, ValueTask> onRowAsync,
         XlsxReadOptions<T>? options = null)
     {
          Steps.Add(async (reader, ct) =>
          {
               await foreach (var item in reader.ReadSheetAsync(sheetName, options, ct))
                    await onRowAsync(item);
          });
          return this;
     }

     public XlsxReadWorkbookMap MapSheet<T>(
         int sheetIndex,
         out List<T> results,
         XlsxReadOptions<T>? options = null)
     {
          var list = new List<T>();
          results = list;

          Steps.Add(async (reader, ct) =>
          {
               await foreach (var item in reader.ReadSheetAsync(sheetIndex, options, ct))
                    list.Add(item);
          });
          return this;
     }

     public XlsxReadWorkbookMap MapSheetAsDictionary(
         string sheetName,
         out List<Dictionary<string, XlsxValue>> results,
         XlsxReadOptions? options = null)
     {
          var list = new List<Dictionary<string, XlsxValue>>();
          results = list;

          Steps.Add(async (reader, ct) =>
          {
               await foreach (var row in reader.ReadSheetAsDictionaryAsync(sheetName, options, ct))
                    list.Add(row);
          });
          return this;
     }

     public XlsxReadWorkbookMap MapSheetAsDictionary(
         int sheetIndex,
         out List<Dictionary<string, XlsxValue>> results,
         XlsxReadOptions? options = null)
     {
          var list = new List<Dictionary<string, XlsxValue>>();
          results = list;

          Steps.Add(async (reader, ct) =>
          {
               await foreach (var row in reader.ReadSheetAsDictionaryAsync(sheetIndex, options, ct))
                    list.Add(row);
          });
          return this;
     }
}

/// <summary>
/// Plan déclaratif d'export multi-onglets (100% Native AOT).
/// </summary>
public sealed class XlsxWriteWorkbookMap
{
     internal List<Func<XlsxWorkbookWriter, CancellationToken, ValueTask>> Steps { get; } = [];

     public XlsxWriteWorkbookMap AddSheet<T>(
         string sheetName,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null)
     {
          options ??= new XlsxWriteOptions<T>();
          options.SheetName = sheetName;

          Steps.Add((writer, ct) => writer.WriteSheetAsync(items, options, ct));
          return this;
     }

     public XlsxWriteWorkbookMap AddSheet<T>(
         string sheetName,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null)
     {
          options ??= new XlsxWriteOptions<T>();
          options.SheetName = sheetName;

          Steps.Add((writer, ct) => writer.WriteSheetAsync(items, options, ct));
          return this;
     }

     public XlsxWriteWorkbookMap AddDictionarySheet(
         string sheetName,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null)
     {
          options ??= new XlsxWriteOptions();
          options.SheetName = sheetName;

          Steps.Add((writer, ct) => writer.WriteDictionarySheetAsync(rows, options, ct));
          return this;
     }
}
