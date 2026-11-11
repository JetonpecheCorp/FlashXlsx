using FlashXlsx.Abstractions;

namespace FlashXlsx;

/// <summary>
/// Constructeur fluide pour configurer l'import multi-onglets (100% Native AOT, zéro boxing).
/// </summary>
public sealed class XlsxReadWorkbookBuilder
{
     internal XlsxReadWorkbookMap Map { get; } = new();

     /// <summary>
     /// Mappe un onglet par son nom vers une liste typée de sortie.
     /// </summary>
     public XlsxReadWorkbookBuilder MapSheet<T>(
         string sheetName,
         out List<T> results,
         XlsxReadOptions<T>? options = null)
     {
          Map.MapSheet(sheetName, out results, options);
          return this;
     }

     /// <summary>
     /// Mappe un onglet par son nom vers un rappel asynchrone par ligne (streaming).
     /// </summary>
     public XlsxReadWorkbookBuilder MapSheet<T>(
         string sheetName,
         Func<T, ValueTask> onRowAsync,
         XlsxReadOptions<T>? options = null)
     {
          Map.MapSheet(sheetName, onRowAsync, options);
          return this;
     }

     /// <summary>
     /// Mappe un onglet par son index (base 0) vers une liste typée de sortie.
     /// </summary>
     public XlsxReadWorkbookBuilder MapSheet<T>(
         int sheetIndex,
         out List<T> results,
         XlsxReadOptions<T>? options = null)
     {
          Map.MapSheet(sheetIndex, out results, options);
          return this;
     }

     /// <summary>
     /// Mappe un onglet par son nom vers une liste de dictionnaires dynamiques.
     /// </summary>
     public XlsxReadWorkbookBuilder MapSheetAsDictionary(
         string sheetName,
         out List<Dictionary<string, XlsxValue>> results,
         XlsxReadOptions? options = null)
     {
          Map.MapSheetAsDictionary(sheetName, out results, options);
          return this;
     }

     /// <summary>
     /// Mappe un onglet par son index (base 0) vers une liste de dictionnaires dynamiques.
     /// </summary>
     public XlsxReadWorkbookBuilder MapSheetAsDictionary(
         int sheetIndex,
         out List<Dictionary<string, XlsxValue>> results,
         XlsxReadOptions? options = null)
     {
          Map.MapSheetAsDictionary(sheetIndex, out results, options);
          return this;
     }
}
