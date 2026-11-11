using FlashXlsx.Abstractions;

namespace FlashXlsx;

/// <summary>
/// Constructeur fluide pour configurer l'export multi-onglets (100% Native AOT).
/// </summary>
public sealed class XlsxWriteWorkbookBuilder
{
     internal XlsxWriteWorkbookMap Map { get; } = new();

     /// <summary>
     /// Ajoute une feuille typée à partir d'une séquence synchrone.
     /// </summary>
     public XlsxWriteWorkbookBuilder AddSheet<T>(
         string sheetName,
         IEnumerable<T> items,
         XlsxWriteOptions<T>? options = null)
     {
          Map.AddSheet(sheetName, items, options);
          return this;
     }

     /// <summary>
     /// Ajoute une feuille typée à partir d'un flux asynchrone.
     /// </summary>
     public XlsxWriteWorkbookBuilder AddSheet<T>(
         string sheetName,
         IAsyncEnumerable<T> items,
         XlsxWriteOptions<T>? options = null)
     {
          Map.AddSheet(sheetName, items, options);
          return this;
     }

     /// <summary>
     /// Ajoute une feuille dynamique de dictionnaires.
     /// </summary>
     public XlsxWriteWorkbookBuilder AddDictionarySheet(
         string sheetName,
         IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
         XlsxWriteOptions? options = null)
     {
          Map.AddDictionarySheet(sheetName, rows, options);
          return this;
     }
}
