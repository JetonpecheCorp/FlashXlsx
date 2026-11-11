using FlashXlsx.Abstractions;

namespace FlashXlsx;

/// <summary>
/// Plan déclaratif d'importation multi-onglets (100% Native AOT, zéro boxing).
/// Permet de configurer l'extraction de plusieurs feuilles au cours d'une seule passe de lecture via <see cref="FlashXlsxSerializer.ReadWorkbookAsync(string, Action{XlsxReadWorkbookBuilder}, FlashXlsx.Abstractions.XlsxReadOptions?, System.Threading.CancellationToken)"/>.
/// </summary>
internal sealed class XlsxReadWorkbookMap
{
     internal List<Func<XlsxWorkbookReader, CancellationToken, ValueTask>> Steps { get; } = [];

     /// <summary>
     /// Associe un onglet par son nom à une liste de résultats typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités à désérialiser.</typeparam>
     /// <param name="sheetName">Le libellé de l'onglet source dans le classeur.</param>
     /// <param name="results">La liste de sortie qui recevra les entités lues.</param>
     /// <param name="options">Options de lecture appliquées à cet onglet.</param>
     /// <returns>L'instance courante pour enchaîner les mappings.</returns>
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

     /// <summary>
     /// Associe un onglet par son nom à un délégué de rappel exécuté pour chaque ligne lue (streaming).
     /// </summary>
     /// <typeparam name="T">Le type des entités à désérialiser.</typeparam>
     /// <param name="sheetName">Le libellé de l'onglet source.</param>
     /// <param name="onRowAsync">Le délégué de traitement asynchrone invoqué à chaque ligne.</param>
     /// <param name="options">Options de lecture appliquées à cet onglet.</param>
     /// <returns>L'instance courante pour enchaîner les mappings.</returns>
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

     /// <summary>
     /// Associe un onglet par son index (base 0) à une liste de résultats typés.
     /// </summary>
     /// <typeparam name="T">Le type des entités à désérialiser.</typeparam>
     /// <param name="sheetIndex">L'index basé sur zéro de l'onglet.</param>
     /// <param name="results">La liste de sortie qui recevra les entités lues.</param>
     /// <param name="options">Options de lecture appliquées à cet onglet.</param>
     /// <returns>L'instance courante pour enchaîner les mappings.</returns>
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

     /// <summary>
     /// Associe un onglet par son nom à une liste de dictionnaires dynamiques.
     /// </summary>
     /// <param name="sheetName">Le nom de l'onglet.</param>
     /// <param name="results">La liste recevant les dictionnaires de colonnes/valeurs.</param>
     /// <param name="options">Options de lecture appliquées à cet onglet.</param>
     /// <returns>L'instance courante pour enchaîner les mappings.</returns>
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

     /// <summary>
     /// Associe un onglet par son index (base 0) à une liste de dictionnaires dynamiques.
     /// </summary>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     /// <param name="results">La liste recevant les dictionnaires de colonnes/valeurs.</param>
     /// <param name="options">Options de lecture appliquées à cet onglet.</param>
     /// <returns>L'instance courante pour enchaîner les mappings.</returns>
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
/// Plan déclaratif d'exportation multi-onglets (100% Native AOT).
/// Permet de déclarer plusieurs feuilles à générer consécutivement au cours d'une seule écriture via <see cref="FlashXlsxSerializer.WriteWorkbookAsync(string, Action{XlsxWriteWorkbookBuilder}, System.Threading.CancellationToken)"/>.
/// </summary>
internal sealed class XlsxWriteWorkbookMap
{
     internal List<Func<XlsxWorkbookWriter, CancellationToken, ValueTask>> Steps { get; } = [];

     /// <summary>
     /// Ajoute une feuille typée à partir d'une collection synchrone d'éléments.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="sheetName">Le nom de l'onglet à créer.</param>
     /// <param name="items">La collection d'éléments à inscrire.</param>
     /// <param name="options">Options de configuration de la feuille.</param>
     /// <returns>L'instance courante pour enchaîner les ajouts.</returns>
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

     /// <summary>
     /// Ajoute une feuille typée à partir d'un flux asynchrone d'éléments.
     /// </summary>
     /// <typeparam name="T">Le type des entités à sérialiser.</typeparam>
     /// <param name="sheetName">Le nom de l'onglet à créer.</param>
     /// <param name="items">Le flux asynchrone d'éléments à inscrire.</param>
     /// <param name="options">Options de configuration de la feuille.</param>
     /// <returns>L'instance courante pour enchaîner les ajouts.</returns>
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

     /// <summary>
     /// Ajoute une feuille dynamique à partir d'une séquence de dictionnaires (clé/valeur).
     /// </summary>
     /// <param name="sheetName">Le nom de l'onglet à créer.</param>
     /// <param name="rows">L'ensemble des lignes représentées sous forme de dictionnaires de colonnes/valeurs.</param>
     /// <param name="options">Options de configuration de la feuille.</param>
     /// <returns>L'instance courante pour enchaîner les ajouts.</returns>
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