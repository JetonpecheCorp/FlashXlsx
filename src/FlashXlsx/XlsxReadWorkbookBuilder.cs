using FlashXlsx.Abstractions;

namespace FlashXlsx;

/// <summary>
/// Constructeur fluide pour configurer l'importation d'un classeur multi-onglets (100% Native AOT, zéro boxing).
/// </summary>
public sealed class XlsxReadWorkbookBuilder
{
    internal XlsxReadWorkbookMap Map { get; } = new();

    /// <summary>
    /// Mappe un onglet désigné par son nom vers une liste typée de sortie.
    /// </summary>
    /// <typeparam name="T">Le type des entités désérialisées pour cet onglet.</typeparam>
    /// <param name="sheetName">Le nom exact de la feuille à extraire.</param>
    /// <param name="results">La liste recevant l'ensemble des instances désérialisées.</param>
    /// <param name="options">Options de lecture spécifiques à cette feuille (en-têtes, alias, gestion d'erreurs).</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxReadWorkbookBuilder MapSheet<T>(
        string sheetName,
        out List<T> results,
        XlsxReadOptions<T>? options = null)
    {
        Map.MapSheet(sheetName, out results, options);
        return this;
    }

    /// <summary>
    /// Mappe un onglet désigné par son nom vers un délégué asynchrone exécuté pour chaque ligne (streaming ligne par ligne).
    /// </summary>
    /// <typeparam name="T">Le type des entités désérialisées pour cet onglet.</typeparam>
    /// <param name="sheetName">Le nom exact de la feuille à lire.</param>
    /// <param name="onRowAsync">La fonction de rappel asynchrone traitant chaque ligne lue.</param>
    /// <param name="options">Options de lecture spécifiques à cette feuille.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxReadWorkbookBuilder MapSheet<T>(
        string sheetName,
        Func<T, ValueTask> onRowAsync,
        XlsxReadOptions<T>? options = null)
    {
        Map.MapSheet(sheetName, onRowAsync, options);
        return this;
    }

    /// <summary>
    /// Mappe un onglet désigné par son index basé sur zéro vers une liste typée de sortie.
    /// </summary>
    /// <typeparam name="T">Le type des entités désérialisées pour cet onglet.</typeparam>
    /// <param name="sheetIndex">L'index basé sur zéro de la feuille à extraire.</param>
    /// <param name="results">La liste recevant l'ensemble des instances désérialisées.</param>
    /// <param name="options">Options de lecture spécifiques à cette feuille.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxReadWorkbookBuilder MapSheet<T>(
        int sheetIndex,
        out List<T> results,
        XlsxReadOptions<T>? options = null)
    {
        Map.MapSheet(sheetIndex, out results, options);
        return this;
    }

    /// <summary>
    /// Mappe un onglet désigné par son nom vers une liste de dictionnaires dynamiques (colonnes / valeurs scalaires sans boxing).
    /// </summary>
    /// <param name="sheetName">Le nom exact de la feuille à extraire.</param>
    /// <param name="results">La liste recevant les lignes sous forme de dictionnaires dynamiques.</param>
    /// <param name="options">Options de lecture spécifiques à cette feuille.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxReadWorkbookBuilder MapSheetAsDictionary(
        string sheetName,
        out List<Dictionary<string, XlsxValue>> results,
        XlsxReadOptions? options = null)
    {
        Map.MapSheetAsDictionary(sheetName, out results, options);
        return this;
    }

    /// <summary>
    /// Mappe un onglet désigné par son index basé sur zéro vers une liste de dictionnaires dynamiques.
    /// </summary>
    /// <param name="sheetIndex">L'index basé sur zéro de la feuille à extraire.</param>
    /// <param name="results">La liste recevant les lignes sous forme de dictionnaires dynamiques.</param>
    /// <param name="options">Options de lecture spécifiques à cette feuille.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxReadWorkbookBuilder MapSheetAsDictionary(
        int sheetIndex,
        out List<Dictionary<string, XlsxValue>> results,
        XlsxReadOptions? options = null)
    {
        Map.MapSheetAsDictionary(sheetIndex, out results, options);
        return this;
    }
}