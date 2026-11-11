using FlashXlsx.Abstractions;

namespace FlashXlsx;

/// <summary>
/// Constructeur fluide pour configurer l'exportation d'un classeur multi-onglets (100% Native AOT).
/// </summary>
public sealed class XlsxWriteWorkbookBuilder
{
    internal XlsxWriteWorkbookMap Map { get; } = new();

    /// <summary>
    /// Ajoute une feuille de données typées à partir d'une séquence synchrone.
    /// </summary>
    /// <typeparam name="T">Le type des entités sérialisées dans la feuille.</typeparam>
    /// <param name="sheetName">Le nom de l'onglet à créer dans Excel.</param>
    /// <param name="items">La collection d'éléments à exporter.</param>
    /// <param name="options">Options d'exportation spécifiques à cet onglet (en-têtes, styles, filtres).</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxWriteWorkbookBuilder AddSheet<T>(
        string sheetName,
        IEnumerable<T> items,
        XlsxWriteOptions<T>? options = null)
    {
        Map.AddSheet(sheetName, items, options);
        return this;
    }

    /// <summary>
    /// Ajoute une feuille de données typées à partir d'un flux asynchrone (streaming).
    /// </summary>
    /// <typeparam name="T">Le type des entités sérialisées dans la feuille.</typeparam>
    /// <param name="sheetName">Le nom de l'onglet à créer dans Excel.</param>
    /// <param name="items">Le flux asynchrone d'éléments à exporter au fur et à mesure.</param>
    /// <param name="options">Options d'exportation spécifiques à cet onglet.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxWriteWorkbookBuilder AddSheet<T>(
        string sheetName,
        IAsyncEnumerable<T> items,
        XlsxWriteOptions<T>? options = null)
    {
        Map.AddSheet(sheetName, items, options);
        return this;
    }

    /// <summary>
    /// Ajoute une feuille dynamique basée sur des dictionnaires clés/valeurs sans schéma rigide (zéro-boxing).
    /// </summary>
    /// <param name="sheetName">Le nom de l'onglet à créer dans Excel.</param>
    /// <param name="rows">La séquence de lignes représentées sous forme de dictionnaires colonnes / valeurs.</param>
    /// <param name="options">Options d'exportation applicables à cette feuille.</param>
    /// <returns>L'instance du constructeur pour enchaîner la configuration.</returns>
    public XlsxWriteWorkbookBuilder AddDictionarySheet(
        string sheetName,
        IEnumerable<IReadOnlyDictionary<string, XlsxValue>> rows,
        XlsxWriteOptions? options = null)
    {
        Map.AddDictionarySheet(sheetName, rows, options);
        return this;
    }
}