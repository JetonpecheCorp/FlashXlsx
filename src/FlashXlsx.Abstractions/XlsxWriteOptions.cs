using FlashXlsx.Abstractions.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Options d'écriture pour le mode dynamique (Dictionary) et socle commun pour toutes les exportations XLSX.
/// </summary>
public class XlsxWriteOptions
{
     private readonly Dictionary<string, string> _translations = new(StringComparer.OrdinalIgnoreCase);

     /// <summary>Nom de l'onglet Excel créé lors de l'exportation (par défaut : "Sheet1").</summary>
     public string SheetName { get; set; } = "Sheet1";

     /// <summary>Indique si la ligne d'en-tête doit être écrite en première ligne (activé par défaut).</summary>
     public bool WriteHeader { get; set; } = true;

     /// <summary>Culture utilisée pour le formatage des données (par défaut : <see cref="CultureInfo.InvariantCulture"/>).</summary>
     public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

     /// <summary>Séparateur décimal utilisé lorsqu'une colonne décimale est formatée avec un motif texte ([XlsxFormat]) ou un séparateur explicite.</summary>
     public char DecimalSeparator { get; set; } = '.';

     /// <summary>Calcule et ajuste automatiquement la largeur des colonnes en fonction de la taille maximale des valeurs écrites.</summary>
     public bool AutoFitColumns { get; set; }

     /// <summary>Fige la ligne d'en-tête pour qu'elle reste visible lors du défilement vertical dans Excel.</summary>
     public bool FreezeHeader { get; set; }

     /// <summary>Génère des listes déroulantes de validation de données natives Excel pour les propriétés énumérées ou contraintes par [XlsxAllowedValues].</summary>
     public bool DropdownValidation { get; set; }

     /// <summary>Fournisseur de notification de progression recevant le nombre cumulé de lignes de données écrites.</summary>
     public IProgress<int>? Progress { get; set; }

     /// <summary>Intervalle en nombre de lignes entre deux notifications de progression (par défaut : 1000).</summary>
     public int ProgressInterval { get; set; } = 1000;

     /// <summary>Comportement en cas d'erreur de sérialisation ou de formatage (par défaut : arrêt à la première erreur).</summary>
     public XlsxErrorMode ErrorMode { get; set; } = XlsxErrorMode.StopOnFirstError;

     /// <summary>Gestionnaire appelé lors de la survenue d'une erreur d'écriture sans allocation d'exception.</summary>
     public XlsxErrorHandler? OnError { get; set; }

     /// <summary>Configuration visuelle personnalisée de la ligne d'en-tête (police en gras, couleur de fond, couleur du texte).</summary>
     public XlsxHeaderStyle? HeaderStyle { get; set; }

     /// <summary>Configuration pour transformer la plage de données exportée en un vrai tableau natif Excel (ListObject / Table).</summary>
     public XlsxTableOptions? ExcelTable { get; set; }

     /// <summary>Libellés textuels personnalisés écrits pour la valeur booléenne <see langword="true"/>.</summary>
     public string[]? DefaultTrueValues { get; private set; }

     /// <summary>Libellés textuels personnalisés écrits pour la valeur booléenne <see langword="false"/>.</summary>
     public string[]? DefaultFalseValues { get; private set; }

     /// <summary>
     /// Mappe un nom de propriété/clé vers l'en-tête traduit à écrire dans Excel.
     /// </summary>
     public IReadOnlyDictionary<string, string> Translations => _translations;

     /// <summary>Politique de sécurité appliquée lors de l'écriture (neutralisation des caractères sensibles '=' ou formules non désirées).</summary>
     public XlsxSecurityPolicy SecurityPolicy { get; set; } = XlsxSecurityPolicy.Sanitize;

     /// <summary>Définit la politique de sécurité pour la neutralisation des formules.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public XlsxWriteOptions WithSecurityPolicy(XlsxSecurityPolicy policy) { SecurityPolicy = policy; return this; }

     /// <summary>Définit le libellé de l'onglet Excel créé.</summary>
     /// <param name="sheetName">Le nom de la feuille.</param>
     public XlsxWriteOptions WithSheetName(string sheetName) { SheetName = sheetName; return this; }

     /// <summary>Active ou désactive la génération de la ligne d'en-tête.</summary>
     /// <param name="writeHeader">Indique s'il faut inscrire les en-têtes.</param>
     public XlsxWriteOptions WithHeader(bool writeHeader) { WriteHeader = writeHeader; return this; }

     /// <summary>Définit la culture appliquée au formatage.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public XlsxWriteOptions WithCulture(CultureInfo culture) { Culture = culture; return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public XlsxWriteOptions WithErrorMode(XlsxErrorMode mode) { ErrorMode = mode; return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs d'écriture sans lever d'exceptions.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public XlsxWriteOptions WithErrorHandler(XlsxErrorHandler handler) { OnError = handler; return this; }

     /// <summary>Attache un observateur de progression de l'écriture.</summary>
     /// <param name="progress">Le canal de notification recevant le nombre de lignes écrites.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement en nombre de lignes.</param>
     public XlsxWriteOptions WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          Progress = progress;
          ProgressInterval = intervalRows > 0 ? intervalRows : 1000;
          return this;
     }

     /// <summary>Définit les libellés textuels personnalisés écrits pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés employés pour vrai.</param>
     /// <param name="falseValues">Libellés employés pour faux.</param>
     public XlsxWriteOptions WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null)
     {
          DefaultTrueValues = trueValues;
          DefaultFalseValues = falseValues;
          return this;
     }

     /// <summary>Définit le caractère séparateur décimal explicite appliqué lors du formatage texte des nombres.</summary>
     /// <param name="separator">Le séparateur décimal (ex: ',' ou '.').</param>
     public XlsxWriteOptions WithDecimalSeparator(char separator)
     {
          DecimalSeparator = separator;
          return this;
     }

     /// <summary>Active ou désactive l'ajustement automatique de la largeur des colonnes à la taille du texte le plus long.</summary>
     /// <param name="autoFit">Indique si le dimensionnement automatique est activé.</param>
     public XlsxWriteOptions WithAutoFitColumns(bool autoFit = true)
     {
          AutoFitColumns = autoFit;
          return this;
     }

     /// <summary>Personnalise le style graphique appliqué à la ligne d'en-tête.</summary>
     /// <param name="bold">Active le texte en gras.</param>
     /// <param name="backgroundHex">Code hexadécimal RVB de la couleur de fond (ex: "F2F2F2").</param>
     /// <param name="textHex">Code hexadécimal RVB de la couleur du texte (ex: "000000").</param>
     public XlsxWriteOptions WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          HeaderStyle = new XlsxHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     /// <summary>Encapsule la plage de données dans un objet Table Excel natif avec style et filtres automatiques.</summary>
     /// <param name="tableName">Nom logique unique du tableau Excel (généré automatiquement si null).</param>
     /// <param name="styleName">Nom du style visuel prédéfini Excel (par défaut : "TableStyleMedium2").</param>
     /// <param name="showRowStripes">Active l'alternance de couleur sur les lignes.</param>
     /// <param name="enableAutoFilter">Active les boutons de filtrage/tri sur chaque en-tête.</param>
     public XlsxWriteOptions WithExcelTable(
         string? tableName = null,
         string styleName = "TableStyleMedium2",
         bool showRowStripes = true,
         bool enableAutoFilter = true)
     {
          ExcelTable = new XlsxTableOptions(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     /// <summary>Active ou désactive le figeage de la première ligne d'en-tête lors du défilement dans Excel.</summary>
     /// <param name="freeze">Indique si la ligne d'en-tête doit être figée.</param>
     public XlsxWriteOptions WithFreezeHeader(bool freeze = true)
     {
          FreezeHeader = freeze;
          return this;
     }

     /// <summary>Active ou désactive la génération automatique de menus déroulants de validation Excel sur les colonnes contraintes.</summary>
     /// <param name="enable">Indique si les validations de listes déroulantes doivent être injectées.</param>
     public XlsxWriteOptions WithDropdownValidation(bool enable = true)
     {
          DropdownValidation = enable;
          return this;
     }

     /// <summary>
     /// Mappe un nom de clé ou de propriété vers un libellé d'en-tête personnalisé dans le fichier Excel.
     /// </summary>
     /// <param name="key">Le nom de la clé logique.</param>
     /// <param name="excelHeader">Le libellé d'en-tête à écrire dans la cellule.</param>
     public XlsxWriteOptions Translate(string key, string excelHeader)
     {
          _translations[key] = excelHeader;
          return this;
     }

     /// <summary>
     /// Injecte un ensemble de traductions d'en-têtes préconfigurées.
     /// </summary>
     /// <param name="translations">Dictionnaire associant une clé à son nom de colonne d'export.</param>
     public XlsxWriteOptions WithTranslations(IReadOnlyDictionary<string, string> translations)
     {
          foreach (var kvp in translations)
               _translations[kvp.Key] = kvp.Value;
          return this;
     }

     internal void SetTranslationInternal(string propertyName, string excelHeader)
     {
          _translations[propertyName] = excelHeader;
     }
}

/// <summary>
/// Options d'écriture fortement typées pour un modèle <typeparamref name="T"/> (100% Native AOT).
/// </summary>
/// <typeparam name="T">Le type de l'entité sérialisée.</typeparam>
public class XlsxWriteOptions<T> : XlsxWriteOptions
{
     /// <summary>Définit le libellé de l'onglet Excel créé.</summary>
     /// <param name="sheetName">Le nom de la feuille.</param>
     public new XlsxWriteOptions<T> WithSheetName(string sheetName) { base.WithSheetName(sheetName); return this; }

     /// <summary>Active ou désactive la génération de la ligne d'en-tête.</summary>
     /// <param name="writeHeader">Indique s'il faut inscrire les en-têtes.</param>
     public new XlsxWriteOptions<T> WithHeader(bool writeHeader) { base.WithHeader(writeHeader); return this; }

     /// <summary>Définit la culture appliquée au formatage.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public new XlsxWriteOptions<T> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }

     /// <summary>Définit les libellés textuels personnalisés écrits pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés employés pour vrai.</param>
     /// <param name="falseValues">Libellés employés pour faux.</param>
     public new XlsxWriteOptions<T> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }

     /// <summary>Définit le caractère séparateur décimal explicite.</summary>
     /// <param name="separator">Le séparateur décimal.</param>
     public new XlsxWriteOptions<T> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public new XlsxWriteOptions<T> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs d'écriture.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public new XlsxWriteOptions<T> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }

     /// <summary>Injecte un ensemble de traductions d'en-têtes préconfigurées.</summary>
     /// <param name="translations">Dictionnaire associant une propriété à son libellé Excel.</param>
     public new XlsxWriteOptions<T> WithTranslations(IReadOnlyDictionary<string, string> translations)
     {
          base.WithTranslations(translations);
          return this;
     }

     /// <summary>Attache un observateur de progression de l'écriture.</summary>
     /// <param name="progress">Le canal de notification de progression.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement.</param>
     public new XlsxWriteOptions<T> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     /// <summary>Définit la politique de sécurité pour la neutralisation des formules.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public new XlsxWriteOptions<T> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     /// <summary>Active ou désactive l'ajustement automatique de la largeur des colonnes.</summary>
     /// <param name="autoFit">Indique si le dimensionnement automatique est activé.</param>
     public new XlsxWriteOptions<T> WithAutoFitColumns(bool autoFit = true)
     {
          base.WithAutoFitColumns(autoFit);
          return this;
     }

     /// <summary>Personnalise le style graphique appliqué à la ligne d'en-tête.</summary>
     /// <param name="bold">Active le texte en gras.</param>
     /// <param name="backgroundHex">Code couleur hexadécimal RVB de fond.</param>
     /// <param name="textHex">Code couleur hexadécimal RVB du texte.</param>
     public new XlsxWriteOptions<T> WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          base.WithHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     /// <summary>Encapsule la plage de données dans un objet Table Excel natif avec style et filtres automatiques.</summary>
     /// <param name="tableName">Nom logique unique du tableau Excel.</param>
     /// <param name="styleName">Nom du style visuel prédéfini Excel.</param>
     /// <param name="showRowStripes">Active l'alternance de couleur sur les lignes.</param>
     /// <param name="enableAutoFilter">Active les boutons de filtrage/tri sur chaque en-tête.</param>
     public new XlsxWriteOptions<T> WithExcelTable(
        string? tableName = null,
        string styleName = "TableStyleMedium2",
        bool showRowStripes = true,
        bool enableAutoFilter = true)
     {
          base.WithExcelTable(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     /// <summary>Active ou désactive le figeage de la première ligne d'en-tête.</summary>
     /// <param name="freeze">Indique si la ligne d'en-tête doit être figée.</param>
     public new XlsxWriteOptions<T> WithFreezeHeader(bool freeze = true)
     {
          base.WithFreezeHeader(freeze);
          return this;
     }

     /// <summary>Active ou désactive la génération automatique de listes déroulantes de validation.</summary>
     /// <param name="enable">Indique si les validations doivent être générées.</param>
     public new XlsxWriteOptions<T> WithDropdownValidation(bool enable = true)
     {
          base.WithDropdownValidation(enable);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à un nom d'en-tête d'export personnalisé sans réflexion.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété sélectionnée.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété (ex: <c>x => x.Prix</c>).</param>
     /// <param name="excelHeader">Le libellé de colonne à écrire dans Excel.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public XlsxWriteOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string excelHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          var propName = PropertyPathHelper.ExtractPropertyName(expression);
          SetTranslationInternal(propName, excelHeader);
          return this;
     }
}

/// <summary>
/// Options d'écriture fortement typées avec contexte métier externe <typeparamref name="TContext"/>.
/// </summary>
/// <typeparam name="T">Le type de l'entité sérialisée.</typeparam>
/// <typeparam name="TContext">Le type de l'objet de contexte partagé passé aux convertisseurs et accesseurs.</typeparam>
public sealed class XlsxWriteOptions<T, TContext> : XlsxWriteOptions<T>
{
     /// <summary>Obtient ou définit l'instance du contexte partagé injecté lors de l'écriture.</summary>
     public TContext Context { get; set; }

     /// <summary>Initialise les options avec le contexte métier spécifié.</summary>
     /// <param name="context">L'instance du contexte d'exécution.</param>
     public XlsxWriteOptions(TContext context)
     {
          Context = context;
     }

     /// <summary>Définit le libellé de l'onglet Excel créé.</summary>
     /// <param name="sheetName">Le nom de la feuille.</param>
     public new XlsxWriteOptions<T, TContext> WithSheetName(string sheetName) { base.WithSheetName(sheetName); return this; }

     /// <summary>Active ou désactive la génération de la ligne d'en-tête.</summary>
     /// <param name="writeHeader">Indique s'il faut inscrire les en-têtes.</param>
     public new XlsxWriteOptions<T, TContext> WithHeader(bool writeHeader) { base.WithHeader(writeHeader); return this; }

     /// <summary>Définit la culture appliquée au formatage.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public new XlsxWriteOptions<T, TContext> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }

     /// <summary>Définit les libellés textuels personnalisés écrits pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés employés pour vrai.</param>
     /// <param name="falseValues">Libellés employés pour faux.</param>
     public new XlsxWriteOptions<T, TContext> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }

     /// <summary>Injecte un ensemble de traductions d'en-têtes préconfigurées.</summary>
     /// <param name="translations">Dictionnaire associant une propriété à son libellé Excel.</param>
     public new XlsxWriteOptions<T, TContext> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }

     /// <summary>Définit le caractère séparateur décimal explicite.</summary>
     /// <param name="separator">Le séparateur décimal.</param>
     public new XlsxWriteOptions<T, TContext> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public new XlsxWriteOptions<T, TContext> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs d'écriture.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public new XlsxWriteOptions<T, TContext> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }

     /// <summary>Attache un observateur de progression de l'écriture.</summary>
     /// <param name="progress">Le canal de notification de progression.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement.</param>
     public new XlsxWriteOptions<T, TContext> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     /// <summary>Active ou désactive l'ajustement automatique de la largeur des colonnes.</summary>
     /// <param name="autoFit">Indique si le dimensionnement automatique est activé.</param>
     public new XlsxWriteOptions<T, TContext> WithAutoFitColumns(bool autoFit = true)
     {
          base.WithAutoFitColumns(autoFit);
          return this;
     }

     /// <summary>Personnalise le style graphique appliqué à la ligne d'en-tête.</summary>
     /// <param name="bold">Active le texte en gras.</param>
     /// <param name="backgroundHex">Code couleur hexadécimal RVB de fond.</param>
     /// <param name="textHex">Code couleur hexadécimal RVB du texte.</param>
     public new XlsxWriteOptions<T, TContext> WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          base.WithHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     /// <summary>Encapsule la plage de données dans un objet Table Excel natif avec style et filtres automatiques.</summary>
     /// <param name="tableName">Nom logique unique du tableau Excel.</param>
     /// <param name="styleName">Nom du style visuel prédéfini Excel.</param>
     /// <param name="showRowStripes">Active l'alternance de couleur sur les lignes.</param>
     /// <param name="enableAutoFilter">Active les boutons de filtrage/tri sur chaque en-tête.</param>
     public new XlsxWriteOptions<T, TContext> WithExcelTable(
        string? tableName = null,
        string styleName = "TableStyleMedium2",
        bool showRowStripes = true,
        bool enableAutoFilter = true)
     {
          base.WithExcelTable(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     /// <summary>Définit la politique de sécurité pour la neutralisation des formules.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public new XlsxWriteOptions<T, TContext> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     /// <summary>Active ou désactive le figeage de la première ligne d'en-tête.</summary>
     /// <param name="freeze">Indique si la ligne d'en-tête doit être figée.</param>
     public new XlsxWriteOptions<T, TContext> WithFreezeHeader(bool freeze = true)
     {
          base.WithFreezeHeader(freeze);
          return this;
     }

     /// <summary>Active ou désactive la génération automatique de listes déroulantes de validation.</summary>
     /// <param name="enable">Indique si les validations doivent être générées.</param>
     public new XlsxWriteOptions<T, TContext> WithDropdownValidation(bool enable = true)
     {
          base.WithDropdownValidation(enable);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à un nom d'en-tête d'export personnalisé sans réflexion.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété sélectionnée.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété.</param>
     /// <param name="excelHeader">Le libellé de colonne à écrire dans Excel.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public new XlsxWriteOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string excelHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, excelHeader, expression);
          return this;
     }
}

/// <summary>
/// Paramètres de mise en forme visuelle appliqués aux cellules de la ligne d'en-tête.
/// </summary>
/// <param name="Bold">Indique si le texte de l'en-tête doit être affiché en gras.</param>
/// <param name="BackgroundHex">Code hexadécimal RVB de la couleur de fond de remplissage (ex: "F2F2F2").</param>
/// <param name="TextHex">Code hexadécimal RVB de la couleur du texte (ex: "000000").</param>
public sealed record XlsxHeaderStyle(
    bool Bold = true,
    string? BackgroundHex = "F2F2F2",
    string? TextHex = "000000"
);

/// <summary>
/// Configuration d'un tableau natif Excel (Table / ListObject) OpenXML.
/// </summary>
/// <param name="TableName">Nom logique du tableau dans Excel (sans espaces, unique au sein du classeur).</param>
/// <param name="StyleName">Nom d'un style de tableau Excel prédéfini (ex: "TableStyleMedium2", "TableStyleLight1").</param>
/// <param name="ShowRowStripes">Active l'alternance de couleur de fond sur les lignes paires et impaires.</param>
/// <param name="EnableAutoFilter">Affiche les flèches déroulantes de filtrage et tri sur les en-têtes.</param>
public sealed record XlsxTableOptions(
    string? TableName = null,
    string StyleName = "TableStyleMedium2",
    bool ShowRowStripes = true,
    bool EnableAutoFilter = true
);