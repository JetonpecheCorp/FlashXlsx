using FlashXlsx.Abstractions.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Options de lecture pour le mode dynamique (Dictionary) et socle commun pour toutes les lectures XLSX.
/// </summary>
public class XlsxReadOptions
{
     private readonly Dictionary<string, string[]> _translations = new(StringComparer.OrdinalIgnoreCase);

     /// <summary>Nom explicite de l'onglet à lire. Prioritaire sur <see cref="SheetIndex"/> si renseigné.</summary>
     public string? SheetName { get; set; }

     /// <summary>Index basé sur zéro de l'onglet à lire (par défaut : 0, soit le premier onglet).</summary>
     public int SheetIndex { get; set; } = 0;

     /// <summary>Indique si la feuille contient une ligne d'en-têtes définissant les noms de colonnes (activé par défaut).</summary>
     public bool HasHeaderRow { get; set; } = true;

     /// <summary>Numéro de ligne Excel (base 1) contenant les en-têtes (par défaut : 1).</summary>
     public int HeaderRowNumber { get; set; } = 1;

     /// <summary>Indique si les valeurs textuelles lues doivent être automatiquement tronquées de leurs espaces de début et de fin.</summary>
     public bool TrimStrings { get; set; } = true;

     /// <summary>Indique si le classeur utilise le système de date 1904 (spécifique aux anciens classeurs Excel Mac).</summary>
     public bool Use1904DateSystem { get; set; }

     /// <summary>Numéro de ligne Excel (base 1) à partir duquel commencer l'extraction des données.</summary>
     public int? StartRow { get; set; }

     /// <summary>Nombre maximal de lignes de données à lire avant d'interrompre le flux.</summary>
     public int? MaxRows { get; set; }

     /// <summary>Indique si les cellules fusionnées doivent répliquer leur valeur sur toutes les cellules de la plage fusionnée.</summary>
     public bool FillMergedCells { get; set; }

     /// <summary>Active la détection automatique de la ligne d'en-tête dans les premières lignes du fichier.</summary>
     public bool AutoDetectHeaderRow { get; set; }

     /// <summary>Nombre maximal de lignes inspectées lors de l'auto-détection d'en-tête (par défaut : 25).</summary>
     public int AutoDetectMaxRows { get; set; } = 25;

     /// <summary>Fournisseur de notification de progression recevant le nombre cumulé de lignes de données lues.</summary>
     public IProgress<int>? Progress { get; set; }

     /// <summary>Intervalle en nombre de lignes entre deux notifications de progression (par défaut : 1000).</summary>
     public int ProgressInterval { get; set; } = 1000;

     /// <summary>Caractère séparateur décimal attendu. '\0' pour auto-détection flexible (tolère '.' et ','), sinon caractère explicite.</summary>
     public char DecimalSeparator { get; set; } = '\0';

     /// <summary>Ignore automatiquement les lignes totalement vides (activé par défaut).</summary>
     public bool SkipEmptyRows { get; set; } = true;

     /// <summary>Culture utilisée pour l'analyse des formats de date, heure et nombres (par défaut : <see cref="CultureInfo.InvariantCulture"/>).</summary>
     public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

     /// <summary>Politique de sécurité appliquée lors de la lecture des cellules (protection contre l'injection de formules / payloads).</summary>
     public XlsxSecurityPolicy SecurityPolicy { get; set; } = XlsxSecurityPolicy.Sanitize;

     /// <summary>Comportement en cas d'erreur de parsing ou de validation (par défaut : arrêt à la première erreur).</summary>
     public XlsxErrorMode ErrorMode { get; set; } = XlsxErrorMode.StopOnFirstError;

     /// <summary>Gestionnaire appelé lors de la survenue d'une erreur de lecture sans allocation d'exception.</summary>
     public XlsxErrorHandler? OnError { get; set; }

     /// <summary>Libellés personnalisés interprétés comme <see langword="true"/> pour les colonnes booléennes.</summary>
     public string[]? DefaultTrueValues { get; private set; }

     /// <summary>Libellés personnalisés interprétés comme <see langword="false"/> pour les colonnes booléennes.</summary>
     public string[]? DefaultFalseValues { get; private set; }

     /// <summary>
     /// Mappe un nom de clé pivot vers un ou plusieurs alias d'en-tête Excel acceptés.
     /// </summary>
     public IReadOnlyDictionary<string, string[]> Translations => _translations;

     /// <summary>Sélectionne l'onglet cible par son nom.</summary>
     /// <param name="sheetName">Le nom de l'onglet à lire.</param>
     public XlsxReadOptions WithSheet(string sheetName) { SheetName = sheetName; return this; }

     /// <summary>Sélectionne l'onglet cible par son index basé sur zéro.</summary>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     public XlsxReadOptions WithSheet(int sheetIndex) { SheetIndex = sheetIndex; SheetName = null; return this; }

     /// <summary>Configure la présence et la position de la ligne d'en-tête.</summary>
     /// <param name="hasHeader">Indique si le document possède une ligne d'en-tête.</param>
     /// <param name="headerRowNumber">Le numéro de ligne Excel (base 1) contenant les en-têtes.</param>
     public XlsxReadOptions WithHeader(bool hasHeader, int headerRowNumber = 1) { HasHeaderRow = hasHeader; HeaderRowNumber = headerRowNumber; return this; }

     /// <summary>Définit la culture utilisée pour l'analyse des valeurs.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public XlsxReadOptions WithCulture(CultureInfo culture) { Culture = culture; return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public XlsxReadOptions WithErrorMode(XlsxErrorMode mode) { ErrorMode = mode; return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs sans lever d'exceptions.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public XlsxReadOptions WithErrorHandler(XlsxErrorHandler handler) { OnError = handler; return this; }

     /// <summary>Définit la politique de sécurité contre l'injection de formules malveillantes.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public XlsxReadOptions WithSecurityPolicy(XlsxSecurityPolicy policy) { SecurityPolicy = policy; return this; }

     /// <summary>Définit la ligne Excel à partir de laquelle démarrer la lecture des données.</summary>
     /// <param name="startRow">Numéro de ligne de début (base 1).</param>
     public XlsxReadOptions WithStartRow(int startRow)
     {
          StartRow = startRow > 0 ? startRow : null;
          return this;
     }

     /// <summary>Limite le nombre total de lignes de données lues.</summary>
     /// <param name="maxRows">Nombre maximal de lignes.</param>
     public XlsxReadOptions WithMaxRows(int maxRows)
     {
          MaxRows = maxRows >= 0 ? maxRows : null;
          return this;
     }

     /// <summary>Active ou désactive la propagation de valeur sur l'ensemble d'une plage fusionnée.</summary>
     /// <param name="enable">Indique si le remplissage des cellules fusionnées est activé.</param>
     public XlsxReadOptions WithFillMergedCells(bool enable = true)
     {
          FillMergedCells = enable;
          return this;
     }

     /// <summary>Configure l'omission automatique des lignes vierges.</summary>
     /// <param name="skip">Indique si les lignes entièrement vides doivent être ignorées.</param>
     public XlsxReadOptions WithSkipEmptyRows(bool skip = true)
     {
          SkipEmptyRows = skip;
          return this;
     }

     /// <summary>Définit les représentations textuelles acceptées pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés reconnus comme vrai.</param>
     /// <param name="falseValues">Libellés reconnus comme faux.</param>
     public XlsxReadOptions WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null)
     {
          DefaultTrueValues = trueValues;
          DefaultFalseValues = falseValues;
          return this;
     }

     /// <summary>Force un caractère séparateur décimal strict pour les nombres à virgule.</summary>
     /// <param name="separator">Le séparateur décimal (ex: ',' ou '.').</param>
     public XlsxReadOptions WithDecimalSeparator(char separator)
     {
          DecimalSeparator = separator;
          return this;
     }

     /// <summary>Active l'analyse automatique des premières lignes pour identifier la position de l'en-tête.</summary>
     /// <param name="enable">Active la recherche automatique.</param>
     /// <param name="maxScanRows">Nombre maximal de lignes analysées.</param>
     public XlsxReadOptions WithAutoDetectHeaderRow(bool enable = true, int maxScanRows = 25)
     {
          AutoDetectHeaderRow = enable;
          AutoDetectMaxRows = maxScanRows > 0 ? maxScanRows : 25;
          return this;
     }

     /// <summary>Attache un observateur de progression du flux de lecture.</summary>
     /// <param name="progress">Le canal de notification recevant le nombre de lignes traitées.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement en nombre de lignes.</param>
     public XlsxReadOptions WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          Progress = progress;
          ProgressInterval = intervalRows > 0 ? intervalRows : 1000;
          return this;
     }

     /// <summary>
     /// Ajoute une traduction ou un ensemble d'alias d'en-tête pour une clé dynamique.
     /// </summary>
     /// <param name="key">Le nom de la colonne ou propriété logique.</param>
     /// <param name="excelHeaders">Les libellés d'en-tête acceptés dans la feuille Excel.</param>
     public XlsxReadOptions Translate(string key, params string[] excelHeaders)
     {
          if (excelHeaders is { Length: > 0 })
               _translations[key] = excelHeaders;
          return this;
     }

     /// <summary>
     /// Injecte un dictionnaire complet d'alias d'en-têtes (ex: chargé depuis un référentiel ou JSON).
     /// </summary>
     /// <param name="translations">Dictionnaire associant une clé logique à son libellé Excel.</param>
     public XlsxReadOptions WithTranslations(IReadOnlyDictionary<string, string> translations)
     {
          foreach (var kvp in translations)
               _translations[kvp.Key] = [kvp.Value];
          return this;
     }

     internal void SetTranslationInternal(string propertyName, string[] excelHeaders)
     {
          if (excelHeaders is { Length: > 0 })
               _translations[propertyName] = excelHeaders;
     }
}

/// <summary>
/// Options de lecture fortement typées pour un modèle <typeparamref name="T"/> (100% Native AOT).
/// </summary>
/// <typeparam name="T">Le type de l'entité désérialisée.</typeparam>
public class XlsxReadOptions<T> : XlsxReadOptions
{
     /// <summary>Sélectionne l'onglet cible par son nom.</summary>
     /// <param name="sheetName">Le nom de l'onglet à lire.</param>
     public new XlsxReadOptions<T> WithSheet(string sheetName) { base.WithSheet(sheetName); return this; }

     /// <summary>Sélectionne l'onglet cible par son index basé sur zéro.</summary>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     public new XlsxReadOptions<T> WithSheet(int sheetIndex) { base.WithSheet(sheetIndex); return this; }

     /// <summary>Configure la présence et la position de la ligne d'en-tête.</summary>
     /// <param name="hasHeader">Indique si le document possède une ligne d'en-tête.</param>
     /// <param name="headerRowNumber">Le numéro de ligne Excel (base 1) contenant les en-têtes.</param>
     public new XlsxReadOptions<T> WithHeader(bool hasHeader, int headerRowNumber = 1) { base.WithHeader(hasHeader, headerRowNumber); return this; }

     /// <summary>Définit la culture utilisée pour l'analyse des valeurs.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public new XlsxReadOptions<T> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public new XlsxReadOptions<T> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs sans lever d'exceptions.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public new XlsxReadOptions<T> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }

     /// <summary>Définit les représentations textuelles acceptées pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés reconnus comme vrai.</param>
     /// <param name="falseValues">Libellés reconnus comme faux.</param>
     public new XlsxReadOptions<T> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }

     /// <summary>Injecte un dictionnaire complet d'alias d'en-têtes.</summary>
     /// <param name="translations">Dictionnaire associant une propriété à son libellé Excel.</param>
     public new XlsxReadOptions<T> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }

     /// <summary>Force un caractère séparateur décimal strict pour les nombres à virgule.</summary>
     /// <param name="separator">Le séparateur décimal.</param>
     public new XlsxReadOptions<T> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }

     /// <summary>Définit la politique de sécurité contre l'injection de formules.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public new XlsxReadOptions<T> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     /// <summary>Définit la ligne Excel à partir de laquelle démarrer la lecture des données.</summary>
     /// <param name="startRow">Numéro de ligne de début (base 1).</param>
     public new XlsxReadOptions<T> WithStartRow(int startRow)
     {
          base.WithStartRow(startRow);
          return this;
     }

     /// <summary>Limite le nombre total de lignes de données lues.</summary>
     /// <param name="maxRows">Nombre maximal de lignes.</param>
     public new XlsxReadOptions<T> WithMaxRows(int maxRows)
     {
          base.WithMaxRows(maxRows);
          return this;
     }

     /// <summary>Active ou désactive la propagation de valeur sur l'ensemble d'une plage fusionnée.</summary>
     /// <param name="enable">Indique si le remplissage des cellules fusionnées est activé.</param>
     public new XlsxReadOptions<T> WithFillMergedCells(bool enable = true)
     {
          base.WithFillMergedCells(enable);
          return this;
     }

     /// <summary>Configure l'omission automatique des lignes vierges.</summary>
     /// <param name="skip">Indique si les lignes vides doivent être ignorées.</param>
     public new XlsxReadOptions<T> WithSkipEmptyRows(bool skip = true)
     {
          base.WithSkipEmptyRows(skip);
          return this;
     }

     /// <summary>Active l'analyse automatique des premières lignes pour identifier la position de l'en-tête.</summary>
     /// <param name="enable">Active la recherche automatique.</param>
     /// <param name="maxScanRows">Nombre maximal de lignes analysées.</param>
     public new XlsxReadOptions<T> WithAutoDetectHeaderRow(bool enable = true, int maxScanRows = 25)
     {
          base.WithAutoDetectHeaderRow(enable, maxScanRows);
          return this;
     }

     /// <summary>Attache un observateur de progression du flux de lecture.</summary>
     /// <param name="progress">Le canal de notification de progression.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement.</param>
     public new XlsxReadOptions<T> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à un ou plusieurs noms d'en-tête Excel sans réflexion.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété sélectionnée.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété (ex: <c>x => x.Nom</c>).</param>
     /// <param name="excelHeaders">Liste des alias d'en-tête Excel reconnus pour cette propriété.</param>
     public XlsxReadOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         params string[] excelHeaders)
         => TranslateCore(excelHeaders, null);

     /// <summary>
     /// Associe une propriété du modèle à un nom d'en-tête Excel en capturant automatiquement le nom de la propriété.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété.</param>
     /// <param name="firstHeader">Le nom de colonne Excel correspondant.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public XlsxReadOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string firstHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          var propName = PropertyPathHelper.ExtractPropertyName(expression);
          SetTranslationInternal(propName, [firstHeader]);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à une liste d'alias d'en-têtes Excel en capturant le nom de la propriété.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété.</param>
     /// <param name="excelHeaders">Tableau des libellés d'en-têtes Excel correspondants.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public XlsxReadOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string[] excelHeaders,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          var propName = PropertyPathHelper.ExtractPropertyName(expression);
          SetTranslationInternal(propName, excelHeaders);
          return this;
     }

     private XlsxReadOptions<T> TranslateCore(string[] excelHeaders, string? expression)
     {
          if (expression is not null)
               SetTranslationInternal(PropertyPathHelper.ExtractPropertyName(expression), excelHeaders);
          return this;
     }
}

/// <summary>
/// Options de lecture fortement typées avec contexte métier externe <typeparamref name="TContext"/>.
/// </summary>
/// <typeparam name="T">Le type de l'entité désérialisée.</typeparam>
/// <typeparam name="TContext">Le type du contexte partagé passé aux convertisseurs et conditions.</typeparam>
public sealed class XlsxReadOptions<T, TContext> : XlsxReadOptions<T>
{
     /// <summary>Obtient ou définit l'instance du contexte partagé injecté lors de la lecture.</summary>
     public TContext Context { get; set; }

     /// <summary>Initialise les options avec le contexte métier spécifié.</summary>
     /// <param name="context">L'instance du contexte d'exécution.</param>
     public XlsxReadOptions(TContext context)
     {
          Context = context;
     }

     /// <summary>Sélectionne l'onglet cible par son nom.</summary>
     /// <param name="sheetName">Le nom de l'onglet.</param>
     public new XlsxReadOptions<T, TContext> WithSheet(string sheetName) { base.WithSheet(sheetName); return this; }

     /// <summary>Sélectionne l'onglet cible par son index basé sur zéro.</summary>
     /// <param name="sheetIndex">L'index de l'onglet (base 0).</param>
     public new XlsxReadOptions<T, TContext> WithSheet(int sheetIndex) { base.WithSheet(sheetIndex); return this; }

     /// <summary>Configure la présence et la position de la ligne d'en-tête.</summary>
     /// <param name="hasHeader">Indique si le document possède une ligne d'en-tête.</param>
     /// <param name="headerRowNumber">Le numéro de ligne Excel (base 1) contenant les en-têtes.</param>
     public new XlsxReadOptions<T, TContext> WithHeader(bool hasHeader, int headerRowNumber = 1) { base.WithHeader(hasHeader, headerRowNumber); return this; }

     /// <summary>Définit la culture utilisée pour l'analyse des valeurs.</summary>
     /// <param name="culture">La culture à appliquer.</param>
     public new XlsxReadOptions<T, TContext> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }

     /// <summary>Définit la stratégie d'interruption ou de poursuite en cas d'erreur.</summary>
     /// <param name="mode">Le mode d'erreur souhaité.</param>
     public new XlsxReadOptions<T, TContext> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }

     /// <summary>Enregistre un gestionnaire de callback pour intercepter les erreurs sans lever d'exceptions.</summary>
     /// <param name="handler">Le délégué de traitement d'erreur.</param>
     public new XlsxReadOptions<T, TContext> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }

     /// <summary>Définit les représentations textuelles acceptées pour les valeurs booléennes.</summary>
     /// <param name="trueValues">Libellés reconnus comme vrai.</param>
     /// <param name="falseValues">Libellés reconnus comme faux.</param>
     public new XlsxReadOptions<T, TContext> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }

     /// <summary>Injecte un dictionnaire complet d'alias d'en-têtes.</summary>
     /// <param name="translations">Dictionnaire associant une propriété à son libellé Excel.</param>
     public new XlsxReadOptions<T, TContext> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }

     /// <summary>Force un caractère séparateur décimal strict pour les nombres à virgule.</summary>
     /// <param name="separator">Le séparateur décimal.</param>
     public new XlsxReadOptions<T, TContext> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }

     /// <summary>Configure l'omission automatique des lignes vierges.</summary>
     /// <param name="skip">Indique si les lignes vides doivent être ignorées.</param>
     public new XlsxReadOptions<T, TContext> WithSkipEmptyRows(bool skip = true)
     {
          base.WithSkipEmptyRows(skip);
          return this;
     }

     /// <summary>Définit la ligne Excel à partir de laquelle démarrer la lecture des données.</summary>
     /// <param name="startRow">Numéro de ligne de début (base 1).</param>
     public new XlsxReadOptions<T, TContext> WithStartRow(int startRow)
     {
          base.WithStartRow(startRow);
          return this;
     }

     /// <summary>Limite le nombre total de lignes de données lues.</summary>
     /// <param name="maxRows">Nombre maximal de lignes.</param>
     public new XlsxReadOptions<T, TContext> WithMaxRows(int maxRows)
     {
          base.WithMaxRows(maxRows);
          return this;
     }

     /// <summary>Active ou désactive la propagation de valeur sur l'ensemble d'une plage fusionnée.</summary>
     /// <param name="enable">Indique si le remplissage des cellules fusionnées est activé.</param>
     public new XlsxReadOptions<T, TContext> WithFillMergedCells(bool enable = true)
     {
          base.WithFillMergedCells(enable);
          return this;
     }

     /// <summary>Définit la politique de sécurité contre l'injection de formules.</summary>
     /// <param name="policy">La politique de sécurité à appliquer.</param>
     public new XlsxReadOptions<T, TContext> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     /// <summary>Active l'analyse automatique des premières lignes pour identifier la position de l'en-tête.</summary>
     /// <param name="enable">Active la recherche automatique.</param>
     /// <param name="maxScanRows">Nombre maximal de lignes analysées.</param>
     public new XlsxReadOptions<T, TContext> WithAutoDetectHeaderRow(bool enable = true, int maxScanRows = 25)
     {
          base.WithAutoDetectHeaderRow(enable, maxScanRows);
          return this;
     }

     /// <summary>Attache un observateur de progression du flux de lecture.</summary>
     /// <param name="progress">Le canal de notification de progression.</param>
     /// <param name="intervalRows">La fréquence d'émission de l'événement.</param>
     public new XlsxReadOptions<T, TContext> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à un nom d'en-tête Excel en capturant le nom de la propriété.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété.</param>
     /// <param name="firstHeader">Le nom de colonne Excel correspondant.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public new XlsxReadOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string firstHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, firstHeader, expression);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à une liste d'alias d'en-têtes Excel en capturant le nom de la propriété.
     /// </summary>
     /// <typeparam name="TProp">Le type de la propriété.</typeparam>
     /// <param name="propertySelector">Expression lambda sélectionnant la propriété.</param>
     /// <param name="excelHeaders">Tableau des libellés d'en-têtes Excel correspondants.</param>
     /// <param name="expression">Nom de l'expression capturé par le compilateur.</param>
     public new XlsxReadOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string[] excelHeaders,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, excelHeaders, expression);
          return this;
     }
}