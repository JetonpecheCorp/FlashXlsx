using FlashXlsx.Abstractions.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Options de lecture pour le mode dynamique (Dictionary) et socle commun.
/// </summary>
public class XlsxReadOptions
{
     private readonly Dictionary<string, string[]> _translations = new(StringComparer.OrdinalIgnoreCase);

     public string? SheetName { get; set; }
     public int SheetIndex { get; set; } = 0;
     public bool HasHeaderRow { get; set; } = true;
     public int HeaderRowNumber { get; set; } = 1;
     public bool TrimStrings { get; set; } = true;
     public bool Use1904DateSystem { get; set; }

     /// <summary>'\0' = auto-détection flexible ('.' et ','), sinon caractère explicite (ex: ',' ou '.').</summary>
     public char DecimalSeparator { get; set; } = '\0';

     /// <summary>Ignore automatiquement les lignes totalement vides (activé par défaut).</summary>
     public bool SkipEmptyRows { get; set; } = true;
     public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
     public XlsxSecurityPolicy SecurityPolicy { get; set; } = XlsxSecurityPolicy.Sanitize;

     public XlsxErrorMode ErrorMode { get; set; } = XlsxErrorMode.StopOnFirstError;
     public XlsxErrorHandler? OnError { get; set; }

     public string[]? DefaultTrueValues { get; private set; }
     public string[]? DefaultFalseValues { get; private set; }

     /// <summary>
     /// Mappe un nom de clé pivot vers un ou plusieurs alias d'en-tête Excel.
     /// </summary>
     public IReadOnlyDictionary<string, string[]> Translations => _translations;

     public XlsxReadOptions WithSheet(string sheetName) { SheetName = sheetName; return this; }
     public XlsxReadOptions WithSheet(int sheetIndex) { SheetIndex = sheetIndex; SheetName = null; return this; }
     public XlsxReadOptions WithHeader(bool hasHeader, int headerRowNumber = 1) { HasHeaderRow = hasHeader; HeaderRowNumber = headerRowNumber; return this; }
     public XlsxReadOptions WithCulture(CultureInfo culture) { Culture = culture; return this; }
     public XlsxReadOptions WithErrorMode(XlsxErrorMode mode) { ErrorMode = mode; return this; }
     public XlsxReadOptions WithErrorHandler(XlsxErrorHandler handler) { OnError = handler; return this; }
     public XlsxReadOptions WithSecurityPolicy(XlsxSecurityPolicy policy) { SecurityPolicy = policy; return this; }

     public XlsxReadOptions WithSkipEmptyRows(bool skip = true)
     {
          SkipEmptyRows = skip;
          return this;
     }

     public XlsxReadOptions WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null)
     {
          DefaultTrueValues = trueValues;
          DefaultFalseValues = falseValues;
          return this;
     }

     public XlsxReadOptions WithDecimalSeparator(char separator)
     {
          DecimalSeparator = separator;
          return this;
     }

     /// <summary>
     /// Ajoute une traduction d'en-tête par nom de clé (mode dynamique ou BDD).
     /// </summary>
     public XlsxReadOptions Translate(string key, params string[] excelHeaders)
     {
          if (excelHeaders is { Length: > 0 })
               _translations[key] = excelHeaders;
          return this;
     }

     /// <summary>
     /// Injecte un dictionnaire de traductions complet (ex: chargé depuis une BDD ou un fichier JSON).
     /// </summary>
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
/// Options de lecture fortement typées pour un modèle T (100% Native AOT).
/// </summary>
public class XlsxReadOptions<T> : XlsxReadOptions
{
     public new XlsxReadOptions<T> WithSheet(string sheetName) { base.WithSheet(sheetName); return this; }
     public new XlsxReadOptions<T> WithSheet(int sheetIndex) { base.WithSheet(sheetIndex); return this; }
     public new XlsxReadOptions<T> WithHeader(bool hasHeader, int headerRowNumber = 1) { base.WithHeader(hasHeader, headerRowNumber); return this; }
     public new XlsxReadOptions<T> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }
     public new XlsxReadOptions<T> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }
     public new XlsxReadOptions<T> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }
     public new XlsxReadOptions<T> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }
     public new XlsxReadOptions<T> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }
     public new XlsxReadOptions<T> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }
     public new XlsxReadOptions<T> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     public new XlsxReadOptions<T> WithSkipEmptyRows(bool skip = true)
     {
          base.WithSkipEmptyRows(skip);
          return this;
     }

     /// <summary>
     /// Associe une propriété du modèle à un ou plusieurs noms d'en-tête Excel sans réflexion.
     /// </summary>
     public XlsxReadOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         params string[] excelHeaders)
         => TranslateCore(excelHeaders, null);

     // Surcharge capturant l'expression C# via CallerArgumentExpression
     public XlsxReadOptions<T> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string firstHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          var propName = PropertyPathHelper.ExtractPropertyName(expression);
          SetTranslationInternal(propName, [firstHeader]);
          return this;
     }

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
/// Options de lecture fortement typées avec contexte métier externe TContext.
/// </summary>
public sealed class XlsxReadOptions<T, TContext> : XlsxReadOptions<T>
{
     public TContext Context { get; set; }

     public XlsxReadOptions(TContext context)
     {
          Context = context;
     }

     public new XlsxReadOptions<T, TContext> WithSheet(string sheetName) { base.WithSheet(sheetName); return this; }
     public new XlsxReadOptions<T, TContext> WithSheet(int sheetIndex) { base.WithSheet(sheetIndex); return this; }
     public new XlsxReadOptions<T, TContext> WithHeader(bool hasHeader, int headerRowNumber = 1) { base.WithHeader(hasHeader, headerRowNumber); return this; }
     public new XlsxReadOptions<T, TContext> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }
     public new XlsxReadOptions<T, TContext> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }
     public new XlsxReadOptions<T, TContext> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }
     public new XlsxReadOptions<T, TContext> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }
     public new XlsxReadOptions<T, TContext> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }
     public new XlsxReadOptions<T, TContext> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }

     public new XlsxReadOptions<T, TContext> WithSkipEmptyRows(bool skip = true)
     {
          base.WithSkipEmptyRows(skip);
          return this;
     }

     public new XlsxReadOptions<T, TContext> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }
     public new XlsxReadOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string firstHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, firstHeader, expression);
          return this;
     }

     public new XlsxReadOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string[] excelHeaders,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, excelHeaders, expression);
          return this;
     }
}