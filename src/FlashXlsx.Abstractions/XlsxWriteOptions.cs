using FlashXlsx.Abstractions.Internal;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Options d'écriture pour le mode dynamique (Dictionary) et socle commun.
/// </summary>
public class XlsxWriteOptions
{
     private readonly Dictionary<string, string> _translations = new(StringComparer.OrdinalIgnoreCase);

     public string SheetName { get; set; } = "Sheet1";
     public bool WriteHeader { get; set; } = true;
     public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
     /// <summary>Séparateur décimal utilisé lorsqu'une colonne décimale est formatée avec un motif texte ([XlsxFormat]) ou un séparateur explicite.</summary>
     public char DecimalSeparator { get; set; } = '.';
     public bool AutoFitColumns { get; set; }
     public bool FreezeHeader { get; set; }
     public bool DropdownValidation { get; set; }
     public IProgress<int>? Progress { get; set; }
     public int ProgressInterval { get; set; } = 1000;
     public XlsxErrorMode ErrorMode { get; set; } = XlsxErrorMode.StopOnFirstError;
     public XlsxErrorHandler? OnError { get; set; }
     public XlsxHeaderStyle? HeaderStyle { get; set; }
     public XlsxTableOptions? ExcelTable { get; set; }

     public string[]? DefaultTrueValues { get; private set; }
     public string[]? DefaultFalseValues { get; private set; }


     /// <summary>
     /// Mappe un nom de propriété/clé vers l'en-tête traduit à écrire dans Excel.
     /// </summary>
     public IReadOnlyDictionary<string, string> Translations => _translations;
     public XlsxSecurityPolicy SecurityPolicy { get; set; } = XlsxSecurityPolicy.Sanitize;

     public XlsxWriteOptions WithSecurityPolicy(XlsxSecurityPolicy policy) { SecurityPolicy = policy; return this; }
     public XlsxWriteOptions WithSheetName(string sheetName) { SheetName = sheetName; return this; }
     public XlsxWriteOptions WithHeader(bool writeHeader) { WriteHeader = writeHeader; return this; }
     public XlsxWriteOptions WithCulture(CultureInfo culture) { Culture = culture; return this; }
     public XlsxWriteOptions WithErrorMode(XlsxErrorMode mode) { ErrorMode = mode; return this; }
     public XlsxWriteOptions WithErrorHandler(XlsxErrorHandler handler) { OnError = handler; return this; }
     public XlsxWriteOptions WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          Progress = progress;
          ProgressInterval = intervalRows > 0 ? intervalRows : 1000;
          return this;
     }

     public XlsxWriteOptions WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null)
     {
          DefaultTrueValues = trueValues;
          DefaultFalseValues = falseValues;
          return this;
     }

     public XlsxWriteOptions WithDecimalSeparator(char separator)
     {
          DecimalSeparator = separator;
          return this;
     }

     public XlsxWriteOptions WithAutoFitColumns(bool autoFit = true)
     {
          AutoFitColumns = autoFit;
          return this;
     }

     public XlsxWriteOptions WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          HeaderStyle = new XlsxHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     public XlsxWriteOptions WithExcelTable(
         string? tableName = null,
         string styleName = "TableStyleMedium2",
         bool showRowStripes = true,
         bool enableAutoFilter = true)
     {
          ExcelTable = new XlsxTableOptions(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     public XlsxWriteOptions WithFreezeHeader(bool freeze = true)
     {
          FreezeHeader = freeze;
          return this;
     }

     public XlsxWriteOptions WithDropdownValidation(bool enable = true)
     {
          DropdownValidation = enable;
          return this;
     }

     public XlsxWriteOptions Translate(string key, string excelHeader)
     {
          _translations[key] = excelHeader;
          return this;
     }

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
/// Options d'écriture fortement typées pour un modèle T (100% Native AOT).
/// </summary>
public class XlsxWriteOptions<T> : XlsxWriteOptions
{
     public new XlsxWriteOptions<T> WithSheetName(string sheetName) { base.WithSheetName(sheetName); return this; }
     public new XlsxWriteOptions<T> WithHeader(bool writeHeader) { base.WithHeader(writeHeader); return this; }
     public new XlsxWriteOptions<T> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }
     public new XlsxWriteOptions<T> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }
     public new XlsxWriteOptions<T> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }
     public new XlsxWriteOptions<T> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }
     public new XlsxWriteOptions<T> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }
     public new XlsxWriteOptions<T> WithTranslations(IReadOnlyDictionary<string, string> translations) 
     { 
          base.WithTranslations(translations); 
          return this; 
     }

     public new XlsxWriteOptions<T> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     public new XlsxWriteOptions<T> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     public new XlsxWriteOptions<T> WithAutoFitColumns(bool autoFit = true)
     {
          base.WithAutoFitColumns(autoFit);
          return this;
     }

     public new XlsxWriteOptions<T> WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          base.WithHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     public new XlsxWriteOptions<T> WithExcelTable(
        string? tableName = null,
        string styleName = "TableStyleMedium2",
        bool showRowStripes = true,
        bool enableAutoFilter = true)
     {
          base.WithExcelTable(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     public new XlsxWriteOptions<T> WithFreezeHeader(bool freeze = true)
     {
          base.WithFreezeHeader(freeze);
          return this;
     }

     public new XlsxWriteOptions<T> WithDropdownValidation(bool enable = true)
     {
          base.WithDropdownValidation(enable);
          return this;
     }

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
/// Options d'écriture fortement typées avec contexte métier externe TContext.
/// </summary>
public sealed class XlsxWriteOptions<T, TContext> : XlsxWriteOptions<T>
{
     public TContext Context { get; set; }

     public XlsxWriteOptions(TContext context)
     {
          Context = context;
     }
     public new XlsxWriteOptions<T, TContext> WithSheetName(string sheetName) { base.WithSheetName(sheetName); return this; }
     public new XlsxWriteOptions<T, TContext> WithHeader(bool writeHeader) { base.WithHeader(writeHeader); return this; }
     public new XlsxWriteOptions<T, TContext> WithCulture(CultureInfo culture) { base.WithCulture(culture); return this; }
     public new XlsxWriteOptions<T, TContext> WithDefaultBooleanValues(string[]? trueValues = null, string[]? falseValues = null) { base.WithDefaultBooleanValues(trueValues, falseValues); return this; }
     public new XlsxWriteOptions<T, TContext> WithTranslations(IReadOnlyDictionary<string, string> translations) { base.WithTranslations(translations); return this; }
     public new XlsxWriteOptions<T, TContext> WithDecimalSeparator(char separator) { base.WithDecimalSeparator(separator); return this; }
     public new XlsxWriteOptions<T, TContext> WithErrorMode(XlsxErrorMode mode) { base.WithErrorMode(mode); return this; }
     public new XlsxWriteOptions<T, TContext> WithErrorHandler(XlsxErrorHandler handler) { base.WithErrorHandler(handler); return this; }
     public new XlsxWriteOptions<T, TContext> WithProgress(IProgress<int> progress, int intervalRows = 1000)
     {
          base.WithProgress(progress, intervalRows);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithAutoFitColumns(bool autoFit = true)
     {
           base.WithAutoFitColumns(autoFit);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithHeaderStyle(bool bold = true, string? backgroundHex = "F2F2F2", string? textHex = "000000")
     {
          base.WithHeaderStyle(bold, backgroundHex, textHex);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithExcelTable(
        string? tableName = null,
        string styleName = "TableStyleMedium2",
        bool showRowStripes = true,
        bool enableAutoFilter = true)
     {
          base.WithExcelTable(tableName, styleName, showRowStripes, enableAutoFilter);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithSecurityPolicy(XlsxSecurityPolicy policy)
     {
          base.WithSecurityPolicy(policy);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithFreezeHeader(bool freeze = true)
     {
          base.WithFreezeHeader(freeze);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> WithDropdownValidation(bool enable = true)
     {
          base.WithDropdownValidation(enable);
          return this;
     }

     public new XlsxWriteOptions<T, TContext> Translate<TProp>(
         Func<T, TProp> propertySelector,
         string excelHeader,
         [CallerArgumentExpression(nameof(propertySelector))] string? expression = null)
     {
          base.Translate(propertySelector, excelHeader, expression);
          return this;
     }
}

public sealed record XlsxHeaderStyle(
    bool Bold = true,
    string? BackgroundHex = "F2F2F2",
    string? TextHex = "000000"
);

public sealed record XlsxTableOptions(
    string? TableName = null,
    string StyleName = "TableStyleMedium2",
    bool ShowRowStripes = true,
    bool EnableAutoFilter = true
);
