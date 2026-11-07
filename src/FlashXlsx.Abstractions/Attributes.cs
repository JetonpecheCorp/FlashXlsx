using System;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Marque un type (class, record, struct) pour la génération de code AOT FlashXlsx.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
public sealed class XlsxSerializableAttribute : Attribute;

/// <summary>
/// Définit le nom d'en-tête par défaut de la colonne dans le fichier Excel.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxNameAttribute(string name) : Attribute
{
     public string Name => name;
}

/// <summary>
/// Ignore systématiquement cette propriété lors de la lecture et de l'écriture.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxIgnoreAttribute : Attribute;

/// <summary>
/// Exige que la cellule contienne une valeur non vide à la lecture.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRequiredAttribute : Attribute;

/// <summary>
/// Exige que la cellule contienne une valeur uniquement si la condition spécifiée retourne true.
/// Le type doit implémenter IXlsxCondition&lt;T&gt; ou IXlsxCondition&lt;T, TContext&gt;.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRequiredIfAttribute(Type conditionType) : Attribute
{
     public Type ConditionType => conditionType;
}

/// <summary>
/// Ignore la propriété si la condition spécifiée retourne true.
/// Le type doit implémenter IXlsxCondition&lt;T&gt; ou IXlsxCondition&lt;T, TContext&gt;.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxIgnoreIfAttribute(Type conditionType) : Attribute
{
     public Type ConditionType => conditionType;
}

/// <summary>
/// Configure les valeurs textuelles reconnues pour un booléen.
/// Si un seul tableau est renseigné, toute valeur absente prend automatiquement la valeur du tableau vide.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxBooleanAttribute(string[]? trueValues = null, string[]? falseValues = null) : Attribute
{
     public string[]? TrueValues => trueValues;
     public string[]? FalseValues => falseValues;
}

public enum XlsxTextCase { None = 0, Upper = 1, Lower = 2, Title = 3 }
public enum XlsxTrimMode { None = 0, Both = 1, Start = 2, End = 3, AllWhitespace = 4 }
public enum XlsxEnumMode { Name = 0, Value = 1, Description = 2 }

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxCaseAttribute(XlsxTextCase textCase) : Attribute
{
     public XlsxTextCase Case { get; } = textCase;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxTrimAttribute(XlsxTrimMode mode = XlsxTrimMode.Both) : Attribute
{
     public XlsxTrimMode Mode { get; } = mode;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRangeAttribute(double min, double max) : Attribute
{
     public double Min { get; } = min;
     public double Max { get; } = max;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxMinAttribute(double min) : Attribute
{
     public double Min { get; } = min;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxMaxAttribute(double max) : Attribute
{
     public double Max { get; } = max;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxLengthAttribute : Attribute
{
     public int Min { get; }
     public int Max { get; }

     public XlsxLengthAttribute(int exactOrMax, bool isExact = false)
     {
          Min = isExact ? exactOrMax : 0;
          Max = exactOrMax;
     }

     public XlsxLengthAttribute(int min, int max)
     {
          Min = min;
          Max = max;
     }
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRegexAttribute(string pattern, bool ignoreCase = false) : Attribute
{
     public string Pattern { get; } = pattern;
     public bool IgnoreCase { get; } = ignoreCase;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxAllowedValuesAttribute(params string[] values) : Attribute
{
     public string[] Values { get; } = values;
     public bool IgnoreCase { get; set; } = true;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxColumnOrderAttribute(int order) : Attribute
{
     public int Order { get; } = order;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxFormulaAttribute(string formula) : Attribute
{
     public string Formula { get; } = formula;
}

[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxEnumAttribute(XlsxEnumMode mode = XlsxEnumMode.Name) : Attribute
{
     public XlsxEnumMode Mode { get; } = mode;
}

[AttributeUsage(AttributeTargets.Field)]
public sealed class XlsxEnumMemberAttribute(string label) : Attribute
{
     public string Label { get; } = label;
}

/// <summary>
/// Définit le séparateur décimal attendu ou écrit pour cette colonne (ex: ',' ou '.').
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxDecimalSeparatorAttribute(char separator) : Attribute
{
     public char Separator { get; } = separator;
}

public enum XlsxDefaultKind
{
     Literal = 0,
     Now = 1,
     UtcNow = 2,
     Today = 3,
     NewGuid = 4,
     EmptyString = 5
}

/// <summary>
/// Définit une valeur par défaut statique ou dynamique (Now, UtcNow, Today, NewGuid)
/// lorsque la cellule Excel est vide à la lecture (et optionnellement à l'écriture).
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxDefaultAttribute : Attribute
{
     public object? Value { get; }
     public XlsxDefaultKind Kind { get; }
     public bool ApplyOnExport { get; set; } = false;

     public XlsxDefaultAttribute(object? value)
     {
          Value = value;
          Kind = XlsxDefaultKind.Literal;
     }

     public XlsxDefaultAttribute(XlsxDefaultKind kind)
     {
          Kind = kind;
     }
}

/// <summary>
/// Spécifie un format exact de date, d'heure ou de nombre (ex: "dd/MM/yyyy" ou "0.00").
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxFormatAttribute(string format) : Attribute
{
     public string Format { get; } = format;
}

/// <summary>
/// Lie directement la propriété à une lettre de colonne Excel (ex: "A", "D", "AA"),
/// utile lorsque le fichier n'a pas de ligne d'en-têtes (HasHeaderRow = false).
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxColumnLetterAttribute(string columnLetter) : Attribute
{
     public string ColumnLetter => columnLetter;
}

/// <summary>
/// Permet de lire/écrire une collection (string[] ou List&lt;string&gt;) dans une seule cellule séparée par un délimiteur.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxSplitAttribute(char separator = ';') : Attribute
{
     public char Separator => separator;
     public bool TrimEntries { get; set; } = true;
     public bool RemoveEmptyEntries { get; set; } = true;
}
