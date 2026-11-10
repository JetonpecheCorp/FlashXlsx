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
/// <param name="name">Le titre de l'en-tête attendu ou généré pour cette propriété.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxNameAttribute(string name) : Attribute
{
     /// <summary>
     /// Obtient le libellé de la colonne associé à cette propriété.
     /// </summary>
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
/// Le type doit implémenter <c>IXlsxCondition&lt;T&gt;</c> ou <c>IXlsxCondition&lt;T, TContext&gt;</c>.
/// </summary>
/// <param name="conditionType">Le type CLR évaluant la condition de présence obligatoire.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRequiredIfAttribute(Type conditionType) : Attribute
{
     /// <summary>
     /// Obtient le type qui évalue la condition.
     /// </summary>
     public Type ConditionType => conditionType;
}

/// <summary>
/// Ignore la propriété si la condition spécifiée retourne true.
/// Le type doit implémenter <c>IXlsxCondition&lt;T&gt;</c> ou <c>IXlsxCondition&lt;T, TContext&gt;</c>.
/// </summary>
/// <param name="conditionType">Le type CLR évaluant la condition d'exclusion.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxIgnoreIfAttribute(Type conditionType) : Attribute
{
     /// <summary>
     /// Obtient le type qui évalue la condition.
     /// </summary>
     public Type ConditionType => conditionType;
}

/// <summary>
/// Configure les valeurs textuelles reconnues pour un booléen.
/// Si un seul tableau est renseigné, toute valeur absente prend automatiquement la valeur du tableau vide.
/// </summary>
/// <param name="trueValues">Ensemble des libellés interprétés comme <see langword="true"/> (ex: "OUI", "VRAI", "1").</param>
/// <param name="falseValues">Ensemble des libellés interprétés comme <see langword="false"/> (ex: "NON", "FAUX", "0").</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxBooleanAttribute(string[]? trueValues = null, string[]? falseValues = null) : Attribute
{
     /// <summary>
     /// Obtient les libellés assimilés à la valeur booléenne <see langword="true"/>.
     /// </summary>
     public string[]? TrueValues => trueValues;

     /// <summary>
     /// Obtient les libellés assimilés à la valeur booléenne <see langword="false"/>.
     /// </summary>
     public string[]? FalseValues => falseValues;
}

/// <summary>
/// Mode de conversion de la casse appliqué à une chaîne de caractères.
/// </summary>
public enum XlsxTextCase
{
     /// <summary>
     /// Aucun changement de casse.
     /// </summary>
     None = 0,

     /// <summary>
     /// Convertit tout le texte en majuscules (ex: "EXEMPLE").
     /// </summary>
     Upper = 1,

     /// <summary>
     /// Convertit tout le texte en minuscules (ex: "exemple").
     /// </summary>
     Lower = 2,

     /// <summary>
     /// Première lettre de chaque mot en majuscule (Title Case).
     /// </summary>
     Title = 3
}

/// <summary>
/// Définit la stratégie de rognage des espaces pour les valeurs textuelles.
/// </summary>
public enum XlsxTrimMode
{
     /// <summary>
     /// Conserve la chaîne telle quelle avec ses espaces d'origine.
     /// </summary>
     None = 0,

     /// <summary>
     /// Supprime les espaces en début et fin de chaîne.
     /// </summary>
     Both = 1,

     /// <summary>
     /// Supprime les espaces au début de la chaîne uniquement.
     /// </summary>
     Start = 2,

     /// <summary>
     /// Supprime les espaces à la fin de la chaîne uniquement.
     /// </summary>
     End = 3,

     /// <summary>
     /// Supprime tous les caractères blancs (y compris internes).
     /// </summary>
     AllWhitespace = 4
}

/// <summary>
/// Stratégie de sérialisation et de désérialisation pour les énumérations.
/// </summary>
public enum XlsxEnumMode
{
     /// <summary>
     /// Utilise le nom textuel du membre de l'enum.
     /// </summary>
     Name = 0,

     /// <summary>
     /// Utilise la valeur sous-jacente (numérique).
     /// </summary>
     Value = 1,

     /// <summary>
     /// Utilise le libellé personnalisé fourni via <see cref="XlsxEnumMemberAttribute"/>.
     /// </summary>
     Description = 2
}

/// <summary>
/// Applique une transformation de casse automatique sur la valeur textuelle lue ou écrite.
/// </summary>
/// <param name="textCase">Le type de casse à appliquer.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxCaseAttribute(XlsxTextCase textCase) : Attribute
{
     /// <summary>
     /// Obtient la transformation de casse configurée.
     /// </summary>
     public XlsxTextCase Case { get; } = textCase;
}

/// <summary>
/// Spécifie la stratégie de nettoyage des espaces sur les colonnes de type texte.
/// </summary>
/// <param name="mode">Le mode de rognage appliqué (par défaut : début et fin).</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxTrimAttribute(XlsxTrimMode mode = XlsxTrimMode.Both) : Attribute
{
     /// <summary>
     /// Obtient le mode de rognage configuré.
     /// </summary>
     public XlsxTrimMode Mode { get; } = mode;
}

/// <summary>
/// Valide que la valeur numérique est comprise dans un intervalle fermé [min, max].
/// </summary>
/// <param name="min">La borne inférieure autorisée.</param>
/// <param name="max">La borne supérieure autorisée.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRangeAttribute(double min, double max) : Attribute
{
     /// <summary>
     /// Obtient la borne minimale autorisée.
     /// </summary>
     public double Min { get; } = min;

     /// <summary>
     /// Obtient la borne maximale autorisée.
     /// </summary>
     public double Max { get; } = max;
}

/// <summary>
/// Valide que la valeur numérique est supérieure ou égale à un minimum spécifié.
/// </summary>
/// <param name="min">La valeur minimale admissible.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxMinAttribute(double min) : Attribute
{
     /// <summary>
     /// Obtient la valeur minimale admissible.
     /// </summary>
     public double Min { get; } = min;
}

/// <summary>
/// Valide que la valeur numérique est inférieure ou égale à un maximum spécifié.
/// </summary>
/// <param name="max">La valeur maximale admissible.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxMaxAttribute(double max) : Attribute
{
     /// <summary>
     /// Obtient la valeur maximale admissible.
     /// </summary>
     public double Max { get; } = max;
}

/// <summary>
/// Valide la longueur en nombre de caractères d'une chaîne de texte.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxLengthAttribute : Attribute
{
     /// <summary>
     /// Obtient la longueur minimale requise.
     /// </summary>
     public int Min { get; }

     /// <summary>
     /// Obtient la longueur maximale autorisée.
     /// </summary>
     public int Max { get; }

     /// <summary>
     /// Initialise une nouvelle contrainte de longueur exacte ou maximale.
     /// </summary>
     /// <param name="exactOrMax">La longueur maximale, ou exacte si <paramref name="isExact"/> vaut true.</param>
     /// <param name="isExact">Indique si la longueur doit être strictement égale à la valeur renseignée.</param>
     public XlsxLengthAttribute(int exactOrMax, bool isExact = false)
     {
          Min = isExact ? exactOrMax : 0;
          Max = exactOrMax;
     }

     /// <summary>
     /// Initialise une nouvelle contrainte d'intervalle de longueur [min, max].
     /// </summary>
     /// <param name="min">La longueur minimale requise.</param>
     /// <param name="max">La longueur maximale autorisée.</param>
     public XlsxLengthAttribute(int min, int max)
     {
          Min = min;
          Max = max;
     }
}

/// <summary>
/// Valide que la valeur textuelle correspond à l'expression régulière donnée.
/// </summary>
/// <param name="pattern">Le motif regex à valider.</param>
/// <param name="ignoreCase">Indique si la casse doit être ignorée lors de la comparaison.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxRegexAttribute(string pattern, bool ignoreCase = false) : Attribute
{
     /// <summary>
     /// Obtient l'expression régulière à tester.
     /// </summary>
     public string Pattern { get; } = pattern;

     /// <summary>
     /// Obtient une valeur indiquant si la comparaison est insensible à la casse.
     /// </summary>
     public bool IgnoreCase { get; } = ignoreCase;
}

/// <summary>
/// Restreint les valeurs autorisées pour la cellule à une liste prédéfinie.
/// </summary>
/// <param name="values">L'ensemble des chaînes admissibles.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxAllowedValuesAttribute(params string[] values) : Attribute
{
     /// <summary>
     /// Obtient la liste des valeurs admises.
     /// </summary>
     public string[] Values { get; } = values;

     /// <summary>
     /// Obtient ou définit si la validation ignore la distinction majuscules/minuscules (vaut <see langword="true"/> par défaut).
     /// </summary>
     public bool IgnoreCase { get; set; } = true;
}

/// <summary>
/// Définit la position et l'ordre d'affichage de la colonne à l'exportation.
/// </summary>
/// <param name="order">L'indice de positionnement (croissant).</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxColumnOrderAttribute(int order) : Attribute
{
     /// <summary>
     /// Obtient la position ordonnée de la colonne.
     /// </summary>
     public int Order { get; } = order;
}

/// <summary>
/// Injecte une formule de calcul Excel native dans la cellule lors de l'exportation.
/// </summary>
/// <param name="formula">La formule Excel (sans le symbole '=' initial, ex: "SUM(A{row}:B{row})").</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxFormulaAttribute(string formula) : Attribute
{
     /// <summary>
     /// Obtient la formule Excel configurée.
     /// </summary>
     public string Formula { get; } = formula;
}

/// <summary>
/// Spécifie la stratégie de conversion appliquée aux propriétés de type énumération.
/// </summary>
/// <param name="mode">Le mode à employer (<see cref="XlsxEnumMode.Name"/>, <see cref="XlsxEnumMode.Value"/> ou <see cref="XlsxEnumMode.Description"/>).</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxEnumAttribute(XlsxEnumMode mode = XlsxEnumMode.Name) : Attribute
{
     /// <summary>
     /// Obtient le mode de sérialisation d'énumération configuré.
     /// </summary>
     public XlsxEnumMode Mode { get; } = mode;
}

/// <summary>
/// Définit le libellé textuel associé à un membre d'énumération lors de l'utilisation de <see cref="XlsxEnumMode.Description"/>.
/// </summary>
/// <param name="label">Le texte exact écrit ou lu dans le classeur Excel.</param>
[AttributeUsage(AttributeTargets.Field)]
public sealed class XlsxEnumMemberAttribute(string label) : Attribute
{
     /// <summary>
     /// Obtient le libellé personnalisé du membre d'énumération.
     /// </summary>
     public string Label { get; } = label;
}

/// <summary>
/// Définit le séparateur décimal attendu ou écrit pour cette colonne (ex: ',' ou '.').
/// </summary>
/// <param name="separator">Le caractère séparateur décimal.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxDecimalSeparatorAttribute(char separator) : Attribute
{
     /// <summary>
     /// Obtient le caractère séparateur décimal configuré.
     /// </summary>
     public char Separator { get; } = separator;
}

/// <summary>
/// Type de générateur ou de source pour la valeur par défaut d'une colonne.
/// </summary>
public enum XlsxDefaultKind
{
     /// <summary>
     /// Valeur littérale statique fournie telle quelle.
     /// </summary>
     Literal = 0,

     /// <summary>
     /// Horodatage local courant (<see cref="DateTime.Now"/>).
     /// </summary>
     Now = 1,

     /// <summary>
     /// Horodatage UTC courant (<see cref="DateTime.UtcNow"/>).
     /// </summary>
     UtcNow = 2,

     /// <summary>
     /// Date du jour courante (<see cref="DateTime.Today"/>).
     /// </summary>
     Today = 3,

     /// <summary>
     /// Nouveau GUID aléatoire (<see cref="Guid.NewGuid"/>).
     /// </summary>
     NewGuid = 4,

     /// <summary>
     /// Chaîne de caractères vide (<see cref="string.Empty"/>).
     /// </summary>
     EmptyString = 5
}

/// <summary>
/// Définit une valeur par défaut statique ou dynamique (Now, UtcNow, Today, NewGuid)
/// lorsque la cellule Excel est vide à la lecture (et optionnellement à l'écriture).
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxDefaultAttribute : Attribute
{
     /// <summary>
     /// Obtient la valeur littérale par défaut lorsque le type est <see cref="XlsxDefaultKind.Literal"/>.
     /// </summary>
     public object? Value { get; }

     /// <summary>
     /// Obtient la nature dynamique ou statique de la valeur par défaut.
     /// </summary>
     public XlsxDefaultKind Kind { get; }

     /// <summary>
     /// Obtient ou définit si la valeur par défaut doit également être appliquée à l'export si la propriété est nulle.
     /// </summary>
     public bool ApplyOnExport { get; set; } = false;

     /// <summary>
     /// Initialise une valeur par défaut statique.
     /// </summary>
     /// <param name="value">La valeur statique à attribuer.</param>
     public XlsxDefaultAttribute(object? value)
     {
          Value = value;
          Kind = XlsxDefaultKind.Literal;
     }

     /// <summary>
     /// Initialise une valeur par défaut dynamique via un mot-clé du moteur.
     /// </summary>
     /// <param name="kind">Le générateur dynamique à employer.</param>
     public XlsxDefaultAttribute(XlsxDefaultKind kind)
     {
          Kind = kind;
     }
}

/// <summary>
/// Spécifie un format exact de date, d'heure ou de nombre (ex: "dd/MM/yyyy" ou "0.00").
/// </summary>
/// <param name="format">Le motif de formatage .NET standard ou personnalisé.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxFormatAttribute(string format) : Attribute
{
     /// <summary>
     /// Obtient le format d'analyse ou d'affichage.
     /// </summary>
     public string Format => format;
}

/// <summary>
/// Lie directement la propriété à une lettre de colonne Excel (ex: "A", "D", "AA"),
/// utile lorsque le fichier n'a pas de ligne d'en-têtes (HasHeaderRow = false).
/// </summary>
/// <param name="columnLetter">La lettre désignant la colonne dans la feuille Excel.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxColumnLetterAttribute(string columnLetter) : Attribute
{
     /// <summary>
     /// Obtient la lettre de colonne attribuée à cette propriété.
     /// </summary>
     public string ColumnLetter => columnLetter;
}

/// <summary>
/// Permet de lire/écrire une collection (string[] ou List&lt;string&gt;) dans une seule cellule séparée par un délimiteur.
/// </summary>
/// <param name="separator">Le caractère délimiteur (par défaut : point-virgule ';').</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxSplitAttribute(char separator = ';') : Attribute
{
     /// <summary>
     /// Obtient le caractère délimiteur utilisé pour séparer les éléments de la collection.
     /// </summary>
     public char Separator => separator;

     /// <summary>
     /// Obtient ou définit si les espaces de chaque entrée doivent être rognés (vaut <see langword="true"/> par défaut).
     /// </summary>
     public bool TrimEntries { get; set; } = true;

     /// <summary>
     /// Obtient ou définit si les segments vides doivent être omis de la collection (vaut <see langword="true"/> par défaut).
     /// </summary>
     public bool RemoveEmptyEntries { get; set; } = true;
}

/// <summary>
/// Associe un convertisseur personnalisé (IXlsxConverter&lt;TProp&gt;, IXlsxRowConverter&lt;TRow, TProp&gt;
/// ou IXlsxRowConverter&lt;TRow, TContext, TProp&gt;) à cette propriété.
/// </summary>
/// <param name="converterType">Le type CLR implémentant l'un des contrats de conversion FlashXlsx.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxConverterAttribute(Type converterType) : Attribute
{
     /// <summary>
     /// Obtient le type du convertisseur personnalisé associé.
     /// </summary>
     public Type ConverterType => converterType;
}

/// <summary>
/// Définit un format numérique natif Excel (ex: "#,##0.00 €", "0.00%", "00000") dans xl/styles.xml.
/// La cellule reste un vrai nombre calculable dans Excel.
/// </summary>
/// <param name="formatCode">Le masque de formattage OpenXML/Excel natif.</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class XlsxNumberFormatAttribute(string formatCode) : Attribute
{
     /// <summary>
     /// Obtient le masque de formatage numérique natif Excel.
     /// </summary>
     public string FormatCode => formatCode;
}