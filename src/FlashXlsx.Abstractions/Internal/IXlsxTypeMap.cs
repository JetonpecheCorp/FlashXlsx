using System;
using System.ComponentModel;

namespace FlashXlsx.Abstractions.Internal;

/// <summary>
/// Contrat généré par Roslyn pour sérialiser/désérialiser un type T sans réflexion.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IXlsxTypeMap<T>
{
     /// <summary>
     /// Nombre total de colonnes mappables (hors [XlsxIgnore]).
     /// </summary>
     int ColumnCount { get; }

     /// <summary>
     /// Retourne les valeurs autorisées pour une liste déroulante Excel ([XlsxAllowedValues] ou membres d'un enum), ou null si aucun.
     /// </summary>
     string[]? GetDropdownValues(int propertyIndex) => null;

     /// <summary>
     /// Retourne l'index de colonne Excel 0-based forcé par [XlsxColumnLetter], ou -1 si non spécifié.
     /// </summary>
     int GetFixedColumnIndex(int propertyIndex) => -1;

     /// <summary>
     /// Retourne le nom C# de la propriété à l'index interne donné (0..ColumnCount-1).
     /// </summary>
     string GetPropertyName(int propertyIndex);

     /// <summary>
     /// Retourne le nom d'en-tête par défaut ([XlsxName] ou nom C#) à l'index interne donné.
     /// </summary>
     string GetDefaultHeaderName(int propertyIndex);

     /// <summary>
     /// Résout un nom de colonne Excel (ou nom C#) vers l'index interne de propriété (-1 si inconnu).
     /// Généré sous forme de switch statique ultra-rapide.
     /// </summary>
     int ResolvePropertyIndex(ReadOnlySpan<char> headerOrPropertyName);

     /// <summary>
     /// Instancie un nouvel objet T (supporte class, record, struct, constructeurs paramétrés).
     /// </summary>
     T CreateInstance();

     /// <summary>
     /// Passe 1 (Lecture) : Parse une cellule brute et affecte la propriété correspondante.
     /// </summary>
     XlsxCellStatus TrySetCellValue(
         ref T instance,
         int propertyIndex,
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         XlsxReadOptions options);

     /// <summary>
     /// Passe 2 (Lecture) : Valide les règles [XlsxRequired], [XlsxRequiredIf] et applique [XlsxIgnoreIf].
     /// </summary>
     bool ValidateRow(
         ref T instance,
         int rowNumber,
         ReadOnlySpan<int> propertyToExcelColIndex,
         ReadOnlySpan<string?> excelHeaders,
         XlsxReadOptions options);

     /// <summary>
     /// Écriture : Extrait la valeur formatée d'une propriété pour l'écrire dans le flux XML.
     /// </summary>
     bool TryFormatCellValue(
         in T instance,
         int propertyIndex,
         int rowNumber,
         Span<char> destination,
         out int charsWritten,
         out string? stringValue,
         out XlsxCellWriteKind writeKind,
         XlsxWriteOptions options);
}

/// <summary>
/// Contrat généré par Roslyn pour les modèles nécessitant un contexte externe TContext.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IXlsxTypeMap<T, TContext> : IXlsxTypeMap<T>
{
     bool ValidateRow(
         ref T instance,
         in TContext context,
         int rowNumber,
         ReadOnlySpan<int> propertyToExcelColIndex,
         ReadOnlySpan<string?> excelHeaders,
         XlsxReadOptions<T, TContext> options);

     bool TryFormatCellValue(
         in T instance,
         in TContext context,
         int propertyIndex,
         int rowNumber,
         Span<char> destination,
         out int charsWritten,
         out string? stringValue,
         out XlsxCellWriteKind writeKind,
         XlsxWriteOptions<T, TContext> options);
}

[EditorBrowsable(EditorBrowsableState.Never)]
public enum XlsxCellWriteKind : byte
{
     Empty = 0,
     Number,
     Date,
     DateTime,
     Time,
     Boolean,
     InlineString,
     Formula
}
