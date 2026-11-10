namespace FlashXlsx.Abstractions;

/// <summary>
/// Comportement du moteur lorsqu'une erreur de donnée est rencontrée.
/// </summary>
public enum XlsxErrorMode : byte
{
     /// <summary>
     /// Notifie l'erreur via OnError et interrompt immédiatement le traitement.
     /// </summary>
     StopOnFirstError,

     /// <summary>
     /// Notifie l'erreur via OnError, laisse la valeur par défaut et poursuit le traitement.
     /// </summary>
     ReportAndContinue
}

/// <summary>
/// État précis d'une cellule ou d'une structure lors du traitement.
/// </summary>
public enum XlsxCellStatus : byte
{
     /// <summary>
     /// La cellule a été lue ou écrite avec succès sans anomalie.
     /// </summary>
     Success,

     /// <summary>
     /// La cellule est vide.
     /// </summary>
     Empty,

     /// <summary>
     /// La cellule est vide ou absente alors qu'elle est marquée comme obligatoire (<see cref="XlsxRequiredAttribute"/> ou <see cref="XlsxRequiredIfAttribute"/>).
     /// </summary>
     RequiredMissing,

     /// <summary>
     /// La chaîne brute n'a pas pu être convertie dans le type cible attendu.
     /// </summary>
     InvalidFormat,

     /// <summary>
     /// La valeur numérique dépasse les limites autorisées par le type cible (dépassement de capacité).
     /// </summary>
     Overflow,

     /// <summary>
     /// La colonne mappée est introuvable dans le fichier ou l'en-tête.
     /// </summary>
     ColumnNotFound,

     /// <summary>
     /// L'onglet demandé par son nom ou son index n'existe pas dans le classeur.
     /// </summary>
     SheetNotFound,

     /// <summary>
     /// La valeur numérique est en dehors des bornes définies (<see cref="XlsxRangeAttribute"/>, <see cref="XlsxMinAttribute"/>, <see cref="XlsxMaxAttribute"/>).
     /// </summary>
     OutOfRange,

     /// <summary>
     /// La longueur de la chaîne est invalide par rapport aux contraintes (<see cref="XlsxLengthAttribute"/>).
     /// </summary>
     LengthInvalid,

     /// <summary>
     /// La valeur ne respecte pas l'expression régulière spécifiée (<see cref="XlsxRegexAttribute"/>).
     /// </summary>
     PatternMismatch,

     /// <summary>
     /// La valeur ne fait pas partie de la liste autorisée (<see cref="XlsxAllowedValuesAttribute"/>).
     /// </summary>
     ValueNotAllowed,

     /// <summary>
     /// Tentative d'injection de formule ou de payload potentiellement malveillant détectée et bloquée.
     /// </summary>
     SecurityViolation,

     /// <summary>
     /// Une exception est survenue lors de l'exécution d'un convertisseur personnalisé (<see cref="XlsxConverterAttribute"/>).
     /// </summary>
     ConverterError,

     /// <summary>
     /// Une erreur d'E/S ou de flux est survenue lors de l'écriture.
     /// </summary>
     WriteError,

     /// <summary>
     /// Le fichier XLSX est corrompu ou ne respecte pas le format OpenXML attendu.
     /// </summary>
     CorruptedFile,

     /// <summary>
     /// Le document est protégé par mot de passe ou encapsulé dans un conteneur binaire chiffré (OLE2/CFB).
     /// </summary>
     EncryptedFile
}

/// <summary>
/// Contexte léger alloué sur la pile (stack) décrivant une erreur sans lever d'exception.
/// </summary>
/// <param name="RowNumber">Le numéro de ligne dans la feuille Excel (base 1).</param>
/// <param name="ColumnIndex">L'index basé sur zéro de la colonne concernée.</param>
/// <param name="ColumnLetter">La lettre de la colonne Excel (ex: "A", "BC").</param>
/// <param name="ColumnName">Le libellé de l'en-tête de colonne s'il existe.</param>
/// <param name="RawValue">La valeur textuelle brute extraite de la cellule.</param>
/// <param name="Status">Le diagnostic précis de l'erreur (<see cref="XlsxCellStatus"/>).</param>
/// <param name="SheetName">Le nom de l'onglet concerné si disponible.</param>
public readonly record struct XlsxErrorContext(
    int RowNumber,
    int ColumnIndex,
    string ColumnLetter,
    string? ColumnName,
    string RawValue,
    XlsxCellStatus Status,
    string? SheetName = null
);

/// <summary>
/// Délégué appelé lors d'une erreur (passage par référence constante 'in' pour éviter toute copie).
/// </summary>
/// <param name="error">Le contexte détaillé de l'erreur rencontrée.</param>
public delegate void XlsxErrorHandler(in XlsxErrorContext error);