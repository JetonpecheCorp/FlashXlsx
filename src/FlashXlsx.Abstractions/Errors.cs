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
     Success,
     Empty,
     RequiredMissing,
     InvalidFormat,
     Overflow,
     ColumnNotFound,
     SheetNotFound,
     OutOfRange,       // [XlsxRange], [XlsxMin], [XlsxMax]
     LengthInvalid,    // [XlsxLength], [XlsxMinLength], [XlsxMaxLength]
     PatternMismatch,  // [XlsxRegex]
     ValueNotAllowed,  // [XlsxAllowedValues]
     SecurityViolation // Injection de formule / payload dangereux bloqué
}

/// <summary>
/// Contexte léger alloué sur la pile (stack) décrivant une erreur sans lever d'exception.
/// </summary>
public readonly record struct XlsxErrorContext(
    int RowNumber,
    int ColumnIndex,
    string ColumnLetter,
    string? ColumnName,
    string RawValue,
    XlsxCellStatus Status
);

/// <summary>
/// Délégué appelé lors d'une erreur (passage par référence constante 'in' pour éviter toute copie).
/// </summary>
public delegate void XlsxErrorHandler(in XlsxErrorContext error);
