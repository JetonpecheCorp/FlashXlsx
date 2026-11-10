using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Identifie le type de donnée primitif stocké dans une valeur dynamique <see cref="XlsxValue"/>.
/// </summary>
public enum XlsxDataType : byte
{
     /// <summary>Cellule vide ou valeur non initialisée.</summary>
     Empty = 0,

     /// <summary>Chaîne de caractères (texte).</summary>
     String,

     /// <summary>Entier signé sur 64 bits (<see cref="long"/>) ou 32 bits (<see cref="int"/>).</summary>
     Int64,

     /// <summary>Nombre à virgule flottante double précision (<see cref="double"/>).</summary>
     Double,

     /// <summary>Nombre décimal haute précision (<see cref="decimal"/>).</summary>
     Decimal,

     /// <summary>Valeur booléenne (<see cref="bool"/>).</summary>
     Boolean,

     /// <summary>Horodatage complet avec date et heure (<see cref="System.DateTime"/>).</summary>
     DateTime,

     /// <summary>Date calendaire sans composante d'heure (<see cref="System.DateOnly"/>).</summary>
     DateOnly,

     /// <summary>Heure de la journée sans composante de date (<see cref="System.TimeOnly"/>).</summary>
     TimeOnly,

     /// <summary>Identifiant global unique (<see cref="System.Guid"/>).</summary>
     Guid
}

/// <summary>
/// Union discriminée sans allocation (zéro boxing) représentant une cellule Excel typée en mode dynamique par dictionnaire.
/// Stocke les primitives directement dans des champs de bits ou structs contigus pour préserver le Garbage Collector.
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly struct XlsxValue : IEquatable<XlsxValue>
{
     private readonly long _bits;
     private readonly decimal _decimal;
     private readonly Guid _guid;
     private readonly string? _string;

     /// <summary>
     /// Obtient le type primitif actuellement stocké dans cette valeur.
     /// </summary>
     public XlsxDataType Type { get; }

     /// <summary>
     /// Indique si la cellule est vide ou correspond à la valeur par défaut.
     /// </summary>
     public bool IsEmpty => Type == XlsxDataType.Empty;

     /// <summary>
     /// Obtient une instance représentant une valeur Excel vide.
     /// </summary>
     public static XlsxValue Empty => default;

     private XlsxValue(XlsxDataType type, long bits = 0, decimal dec = 0m, Guid guid = default, string? str = null)
     {
          Type = type;
          _bits = bits;
          _decimal = dec;
          _guid = guid;
          _string = str;
     }

     // --- Conversions implicites (Zéro boxing) ---

     /// <summary>Convertit implicitement une chaîne de caractères en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">La chaîne source, ou <see langword="null"/> pour une valeur vide.</param>
     public static implicit operator XlsxValue(string? value)
         => value is null ? Empty : new(XlsxDataType.String, str: value);

     /// <summary>Convertit implicitement un entier 32 bits en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">L'entier source.</param>
     public static implicit operator XlsxValue(int value)
         => new(XlsxDataType.Int64, bits: value);

     /// <summary>Convertit implicitement un entier 64 bits en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">L'entier source.</param>
     public static implicit operator XlsxValue(long value)
         => new(XlsxDataType.Int64, bits: value);

     /// <summary>Convertit implicitement un double en <see cref="XlsxValue"/> sans allocation.</summary>
     /// <param name="value">Le double source.</param>
     public static implicit operator XlsxValue(double value)
         => new(XlsxDataType.Double, bits: BitConverter.DoubleToInt64Bits(value));

     /// <summary>Convertit implicitement un décimal en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">Le décimal source.</param>
     public static implicit operator XlsxValue(decimal value)
         => new(XlsxDataType.Decimal, dec: value);

     /// <summary>Convertit implicitement un booléen en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">Le booléen source.</param>
     public static implicit operator XlsxValue(bool value)
         => new(XlsxDataType.Boolean, bits: value ? 1L : 0L);

     /// <summary>Convertit implicitement un <see cref="DateTime"/> en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">L'horodatage source.</param>
     public static implicit operator XlsxValue(DateTime value)
         => new(XlsxDataType.DateTime, bits: value.ToBinary());

     /// <summary>Convertit implicitement un <see cref="DateOnly"/> en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">La date source.</param>
     public static implicit operator XlsxValue(DateOnly value)
         => new(XlsxDataType.DateOnly, bits: value.DayNumber);

     /// <summary>Convertit implicitement un <see cref="TimeOnly"/> en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">L'heure source.</param>
     public static implicit operator XlsxValue(TimeOnly value)
         => new(XlsxDataType.TimeOnly, bits: value.Ticks);

     /// <summary>Convertit implicitement un <see cref="Guid"/> en <see cref="XlsxValue"/>.</summary>
     /// <param name="value">L'identifiant unique source.</param>
     public static implicit operator XlsxValue(Guid value)
         => new(XlsxDataType.Guid, guid: value);

     // --- Accesseurs typés ---

     /// <summary>
     /// Restitue la valeur sous forme de texte formaté, ou <see langword="null"/> si la cellule est vide.
     /// </summary>
     /// <returns>La chaîne formatée correspondante ou <see langword="null"/>.</returns>
     public string? AsString() => Type switch
     {
          XlsxDataType.Empty => null,
          XlsxDataType.String => _string,
          XlsxDataType.Int64 => _bits.ToString(CultureInfo.InvariantCulture),
          XlsxDataType.Double => BitConverter.Int64BitsToDouble(_bits).ToString(CultureInfo.InvariantCulture),
          XlsxDataType.Decimal => _decimal.ToString(CultureInfo.InvariantCulture),
          XlsxDataType.Boolean => _bits != 0 ? "true" : "false",
          XlsxDataType.DateTime => DateTime.FromBinary(_bits).ToString("O"),
          XlsxDataType.DateOnly => DateOnly.FromDayNumber((int)_bits).ToString("O"),
          XlsxDataType.TimeOnly => new TimeOnly(_bits).ToString("O"),
          XlsxDataType.Guid => _guid.ToString(),
          _ => null
     };

     /// <summary>
     /// Extrait la valeur sous forme d'entier 64 bits (<see cref="long"/>). Convertit les nombres à virgule par troncature.
     /// </summary>
     /// <returns>La valeur sous forme de long, ou 0 si incompatible.</returns>
     public long AsInt64() => Type switch
     {
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Double => (long)BitConverter.Int64BitsToDouble(_bits),
          XlsxDataType.Decimal => (long)_decimal,
          XlsxDataType.Boolean => _bits,
          _ => 0L
     };

     /// <summary>
     /// Extrait la valeur sous forme d'entier 32 bits (<see cref="int"/>).
     /// </summary>
     /// <returns>La valeur sous forme d'entier, ou 0 si incompatible.</returns>
     public int AsInt32() => (int)AsInt64();

     /// <summary>
     /// Extrait la valeur sous forme de nombre à virgule flottante (<see cref="double"/>).
     /// </summary>
     /// <returns>La valeur sous forme de double, ou 0.0 si incompatible.</returns>
     public double AsDouble() => Type switch
     {
          XlsxDataType.Double => BitConverter.Int64BitsToDouble(_bits),
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Decimal => (double)_decimal,
          _ => 0d
     };

     /// <summary>
     /// Extrait la valeur sous forme de décimal haute précision (<see cref="decimal"/>).
     /// </summary>
     /// <returns>La valeur sous forme de décimal, ou 0m si incompatible.</returns>
     public decimal AsDecimal() => Type switch
     {
          XlsxDataType.Decimal => _decimal,
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Double => (decimal)BitConverter.Int64BitsToDouble(_bits),
          _ => 0m
     };

     /// <summary>
     /// Extrait la valeur sous forme de booléen.
     /// </summary>
     /// <returns><see langword="true"/> si la valeur est évaluée à vrai ; sinon, <see langword="false"/>.</returns>
     public bool AsBoolean() => Type switch
     {
          XlsxDataType.Boolean or XlsxDataType.Int64 => _bits != 0,
          _ => false
     };

     /// <summary>
     /// Extrait la valeur sous forme d'horodatage <see cref="DateTime"/>.
     /// </summary>
     /// <returns>L'horodatage correspondant, ou <see cref="DateTime.MinValue"/> si incompatible.</returns>
     public DateTime AsDateTime() => Type == XlsxDataType.DateTime
         ? DateTime.FromBinary(_bits)
         : default;

     /// <summary>
     /// Extrait la valeur sous forme de date pure <see cref="DateOnly"/>.
     /// </summary>
     /// <returns>L'instance <see cref="DateOnly"/> correspondante, ou la valeur par défaut.</returns>
     public DateOnly AsDateOnly() => Type switch
     {
          XlsxDataType.DateOnly => DateOnly.FromDayNumber((int)_bits),
          XlsxDataType.DateTime => DateOnly.FromDateTime(DateTime.FromBinary(_bits)),
          _ => default
     };

     /// <summary>
     /// Extrait la valeur sous forme d'heure pure <see cref="TimeOnly"/>.
     /// </summary>
     /// <returns>L'instance <see cref="TimeOnly"/> correspondante, ou la valeur par défaut.</returns>
     public TimeOnly AsTimeOnly() => Type switch
     {
          XlsxDataType.TimeOnly => new TimeOnly(_bits),
          XlsxDataType.DateTime => TimeOnly.FromDateTime(DateTime.FromBinary(_bits)),
          _ => default
     };

     /// <summary>
     /// Extrait la valeur sous forme d'identifiant unique <see cref="Guid"/>.
     /// </summary>
     /// <returns>Le GUID extrait ou <see cref="Guid.Empty"/> si incompatible.</returns>
     public Guid AsGuid() => Type == XlsxDataType.Guid ? _guid : Guid.Empty;

     /// <inheritdoc/>
     public override string ToString() => AsString() ?? string.Empty;

     /// <inheritdoc/>
     public bool Equals(XlsxValue other)
         => Type == other.Type
         && _bits == other._bits
         && _decimal == other._decimal
         && _guid == other._guid
         && _string == other._string;

     /// <inheritdoc/>
     public override bool Equals(object? obj) => obj is XlsxValue other && Equals(other);

     /// <inheritdoc/>
     public override int GetHashCode() => HashCode.Combine(Type, _bits, _decimal, _guid, _string);
}