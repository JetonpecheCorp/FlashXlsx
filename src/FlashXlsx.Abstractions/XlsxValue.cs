using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace FlashXlsx.Abstractions;

public enum XlsxDataType : byte
{
     Empty = 0,
     String,
     Int64,
     Double,
     Decimal,
     Boolean,
     DateTime,
     DateOnly,
     TimeOnly,
     Guid
}

/// <summary>
/// Union discriminée 100% zéro-boxing pour le mapping dynamique par dictionnaire (.NET 10).
/// </summary>
[StructLayout(LayoutKind.Auto)]
public readonly struct XlsxValue : IEquatable<XlsxValue>
{
     private readonly long _bits;
     private readonly decimal _decimal;
     private readonly Guid _guid;
     private readonly string? _string;

     public XlsxDataType Type { get; }
     public bool IsEmpty => Type == XlsxDataType.Empty;

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
     public static implicit operator XlsxValue(string? value)
         => value is null ? Empty : new(XlsxDataType.String, str: value);

     public static implicit operator XlsxValue(int value)
         => new(XlsxDataType.Int64, bits: value);

     public static implicit operator XlsxValue(long value)
         => new(XlsxDataType.Int64, bits: value);

     public static implicit operator XlsxValue(double value)
         => new(XlsxDataType.Double, bits: BitConverter.DoubleToInt64Bits(value));

     public static implicit operator XlsxValue(decimal value)
         => new(XlsxDataType.Decimal, dec: value);

     public static implicit operator XlsxValue(bool value)
         => new(XlsxDataType.Boolean, bits: value ? 1L : 0L);

     public static implicit operator XlsxValue(DateTime value)
         => new(XlsxDataType.DateTime, bits: value.ToBinary());

     public static implicit operator XlsxValue(DateOnly value)
         => new(XlsxDataType.DateOnly, bits: value.DayNumber);

     public static implicit operator XlsxValue(TimeOnly value)
         => new(XlsxDataType.TimeOnly, bits: value.Ticks);

     public static implicit operator XlsxValue(Guid value)
         => new(XlsxDataType.Guid, guid: value);

     // --- Accesseurs typés ---
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

     public long AsInt64() => Type switch
     {
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Double => (long)BitConverter.Int64BitsToDouble(_bits),
          XlsxDataType.Decimal => (long)_decimal,
          XlsxDataType.Boolean => _bits,
          _ => 0L
     };

     public int AsInt32() => (int)AsInt64();

     public double AsDouble() => Type switch
     {
          XlsxDataType.Double => BitConverter.Int64BitsToDouble(_bits),
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Decimal => (double)_decimal,
          _ => 0d
     };

     public decimal AsDecimal() => Type switch
     {
          XlsxDataType.Decimal => _decimal,
          XlsxDataType.Int64 => _bits,
          XlsxDataType.Double => (decimal)BitConverter.Int64BitsToDouble(_bits),
          _ => 0m
     };

     public bool AsBoolean() => Type switch
     {
          XlsxDataType.Boolean or XlsxDataType.Int64 => _bits != 0,
          _ => false
     };

     public DateTime AsDateTime() => Type == XlsxDataType.DateTime
         ? DateTime.FromBinary(_bits)
         : default;

     public DateOnly AsDateOnly() => Type switch
     {
          XlsxDataType.DateOnly => DateOnly.FromDayNumber((int)_bits),
          XlsxDataType.DateTime => DateOnly.FromDateTime(DateTime.FromBinary(_bits)),
          _ => default
     };

     public TimeOnly AsTimeOnly() => Type switch
     {
          XlsxDataType.TimeOnly => new TimeOnly(_bits),
          XlsxDataType.DateTime => TimeOnly.FromDateTime(DateTime.FromBinary(_bits)),
          _ => default
     };

     public Guid AsGuid() => Type == XlsxDataType.Guid ? _guid : Guid.Empty;

     public override string ToString() => AsString() ?? string.Empty;

     public bool Equals(XlsxValue other)
         => Type == other.Type
         && _bits == other._bits
         && _decimal == other._decimal
         && _guid == other._guid
         && _string == other._string;

     public override bool Equals(object? obj) => obj is XlsxValue other && Equals(other);
     public override int GetHashCode() => HashCode.Combine(Type, _bits, _decimal, _guid, _string);
}