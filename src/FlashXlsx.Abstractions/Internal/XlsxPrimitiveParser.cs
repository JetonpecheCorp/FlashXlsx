using System;
using System.Buffers;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Text;

namespace FlashXlsx.Abstractions.Internal;

/// <summary>
/// Parseurs primitifs haute performance basés sur ReadOnlySpan&lt;char&gt; (zéro allocation).
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class XlsxPrimitiveParser
{
     public static XlsxCellStatus TryParseInt32(ReadOnlySpan<char> span, CultureInfo culture, out int value)
     {
          value = 0;
          span = StripCurrencyAndSpaces(span);
          if (span.IsEmpty) return XlsxCellStatus.Empty;
          if (int.TryParse(span, NumberStyles.Integer, culture, out value)) return XlsxCellStatus.Success;
          return long.TryParse(span, NumberStyles.Integer, culture, out _) ? XlsxCellStatus.Overflow : XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseInt64(ReadOnlySpan<char> span, CultureInfo culture, out long value)
     {
          value = 0;
          span = StripCurrencyAndSpaces(span);
          if (span.IsEmpty) return XlsxCellStatus.Empty;
          return long.TryParse(span, NumberStyles.Integer, culture, out value) ? XlsxCellStatus.Success : XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseDecimal(ReadOnlySpan<char> span, CultureInfo culture, out decimal value)
     {
          value = 0m;
          span = StripCurrencyAndSpaces(span);
          if (span.IsEmpty) return XlsxCellStatus.Empty;
          return decimal.TryParse(span, NumberStyles.Number, culture, out value) ? XlsxCellStatus.Success : XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseDouble(ReadOnlySpan<char> span, CultureInfo culture, out double value)
     {
          value = 0d;
          span = StripCurrencyAndSpaces(span);
          if (span.IsEmpty) return XlsxCellStatus.Empty;
          return double.TryParse(span, NumberStyles.Float | NumberStyles.AllowThousands, culture, out value) ? XlsxCellStatus.Success : XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseDateTime(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         CultureInfo culture,
         out DateTime result,
         bool use1904DateSystem = false)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (isNumericCell && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate))
          {
               double adjusted = use1904DateSystem ? oaDate + 1462.0 : oaDate;
               if (adjusted is <= -657435.0 or >= 2958466.0)
                    return XlsxCellStatus.InvalidFormat;

               result = DateTime.FromOADate(adjusted);
               return XlsxCellStatus.Success;
          }

          if (DateTime.TryParse(rawValue, culture, DateTimeStyles.None, out result) ||
              DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
          {
               return XlsxCellStatus.Success;
          }

          return XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseDateOnly(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         CultureInfo culture,
         out DateOnly result,
         bool use1904DateSystem = false)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (isNumericCell && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate))
          {
               double adjusted = use1904DateSystem ? oaDate + 1462.0 : oaDate;
               if (adjusted is <= -657435.0 or >= 2958466.0)
                    return XlsxCellStatus.InvalidFormat;

               DateTime dt = DateTime.FromOADate(adjusted);
               result = DateOnly.FromDateTime(dt);
               return XlsxCellStatus.Success;
          }

          if (DateOnly.TryParse(rawValue, culture, DateTimeStyles.None, out result) ||
              DateOnly.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
          {
               return XlsxCellStatus.Success;
          }

          if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDt))
          {
               result = DateOnly.FromDateTime(parsedDt);
               return XlsxCellStatus.Success;
          }

          return XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseGuid(ReadOnlySpan<char> span, out Guid value)
     {
          value = Guid.Empty;
          if (span.IsEmpty) return XlsxCellStatus.Empty;
          return Guid.TryParse(span, out value) ? XlsxCellStatus.Success : XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseBoolean(
         ReadOnlySpan<char> span,
         string[]? trueValues,
         string[]? falseValues,
         out bool value)
     {
          value = false;
          bool hasTrue = trueValues is { Length: > 0 };
          bool hasFalse = falseValues is { Length: > 0 };

          // Standard Excel par défaut si aucun tableau n'est configuré
          if (!hasTrue && !hasFalse)
          {
               if (span.IsEmpty) return XlsxCellStatus.Empty;
               if (span.Equals("1", StringComparison.Ordinal) || span.Equals("true", StringComparison.OrdinalIgnoreCase))
               {
                    value = true;
                    return XlsxCellStatus.Success;
               }
               if (span.Equals("0", StringComparison.Ordinal) || span.Equals("false", StringComparison.OrdinalIgnoreCase))
               {
                    value = false;
                    return XlsxCellStatus.Success;
               }
               return XlsxCellStatus.InvalidFormat;
          }

          if (hasTrue)
          {
               foreach (var t in trueValues!)
               {
                    if (span.Equals(t.AsSpan(), StringComparison.OrdinalIgnoreCase))
                    {
                         value = true;
                         return XlsxCellStatus.Success;
                    }
               }
          }

          if (hasFalse)
          {
               foreach (var f in falseValues!)
               {
                    if (span.Equals(f.AsSpan(), StringComparison.OrdinalIgnoreCase))
                    {
                         value = false;
                         return XlsxCellStatus.Success;
                    }
               }
          }

          // Si un seul tableau est rempli, toute autre valeur prend automatiquement la valeur du tableau vide
          if (hasTrue && !hasFalse)
          {
               value = false;
               return XlsxCellStatus.Success;
          }

          if (!hasTrue && hasFalse)
          {
               value = true;
               return XlsxCellStatus.Success;
          }

          return XlsxCellStatus.InvalidFormat;
     }

     /// <summary>
     /// Parse un decimal avec gestion flexible du séparateur ('.', ',', ou auto-détection '\0') et nettoyage automatique des devises.
     /// </summary>
     public static XlsxCellStatus TryParseDecimalFlexible(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         char preferredSeparator, // '\0' = auto-détection flexible, sinon ',' ou '.'
         CultureInfo culture,
         out decimal result)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          // Dans un fichier XLSX, une cellule numérique native utilise toujours '.' en XML
          if (isNumericCell && decimal.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
               return XlsxCellStatus.Success;

          ReadOnlySpan<char> cleaned = StripCurrencyAndSpaces(rawValue);
          if (cleaned.IsEmpty)
               return XlsxCellStatus.Empty;

          if (NormalizeDecimalSpan(cleaned, preferredSeparator, stackalloc char[64], out int written, out ReadOnlySpan<char> normalized))
          {
               if (decimal.TryParse(normalized.Slice(0, written), NumberStyles.Float, CultureInfo.InvariantCulture, out result))
                    return XlsxCellStatus.Success;
          }

          if (decimal.TryParse(cleaned, NumberStyles.Number, culture, out result))
               return XlsxCellStatus.Success;

          return XlsxCellStatus.InvalidFormat;
     }

     /// <summary>
     /// Parse un double avec gestion flexible du séparateur ('.', ',', ou auto-détection '\0') et nettoyage automatique des devises.
     /// </summary>
     public static XlsxCellStatus TryParseDoubleFlexible(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         char preferredSeparator,
         CultureInfo culture,
         out double result)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (isNumericCell && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
               return XlsxCellStatus.Success;

          ReadOnlySpan<char> cleaned = StripCurrencyAndSpaces(rawValue);
          if (cleaned.IsEmpty)
               return XlsxCellStatus.Empty;

          if (NormalizeDecimalSpan(cleaned, preferredSeparator, stackalloc char[64], out int written, out ReadOnlySpan<char> normalized))
          {
               if (double.TryParse(normalized.Slice(0, written), NumberStyles.Float, CultureInfo.InvariantCulture, out result))
                    return XlsxCellStatus.Success;
          }

          if (double.TryParse(cleaned, NumberStyles.Number, culture, out result))
               return XlsxCellStatus.Success;

          return XlsxCellStatus.InvalidFormat;
     }

     /// <summary>
     /// Retire sans allocation les symboles monétaires (€, $, £, ¥, ₹, CHF, EUR, USD, GBP, CAD) et espaces en début/fin.
     /// </summary>
     private static ReadOnlySpan<char> StripCurrencyAndSpaces(ReadOnlySpan<char> span)
     {
          span = TrimAllSpaces(span);
          if (span.IsEmpty)
               return span;

          // Symboles ou codes monétaires en préfixe (ex: "$1,250.50", "CHF 120.00", "EUR 99,90")
          span = TrimCurrencyPrefix(span);
          span = TrimAllSpaces(span);

          // Symboles ou codes monétaires en suffixe (ex: "1 250,50 €", "120.00 CHF", "99,90 EUR")
          span = TrimCurrencySuffix(span);
          span = TrimAllSpaces(span);

          return span;
     }

     private static ReadOnlySpan<char> TrimAllSpaces(ReadOnlySpan<char> span)
     {
          int start = 0;
          int end = span.Length - 1;
          while (start <= end && IsSpaceChar(span[start])) start++;
          while (end >= start && IsSpaceChar(span[end])) end--;
          return start > end ? ReadOnlySpan<char>.Empty : span.Slice(start, end - start + 1);
     }

     private static bool IsSpaceChar(char c) => c == ' ' || c == '\t' || c == '\u00A0' || c == '\u202F';

     private static bool IsCurrencySymbol(char c) => c is '€' or '$' or '£' or '¥' or '₹' or '₩' or '₽';

     private static ReadOnlySpan<char> TrimCurrencyPrefix(ReadOnlySpan<char> span)
     {
          if (span.IsEmpty) return span;
          if (IsCurrencySymbol(span[0]))
               return span[1..];

          if (span.Length >= 3)
          {
               ReadOnlySpan<char> prefix3 = span[..3];
               if (prefix3.Equals("CHF", StringComparison.OrdinalIgnoreCase) ||
                   prefix3.Equals("EUR", StringComparison.OrdinalIgnoreCase) ||
                   prefix3.Equals("USD", StringComparison.OrdinalIgnoreCase) ||
                   prefix3.Equals("GBP", StringComparison.OrdinalIgnoreCase) ||
                   prefix3.Equals("CAD", StringComparison.OrdinalIgnoreCase))
               {
                    return span[3..];
               }
          }
          return span;
     }

     private static ReadOnlySpan<char> TrimCurrencySuffix(ReadOnlySpan<char> span)
     {
          if (span.IsEmpty) return span;
          if (IsCurrencySymbol(span[^1]))
               return span[..^1];

          if (span.Length >= 3)
          {
               ReadOnlySpan<char> suffix3 = span[^3..];
               if (suffix3.Equals("CHF", StringComparison.OrdinalIgnoreCase) ||
                   suffix3.Equals("EUR", StringComparison.OrdinalIgnoreCase) ||
                   suffix3.Equals("USD", StringComparison.OrdinalIgnoreCase) ||
                   suffix3.Equals("GBP", StringComparison.OrdinalIgnoreCase) ||
                   suffix3.Equals("CAD", StringComparison.OrdinalIgnoreCase))
               {
                    return span[..^3];
               }
          }
          return span;
     }

     private static bool NormalizeDecimalSpan(
         ReadOnlySpan<char> input,
         char preferredSeparator,
         Span<char> stackBuffer,
         out int written,
         out ReadOnlySpan<char> resultSpan)
     {
          written = 0;
          resultSpan = default;

          // Détermination du séparateur décimal effectif si mode auto ('\0')
          char decSep = preferredSeparator;
          if (decSep == '\0')
          {
               int lastDot = input.LastIndexOf('.');
               int lastComma = input.LastIndexOf(',');
               if (lastDot >= 0 && lastComma >= 0)
               {
                    // Format avec séparateur de milliers + décimales (ex: "1,234.56" -> '.' ou "1.234,56" -> ',')
                    decSep = lastComma > lastDot ? ',' : '.';
               }
               else if (lastComma >= 0)
               {
                    decSep = ',';
               }
               else
               {
                    decSep = '.';
               }
          }

          char thousandsSep = decSep == ',' ? '.' : ',';
          char[]? rented = input.Length > stackBuffer.Length ? ArrayPool<char>.Shared.Rent(input.Length) : null;
          Span<char> dest = rented is not null ? rented : stackBuffer;

          try
          {
               for (int i = 0; i < input.Length; i++)
               {
                    char c = input[i];
                    // Ignore les espaces classiques, insécables (\u00A0, \u202F) et séparateurs de milliers
                    if (IsSpaceChar(c) || c == thousandsSep)
                         continue;

                    dest[written++] = (c == decSep) ? '.' : c;
               }

               if (rented is not null)
               {
                    resultSpan = dest.Slice(0, written).ToArray();
               }
               else
               {
                    resultSpan = stackBuffer;
               }
               return written > 0;
          }
          finally
          {
               if (rented is not null)
                    ArrayPool<char>.Shared.Return(rented);
          }
     }

     public static XlsxCellStatus TryParseTimeOnly(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         CultureInfo culture,
         out TimeOnly result)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (isNumericCell && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate))
          {
               // Extraction de la partie fractionnaire du jour (gère aussi un DateTime complet)
               double fractionalDay = oaDate - Math.Floor(oaDate);
               long totalMilliseconds = (long)Math.Round(fractionalDay * 86_400_000.0);
               if (totalMilliseconds >= 86_400_000L)
                    totalMilliseconds = 0;

               result = TimeOnly.FromTimeSpan(TimeSpan.FromMilliseconds(totalMilliseconds));
               return XlsxCellStatus.Success;
          }

          if (TimeOnly.TryParse(rawValue, culture, DateTimeStyles.None, out result) ||
              TimeOnly.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
          {
               return XlsxCellStatus.Success;
          }

          return XlsxCellStatus.InvalidFormat;
     }

     public static XlsxCellStatus TryParseDateTimeExact(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         string format,
         CultureInfo culture,
         out DateTime result,
         bool use1904DateSystem = false)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (!isNumericCell && DateTime.TryParseExact(rawValue, format, culture, DateTimeStyles.None, out result))
               return XlsxCellStatus.Success;

          return TryParseDateTime(rawValue, isNumericCell, culture, out result, use1904DateSystem);
     }

     public static XlsxCellStatus TryParseDateOnlyExact(
         ReadOnlySpan<char> rawValue,
         bool isNumericCell,
         string format,
         CultureInfo culture,
         out DateOnly result,
         bool use1904DateSystem = false)
     {
          result = default;
          if (rawValue.IsEmpty)
               return XlsxCellStatus.Empty;

          if (!isNumericCell && DateOnly.TryParseExact(rawValue, format, culture, DateTimeStyles.None, out result))
               return XlsxCellStatus.Success;

          return TryParseDateOnly(rawValue, isNumericCell, culture, out result, use1904DateSystem);
     }

     public static string[] SplitStringList(ReadOnlySpan<char> rawValue, char separator, bool trimEntries, bool removeEmpty)
     {
          if (rawValue.IsEmpty)
               return [];

          var list = new List<string>();
          while (true)
          {
               int idx = rawValue.IndexOf(separator);
               ReadOnlySpan<char> segment = idx >= 0 ? rawValue.Slice(0, idx) : rawValue;
               if (trimEntries)
                    segment = segment.Trim();

               if (!removeEmpty || !segment.IsEmpty)
                    list.Add(segment.ToString());

               if (idx < 0)
                    break;

               rawValue = rawValue.Slice(idx + 1);
          }
          return [.. list];
     }
}