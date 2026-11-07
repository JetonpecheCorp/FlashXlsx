using System;

namespace FlashXlsx.Abstractions;

internal static class PropertyPathHelper
{
     public static string ExtractPropertyName(string? expression)
     {
          if (string.IsNullOrWhiteSpace(expression))
               throw new ArgumentException("L'expression de sélection de propriété ne peut pas être vide.", nameof(expression));

          ReadOnlySpan<char> span = expression.AsSpan().Trim();
          int lastDot = span.LastIndexOf('.');

          if (lastDot < 0 || lastDot == span.Length - 1)
               return span.ToString();

          return span.Slice(lastDot + 1).Trim().ToString();
     }
}
