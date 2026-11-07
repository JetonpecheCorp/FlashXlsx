using System;
using System.Buffers;
using System.Runtime.CompilerServices;

namespace FlashXlsx.Abstractions.Internal;

public enum XlsxSecurityPolicy
{
     /// <summary>Échappe les formules dangereuses à l'export avec une apostrophe (') et nettoie les caractères de contrôle.</summary>
     Sanitize = 0,
     /// <summary>Rejette la cellule avec XlsxCellStatus.SecurityViolation dès détection d'une formule ou d'un payload XSS.</summary>
     Reject = 1,
     /// <summary>Désactive la protection (uniquement pour des données internes de confiance).</summary>
     None = 2
}

public static class XlsxSecuritySanitizer
{
     private static readonly SearchValues<char> DangerousFormulaPrefixes =
         SearchValues.Create(['=', '+', '-', '@', '\t', '\r']);

     /// <summary>
     /// Vérifie si une valeur brute contient une tentative d'injection de formule Excel ou de XSS basique.
     /// </summary>
     [MethodImpl(MethodImplOptions.AggressiveInlining)]
     public static bool IsPotentiallyDangerous(ReadOnlySpan<char> value)
     {
          if (value.IsEmpty)
               return false;

          // 1. Détection d'injection de formule en début de chaîne (après espaces éventuels)
          ReadOnlySpan<char> trimmedStart = value.TrimStart();
          if (!trimmedStart.IsEmpty && DangerousFormulaPrefixes.Contains(trimmedStart[0]))
               return true;

          // 2. Détection XSS / protocoles actifs
          if (value.Contains("<script", StringComparison.OrdinalIgnoreCase) ||
              value.Contains("</script", StringComparison.OrdinalIgnoreCase) ||
              value.Contains("javascript:", StringComparison.OrdinalIgnoreCase) ||
              value.Contains("vbscript:", StringComparison.OrdinalIgnoreCase) ||
              value.Contains("onerror=", StringComparison.OrdinalIgnoreCase) ||
              value.Contains("onload=", StringComparison.OrdinalIgnoreCase))
          {
               return true;
          }

          return false;
     }

     /// <summary>
     /// Neutralise une chaîne à l'écriture (sauf si la colonne est explicitement déclarée avec [XlsxFormula]).
     /// Préfixe par une apostrophe (') si elle commence par =, +, -, @, \t, \r et supprime les caractères de contrôle XML invalides.
     /// </summary>
     public static string SanitizeForExport(string? input, bool isExplicitFormula = false)
     {
          if (string.IsNullOrEmpty(input))
               return string.Empty;

          ReadOnlySpan<char> span = input.AsSpan();
          bool needsFormulaEscape = !isExplicitFormula && !span.IsEmpty && DangerousFormulaPrefixes.Contains(span[0]);

          // Nettoyage des caractères de contrôle ASCII interdits en XML 1.0 (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F)
          bool hasInvalidXmlChars = false;
          for (int i = 0; i < span.Length; i++)
          {
               char c = span[i];
               if (c < 0x20 && c != '\t' && c != '\n' && c != '\r')
               {
                    hasInvalidXmlChars = true;
                    break;
               }
          }

          if (!needsFormulaEscape && !hasInvalidXmlChars)
               return input;

          int extra = needsFormulaEscape ? 1 : 0;
          return string.Create(span.Length + extra, (input, needsFormulaEscape), static (dest, state) =>
          {
               int pos = 0;
               if (state.needsFormulaEscape)
                    dest[pos++] = '\'';

               ReadOnlySpan<char> src = state.input.AsSpan();
               for (int i = 0; i < src.Length; i++)
               {
                    char c = src[i];
                    dest[pos++] = (c < 0x20 && c != '\t' && c != '\n' && c != '\r') ? ' ' : c;
               }
          });
     }

     /// <summary>
     /// Applique les transformations [XlsxTrim] et [XlsxCase] sans allouer inutilement.
     /// </summary>
     public static string TransformText(string input, XlsxTrimMode trimMode, XlsxTextCase textCase)
     {
          if (string.IsNullOrEmpty(input))
               return input;

          string result = trimMode switch
          {
               XlsxTrimMode.Both => input.Trim(),
               XlsxTrimMode.Start => input.TrimStart(),
               XlsxTrimMode.End => input.TrimEnd(),
               XlsxTrimMode.AllWhitespace => RemoveAllWhitespace(input),
               _ => input
          };

          return textCase switch
          {
               XlsxTextCase.Upper => result.ToUpperInvariant(),
               XlsxTextCase.Lower => result.ToLowerInvariant(),
               XlsxTextCase.Title => System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(result.ToLowerInvariant()),
               _ => result
          };
     }

     private static string RemoveAllWhitespace(string input)
     {
          int count = 0;
          for (int i = 0; i < input.Length; i++)
               if (!char.IsWhiteSpace(input[i])) count++;

          if (count == input.Length)
               return input;

          return string.Create(count, input, static (dest, src) =>
          {
               int pos = 0;
               for (int i = 0; i < src.Length; i++)
               {
                    if (!char.IsWhiteSpace(src[i]))
                         dest[pos++] = src[i];
               }
          });
     }
}