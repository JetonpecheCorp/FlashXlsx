using System;
using System.Collections.Generic;
using System.Text;

namespace FlashXlsx.Core;

internal static class XlsxCoordinateHelper
{
     /// <summary>
     /// Extrait l'index de colonne (base 0) depuis une référence de cellule Excel (ex: "C14" -> 2, "AA1" -> 26).
     /// Zéro allocation.
     /// </summary>
     public static int ParseColumnIndex(ReadOnlySpan<char> cellReference)
     {
          int columnIndex = 0;
          for (int i = 0; i < cellReference.Length; i++)
          {
               char c = cellReference[i];
               if ((uint)(c - 'A') <= ('Z' - 'A'))
               {
                    columnIndex = (columnIndex * 26) + (c - 'A' + 1);
               }
               else if ((uint)(c - 'a') <= ('z' - 'a'))
               {
                    columnIndex = (columnIndex * 26) + (c - 'a' + 1);
               }
               else
               {
                    // Dès qu'on atteint les chiffres (ex: '1' dans "C14"), on s'arrête
                    break;
               }
          }
          return columnIndex - 1;
     }

     /// <summary>
     /// Convertit un index de colonne (base 0) en lettres Excel (ex: 0 -> "A", 26 -> "AA").
     /// </summary>
     public static string GetColumnLetter(int columnIndex)
     {
          if (columnIndex < 0) return string.Empty;

          Span<char> buffer = stackalloc char[8];
          int pos = buffer.Length;
          int current = columnIndex + 1;

          while (current > 0)
          {
               current--;
               buffer[--pos] = (char)('A' + (current % 26));
               current /= 26;
          }

          return buffer[pos..].ToString();
     }

     /// <summary>
     /// Écrit directement la référence de cellule (ex: "C14") dans un buffer Span sans allouer de string.
     /// </summary>
     public static int FormatCellReference(int columnIndex, int rowNumber, Span<char> destination)
     {
          Span<char> letters = stackalloc char[8];
          int pos = letters.Length;
          int current = columnIndex + 1;

          while (current > 0)
          {
               current--;
               letters[--pos] = (char)('A' + (current % 26));
               current /= 26;
          }

          ReadOnlySpan<char> colSpan = letters[pos..];
          colSpan.CopyTo(destination);

          rowNumber.TryFormat(destination[colSpan.Length..], out int digitsWritten);
          return colSpan.Length + digitsWritten;
     }
}
