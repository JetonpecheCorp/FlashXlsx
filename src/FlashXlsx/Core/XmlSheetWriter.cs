using FlashXlsx.Abstractions.Internal;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FlashXlsx.Core;

internal sealed record XlsxDropdownRule(int ColumnIndex, IReadOnlyList<string> Values);

internal sealed class XmlSheetWriter : IDisposable
{
     private static readonly SearchValues<char> XmlEscapeChars = SearchValues.Create("&<>\"'");

     private readonly Stream _targetStream;
     private readonly MemoryStream? _bufferedSheetData;
     private readonly Stream _activeStream;
     private readonly ArrayBufferWriter<byte> _buffer;
     private readonly char[] _charScratch;
     private readonly int[]? _colMaxWidths;
     private readonly bool _hasHeaderStyle;
     private readonly bool _hasExcelTable;
     private readonly bool _freezeHeader;
     private readonly IReadOnlyList<XlsxDropdownRule>? _dropdowns;
     private readonly IReadOnlyDictionary<int, int>? _columnStyleIndices;

     private int _currentRowNumber;
     private bool _isHeaderRow;

     public int RowCount { get; private set; }
     public int ColumnCount { get; }

     public XmlSheetWriter(
         Stream entryStream,
         int columnCount = 0,
         bool autoFitColumns = false,
         bool hasHeaderStyle = false,
         bool hasExcelTable = false,
         bool freezeHeader = false,
         IReadOnlyList<XlsxDropdownRule>? dropdowns = null,
         IReadOnlyDictionary<int, int>? columnStyleIndices = null)
     {
          _targetStream = entryStream;
          ColumnCount = columnCount;
          _hasHeaderStyle = hasHeaderStyle;
          _hasExcelTable = hasExcelTable;
          _freezeHeader = freezeHeader;
          _dropdowns = dropdowns;
          _columnStyleIndices = columnStyleIndices;
          _buffer = new ArrayBufferWriter<byte>(65536);
          _charScratch = ArrayPool<char>.Shared.Rent(256);

          if (autoFitColumns && columnCount > 0)
          {
               _colMaxWidths = new int[columnCount];
               _bufferedSheetData = new MemoryStream(65536);
               _activeStream = _bufferedSheetData;
               WriteRawUtf8("<sheetData>"u8);
          }
          else
          {
               _activeStream = _targetStream;
               WriteXmlHeader();
               WriteRawUtf8("<sheetData>"u8);
          }
     }

     private void WriteXmlHeader()
     {
          WriteRawUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8);
          if (_hasExcelTable)
               WriteRawUtf8("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8);
          else
               WriteRawUtf8("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8);

          if (_freezeHeader)
          {
               WriteRawUtf8("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>"u8);
          }
     }

     public void StartRow(int rowNumber, bool isHeader = false)
     {
          _currentRowNumber = rowNumber;
          _isHeaderRow = isHeader;
          if (rowNumber > RowCount)
               RowCount = rowNumber;

          WriteRawUtf8("<row r=\""u8);
          WriteInt32(rowNumber);
          WriteRawUtf8("\">"u8);
     }

     public void EndRow()
     {
          WriteRawUtf8("</row>"u8);
          if (_buffer.WrittenCount >= 32768)
               FlushBuffer();
     }

     public void WriteCell(
         int columnIndex,
         XlsxCellWriteKind kind,
         ReadOnlySpan<char> formattedSpan,
         string? stringValue)
     {
          if (kind == XlsxCellWriteKind.Empty)
               return;

          if (_colMaxWidths is not null && (uint)columnIndex < (uint)_colMaxWidths.Length)
          {
               int displayLen = kind switch
               {
                    XlsxCellWriteKind.InlineString => stringValue?.Length ?? formattedSpan.Length,
                    XlsxCellWriteKind.Date => 10,
                    XlsxCellWriteKind.DateTime => 19,
                    XlsxCellWriteKind.Time => 8,
                    XlsxCellWriteKind.Formula => 12,
                    _ => formattedSpan.Length + 4
               };
               if (displayLen > _colMaxWidths[columnIndex])
                    _colMaxWidths[columnIndex] = displayLen;
          }

          int refLen = XlsxCoordinateHelper.FormatCellReference(columnIndex, _currentRowNumber, _charScratch);
          ReadOnlySpan<char> cellRef = _charScratch.AsSpan(0, refLen);

          switch (kind)
          {
               case XlsxCellWriteKind.Number:
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    if (_columnStyleIndices is not null && _columnStyleIndices.TryGetValue(columnIndex, out int numStyleIdx))
                    {
                         WriteRawUtf8("\" s=\""u8);
                         WriteInt32(numStyleIdx);
                    }
                    WriteRawUtf8("\"><v>"u8);
                    WriteUtf16AsUtf8(formattedSpan);
                    WriteRawUtf8("</v></c>"u8);
                    break;

               case XlsxCellWriteKind.Date:
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    WriteRawUtf8("\" s=\"1\"><v>"u8);
                    WriteUtf16AsUtf8(formattedSpan);
                    WriteRawUtf8("</v></c>"u8);
                    break;

               case XlsxCellWriteKind.DateTime:
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    WriteRawUtf8("\" s=\"2\"><v>"u8);
                    WriteUtf16AsUtf8(formattedSpan);
                    WriteRawUtf8("</v></c>"u8);
                    break;

               case XlsxCellWriteKind.Time:
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    WriteRawUtf8("\" s=\"3\"><v>"u8);
                    WriteUtf16AsUtf8(formattedSpan);
                    WriteRawUtf8("</v></c>"u8);
                    break;

               case XlsxCellWriteKind.Boolean:
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    WriteRawUtf8("\" t=\"b\"><v>"u8);
                    WriteUtf16AsUtf8(formattedSpan);
                    WriteRawUtf8("</v></c>"u8);
                    break;

               case XlsxCellWriteKind.InlineString:
                    ReadOnlySpan<char> text = stringValue is not null ? stringValue.AsSpan() : formattedSpan;
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    if (_isHeaderRow && _hasHeaderStyle)
                         WriteRawUtf8("\" s=\"4\" t=\"inlineStr\"><is><t>"u8);
                    else
                         WriteRawUtf8("\" t=\"inlineStr\"><is><t>"u8);
                    WriteEscapedXmlText(text);
                    WriteRawUtf8("</t></is></c>"u8);
                    break;

               case XlsxCellWriteKind.Formula:
                    if (string.IsNullOrEmpty(stringValue))
                         return;

                    string expr = stringValue.StartsWith('=') ? stringValue.Substring(1) : stringValue;
                    WriteRawUtf8("<c r=\""u8);
                    WriteUtf16AsUtf8(cellRef);
                    if (_columnStyleIndices is not null && _columnStyleIndices.TryGetValue(columnIndex, out int formStyleIdx))
                    {
                         WriteRawUtf8("\" s=\""u8);
                         WriteInt32(formStyleIdx);
                    }
                    WriteRawUtf8("\"><f>"u8);
                    WriteEscapedXmlText(expr.AsSpan());
                    WriteRawUtf8("</f></c>"u8);
                    break;
          }
     }

     public async ValueTask CompleteAsync(CancellationToken ct = default)
     {
          WriteRawUtf8("</sheetData>"u8);

          if (_dropdowns is { Count: > 0 })
          {
               int lastRow = Math.Max(2, RowCount);
               WriteRawUtf8("<dataValidations count=\""u8);
               WriteInt32(_dropdowns.Count);
               WriteRawUtf8("\">"u8);

               for (int i = 0; i < _dropdowns.Count; i++)
               {
                    var dv = _dropdowns[i];
                    string colLetter = XlsxCoordinateHelper.GetColumnLetter(dv.ColumnIndex);
                    string sqref = $"{colLetter}2:{colLetter}{lastRow}";
                    string joinedList = string.Join(",", dv.Values).Replace("\"", "\"\"");

                    WriteRawUtf8("<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\""u8);
                    WriteUtf16AsUtf8(sqref.AsSpan());
                    WriteRawUtf8("\"><formula1>\""u8);
                    WriteEscapedXmlText(joinedList.AsSpan());
                    WriteRawUtf8("\"</formula1></dataValidation>"u8);
               }

               WriteRawUtf8("</dataValidations>"u8);
          }

          if (_hasExcelTable)
          {
               WriteRawUtf8("<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>"u8);
          }

          WriteRawUtf8("</worksheet>"u8);
          await FlushBufferAsync(ct);

          if (_bufferedSheetData is not null && _colMaxWidths is not null)
          {
               var headerWriter = new ArrayBufferWriter<byte>(4096);
               void WriteHeaderUtf8(ReadOnlySpan<byte> b) { b.CopyTo(headerWriter.GetSpan(b.Length)); headerWriter.Advance(b.Length); }

               WriteHeaderUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8);
               if (_hasExcelTable)
                    WriteHeaderUtf8("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8);
               else
                    WriteHeaderUtf8("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8);

               if (_freezeHeader)
               {
                    WriteHeaderUtf8("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>"u8);
               }

               WriteHeaderUtf8("<cols>"u8);
               for (int c = 0; c < _colMaxWidths.Length; c++)
               {
                    double width = Math.Clamp(_colMaxWidths[c] + 5.0, 10.0, 80.0);
                    string colTag = string.Format(
                        CultureInfo.InvariantCulture,
                        "<col min=\"{0}\" max=\"{0}\" width=\"{1:0.##}\" bestFit=\"1\" customWidth=\"1\"/>",
                        c + 1,
                        width);
                    int maxBytes = Encoding.UTF8.GetMaxByteCount(colTag.Length);
                    int written = Encoding.UTF8.GetBytes(colTag, headerWriter.GetSpan(maxBytes));
                    headerWriter.Advance(written);
               }
               WriteHeaderUtf8("</cols>"u8);

               await _targetStream.WriteAsync(headerWriter.WrittenMemory, ct);
               _bufferedSheetData.Position = 0;
               await _bufferedSheetData.CopyToAsync(_targetStream, ct);
               await _bufferedSheetData.DisposeAsync();
          }
     }

     private void WriteEscapedXmlText(ReadOnlySpan<char> text)
     {
          while (!text.IsEmpty)
          {
               int idx = text.IndexOfAny(XmlEscapeChars);
               if (idx < 0)
               {
                    WriteUtf16AsUtf8(text);
                    break;
               }

               if (idx > 0)
                    WriteUtf16AsUtf8(text[..idx]);

               switch (text[idx])
               {
                    case '&': WriteRawUtf8("&amp;"u8); break;
                    case '<': WriteRawUtf8("&lt;"u8); break;
                    case '>': WriteRawUtf8("&gt;"u8); break;
                    case '"': WriteRawUtf8("&quot;"u8); break;
                    case '\'': WriteRawUtf8("&apos;"u8); break;
               }

               text = text[(idx + 1)..];
          }
     }

     private void WriteRawUtf8(ReadOnlySpan<byte> utf8)
     {
          var span = _buffer.GetSpan(utf8.Length);
          utf8.CopyTo(span);
          _buffer.Advance(utf8.Length);
     }

     private void WriteUtf16AsUtf8(ReadOnlySpan<char> chars)
     {
          if (chars.IsEmpty) return;
          int maxBytes = Encoding.UTF8.GetMaxByteCount(chars.Length);
          var span = _buffer.GetSpan(maxBytes);
          int written = Encoding.UTF8.GetBytes(chars, span);
          _buffer.Advance(written);
     }

     private void WriteInt32(int value)
     {
          var span = _buffer.GetSpan(16);
          value.TryFormat(span, out int written, provider: CultureInfo.InvariantCulture);
          _buffer.Advance(written);
     }

     private void FlushBuffer()
     {
          if (_buffer.WrittenCount > 0)
          {
               _activeStream.Write(_buffer.WrittenSpan);
               _buffer.Clear();
          }
     }

     private async ValueTask FlushBufferAsync(CancellationToken ct)
     {
          if (_buffer.WrittenCount > 0)
          {
               await _activeStream.WriteAsync(_buffer.WrittenMemory, ct);
               _buffer.Clear();
          }
     }

     public void Dispose()
     {
          FlushBuffer();
          ArrayPool<char>.Shared.Return(_charScratch);
     }
}