using System.Buffers;
using System.Buffers.Binary;
using System.Globalization;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// Writes a document as SYLK, byte for byte as Excel 1.00's Save As with the SYLK type does (spec Appendix C).
/// </summary>
/// <remarks>
/// Each record is built in one reusable buffer of Mac OS Roman bytes and written to the output with a carriage
/// return. Cells are read from the file's own bytes wherever they are untouched, so the export creates no cell
/// objects.
/// </remarks>
internal sealed class SylkWriter : IDisposable
{
    // Excel's widths are in pixels; SYLK's in characters, and a column of 10 needs no record.
    private const int StandardWidth = 10;

    // Characters that Excel writes as ESC, 'N' and a letter (spec Appendix C): each pair is the letter and the Mac
    // OS Roman character, from Excel 1.00's table at PCOD 0 offset 0x2CB6.
    private static ReadOnlySpan<byte> PlainLetters =>
        [(byte)'`', 0xBD, (byte)'a', 0xAE, (byte)'c', 0xBB, (byte)'i', 0xAF, (byte)'j', 0xCE, (byte)'k', 0xBC, (byte)'q', 0xBE, (byte)'x', 0xA0, (byte)'y', 0xBF, (byte)'z', 0xCF, (byte)'{', 0xA7];

    private readonly ExcelDocument _document;
    private readonly IBufferWriter<byte> _output;
    private byte[] _line = ArrayPool<byte>.Shared.Rent(512);
    private int _length;

    // What the last records said, which Excel leaves out when it repeats: the cell position (Y, X) and an array
    // formula's or table's last cell (R, C), all 1-based.
    private int _row = -1;
    private int _column = -1;
    private int _sharedRow = -1;
    private int _sharedColumn = -1;

    // The default format (the F;D record). Excel writes a cell's format only where it differs from this.
    private CellFormat _default = new(-1, -1, 0);

    public SylkWriter(ExcelDocument document, IBufferWriter<byte> output)
    {
        _document = document;
        _output = output;
    }

    public void Dispose()
    {
        ArrayPool<byte>.Shared.Return(_line);
        _line = [];
    }

    public void Write()
    {
        _document.UpdateExtent();
        var state = _document.State;
        Append("ID;PXL;N;E"u8);
        EndRecord(escape: false);

        WriteNumberFormats();

        // The grid: rows 0 … last row − 1 and columns 0 … last column − 1 hold cells. In the last-index form a
        // margin follows, whose row and column hold the whole-column and whole-row formats (spec §6.1).
        int rows = state.GridRows;
        int columns = state.GridColumns;
        bool margin = state.HasGridMargin;
        WriteFormat(margin ? GetFormat(rows, columns) : default, 'D', 0);

        if (rows != 0)
        {
            // Written through the cell position, so the first cell leaves out a Y or X that equals the grid's size.
            Append("B"u8);
            AppendPosition(rows - 1, columns - 1);
            EndRecord(escape: false);
        }

        WriteOptions();
        WriteColumnWidths();
        if (margin)
        {
            for (int column = 0; column < columns; column++)
            {
                WriteFormat(GetFormat(rows, column), 'C', column + 1);
            }

            for (int row = 0; row < rows; row++)
            {
                WriteFormat(GetFormat(row, columns), 'R', row + 1);
            }
        }

        WriteNames();
        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                WriteCell(row, column);
            }
        }

        Append("E"u8);
        EndRecord(escape: false);
    }

    // P records: every format, built in and the document's own, as the Format Number dialog shows it.
    private void WriteNumberFormats()
    {
        foreach (string text in s_builtInText.Value)
        {
            WriteNumberFormat(text);
        }

        foreach (var format in _document.NumberFormats)
        {
            WriteNumberFormat(format.ToText(ExcelVersion.Excel100));
        }
    }

    private static readonly Lazy<string[]> s_builtInText = new(() => [.. ExcelNumberFormat.BuiltIn.Select(f => f.ToText(ExcelVersion.Excel100))]);

    private void WriteNumberFormat(string text)
    {
        Append("P;P"u8);
        int start = _length;
        AppendText(text);
        DoubleSemicolons(start);
        EndRecord(escape: true);
    }

    // The O record: document options.
    private void WriteOptions()
    {
        var state = _document.State;
        Append("O"u8);
        if (state.IsProtected)
        {
            Append(";P"u8);
        }

        if (state.ReferenceStyle == ExcelReferenceStyle.A1)
        {
            Append(";L"u8);
        }

        if (state.CalculationMode == ExcelCalculationMode.Manual)
        {
            Append(";M"u8);
        }

        if (!state.FullPrecision)
        {
            Append(";R"u8);
        }

        if (_document.Kind == ExcelDocumentKind.MacroSheet)
        {
            Append(";E"u8);
        }

        if (state.Iteration)
        {
            AppendField('A', state.MaxIterations);
            AppendByte((byte)' ');
            AppendNumber(state.MaxChange);
        }

        EndRecord(escape: false);
    }

    // F;W records: runs of columns with the same width in characters, except the standard width.
    private void WriteColumnWidths()
    {
        var state = _document.State;
        int digit = DigitWidth(state);
        int runStart = 0;
        int runWidth = StandardWidth;
        for (int column = 0; column < ExcelCellGrid.MaxColumns; column++)
        {
            int pixels = state.GetColumnWidth(column);
            int width = pixels == 0 || digit == 0 ? 0 : (pixels + digit / 2 - 5) / digit;
            if (width != runWidth)
            {
                WriteWidths(runStart + 1, column, runWidth);
                runStart = column;
                runWidth = width;
            }
        }

        WriteWidths(runStart + 1, ExcelCellGrid.MaxColumns, runWidth);
    }

    // Excel works out the width of 0 from the standard font when it opens a file, rather than trusting the cached
    // value: Expenses, saved by a build before 1.00, caches 8 for Geneva 10, which is 7 wide.
    private static int DigitWidth(ExcelSheetState state)
        => state.StandardFont == 3 && state.StandardFontSize == 10 ? 7 : state.DigitWidth;

    private void WriteWidths(int first, int last, int width)
    {
        if (width == StandardWidth)
        {
            return;
        }

        Append("F;W"u8);
        AppendNumber(first);
        AppendByte((byte)' ');
        AppendNumber(last);
        AppendByte((byte)' ');
        AppendNumber(width);
        EndRecord(escape: false);
    }

    // An F record. kind is 'D' for the defaults, 'C' or 'R' for a whole column or row (index is 1-based), and 0
    // for a cell, which it places with Y and X.
    private void WriteFormat(CellFormat format, char kind, int index, int row = 0, int column = 0)
    {
        if (format == _default)
        {
            return;
        }

        Append("F"u8);
        if (format.Number != _default.Number || format.Alignment != _default.Alignment)
        {
            AppendField('P', format.Number);
            AppendFormatCode(kind == 'D' ? 'D' : 'F', format);
            if (kind == 'D')
            {
                AppendNumber(StandardWidth);
            }
        }

        if (format.Style != _default.Style)
        {
            Append(";S"u8);
            int style = format.Style;
            AppendFlag(style, CellWords.ItalicBit, 'I');
            AppendFlag(style, CellWords.BoldBit, 'D');
            AppendFlag(style, CellWords.TopBorderBit, 'T');
            AppendFlag(style, CellWords.LeftBorderBit, 'L');
            AppendFlag(style, CellWords.BottomBorderBit, 'B');
            AppendFlag(style, CellWords.RightBorderBit, 'R');
        }

        switch (kind)
        {
            case 'C':
            case 'R':
                AppendField(kind, index);
                break;
            case 'D':
                WriteDefaultsTail();
                _default = format;
                break;
            default:
                AppendPosition(row, column);
                break;
        }

        EndRecord(escape: false);

        void AppendFlag(int style, int bit, char letter)
        {
            if ((style & (1 << bit)) != 0)
            {
                AppendByte((byte)letter);
            }
        }
    }

    // The rest of the F;D record: the front window's display options and the standard font.
    private void WriteDefaultsTail()
    {
        var window = _document.Windows.Count > 0 ? _document.Windows[0] : null;
        bool formulas = window?.DisplayFormulas ?? _document.Kind == ExcelDocumentKind.MacroSheet;
        if (formulas)
        {
            Append(";E"u8);
        }

        if (!(window?.DisplayGridlines ?? true))
        {
            Append(";G"u8);
        }

        if (!(window?.DisplayHeadings ?? true))
        {
            Append(";H"u8);
        }

        AppendField('N', _document.State.StandardFont);
        AppendByte((byte)' ');
        AppendNumber(_document.State.StandardFontSize);
    }

    // ;F or ;D, then the format's kind, its decimals and the alignment: G0G for General, $2C for currency with two
    // decimals, centred.
    private void AppendFormatCode(char field, CellFormat format)
    {
        var (code, digits) = FormatCode(format.Number);
        AppendByte((byte)';');
        AppendByte((byte)field);
        AppendByte((byte)(code == 'F' && digits == 0 ? 'I' : code));
        AppendNumber(digits);
        AppendByte((byte)((ExcelAlignment)format.Alignment switch
        {
            ExcelAlignment.General => 'G',
            ExcelAlignment.Left => 'L',
            ExcelAlignment.Center => 'C',
            ExcelAlignment.Right => 'R',
            _ => 'X',
        }));
    }

    // Classifies a format by its positive section (spec Appendix C): a number with % (two more scaling digits
    // than decimals), with a currency symbol first, or plain; scientific; or General for anything else, dates
    // included.
    private (char Code, int Digits) FormatCode(int index)
    {
        var format = index < ExcelNumberFormat.BuiltInCount ? ExcelNumberFormat.BuiltIn[index]
            : index - ExcelNumberFormat.BuiltInCount < _document.NumberFormats.Count ? _document.NumberFormats[index - ExcelNumberFormat.BuiltInCount]
            : null;
        var section = format is null ? default : format.GetSection(ExcelNumberFormatSection.Positive);
        if (section.Length < 4)
        {
            return ('G', 0);
        }

        int digits = section[2];
        return section[0] switch
        {
            0x03 when section[3] >= digits + 2 => ('%', digits),
            0x03 when section.Length >= 6 && section[4] == 0x22 => ('$', digits),
            0x03 => ('F', digits),
            0x1C => ('E', digits),
            _ => ('G', 0),
        };
    }

    // NN records: each defined name with a definition, with ;F for a function macro and ;G and its key for a
    // command.
    private void WriteNames()
    {
        Span<char> initial = stackalloc char[256];
        foreach (var name in _document.Names)
        {
            if (name.Definition.Length == 0)
            {
                continue;
            }

            Append("NN;N"u8);
            AppendText(name.Name);
            if ((name.Flags & ExcelNameFlags.Macro) != 0)
            {
                if ((name.RawMacroKind & 0x80) != 0)
                {
                    Append(";F"u8);
                }

                if ((name.RawMacroKind & 0x40) != 0)
                {
                    Append(";G"u8);
                    if (name.RawShortcutKey != 0)
                    {
                        AppendByte(name.RawShortcutKey);
                    }
                }
            }

            Append(";E"u8);
            int start = _length;
            var text = new TextBuffer(initial);
            try
            {
                FormulaDecompiler.AppendDefinitionR1C1(ref text, name, _document);
                AppendText(text.AsSpan());
            }
            finally
            {
                text.Dispose();
            }

            DoubleSemicolons(start);
            EndRecord(escape: true);
        }
    }

    private void WriteCell(int row, int column)
    {
        ReadOnlySpan<byte> data;
        ushort word1, word2;
        ExcelCell? cell = null;
        if (!_document.Cells.TryGetRawCell(row, column, out ushort word0, out word1, out word2, out data))
        {
            cell = _document.Cells.Find(row, column);
            _document.Cells.GetWords(row, column, out word0, out word1, out word2);
        }

        WriteFormat(new CellFormat(CellWords.NumberFormat(word1), (int)CellWords.Alignment(word1), word2 & 0x3F), (char)0, 0, row, column);

        var type = CellWords.Type(word1);
        bool locked = CellWords.Has(word2, CellWords.LockedBit);
        bool hidden = CellWords.Has(word2, CellWords.HiddenBit);
        if (type == ExcelCellType.Empty && locked && !hidden)
        {
            return;
        }

        Append("C"u8);
        AppendPosition(row, column);
        if (!locked)
        {
            Append(";N"u8);
        }

        if (hidden)
        {
            Append(";H"u8);
        }

        switch (type)
        {
            case ExcelCellType.Empty:
                break;
            case ExcelCellType.Integer:
                Append(";K"u8);
                AppendNumber(word0);
                break;
            case ExcelCellType.Number:
                Append(";K"u8);
                AppendNumber(cell is null ? BinaryPrimitives.ReadDoubleBigEndian(data) : cell.Value.Number);
                break;
            case ExcelCellType.Text:
                Append(";K"u8);
                if (cell is null)
                {
                    AppendTextValue(data.Slice(1, data[0]));
                }
                else
                {
                    AppendTextValue(cell.Value.Text);
                }

                break;
            case ExcelCellType.BooleanOrError:
                Append(";K"u8);
                AppendValue(word0 >> 8 == 0 ? ExcelValue.FromBoolean((word0 & 0xFF) != 0) : ExcelValue.FromErrorCode((byte)word0));
                break;
            default:
                WriteFormulaFields(row, column, cell, data);
                return;
        }

        EndRecord(escape: true);
    }

    private void WriteFormulaFields(int row, int column, ExcelCell? cell, ReadOnlySpan<byte> record)
    {
        ReadOnlySpan<byte> expression;
        ReadOnlySpan<byte> fixups;
        Append(";K"u8);
        if (cell is null)
        {
            ExcelFormula.MeasureRecord(record, out int fixupsOffset, out int textOffset);
            expression = record.Slice(15, record[14]);
            fixups = record[fixupsOffset..];
            if (textOffset >= 0)
            {
                // A text result follows the expression: copy its bytes rather than make a string.
                AppendTextValue(record.Slice(textOffset + 1, record[textOffset]));
            }
            else
            {
                AppendValue(ExcelValue.FromFormulaBytes(record, null));
            }
        }
        else
        {
            var formula = cell.Formula!;
            expression = formula.Expression;
            fixups = formula.Fixups;
            AppendValue(formula.Result);
        }

        var at = new ExcelCellAddress(row, column);
        if (FormulaTokens.IsSharedReference(expression))
        {
            WriteSharedFields(at, expression, fixups);
            return;
        }

        Append(";E"u8);
        AppendFormula(expression, fixups, at);
        EndRecord(escape: true);
        WriteExternalLink(expression, fixups, at, default);
    }

    // An array formula or a table (spec §12.6). The anchor gives the range's last cell, then ;M and the shared
    // expression or ;T and the table's inputs; every other cell gives ;I and the anchor.
    private void WriteSharedFields(ExcelCellAddress at, ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups)
    {
        bool table = FormulaTokens.Normalize(expression[0]) == 0x02;
        bool anchor = fixups.Length >= 2 && BinaryPrimitives.ReadUInt16BigEndian(fixups) == FormulaTokens.SharedAnchorMarker;
        if (!anchor)
        {
            Append(";I"u8);
            if (fixups.Length >= 3)
            {
                AppendShared(fixups[2], BinaryPrimitives.ReadUInt16BigEndian(fixups) & RowWord.RowMask);
            }

            EndRecord(escape: true);
            return;
        }

        // The shared record: first row, last row, first column, last column.
        if (fixups.Length < (table ? 3 + 18 : 13) || (!table && fixups.Length < 13 + fixups[12]))
        {
            throw new ExcelFormatException($"The shared record of {at} runs past the end of its cell data.");
        }

        var shared = fixups.Slice(3);
        int lastRow = BinaryPrimitives.ReadUInt16BigEndian(shared.Slice(2)) & RowWord.RowMask;
        AppendShared(shared[5], lastRow);
        if (table)
        {
            Append(";T"u8);
            AppendTableInputs(shared.Slice(0, 18));
            EndRecord(escape: true);
            return;
        }

        int length = shared[9];
        var sharedExpression = shared.Slice(10, length);
        var sharedFixups = shared.Slice(10 + length);
        Append(";M"u8);
        AppendFormula(sharedExpression, sharedFixups, at);
        EndRecord(escape: true);
        var range = new ExcelRange(at, new ExcelCellAddress(lastRow, shared[5]));
        WriteExternalLink(sharedExpression, sharedFixups, at, range);
    }

    private void AppendShared(int column, int row)
    {
        if (row + 1 != _sharedRow)
        {
            _sharedRow = row + 1;
            AppendField('R', _sharedRow);
        }

        if (column + 1 != _sharedColumn)
        {
            _sharedColumn = column + 1;
            AppendField('C', _sharedColumn);
        }
    }

    // TABLE's arguments: the row input, then the column input, as absolute references, #REF! for a deleted one.
    private void AppendTableInputs(ReadOnlySpan<byte> record)
    {
        byte flags = record[14];
        bool firstIsRow = (flags & 0x40) != 0;
        bool hasSecond = (flags & 0x80) != 0;
        if (firstIsRow)
        {
            AppendInput(record, 9);
        }

        AppendByte((byte)',');
        if (!firstIsRow)
        {
            AppendInput(record, 9);
        }
        else if (hasSecond)
        {
            AppendInput(record, 15);
        }

        void AppendInput(ReadOnlySpan<byte> record, int at)
        {
            ushort word = BinaryPrimitives.ReadUInt16BigEndian(record.Slice(at + 1));
            if ((word & 0x01) != 0)
            {
                Append("#REF!"u8);
                return;
            }

            AppendField('R', (word >> 2) + 1, separator: false);
            AppendField('C', record[at] + 1, separator: false);
        }
    }

    private void AppendFormula(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups, ExcelCellAddress at)
    {
        int start = _length;
        var text = new TextBuffer(stackalloc char[256]);
        try
        {
            FormulaDecompiler.AppendR1C1(ref text, expression, fixups, _document, at);
            AppendText(text.AsSpan());
        }
        finally
        {
            text.Dispose();
        }

        DoubleSemicolons(start);
    }

    // An NE record for a formula that is nothing but an external reference (=Budget!A1): the document, the
    // referenced cells or name, and the cell or array range that uses it.
    private void WriteExternalLink(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups, ExcelCellAddress at, ExcelRange range)
    {
        if (!FormulaTokens.TryGetExternalLink(expression, fixups, out int document, out int referenceOffset, out ushort index))
        {
            return;
        }

        Append("NE;F"u8);
        if (document >= 1 && document <= _document.ExternalDocuments.Count)
        {
            AppendText(_document.ExternalDocuments[document - 1].Name);
        }

        Append(";S"u8);
        byte code = FormulaTokens.Normalize(expression[referenceOffset]);
        int p = referenceOffset + 1;
        switch (code)
        {
            case 0x23:
            {
                var names = document >= 1 && document <= _document.ExternalDocuments.Count ? _document.ExternalDocuments[document - 1].Names : null;
                if (names is not null && index >= 1 && index <= names.Count)
                {
                    AppendText(names[index - 1]);
                }

                break;
            }

            case 0x24:
                AppendAbsoluteArea(RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p))), expression[p + 2],
                    RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p))), expression[p + 2]);
                break;
            case 0x25:
                AppendAbsoluteArea(RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p))), expression[p + 4],
                    RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p + 2))), expression[p + 5]);
                break;
        }

        Append(";E"u8);
        if (range.First == range.Last)
        {
            AppendAbsoluteArea(at.Row, at.Column, at.Row, at.Column);
        }
        else
        {
            AppendAbsoluteArea(range.First.Row, range.First.Column, range.Last.Row, range.Last.Column);
        }

        EndRecord(escape: true);
    }

    private void AppendAbsoluteArea(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        AppendField('R', firstRow + 1, separator: false);
        AppendField('C', firstColumn + 1, separator: false);
        if (firstRow != lastRow || firstColumn != lastColumn)
        {
            AppendByte((byte)':');
            AppendField('R', lastRow + 1, separator: false);
            AppendField('C', lastColumn + 1, separator: false);
        }
    }

    // ;Y and ;X where they differ from the last ones written.
    private void AppendPosition(int row, int column)
    {
        if (row + 1 != _row)
        {
            _row = row + 1;
            AppendField('Y', _row);
        }

        if (column + 1 != _column)
        {
            _column = column + 1;
            AppendField('X', _column);
        }
    }

    // A value after ;K: a number as General shows it, text in quotes with its quotes as they are, TRUE, FALSE or an
    // error.
    private void AppendValue(ExcelValue value)
    {
        switch (value.Kind)
        {
            case ExcelValueKind.Number:
                AppendNumber(value.Number);
                break;
            case ExcelValueKind.Text:
                AppendTextValue(value.Text);
                break;
            case ExcelValueKind.Boolean:
                AppendText(BooleanText.Of(value.Boolean));
                break;
            case ExcelValueKind.Error:
                AppendText(value.Error.ToText());
                break;
        }
    }

    private void AppendTextValue(ReadOnlySpan<byte> macRoman)
    {
        AppendByte((byte)'"');
        int start = _length;
        Append(macRoman);
        DoubleSemicolons(start);
        AppendByte((byte)'"');
    }

    private void AppendTextValue(string text)
    {
        AppendByte((byte)'"');
        int start = _length;
        AppendText(text);
        DoubleSemicolons(start);
        AppendByte((byte)'"');
    }

    private void AppendField(char letter, int value, bool separator = true)
    {
        if (separator)
        {
            AppendByte((byte)';');
        }

        AppendByte((byte)letter);
        AppendNumber(value);
    }

    private void AppendNumber(int value)
    {
        Ensure(11);
        value.TryFormat(_line.AsSpan(_length), out int written, default, CultureInfo.InvariantCulture);
        _length += written;
    }

    // As General shows it with room for 21 characters (19.0c in Excel 1.00).
    private void AppendNumber(double value)
    {
        Span<char> text = stackalloc char[ExcelValueFormatter.GeneralSize(21)];
        AppendText(text[..ExcelValueFormatter.FormatGeneral(value, 21, text)]);
    }

    private void AppendText(ReadOnlySpan<char> text)
    {
        Ensure(text.Length);
        foreach (char c in text)
        {
            _line[_length++] = MacRoman.GetByte(c);
        }
    }

    private void Append(ReadOnlySpan<byte> bytes)
    {
        Ensure(bytes.Length);
        bytes.CopyTo(_line.AsSpan(_length));
        _length += bytes.Length;
    }

    private void AppendByte(byte b)
    {
        Ensure(1);
        _line[_length++] = b;
    }

    // SYLK separates fields with ';', so a ';' inside one is doubled.
    private void DoubleSemicolons(int start)
    {
        int count = _line.AsSpan(start, _length - start).Count((byte)';');
        if (count == 0)
        {
            return;
        }

        Ensure(count);
        for (int from = _length - 1, to = _length + count - 1; from >= start; from--)
        {
            _line[to--] = _line[from];
            if (_line[from] == ';')
            {
                _line[to--] = (byte)';';
            }
        }

        _length += count;
    }

    private void Ensure(int more)
    {
        if (_length + more > _line.Length)
        {
            var larger = ArrayPool<byte>.Shared.Rent(Math.Max(_line.Length * 2, _length + more));
            _line.AsSpan(0, _length).CopyTo(larger);
            ArrayPool<byte>.Shared.Return(_line);
            _line = larger;
        }
    }

    // Writes the record and a carriage return. In records that carry text, a character below 32 or above 126 is
    // written as ESC and two or three printable characters (spec Appendix C).
    private void EndRecord(bool escape)
    {
        var line = _line.AsSpan(0, _length);
        _length = 0;
        if (!escape || !line.ContainsAnyExceptInRange((byte)32, (byte)126))
        {
            var span = _output.GetSpan(line.Length + 1);
            line.CopyTo(span);
            span[line.Length] = (byte)'\r';
            _output.Advance(line.Length + 1);
            return;
        }

        var destination = _output.GetSpan(line.Length * 4 + 1);
        int n = 0;
        foreach (byte b in line)
        {
            if (b is >= 32 and < 127)
            {
                destination[n++] = b;
                continue;
            }

            destination[n++] = 0x1B;
            if (TryGetEscapeLetters(b, out byte accent, out byte letter))
            {
                destination[n++] = (byte)'N';
                if (accent != 0)
                {
                    destination[n++] = accent;
                }

                destination[n++] = letter;
            }
            else
            {
                destination[n++] = (byte)((b >> 4) + 0x20);
                destination[n++] = (byte)((b & 0x0F) + 0x30);
            }
        }

        destination[n++] = (byte)'\r';
        _output.Advance(n);
    }

    // Excel's ESC N forms: a letter alone for a character from its first table, or an accent (A grave, B acute,
    // C circumflex, D tilde, H umlaut, J ring, K cedilla) and the letter it sits on.
    private static bool TryGetEscapeLetters(byte character, out byte accent, out byte letter)
    {
        var plain = PlainLetters;
        for (int i = 0; i < plain.Length; i += 2)
        {
            if (plain[i + 1] == character)
            {
                accent = 0;
                letter = plain[i];
                return true;
            }
        }

        foreach (var (code, pairs) in s_accents)
        {
            for (int i = 0; i < pairs.Length; i += 2)
            {
                if (pairs[i + 1] == character)
                {
                    accent = code;
                    letter = pairs[i];
                    return true;
                }
            }
        }

        accent = letter = 0;
        return false;
    }

    // Excel 1.00's accent table (PCOD 0 offset 0x2CA5): for each accent letter, pairs of the base letter and the Mac
    // OS Roman character.
    private static readonly (byte Code, byte[] Pairs)[] s_accents =
    [
        ((byte)'A', [(byte)'A', 0xCB, (byte)'a', 0x88, (byte)'e', 0x8F, (byte)'i', 0x93, (byte)'o', 0x98, (byte)'u', 0x9D]),
        ((byte)'B', [(byte)' ', 0xAB, (byte)'E', 0x83, (byte)'a', 0x87, (byte)'e', 0x8E, (byte)'i', 0x92, (byte)'o', 0x97, (byte)'u', 0x9C]),
        ((byte)'C', [(byte)'a', 0x89, (byte)'e', 0x90, (byte)'i', 0x94, (byte)'o', 0x99, (byte)'u', 0x9E]),
        ((byte)'D', [(byte)'A', 0xCC, (byte)'N', 0x84, (byte)'O', 0xCD, (byte)'a', 0x8B, (byte)'n', 0x96, (byte)'o', 0x9B]),
        ((byte)'H', [(byte)' ', 0xAC, (byte)'A', 0x80, (byte)'O', 0x85, (byte)'U', 0x86, (byte)'a', 0x8A, (byte)'e', 0x91, (byte)'i', 0x95, (byte)'o', 0x9A, (byte)'u', 0x9F, (byte)'y', 0xD8]),
        ((byte)'J', [(byte)' ', 0xA1, (byte)'A', 0x81, (byte)'a', 0x8C]),
        ((byte)'K', [(byte)'C', 0x82, (byte)'c', 0x8D]),
    ];

    private CellFormat GetFormat(int row, int column)
    {
        _document.Cells.GetWords(row, column, out _, out ushort word1, out ushort word2);
        return new CellFormat(CellWords.NumberFormat(word1), (int)CellWords.Alignment(word1), word2 & 0x3F);
    }

    private readonly record struct CellFormat(int Number, int Alignment, int Style);
}
