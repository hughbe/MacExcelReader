using System.Buffers.Binary;
using System.Globalization;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// Converts parsed expressions (spec §15) to A1-style formula text.
/// </summary>
internal static class FormulaDecompiler
{
    /// <summary>
    /// Converts a parsed expression to text, without the leading '='.
    /// </summary>
    /// <param name="expression">The expression bytes.</param>
    /// <param name="fixupWords">Name and external-document indices in token order (cell formulas).</param>
    /// <param name="document">The document used to resolve names and external documents; may be null.</param>
    /// <param name="fixupsInPlace">True for name definitions, whose indices are stored inside the tokens.</param>
    public static string Decompile(ReadOnlySpan<byte> expression, IReadOnlyList<ushort> fixupWords, ExcelDocument? document, bool fixupsInPlace)
        => Decompile(expression, fixupWords, document, fixupsInPlace, null);

    /// <summary>
    /// Converts a parsed expression to text, without the leading '=', rendering array constants from their
    /// data blocks (spec §12.5), one per <c>0x20</c> token in order.
    /// </summary>
    public static string Decompile(ReadOnlySpan<byte> expression, IReadOnlyList<ushort> fixupWords, ExcelDocument? document, bool fixupsInPlace, IReadOnlyList<byte[]>? arrays)
        => Decompile(expression, fixupWords, document, fixupsInPlace, arrays, default);

    /// <summary>
    /// Converts a parsed expression to text, without the leading '=', placing the relative references of a name
    /// definition (tokens <c>0x2C</c>, <c>0x2D</c>), which hold offsets (spec §11), relative to
    /// <paramref name="relativeTo"/>, as Excel shows them with that cell active.
    /// </summary>
    public static string Decompile(ReadOnlySpan<byte> expression, IReadOnlyList<ushort> fixupWords, ExcelDocument? document, bool fixupsInPlace, IReadOnlyList<byte[]>? arrays, ExcelCellAddress relativeTo)
    {
        ArgumentNullException.ThrowIfNull(fixupWords);
        var fixups = new FixupSource(fixupWords, arrays, fixupsInPlace);
        return Decompile(expression, ref fixups, document, relativeTo, default, default);
    }

    /// <summary>
    /// Converts a cell formula's expression to text between <paramref name="prefix"/> and <paramref name="suffix"/>,
    /// reading name indices and array constants from its fix-up bytes as it goes.
    /// </summary>
    internal static string Decompile(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixupBytes, ExcelDocument? document, string prefix, string suffix)
    {
        var fixups = new FixupSource(fixupBytes);
        return Decompile(expression, ref fixups, document, default, prefix, suffix);
    }

    private static string Decompile(ReadOnlySpan<byte> expression, ref FixupSource fixups, ExcelDocument? document, ExcelCellAddress relativeTo, ReadOnlySpan<char> prefix, ReadOnlySpan<char> suffix)
    {
        var text = new TextBuffer(stackalloc char[256]);
        try
        {
            text.Append(prefix);
            Write(ref text, expression, fixups, document, new References(relativeTo, r1c1: false));
            text.Append(suffix);
            return text.ToString();
        }
        finally
        {
            text.Dispose();
        }
    }

    /// <summary>
    /// Appends a cell formula's expression in R1C1 style, as Excel's SYLK export writes it: relative references as
    /// offsets from <paramref name="cell"/>, such as <c>R[-1]C[+2]</c>.
    /// </summary>
    internal static void AppendR1C1(ref TextBuffer text, ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixupBytes, ExcelDocument? document, ExcelCellAddress cell)
    {
        var fixups = new FixupSource(fixupBytes);
        Write(ref text, expression, fixups, document, new References(cell, r1c1: true));
    }

    /// <summary>
    /// Appends a name definition in R1C1 style. Its relative references already hold offsets (spec §11).
    /// </summary>
    internal static void AppendDefinitionR1C1(ref TextBuffer text, ExcelDefinedName name, ExcelDocument? document)
    {
        var fixups = new FixupSource([], name.HasArrayConstants ? name.ArrayConstants : null, inPlace: true);
        Write(ref text, name.Definition, fixups, document, new References(default, r1c1: true));
    }

    // How references are written: A1 style, with a definition's relative parts placed relative to Origin, or R1C1
    // style, with a cell formula's relative parts written as offsets from Origin, the formula's cell.
    private readonly struct References(ExcelCellAddress origin, bool r1c1)
    {
        public ExcelCellAddress Origin { get; } = origin;

        public bool R1C1 { get; } = r1c1;
    }

    private static void Write(ref TextBuffer text, scoped ReadOnlySpan<byte> expression, FixupSource fixups, ExcelDocument? document, References references)
    {
        // In reverse Polish the operands of each operator are the last entries on the stack, and their text is the
        // end of the buffer, in order, so an entry needs only its start: combining operands inserts the operator
        // text between them and drops the later entries.
        Span<int> starts = stackalloc int[256];
        Span<char> numberText = stackalloc char[ExcelValueFormatter.GeneralSize(20)];
        int count = 0;
        {
            int external = 0;
            int pendingExternal = -1;

            foreach (var token in FormulaTokens.Enumerate(expression))
            {
                int p = token.Offset + 1;
                switch (token.Code)
                {
                    case 0x01:
                    case 0x02:
                        Push(ref text, starts, ref count);
                        text.Append(token.Code == 0x01 ? ExcelFormula.ArrayFormulaPlaceholder : ExcelFormula.TablePlaceholder);
                        break;
                    case >= 0x03 and <= 0x11:
                        Require(ref text, starts, ref count, 2);
                        text.Insert(starts[count - 1], FormulaTokens.BinaryOperatorText(token.Code));
                        count--;
                        break;
                    case 0x12:
                    case 0x13:
                        Require(ref text, starts, ref count, 1);
                        text.Insert(starts[count - 1], token.Code == 0x12 ? "+" : "-");
                        break;
                    case 0x14:
                        Require(ref text, starts, ref count, 1);
                        text.Append('%');
                        break;
                    case 0x15:
                        Require(ref text, starts, ref count, 1);
                        text.Insert(starts[count - 1], "(");
                        text.Append(')');
                        break;
                    case 0x16:
                        Push(ref text, starts, ref count);
                        break;
                    case 0x17:
                        Push(ref text, starts, ref count);
                        text.Append('"');
                        text.AppendMacRoman(expression.Slice(p + 1, expression[p]), escapeQuotes: true);
                        text.Append('"');
                        break;
                    case 0x19:
                        if ((expression[p] & 0x10) != 0)
                        {
                            Require(ref text, starts, ref count, 1);
                            text.Insert(starts[count - 1], "SUM(");
                            text.Append(')');
                        }

                        break;
                    case 0x1A:
                        pendingExternal = fixups.NextWord(expression, token);
                        external = pendingExternal;
                        break;
                    case 0x1B:
                        external = 0;
                        break;
                    case 0x1C:
                        Push(ref text, starts, ref count);
                        text.Append(((ExcelError)expression[p]).ToText());
                        break;
                    case 0x1D:
                        Push(ref text, starts, ref count);
                        text.Append(BooleanText.Of(expression[p] != 0));
                        break;
                    case 0x1E:
                        Push(ref text, starts, ref count);
                        text.AppendFormatted(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p)), default);
                        break;
                    case 0x1F:
                        Push(ref text, starts, ref count);
                        // Excel 1.00 writes a number constant as General shows it, with at most 14 significant digits (spec
                        // §15.4), so text with more digits comes back shorter.
                        text.Append(numberText[..ExcelValueFormatter.FormatGeneral(BinaryPrimitives.ReadDoubleBigEndian(expression.Slice(p)), 20, numberText)]);
                        break;
                    case 0x20:
                        Push(ref text, starts, ref count);
                        if (fixups.TryNextArray(out var block))
                        {
                            AppendArray(ref text, block);
                        }
                        else
                        {
                            text.Append("{array}");
                        }

                        break;
                    case 0x21:
                    case 0x39:
                    {
                        // 0x39: Excel 1.5's form for functions numbered 256 and up.
                        int number = expression[p] + (token.Code == 0x39 ? 256 : 0);
                        var function = ExcelFunctions.Get(number);
                        AppendCall(ref text, starts, ref count, function, number, function?.MinArguments ?? 0, dialog: false);
                        break;
                    }

                    case 0x22:
                    case 0x3A:
                    {
                        // Bit 7 of the argument count marks a command's dialog-box form, such as RUN?().
                        int number = expression[p + 1] + (token.Code == 0x3A ? 256 : 0);
                        AppendCall(ref text, starts, ref count, ExcelFunctions.Get(number), number, expression[p] & 0x7F, dialog: (expression[p] & 0x80) != 0);
                        break;
                    }

                    case 0x23:
                    case 0x30:
                    {
                        ushort index = fixups.NextWord(expression, token);
                        PushOperand(ref text, starts, ref count, document, ref pendingExternal);
                        // Inside an external reference a name token indexes that document's names (spec §10).
                        if (external != 0)
                        {
                            AppendExternalName(ref text, document, external, index);
                        }
                        else
                        {
                            AppendName(ref text, document, index);
                        }

                        break;
                    }

                    case 0x24:
                    case 0x31:
                    case 0x37:
                    case 0x2C:
                    {
                        PushOperand(ref text, starts, ref count, document, ref pendingExternal);
                        ushort rowWord = BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p));
                        bool definition = token.Code == 0x2C;
                        if (references.R1C1)
                        {
                            AppendCellR1C1(ref text, rowWord, expression[p + 2], references.Origin, offsets: definition);
                        }
                        else
                        {
                            AppendCell(ref text, rowWord, expression[p + 2], definition ? references.Origin : default);
                        }

                        break;
                    }

                    case 0x25:
                    case 0x32:
                    case 0x2D:
                    {
                        // A definition's area (0x2D) holds offsets for its relative parts (spec §11).
                        bool definition = token.Code == 0x2D;
                        PushOperand(ref text, starts, ref count, document, ref pendingExternal);
                        AppendArea(ref text, BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p)), BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p + 2)),
                            expression[p + 4], expression[p + 5], definition, references);
                        break;
                    }

                    case 0x2A:
                    case 0x2B:
                        PushOperand(ref text, starts, ref count, document, ref pendingExternal);
                        text.Append(ExcelError.Reference.ToText());
                        break;
                    case FormulaTokens.MemoryArea:
                        // The areas it stands for are in the fix-ups; the tokens that follow give the text.
                        fixups.SkipMemoryArea();
                        break;
                    default:
                        // Other memory tokens (0x27–0x29, 0x2E, 0x2F, 0x33–0x36) carry no text of their own.
                        break;
                }
            }

            // Operands left over are written one after another, separated by spaces.
            for (int i = count - 1; i >= 1; i--)
            {
                text.Insert(starts[i], " ");
            }
        }
    }

    // Starts a new operand at the end of the text.
    private static void Push(ref TextBuffer text, scoped Span<int> starts, ref int count)
    {
        if (count == starts.Length)
        {
            throw new ExcelFormatException("The formula has too many operands.");
        }

        starts[count++] = text.Length;
    }

    // Starts an operand that an external sheet prefix (a 0x1A token before it) applies to, writing the prefix.
    private static void PushOperand(ref TextBuffer text, scoped Span<int> starts, ref int count, ExcelDocument? document, ref int pendingExternal)
    {
        Push(ref text, starts, ref count);
        if (pendingExternal >= 0)
        {
            AppendExternalDocumentName(ref text, document, pendingExternal);
            text.Append('!');
            pendingExternal = -1;
        }
    }

    // Makes sure the stack has at least `needed` operands, adding "?" for missing ones before the others, where the
    // missing (earlier) operands would have been.
    private static void Require(ref TextBuffer text, scoped Span<int> starts, ref int count, int needed)
    {
        while (count < needed)
        {
            int at = count == 0 ? text.Length : starts[0];
            text.Insert(at, "?");
            for (int i = count; i > 0; i--)
            {
                starts[i] = starts[i - 1] + 1;
            }

            starts[0] = at;
            count++;
        }
    }

    // Combines the last `argumentCount` operands into a call: NAME(a,b,c).
    private static void AppendCall(ref TextBuffer text, scoped Span<int> starts, ref int count, ExcelFunction? function, int number, int argumentCount, bool dialog)
    {
        Require(ref text, starts, ref count, argumentCount);
        int first = count - argumentCount;
        for (int i = count - 1; i > first; i--)
        {
            text.Insert(starts[i], ",");
        }

        // Function 255 calls a user-defined function macro; its first argument is the macro.
        if (number == 255 && argumentCount > 0)
        {
            int end = argumentCount > 1 ? starts[first + 1] : text.Length;
            if (argumentCount > 1)
            {
                // The comma after the macro becomes the opening parenthesis.
                text.Remove(end, 1);
            }

            text.Insert(end, "(");
            text.Append(')');
            count = first + 1;
            return;
        }

        int start = argumentCount > 0 ? starts[first] : text.Length;
        if (argumentCount == 0)
        {
            Push(ref text, starts, ref count);
            first = count - 1;
        }

        Span<char> fallback = stackalloc char[16];
        ReadOnlySpan<char> name = function?.Name ?? FunctionFallback(fallback, number);
        text.Insert(start, dialog ? "?(" : "(");
        text.Insert(start, name);
        text.Append(')');
        count = first + 1;
    }

    private static ReadOnlySpan<char> FunctionFallback(Span<char> buffer, int number)
    {
        "FUNCTION".AsSpan().CopyTo(buffer);
        number.TryFormat(buffer[8..], out int written, default, CultureInfo.InvariantCulture);
        return buffer[..(8 + written)];
    }

    /// <summary>
    /// Formats an array-constant block (column count, row count, then elements) as <c>{1,2;3,4}</c>.
    /// Element type 1 is a number, 2 a string, 4 TRUE or FALSE and 16 an error (spec §12.5).
    /// </summary>
    public static string ArrayText(ReadOnlySpan<byte> block)
    {
        var text = new TextBuffer(stackalloc char[256]);
        try
        {
            AppendArray(ref text, block);
            return text.ToString();
        }
        finally
        {
            text.Dispose();
        }
    }

    private static void AppendArray(ref TextBuffer text, ReadOnlySpan<byte> block)
    {
        int columns = block[0] == 0 ? 256 : block[0];
        int rows = BinaryPrimitives.ReadUInt16BigEndian(block.Slice(1));
        text.Append('{');
        int p = 3;
        for (int r = 0; r < rows; r++)
        {
            if (r > 0)
            {
                text.Append(';');
            }

            for (int c = 0; c < columns; c++)
            {
                if (c > 0)
                {
                    text.Append(',');
                }

                byte type = block[p++];
                switch (type)
                {
                    case 2:
                        text.Append('"');
                        text.AppendMacRoman(block.Slice(p + 1, block[p]), escapeQuotes: true);
                        text.Append('"');
                        p += 1 + block[p];
                        continue;
                    case 1:
                        text.AppendFormatted(BinaryPrimitives.ReadDoubleBigEndian(block.Slice(p)), "R");
                        break;
                    case 4:
                        text.Append(BooleanText.Of(block[p + 1] != 0));
                        break;
                    case 16:
                        text.Append(((ExcelError)block[p + 1]).ToText());
                        break;
                    default:
                        text.Append('?');
                        text.AppendFormatted(type, default);
                        break;
                }

                p += 8;
            }
        }

        text.Append('}');
    }

    // Where name indices, external-document indices and array constants come from: inside the tokens (a name
    // definition), lists of them, or a cell formula's fix-up bytes, read in token order (spec §12.5).
    private ref struct FixupSource
    {
        private readonly IReadOnlyList<ushort>? _words;
        private readonly IReadOnlyList<byte[]>? _arrays;
        private readonly bool _inPlace;
        private readonly ReadOnlySpan<byte> _bytes;
        private readonly bool _fromBytes;
        private int _position;
        private int _wordIndex;
        private int _arrayIndex;

        public FixupSource(IReadOnlyList<ushort> words, IReadOnlyList<byte[]>? arrays, bool inPlace)
        {
            _words = words;
            _arrays = arrays;
            _inPlace = inPlace;
        }

        public FixupSource(ReadOnlySpan<byte> bytes)
        {
            _bytes = bytes;
            _fromBytes = true;
        }

        public ushort NextWord(scoped ReadOnlySpan<byte> expression, FormulaToken token)
        {
            if (_inPlace)
            {
                return token.Code == FormulaTokens.ExternalSheet
                    ? expression[token.Offset + 7]
                    : BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(token.Offset + 1));
            }

            if (_fromBytes)
            {
                return _bytes.IsEmpty ? (ushort)0 : FormulaTokens.ReadWord(_bytes, ref _position);
            }

            return _wordIndex < _words!.Count ? _words[_wordIndex++] : (ushort)0;
        }

        public bool TryNextArray(out ReadOnlySpan<byte> block)
        {
            if (_fromBytes)
            {
                if (_bytes.IsEmpty)
                {
                    block = default;
                    return false;
                }

                int start = _position;
                FormulaTokens.SkipArrayBlock(_bytes, ref _position);
                block = _bytes[start.._position];
                return true;
            }

            if (_arrays is not null && _arrayIndex < _arrays.Count)
            {
                block = _arrays[_arrayIndex++];
                return true;
            }

            block = default;
            return false;
        }

        public void SkipMemoryArea()
        {
            if (_fromBytes && !_bytes.IsEmpty)
            {
                FormulaTokens.Skip(_bytes, ref _position, FormulaTokens.ReadWord(_bytes, ref _position) * 6);
            }
        }
    }

    /// <summary>
    /// Formats a cell reference from a row word (bit 15 = relative row, bit 14 = relative column) and a column.
    /// </summary>
    public static string CellText(ushort rowWord, byte column)
    {
        Span<char> buffer = stackalloc char[16];
        var text = new TextBuffer(buffer);
        AppendCell(ref text, rowWord, column, default);
        return text.ToString();
    }

    // A reference. In a name definition its relative row and column are offsets from the cell the name is used in
    // (spec §11), 14-bit and 8-bit two's complement, shown relative to `relativeTo`. Excel wraps an offset past the
    // edge of the sheet round to the other side, so B1's R[-3] is row 16382. In a cell formula `relativeTo` is A1.
    private static void AppendCell(ref TextBuffer text, ushort rowWord, byte column, ExcelCellAddress relativeTo)
    {
        (rowWord, column) = RowWord.Move(rowWord, column, relativeTo.Row, relativeTo.Column);
        if (!RowWord.IsColumnRelative(rowWord))
        {
            text.Append('$');
        }

        AppendColumnName(ref text, column);
        if (!RowWord.IsRowRelative(rowWord))
        {
            text.Append('$');
        }

        text.AppendFormatted(RowWord.Row(rowWord) + 1, default);
    }

    // An area. One that spans every column is written as rows, 1:3 or $18:$19, and one that spans every row as columns,
    // A:B; Excel 1.00 shows =A1:IV1 as =1:1 and =$A$1:$IV$16384 as =$1:$16384. A row or column is absolute when its
    // own flag says so.
    private static void AppendArea(ref TextBuffer text, ushort firstRowWord, ushort lastRowWord, byte firstColumn, byte lastColumn, bool definition, References references)
    {
        if (references.R1C1)
        {
            AppendAreaR1C1(ref text, firstRowWord, lastRowWord, firstColumn, lastColumn, references.Origin, definition);
            return;
        }

        var relativeTo = definition ? references.Origin : default;
        (firstRowWord, firstColumn) = RowWord.Move(firstRowWord, firstColumn, relativeTo.Row, relativeTo.Column);
        (lastRowWord, lastColumn) = RowWord.Move(lastRowWord, lastColumn, relativeTo.Row, relativeTo.Column);
        if (firstColumn == 0 && lastColumn == ExcelCellGrid.MaxColumns - 1)
        {
            AppendRow(ref text, firstRowWord);
            text.Append(':');
            AppendRow(ref text, lastRowWord);
        }
        else if (RowWord.Row(firstRowWord) == 0 && RowWord.Row(lastRowWord) == ExcelCellGrid.MaxRows - 1)
        {
            AppendColumn(ref text, firstRowWord, firstColumn);
            text.Append(':');
            AppendColumn(ref text, lastRowWord, lastColumn);
        }
        else
        {
            AppendCell(ref text, firstRowWord, firstColumn, default);
            text.Append(':');
            AppendCell(ref text, lastRowWord, lastColumn, default);
        }

        static void AppendRow(ref TextBuffer text, ushort rowWord)
        {
            if (!RowWord.IsRowRelative(rowWord))
            {
                text.Append('$');
            }

            text.AppendFormatted(RowWord.Row(rowWord) + 1, default);
        }

        static void AppendColumn(ref TextBuffer text, ushort rowWord, byte column)
        {
            if (!RowWord.IsColumnRelative(rowWord))
            {
                text.Append('$');
            }

            AppendColumnName(ref text, column);
        }
    }

    // A reference in R1C1 style: R5C2 for an absolute row and column, R[-1]C[+2] for relative ones, and R or C for
    // an offset of 0. A cell formula's relative parts are positions, written as offsets from `origin`; a name
    // definition's are already offsets (spec §11).
    private static void AppendCellR1C1(ref TextBuffer text, ushort rowWord, byte column, ExcelCellAddress origin, bool offsets)
    {
        AppendRowR1C1(ref text, rowWord, origin, offsets);
        AppendColumnR1C1(ref text, rowWord, column, origin, offsets);
    }

    // An area in R1C1 style. As in A1 style, one that spans every column is written as rows (R1:R3) and one that
    // spans every row as columns (C1:C2).
    private static void AppendAreaR1C1(ref TextBuffer text, ushort firstRowWord, ushort lastRowWord, byte firstColumn, byte lastColumn, ExcelCellAddress origin, bool offsets)
    {
        // Unlike A1 style, a single row or column is written once: R[-20], not R[-20]:R[-20].
        if (firstColumn == 0 && lastColumn == ExcelCellGrid.MaxColumns - 1)
        {
            AppendRowR1C1(ref text, firstRowWord, origin, offsets);
            if (lastRowWord != firstRowWord)
            {
                text.Append(':');
                AppendRowR1C1(ref text, lastRowWord, origin, offsets);
            }
        }
        else if (RowWord.Row(firstRowWord) == 0 && RowWord.Row(lastRowWord) == ExcelCellGrid.MaxRows - 1)
        {
            AppendColumnR1C1(ref text, firstRowWord, firstColumn, origin, offsets);
            if (lastColumn != firstColumn || (lastRowWord & RowWord.ColumnRelative) != (firstRowWord & RowWord.ColumnRelative))
            {
                text.Append(':');
                AppendColumnR1C1(ref text, lastRowWord, lastColumn, origin, offsets);
            }
        }
        else
        {
            AppendCellR1C1(ref text, firstRowWord, firstColumn, origin, offsets);
            text.Append(':');
            AppendCellR1C1(ref text, lastRowWord, lastColumn, origin, offsets);
        }
    }

    private static void AppendRowR1C1(ref TextBuffer text, ushort rowWord, ExcelCellAddress origin, bool offsets)
    {
        text.Append('R');
        int row = RowWord.Row(rowWord);
        if (!RowWord.IsRowRelative(rowWord))
        {
            text.AppendFormatted(row + 1, default);
        }
        else
        {
            // Excel 1.00 reads a definition's 14-bit row offset with the relative-column flag (bit 14) as its sign
            // bit, inverted: with an absolute column −1 is R[-1], but with a relative column −3 is R[+16381]
            // (Relative Names sample; spec Appendix C).
            AppendOffset(ref text, offsets ? row - (RowWord.IsColumnRelative(rowWord) ? 0 : RowWord.RowMask + 1) : row - origin.Row);
        }
    }

    private static void AppendColumnR1C1(ref TextBuffer text, ushort rowWord, byte column, ExcelCellAddress origin, bool offsets)
    {
        text.Append('C');
        if (!RowWord.IsColumnRelative(rowWord))
        {
            text.AppendFormatted(column + 1, default);
        }
        else
        {
            AppendOffset(ref text, offsets ? (sbyte)column : column - origin.Column);
        }
    }

    private static void AppendOffset(ref TextBuffer text, int offset)
    {
        if (offset == 0)
        {
            return;
        }

        text.Append('[');
        if (offset > 0)
        {
            text.Append('+');
        }

        text.AppendFormatted(offset, default);
        text.Append(']');
    }

    /// <summary>
    /// Gets the letters of a zero-based column index (0 = A, 255 = IV).
    /// </summary>
    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        Span<char> buffer = stackalloc char[8];
        var text = new TextBuffer(buffer);
        AppendColumnName(ref text, column);
        return text.ToString();
    }

    private static void AppendColumnName(ref TextBuffer text, int column)
    {
        // Bijective base 26: A–Z, then AA–ZZ, and so on; at most 7 letters for an int.
        Span<char> letters = stackalloc char[8];
        int length = 0;
        for (column++; column > 0; column = (column - 1) / 26)
        {
            letters[^++length] = (char)('A' + (column - 1) % 26);
        }

        text.Append(letters[^length..]);
    }

    private static void AppendName(ref TextBuffer text, ExcelDocument? document, ushort index)
    {
        if (document is not null && index >= 1 && index <= document.Names.Count)
        {
            text.Append(document.Names[index - 1].Name);
        }
        else
        {
            text.Append("NAME");
            text.AppendFormatted(index, default);
        }
    }

    private static void AppendExternalName(ref TextBuffer text, ExcelDocument? document, int external, ushort index)
    {
        if (document is not null && external <= document.ExternalDocuments.Count
            && document.ExternalDocuments[external - 1].Names is var names && index >= 1 && index <= names.Count)
        {
            text.Append(names[index - 1]);
        }
        else
        {
            text.Append("NAME");
            text.AppendFormatted(index, default);
        }
    }

    private static void AppendExternalDocumentName(ref TextBuffer text, ExcelDocument? document, int index)
    {
        // External-document indices are 1-based. A document named "\0" (or "") means the active sheet,
        // written as !A1.
        if (document is null || index < 1 || index > document.ExternalDocuments.Count)
        {
            text.Append("DOC");
            text.AppendFormatted(index, default);
            return;
        }

        var name = document.ExternalDocuments[index - 1].Name;
        if (name.Length == 0 || name == "\0")
        {
            return;
        }

        foreach (char c in name)
        {
            if (!(char.IsLetterOrDigit(c) || c is '.' or '_'))
            {
                // Quoted, with quotes doubled: 'Budget Plan''s'.
                text.Append('\'');
                foreach (char d in name)
                {
                    if (d == '\'')
                    {
                        text.Append('\'');
                    }

                    text.Append(d);
                }

                text.Append('\'');
                return;
            }
        }

        text.Append(name);
    }
}
