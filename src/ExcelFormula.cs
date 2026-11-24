using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A formula record (spec §12.3): the current value, the parsed expression and the fix-ups that follow it.
/// </summary>
/// <remarks>
/// The record is held as a single byte array in its on-disk layout:
/// value (8) · reserved (3) · column (1) · row word (2) · expression length (1) · expression ·
/// text result (if any) · fix-ups.
/// </remarks>
public sealed class ExcelFormula
{
    private const int ExpressionOffset = 15;

    private readonly byte[] _record;
    private readonly int _fixupsOffset;

    // Offset of a text result still undecoded in _record, or -1.
    private int _rawTextOffset = -1;
    private string? _text;

    /// <summary>
    /// Creates a formula from a parsed expression.
    /// </summary>
    /// <param name="expression">The parsed expression (spec §15), at most 255 bytes.</param>
    /// <param name="fixups">
    /// The fix-up bytes that follow the expression (spec §12.5): one word per name token (its 1-based
    /// index in the name list) and per external-sheet token, in token order. Empty if none are needed.
    /// </param>
    public ExcelFormula(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups = default)
    {
        if (expression.Length > 255)
        {
            throw new ArgumentException("An expression cannot be longer than 255 bytes.", nameof(expression));
        }

        _ = FormulaTokens.CountTokens(expression);
        _record = new byte[ExpressionOffset + expression.Length + fixups.Length];
        _record[14] = (byte)expression.Length;
        expression.CopyTo(_record.AsSpan(ExpressionOffset));
        _fixupsOffset = ExpressionOffset + expression.Length;
        fixups.CopyTo(_record.AsSpan(_fixupsOffset));
    }

    private ExcelFormula(byte[] record, int fixupsOffset, int rawTextOffset)
    {
        _record = record;
        _fixupsOffset = fixupsOffset;
        _rawTextOffset = rawTextOffset;
    }

    /// <summary>Gets the parsed expression.</summary>
    public ReadOnlySpan<byte> Expression => _record.AsSpan(ExpressionOffset, _record[14]);

    /// <summary>Gets the raw fix-up bytes that follow the expression (and the text result, if any).</summary>
    public ReadOnlySpan<byte> Fixups => _record.AsSpan(_fixupsOffset);

    /// <summary>
    /// Gets bytes 8–10 of the record: the formula's link in Excel's list of formulas waiting to be recalculated
    /// (spec §12.3). Open puts a formula whose first two bytes here are neither <c>00 00</c> nor <c>FF FE</c> back on
    /// the list, so Excel recalculates it as soon as the file is open; formulas the library creates have zeros.
    /// </summary>
    public ReadOnlySpan<byte> Reserved => _record.AsSpan(8, 3);

    /// <summary>Gets the column stored in the record (the writer replaces it with the cell's column).</summary>
    public byte StoredColumn => _record[11];

    /// <summary>Gets the row word stored in the record: the row in bits 13–0 and two flags in bits 15–14.</summary>
    public ushort StoredRowWord => BinaryPrimitives.ReadUInt16BigEndian(_record.AsSpan(12));

    /// <summary>Gets the raw 8 value bytes.</summary>
    public ReadOnlySpan<byte> RawValue => _record.AsSpan(0, 8);

    // Whether the current value is a number: the first two value bytes are not FF FF (spec §12.4).
    internal bool HasNumericResult => _record[0] != 0xFF || _record[1] != 0xFF;

    private bool IsTextResult => _record[0] == 0xFF && _record[1] == 0xFF && _record[2] == 0;

    private string? TextResult
    {
        get
        {
            if (_rawTextOffset >= 0)
            {
                _text = MacRoman.GetString(_record.AsSpan(_rawTextOffset + 1, _record[_rawTextOffset]));
                _rawTextOffset = -1;
            }

            return _text;
        }
    }

    /// <summary>Gets the formula's current value: what Excel last calculated, and shows.</summary>
    public ExcelValue Result => ExcelValue.FromFormulaBytes(RawValue, IsTextResult ? TextResult : null);

    /// <summary>
    /// Sets the current value: a number, text, TRUE or FALSE, or an error (<see cref="ExcelValue.FromError"/>).
    /// The library does not calculate, so this is the value Excel shows until it recalculates.
    /// </summary>
    /// <exception cref="ArgumentException">The value is empty, which a formula cannot have, text longer than 255
    /// characters, or a number that is infinite or not a number.</exception>
    public void SetResult(ExcelValue value)
    {
        if (value.Kind == ExcelValueKind.Text)
        {
            PascalText.Check(value.Text, nameof(value));
        }
        else if (value.Kind == ExcelValueKind.Number && !double.IsFinite(value.Number))
        {
            throw new ArgumentOutOfRangeException(nameof(value), value.Number, "Excel numbers are finite: it gives #NUM! for an overflow.");
        }

        value.WriteFormulaBytes(_record.AsSpan(0, 8));
        _rawTextOffset = -1;
        _text = value.Kind == ExcelValueKind.Text ? value.Text : null;
    }

    /// <summary>
    /// Decodes the current value of a formula record in a file without creating the formula (spec §12.3).
    /// </summary>
    internal static ExcelValue ReadResult(ReadOnlySpan<byte> record)
    {
        if (record.Length < ExpressionOffset)
        {
            throw new ExcelFormatException("A formula record runs past the end of the file.");
        }

        string? text = null;
        if (record[0] == 0xFF && record[1] == 0xFF && record[2] == 0)
        {
            int at = ExpressionOffset + record[14];
            if (at >= record.Length || at + 1 + record[at] > record.Length)
            {
                throw new ExcelFormatException("A formula's text result runs past the end of the file.");
            }

            text = MacRoman.GetString(record.Slice(at + 1, record[at]));
        }

        return ExcelValue.FromFormulaBytes(record, text);
    }

    /// <summary>
    /// Compiles A1-style formula text, such as <c>=SUM($A$1:$A$5)*2</c>, into a formula (spec §15), choosing token
    /// classes and <c>IF</c> jump tokens the way Excel 1.x does.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Supported: numbers, strings, TRUE and FALSE, error values, cell references and areas in A1 notation
    /// (with <c>$</c> for absolute parts), external references (<c>Budget!A1</c>, <c>'My Doc'!A1</c>),
    /// defined names, every Excel 1.x function and macro command, calls to user-defined function macros,
    /// the operators <c>: , % ^ * / + - &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;=</c> and parentheses,
    /// references to the active sheet (<c>!A1</c>) and the dialog-box forms of commands (<c>RUN?()</c>).
    /// </para>
    /// <para>
    /// Like Excel, it compiles IF and CHOOSE with jump attributes, SUM with one argument to a SUM
    /// attribute, folds a minus sign into a number constant, starts formulas that call volatile
    /// functions (RAND, NOW, INDEX, …) with a volatile attribute, and wraps ranges built from function results
    /// (<c>INDEX(A1:C3,1,1):C3</c>) in a memory-function token.
    /// </para>
    /// <para>
    /// Excel's own parser was not disassembled in full. The token classes and <c>IF</c> encoding are derived from
    /// Excel's function table and from every formula in the sample documents, all of which the parser
    /// reproduces byte for byte (apart from bytes left over from memory inside name tokens, which Excel rebuilds when
    /// it opens a file).
    /// </para>
    /// </remarks>
    /// <param name="text">The formula text; the leading '=' is optional.</param>
    /// <param name="document">Resolves defined names and external documents; may be null if there are none.</param>
    /// <param name="version">
    /// The Excel release whose functions and encoding to use. Excel 1.00–1.03 (the default) produce formulas every
    /// Excel 1.x can read; Excel 1.5 adds functions and encodes a few things differently (spec §15.4).
    /// </param>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public static ExcelFormula Parse(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
        => ExcelFormulaParser.Parse(text, document, version);

    /// <summary>
    /// Gets the name indices and external-document indices from the fix-ups, in token order, without allocating.
    /// </summary>
    public FormulaTokens.FixupWordCollection FixupWords => new(Expression, Fixups);

    /// <summary>
    /// Converts the formula to A1-style text, such as <c>=SUM(A1:A5)</c>.
    /// </summary>
    /// <param name="document">The document, used to resolve names and external documents; may be null.</param>
    public string ToText(ExcelDocument? document = null)
    {
        if (IsArrayFormula || IsTable)
        {
            return SharedText(document);
        }

        return FormulaDecompiler.Decompile(Expression, Fixups, document, "=", string.Empty);
    }

    /// <summary>
    /// Writes the formula's A1-style text, as <see cref="ToText"/> gives it, into <paramref name="destination"/>
    /// without allocating (an array formula or table cell aside, whose text is built from its shared record).
    /// </summary>
    /// <returns>False if the text does not fit in <paramref name="destination"/>.</returns>
    public bool TryFormatText(Span<char> destination, out int charsWritten, ExcelDocument? document = null)
    {
        if (IsArrayFormula || IsTable)
        {
            string shared = SharedText(document);
            charsWritten = shared.TryCopyTo(destination) ? shared.Length : 0;
            return charsWritten == shared.Length;
        }

        var text = new TextBuffer(stackalloc char[256]);
        try
        {
            FormulaDecompiler.Append(ref text, Expression, Fixups, document, "=", default);
            return text.TryCopyTo(destination, out charsWritten);
        }
        finally
        {
            text.Dispose();
        }
    }

    /// <summary>Gets whether the cell is part of an array formula (spec §12.6).</summary>
    public bool IsArrayFormula => FormulaTokens.IsSharedReference(Expression) && FormulaTokens.Normalize(Expression[0]) == 0x01;

    /// <summary>Gets whether the cell is part of a data table made with the Table command (spec §12.6).</summary>
    public bool IsTable => FormulaTokens.IsSharedReference(Expression) && FormulaTokens.Normalize(Expression[0]) == 0x02;

    /// <summary>
    /// Gets whether this cell holds the shared record of its array formula or table (the top-left cell);
    /// other cells of the range refer to it.
    /// </summary>
    public bool IsSharedAnchor => (IsArrayFormula || IsTable) && Fixups.Length >= 2 && BinaryPrimitives.ReadUInt16BigEndian(Fixups) == FormulaTokens.SharedAnchorMarker;

    /// <summary>
    /// Gets the cell that holds the shared record, for a cell of an array formula or table.
    /// </summary>
    public ExcelCellAddress? SharedAnchor
    {
        get
        {
            if (!(IsArrayFormula || IsTable) || Fixups.Length < 3)
            {
                return null;
            }

            if (IsSharedAnchor)
            {
                // The anchor's shared header starts with the range: first row word, last row word, first column.
                return Fixups.Length >= 13 ? new ExcelCellAddress(RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(Fixups.Slice(3))), Fixups[7]) : null;
            }

            return new ExcelCellAddress(RowWord.Row(BinaryPrimitives.ReadUInt16BigEndian(Fixups)), Fixups[2]);
        }
    }

    /// <summary>
    /// Gets the range an array formula or table covers, from its anchor's shared header.
    /// </summary>
    public ExcelRange? SharedRange
        => IsSharedAnchor && Fixups.Length >= 13
            ? new ExcelRange(BinaryPrimitives.ReadUInt16BigEndian(Fixups.Slice(3)), Fixups[7], BinaryPrimitives.ReadUInt16BigEndian(Fixups.Slice(5)), Fixups[8])
            : null;

    /// <summary>
    /// Gets a table's input cells, from its anchor's 18-byte record (spec §12.6): the row input and the column
    /// input, either of which may be absent. Null for anything but a table's anchor.
    /// </summary>
    public (ExcelCellAddress? RowInput, ExcelCellAddress? ColumnInput)? TableInputs
    {
        get
        {
            if (!IsTable || !IsSharedAnchor || Fixups.Length < 3 + 18)
            {
                return null;
            }

            // Record bytes 9–10 and 15–16: an input's column byte, then a word with its row in bits 15–2. Byte 14:
            // bit 6 = the first input is the row input, bit 7 = there is a second input (the column input).
            var record = Fixups.Slice(3, 18);
            var first = new ExcelCellAddress(BinaryPrimitives.ReadUInt16BigEndian(record.Slice(10)) >> 2, record[9]);
            var second = new ExcelCellAddress(BinaryPrimitives.ReadUInt16BigEndian(record.Slice(16)) >> 2, record[15]);
            byte flags = record[14];
            return (flags & 0x40) != 0 ? (first, (flags & 0x80) != 0 ? second : null) : (null, first);
        }
    }

    /// <summary>
    /// Gets whether the table's row input cell has been deleted (bit 0 of the input's row word); Excel then shows
    /// <c>#REF!</c> in its place.
    /// </summary>
    public bool RowInputDeleted => TableInputDeleted(row: true);

    /// <summary>Gets whether the table's column input cell has been deleted.</summary>
    public bool ColumnInputDeleted => TableInputDeleted(row: false);

    private bool TableInputDeleted(bool row)
    {
        if (!IsTable || !IsSharedAnchor || Fixups.Length < 21)
        {
            return false;
        }

        var record = Fixups.Slice(3, 18);
        bool firstIsRow = (record[14] & 0x40) != 0;
        bool hasSecond = (record[14] & 0x80) != 0;
        // The first input is the row input when bit 6 is set, otherwise the column input; the second is the column input.
        int word = row ? (firstIsRow ? 10 : -1) : (firstIsRow ? (hasSecond ? 16 : -1) : 10);
        return word >= 0 && (record[word + 1] & 0x01) != 0;
    }

    /// <summary>
    /// Gets the shared expression of an array formula, held by its anchor cell; empty for other formulas.
    /// </summary>
    public ReadOnlySpan<byte> SharedExpression
        => IsArrayFormula && IsSharedAnchor && Fixups.Length >= 13 && Fixups.Length >= 13 + Fixups[12] ? Fixups.Slice(13, Fixups[12]) : default;

    /// <summary>Gets the fix-ups of the shared expression (see <see cref="SharedExpression"/>).</summary>
    public ReadOnlySpan<byte> SharedFixups
        => SharedExpression.IsEmpty ? default : Fixups.Slice(13 + Fixups[12]);

    private static string Absolute(ExcelCellAddress? address)
    {
        if (address is null)
        {
            return string.Empty;
        }

        string text = address.ToString()!;
        int digits = text.AsSpan().IndexOfAnyInRange('0', '9');
        return "$" + text[..digits] + "$" + text[digits..];
    }

    private string SharedText(ExcelDocument? document)
    {
        if (!IsSharedAnchor)
        {
            // Another cell holds the shared record.
            if (document is not null && SharedAnchor is { } anchor
                && document.Cells.Find(anchor.Row, anchor.Column)?.Formula is { IsSharedAnchor: true } shared)
            {
                return shared.SharedText(document);
            }

            return IsTable ? TablePlaceholder : ArrayFormulaPlaceholder;
        }

        if (IsTable)
        {
            // The inputs are fixed cells: Excel's GET.FORMULA gives them as absolute references (R7C25).
            return TableInputs is var (row, column)
                ? $"{{=TABLE({(RowInputDeleted ? ExcelError.Reference.ToText() : Absolute(row))},{(ColumnInputDeleted ? ExcelError.Reference.ToText() : Absolute(column))})}}"
                : TablePlaceholder;
        }

        if (Fixups.Length < 13)
        {
            return ArrayFormulaPlaceholder;
        }

        // Anchor: 0x4000, 0, a 10-byte header whose last byte is the shared expression's length, the shared
        // expression, then its fix-ups.
        int length = Fixups[12];
        if (Fixups.Length < 13 + length)
        {
            return ArrayFormulaPlaceholder;
        }

        var expression = Fixups.Slice(13, length);
        var fixups = Fixups.Slice(13 + length);
        return FormulaDecompiler.Decompile(expression, fixups, document, "{=", "}");
    }

    /// <summary>
    /// Measures a formula record at the start of <paramref name="tail"/>.
    /// </summary>
    internal static int MeasureRecord(ReadOnlySpan<byte> tail, out int fixupsOffset, out int rawTextOffset)
    {
        if (tail.Length < ExpressionOffset)
        {
            throw new ExcelFormatException("A formula record runs past the end of the file.");
        }

        int expressionLength = tail[14];
        int position = ExpressionOffset + expressionLength;
        if (position > tail.Length)
        {
            throw new ExcelFormatException("A formula's expression runs past the end of the file.");
        }

        rawTextOffset = -1;
        if (tail[0] == 0xFF && tail[1] == 0xFF && tail[2] == 0)
        {
            if (position >= tail.Length || position + 1 + tail[position] > tail.Length)
            {
                throw new ExcelFormatException("A formula's text result runs past the end of the file.");
            }

            rawTextOffset = position;
            position += 1 + tail[position];
        }

        fixupsOffset = position;
        return position + FormulaTokens.MeasureFixups(tail.Slice(ExpressionOffset, expressionLength), tail.Slice(position));
    }

    /// <summary>
    /// Measures the cell data of a cell of the given type at the start of <paramref name="tail"/>.
    /// </summary>
    internal static int MeasureCellData(ExcelCellType type, ReadOnlySpan<byte> tail)
    {
        int length = type switch
        {
            ExcelCellType.Number => 8,
            ExcelCellType.Text => tail.IsEmpty ? 1 : 1 + tail[0],
            ExcelCellType.NumericFormula or ExcelCellType.Formula => MeasureRecord(tail, out _, out _),
            _ => 0,
        };

        if (length > tail.Length)
        {
            throw new ExcelFormatException("Cell data runs past the end of the file.");
        }

        return length;
    }

    /// <summary>
    /// Parses a formula record at <paramref name="offset"/>, copying it into a single array.
    /// </summary>
    internal static ExcelFormula Read(ReadOnlySpan<byte> data, int offset)
    {
        if ((uint)offset >= (uint)data.Length)
        {
            throw new ExcelFormatException($"A formula record at file offset {offset} lies outside the file.");
        }

        var tail = data.Slice(offset);
        int length = MeasureRecord(tail, out int fixupsOffset, out int rawTextOffset);
        return new ExcelFormula(tail.Slice(0, length).ToArray(), fixupsOffset, rawTextOffset);
    }

    // The text of an array formula or table whose shared record cannot be read.
    internal const string ArrayFormulaPlaceholder = "{array formula}";
    internal const string TablePlaceholder = "{=TABLE()}";

    // Renumbers the name tokens after the document's names are reordered (map[old] = new, 1-based).
    internal void RemapNames(ReadOnlySpan<ushort> map)
        => FormulaTokens.RemapNames(Expression, _record.AsSpan(_fixupsOffset), map, definition: false);

    internal void Write(BigEndianWriter writer, int row, int column)
    {
        writer.WriteBytes(_record.AsSpan(0, 11));
        writer.WriteByte((byte)column);
        writer.WriteUInt16((ushort)((StoredRowWord & RowWord.Flags) | RowWord.Row((ushort)row)));
        writer.WriteBytes(_record.AsSpan(14, 1 + _record[14]));
        if (IsTextResult)
        {
            if (_rawTextOffset >= 0)
            {
                writer.WriteBytes(_record.AsSpan(_rawTextOffset, 1 + _record[_rawTextOffset]));
            }
            else
            {
                writer.WritePascalString(_text ?? string.Empty);
            }
        }

        writer.WriteBytes(Fixups);
    }
}
