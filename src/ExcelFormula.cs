using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A formula record (spec §8.3): the current value, the parsed expression and the fix-ups that follow it.
/// </summary>
public sealed class ExcelFormula
{
    private byte[] _valueBytes;
    private string? _textResult;

    /// <summary>
    /// Creates a formula from a parsed expression.
    /// </summary>
    /// <param name="expression">The parsed expression (spec §9), at most 255 bytes.</param>
    /// <param name="fixups">
    /// The fix-up bytes that follow the expression (spec §8.5): one word per name token (its 1-based
    /// index in the name list) and per external-sheet token, in token order. Empty if none are needed.
    /// </param>
    public ExcelFormula(byte[] expression, byte[]? fixups = null)
    {
        ArgumentNullException.ThrowIfNull(expression);
        if (expression.Length > 255)
        {
            throw new ArgumentException("An expression cannot be longer than 255 bytes.", nameof(expression));
        }

        // Validate the token stream.
        _ = FormulaTokens.Enumerate(expression).Count();
        Expression = expression;
        Fixups = fixups ?? [];
        Reserved = new byte[3];
        _valueBytes = ExcelFormulaResult.Encode(ExcelFormulaResultKind.Number, 0, false, 0);
    }

    private ExcelFormula(byte[] valueBytes, byte[] reserved, byte column, ushort rowWord, byte[] expression, byte[] fixups, string? text)
    {
        _valueBytes = valueBytes;
        Reserved = reserved;
        StoredColumn = column;
        StoredRowWord = rowWord;
        Expression = expression;
        Fixups = fixups;
        _textResult = text;
    }

    /// <summary>
    /// Gets the parsed expression.
    /// </summary>
    public byte[] Expression { get; }

    /// <summary>
    /// Gets the raw fix-up bytes that follow the expression (and the text result, if any).
    /// </summary>
    public byte[] Fixups { get; }

    /// <summary>
    /// Gets bytes 8–10 of the record: two unidentified bytes and a recalculation scratch byte.
    /// </summary>
    public byte[] Reserved { get; }

    /// <summary>
    /// Gets the column stored in the record (the writer replaces it with the cell's column).
    /// </summary>
    public byte StoredColumn { get; internal set; }

    /// <summary>
    /// Gets the row word stored in the record: the row in bits 13–0 and two flags in bits 15–14.
    /// </summary>
    public ushort StoredRowWord { get; internal set; }

    /// <summary>
    /// Gets the raw 8 value bytes.
    /// </summary>
    public ReadOnlySpan<byte> RawValue => _valueBytes;

    /// <summary>
    /// Gets the formula's current value.
    /// </summary>
    public ExcelFormulaResult Result => new(_valueBytes, _textResult);

    /// <summary>Sets the current value to a number.</summary>
    public void SetResult(double value)
    {
        _valueBytes = ExcelFormulaResult.Encode(ExcelFormulaResultKind.Number, value, false, 0);
        _textResult = null;
    }

    /// <summary>Sets the current value to text.</summary>
    public void SetResult(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        _valueBytes = ExcelFormulaResult.Encode(ExcelFormulaResultKind.Text, 0, false, 0);
        _textResult = value;
    }

    /// <summary>Sets the current value to a logical value.</summary>
    public void SetResult(bool value)
    {
        _valueBytes = ExcelFormulaResult.Encode(ExcelFormulaResultKind.Boolean, 0, value, 0);
        _textResult = null;
    }

    /// <summary>Sets the current value to an error code.</summary>
    public void SetErrorResult(byte errorCode)
    {
        _valueBytes = ExcelFormulaResult.Encode(ExcelFormulaResultKind.Error, 0, false, errorCode);
        _textResult = null;
    }

    /// <summary>
    /// Gets the name indices and external-document indices from the fix-ups, in token order.
    /// </summary>
    public IReadOnlyList<ushort> GetFixupWords() => FormulaTokens.ReadFixupWords(Expression, Fixups);

    /// <summary>
    /// Converts the formula to A1-style text, such as <c>=SUM(A1:A5)</c>.
    /// </summary>
    /// <param name="document">The document, used to resolve names and external documents; may be null.</param>
    public string ToText(ExcelDocument? document = null)
        => "=" + FormulaDecompiler.Decompile(Expression, GetFixupWords(), document, fixupsInPlace: false);

    internal static ExcelFormula Read(BigEndianReader reader)
    {
        var value = reader.ReadBytes(8);
        var reserved = reader.ReadBytes(3);
        byte column = reader.ReadByte();
        ushort rowWord = reader.ReadUInt16();
        var expression = reader.ReadBytes(reader.ReadByte());
        string? text = null;
        if (value[0] == 0xFF && value[1] == 0xFF && value[2] == 0)
        {
            text = reader.ReadPascalString();
        }

        var fixups = FormulaTokens.ReadFixups(expression, reader);
        return new ExcelFormula(value, reserved, column, rowWord, expression, fixups, text);
    }

    internal void Write(BigEndianWriter writer, int row, int column)
    {
        writer.WriteBytes(_valueBytes);
        writer.WriteBytes(Reserved);
        writer.WriteByte((byte)column);
        writer.WriteUInt16((ushort)((StoredRowWord & 0xC000) | (row & 0x3FFF)));
        writer.WriteByte((byte)Expression.Length);
        writer.WriteBytes(Expression);
        if (Result.Kind == ExcelFormulaResultKind.Text)
        {
            writer.WritePascalString(_textResult ?? string.Empty);
        }

        writer.WriteBytes(Fixups);
    }
}
