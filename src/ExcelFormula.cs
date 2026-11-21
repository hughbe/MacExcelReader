using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A formula record (spec §8.3): the current value, the parsed expression and the fix-ups that follow it.
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
    /// <param name="expression">The parsed expression (spec §9), at most 255 bytes.</param>
    /// <param name="fixups">
    /// The fix-up bytes that follow the expression (spec §8.5): one word per name token (its 1-based
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

    /// <summary>Gets bytes 8–10 of the record: two unidentified bytes and a recalculation scratch byte.</summary>
    public ReadOnlySpan<byte> Reserved => _record.AsSpan(8, 3);

    /// <summary>Gets the column stored in the record (the writer replaces it with the cell's column).</summary>
    public byte StoredColumn => _record[11];

    /// <summary>Gets the row word stored in the record: the row in bits 13–0 and two flags in bits 15–14.</summary>
    public ushort StoredRowWord => BinaryPrimitives.ReadUInt16BigEndian(_record.AsSpan(12));

    /// <summary>Gets the raw 8 value bytes.</summary>
    public ReadOnlySpan<byte> RawValue => _record.AsSpan(0, 8);

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

    /// <summary>Gets the formula's current value.</summary>
    public ExcelFormulaResult Result => new(RawValue, IsTextResult ? TextResult : null);

    /// <summary>Sets the current value to a number.</summary>
    public void SetResult(double value) => SetValue(ExcelFormulaResultKind.Number, value, false, 0, null);

    /// <summary>Sets the current value to text.</summary>
    public void SetResult(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        SetValue(ExcelFormulaResultKind.Text, 0, false, 0, value);
    }

    /// <summary>Sets the current value to a logical value.</summary>
    public void SetResult(bool value) => SetValue(ExcelFormulaResultKind.Boolean, 0, value, 0, null);

    /// <summary>Sets the current value to an error code.</summary>
    public void SetErrorResult(byte errorCode) => SetValue(ExcelFormulaResultKind.Error, 0, false, errorCode, null);

    private void SetValue(ExcelFormulaResultKind kind, double number, bool boolean, byte errorCode, string? text)
    {
        ExcelFormulaResult.Encode(kind, number, boolean, errorCode, _record.AsSpan(0, 8));
        _rawTextOffset = -1;
        _text = text;
    }

    /// <summary>Gets the name indices and external-document indices from the fix-ups, in token order.</summary>
    public IReadOnlyList<ushort> GetFixupWords() => FormulaTokens.ReadFixupWords(Expression, Fixups);

    /// <summary>
    /// Converts the formula to A1-style text, such as <c>=SUM(A1:A5)</c>.
    /// </summary>
    /// <param name="document">The document, used to resolve names and external documents; may be null.</param>
    public string ToText(ExcelDocument? document = null)
        => "=" + FormulaDecompiler.Decompile(Expression, GetFixupWords(), document, fixupsInPlace: false);

    /// <summary>
    /// Measures a formula record at the start of <paramref name="tail"/>.
    /// </summary>
    private static int MeasureRecord(ReadOnlySpan<byte> tail, out int fixupsOffset, out int rawTextOffset)
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

    internal void Write(BigEndianWriter writer, int row, int column)
    {
        writer.WriteBytes(_record.AsSpan(0, 11));
        writer.WriteByte((byte)column);
        writer.WriteUInt16((ushort)((StoredRowWord & 0xC000) | (row & 0x3FFF)));
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
