using System.Buffers.Binary;

namespace MacExcelReader;

/// <summary>
/// The kind of a formula's current value.
/// </summary>
public enum ExcelFormulaResultKind
{
    /// <summary>A number.</summary>
    Number,

    /// <summary>Text; the text follows the expression in the formula record.</summary>
    Text,

    /// <summary>A logical value.</summary>
    Boolean,

    /// <summary>An error value.</summary>
    Error,

    /// <summary>An unrecognised encoding.</summary>
    Unknown,
}

/// <summary>
/// A formula's current value, decoded from the first 8 bytes of the formula record (spec §8.4).
/// </summary>
public readonly struct ExcelFormulaResult
{
    internal ExcelFormulaResult(ReadOnlySpan<byte> valueBytes, string? text)
    {
        if (valueBytes[0] != 0xFF || valueBytes[1] != 0xFF)
        {
            Kind = ExcelFormulaResultKind.Number;
            Number = BinaryPrimitives.ReadDoubleBigEndian(valueBytes);
            return;
        }

        switch (valueBytes[2])
        {
            case 0:
                Kind = ExcelFormulaResultKind.Text;
                Text = text ?? string.Empty;
                break;
            case 1:
                Kind = ExcelFormulaResultKind.Boolean;
                Boolean = valueBytes[4] != 0;
                break;
            case 2:
                Kind = ExcelFormulaResultKind.Error;
                ErrorCode = valueBytes[4];
                break;
            default:
                Kind = ExcelFormulaResultKind.Unknown;
                break;
        }
    }

    /// <summary>Gets the kind of value.</summary>
    public ExcelFormulaResultKind Kind { get; }

    /// <summary>Gets the number, for <see cref="ExcelFormulaResultKind.Number"/>.</summary>
    public double Number { get; }

    /// <summary>Gets the text, for <see cref="ExcelFormulaResultKind.Text"/>.</summary>
    public string? Text { get; }

    /// <summary>Gets the logical value, for <see cref="ExcelFormulaResultKind.Boolean"/>.</summary>
    public bool Boolean { get; }

    /// <summary>Gets the error code, for <see cref="ExcelFormulaResultKind.Error"/>.</summary>
    public byte ErrorCode { get; }

    /// <summary>
    /// Gets the value as an object: a double, string, bool, or the error text such as "#VALUE!".
    /// </summary>
    public object? Value => Kind switch
    {
        ExcelFormulaResultKind.Number => Number,
        ExcelFormulaResultKind.Text => Text,
        ExcelFormulaResultKind.Boolean => Boolean,
        ExcelFormulaResultKind.Error => ErrorText(ErrorCode),
        _ => null,
    };

    /// <summary>
    /// Gets the display text of an error code (the BIFF-compatible codes).
    /// </summary>
    public static string ErrorText(byte code) => code switch
    {
        0 => "#NULL!",
        7 => "#DIV/0!",
        15 => "#VALUE!",
        23 => "#REF!",
        29 => "#NAME?",
        36 => "#NUM!",
        42 => "#N/A",
        _ => $"#ERR{code}",
    };

    /// <summary>
    /// Encodes a result as the 8 value bytes of a formula record.
    /// </summary>
    internal static void Encode(ExcelFormulaResultKind kind, double number, bool boolean, byte errorCode, Span<byte> bytes)
    {
        bytes.Slice(0, 8).Clear();
        switch (kind)
        {
            case ExcelFormulaResultKind.Number:
                BinaryPrimitives.WriteDoubleBigEndian(bytes, number);
                break;
            case ExcelFormulaResultKind.Text:
                bytes[0] = bytes[1] = 0xFF;
                break;
            case ExcelFormulaResultKind.Boolean:
                bytes[0] = bytes[1] = 0xFF;
                bytes[2] = 1;
                bytes[4] = (byte)(boolean ? 1 : 0);
                break;
            case ExcelFormulaResultKind.Error:
                bytes[0] = bytes[1] = 0xFF;
                bytes[2] = 2;
                bytes[4] = errorCode;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }
}
