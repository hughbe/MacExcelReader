using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The kind of an <see cref="ExcelValue"/>.
/// </summary>
public enum ExcelValueKind : byte
{
    /// <summary>No value: an empty cell.</summary>
    Empty,

    /// <summary>A number. Integer cells are numbers too.</summary>
    Number,

    /// <summary>Text.</summary>
    Text,

    /// <summary>TRUE or FALSE.</summary>
    Boolean,

    /// <summary>An error such as <c>#N/A</c>.</summary>
    Error,
}

/// <summary>
/// A cell's value or a formula's current value: a number, text, TRUE or FALSE, an error, or empty. A struct, so
/// reading a value does not allocate (text aside, which is a string).
/// </summary>
/// <remarks>
/// Numbers, text and logical values convert implicitly: <c>cell.SetValue(2.5)</c>, <c>formula.SetResult("total")</c>.
/// Errors do not, since a literal <c>0</c> would then be ambiguous: use <see cref="FromError"/>.
/// </remarks>
public readonly struct ExcelValue : IEquatable<ExcelValue>
{
    private readonly double _number;
    private readonly string? _text;

    private ExcelValue(ExcelValueKind kind, double number = 0, string? text = null)
    {
        Kind = kind;
        _number = number;
        _text = text;
    }

    /// <summary>Gets the empty value.</summary>
    public static ExcelValue Empty => default;

    /// <summary>Gets the kind of value.</summary>
    public ExcelValueKind Kind { get; }

    /// <summary>Gets whether this is the empty value.</summary>
    public bool IsEmpty => Kind == ExcelValueKind.Empty;

    /// <summary>Gets the number.</summary>
    /// <exception cref="InvalidOperationException">The value is not a number.</exception>
    public double Number => Kind == ExcelValueKind.Number ? _number : throw WrongKind(ExcelValueKind.Number);

    /// <summary>Gets the text.</summary>
    /// <exception cref="InvalidOperationException">The value is not text.</exception>
    public string Text => Kind == ExcelValueKind.Text ? _text! : throw WrongKind(ExcelValueKind.Text);

    /// <summary>Gets TRUE or FALSE.</summary>
    /// <exception cref="InvalidOperationException">The value is not TRUE or FALSE.</exception>
    public bool Boolean => Kind == ExcelValueKind.Boolean ? _number != 0 : throw WrongKind(ExcelValueKind.Boolean);

    /// <summary>Gets the error.</summary>
    /// <exception cref="InvalidOperationException">The value is not an error.</exception>
    public ExcelError Error => Kind == ExcelValueKind.Error ? (ExcelError)_number : throw WrongKind(ExcelValueKind.Error);

    /// <summary>Gets the number, if the value is one.</summary>
    public bool TryGetNumber(out double value)
    {
        value = Kind == ExcelValueKind.Number ? _number : 0;
        return Kind == ExcelValueKind.Number;
    }

    /// <summary>Gets the text, if the value is text.</summary>
    public bool TryGetText([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? value)
    {
        value = Kind == ExcelValueKind.Text ? _text : null;
        return value is not null;
    }

    /// <summary>Creates a number.</summary>
    public static ExcelValue FromNumber(double value) => new(ExcelValueKind.Number, value);

    /// <summary>Creates text, at most 255 Mac OS Roman characters when stored.</summary>
    public static ExcelValue FromText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new(ExcelValueKind.Text, text: value);
    }

    /// <summary>Creates TRUE or FALSE.</summary>
    public static ExcelValue FromBoolean(bool value) => new(ExcelValueKind.Boolean, value ? 1 : 0);

    /// <summary>Creates an error.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The error is not one of Excel's.</exception>
    public static ExcelValue FromError(ExcelError error)
    {
        if (!Enum.IsDefined(error))
        {
            throw new ArgumentOutOfRangeException(nameof(error), error, "Not an Excel error.");
        }

        return new(ExcelValueKind.Error, (byte)error);
    }

    // An error code read from a file, which may not be one of Excel's.
    internal static ExcelValue FromErrorCode(byte code) => new(ExcelValueKind.Error, code);

    /// <summary>Converts a number.</summary>
    public static implicit operator ExcelValue(double value) => FromNumber(value);

    /// <summary>Converts text.</summary>
    public static implicit operator ExcelValue(string value) => FromText(value);

    /// <summary>Converts TRUE or FALSE.</summary>
    public static implicit operator ExcelValue(bool value) => FromBoolean(value);

    /// <summary>
    /// Gets the value as Excel shows it with the General format: <c>12.5</c>, <c>TRUE</c>, <c>#N/A</c>, the text,
    /// or an empty string.
    /// </summary>
    public override string ToString() => Kind switch
    {
        ExcelValueKind.Number => ExcelValueFormatter.FormatGeneral(_number),
        ExcelValueKind.Text => _text!,
        ExcelValueKind.Boolean => BooleanText.Of(_number != 0),
        ExcelValueKind.Error => ((ExcelError)_number).ToText(),
        _ => string.Empty,
    };

    /// <inheritdoc/>
    public bool Equals(ExcelValue other)
        => Kind == other.Kind && _number.Equals(other._number) && string.Equals(_text, other._text, StringComparison.Ordinal);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is ExcelValue other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Kind, _number, _text);

    /// <summary>Compares two values.</summary>
    public static bool operator ==(ExcelValue left, ExcelValue right) => left.Equals(right);

    /// <summary>Compares two values.</summary>
    public static bool operator !=(ExcelValue left, ExcelValue right) => !left.Equals(right);

    private InvalidOperationException WrongKind(ExcelValueKind wanted)
        => new($"The value is {Kind}, not {wanted}.");

    /// <summary>
    /// Decodes a formula's current value from the first 8 bytes of its record (spec §12.4). A text result's text
    /// is passed in, since it follows the expression. An encoding Excel does not use reads as empty.
    /// </summary>
    internal static ExcelValue FromFormulaBytes(ReadOnlySpan<byte> valueBytes, string? text)
    {
        if (valueBytes[0] != 0xFF || valueBytes[1] != 0xFF)
        {
            return FromNumber(BinaryPrimitives.ReadDoubleBigEndian(valueBytes));
        }

        return valueBytes[2] switch
        {
            0 => new(ExcelValueKind.Text, text: text ?? string.Empty),
            1 => FromBoolean(valueBytes[4] != 0),
            2 => new(ExcelValueKind.Error, valueBytes[4]),
            _ => Empty,
        };
    }

    /// <summary>
    /// Encodes the value as the 8 value bytes of a formula record; a text result's text goes after the expression.
    /// </summary>
    /// <exception cref="ArgumentException">The value is empty, which a formula cannot have.</exception>
    internal void WriteFormulaBytes(Span<byte> bytes)
    {
        bytes.Slice(0, 8).Clear();
        switch (Kind)
        {
            case ExcelValueKind.Number:
                BinaryPrimitives.WriteDoubleBigEndian(bytes, _number);
                break;
            case ExcelValueKind.Text:
                bytes[0] = bytes[1] = 0xFF;
                break;
            case ExcelValueKind.Boolean:
                bytes[0] = bytes[1] = 0xFF;
                bytes[2] = 1;
                bytes[4] = (byte)(_number != 0 ? 1 : 0);
                break;
            case ExcelValueKind.Error:
                bytes[0] = bytes[1] = 0xFF;
                bytes[2] = 2;
                bytes[4] = (byte)_number;
                break;
            default:
                throw new ArgumentException("A formula's value cannot be empty.");
        }
    }
}
