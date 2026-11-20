using System.Globalization;
using System.Text;

namespace MacExcelReader.Charts;

/// <summary>
/// The 8-byte decimal number format used in chart files (spec §16.3): a sign bit, a decimal exponent
/// biased by 64, and 14 BCD digits.
/// </summary>
public readonly struct MicrosoftDecimal : IEquatable<MicrosoftDecimal>
{
    /// <summary>The size of a value in bytes.</summary>
    public const int Size = 8;

    private readonly ulong _bits;

    private MicrosoftDecimal(ulong bits) => _bits = bits;

    /// <summary>Gets zero.</summary>
    public static MicrosoftDecimal Zero => default;

    /// <summary>
    /// Creates a value from its 8 stored bytes.
    /// </summary>
    public static MicrosoftDecimal FromBytes(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < Size)
        {
            throw new ArgumentException("A decimal value is 8 bytes.", nameof(bytes));
        }

        ulong bits = 0;
        for (int i = 0; i < Size; i++)
        {
            bits = (bits << 8) | bytes[i];
        }

        return new MicrosoftDecimal(bits);
    }

    /// <summary>
    /// Writes the 8 stored bytes.
    /// </summary>
    public void WriteTo(Span<byte> destination)
    {
        for (int i = 0; i < Size; i++)
        {
            destination[i] = (byte)(_bits >> (56 - 8 * i));
        }
    }

    /// <summary>Gets the 8 stored bytes.</summary>
    public byte[] ToBytes()
    {
        var bytes = new byte[Size];
        WriteTo(bytes);
        return bytes;
    }

    /// <summary>
    /// Converts a double, rounding to 14 significant digits.
    /// </summary>
    public static MicrosoftDecimal FromDouble(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value), "Only finite values can be stored.");
        }

        if (value == 0)
        {
            return Zero;
        }

        bool negative = value < 0;
        string scientific = Math.Abs(value).ToString("E13", CultureInfo.InvariantCulture);
        int ePos = scientific.IndexOf('E');
        string digits = scientific[..ePos].Replace(".", string.Empty);
        int exponent = int.Parse(scientific[(ePos + 1)..], CultureInfo.InvariantCulture) + 1;
        if (exponent > 63)
        {
            exponent = 63;
            digits = new string('9', 14);
        }
        else if (exponent < -63)
        {
            exponent = -63;
            digits = "1" + new string('0', 13);
        }

        ulong bits = (ulong)((negative ? 0x80 : 0) | (exponent + 64)) << 56;
        for (int i = 0; i < 14; i++)
        {
            bits |= (ulong)(digits[i] - '0') << (52 - 4 * i);
        }

        return new MicrosoftDecimal(bits);
    }

    /// <summary>
    /// Converts the value to a double.
    /// </summary>
    public double ToDouble()
    {
        if (_bits == 0)
        {
            return 0;
        }

        int head = (int)(_bits >> 56);
        int exponent = (head & 0x7F) - 64;
        var digits = new StringBuilder(14);
        for (int i = 0; i < 14; i++)
        {
            int digit = (int)((_bits >> (52 - 4 * i)) & 0xF);
            digits.Append((char)('0' + Math.Min(digit, 9)));
        }

        double value = double.Parse($"0.{digits}E{exponent}", CultureInfo.InvariantCulture);
        return (head & 0x80) != 0 ? -value : value;
    }

    /// <inheritdoc/>
    public bool Equals(MicrosoftDecimal other) => _bits == other._bits;

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is MicrosoftDecimal other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => _bits.GetHashCode();

    /// <summary>Compares two values for equality.</summary>
    public static bool operator ==(MicrosoftDecimal left, MicrosoftDecimal right) => left.Equals(right);

    /// <summary>Compares two values for inequality.</summary>
    public static bool operator !=(MicrosoftDecimal left, MicrosoftDecimal right) => !left.Equals(right);

    /// <inheritdoc/>
    public override string ToString() => ToDouble().ToString("R", CultureInfo.InvariantCulture);
}
