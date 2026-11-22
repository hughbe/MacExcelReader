using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// Formats values the way Excel 1.x displays them, through the tokens of their number format (spec §7). Dates use
/// the 1904 date system: serial 0 is 1 January 1904.
/// </summary>
public static class ExcelValueFormatter
{
    private static readonly DateTime s_epoch = new(1904, 1, 1);

    /// <summary>
    /// Gets whether a format index is a date format the way Excel decides it (<c>34.04</c> in 1.00; spec §7): one of
    /// the built-in date and time formats 10–18, or a user-defined format of <paramref name="document"/> that shows
    /// a date (<see cref="ExcelNumberFormat.IsDate"/>).
    /// </summary>
    public static bool IsDateFormat(int index, ExcelDocument? document = null)
        => index is >= 10 and <= 18
            || (index >= ExcelNumberFormat.BuiltInCount && document?.GetNumberFormat(index)?.IsDate == true);

    /// <summary>
    /// Converts a 1904-system serial number to a date and time.
    /// </summary>
    public static DateTime ToDateTime(double serial) => s_epoch.AddDays(serial);

    /// <summary>
    /// Converts a date and time to a 1904-system serial number.
    /// </summary>
    public static double ToSerial(DateTime value) => (value - s_epoch).TotalDays;

    /// <summary>
    /// Formats a number with the format a cell's format index refers to. User-defined formats (index 19 and up)
    /// need the document that defines them; without it, or past its last format, the number is shown as General.
    /// </summary>
    public static string Format(double value, int formatIndex, ExcelDocument? document = null)
    {
        var format = formatIndex < 0 ? null
            : document is not null ? document.GetNumberFormat(formatIndex)
            : formatIndex < ExcelNumberFormat.BuiltInCount ? ExcelNumberFormat.BuiltIn[formatIndex] : null;
        return format?.Format(value) ?? FormatGeneral(value);
    }

    /// <summary>
    /// Formats a number the way the General format does when Excel turns a number into text (as <c>""&amp;A1</c>
    /// does): rounded to 14 significant digits, written out in full from 1E-18 up to just below 1E+20 and with an
    /// exponent outside that range (spec §7). This is General with room for 20 characters.
    /// </summary>
    public static string FormatGeneral(double value) => FormatGeneral(value, 20);

    /// <summary>
    /// Formats a number the way the General format shows it in a column with room for <paramref name="width"/>
    /// characters, not counting a minus sign (spec §7, <c>19.0c</c> in Excel 1.00): as many digits as fit,
    /// switching to an exponent when the number does not fit written out.
    /// </summary>
    public static string FormatGeneral(double value, int width)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(width, 1);
        int size = GeneralSize(width);
        char[]? rented = size > 256 ? System.Buffers.ArrayPool<char>.Shared.Rent(size) : null;
        Span<char> buffer = rented ?? stackalloc char[256];
        try
        {
            return new string(buffer[..FormatGeneral(value, width, buffer)]);
        }
        finally
        {
            if (rented is not null)
            {
                System.Buffers.ArrayPool<char>.Shared.Return(rented);
            }
        }
    }

    // The most characters General can write with room for `width`: a sign, the digits written out with a point
    // and leading zeros, or a mantissa and an exponent.
    internal static int GeneralSize(int width) => width + 32;

    /// <summary>
    /// Formats a number as <see cref="FormatGeneral(double, int)"/> does into <paramref name="destination"/>, which
    /// must hold <see cref="GeneralSize"/> characters, and returns the number written.
    /// </summary>
    internal static int FormatGeneral(double value, int width, Span<char> destination)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            value.TryFormat(destination, out int special, default, System.Globalization.CultureInfo.InvariantCulture);
            return special;
        }

        int n = 0;
        if (value < 0)
        {
            destination[n++] = '-';
        }

        Span<char> buffer = stackalloc char[15];
        int length = Num2Str(Math.Abs(value), buffer, out int exponent);
        Span<char> digits = buffer[..length];

        // exponent counts the digits before the point: the value is 0.digits × 10^exponent.
        bool narrow = exponent is >= -3 and < 5 && width < 5;
        bool scientific;
        if (exponent > width)
        {
            scientific = !narrow;
        }
        else
        {
            int room = width + exponent - 2;
            scientific = room <= 0 || (exponent < -3 && room < digits.Length)
                ? width > 4 && !narrow
                : false;
        }

        if (!scientific)
        {
            int keep = exponent == 0 && width == 1 ? 0 : width;
            if (!(exponent == 0 && width == 1) && exponent < keep)
            {
                keep--;
                if (exponent <= 0)
                {
                    keep += exponent - 1;
                }
            }

            if (keep < digits.Length)
            {
                digits = buffer[..Round(buffer, digits.Length, ref exponent, keep)];
                if (exponent > width && !narrow)
                {
                    scientific = true;
                }
            }

            if (!scientific)
            {
                return n + Fixed(digits, exponent, destination[n..]);
            }
        }

        int mantissa = width <= 6 ? 1 : width - 6 + (exponent is > -100 and < 100 ? 1 : 0);
        if (mantissa < digits.Length)
        {
            digits = buffer[..Round(buffer, digits.Length, ref exponent, mantissa)];
        }

        if (digits.Length == 0)
        {
            buffer[0] = '0';
            digits = buffer[..1];
        }

        destination[n++] = digits[0];
        if (digits.Length > 1)
        {
            destination[n++] = '.';
            digits[1..].CopyTo(destination[n..]);
            n += digits.Length - 1;
        }

        int shown = exponent - 1;
        destination[n++] = 'E';
        destination[n++] = shown < 0 ? '-' : '+';
        Math.Abs(shown).TryFormat(destination[n..], out int written, Math.Abs(shown) >= 100 ? "000" : "00", System.Globalization.CultureInfo.InvariantCulture);
        return n + written;
    }

    // The digits written out: 1234, 12.34 or 0.001234.
    private static int Fixed(ReadOnlySpan<char> digits, int exponent, Span<char> destination)
    {
        if (digits.Length == 0)
        {
            destination[0] = '0';
            return 1;
        }

        if (exponent >= digits.Length)
        {
            digits.CopyTo(destination);
            destination[digits.Length..exponent].Fill('0');
            return exponent;
        }

        if (exponent > 0)
        {
            digits[..exponent].CopyTo(destination);
            destination[exponent] = '.';
            digits[exponent..].CopyTo(destination[(exponent + 1)..]);
            return digits.Length + 1;
        }

        destination[0] = '0';
        destination[1] = '.';
        destination.Slice(2, -exponent).Fill('0');
        digits.CopyTo(destination[(2 - exponent)..]);
        return 2 - exponent + digits.Length;
    }

    // Excel's decimal conversion: at most 14 significant digits, trailing zeros dropped, and the exponent that makes
    // the value 0.digits × 10^exponent. A tie at the 14th digit rounds down (99999999999999.5 gives 99999999999999).
    // Writes the digits to `result` and returns how many there are.
    private static int Num2Str(double magnitude, Span<char> result, out int exponent)
    {
        exponent = 1;
        if (magnitude == 0)
        {
            return 0;
        }

        // .NET gives the first 17 significant digits correctly rounded: d.dddd…E±x. They settle the 14 Excel keeps,
        // unless the 15th starts an apparent tie, 5 and then zeros, which the exact conversion has to decide. (A
        // value just below a tie rounds up to one at 17 digits, so it takes the exact path too.)
        Span<char> text = stackalloc char[32];
        magnitude.TryFormat(text, out int length, "E16", System.Globalization.CultureInfo.InvariantCulture);
        Span<char> digits = stackalloc char[17];
        digits[0] = text[0];
        text.Slice(2, 16).CopyTo(digits[1..]);
        if (digits[14] == '5' && !digits[15..].ContainsAnyExcept('0'))
        {
            var (exact, exactExponent) = Num2StrExact(magnitude);
            exact.CopyTo(result);
            exponent = exactExponent;
            return exact.Length;
        }

        exponent = int.Parse(text[19..length], System.Globalization.CultureInfo.InvariantCulture) + 1;
        long kept = long.Parse(digits[..14], System.Globalization.CultureInfo.InvariantCulture);
        if (digits[14] >= '5')
        {
            kept++;
            if (kept == 100_000_000_000_000)
            {
                kept = 10_000_000_000_000;
                exponent++;
            }
        }

        kept.TryFormat(result, out _, default, System.Globalization.CultureInfo.InvariantCulture);
        return result[..14].TrimEnd('0').Length;
    }

    // The same conversion done exactly, for a value whose 15th digit may be an exact tie.
    private static (string Digits, int Exponent) Num2StrExact(double magnitude)
    {
        long bits = BitConverter.DoubleToInt64Bits(magnitude);
        int binaryExponent = (int)((bits >> 52) & 0x7FF);
        long mantissa = bits & 0xFFFFFFFFFFFFFL;
        if (binaryExponent == 0)
        {
            binaryExponent = 1;
        }
        else
        {
            mantissa |= 1L << 52;
        }

        binaryExponent -= 1075;
        System.Numerics.BigInteger numerator = mantissa, denominator = 1;
        if (binaryExponent >= 0)
        {
            numerator <<= binaryExponent;
        }
        else
        {
            denominator <<= -binaryExponent;
        }

        // The exact decimal digits: scale so that the quotient has at least 15 digits.
        int scale = 0;
        while (numerator / denominator < System.Numerics.BigInteger.Pow(10, 14))
        {
            numerator *= 10;
            scale++;
        }

        while (numerator / denominator >= System.Numerics.BigInteger.Pow(10, 15))
        {
            denominator *= 10;
            scale--;
        }

        var quotient = System.Numerics.BigInteger.DivRem(numerator, denominator, out var remainder);
        string fifteen = quotient.ToString(System.Globalization.CultureInfo.InvariantCulture);
        int exponent = 15 - scale;
        var kept = System.Numerics.BigInteger.Parse(fifteen[..14], System.Globalization.CultureInfo.InvariantCulture);
        int next = fifteen[14] - '0';
        if (next > 5 || (next == 5 && !remainder.IsZero))
        {
            kept++;
        }

        string digits = kept.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (digits.Length > 14)
        {
            digits = digits[..14];
            exponent++;
        }

        return (digits.TrimEnd('0'), exponent);
    }

    // Rounds the digits in `digits` to `keep` significant digits, half up, as Excel's General does, and returns how
    // many are left once trailing zeros are dropped.
    private static int Round(Span<char> digits, int length, ref int exponent, int keep)
    {
        if (keep <= 0)
        {
            if (keep == 0 && length > 0 && digits[0] >= '5')
            {
                digits[0] = '1';
                exponent++;
                return 1;
            }

            exponent = 1;
            return 0;
        }

        if (digits[keep] >= '5')
        {
            int i = keep - 1;
            while (i >= 0 && digits[i] == '9')
            {
                digits[i--] = '0';
            }

            if (i < 0)
            {
                digits[0] = '1';
                exponent++;
                return 1;
            }

            digits[i]++;
        }

        return digits[..keep].TrimEnd('0').Length;
    }

    /// <summary>
    /// Formats a cell's value for display using its number format; TRUE, FALSE and errors show as they are. Pass the
    /// document to use its user-defined formats, including Excel 1.5's text sections.
    /// </summary>
    public static string FormatCell(ExcelCell cell, ExcelDocument? document = null)
    {
        ArgumentNullException.ThrowIfNull(cell);
        return FormatValue(cell.Value, cell.NumberFormat, document, null);
    }

    /// <summary>
    /// Formats a cell's value for display, as <see cref="FormatCell(ExcelCell, ExcelDocument?)"/> does, without
    /// creating the cell.
    /// </summary>
    public static string FormatCell(ExcelCellRef cell, ExcelDocument? document = null)
        => FormatValue(cell.Value, cell.NumberFormat, document, null);

    /// <summary>
    /// Formats a cell's value as Excel 1.00 shows it in its column (spec §7): General fits the column, fills repeat
    /// across it, and a number too wide for it shows as a row of <c>#</c>. Text is returned whole, since Excel lets
    /// it run into the next cells. Excel measures in pixels with the sheet's font; this counts the column's width
    /// in digit widths (the cached <see cref="ExcelSheetState.DigitWidth"/>, or the bold one for a bold cell), so it
    /// is an approximation.
    /// </summary>
    public static string FormatCell(ExcelCell cell, ExcelDocument document, int column)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentNullException.ThrowIfNull(document);
        return FormatValue(cell.Value, cell.NumberFormat, document, ColumnDigits(document, column, cell.Bold));
    }

    /// <summary>
    /// Formats a cell's value as Excel 1.00 shows it in its column, as
    /// <see cref="FormatCell(ExcelCell, ExcelDocument, int)"/> does, without creating the cell.
    /// </summary>
    public static string FormatCell(ExcelCellRef cell, ExcelDocument document, int column)
    {
        ArgumentNullException.ThrowIfNull(document);
        return FormatValue(cell.Value, cell.NumberFormat, document, ColumnDigits(document, column, cell.Bold));
    }

    private static int ColumnDigits(ExcelDocument document, int column, bool bold)
        => document.State.GetColumnWidth(column) / Math.Max(1, (int)(bold ? document.State.BoldDigitWidth : document.State.DigitWidth));

    private static string FormatValue(ExcelValue value, int numberFormat, ExcelDocument? document, int? width)
    {
        switch (value.Kind)
        {
            case ExcelValueKind.Number:
                if (width is not { } columns)
                {
                    return Format(value.Number, numberFormat, document);
                }

                var format = document?.GetNumberFormat(numberFormat) ?? ExcelNumberFormat.BuiltIn[0];
                return format.Format(value.Number, columns);
            case ExcelValueKind.Text:
                // Excel 1.5's text section, when the format has one.
                return document?.GetNumberFormat(numberFormat)?.Format(value.Text) ?? value.Text;
            case ExcelValueKind.Empty:
                return string.Empty;
            default:
                // TRUE, FALSE and errors show as they are, whatever the cell's format.
                return value.ToString();
        }
    }
}
