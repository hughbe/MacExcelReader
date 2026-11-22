using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// Formats values the way Excel 1.x displays them, through the tokens of their number format (spec §7). Dates use
/// the 1904 date system: serial 0 is 1 January 1904.
/// </summary>
public static class ExcelValueFormatter
{
    /// <summary>The number of built-in formats.</summary>
    public const int BuiltInFormatCount = 19;

    private static readonly DateTime s_epoch = new(1904, 1, 1);

    /// <summary>
    /// Gets the text of a built-in format, such as "#,##0.00", or null for a user-defined index.
    /// </summary>
    public static string? GetBuiltInFormat(int index)
        => index >= 0 && index < BuiltInFormatCount ? ExcelNumberFormat.BuiltIn[index].ToText() : null;

    /// <summary>
    /// Gets whether a format index is a date format the way Excel decides it (<c>34.04</c> in 1.00; spec §7): one of
    /// the built-in date and time formats 10–18, or a user-defined format of <paramref name="document"/> that shows
    /// a date (<see cref="ExcelNumberFormat.IsDate"/>).
    /// </summary>
    public static bool IsDateFormat(int index, ExcelDocument? document = null)
        => index is >= 10 and <= 18
            || (index >= BuiltInFormatCount && document?.GetNumberFormat(index)?.IsDate == true);

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
            : formatIndex < BuiltInFormatCount ? ExcelNumberFormat.BuiltIn[formatIndex] : null;
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
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        string sign = value < 0 ? "-" : string.Empty;
        var (digits, exponent) = Num2Str(Math.Abs(value));

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
                (digits, exponent) = Round(digits, exponent, keep);
                if (exponent > width && !narrow)
                {
                    scientific = true;
                }
            }

            if (!scientific)
            {
                return sign + Fixed(digits, exponent);
            }
        }

        int mantissa = width <= 6 ? 1 : width - 6 + (exponent is > -100 and < 100 ? 1 : 0);
        if (mantissa < digits.Length)
        {
            (digits, exponent) = Round(digits, exponent, mantissa);
        }

        if (digits.Length == 0)
        {
            digits = "0";
        }

        int shown = exponent - 1;
        string body = digits.Length > 1 ? digits[0] + "." + digits[1..] : digits;
        return $"{sign}{body}E{(shown < 0 ? '-' : '+')}{Math.Abs(shown).ToString(Math.Abs(shown) >= 100 ? "000" : "00", System.Globalization.CultureInfo.InvariantCulture)}";
    }

    private static string Fixed(string digits, int exponent)
    {
        if (digits.Length == 0)
        {
            return "0";
        }

        if (exponent >= digits.Length)
        {
            return digits + new string('0', exponent - digits.Length);
        }

        return exponent > 0
            ? digits[..exponent] + "." + digits[exponent..]
            : "0." + new string('0', -exponent) + digits;
    }

    // Excel's decimal conversion: at most 14 significant digits, trailing zeros dropped, and the exponent that makes
    // the value 0.digits × 10^exponent. A tie at the 14th digit rounds down (99999999999999.5 gives 99999999999999).
    private static (string Digits, int Exponent) Num2Str(double magnitude)
    {
        if (magnitude == 0)
        {
            return (string.Empty, 1);
        }

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

    // Rounds a digit string to the given number of significant digits, half up, as Excel's General does.
    private static (string Digits, int Exponent) Round(string digits, int exponent, int keep)
    {
        if (keep <= 0)
        {
            return keep == 0 && digits.Length > 0 && digits[0] >= '5' ? ("1", exponent + 1) : (string.Empty, 1);
        }

        var chars = digits[..keep].ToCharArray();
        if (digits[keep] >= '5')
        {
            int i = keep - 1;
            while (i >= 0 && chars[i] == '9')
            {
                chars[i--] = '0';
            }

            if (i < 0)
            {
                return ("1", exponent + 1);
            }

            chars[i]++;
        }

        return (new string(chars).TrimEnd('0'), exponent);
    }

    /// <summary>
    /// Formats a cell's value for display using its number format; logical values show as TRUE/FALSE and errors as
    /// their error text. Pass the document to use its user-defined formats, including Excel 1.5's text sections.
    /// </summary>
    public static string FormatCell(ExcelCell cell, ExcelDocument? document = null)
    {
        return FormatCellCore(cell, document, null);
    }

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
        int digit = cell.Bold ? document.State.BoldDigitWidth : document.State.DigitWidth;
        return FormatCellCore(cell, document, document.State.GetColumnWidth(column) / Math.Max(1, digit));
    }

    private static string FormatCellCore(ExcelCell cell, ExcelDocument? document, int? width)
    {
        ArgumentNullException.ThrowIfNull(cell);
        if (cell.TryGetNumber(out double number))
        {
            if (width is not { } columns)
            {
                return Format(number, cell.NumberFormat, document);
            }

            var format = document?.GetNumberFormat(cell.NumberFormat) ?? ExcelNumberFormat.BuiltIn[0];
            return format.Format(number, columns);
        }

        // TRUE, FALSE and errors show as they are, whatever the cell's format.
        if (cell.Type == ExcelCellType.BooleanOrError)
        {
            return cell.IsError ? cell.Error.ToText() : BooleanText.Of(cell.BooleanValue);
        }

        if (cell.Formula?.Result is { Kind: ExcelFormulaResultKind.Error } result)
        {
            return result.Error.ToText();
        }

        return cell.Value switch
        {
            bool b => BooleanText.Of(b),
            string s => document?.GetNumberFormat(cell.NumberFormat)?.Format(s) ?? s,
            _ => string.Empty,
        };
    }
}
