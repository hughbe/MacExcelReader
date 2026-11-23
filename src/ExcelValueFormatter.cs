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
    /// switching to an exponent when the number does not fit written out. A negative number that rounds to nothing
    /// shows as <c>0</c>, without its sign (<c>19.0d</c>).
    /// </summary>
    public static string FormatGeneral(double value, int width)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(width);
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
                if (digits.Length == 0)
                {
                    // Rounded to nothing: Excel writes 0 over the minus sign (19.0d).
                    destination[0] = '0';
                    return 1;
                }

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
        return FormatValue(cell.Value, cell.NumberFormat, document);
    }

    /// <summary>
    /// Formats a cell's value for display, as <see cref="FormatCell(ExcelCell, ExcelDocument?)"/> does, without
    /// creating the cell.
    /// </summary>
    public static string FormatCell(ExcelCellRef cell, ExcelDocument? document = null)
        => FormatValue(cell.Value, cell.NumberFormat, document);
    /// <summary>
    /// Formats a cell's value as Excel shows it in its column (spec Appendix D): General fits the column, fills repeat
    /// across it, and a number, TRUE, FALSE or error too wide for it shows as a row of <c>#</c>. Text is returned
    /// whole, since Excel lets it run into the next cells (<see cref="LayoutCell"/>). Excel measures in pixels with
    /// the sheet's standard font; for a font or size missing from Excel 1.00's System, every character counts as one
    /// cached digit width (<see cref="ExcelSheetState.DigitWidth"/>), an approximation. Excel 1.00 and 1.5 draw
    /// alike; pass <see cref="ExcelVersion.Excel103"/> for Excel 1.03, which differs in font sizes the System has no
    /// bitmap for.
    /// </summary>
    public static string FormatCell(ExcelCell cell, ExcelDocument document, int column, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentNullException.ThrowIfNull(document);
        return Fit(cell.Value, cell.NumberFormat, cell.Bold, cell.Italic, document, column, version).Text;
    }

    /// <summary>
    /// Formats a cell's value as Excel shows it in its column, as
    /// <see cref="FormatCell(ExcelCell, ExcelDocument, int, ExcelVersion)"/> does, without creating the cell.
    /// </summary>
    public static string FormatCell(ExcelCellRef cell, ExcelDocument document, int column, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(document);
        return Fit(cell.Value, cell.NumberFormat, cell.Bold, cell.Italic, document, column, version).Text;
    }

    /// <summary>
    /// Works out how Excel draws a cell (spec Appendix D), as 1.00 and 1.5 do, or 1.03 with
    /// <see cref="ExcelVersion.Excel103"/>: the text it shows, where it starts and the stretch of the
    /// row it is clipped to, in pixels from the left edge of column A. Returns null for a cell with no value or in a
    /// hidden column.
    /// </summary>
    public static ExcelCellDisplay? LayoutCell(ExcelDocument document, int row, int column, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, ExcelCellGrid.MaxColumns);
        var cell = document.Cells.Find(row, column);
        var state = document.State;
        int width = state.GetColumnWidth(column);
        if (cell is null || cell.Type == ExcelCellType.Empty || width == 0)
        {
            return null;
        }

        int left = 0;
        for (int c = 0; c < column; c++)
        {
            left += state.GetColumnWidth(c);
        }

        int right = left + width;
        var value = cell.Value;
        var (text, measure, italicExtra, hashes) = Fit(value, cell.NumberFormat, cell.Bold, cell.Italic, document, column, version);
        var alignment = hashes ? ExcelAlignment.Center : cell.Alignment;
        if (alignment == ExcelAlignment.General)
        {
            // General puts text on the left, numbers on the right and TRUE, FALSE and errors in the middle (19.0e).
            alignment = value.Kind switch
            {
                ExcelValueKind.Text => ExcelAlignment.Left,
                ExcelValueKind.Number => ExcelAlignment.Right,
                _ => ExcelAlignment.Center,
            };
        }

        if (alignment == ExcelAlignment.Fill)
        {
            // A fill runs on over the following blank cells that are Fill aligned, and hidden columns (19.06), and
            // repeats its text as many whole times as fit, at least once (19.0b).
            int last = column;
            int runRight = right;
            for (int c = column + 1; c < ExcelCellGrid.MaxColumns; c++)
            {
                int w = state.GetColumnWidth(c);
                if (w != 0 && !(document.Cells.GetCellType(row, c) == ExcelCellType.Empty && document.Cells.GetAlignment(row, c) == ExcelAlignment.Fill))
                {
                    break;
                }

                last = c;
                runRight += w;
            }

            int once = measure.Width(text);
            int count = once == 0 ? 0 : Math.Max(1, (runRight - left - 3 - italicExtra) / once);
            var repeated = new System.Text.StringBuilder(text.Length * count);
            for (int i = 0; i < count; i++)
            {
                repeated.Append(text);
            }

            text = repeated.ToString();
            return new ExcelCellDisplay(text, alignment, left + 2, once * count + italicExtra, left, runRight - 1);
        }

        int textWidth = measure.Width(text) + italicExtra;
        int x = alignment switch
        {
            ExcelAlignment.Right => right - 2 - textWidth,
            ExcelAlignment.Center => (left + right - textWidth) >> 1,
            _ => left + 2,
        };

        int clipLeft = left;
        int clipRight = right - 1;
        if (value.Kind == ExcelValueKind.Text)
        {
            // Text runs into blank neighbours, on the side it is aligned away from, until a cell with a value
            // (19.07, 19.04, 19.05). Formatting alone does not stop it, and hidden columns let it through.
            if (alignment != ExcelAlignment.Right)
            {
                for (int c = column + 1; x + textWidth > clipRight && c < ExcelCellGrid.MaxColumns && document.Cells.GetCellType(row, c) == ExcelCellType.Empty; c++)
                {
                    clipRight += state.GetColumnWidth(c);
                }
            }

            if (alignment != ExcelAlignment.Left)
            {
                for (int c = column - 1; x < clipLeft && c >= 0 && document.Cells.GetCellType(row, c) == ExcelCellType.Empty; c--)
                {
                    clipLeft -= state.GetColumnWidth(c);
                }
            }
        }

        return new ExcelCellDisplay(text, alignment, x, textWidth, clipLeft, clipRight);
    }

    /// <summary>
    /// Gets the height of every row of a sheet, and how far below a row's top its text sits, in pixels (spec
    /// Appendix D): the standard font's ascent + descent + 1, with the baseline one ascent down. 13 and 10 in
    /// Geneva 10. For a font family Excel 1.00's System lacks, the cached values in the sheet state are used.
    /// </summary>
    public static (int Height, int Baseline) GetRowMetrics(ExcelDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var state = document.State;
        return MacFont.Find(state.StandardFont, state.StandardFontSize) is { } font
            ? (font.RowHeight, font.Ascent)
            : (state.CachedRowHeight, state.CachedAscent);
    }

    // A value as Excel 1.00 fits it into its column (19.0a, 19.0e): the text, how it was measured, the room italic
    // takes, and whether it became a row of #s.
    private static (string Text, TextMeasure Measure, int ItalicExtra, bool Hashes) Fit(
        ExcelValue value, int numberFormat, bool bold, bool italic, ExcelDocument document, int column, ExcelVersion version)
    {
        var state = document.State;
        var font = MacFont.Find(state.StandardFont, state.StandardFontSize);
        var measure = font is not null
            ? TextMeasure.Pixels(font, bold, version)
            : TextMeasure.Fixed(bold ? state.BoldDigitWidth : state.DigitWidth);
        int italicExtra = !italic ? 0 : font?.ItalicExtra ?? state.HalfAscent;

        // The room is the column less 2 pixels each side and the italic allowance.
        int room = state.GetColumnWidth(column) - 4 - italicExtra;
        switch (value.Kind)
        {
            case ExcelValueKind.Number:
            {
                double number = value.Number;
                var format = document.GetNumberFormat(numberFormat);
                string text;
                if (format is null || numberFormat == 0)
                {
                    // General: as many characters as there are digit widths, less a minus sign (19.0d).
                    int digits = (room - (number < 0 ? measure.Width('-') : 0)) / measure.DigitWidth;
                    text = FormatGeneral(number, Math.Max(0, digits));
                }
                else
                {
                    text = format.Format(number, room, measure);
                    if (text.Length > 0 && !text.AsSpan().ContainsAnyExcept('#'))
                    {
                        return (text, measure, italicExtra, true);
                    }
                }

                return measure.FitWidth(text) > room
                    ? (ExcelNumberFormat.Hashes(room, measure), measure, italicExtra, true)
                    : (text, measure, italicExtra, false);
            }

            case ExcelValueKind.Text:
                return (document.GetNumberFormat(numberFormat)?.Format(value.Text) ?? value.Text, measure, italicExtra, false);
            case ExcelValueKind.Empty:
                return (string.Empty, measure, italicExtra, false);
            default:
            {
                // TRUE, FALSE and errors show as they are, whatever the cell's format, or as #s (19.0e).
                string text = value.ToString();
                return measure.FitWidth(text) > room
                    ? (ExcelNumberFormat.Hashes(room, measure), measure, italicExtra, true)
                    : (text, measure, italicExtra, false);
            }
        }
    }

    private static string FormatValue(ExcelValue value, int numberFormat, ExcelDocument? document)
    {
        switch (value.Kind)
        {
            case ExcelValueKind.Number:
                return Format(value.Number, numberFormat, document);
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
