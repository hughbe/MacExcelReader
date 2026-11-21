using System.Globalization;

namespace MacExcelReader;

/// <summary>
/// Formats values the way Excel 1.x displays them with its 19 built-in number formats (spec §6.1).
/// Dates use the 1904 date system: serial 0 is 1 January 1904.
/// </summary>
public static class ExcelValueFormatter
{
    /// <summary>The number of built-in formats.</summary>
    public const int BuiltInFormatCount = 19;

    private static readonly DateTime s_epoch = new(1904, 1, 1);

    private static readonly string[] s_formatNames =
    [
        "General", "0", "0.00", "#,##0", "#,##0.00", "$#,##0 ;($#,##0)", "$#,##0.00 ;($#,##0.00)", "0%", "0.00%",
        "0.00E+00", "m/d/yy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss",
        "m/d/yy h:mm",
    ];

    /// <summary>
    /// Gets the text of a built-in format, such as "#,##0.00", or null for a user-defined index.
    /// </summary>
    public static string? GetBuiltInFormat(int index)
        => index >= 0 && index < s_formatNames.Length ? s_formatNames[index] : null;

    /// <summary>
    /// Gets whether a built-in format displays a date or time.
    /// </summary>
    public static bool IsDateFormat(int index) => index is >= 10 and <= 18;

    /// <summary>
    /// Converts a 1904-system serial number to a date and time.
    /// </summary>
    public static DateTime ToDateTime(double serial) => s_epoch.AddDays(serial);

    /// <summary>
    /// Converts a date and time to a 1904-system serial number.
    /// </summary>
    public static double ToSerial(DateTime value) => (value - s_epoch).TotalDays;

    /// <summary>
    /// Formats a number with a built-in format. User-defined formats (index 19 and up) use General.
    /// </summary>
    public static string Format(double value, int formatIndex)
    {
        var invariant = CultureInfo.InvariantCulture;
        switch (formatIndex)
        {
            case 1:
                return Math.Round(value, MidpointRounding.AwayFromZero).ToString("0", invariant);
            case 2:
                return value.ToString("0.00", invariant);
            case 3:
                return value.ToString("#,##0", invariant);
            case 4:
                return value.ToString("#,##0.00", invariant);
            case 5:
                return value < 0 ? "($" + (-value).ToString("#,##0", invariant) + ")" : "$" + value.ToString("#,##0", invariant) + " ";
            case 6:
                return value < 0 ? "($" + (-value).ToString("#,##0.00", invariant) + ")" : "$" + value.ToString("#,##0.00", invariant) + " ";
            case 7:
                return (value * 100).ToString("0", invariant) + "%";
            case 8:
                return (value * 100).ToString("0.00", invariant) + "%";
            case 9:
                return value.ToString("0.00E+00", invariant);
            case >= 10 and <= 18:
                return FormatDate(value, formatIndex);
            default:
                return FormatGeneral(value);
        }
    }

    /// <summary>
    /// Formats a number the way the General format does: up to 11 significant characters.
    /// </summary>
    public static string FormatGeneral(double value)
    {
        var invariant = CultureInfo.InvariantCulture;
        string text = value.ToString("G11", invariant);
        if (text.Contains('E'))
        {
            text = value.ToString("0.#####E+00", invariant);
        }

        return text;
    }

    private static string FormatDate(double serial, int formatIndex)
    {
        if (serial < 0 || serial > 2_958_465)
        {
            return FormatGeneral(serial);
        }

        var date = ToDateTime(serial);
        var invariant = CultureInfo.InvariantCulture;
        return formatIndex switch
        {
            10 => date.ToString("M/d/yy", invariant),
            11 => date.ToString("d-MMM-yy", invariant),
            12 => date.ToString("d-MMM", invariant),
            13 => date.ToString("MMM-yy", invariant),
            14 => date.ToString("h:mm tt", invariant),
            15 => date.ToString("h:mm:ss tt", invariant),
            16 => date.ToString("H:mm", invariant),
            17 => date.ToString("H:mm:ss", invariant),
            _ => date.ToString("M/d/yy H:mm", invariant),
        };
    }

    /// <summary>
    /// Formats a cell's value for display using its number format; text is returned as is, logical values
    /// as TRUE/FALSE and errors as their error text.
    /// </summary>
    public static string FormatCell(ExcelCell cell)
    {
        ArgumentNullException.ThrowIfNull(cell);
        if (cell.TryGetNumber(out double number))
        {
            return Format(number, cell.NumberFormat);
        }

        return cell.Value switch
        {
            null => string.Empty,
            bool b => b ? "TRUE" : "FALSE",
            string s => s,
            var other => Convert.ToString(other, CultureInfo.InvariantCulture) ?? string.Empty,
        };
    }
}
