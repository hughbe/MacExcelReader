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
    /// Gets whether a built-in format displays a date or time.
    /// </summary>
    public static bool IsDateFormat(int index) => index is >= 0 and < BuiltInFormatCount && ExcelNumberFormat.BuiltIn[index].IsDate;

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
    /// Formats a number the way the General format does: up to 11 significant characters.
    /// </summary>
    public static string FormatGeneral(double value)
    {
        var invariant = System.Globalization.CultureInfo.InvariantCulture;
        string text = value.ToString("G11", invariant);
        if (text.Contains('E'))
        {
            text = value.ToString("0.#####E+00", invariant);
        }

        return text;
    }

    /// <summary>
    /// Formats a cell's value for display using its number format; logical values show as TRUE/FALSE and errors as
    /// their error text. Pass the document to use its user-defined formats, including Excel 1.5's text sections.
    /// </summary>
    public static string FormatCell(ExcelCell cell, ExcelDocument? document = null)
    {
        ArgumentNullException.ThrowIfNull(cell);
        if (cell.TryGetNumber(out double number))
        {
            return Format(number, cell.NumberFormat, document);
        }

        return cell.Value switch
        {
            bool b => b ? "TRUE" : "FALSE",
            string s => document?.GetNumberFormat(cell.NumberFormat)?.Format(s) ?? s,
            _ => string.Empty,
        };
    }
}
