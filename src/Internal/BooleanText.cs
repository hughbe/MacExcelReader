namespace MacExcelReader.Internal;

/// <summary>
/// The text of Excel's logical values.
/// </summary>
internal static class BooleanText
{
    public const string True = "TRUE";
    public const string False = "FALSE";

    public static string Of(bool value) => value ? True : False;

    /// <summary>Parses <c>TRUE</c> or <c>FALSE</c>, ignoring case.</summary>
    public static bool TryParse(string text, out bool value) => TryParse(text.AsSpan(), out value);

    /// <summary>Parses <c>TRUE</c> or <c>FALSE</c>, ignoring case.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out bool value)
    {
        value = text.Equals(True, StringComparison.OrdinalIgnoreCase);
        return value || text.Equals(False, StringComparison.OrdinalIgnoreCase);
    }
}
