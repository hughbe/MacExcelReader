namespace MacExcelReader.Internal;

/// <summary>
/// Checks text stored as a Pascal string: at most 255 Mac OS Roman bytes, one per character.
/// </summary>
internal static class PascalText
{
    public const int MaxLength = 255;

    public static string Check(string value, string parameterName)
    {
        ArgumentNullException.ThrowIfNull(value, parameterName);
        if (value.Length > MaxLength)
        {
            throw new ArgumentException($"The text is {value.Length} characters long; a file stores at most {MaxLength}.", parameterName);
        }

        return value;
    }

    public static string? CheckOptional(string? value, string parameterName) => value is null ? null : Check(value, parameterName);
}
