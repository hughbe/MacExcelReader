namespace MacExcelReader.Internal;

/// <summary>
/// How Excel 1.00 compares defined names (spec §11): letter by letter, ignoring case and accents, with Æ and æ as AE,
/// Œ and œ as OE, and Ø and ø as O. Names with the same key are the same name, and Excel keeps its names in the order
/// of their keys. Other characters (digits, '.', '_') compare by their codes.
/// </summary>
internal static class NameKey
{
    /// <summary>Gets the key Excel orders and matches a name by.</summary>
    public static string Of(ReadOnlySpan<char> name)
    {
        Span<char> buffer = name.Length <= 128 ? stackalloc char[name.Length * 2] : new char[name.Length * 2];
        int length = 0;
        foreach (char c in name)
        {
            switch (c)
            {
                case 'Æ' or 'æ':
                    buffer[length++] = 'A';
                    buffer[length++] = 'E';
                    break;
                case 'Œ' or 'œ':
                    buffer[length++] = 'O';
                    buffer[length++] = 'E';
                    break;
                case 'Ø' or 'ø':
                    buffer[length++] = 'O';
                    break;
                default:
                    buffer[length++] = char.ToUpperInvariant(BaseLetter(c));
                    break;
            }
        }

        return new string(buffer[..length]);
    }

    /// <summary>Gets whether two names are the same name to Excel.</summary>
    public static bool Equal(ReadOnlySpan<char> a, ReadOnlySpan<char> b) => a.Equals(b, StringComparison.OrdinalIgnoreCase) || Of(a) == Of(b);

    // A letter with its accent removed; anything else as it is.
    private static char BaseLetter(char c)
    {
        if (c < 0x80)
        {
            return c;
        }

        var decomposed = c.ToString().Normalize(System.Text.NormalizationForm.FormD);
        return decomposed.Length > 0 && decomposed[0] < 0x80 && char.IsLetter(decomposed[0]) ? decomposed[0] : c;
    }
}
