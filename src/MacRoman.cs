using System.Text;

namespace MacExcelReader;

/// <summary>
/// Converts between .NET strings and the Mac OS Roman character set used for all text in Excel 1.x files.
/// </summary>
public static class MacRoman
{
    // Characters 0x80–0xFF of Mac OS Roman.
    private const string HighHalf =
        "ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü" +
        "†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø" +
        "¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ" +
        "‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ";

    private static readonly Dictionary<char, byte> s_reverse = BuildReverse();

    private static Dictionary<char, byte> BuildReverse()
    {
        var map = new Dictionary<char, byte>(128);
        for (int i = 0; i < HighHalf.Length; i++)
        {
            map[HighHalf[i]] = (byte)(0x80 + i);
        }

        return map;
    }

    /// <summary>
    /// Decodes Mac OS Roman bytes.
    /// </summary>
    public static string GetString(ReadOnlySpan<byte> bytes)
    {
        if (bytes.IsEmpty)
        {
            return string.Empty;
        }

        // Most text is plain ASCII, which decodes with a vectorised fast path.
        if (Ascii.IsValid(bytes))
        {
            return Encoding.ASCII.GetString(bytes);
        }

        Span<char> chars = bytes.Length <= 256 ? stackalloc char[bytes.Length] : new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++)
        {
            byte b = bytes[i];
            chars[i] = b < 0x80 ? (char)b : HighHalf[b - 0x80];
        }

        return new string(chars);
    }

    /// <summary>
    /// Encodes a string as Mac OS Roman. Characters with no Mac OS Roman equivalent become '?'.
    /// </summary>
    public static byte[] GetBytes(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        var result = new byte[value.Length];
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (c < 0x80)
            {
                result[i] = (byte)c;
            }
            else if (s_reverse.TryGetValue(c, out var b))
            {
                result[i] = b;
            }
            else
            {
                result[i] = (byte)'?';
            }
        }

        return result;
    }
}
