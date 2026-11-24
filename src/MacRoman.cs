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
    /// Decodes Mac OS Roman bytes into <paramref name="chars"/>, one character per byte, without allocating.
    /// </summary>
    /// <returns>The number of characters written: the number of bytes.</returns>
    /// <exception cref="ArgumentException"><paramref name="chars"/> is shorter than <paramref name="bytes"/>.</exception>
    public static int GetChars(ReadOnlySpan<byte> bytes, Span<char> chars)
    {
        if (chars.Length < bytes.Length)
        {
            throw new ArgumentException("The destination is too short: Mac OS Roman has one character per byte.", nameof(chars));
        }

        for (int i = 0; i < bytes.Length; i++)
        {
            chars[i] = GetChar(bytes[i]);
        }

        return bytes.Length;
    }

    /// <summary>
    /// Encodes a string as Mac OS Roman. Characters with no Mac OS Roman equivalent become '?'.
    /// </summary>
    public static byte[] GetBytes(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        var result = new byte[value.Length];
        GetBytes(value, result);
        return result;
    }

    /// <summary>
    /// Encodes characters as Mac OS Roman into <paramref name="bytes"/>, one byte per character. Characters with no
    /// Mac OS Roman equivalent become '?'.
    /// </summary>
    /// <returns>The number of bytes written: the number of characters.</returns>
    /// <exception cref="ArgumentException"><paramref name="bytes"/> is shorter than <paramref name="chars"/>.</exception>
    public static int GetBytes(ReadOnlySpan<char> chars, Span<byte> bytes)
    {
        if (bytes.Length < chars.Length)
        {
            throw new ArgumentException("The destination is too short: Mac OS Roman needs one byte per character.", nameof(bytes));
        }

        // Most text is plain ASCII, which encodes with a vectorised fast path.
        if (Ascii.IsValid(chars))
        {
            return Encoding.ASCII.GetBytes(chars, bytes);
        }

        for (int i = 0; i < chars.Length; i++)
        {
            bytes[i] = GetByte(chars[i]);
        }

        return chars.Length;
    }

    /// <summary>Decodes one Mac OS Roman byte.</summary>
    internal static char GetChar(byte b) => b < 0x80 ? (char)b : HighHalf[b - 0x80];

    /// <summary>Encodes one character as Mac OS Roman, or '?' when it has no equivalent.</summary>
    internal static byte GetByte(char c)
        => c < 0x80 ? (byte)c : s_reverse.TryGetValue(c, out byte b) ? b : (byte)'?';
}
