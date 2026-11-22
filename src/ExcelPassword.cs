namespace MacExcelReader;

/// <summary>
/// The document-protection password hash (spec §6.2).
/// </summary>
public static class ExcelPassword
{
    /// <summary>
    /// Computes the 16-bit hash Excel stores for a protection password. An empty password hashes to 0.
    /// </summary>
    public static ushort Hash(string password)
    {
        ArgumentNullException.ThrowIfNull(password);
        var bytes = MacRoman.GetBytes(password);
        if (bytes.Length == 0)
        {
            return 0;
        }

        int hash = 0;
        for (int i = bytes.Length; i >= 0; i--)
        {
            int value = i == 0 ? bytes.Length : bytes[i - 1];
            hash = ((hash << 1) & 0x7FFF) | ((hash >> 14) & 1);
            hash ^= value;
        }

        return (ushort)(hash ^ 0xCE4B);
    }
}
