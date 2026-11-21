namespace MacExcelReader.Tests;

/// <summary>
/// Helpers shared by the tests.
/// </summary>
internal static class TestData
{
    public static string SamplePath(string name) => Path.Combine(AppContext.BaseDirectory, "Samples", name);

    public static byte[] Sample(string name) => File.ReadAllBytes(SamplePath(name));

    public static ExcelDocument Load(string name) => ExcelDocument.Load(SamplePath(name));

    /// <summary>
    /// Writes a document and reads it back.
    /// </summary>
    public static ExcelDocument RoundTrip(ExcelDocument document) => ExcelDocument.Read(document.ToArray());

    /// <summary>
    /// Builds a relative cell reference token (class 0x44) for a row and column.
    /// </summary>
    public static byte[] Reference(int row, int column, bool rowRelative = true, bool columnRelative = true)
    {
        int word = row | (rowRelative ? 0x8000 : 0) | (columnRelative ? 0x4000 : 0);
        return [0x44, (byte)(word >> 8), (byte)word, (byte)column];
    }

    public static byte[] Integer(ushort value) => [0x1E, (byte)(value >> 8), (byte)value];

    public static byte[] Concat(params byte[][] parts) => parts.SelectMany(p => p).ToArray();
}
