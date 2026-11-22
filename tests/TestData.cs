namespace MacExcelReader.Tests;

/// <summary>
/// Helpers shared by the tests.
/// </summary>
internal static class TestData
{
    private static readonly string SamplesDirectory = Path.Combine(AppContext.BaseDirectory, "Samples");

    /// <summary>
    /// Finds a sample by its Mac file name, in whichever version folder holds it ("Excel 1.00", "Excel 1.03",
    /// "Excel 1.5"). A ".xls" or ".xlm" suffix is ignored, so "Expenses.xls" finds Expenses.
    /// </summary>
    public static string SamplePath(string name)
    {
        string bare = Path.GetExtension(name) is ".xls" or ".xlm" ? Path.GetFileNameWithoutExtension(name) : name;
        return Directory.GetFiles(SamplesDirectory, bare, SearchOption.AllDirectories).Single();
    }

    /// <summary>Gets every sample file, as "folder/name" relative to the samples directory.</summary>
    public static IEnumerable<string> AllSamples()
        => Directory.GetFiles(SamplesDirectory, "*", SearchOption.AllDirectories)
            .Where(p => !Path.GetFileName(p).StartsWith('.'))
            .Select(p => Path.GetRelativePath(SamplesDirectory, p))
            .Order(StringComparer.Ordinal);

    /// <summary>Gets every worksheet and macro sheet sample (magic 0x5327).</summary>
    public static TheoryData<string> AllSheets()
        => new(AllSamples().Where(p => File.ReadAllBytes(Path.Combine(SamplesDirectory, p))[0] == 0x53));

    /// <summary>Gets every chart sample (magic 0x018C).</summary>
    public static TheoryData<string> AllCharts()
        => new(AllSamples().Where(p => File.ReadAllBytes(Path.Combine(SamplesDirectory, p))[0] == 0x01));

    /// <summary>Gets the full path of a sample given as "folder/name".</summary>
    public static string SampleFile(string relative) => Path.Combine(SamplesDirectory, relative);

    /// <summary>
    /// Gets the Excel release whose encoding a sample follows, from the length of its embedded code segment,
    /// which identifies the build that saved it (spec §3.1).
    /// </summary>
    public static ExcelVersion SavedBy(byte[] file) => ((file[14] << 8) | file[15]) switch
    {
        // Excel 1.04's segment has the same length as 1.00's but imports its host as 0x740A (spec §2.3).
        0x326 when ImportWord(file) == 0x740A => ExcelVersion.Excel104,
        0x326 or 0x38A => ExcelVersion.Excel100,
        _ => ExcelVersion.Excel150,
    };

    // Segment 1 starts at offset 24: a word, a word, the procedure count, one offset per procedure, then the import.
    private static int ImportWord(byte[] file)
    {
        int procedures = (file[28] << 8) | file[29];
        int at = 24 + 6 + 2 * procedures;
        return (file[at] << 8) | file[at + 1];
    }

    /// <summary>
    /// Gets whether a sample was saved by the pre-release Excel 1.5 build (spec §2.3), recognised by a hash of its
    /// embedded code. That build wrote FORMULA with one argument in the variable-count form.
    /// </summary>
    public static bool IsPreRelease15(byte[] file)
    {
        int length = (file[14] << 8) | file[15];
        if (length != 0x32A)
        {
            return false;
        }

        var hash = System.Security.Cryptography.MD5.HashData(file.AsSpan(28, length - 14));
        return Convert.ToHexString(hash, 0, 4) == "63CE4783";
    }

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
