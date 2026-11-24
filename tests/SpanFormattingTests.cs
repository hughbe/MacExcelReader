using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// The span-based formatting methods give the same text as the string ones, report a destination too small, and
/// allocate nothing on the common paths.
/// </summary>
public class SpanFormattingTests
{
    public static TheoryData<string> Samples => AllSheets();

    [Theory]
    [MemberData(nameof(Samples))]
    public void TryFormat_MatchesTheStringMethods(string name)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(name)));
        var buffer = new char[1024];
        foreach (var entry in document.Cells.NonEmptyCells)
        {
            Assert.True(ExcelValueFormatter.TryFormatCell(entry, document, buffer, out int written));
            Assert.Equal(ExcelValueFormatter.FormatCell(entry, document), new string(buffer, 0, written));

            Assert.True(entry.Value.TryFormat(buffer, out written));
            Assert.Equal(entry.Value.ToString(), new string(buffer, 0, written));

            if (entry.Type is ExcelCellType.Formula or ExcelCellType.NumericFormula)
            {
                var formula = entry.GetCell().Formula!;
                Assert.True(formula.TryFormatText(buffer, out written, document));
                Assert.Equal(formula.ToText(document), new string(buffer, 0, written));
                Assert.Equal(formula.FixupWords.Count, formula.FixupWords.ToArray().Length);
            }

            Assert.True(ExcelValueFormatter.TryFormatCell(entry.GetCell(), document, buffer, out written));
            Assert.Equal(ExcelValueFormatter.FormatCell(entry.GetCell(), document), new string(buffer, 0, written));
        }
    }

    [Fact]
    public void TryFormat_ReportsADestinationTooSmall()
    {
        Span<char> small = stackalloc char[3];
        Assert.False(ExcelValueFormatter.TryFormat(1234.5, 4, null, small, out int written));
        Assert.Equal(0, written);
        Assert.False(ExcelValue.FromText("long text").TryFormat(small, out _));
        Assert.False(ExcelFormula.Parse("=SUM(A1:A5)").TryFormatText(small, out _));
        Assert.True(ExcelValueFormatter.TryFormat(12, 0, null, small, out written));
        Assert.Equal("12", small[..written].ToString());
    }

    private static readonly int[] Formats = [0, 2, 4, 7, 9, 12, 14, 16, 18];

    [Fact]
    public void TryFormat_DoesNotAllocate()
    {
        var document = Load("DB.FORM");
        var buffer = new char[512];
        long total = 0;
        void Run()
        {
            foreach (var entry in document.Cells.NonEmptyCells)
            {
                ExcelValueFormatter.TryFormatCell(entry, document, buffer, out int written);
                total += written;
            }

            foreach (int format in Formats)
            {
                ExcelValueFormatter.TryFormat(40123.456, format, null, buffer, out int written);
                total += written;
            }
        }

        Run();
        long before = GC.GetAllocatedBytesForCurrentThread();
        Run();
        Assert.Equal(0, GC.GetAllocatedBytesForCurrentThread() - before);
        Assert.True(total > 0);
    }

    [Theory]
    [InlineData(0, 0, "A1")]
    [InlineData(18, 1, "B19")]
    [InlineData(16383, 255, "IV16384")]
    [InlineData(16384, 256, "IW16385")]
    public void Address_TryFormatMatchesToString(int row, int column, string expected)
    {
        var address = new ExcelCellAddress(row, column);
        Span<char> buffer = stackalloc char[16];
        Assert.True(address.TryFormat(buffer, out int written));
        Assert.Equal(expected, buffer[..written].ToString());
        Assert.Equal(expected, address.ToString());
        Assert.Equal(expected, $"{address}");
        Assert.False(address.TryFormat(buffer[..(expected.Length - 1)], out _));
    }

    [Theory]
    [InlineData("B2:C5")]
    [InlineData("A1")]
    [InlineData("A1:IV16384")]
    public void Range_TryFormatMatchesToString(string text)
    {
        var range = ExcelRange.Parse(text);
        Span<char> buffer = stackalloc char[24];
        Assert.True(range.TryFormat(buffer, out int written));
        Assert.Equal(text, buffer[..written].ToString());
        Assert.Equal(text, range.ToString());
        Assert.False(range.TryFormat(buffer[..(text.Length - 1)], out _));
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void Definition_TryFormatMatchesGetDefinitionText(string name)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(name)));
        var buffer = new char[1024];
        foreach (var definedName in document.Names)
        {
            Assert.True(definedName.TryFormatDefinitionText(buffer, out int written, document));
            Assert.Equal(definedName.GetDefinitionText(document), new string(buffer, 0, written));
        }
    }

    [Fact]
    public void Read_NamesReferToTheFile()
    {
        // Name definitions read from a file are slices of it, and saving writes them back unchanged.
        var bytes = Sample("DB.FORM");
        var document = ExcelDocument.Read(bytes);
        Assert.Contains(document.Names, n => n.Definition.Length > 0);
        Assert.Equal(bytes, document.ToArray());
    }
}
