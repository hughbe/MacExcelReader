namespace MacExcelReader.Tests;

public class NumberFormatParseTests
{
    private static string DataFile(string name) => Path.Combine(AppContext.BaseDirectory, "Data", name);

    [Theory]
    [InlineData("Format Compiler", "number-formats-1.0.txt", ExcelVersion.Excel100)]
    [InlineData("Format Compiler 1.5", "number-formats-1.5.txt", ExcelVersion.Excel150)]
    public void Parse_MatchesExcelsFormatNumber(string sample, string list, ExcelVersion version)
    {
        // Each row of the sample got the format on the same line of the list with FORMAT.NUMBER, in Excel 1.00 or
        // Excel 1.5 (spec §7); Excel reuses an existing format where it can, so compare the format each cell uses.
        var document = TestData.Load(sample);
        string[] texts = File.ReadAllLines(DataFile(list));
        for (int row = 0; row < texts.Length; row++)
        {
            var excel = document.GetNumberFormat(document.Cells[row, 0].NumberFormat)!;
            var ours = ExcelNumberFormat.Parse(texts[row], version);
            Assert.True(
                excel.Descriptor.AsSpan().SequenceEqual(ours.Descriptor) && excel.Body.AsSpan().SequenceEqual(ours.Body),
                $"{texts[row]}: Excel {Convert.ToHexString(excel.Descriptor)} {Convert.ToHexString(excel.Body)}, ours {Convert.ToHexString(ours.Descriptor)} {Convert.ToHexString(ours.Body)}");
        }
    }

    [Fact]
    public void Parse_RecompilesEveryBuiltInFormat()
    {
        foreach (var format in ExcelNumberFormat.BuiltIn)
        {
            Assert.Equal(format.Body, ExcelNumberFormat.Parse(format.ToText()).Body);
            Assert.Equal(format.Descriptor, ExcelNumberFormat.Parse(format.ToText()).Descriptor);
        }
    }

    public static TheoryData<string> Sheets => TestData.AllSheets();

    [Theory]
    [MemberData(nameof(Sheets))]
    public void AddNumberFormat_RecompilesEverySampleFormat(string sample)
    {
        // Text from ToText compiles back to the stored bytes, in the form the file's version keeps (spec §7).
        var document = ExcelDocument.Read(File.ReadAllBytes(TestData.SampleFile(sample)));
        var version = sample.StartsWith("Excel 1.5", StringComparison.Ordinal) ? ExcelVersion.Excel150 : ExcelVersion.Excel100;
        foreach (var format in document.NumberFormats)
        {
            var target = ExcelDocument.CreateWorksheet();
            target.Version = document.Version;
            int index = target.AddNumberFormat(format.ToText(), version);
            Assert.Equal(ExcelNumberFormat.BuiltInCount, index);
            Assert.Equal(format.Descriptor, target.NumberFormats[0].Descriptor);
            Assert.Equal(format.Body, target.NumberFormats[0].Body);
            Assert.Equal(document.Version == ExcelFormatVersion.Version18 && format.RequiresVersion18, format.RequiresVersion18);
        }
    }

    [Theory]
    [InlineData("x0")]              // a letter Excel 1.00 does not know
    [InlineData("yyy")]             // a year is yy or yyyy
    [InlineData("General 0")]       // General with digits
    [InlineData("0;0;0;0")]         // four sections
    [InlineData("0.00E")]           // E needs a sign
    [InlineData("\"abc")]          // an unclosed quote
    [InlineData("*")]               // a fill needs its character
    [InlineData("d 0")]             // a date with digits
    [InlineData("[red]0")]          // colours are Excel 1.5's
    public void Parse_RejectsWhatExcel100Rejects(string text)
    {
        Assert.Throws<FormatException>(() => ExcelNumberFormat.Parse(text));
    }

    [Theory]
    [InlineData("0;0;0;@;0")]       // nothing after the text section
    [InlineData("0,%")]             // scaling with %
    [InlineData("@0")]              // @ with digits
    [InlineData("[purple]0")]       // an unknown colour
    [InlineData("[red][blue]0")]    // two colours
    public void Parse_RejectsWhatExcel15Rejects(string text)
    {
        Assert.Throws<FormatException>(() => ExcelNumberFormat.Parse(text, ExcelVersion.Excel150));
    }

    [Fact]
    public void Parse_FollowsEachReleasesRules()
    {
        // Excel 1.00 pads a thousands format to four positions; a trailing comma only turns thousands on.
        Assert.Equal("#,##0", ExcelNumberFormat.Parse("#,0").ToText());
        Assert.Equal("#,##0", ExcelNumberFormat.Parse("#,##0,").ToText());

        // Excel 1.5 scales by trailing commas, keeps unknown characters, and appends General as the text section.
        var scaled = ExcelNumberFormat.Parse("#,##0,", ExcelVersion.Excel150);
        Assert.Equal("#,##0,", scaled.ToText());
        Assert.True(scaled.RequiresVersion18);
        Assert.Equal("\\x0", ExcelNumberFormat.Parse("x0", ExcelVersion.Excel150).ToText());
        Assert.Equal(0x01, ExcelNumberFormat.Parse("0.00", ExcelVersion.Excel150).GetSection(ExcelNumberFormatSection.Text)[0]);
    }

    // Excel 1.00 writes the negative section of any format with more than one, and leaves out a zero section
    // with the positive section's bytes, even an empty one; Excel 1.5 leaves out only sections that repeat a
    // non-empty positive one. Both show what Excel 1.00 exported to SYLK for `Format Compiler`.
    [Theory]
    [InlineData(";;", ";", ";;")]
    [InlineData("0.0;;", "0.0;;", "0.0;;")]
    [InlineData("0;(0)", "0;(0)", "0;(0)")]
    [InlineData("0.00", "0.00", "0.00")]
    public void ToText_FollowsEachReleasesSectionRule(string format, string excel100, string excel150)
    {
        var compiled = ExcelNumberFormat.Parse(format);
        Assert.Equal(excel100, compiled.ToText(ExcelVersion.Excel100));
        Assert.Equal(excel150, compiled.ToText(ExcelVersion.Excel150));
    }

    [Fact]
    public void AddNumberFormat_NeedsVersion18ForExcel15Formats()
    {
        var document = ExcelDocument.CreateWorksheet();
        Assert.Equal(19, document.AddNumberFormat("0.0%"));
        Assert.Equal(20, document.AddNumberFormat("0.000", ExcelVersion.Excel150));   // version 17: three slots
        Assert.Equal(7, document.NumberFormats[1].Descriptor.Length);
        Assert.Throws<InvalidOperationException>(() => document.AddNumberFormat("[red]0", ExcelVersion.Excel150));

        document.Version = ExcelFormatVersion.Version18;
        var reread = ExcelDocument.CreateWorksheet();
        reread.Version = ExcelFormatVersion.Version18;
        Assert.Equal(19, reread.AddNumberFormat("[red]0", ExcelVersion.Excel150));
        reread.Cells["A1"].SetNumber(-2);
        reread.Cells["A1"].NumberFormat = 19;
        var saved = ExcelDocument.Read(reread.ToArray());
        Assert.Equal("-2", ExcelValueFormatter.FormatCell(saved.Cells["A1"], saved));

        // The new sheet's 41-byte window state is saved as Excel 1.5's 129 bytes, gridline colour automatic.
        var window = Assert.Single(saved.Windows);
        Assert.True(window.HasVersion18State);
        Assert.Null(window.GridlineColor);
        Assert.False(window.PanesFrozen);
    }
}
