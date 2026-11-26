using System.Globalization;

namespace MacExcelReader.Tests;

/// <summary>
/// How Excel 1.00 draws cells (docs/Display.md), checked against what it showed on screen for sheets built by the library
/// (docs/Testing.md). Unless a test says otherwise, the sheet uses the default font, Geneva 10, and Excel 1.00, 1.03 and
/// 1.5 drew it the same, pixel for pixel.
/// </summary>
public class DisplayTests
{
    private static readonly ushort[] s_fitWidths = [12, 15, 18, 19, 23, 25, 26, 32, 33, 39, 47, 61, 75];

    public static TheoryData<string, string, string[]> FitRows()
    {
        var data = new TheoryData<string, string, string[]>();
        foreach (string line in File.ReadLines(Path.Combine(AppContext.BaseDirectory, "Data", "display-1.00.txt")))
        {
            if (line.StartsWith('#'))
            {
                continue;
            }

            string[] fields = line.Split('\t');
            data.Add(fields[0], fields[1], fields[2..]);
        }

        return data;
    }

    public static TheoryData<string, string, string[], ExcelVersion> FitRowsByRelease()
    {
        // Excel 1.00, 1.03 and 1.5 showed the Fit sheet the same, pixel for pixel.
        var data = new TheoryData<string, string, string[], ExcelVersion>();
        foreach (var row in FitRows())
        {
            foreach (var version in new[] { ExcelVersion.Excel100, ExcelVersion.Excel103, ExcelVersion.Excel150 })
            {
                data.Add((string)row[0], (string)row[1], (string[])row[2], version);
            }
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(FitRowsByRelease))]
    public void General_FitsTheColumnAsExcelShowsIt(string face, string value, string[] expected, ExcelVersion version)
    {
        var document = ExcelDocument.CreateWorksheet();
        for (int c = 0; c < s_fitWidths.Length; c++)
        {
            document.State.SetColumnWidth(c, s_fitWidths[c]);
            var cell = document.Cells[0, c];
            cell.SetNumber(double.Parse(value, CultureInfo.InvariantCulture));
            cell.Bold = face == "bold";
            cell.Italic = face == "italic";
        }

        string[] actual = [.. Enumerable.Range(0, s_fitWidths.Length).Select(c => ExcelValueFormatter.FormatCell(document.Cells[0, c], document, c, version))];
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(-0.4, new[] { "0", "0", "-0.4", "-0.4", "-0.4" })]
    [InlineData(-0.001234, new[] { "0", "0", "0", "0", "-0.001" })]
    [InlineData(-0.04, new[] { "0", "0", "0", "-0.04", "-0.04" })]
    [InlineData(-1E-10, new[] { "0", "0", "0", "0", "-1E-10" })]
    [InlineData(-0.6, new[] { "-1", "-1", "-0.6", "-0.6", "-0.6" })]
    [InlineData(0.4, new[] { "0", "0.4", "0.4", "0.4", "0.4" })]
    public void General_NegativeRoundingToNothing_ShowsZero(double value, string[] expected)
    {
        // Columns 19, 25, 33, 40 and 47 pixels wide (the Zero sheet).
        ushort[] widths = [19, 25, 33, 40, 47];
        var document = ExcelDocument.CreateWorksheet();
        for (int c = 0; c < widths.Length; c++)
        {
            document.State.SetColumnWidth(c, widths[c]);
            document.Cells[0, c].SetNumber(value);
        }

        Assert.Equal(expected, Enumerable.Range(0, widths.Length).Select(c => ExcelValueFormatter.FormatCell(document.Cells[0, c], document, c)));
    }

    // The Spill sheet: default 75-pixel columns, so column B spans pixels 75-149.
    private static ExcelDocument Spill()
    {
        var d = ExcelDocument.CreateWorksheet();
        d.Cells["B1"].SetText("General text spills right");
        d.Cells["B2"].SetText("Blocked by C2 text");
        d.Cells["C2"].SetText("x");
        d.Cells["D3"].SetText("Right aligned spills left");
        d.Cells["D3"].Alignment = ExcelAlignment.Right;
        d.Cells["C4"].SetText("Centred text spills both ways");
        d.Cells["C4"].Alignment = ExcelAlignment.Center;
        d.Cells["B5"].SetNumber(123456789012);
        d.Cells["B5"].NumberFormat = 2;
        d.Cells["B6"].SetBoolean(true);
        d.Cells["C6"].SetError(ExcelError.NotAvailable);
        d.Cells["D6"].SetText("abc");
        d.Cells["E6"].SetNumber(42);
        d.Cells["B7"].SetText("ab");
        d.Cells["B7"].Alignment = ExcelAlignment.Fill;
        d.Cells["C7"].Alignment = ExcelAlignment.Fill;
        d.Cells["D7"].Alignment = ExcelAlignment.Fill;
        d.Cells["B8"].SetText("xyz");
        d.Cells["B8"].Alignment = ExcelAlignment.Fill;
        d.Cells["B9"].SetText("Left");
        d.Cells["B9"].Alignment = ExcelAlignment.Left;
        d.Cells["C9"].SetText("Right");
        d.Cells["C9"].Alignment = ExcelAlignment.Right;
        d.Cells["D9"].SetText("Mid");
        d.Cells["D9"].Alignment = ExcelAlignment.Center;
        d.Cells["E9"].SetText("Ital");
        d.Cells["E9"].Italic = true;
        d.Cells["E9"].Alignment = ExcelAlignment.Right;
        d.Cells["B11"].SetText("Long text into a number");
        d.Cells["C11"].SetNumber(1);
        d.Cells["C12"].SetText("Right aligned blocked");
        d.Cells["C12"].Alignment = ExcelAlignment.Right;
        d.Cells["A12"].SetNumber(9);
        d.Cells["B13"].SetText("Empty-text neighbour");
        d.Cells["C13"].SetText("");
        d.Cells["B14"].SetText("Spill over formatted blank");
        d.Cells["C14"].Bold = true;
        return d;
    }

    [Theory]
    // Text starts 2 pixels in from the left, ends 2 in from the right, or is centred on the cell.
    [InlineData("B9", "Left", ExcelAlignment.Left, 77)]
    [InlineData("C9", "Right", ExcelAlignment.Right, 197)]
    [InlineData("D9", "Mid", ExcelAlignment.Center, 253)]
    [InlineData("E9", "Ital", ExcelAlignment.Right, 351)]  // italic: 5 pixels more room on the right
    // General: TRUE and errors in the middle, text on the left, numbers on the right.
    [InlineData("B6", "TRUE", ExcelAlignment.Center, 99)]
    [InlineData("C6", "#N/A", ExcelAlignment.Center, 173)]
    [InlineData("D6", "abc", ExcelAlignment.Left, 227)]
    [InlineData("E6", "42", ExcelAlignment.Right, 359)]
    // A number too wide shows as many #s as fit, centred.
    [InlineData("B5", "#######", ExcelAlignment.Center, 81)]
    public void Layout_PlacesTextAsExcelDoes(string address, string text, ExcelAlignment alignment, int x)
    {
        var document = Spill();
        var cell = ExcelCellAddress.Parse(address);
        var display = ExcelValueFormatter.LayoutCell(document, cell.Row, cell.Column)!.Value;
        Assert.Equal((text, alignment, x), (display.Text, display.Alignment, display.X));
    }

    [Theory]
    [InlineData("B1", 75, 224)]   // runs into C
    [InlineData("B2", 75, 149)]   // stopped by C2
    [InlineData("D3", 150, 299)]  // right aligned: runs left into C
    [InlineData("C4", 75, 299)]   // centred: both ways
    [InlineData("B11", 75, 149)]  // stopped by a number
    [InlineData("C12", 75, 224)]  // right aligned, runs into B only as far as it needs
    [InlineData("B13", 75, 149)]  // empty text is a value
    [InlineData("B14", 75, 224)]  // formatting alone is not
    public void Layout_TextRunsIntoBlankNeighbours(string address, int clipLeft, int clipRight)
    {
        var document = Spill();
        var cell = ExcelCellAddress.Parse(address);
        var display = ExcelValueFormatter.LayoutCell(document, cell.Row, cell.Column)!.Value;
        Assert.Equal((clipLeft, clipRight), (display.ClipLeft, display.ClipRight));
    }

    [Fact]
    public void Layout_FillRepeatsAcrossItsRun()
    {
        var document = Spill();
        var run = ExcelValueFormatter.LayoutCell(document, 6, 1)!.Value;
        Assert.Equal(string.Concat(Enumerable.Repeat("ab", 20)), run.Text);
        Assert.Equal((77, 75, 299), (run.X, run.ClipLeft, run.ClipRight));
        Assert.Equal("xyzxyzxyzxyz", ExcelValueFormatter.LayoutCell(document, 7, 1)!.Value.Text);
    }

    // The More sheet: B is 30 pixels wide and D is hidden.
    private static ExcelDocument More()
    {
        var d = ExcelDocument.CreateWorksheet();
        d.State.SetColumnWidth(1, 30);
        d.State.SetColumnWidth(3, 0);
        d.Cells["B1"].SetBoolean(true);
        d.Cells["B2"].SetError(ExcelError.DivideByZero);
        d.Cells["B3"].SetBoolean(false);
        d.Cells["C4"].SetNumber(12);
        d.Cells["C4"].Alignment = ExcelAlignment.Fill;
        d.Cells["C5"].SetNumber(1.5);
        d.Cells["C5"].NumberFormat = 2;
        d.Cells["C5"].Alignment = ExcelAlignment.Fill;
        d.Cells["C6"].SetText("Spills across the hidden column D");
        d.Cells["E7"].SetText("This fill text is longer than the cell");
        d.Cells["E7"].Alignment = ExcelAlignment.Fill;
        d.Cells["E8"].SetText("Spill over empty fill cell");
        d.Cells["F8"].Alignment = ExcelAlignment.Fill;
        d.Cells["E9"].SetText("x");
        d.Cells["F9"].SetText("Centred long text blocked left");
        d.Cells["F9"].Alignment = ExcelAlignment.Center;
        d.Cells["A10"].SetText("Right text at column A");
        d.Cells["A10"].Alignment = ExcelAlignment.Right;
        d.Cells["E11"].SetText("Text before a formatted blank and a number");
        d.Cells["F11"].Alignment = ExcelAlignment.Right;
        d.Cells["G11"].SetNumber(3);
        d.Cells["B12"].SetBoolean(true);
        d.Cells["B12"].Bold = true;
        return d;
    }

    [Theory]
    [InlineData("B1", "TRUE")]          // 26 pixels: just fits the 26 the column leaves
    [InlineData("B2", "##")]            // #DIV/0! and FALSE do not
    [InlineData("B3", "##")]
    [InlineData("B12", "##")]           // nor does bold TRUE, a pixel wider per letter
    [InlineData("C4", "1212121212")]    // numbers fill too
    [InlineData("C5", "1.501.501.50")]  // (75 - 3) / 24 = 3 copies
    [InlineData("E7", "This fill text is longer than the cell")]  // too long to repeat: once, cut off
    public void Layout_ShowsWhatExcelShows(string address, string text)
    {
        var document = More();
        var cell = ExcelCellAddress.Parse(address);
        Assert.Equal(text, ExcelValueFormatter.LayoutCell(document, cell.Row, cell.Column)!.Value.Text);
    }

    [Theory]
    [InlineData("C6", 105, 329)]  // runs across hidden D, through E into F (A 75, B 30, C 75, D 0)
    [InlineData("E7", 180, 254)]  // a fill never runs into a blank that is not Fill aligned
    [InlineData("E8", 180, 329)]  // a blank Fill-aligned neighbour does not stop text
    [InlineData("F9", 255, 404)]  // centred, stopped on the left by E9, runs right into G
    [InlineData("A10", 0, 74)]    // right aligned in A: cut off at the sheet's left edge
    [InlineData("E11", 180, 329)] // runs over the formatted blank F11, stops at the number in G11
    public void Layout_ClipsAsExcelDoes(string address, int clipLeft, int clipRight)
    {
        var document = More();
        var cell = ExcelCellAddress.Parse(address);
        var display = ExcelValueFormatter.LayoutCell(document, cell.Row, cell.Column)!.Value;
        Assert.Equal((clipLeft, clipRight), (display.ClipLeft, display.ClipRight));
    }

    [Fact]
    public void Layout_SkipsBlankAndHiddenCells()
    {
        var document = More();
        document.Cells["D1"].SetText("hidden");
        Assert.Null(ExcelValueFormatter.LayoutCell(document, 0, 3));
        Assert.Null(ExcelValueFormatter.LayoutCell(document, 0, 5));
    }

    [Theory]
    [MemberData(nameof(TestData.AllSheets), MemberType = typeof(TestData))]
    public void Layout_HandlesEverySampleCell(string sample)
    {
        var document = ExcelDocument.Read(TestData.Sample(sample));
        foreach (var cell in document.Cells.NonEmptyCells)
        {
            var display = ExcelValueFormatter.LayoutCell(document, cell.Row, cell.Column);
            if (display is { } shown)
            {
                Assert.True(shown.ClipLeft <= shown.ClipRight);
            }
        }
    }

    [Theory]
    // Sizes the System has no bitmap for (the Gen11, Gen11b, Gen16 and Mon12 sheets). QuickDraw scales the next
    // larger size, or Monaco 9, and Excel adds a pixel to the widths of 0 and # it divides by.
    [InlineData(3, 11, 40, 123456789, "1E+08")]
    [InlineData(3, 11, 40, 1234.5678, "1235")]
    [InlineData(3, 16, 40, 123456789, "##")]
    [InlineData(3, 16, 40, 1234.5678, "##")]
    [InlineData(4, 12, 40, 123456789, "####")]
    [InlineData(4, 12, 40, 1234.5678, "1235")]
    [InlineData(3, 11, 33, 1234, "1230")]   // a digit counts as 8 pixels (7.33, rounded down, + 1): 29 / 8 = 3
    [InlineData(3, 11, 36, 1234, "1234")]
    [InlineData(3, 11, 37, 12345, "###")]   // 1E+04 is 34 pixels; # counts as 9 (8.25 + 1)
    [InlineData(3, 11, 38, 12345, "1E+04")]
    [InlineData(3, 11, 43, 12345, "1E+04")]
    public void ScaledFont_FitsAsExcelShowsIt(int font, int size, int width, double value, string expected)
    {
        var document = ExcelDocument.CreateWorksheet();
        document.State.StandardFont = (ushort)font;
        document.State.StandardFontSize = (ushort)size;
        document.State.SetColumnWidth(0, (ushort)width);
        document.Cells[0, 0].SetNumber(value);
        Assert.Equal(expected, ExcelValueFormatter.FormatCell(document.Cells[0, 0], document, 0));
    }

    [Theory]
    [InlineData(33)]
    [InlineData(43)]
    public void ScaledFont_TrueFitsByItsScaledWidth(int width)
    {
        // TRUE is 27.5 pixels in Geneva 11; Excel adds nothing when it checks the fit.
        var document = ExcelDocument.CreateWorksheet();
        document.State.StandardFont = 3;
        document.State.StandardFontSize = 11;
        document.State.SetColumnWidth(0, (ushort)width);
        document.Cells[0, 0].SetBoolean(true);
        Assert.Equal("TRUE", ExcelValueFormatter.FormatCell(document.Cells[0, 0], document, 0));
    }

    [Theory]
    [InlineData(3, 10, 13, 10)]  // measured on screen: ink on rows 3-10 of 13, descenders to 12
    [InlineData(3, 11, 15, 11)]  // scaled from Geneva 12: descent 2.75 rounds to 3
    [InlineData(3, 16, 21, 16)]  // scaled from Geneva 18
    [InlineData(4, 12, 16, 12)]  // scaled up from Monaco 9
    public void RowMetrics_ComeFromTheFont(int font, int size, int height, int baseline)
    {
        var document = ExcelDocument.CreateWorksheet();
        document.State.StandardFont = (ushort)font;
        document.State.StandardFontSize = (ushort)size;
        Assert.Equal((height, baseline), ExcelValueFormatter.GetRowMetrics(document));
    }

    // The Gen11b sheet: Geneva 11, which Excel draws from Geneva 12, in columns 33 to 43 pixels wide.
    [Theory]
    [InlineData(ExcelVersion.Excel100, 12345.0, "### ### ### ### ### 1E+04 1E+04 1E+04 1E+04 1E+04 1E+04")]
    [InlineData(ExcelVersion.Excel150, 12345.0, "### ### ### ### ### 1E+04 1E+04 1E+04 1E+04 1E+04 1E+04")]
    [InlineData(ExcelVersion.Excel103, 12345.0, "### ### ### ### ### ### ### #### #### #### 1E+04")]
    [InlineData(ExcelVersion.Excel100, 1234.0, "1230 1230 1230 1234 1234 1234 1234 1234 1234 1234 1234")]
    [InlineData(ExcelVersion.Excel150, 1234.0, "1230 1230 1230 1234 1234 1234 1234 1234 1234 1234 1234")]
    [InlineData(ExcelVersion.Excel103, 1234.0, "### ### ### ### 1234 1234 1234 1234 1234 1234 1234")]
    public void ScaledFont_Excel103AddsAPixelPerCharacter(ExcelVersion version, double value, string expected)
    {
        var document = ExcelDocument.CreateWorksheet();
        document.State.StandardFont = 3;
        document.State.StandardFontSize = 11;
        for (int c = 0; c < 11; c++)
        {
            document.State.SetColumnWidth(c, (ushort)(33 + c));
            document.Cells[0, c].SetNumber(value);
        }

        Assert.Equal(expected, string.Join(' ', Enumerable.Range(0, 11).Select(c => ExcelValueFormatter.FormatCell(document.Cells[0, c], document, c, version))));
    }

    [Theory]
    [InlineData(ExcelVersion.Excel100, "TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE")]
    [InlineData(ExcelVersion.Excel150, "TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE")]
    [InlineData(ExcelVersion.Excel103, "### ### ### TRUE TRUE TRUE TRUE TRUE TRUE TRUE TRUE")]
    public void ScaledFont_Excel103FitsTrueWithAPixelPerLetter(ExcelVersion version, string expected)
    {
        // TRUE is 28 pixels in Geneva 11: it fits 29 pixels of room in 1.00 and 1.5, and needs 32 in 1.03.
        var document = ExcelDocument.CreateWorksheet();
        document.State.StandardFont = 3;
        document.State.StandardFontSize = 11;
        for (int c = 0; c < 11; c++)
        {
            document.State.SetColumnWidth(c, (ushort)(33 + c));
            document.Cells[0, c].SetBoolean(true);
        }

        Assert.Equal(expected, string.Join(' ', Enumerable.Range(0, 11).Select(c => ExcelValueFormatter.FormatCell(document.Cells[0, c], document, c, version))));
    }

    [Fact]
    public void MarginColumn_IsNeverDrawn()
    {
        // A damaged file can hold a value in the table's margin column, one past IV (spec §6.1). Excel draws nothing
        // there, so it formats and lays out as a hidden column does rather than throwing.
        var document = ExcelDocument.CreateWorksheet();
        document.Cells.Grow(1, ExcelCellGrid.MaxColumns + 1);
        var margin = document.Cells.Find(0, ExcelCellGrid.MaxColumns)!;
        margin.SetNumber(12345);
        document.Cells[0, 255].SetNumber(12345);
        document.State.SetColumnWidth(255, 0);
        Assert.Equal(
            ExcelValueFormatter.FormatCell(document.Cells[0, 255], document, 255),
            ExcelValueFormatter.FormatCell(margin, document, ExcelCellGrid.MaxColumns));
        Assert.Null(ExcelValueFormatter.LayoutCell(document, 0, 255));
        Assert.Null(ExcelValueFormatter.LayoutCell(document, 0, ExcelCellGrid.MaxColumns));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelValueFormatter.LayoutCell(document, 0, ExcelCellGrid.MaxColumns + 1));
    }
}
