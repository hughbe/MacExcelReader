using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// The public API's value types, helpers and checks.
/// </summary>
public class ApiTests
{
    [Fact]
    public void ExcelValue_KindsAccessorsAndText()
    {
        ExcelValue number = 12.5, text = "total", logical = true;
        var error = ExcelValue.FromError(ExcelError.NotAvailable);
        Assert.Equal(ExcelValueKind.Number, number.Kind);
        Assert.Equal(12.5, number.Number);
        Assert.Equal("12.5", number.ToString());
        Assert.Equal("total", text.Text);
        Assert.True(logical.Boolean);
        Assert.Equal("TRUE", logical.ToString());
        Assert.Equal("#N/A", error.ToString());
        Assert.True(ExcelValue.Empty.IsEmpty);
        Assert.Equal(string.Empty, ExcelValue.Empty.ToString());

        Assert.Throws<InvalidOperationException>(() => _ = text.Number);
        Assert.False(text.TryGetNumber(out _));
        Assert.True(text.TryGetText(out string? value));
        Assert.Equal("total", value);
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelValue.FromError((ExcelError)1));

        Assert.Equal(ExcelValue.FromNumber(12.5), number);
        Assert.NotEqual<ExcelValue>(1.0, true);
        Assert.Equal(((ExcelValue)"a").GetHashCode(), ExcelValue.FromText("a").GetHashCode());
    }

    [Fact]
    public void SetValue_StoresEachKindAsExcelDoes()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetValue(7);
        document.Cells["A2"].SetValue(-7);
        document.Cells["A3"].SetValue("text");
        document.Cells["A4"].SetValue(false);
        document.Cells["A5"].SetValue(ExcelValue.FromError(ExcelError.DivideByZero));
        document.Cells["A6"].SetValue(1);
        document.Cells["A6"].SetValue(ExcelValue.Empty);

        var reread = RoundTrip(document);
        Assert.Equal(ExcelCellType.Integer, reread.Cells["A1"].Type);
        Assert.Equal(ExcelCellType.Number, reread.Cells["A2"].Type);
        Assert.Equal("text", reread.Cells["A3"].Value.Text);
        Assert.False(reread.Cells["A4"].Value.Boolean);
        Assert.Equal(ExcelError.DivideByZero, reread.Cells["A5"].Value.Error);
        Assert.Equal(ExcelCellType.Empty, reread.Cells["A6"].Type);
    }

    [Fact]
    public void NonEmptyCells_ForeachAllocatesNothing()
    {
        // Reading values allocates only the strings of text cells, which this skips.
        var document = ExcelDocument.Read(Sample("Expenses"), copyData: false);
        double total = Sum(document);
        long before = GC.GetAllocatedBytesForCurrentThread();
        total += Sum(document);

        Assert.Equal(0, GC.GetAllocatedBytesForCurrentThread() - before);
        Assert.NotEqual(0, total);
        Assert.Equal(document.Cells.CountNonEmpty(), document.Cells.NonEmptyCells.Count());

        static double Sum(ExcelDocument document)
        {
            double total = 0;
            foreach (var cell in document.Cells.NonEmptyCells)
            {
                if (cell.Type != ExcelCellType.Text && cell.Value.TryGetNumber(out double value))
                {
                    total += value;
                }
            }

            return total;
        }
    }

    [Fact]
    public void ExcelRange_ParsesAndFormats()
    {
        var range = ExcelRange.Parse("$B$2:C5");
        Assert.Equal(new ExcelRange(1, 1, 4, 2), range);
        Assert.Equal("B2:C5", range.ToString());
        Assert.Equal(4, range.RowCount);
        Assert.Equal(2, range.ColumnCount);
        Assert.True(range.Contains(ExcelCellAddress.Parse("C3")));
        Assert.False(range.Contains(ExcelCellAddress.Parse("D3")));
        Assert.Equal("B2", ExcelRange.Parse("B2").ToString());
        Assert.False(ExcelRange.TryParse("C5:B2", out _));
        Assert.Throws<FormatException>(() => ExcelRange.Parse("nope"));
        Assert.Throws<ArgumentException>(() => ExcelDocument.CreateWorksheet().SetArrayFormula(new ExcelRange(2, 0, 0, 0), "=1"));
    }

    [Fact]
    public void Grid_AndAddressHelpers()
    {
        var document = ExcelDocument.CreateWorksheet();
        var address = ExcelCellAddress.Parse("C4");
        document.Cells[address].SetValue(3);
        Assert.Same(document.Cells["C4"], document.Cells[address]);
        Assert.Same(document.Cells["C4"], document.Cells.Find(address));
        Assert.Equal("IV", ExcelCellAddress.ColumnName(255));
        Assert.Equal("$B5", FormulaTokens.ReferenceText(0x8004, 1));
    }

    [Fact]
    public void DefinedName_MacroKindAndShortcutKey()
    {
        var name = new ExcelDefinedName("Go") { MacroKind = ExcelMacroKind.Command, ShortcutKey = 'g' };
        Assert.True(name.Flags.HasFlag(ExcelNameFlags.Macro));
        Assert.Equal(0x40, name.RawMacroKind);
        Assert.Equal((byte)'g', name.RawShortcutKey);

        name.MacroKind = ExcelMacroKind.Function;
        Assert.Equal(0x80, name.RawMacroKind);
        name.MacroKind = ExcelMacroKind.None;
        Assert.False(name.Flags.HasFlag(ExcelNameFlags.Macro));
        name.ShortcutKey = null;
        Assert.Equal(0, name.RawShortcutKey);
    }

    [Fact]
    public void ProtectAndUnprotect()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Protect("secret");
        Assert.True(document.State.IsProtected);
        Assert.Equal(ExcelPassword.Hash("secret"), document.State.PasswordHash);
        Assert.False(document.Unprotect("wrong"));
        Assert.True(document.State.IsProtected);
        Assert.True(document.Unprotect("secret"));
        Assert.False(document.State.IsProtected);

        document.Protect();
        Assert.True(document.Unprotect());
    }

    [Fact]
    public void Document_RejectsChartsAndUnknownVersions()
    {
        var document = ExcelDocument.CreateWorksheet();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Kind = ExcelDocumentKind.Chart);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Version = (ExcelFormatVersion)19);
        document.Kind = ExcelDocumentKind.MacroSheet;
        Assert.Equal("XLPG", document.FinderType);
    }

    [Fact]
    public async Task AsyncLoadAndSave_RoundTrip()
    {
        string path = Path.Combine(Path.GetTempPath(), $"MacExcelReader-{Guid.NewGuid():N}");
        try
        {
            var bytes = Sample("Expenses");
            await File.WriteAllBytesAsync(path, bytes);
            var document = await ExcelDocument.LoadAsync(path);
            await document.SaveAsync(path);
            Assert.Equal(bytes, await File.ReadAllBytesAsync(path));

            using var stream = new MemoryStream();
            await document.WriteAsync(stream);
            stream.Position = 0;
            Assert.Equal(bytes, (await ExcelDocument.ReadAsync(stream)).ToArray());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ChartViews_AreCreatedOnce()
    {
        var chart = ChartDocument.Create(ChartType.ColumnOrBar).MainChart;
        Assert.Same(chart.ValueScale, chart.ValueScale);
        Assert.Same(chart.CategoryAxis.Line, chart.CategoryAxis.Line);
        chart.ValueScale.Minimum = 5;
        Assert.Equal(5, chart.ValueScale.Minimum);
    }

    [Fact]
    public void NumberFormats_AreEqualByTheirBytes()
    {
        var parsed = ExcelNumberFormat.Parse("0.00");
        Assert.Equal(ExcelNumberFormat.BuiltIn[2], parsed);
        Assert.Equal(ExcelNumberFormat.BuiltIn[2].GetHashCode(), parsed.GetHashCode());
        Assert.NotEqual(ExcelNumberFormat.BuiltIn[1], parsed);
    }
}
