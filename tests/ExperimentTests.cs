using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// Files Excel itself saved while running experiments in the test harness (docs/Testing.md): each was built by the
/// library or copied from a sample, changed by Excel through macro commands, and saved by Excel. Each test pins down
/// what the experiment showed about the format.
/// </summary>
public class ExperimentTests
{
    [Theory]
    [InlineData("C2", "#DIV/0!", 0x07)]
    [InlineData("C3", "#NUM!", 0x24)]
    [InlineData("C4", "#N/A", 0x2A)]
    [InlineData("C5", "#REF!", 0x17)]
    [InlineData("C6", "#VALUE!", 0x0F)]
    [InlineData("C7", "#NULL!", 0x00)]
    [InlineData("C8", "#NAME?", 0x1D)]
    public void Errors_UseTheErrorCodes(string address, string text, byte code)
    {
        // Each formula entered with FORMULA and calculated by Excel 1.03 (spec §12.4).
        var result = Load("Errors").Cells[address].Formula!.Result;
        Assert.Equal(ExcelFormulaResultKind.Error, result.Kind);
        Assert.Equal(code, result.ErrorCode);
        Assert.Equal(text, result.Value);
    }

    [Fact]
    public void Errors_EmptyIntersectionCompilesToMemoryErrorNull()
    {
        // FORMULA("=R1C1 R1C2"): two cells that do not overlap. Excel writes a memory error (27) holding #NULL!.
        var document = Load("Errors");
        var formula = document.Cells["C7"].Formula!;
        Assert.Equal(0x47, formula.Expression[0]);
        Assert.Equal(0x00, formula.Expression[2]);
        Assert.Equal(
            FormulaTokens.ClearStaleBytes(formula.Expression),
            FormulaTokens.ClearStaleBytes(ExcelFormula.Parse("=$A$1 $B$1", document).Expression));
    }

    [Fact]
    public void MemoryError_UnionNeedingOneValueCompilesToMemoryErrorValue()
    {
        // FORMULA("=(R1C1,R2C2)"): a union where a single value is needed. Excel writes a memory error (27) holding
        // #VALUE!, and no cached areas (spec §15.2).
        var document = Load("Memory Error");
        var formula = document.Cells["B13"].Formula!;
        Assert.Equal(0x47, formula.Expression[0]);
        Assert.Equal(0x0F, formula.Expression[2]);
        Assert.Empty(formula.Fixups.ToArray());
        var parsed = ExcelFormula.Parse("=($A$1,$B$2)", document);
        Assert.Equal(FormulaTokens.ClearStaleBytes(formula.Expression), FormulaTokens.ClearStaleBytes(parsed.Expression));
        Assert.Empty(parsed.Fixups.ToArray());
    }

    [Fact]
    public void Recalculated_StateWordMakesOpenRecalculate()
    {
        // Written by the library with every result 0 and state word 0x018 = 1; Excel 1.03 recalculated it on Open.
        // The same sheet with 0x018 = 0 kept its zeros even after CALCULATE.NOW (spec §6).
        var document = Load("Recalculated");
        Assert.Equal((byte)0x07, document.Cells["B2"].Formula!.Result.ErrorCode);
        Assert.Equal("#NUM!", document.Cells["B3"].Value);
        Assert.Equal("#N/A", document.Cells["B4"].Value);
    }

    [Theory]
    [InlineData("C", 1, 12)]
    [InlineData("D", 10, 75)]
    [InlineData("E", 20, 145)]
    [InlineData("F", 0, 0)]
    [InlineData("G", 2.5, 23)]
    public void ColumnWidths_ArePixels(string column, double characters, int pixels)
    {
        // COLUMN.WIDTH(characters): the width of "0" in the standard font (7 pixels) per character, plus 5,
        // rounded; 0 hides the column (spec §6, 0x042).
        var state = Load("Column Widths").State;
        int index = ExcelCellAddress.Parse(column + "1").Column;
        Assert.Equal(pixels, state.GetColumnWidth(index));
        Assert.Equal(pixels, characters == 0 ? 0 : (int)Math.Round(characters * 7 + 5, MidpointRounding.AwayFromZero));
    }

    [Fact]
    public void TableRowInput_HasOnlyARowInput()
    {
        // TABLE(Y7,) on A5:C6: byte 14 is 0x40, the first input is the row input and there is no second.
        var formula = Load("Table Row Input").Cells["B6"].Formula!;
        Assert.True(formula.IsTable);
        Assert.Equal(0x40, formula.Fixups[3 + 14]);
        Assert.Equal((ExcelCellAddress.Parse("Y7"), (ExcelCellAddress?)null), formula.TableInputs);
        Assert.Equal("{=TABLE($Y$7,)}", formula.ToText(Load("Table Row Input")));
    }

    [Fact]
    public void TableOneArgument_StoresTwoInputs()
    {
        // The macro command TABLE(Y7) with one argument takes a second input from an uncleared buffer: P1 here (spec §12.6).
        var formula = Load("Table One Argument").Cells["B6"].Formula!;
        Assert.Equal(0xC0, formula.Fixups[3 + 14]);
        Assert.Equal((ExcelCellAddress.Parse("Y7"), ExcelCellAddress.Parse("P1")), formula.TableInputs);
    }

    [Theory]
    [InlineData("Calculation No Tables", -1)]
    [InlineData("Calculation Manual", 0)]
    public void CalculationMode_Values(string name, short mode)
    {
        // CALCULATION(2) is automatic except tables, CALCULATION(3) manual (spec §6, 0x002).
        Assert.Equal(mode, Load(name).State.CalculationMode);
    }

    [Fact]
    public void PageBreaks_SetPageBreakAboveAndLeftOfTheActiveCell()
    {
        // SET.PAGE.BREAK with C5 active (spec §14).
        var document = Load("Page Breaks");
        Assert.Equal([3], document.RowPageBreaks);
        Assert.Equal([(byte)1], document.ColumnPageBreaks);
        Assert.Equal(document.ToArray(), Sample("Page Breaks"));
    }

    [Fact]
    public void MacroNames_KindByte()
    {
        // DEFINE.NAME("Cmd",C1,2,"q"), DEFINE.NAME("Fn",C2,1) and DEFINE.NAME("Plain",C3) (spec §11).
        var names = Load("Macro Names").Names;
        var command = names.Single(n => n.Name == "Cmd");
        Assert.True(command.IsCommand);
        Assert.Equal((byte)'q', command.ShortcutKey);
        Assert.Equal(0x44, command.MacroKind);   // the low nibble is junk (spec §11)
        var function = names.Single(n => n.Name == "Fn");
        Assert.True(function.IsFunctionMacro);
        Assert.Equal(ExcelDefinedName.FunctionKind, function.MacroKind);
        var plain = names.Single(n => n.Name == "Plain");
        Assert.False(plain.IsCommand);
        Assert.False(plain.IsFunctionMacro);
        Assert.All([command, function, plain], n => Assert.NotEqual(0, n.Flags & ExcelDefinedName.ReferenceFlag));
    }

    [Fact]
    public void Links_ActiveSheetEntryHasTheNulName()
    {
        // FORMULA("=Other!A1") and FORMULA("=!A1+1") in R1C1 mode: A1 is a name there, on Other and on the active
        // sheet, whose entry is named "\0" (spec §8).
        var documents = Load("Links").ExternalDocuments;
        Assert.Equal(["Other", "\0"], documents.Select(d => d.Name));
        Assert.All(documents, d => Assert.Equal(["A1"], d.Names));
    }

    [Fact]
    public void PieExploded_PointFormatHoldsTheExplosion()
    {
        // GALLERY.PIE(3): the first slice pulled out by 25% of the radius, in a data-point format record.
        var chart = ChartDocument.Load(SamplePath("Pie Exploded")).MainChart;
        Assert.Equal(ChartType.Pie, chart.Type);
        Assert.Single(chart.PointFormats);
        Assert.Equal(25, chart.GetPointFormat(0).Explosion);
        Assert.Equal(0, chart.DefaultSeriesFormat.Explosion);
    }

    [Fact]
    public void PieExplodedAll_DefaultSeriesFormatHoldsTheExplosion()
    {
        // GALLERY.PIE(4): every slice pulled out by 20%.
        var chart = ChartDocument.Load(SamplePath("Pie Exploded All")).MainChart;
        Assert.Empty(chart.PointFormats);
        Assert.Equal(20, chart.DefaultSeriesFormat.Explosion);
    }

    [Fact]
    public void BarStacked_OverlapsFully()
    {
        // GALLERY.BAR(3): stacked bars, overlapped 100%.
        var chart = ChartDocument.Load(SamplePath("Bar Stacked")).MainChart;
        Assert.True(chart.IsBar);
        Assert.True(chart.IsStacked);
        Assert.Equal(-100, chart.Overlap);
    }

    [Fact]
    public void Combination_OverlayRecordSetsTheFourOptionFlags()
    {
        // COMBINATION(1) on a line chart: a column main chart with a line overlay, which Excel writes first. Its record has (1, 0) at offset 18 and all
        // four option flags (offset 22, bits 11–8) set; the main chart's are clear (spec §16.5).
        var document = ChartDocument.Load(SamplePath("Combination"));
        var overlay = document.OverlayChart!;
        Assert.Equal([0x00, 0x01, 0x00, 0x00, 0x0F, 0x00], overlay.Record[18..24]);
        Assert.Equal(ChartType.Line, overlay.Type);
        Assert.Equal(ChartType.ColumnOrBar, document.MainChart.Type);
        Assert.Equal(0, document.MainChart.Record[22] & 0x0F);
    }

    [Fact]
    public void Combination_OverlaySharesTheMainChartsScale()
    {
        var document = ChartDocument.Load(SamplePath("Combination"));
        Assert.True(document.OverlayChart!.SharesValueAxisScale);
        Assert.True(document.OverlayChart.SharesXAxisScale);
        Assert.False(document.MainChart.SharesValueAxisScale);
        Assert.Equal(0, document.OverlayChart.CategoryCount);
        Assert.Equal(12, document.MainChart.CategoryCount);   // PROFIT CHART's 12 revenue figures
    }

    [Fact]
    public void AreaLabels_LabelEachSeries()
    {
        // GALLERY.AREA(5): one data label per series, from built-in text object -2 (spec §16.8).
        var format = ChartDocument.Load(SamplePath("Area Labels")).MainChart.DefaultSeriesFormat;
        Assert.True(format.SeriesLabel);
        Assert.False(format.PointLabels);
    }

    [Fact]
    public void PiePercentages_LabelEachPoint()
    {
        // GALLERY.PIE(6): a label on every slice, from built-in text object -1, with bit 9 (percentages) too.
        var chart = ChartDocument.Load(SamplePath("Pie Percentages")).MainChart;
        Assert.True(chart.DefaultSeriesFormat.PointLabels);
        Assert.False(chart.DefaultSeriesFormat.SeriesLabel);
        Assert.True(chart.DefaultSeriesFormat.PercentLabels);
    }

    [Fact]
    public void LineLogScale_SetsTheLogarithmicFlag()
    {
        // GALLERY.LINE(6).
        Assert.True(ChartDocument.Load(SamplePath("Line Log Scale")).MainChart.ValueScale.Logarithmic);
    }

    [Fact]
    public void Scatter_HasAnXAxisScale()
    {
        // GALLERY.SCATTER(1): the category axis becomes a value axis with its own scale record.
        var chart = ChartDocument.Load(SamplePath("Scatter")).MainChart;
        Assert.Equal(ChartType.Scatter, chart.Type);
        Assert.NotNull(chart.XScale);
    }

    [Fact]
    public void AreaDropLines_SetsTheDropLinePattern()
    {
        // GALLERY.AREA(3): drop lines, the line format at offset 254.
        var chart = ChartDocument.Load(SamplePath("Area Drop Lines")).MainChart;
        Assert.Equal(ChartType.Area, chart.Type);
        Assert.Equal(1, chart.Record[255]);
    }

    [Theory]
    [InlineData("Colour Formats")]
    [InlineData("Gridline Colour")]
    [InlineData("Frozen Panes")]
    public void Excel15Features_SaveAsVersion18(string name)
    {
        // Excel 1.5 saves version 18 for colour or @ number formats, gridline colours and frozen panes (spec §2.1).
        Assert.Equal(ExcelDocument.Excel15FormatVersion, Load(name).Version);
    }

    [Fact]
    public void Literals_IntegerConstantsStopAt65535()
    {
        // Excel 1.00: FORMULA("=65535") and FORMULA("=65536") (spec §15.4).
        var document = Load("Literals");
        Assert.Equal([0x1E, 0xFF, 0xFF], document.Cells["C2"].Formula!.Expression.ToArray());
        Assert.Equal(0x1F, document.Cells["C3"].Formula!.Expression[0]);
    }

    [Fact]
    public void Literals_DatesUseThe1904System()
    {
        // Excel 1.00: DATE(1984,1,1) is serial number 29220 (spec §12.3).
        Assert.Equal(29220.0, Convert.ToDouble(Load("Literals").Cells["C4"].Value));
    }

    [Fact]
    public void Literals_Excel100CompilesIntersectionsAndDeletedReferences()
    {
        // Excel 1.00 reads a space as the intersection operator: two cells that do not overlap give a memory error
        // holding #NULL!. "=#REF!" is a deleted reference (2A), not the error constant (spec §15.4).
        var document = Load("Literals");
        var intersection = document.Cells["C5"].Formula!;
        Assert.Equal(0x47, intersection.Expression[0]);
        Assert.Equal(0x00, intersection.Expression[2]);
        Assert.Equal(0x0F, intersection.Expression[^1]);
        Assert.Equal(0x4A, document.Cells["C7"].Formula!.Expression[0]);
    }

    [Theory]
    [InlineData("C2", 0x60)]   // =SUM({1,2,3})
    [InlineData("C4", 0x40)]   // ={1,2}+1
    [InlineData("C5", 0x40)]   // ={1,2;3,4}
    [InlineData("C6", 0x60)]   // =TRANSPOSE({1,2})
    public void ArrayConstants_TakeArrayClassOnlyAsArguments(string address, byte token)
    {
        // Excel 1.00 (spec §15.4).
        var document = Load("Array Constants");
        var formula = document.Cells[address].Formula!;
        Assert.Equal(token, formula.Expression.ToArray().First(b => (b & 0x1F) == 0x00 && b >= 0x40));
        Assert.Equal(
            FormulaTokens.ClearStaleBytes(formula.Expression),
            FormulaTokens.ClearStaleBytes(ExcelFormula.Parse(formula.ToText(document), document).Expression));
    }

    [Fact]
    public void FillFormat_StarIsTheRepeatToken()
    {
        // Excel 1.00: FORMAT.NUMBER("*-0") stores 0A 2D (fill with "-"); "\-0" stores the literal 02 2D (spec §7).
        var formats = Load("Fill Format").NumberFormats;
        Assert.Equal([0x03, 0x01, 0x00, 0x00, 0x0A, 0x2D, 0x06, 0x01], formats[0].Body);
        Assert.Equal([0x03, 0x01, 0x00, 0x00, 0x02, 0x2D, 0x06, 0x01], formats[1].Body);
    }

    [Fact]
    public void TableDeletedInput_ShowsRef()
    {
        // Excel 1.00: Y7, the row input of TABLE(Y7,X5), deleted with EDIT.DELETE(2). GET.FORMULA then gives
        // =TABLE(#REF!,R5C24) (spec §12.6).
        var document = Load("Table Deleted Input");
        var formula = document.Cells["B2"].Formula!;
        Assert.True(formula.RowInputDeleted);
        Assert.False(formula.ColumnInputDeleted);
        Assert.Equal("{=TABLE(#REF!,$X$5)}", formula.ToText(document));
    }

    [Fact]
    public void LinkOpen_EntryStillStartsWithZeros()
    {
        // Excel 1.00 saved this with Src, the linked sheet, open (spec §8).
        var bytes = Sample("Link Open");
        var document = ExcelDocument.Read(bytes);
        Assert.Equal(["Src"], document.ExternalDocuments.Select(d => d.Name));
        int name = bytes.AsSpan().IndexOf("\u0003Src"u8);
        Assert.True(name >= 6);
        Assert.Equal(new byte[6], bytes[(name - 6)..name]);
    }

    [Fact]
    public void OverlayAxis_ShowingAnOverlayAxisSetsItsLinePattern()
    {
        // Excel 1.00: Chart > Axes, "Overlay Chart Axis" under Category Axes, on the Combination sample. The overlay's
        // category axis line pattern becomes 1; its option flags do not change (spec §16.5).
        var overlay = ChartDocument.Load(SamplePath("Overlay Axis")).OverlayChart!;
        Assert.Equal(1, overlay.CategoryAxis.AxisLinePattern);
        Assert.Equal(0, overlay.ValueAxis.AxisLinePattern);
        Assert.Equal(0x0F, overlay.Record[22]);
    }

    [Theory]
    [InlineData("PROFIT CHART")]
    [InlineData("CLOCK.CH")]
    public void ChartLayout_DrawingRectangleIsCentredInTheWindow(string name)
    {
        // The margins either side of the chart area add up to the window's drawing size (spec §16.5, §16.6).
        var document = ChartDocument.Load(SamplePath(name));
        var chart = document.MainChart;
        var (top, left) = chart.DrawingMargins;
        Assert.InRange(document.ChartHeight - (2 * top + chart.ChartAreaBox.Height), -1, 1);
        Assert.InRange(document.ChartWidth - (2 * left + chart.ChartAreaBox.Width), -1, 1);
    }

    [Fact]
    public void ChartLayout_ValueAxisLineIsAtThePlotAreasLeftEdge()
    {
        // PROFIT CHART's value axis crosses before the first category (spec §16.5).
        var chart = ChartDocument.Load(SamplePath("PROFIT CHART")).MainChart;
        Assert.Equal(chart.PlotAreaBox.Left, chart.ValueAxis.LineAt);
    }

    [Fact]
    public void PointFormats_SingleSeriesChartsKeyThemByPoint()
    {
        // Pie format 3 pulls out the first slice: the override's series byte holds point 0 (spec §16.8).
        var chart = ChartDocument.Load(SamplePath("Pie Exploded")).MainChart;
        Assert.Equal(0, chart.GetPointFormat(0).SeriesNumber);
    }
}
