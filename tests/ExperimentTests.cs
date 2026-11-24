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
    public void Errors_UseTheErrorCodes(string address, string text, int code)
    {
        // Each formula entered with FORMULA and calculated by Excel 1.03 (spec §12.4).
        var result = Load("Errors").Cells[address].Formula!.Result;
        Assert.Equal(ExcelValueKind.Error, result.Kind);
        Assert.Equal((ExcelError)code, result.Error);
        Assert.Equal(text, result.Error.ToText());
        Assert.Equal(text, result.ToString());
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
        Assert.Equal(ExcelError.DivideByZero, document.Cells["B2"].Formula!.Result.Error);
        Assert.Equal(ExcelError.Number, document.Cells["B3"].Value.Error);
        Assert.Equal(ExcelError.NotAvailable, document.Cells["B4"].Value.Error);
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
    [InlineData("Calculation No Tables", ExcelCalculationMode.AutomaticExceptTables)]
    [InlineData("Calculation Manual", ExcelCalculationMode.Manual)]
    public void CalculationMode_Values(string name, ExcelCalculationMode mode)
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
        Assert.Equal(ExcelMacroKind.Command, command.MacroKind);
        Assert.Equal('q', command.ShortcutKey);
        Assert.Equal(0x44, command.RawMacroKind);   // the low nibble is junk (spec §11)
        var function = names.Single(n => n.Name == "Fn");
        Assert.Equal(ExcelMacroKind.Function, function.MacroKind);
        Assert.Equal(0x80, function.RawMacroKind);
        var plain = names.Single(n => n.Name == "Plain");
        Assert.Equal(ExcelMacroKind.None, plain.MacroKind);
        Assert.All([command, function, plain], n => Assert.True(n.Flags.HasFlag(ExcelNameFlags.Reference)));
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

    [Theory]
    [InlineData("PROFIT CHART", 0, 12)]
    [InlineData("Pie Exploded", 12, 12)]
    [InlineData("Point Markers", 12, 12)]
    [InlineData("Library Column", 3, 0)]
    public void ChartHeader_Byte5IsExcel15sCategoryCount(string sample, int byte4, int byte5)
    {
        // Excel 1.5 writes the category count as a word over header bytes 4-5; Excel 1.00 and 1.03 write only byte
        // 4, so charts that began as PROFIT CHART (saved by 1.5) keep its 12 in byte 5, and a chart the library made
        // and Excel 1.00 saved keeps the library's 0 (spec §16.5).
        var chart = ChartDocument.Load(SamplePath(sample)).MainChart;
        Assert.Equal(byte4, chart.CategoryCount);
        Assert.Equal(byte5, chart.Record[5]);
        Assert.Equal((byte4 << 8) | byte5, chart.Excel15CategoryCount);
    }

    [Theory]
    [InlineData("Legend", true)]
    [InlineData("Library Column", false)]
    public void ChartHeader_Bit15IsTheLegend(string sample, bool hasLegend)
    {
        // Chart > Add Legend in Excel 1.00 on a library column chart, then File > Save (spec §16.5).
        var chart = ChartDocument.Load(SamplePath(sample)).MainChart;
        Assert.Equal(hasLegend, chart.HasLegend);
        Assert.Equal(LegendPosition.Bottom, chart.LegendPosition);
    }

    [Theory]
    [InlineData("B4", "2")]          // DCOUNT(D1:D4,1,F1:F2)
    [InlineData("B5", "#VALUE!")]    // field 0
    [InlineData("B6", "2")]          // empty field: "not given", because the argument's type has bit 0x80
    [InlineData("B7", "#VALUE!")]    // field 2, past the database's one column
    public void DCountEmptyField_IsNotGiven(string address, string value)
    {
        // Re-entered with FORMULA by the harness in Excel 1.00 and saved (spec §15.3).
        Assert.Equal(value, Load("DCount Empty Field").Cells[address].ToString());
    }

    [Fact]
    public void Excel101_NewSheetHasTheExtraHeaderAndTranslatedNames()
    {
        // NEW(1) and SAVE.AS in the French Excel 1.01: a one-byte extra header holding 0, and a new sheet's five
        // built-in names from its strings (spec §5, §11).
        var document = Load("New Sheet French");
        Assert.Equal([0], document.ExtraHeader);
        Assert.Equal(208, document.FunctionCount);
        Assert.Equal(
            ["Base_de_données", "Critères", "Enregistrement", "Impression_des_titres", "Zone_d_impression"],
            document.Names.Select(n => n.Name));
    }

    [Fact]
    public void Excel101_AddsItsNamesToAFileAndKeepsTheTokens()
    {
        // A library sheet opened and saved by the French Excel 1.01: it adds its own five built-in names to the
        // English ones, and the formulas keep the library's tokens (spec §11, §15.4).
        var document = Load("Formulas French");
        Assert.Equal(10, document.Names.Count);
        Assert.Contains(document.Names, n => n.Name == "Critères");
        Assert.Contains(document.Names, n => n.Name == "Criteria");
        foreach (var address in new[] { "B4", "B5", "B6", "B9" })
        {
            var formula = document.Cells[address].Formula!;
            Assert.Equal(
                FormulaTokens.ClearStaleBytes(ExcelFormula.Parse(formula.ToText(document), document).Expression),
                FormulaTokens.ClearStaleBytes(formula.Expression));
        }
    }

    [Theory]
    [InlineData("A7", "=SOMME(L[-3]C[2],L[-2]C[2])")]
    [InlineData("A9", "=ARRONDI(L[-4]C[2],0)")]
    [InlineData("A10", "={1.2;3.4}")]
    [InlineData("A12", "=VRAI")]
    public void Excel101_GetFormulaIsFrench(string address, string text)
    {
        // GET.FORMULA in the French Excel 1.01 on an English system: French function names, L for row in R1C1, a
        // comma between arguments and a full stop between array columns (spec §15.4).
        Assert.Equal(text, Load("Driver French").Cells[address].Value);
    }

    [Theory]
    [InlineData(ExcelVersion.Excel100, 2)]
    [InlineData(ExcelVersion.Excel104, 3)]
    [InlineData(ExcelVersion.Excel150, 3)]
    public void Open_GainsItsThirdArgumentInExcel104(ExcelVersion version, int maxArguments)
    {
        // From the function tables of Excel 1.00, 1.04 and 1.5 (spec §2.1).
        Assert.Equal(maxArguments, ExcelFunctions.Get(106, version)!.MaxArguments);
        Assert.Equal(version >= ExcelVersion.Excel104, ExcelFunctions.Get(208, version) is not null);
    }

    [Fact]
    public void Excel104_SavesItsChanges()
    {
        // CHANGE.LINK and a three-argument OPEN, re-entered with FORMULA and saved by Excel 1.04 (spec §2.1, §2.3, §5).
        var bytes = Sample("Changes 1.04");
        var document = ExcelDocument.Read(bytes);
        Assert.Equal(ExcelVersion.Excel104, SavedBy(bytes));
        Assert.Equal(209, document.FunctionCount);               // one more than CHANGE.LINK (208)
        Assert.Equal([0], document.ExtraHeader);
        Assert.Equal([0x42, 0x02, 0xD0], document.Cells["B4"].Formula!.Expression[^3..].ToArray()); // CHANGE.LINK (208) with 2 arguments
        Assert.Equal([0x42, 0x03, 0x6A], document.Cells["B6"].Formula!.Expression[^3..].ToArray()); // OPEN with 3 arguments
    }

    [Fact]
    public void Excel104_WritesTheFunctionCountFromTheFormulas()
    {
        // The library wrote 209 with no function above 207; Excel 1.04 saved it back with 208 (spec §5).
        var bytes = Sample("No New Functions 1.04");
        Assert.Equal(ExcelVersion.Excel104, SavedBy(bytes));
        Assert.Equal(208, ExcelDocument.Read(bytes).FunctionCount);
    }

    [Theory]
    [InlineData("Changes 1.06", 209)]
    [InlineData("No New Functions 1.06", 208)]
    public void Excel106_SavesAs104Does(string sample, int functionCount)
    {
        // The same sheets as the 1.04 experiment, run in Excel 1.06: the same tokens, embedded code (import word
        // 0x740A, so a 1.06 file looks like a 1.04 one) and function-count rule (spec §2.1, §2.3, §5).
        var bytes = Sample(sample);
        var document = ExcelDocument.Read(bytes);
        Assert.Equal(ExcelVersion.Excel104, SavedBy(bytes));
        Assert.Equal(functionCount, document.FunctionCount);
        Assert.Equal([0], document.ExtraHeader);
        var original = Load(sample.Replace("1.06", "1.04"));
        foreach (var entry in original.Cells.NonEmptyCells)
        {
            int row = entry.Row;
            int column = entry.Column;
            var cell = entry.GetCell();
            if (cell.Formula is { } formula)
            {
                Assert.Equal(
                    FormulaTokens.ClearStaleBytes(formula.Expression),
                    FormulaTokens.ClearStaleBytes(document.Cells[row, column].Formula!.Expression));
            }
        }
    }

    [Theory]
    [InlineData("Chart 102 Categories", 12, 102)]
    [InlineData("Chart 103 Categories", 1, 103)]
    public void Excel15_MarksChartsWithMoreThan102CategoriesAsVersion1(string sample, int version, int categories)
    {
        // Library line charts with 102 and 103 points, opened and saved by Excel 1.5. It writes version 1, which
        // releases before 1.5 refuse, when a chart has more than 102 categories (spec §16.2).
        var document = ChartDocument.Load(SamplePath(sample));
        Assert.Equal(version, document.Version);
        Assert.Equal(categories, document.MainChart.Excel15CategoryCount);
        Assert.Equal(File.ReadAllBytes(SamplePath(sample)), document.ToArray());
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
        Assert.True(ChartDocument.Load(SamplePath("Line Log Scale")).MainChart.ValueScale.IsLogarithmic);
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
    [InlineData("Format Text Section")]
    [InlineData("Format Scaling")]
    public void Excel15Features_SaveAsVersion18(string name)
    {
        // Excel 1.5 saves version 18 for colour or @ number formats, a text section, trailing-comma scaling,
        // gridline colours and frozen panes (spec §2.1).
        Assert.Equal(ExcelFormatVersion.Version18, Load(name).Version);
    }

    [Theory]
    [InlineData("Format Two Sections")]
    [InlineData("SIMUL EQUATIONS Saved")]
    public void Excel15_SavesVersion17WhenNothingNeedsIt(string name)
    {
        // FORMAT.NUMBER("0.0;-0.0"), and Microsoft's version-18 SIMUL EQUATIONS opened and saved in the released
        // Excel 1.5: neither has anything only Excel 1.5 reads, so both save as version 17 (spec §2.1).
        var document = Load(name);
        Assert.Equal(ExcelFormatVersion.Version17, document.Version);
        Assert.Equal(ExcelWindow.StateSize, document.Windows[0].State.Length);
    }

    [Fact]
    public void Excel15_TextSectionFollowsTheOtherThree()
    {
        // FORMAT.NUMBER("0;-0;0;""t"""): the fourth section is the quoted string, 09 01 "t" (spec §7).
        var format = Assert.Single(Load("Format Text Section").NumberFormats);
        Assert.Equal(0x14, format.Descriptor[3]);
        Assert.Equal(3, format.Descriptor[7]);
        Assert.Equal([0x09, 0x01, (byte)'t'], format.Body[0x14..]);
    }

    [Fact]
    public void Excel15_TrailingCommaScalesByAThousand()
    {
        // FORMAT.NUMBER("#,##0,"): the header's third byte is c = b - 3 for the one trailing comma (spec §7).
        var format = Assert.Single(Load("Format Scaling").NumberFormats);
        Assert.Equal([0x03, 0x04, 0x00, 0xFD], format.Body[..4]);
    }

    [Fact]
    public void Excel15_FrozenPanesStoreTheSplitCell()
    {
        // SPLIT(1,2) then FREEZE.PANES(TRUE): frozen, with the split at row 2 and column 1, counting from 0 (spec §13).
        var window = Load("Frozen Panes").Windows[0];
        Assert.True(window.PanesFrozen);
        Assert.Equal((ushort)2, window.FrozenRow);
        Assert.Equal((ushort)1, window.FrozenColumn);
        Assert.Null(window.GridlineColor);
    }

    [Fact]
    public void Excel15_GridlineColour()
    {
        // DISPLAY(FALSE,TRUE,TRUE,TRUE,3): red (spec §13).
        var window = Load("Gridline Colour").Windows[0];
        Assert.Equal(ExcelColor.Red, window.GridlineColor);
        Assert.False(window.PanesFrozen);
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
        Assert.Equal(29220.0, Load("Literals").Cells["C4"].Value.Number);
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
        Assert.Equal(ChartPattern.Solid, overlay.CategoryAxis.Line.Pattern);
        Assert.Equal(ChartPattern.None, overlay.ValueAxis.Line.Pattern);
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
        Assert.InRange(document.ChartHeight - (2 * top + chart.ChartArea.Box.Height), -1, 1);
        Assert.InRange(document.ChartWidth - (2 * left + chart.ChartArea.Box.Width), -1, 1);
    }

    [Fact]
    public void ChartLayout_ValueAxisLineIsAtThePlotAreasLeftEdge()
    {
        // PROFIT CHART's value axis crosses before the first category (spec §16.5).
        var chart = ChartDocument.Load(SamplePath("PROFIT CHART")).MainChart;
        Assert.Equal(chart.PlotArea.Box.Left, chart.ValueAxis.LineAt);
    }

    [Fact]
    public void PointFormats_SingleSeriesChartsKeyThemByPoint()
    {
        // Pie format 3 pulls out the first slice: the override's series byte holds point 0 (spec §16.8).
        var chart = ChartDocument.Load(SamplePath("Pie Exploded")).MainChart;
        Assert.Equal(0, chart.GetPointFormat(0).SeriesOrPoint);
    }

    [Fact]
    public void Arrow_IsAFreeObjectWithTwoEnds()
    {
        // Excel 1.00: Chart > Add Arrow on PROFIT CHART, whose own fifth text object is already an arrow (spec §16.7).
        var texts = ChartDocument.Load(SamplePath("Arrow")).MainChart.Texts;
        Assert.Equal(6, texts.Count);
        Assert.All(texts.Take(4), t => Assert.False(t.IsArrow));
        Assert.All(texts.Skip(4), t =>
        {
            Assert.True(t.IsArrow);
            Assert.Equal(ChartTextKind.FreeText, t.Kind);
            Assert.True(t.IsAutomaticText);
            Assert.Null(t.Text);
        });
        Assert.True(ChartDocument.Load(SamplePath("PROFIT CHART")).MainChart.Texts[4].IsArrow);
    }

    [Theory]
    [InlineData("Library Column")]
    [InlineData("Library Line")]
    [InlineData("Library Pie")]
    [InlineData("Library Scatter")]
    public void LibraryCharts_SeriesHeaderMatchesExcels(string name)
    {
        // Charts built with ChartDocument.Create and ChartSeries.Create, which Excel 1.00 drew and saved. Excel's series
        // header matches the library's apart from bytes 22–25, which it copies from memory, and so does the default
        // series format (spec §16.8, §16.9).
        var saved = ChartDocument.Load(SamplePath(name)).Series[0];
        var built = ChartDocument.Create(ChartType.Line);
        built.Series.Add(ChartSeries.Create(saved.Name, saved.Values.Select(v => v.ToDouble()), saved.SourceDocument, saved.ValuesReference));
        var ours = ChartDocument.Read(built.ToArray()).Series[0];
        var a = ours.Header.ToArray();
        var b = saved.Header.ToArray();
        a.AsSpan(22, 4).Clear();
        b.AsSpan(22, 4).Clear();
        Assert.Equal(b, a);
        Assert.Equal(saved.FirstLabel, ours.FirstLabel);
        var excelChart = ChartDocument.Load(SamplePath(name)).MainChart;
        Assert.Equal(excelChart.Record[366..378], Chart.Create(excelChart.Type).Record[366..378]);
    }

    [Fact]
    public void TableRelinked_SavedLinkIsCleared()
    {
        // Excel 1.00 opened a table whose record had 00 01 00 at offset 6 (a list link) and saved it back with zeros.
        // Byte 14's low bits came back as 0E (spec §12.6).
        var formula = Load("Table Relinked").Cells["B6"].Formula!;
        var record = formula.Fixups.Slice(3, 18).ToArray();
        Assert.Equal(new byte[3], record[6..9]);
        Assert.Equal(0xCE, record[14]);
    }

    [Theory]
    [InlineData("Arrow", 4, ChartPattern.Solid, ChartLineWeight.Hairline, ArrowHeadWidth.Medium, ArrowHeadLength.Short, ArrowHeadStyle.Filled)]
    [InlineData("Arrow Patterns", 5, ChartPattern.DarkGray, ChartLineWeight.Medium, ArrowHeadWidth.Narrow, ArrowHeadLength.Long, ArrowHeadStyle.Open)]
    [InlineData("Arrow Heads", 5, ChartPattern.DarkGray, ChartLineWeight.Medium, ArrowHeadWidth.Narrow, ArrowHeadLength.None, ArrowHeadStyle.None)]
    public void Arrows_PatternWeightAndHead(string name, int index, ChartPattern pattern, ChartLineWeight weight, ArrowHeadWidth width, ArrowHeadLength length, ArrowHeadStyle style)
    {
        // PROFIT CHART's arrow, and an arrow set in Format > Patterns in Excel 1.00: long and open, then no head
        // (spec §16.7).
        var arrow = ChartDocument.Load(SamplePath(name)).MainChart.Texts[index];
        Assert.True(arrow.IsArrow);
        Assert.Equal(pattern, arrow.ArrowLine.Pattern);
        Assert.Equal(weight, arrow.ArrowLine.Weight);
        Assert.Equal(width, arrow.HeadWidth);
        Assert.Equal(length, arrow.HeadLength);
        Assert.Equal(style, arrow.HeadStyle);
    }

    [Fact]
    public void PointMarkers_OverridesAreKeyedByPoint()
    {
        // Excel 1.00: Format > Patterns on two markers of PROFIT CHART's single series, choosing X. Each override
        // record holds the point number and marker 7; the series keeps its hollow diamonds (spec §16.8).
        var chart = ChartDocument.Load(SamplePath("Point Markers")).MainChart;
        Assert.Equal(ChartMarker.Diamond, chart.DefaultSeriesFormat.Marker);
        Assert.Equal([8, 10], Enumerable.Range(0, chart.PointFormats.Count).Select(i => (int)chart.GetPointFormat(i).SeriesOrPoint));
        Assert.All(Enumerable.Range(0, chart.PointFormats.Count), i => Assert.Equal(ChartMarker.X, chart.GetPointFormat(i).Marker));
    }

    [Fact]
    public void ColumnFill_OverrideHoldsBorderAndArea()
    {
        // Excel 1.00: Format > Patterns on the first series of a column chart: vertical stripes, a grey thick border
        // (spec §16.8). Bytes 5, 9 and 10 stay 0.
        var chart = ChartDocument.Load(SamplePath("Column Fill")).MainChart;
        var format = chart.GetPointFormat(0);
        Assert.Equal(0, format.SeriesOrPoint);
        Assert.Equal(ChartPattern.VerticalStripes, format.AreaPattern);
        Assert.False(format.AutomaticArea);
        Assert.Equal(ChartPattern.Gray, format.Line.Pattern);
        Assert.Equal(ChartLineWeight.Thick, format.Line.Weight);
        Assert.False(format.AutomaticLine);
        Assert.Equal(new byte[3], new[] { chart.PointFormats[0][5], chart.PointFormats[0][9], chart.PointFormats[0][10] });
    }

    [Fact]
    public void EmptyArguments_CompileEverywhereButSeriesIsRefused()
    {
        // Excel 1.00 macro sheet, formulas entered with FORMULA (spec §15.3). An empty argument compiles to 16 whether
        // or not the function's type byte has 0x80; SERIES (flag 0x02) is refused, leaving C6 empty.
        var document = Load("Empty Arguments");
        Assert.Equal([0x1E, 0x00, 0x01, 0x16, 0x41, 0x1B], document.Cells["C2"].Formula!.Expression.ToArray());
        Assert.Equal(1, document.Cells["C2"].Value.Number);                    // ROUND(1,) is ROUND(1,0)
        Assert.Equal(ExcelError.Value, document.Cells["C3"].Value.Error);      // MID("abc",,1): start 0
        Assert.Equal(ExcelCellType.Empty, document.Cells["C6"].Type);
    }

    [Fact]
    public void Excel15_FunctionsFrom256UseTheHighTokens()
    {
        // Every formula of functions 256–291 from tests/Data/formulas-1.5.txt, re-entered by Excel 1.5: LEFT (256)
        // with one argument is 3A (value class 5A), 1, 00, and the count is the highest function used plus 1 (spec §5, §15.2).
        var document = Load("Functions 256 And Up");
        Assert.Equal([0x44, 0xC0, 0x00, 0x00, 0x5A, 0x01, 0x00], document.Cells["B4"].Formula!.Expression.ToArray());
        Assert.Equal(292, document.FunctionCount);
        Assert.Empty(document.ExtraFunctionNames);
    }

    [Fact]
    public void Excel15_SavesDataDeleteDialogInTheFixedForm()
    {
        // formulas-1.5 sheet 5 re-entered in Excel 1.5: DATA.DELETE?() comes back as 41 82 (spec §15.4).
        Assert.Equal([0x41, 0x82], Load("Formulas 1.5 Sheet 5").Cells["B38"].Formula!.Expression.ToArray());
    }

    [Fact]
    public void Excel15_ArrayFormulaOperandsTakeArrayClassOnlyInsideArrayArguments()
    {
        // formulas-1.5 sheet 2 re-entered in Excel 1.5 with FORMULA.ARRAY (spec §15.4).
        var document = Load("Formulas 1.5 Sheet 2");
        Assert.Equal(0x45, document.Cells["B60"].Formula!.SharedExpression[0]);   // {=A1:A2*2}
        Assert.Equal(0x65, document.Cells["B64"].Formula!.SharedExpression[0]);   // {=SUM(A1:A2*B1:B2)}
    }

    [Fact]
    public void ArrayFormulaOperands_TakeValueClassAtTheTop()
    {
        // Excel 1.00 re-entering array formulas with FORMULA.ARRAY: an operator at the top of an array formula
        // leaves its operands in value class, as Excel 1.5 does (spec §15.4).
        var document = Load("Array Formula Classes");
        Assert.Equal(0x45, document.Cells["B4"].Formula!.SharedExpression[0]);    // {=A1:A2*2}
        Assert.Equal(0x45, document.Cells["B10"].Formula!.SharedExpression[0]);   // {=A1:B1+A2:B2}
        Assert.Equal(0x45, document.Cells["B12"].Formula!.SharedExpression[0]);   // {=-A1:A2}
    }

    [Theory]
    [InlineData("B8", 0x41)]    // =IF(A1,NOW()): NOW in value class
    [InlineData("B21", 0x41)]   // =IF(ISERROR(A1),1), typed: ISERROR in value class
    [InlineData("B22", 0x41)]   // =IF(A1,GOTO(B1)), typed: GOTO in value class
    public void IfArguments_FollowTheArgumentRules(string address, byte expected)
    {
        // Excel 1.00 compiling through FORMULA and by typing (spec §15.4).
        var expression = Load("If Classes").Cells[address].Formula!.Expression.ToArray();
        Assert.Contains(expected, expression.Where(b => (b & 0x9F) is 0x01 or 0x02));
        Assert.DoesNotContain(expression, b => b is 0x21 or 0x22);
    }

    // Every number format the library knows, by the text Excel's TEXT function was given for it.
    private static Dictionary<string, ExcelNumberFormat> FormatsByText()
    {
        var formats = new Dictionary<string, ExcelNumberFormat>();
        foreach (var format in ExcelNumberFormat.BuiltIn.Concat(TestData.AllSamples()
            .Select(s => File.ReadAllBytes(TestData.SampleFile(s))).Where(b => b[0] == 0x53)
            .SelectMany(b => ExcelDocument.Read(b).NumberFormats)))
        {
            formats.TryAdd(format.ToText(), format);
        }

        // The date-limit run's format, which no sample stores: m/d/yy h:mm:ss in tokens.
        var dateTime = new ExcelNumberFormat(
            [0, 0, 0, 16, 16, 16, 1], Convert.FromHexString("0F022F0B022F190220130 23A16023A18".Replace(" ", "")));
        formats.TryAdd(dateTime.ToText(), dateTime);
        return formats;
    }

    [Theory]
    [InlineData("TEXT Formats 1")]
    [InlineData("TEXT Formats 2")]
    [InlineData("TEXT Date Limits")]
    [InlineData("TEXT Formats 1.5")]
    public void NumberFormats_MatchExcelsText(string name)
    {
        // Each row: a format's text, a value and =TEXT(value,format), recalculated by Excel 1.00 or 1.5. TEXT gives
        // #VALUE! for General and fills, and for dates past 6 February 2040, where the library shows General.
        var formats = FormatsByText();
        var document = Load(name);
        int compared = 0;
        for (int row = 0; row < document.Cells.RowCount; row++)
        {
            if (document.Cells.Find(row, 2)?.Formula?.Result is not { Kind: ExcelValueKind.Text } result)
            {
                continue;
            }

            var format = formats[document.Cells[row, 0].TextValue!];
            Assert.Equal(result.Text, format.Format(document.Cells[row, 1].NumberValue));
            compared++;
        }

        Assert.NotEqual(0, compared);
    }

    [Theory]
    [InlineData("General As Text")]
    [InlineData("General As Text 2")]
    [InlineData("General As Text 1.5")]
    [InlineData("General As Text 2 1.5")]
    public void General_MatchesExcelsNumberToText(string name)
    {
        // =""&B1 in Excel 1.00 and 1.5: 14 significant digits, written in full from 1E-18 to below 1E+20 (spec §7). The exact
        // half 99999999999999.5 is the one value Excel rounds down.
        var document = Load(name);
        for (int row = 0; row < document.Cells.RowCount; row++)
        {
            if (document.Cells.Find(row, 2)?.Formula is not { } formula || document.Cells[row, 1].NumberValue == 99999999999999.5)
            {
                continue;
            }

            Assert.Equal(formula.Result.Text, ExcelValueFormatter.FormatGeneral(document.Cells[row, 1].NumberValue));
        }
    }
    [Theory]
    [InlineData("Constant Cells")]
    [InlineData("Constant Cells 1.5")]
    public void ConstantCells_TrueFalseAndErrorsAreType4(string name)
    {
        // Excel 1.00 and 1.5: FORMULA("TRUE") in D1 and FORMULA("#REF!") in D2 (spec §9). Word 0 holds a kind byte,
        // 0 for TRUE or FALSE and 1 for an error, then the value or the error code, like BIFF2's BOOLERR.
        var document = Load(name);
        var d1 = document.Cells["D1"];
        Assert.Equal(ExcelCellType.BooleanOrError, d1.Type);
        Assert.False(d1.IsError);
        Assert.True(d1.BooleanValue);
        Assert.Equal(true, d1.Value);
        Assert.Equal("TRUE", ExcelValueFormatter.FormatCell(d1, document));

        var d2 = document.Cells["D2"];
        Assert.True(d2.IsError);
        Assert.Equal(ExcelError.Reference, d2.Error);
        Assert.Equal(ExcelError.Reference, d2.Value.Error);
        Assert.Equal("#REF!", d2.ToString());
        Assert.Equal("#REF!", ExcelValueFormatter.FormatCell(d2, document));
    }

    [Theory]
    [InlineData("Constant Cells")]
    [InlineData("Constant Cells 1.5")]
    public void ConstantCells_IntegersAreUnsigned(string name)
    {
        // Typed 7, -7, 32767, 40000 and 0 in D3:D7: whole numbers from 0 to 65535 are integer cells, Excel shows
        // 40000 for D6, and a negative number is a number cell (spec §9).
        var document = Load(name);
        Assert.Equal(7, document.Cells["D3"].IntegerValue);
        Assert.Equal(ExcelCellType.Number, document.Cells["D4"].Type);
        Assert.Equal(-7.0, document.Cells["D4"].NumberValue);
        Assert.Equal(32767, document.Cells["D5"].IntegerValue);
        Assert.Equal(40000, document.Cells["D6"].IntegerValue);
        Assert.True(document.Cells["D6"].TryGetNumber(out double d6));
        Assert.Equal(40000.0, d6);
        Assert.Equal(0, document.Cells["D7"].IntegerValue);

        var cell = new ExcelCell();
        cell.SetValue(-7);
        Assert.Equal(ExcelCellType.Number, cell.Type);
        cell.SetValue(40000);
        Assert.Equal(ExcelCellType.Integer, cell.Type);
    }

    [Fact]
    public void ArrayConstants_HoldTrueFalseAndErrors()
    {
        // Elements of type 4 (TRUE, FALSE) and 16 (an error) hold the value in their second byte, as BIFF's types 4
        // and 16 do (spec §12.5).
        var worksheet = Load("Constant Cells");
        var e1 = worksheet.Cells["E1"].Formula!;
        Assert.Equal("={1,FALSE;#DIV/0!,\"b\"}", e1.ToText(worksheet));
        var parsed = ExcelFormula.Parse("={1,FALSE;#DIV/0!,\"b\"}", worksheet);
        Assert.Equal(e1.Expression[..4].ToArray(), parsed.Expression[..4].ToArray());   // the rest is left over from memory
        Assert.Equal(e1.Fixups.ToArray(), parsed.Fixups.ToArray());
        Assert.Equal(true, worksheet.Cells["E2"].Value);

        var macros = Load("Array Constant Types");
        Assert.Equal("={TRUE,#N/A,\"a\",1}", macros.Cells["D1"].Formula!.ToText(macros));
        Assert.Equal("{FALSE,#REF!}", macros.Names.Single(n => n.Name == "ArrName").GetDefinitionText(macros));
    }

    [Theory]
    [InlineData("Relative Names")]
    [InlineData("Relative Names 1.5")]
    public void RelativeNames_HoldOffsetsFromTheCellThatUsesThem(string name)
    {
        // With D5 active: DEFINE.NAME("Up3","=R[-3]C[-2]"), "UpArea" R[-3]C[-2]:R[-2]C[-2], "Up1Abs" R[-1]C2 and
        // "Left" R5C[-2]. A relative row or column is stored as an offset (spec §11), so =Up3 in D5 is B2 (11) and
        // in D6 is B3 (22); =Up1Abs in F5 is B4 (33) and in F6 is B5 (44).
        var document = Load(name);
        var d5 = new ExcelCellAddress(4, 3);
        var names = document.Names.ToDictionary(n => n.Name);
        Assert.Equal([0x2C, 0xFF, 0xFD, 0xFE], names["Up3"].Definition);
        Assert.Equal("B2", names["Up3"].GetDefinitionText(document, d5));
        Assert.Equal("IU16382", names["Up3"].GetDefinitionText(document));
        Assert.Equal("B2:B3", names["UpArea"].GetDefinitionText(document, d5));
        Assert.Equal("$B4", names["Up1Abs"].GetDefinitionText(document, d5));
        Assert.Equal("B$5", names["Left"].GetDefinitionText(document, d5));
        Assert.Equal(11.0, document.Cells["D5"].Formula!.Result.Number);
        Assert.Equal(22.0, document.Cells["D6"].Formula!.Result.Number);
        Assert.Equal(33.0, document.Cells["F5"].Formula!.Result.Number);
        Assert.Equal(44.0, document.Cells["F6"].Formula!.Result.Number);

        foreach (var (defined, text) in new[] { ("Up3", "=B2"), ("UpArea", "=B2:B3"), ("Up1Abs", "=$B4"), ("Left", "=B$5") })
        {
            var copy = new ExcelDefinedName(defined);
            copy.SetDefinition(text, d5, document);
            Assert.Equal(names[defined].Definition, copy.Definition);
        }
    }

    [Fact]
    public void RecalculationOrder_FollowsDependencies()
    {
        // Every formula marked waiting (offset 8 = FF FF) with a stored 0, opened by Excel 1.00: one recalculation
        // gets chains right whichever way they run (docs/Calculation.md).
        var document = Load("Recalculation Order");
        for (int row = 1; row <= 9; row++)
        {
            Assert.Equal(11 - row, document.Cells[$"A{row}"].Formula!.Result.Number);   // A1 = A2+1 … A10 = 1
            Assert.Equal(row + 1, document.Cells[$"B{row + 1}"].Formula!.Result.Number); // B1 = 1, B2 = B1+1 …
        }

        Assert.Equal(12, document.Cells["C1"].Formula!.Result.Number);  // =D1*2, D1 =E1*2, E1 3
        Assert.Equal(180, document.Cells["C3"].Formula!.Result.Number); // =SUM of four formulas
        Assert.Equal(15, document.Cells["C8"].Formula!.Result.Number);  // =E8+C9 with E8 and C9 formulas
    }

    [Fact]
    public void CircularReference_WithoutIteration_KeepsTheValues()
    {
        // A1 = B1+1, B1 = A1+1 and A3 = A1*10, all waiting with stored 0: none is calculated.
        var document = Load("Circular No Iteration");
        Assert.All(["A1", "B1", "A3"], a => Assert.Equal(0, document.Cells[a].Formula!.Result.Number));
    }

    [Theory]
    [InlineData("Iteration Count", 100, 0.001, 100)]                // A1 = A1+1: exactly the maximum iterations
    [InlineData("Iteration Five", 5, 0.001, 5)]
    [InlineData("Iteration Converge", 100, 0.001, 1.9990234375)]    // A1 = A1/2+1: stops after 11, change < 0.001
    [InlineData("Iteration Big Change", 100, 0.5, 1.75)]            // a change of exactly 0.5 does not stop it
    public void Iteration_StopsAsExcelDoes(string sample, int maxIterations, double maxChange, double a1)
    {
        // CALCULATION(1,TRUE,max,change), then FORMULA in A1 (docs/Calculation.md).
        var document = Load(sample);
        Assert.True(document.State.Iteration);
        Assert.Equal(maxIterations, document.State.MaxIterations);
        Assert.Equal(maxChange, document.State.MaxChange);
        Assert.Equal(a1, document.Cells["A1"].Formula!.Result.Number);
    }

    [Fact]
    public void Iteration_CalculatesTheEnteredCellFirst()
    {
        // B1 = A1+1 already there, A1 = B1+1 entered: each of 100 iterations does A1, then B1.
        var document = Load("Iteration Pair");
        Assert.Equal(199, document.Cells["A1"].Formula!.Result.Number);
        Assert.Equal(200, document.Cells["B1"].Formula!.Result.Number);
    }

    [Fact]
    public void CalculationSettings_AreExcelsNotTheFiles()
    {
        // Built with iteration on and 5 iterations, opened while another document was open:
        // Open drops the file's settings and Save writes Excel's (spec §6).
        var document = Load("Iteration Settings Kept");
        Assert.False(document.State.Iteration);
        Assert.Equal(100, document.State.MaxIterations);
        Assert.Equal(0.001, document.State.MaxChange);
        Assert.Equal(0, document.Cells["A1"].Formula!.Result.Number);
    }

    // An empty worksheet cut down as in the smallest-file experiment: no names, no window record, no page breaks or
    // print record, so the sections after the cell data are just their terminators (docs/Testing.md).
    private static byte[] MinimalBare()
    {
        byte[] full = ExcelDocument.CreateWorksheet().ToArray();
        int names = ExcelFileLayout.Read(full).GetSectionOffset(ExcelSection.DefinedNames);
        return [.. full.AsSpan(0, names), 0x00, 0x00, 0x00, 0x00, 0x00];
    }

    [Fact]
    public void Minimal_SmallestFileEveryReleaseOpens()
    {
        // 701 bytes: Excel 1.00, 1.03 and 1.5 all open it; cut any shorter and they refuse it.
        byte[] bare = MinimalBare();
        Assert.Equal(701, bare.Length);
        var document = ExcelDocument.Read(bare);
        Assert.Empty(document.Names);
        Assert.Empty(document.Windows);
    }

    [Theory]
    [InlineData("Excel 1.00/Minimal Bare")]
    [InlineData("Excel 1.03/Minimal Bare")]
    [InlineData("Excel 1.5/Minimal Bare")]
    [InlineData("Excel 1.00/Minimal No Window")]
    public void Minimal_ExcelAddsItsNamesButNoWindow(string sample)
    {
        // Saved back, the sheet has Excel's five built-in names again, undefined, and still no window: a document
        // opened without a window record has none.
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(sample)));
        Assert.Equal(["Criteria", "Database", "Print_Area", "Print_Titles", "Recorder"], document.Names.Select(n => n.Name));
        Assert.Empty(document.Windows);
    }

    [Fact]
    public void Minimal_MissingThePrintSection_IsRefused()
    {
        // Excel 1.00 and 1.5 refuse a file that stops before the 3 bytes of section 7; Excel 1.03 crashes on it.
        byte[] cut = ExcelDocument.CreateWorksheet().ToArray()[..^3];
        Assert.ThrowsAny<Exception>(() => ExcelDocument.Read(cut));
    }

    [Fact]
    public void FunctionNameList_WrittenByABuildWithASmallerBaseCount()
    {
        // Saved by a copy of Excel 1.00 patched to a base count of 200, after FORMULA entered =FULL() (function
        // 202): the count is the table length, 208, and section 7 starts with the names of functions 200–207 as
        // Pascal strings before the page breaks (spec §5, §14). Released 1.00 never writes the list, since its base
        // count is 208 too, and a reader cannot tell from the file that it is there.
        byte[] file = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "Function Names 1.00 Patched"));
        var layout = ExcelFileLayout.Read(file);
        Assert.Equal(208, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(file.AsSpan(layout.DataOffset)));
        int position = layout.GetSectionOffset(ExcelSection.PrintSettings);
        var names = new List<string>();
        for (int i = 0; i < 8; i++)
        {
            int length = file[position];
            names.Add(MacRoman.GetString(file.AsSpan(position + 1, length)));
            position += 1 + length;
        }

        Assert.Equal(["MOVE", "SIZE", "FULL", "8", "9", "CANCEL.COPY", "UNLOCKED.NEXT", "UNLOCKED.PREV"], names);
        Assert.Equal([0, 0, 0], file[position..]);
        Assert.Equal(["MOVE", "SIZE", "FULL"], Enumerable.Range(200, 3).Select(n => ExcelFunctions.Get(n)!.Name));
    }

    [Fact]
    public void ChartPointFormatListHead_IsZeroWhenThereAreNoPointFormats()
    {
        // Record offset 362 heads Excel's list of data-point formats. With junk there and no formats, Excel 1.00 kept
        // it, lost the chart's series and saved garbage records (docs/Testing.md), so a writer must clear it.
        var chart = ChartDocument.Read(Sample("Column Fill"));
        int record = System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(Sample("Column Fill").AsSpan(20)) + 8;
        Assert.NotEqual(0, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(Sample("Column Fill").AsSpan(record + 362)));
        chart.Charts[0].PointFormats.Clear();
        byte[] written = chart.ToArray();
        Assert.Equal(0, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(written.AsSpan(record + 362)));
        Assert.Equal(0, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(written.AsSpan(record + 364)));
    }

    [Fact]
    public void ChartPointFormatListHead_JunkLosesTheSeries()
    {
        // Area Drop Lines (no data-point formats) with A5 A5 at record offset 362, opened and saved by Excel 1.00:
        // Open kept the word as the head of its list of point formats, and Save wrote the chart with its series'
        // name and references emptied. With 0 there, as in the original, the series survives (spec §16.5).
        var original = ChartDocument.Read(Sample("Area Drop Lines"));
        var saved = ChartDocument.Read(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "Point Format Head Junk 1.00")));
        Assert.Equal("Margin", original.Series[0].Name);
        Assert.Equal("", saved.Series[0].Name);
        Assert.Equal("", saved.Series[0].ValuesReference);
    }

    [Fact]
    public void EdgeLastRow_KeepsAMarginRowPastTheSheet()
    {
        // A value in A16384 (and A1): Excel 1.00 saved a 16385-row table, its blank margin past the sheet's last row
        // (spec §6.1). IV1, IV16384 and an array at the corner ran it out of memory and were not kept.
        var document = Load("Edge Last Row");
        Assert.Equal(16385, document.Cells.RowCount);
        Assert.Equal(1, document.Cells["A16384"].Value.Number);
        Assert.Equal(ExcelCellType.Empty, document.Cells.GetCellType(16384, 0));
        Assert.Equal(document.ToArray(), Sample("Edge Last Row"));
    }

    [Fact]
    public void ThreeWindows_AreSavedFrontToBack()
    {
        // NEW.WINDOW twice, each with its own SELECT, MOVE and SIZE: the newest window is first.
        var document = Load("Three Windows");
        Assert.Equal(["B10:D12", "C5", "A1"], document.Windows.Select(w => w.Selection[0].ToString()));
        Assert.Equal([(107, 259, 269, 479), (77, 39, 209, 239), (61, 9, 335, 502)],
            document.Windows.Select(w => ((int)w.Top, (int)w.Left, (int)w.Bottom, (int)w.Right)));
    }

    [Fact]
    public void ChartPageSetup_HoldsMarginsHeaderAndFooter()
    {
        // PAGE.SETUP("Top","Bottom",0.5,0.6,0.7,0.8) on a chart (left, right, top, bottom): the record stores the
        // margins top, bottom, left, right as decimal numbers, then the header and footer (spec §16.6).
        var chart = ChartDocument.Read(Sample("Chart Page Setup"));
        byte[] setup = chart.PageSetup!;
        double[] margins = [.. Enumerable.Range(0, 4).Select(i => MicrosoftDecimal.FromBytes(setup.AsSpan(120 + 8 * i, 8)).ToDouble())];
        Assert.Equal([0.7, 0.8, 0.5, 0.6], margins);
        Assert.Equal("Top", MacRoman.GetString(setup.AsSpan(153, setup[152])));
        Assert.Equal("Bottom", MacRoman.GetString(setup.AsSpan(409, setup[408])));
    }

    [Fact]
    public void Workspace_ListsEveryWindowOfEachDocument()
    {
        // Excel 1.00 quit with Driver, a chart, a sheet with three windows and another sheet open.
        var workspace = ExcelWorkspace.Read(Sample("Resume Excel 3"));
        Assert.Equal(6, workspace.Windows.Count);
        Assert.Equal(3, workspace.Windows.Count(w => w.DocumentName.EndsWith("Windows saved", StringComparison.Ordinal)));
    }

    [Theory]
    [InlineData("Resume Excel 1.03", ExcelDocumentKind.Worksheet, "Excel Program:Sample Documents:Expenses", "Excel Program:", 23)]
    [InlineData("Resume Excel 1.5", ExcelDocumentKind.MacroSheet, "Excel 1.5 Test:Formulas 1.5", "Excel 1.5 Test:", 0)]
    public void QuitWithADocumentOpen_WritesTheSameWorkspaceAs100(string name, ExcelDocumentKind kind, string document, string folder, int directory)
    {
        // Excel 1.03 (to its own disk) and 1.5 (to the startup disk's System Folder), quit with one document
        // open and the disks flushed by Shut Down (spec Appendix B).
        var bytes = Sample(name);
        var workspace = ExcelWorkspace.Read(bytes);
        var window = Assert.Single(workspace.Windows);
        Assert.Equal((kind, document, folder, directory), (window.Kind, window.DocumentName, window.Folder, (int)window.Directory));
        Assert.Equal(bytes, workspace.ToArray());
    }

    [Fact]
    public void EdgeColumn_KeepsAMarginColumnPastTheSheet()
    {
        // Values in IV1 and IV2 and an array in IU3:IV4: Excel 1.00 saved a 257-column table (spec §6.1).
        var document = Load("Edge Column");
        Assert.Equal(257, document.Cells.ColumnCount);
        Assert.Equal(2, document.Cells["IV2"].Value.Number);
        Assert.True(document.Cells["IV4"].Formula!.IsArrayFormula);
    }

    [Fact]
    public void ValueInTheLastRowOrColumn_GetsTheMarginExcelWrites()
    {
        // Built by the library, a sheet with values in A16384 and IV1 gets the same margin row and column.
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A16384"].SetNumber(1);
        document.Cells["IV1"].SetNumber(2);
        var reread = ExcelDocument.Read(document.ToArray());
        Assert.Equal((16385, 257), (reread.Cells.RowCount, reread.Cells.ColumnCount));
    }

    [Fact]
    public void ThreeWindows_Excel103SavesThemAs100Does()
    {
        var document = Load("Three Windows 1.03");
        Assert.Equal(["B10:D12", "C5", "A1"], document.Windows.Select(w => w.Selection[0].ToString()));
    }

    [Fact]
    public void ThreeWindows_Excel15KeepsEachWindowsVersion18Settings()
    {
        // One window frozen after SPLIT(2,1), another with DISPLAY(..., 3) (red): version 18, three 129-byte states,
        // each with its own panes and colour (spec §13).
        var document = Load("Three Windows 1.5");
        Assert.Equal(ExcelFormatVersion.Version18, document.Version);
        Assert.All(document.Windows, w => Assert.Equal(129, w.State.Length));
        Assert.Equal([(0, 0, 2), (1, 1, 0), (0, 1, 0)], document.Windows.Select(w => ((int)w.State[121], (int)w.State[127], (int)w.State[128])));
        Assert.Equal((1, 2), (document.Windows[1].State[123] << 8 | document.Windows[1].State[124], document.Windows[1].State[125] << 8 | document.Windows[1].State[126]));
    }

    [Fact]
    public void NaNCells_ExcelKeepsThemAndShowsTheExponentField()
    {
        // A1–A4 were written by the library as NaN, +infinity, -infinity and -NaN, with formulas on each and the
        // sheet marked for recalculation. Excel kept the values, calculated with them as IEEE doubles, and turned
        // them into text as its decimal conversion reads the exponent field 0x7FF (spec §12.1).
        var document = Load("NaN Cells");
        Assert.True(double.IsNaN(document.Cells["A1"].NumberValue));
        Assert.Equal(double.PositiveInfinity, document.Cells["A2"].NumberValue);
        Assert.Equal(double.NegativeInfinity, document.Cells["A3"].NumberValue);
        Assert.Equal(
            ["-2.6965397022935E+308", "1.7976931348623E+308", "-1.7976931348623E+308", "-2.6965397022935E+308"],
            Enumerable.Range(1, 4).Select(r => document.Cells[$"B{r}"].Value.Text));
        Assert.All(Enumerable.Range(1, 4), r => Assert.Equal(document.Cells[$"B{r}"].Value.Text, ExcelValueFormatter.FormatGeneral(document.Cells[$"A{r}"].NumberValue)));
        Assert.All(Enumerable.Range(1, 4), r => Assert.True(document.Cells[$"C{r}"].Value.Boolean));       // =A1=A1
        Assert.All(Enumerable.Range(1, 4), r => Assert.False(document.Cells[$"E{r}"].Value.Boolean));      // ISERROR
        Assert.True(double.IsNaN(document.Cells["F1"].Value.Number));                                     // =A1+1
        Assert.Equal(double.PositiveInfinity, document.Cells["F2"].Value.Number);
        Assert.Equal([false, true, false, false], Enumerable.Range(1, 4).Select(r => document.Cells[$"G{r}"].Value.Boolean)); // >0

        // Excel never makes such a value itself, so the library will not write one.
        var cell = ExcelDocument.CreateWorksheet().Cells["A1"];
        Assert.Throws<ArgumentOutOfRangeException>(() => cell.SetNumber(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => cell.SetValue(double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelFormula.Parse("=1").SetResult(double.NegativeInfinity));
    }

    [Fact]
    public void HugeConstants_OverflowIsNumErrorAndTooLargeConstantsAreRefused()
    {
        // Entered with FORMULA: =1E300*1E300 gives #NUM!, the constant 1E400 is taken as text, and =1E400 (A5) is
        // refused. =1E308*10 (A2) was refused too.
        var document = Load("Huge Constants");
        Assert.Equal(ExcelError.Number, document.Cells["A3"].Value.Error);
        Assert.Equal("=1E+300*1E+300", document.Cells["A3"].Formula!.ToText(document));
        Assert.Equal("1E400", document.Cells["A4"].TextValue);
        Assert.Equal(ExcelCellType.Empty, document.Cells.GetCellType(1, 0));
        Assert.Equal(ExcelCellType.Empty, document.Cells.GetCellType(4, 0));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=1E400"));
    }

    [Fact]
    public void NamesUnsorted_ExcelSearchesTheNamesInOrder()
    {
        // The library wrote Zeta and Alpha after the built-in names, out of order. Formulas holding their indices
        // worked, but typing =Alpha*10 or =Beta+Alpha+Zeta gave #NAME? while =Zeta*10 and =Aardvark+Zeta worked: Excel
        // finds a typed name by searching the order (spec §11). DEFINE.NAME put Aardvark and Beta in their places, and
        // the failed lookup added an undefined Alpha in its place; Excel left Zeta and Alpha where they were.
        var document = Load("Names Unsorted");
        Assert.Equal(
            ["Aardvark", "Alpha", "Beta", "Criteria", "Database", "Print_Area", "Print_Titles", "Recorder", "Zeta", "Alpha"],
            document.Names.Select(n => n.Name));
        Assert.Equal((1.0, 2.0), (document.Cells["B1"].Value.Number, document.Cells["B2"].Value.Number));
        Assert.Equal(ExcelError.Name, document.Cells["B3"].Value.Error);
        Assert.Equal(10, document.Cells["B4"].Value.Number);
        Assert.Equal(ExcelError.Name, document.Cells["B5"].Value.Error);
        Assert.Equal(8, document.Cells["B6"].Value.Number);

        // Saved unchanged, the names keep Excel's order; once they change, the library sorts them and renumbers the
        // formulas that use them.
        Assert.Equal(Sample("Names Unsorted"), document.ToArray());
        document.Names.RemoveAt(1);   // the undefined Alpha; B3 and B5 point at it
        document.Cells["B3"].Clear();
        document.Cells["B5"].Clear();
        foreach (var cell in new[] { "B1", "B2", "B4", "B6" })
        {
            var formula = document.Cells[cell].Formula!;
            var words = formula.GetFixupWords().Select(w => w > 1 ? w - 1 : w).Select(w => (byte[])[(byte)(w >> 8), (byte)w]).SelectMany(b => b).ToArray();
            document.Cells[cell].SetFormula(new ExcelFormula(formula.Expression, words));
        }

        var sorted = ExcelDocument.Read(document.ToArray());
        Assert.Equal(["Aardvark", "Alpha", "Beta", "Criteria", "Database", "Print_Area", "Print_Titles", "Recorder", "Zeta"], sorted.Names.Select(n => n.Name));
        Assert.Equal(["=Zeta", "=Alpha", "=Zeta*10", "=Aardvark+Zeta"], new[] { "B1", "B2", "B4", "B6" }.Select(c => sorted.Cells[c].Formula!.ToText(sorted)));
    }

    [Fact]
    public void ManyDecimals_FormatsCompileAsTheLibraryDoesAndShowPlainDecimals()
    {
        // FORMAT.NUMBER gave column C formats of 15 to 200 decimal places. From 128 on, the header's b and c are 0x80
        // or more; Excel 1.00 still shows plain decimals, and TEXT agrees (spec §7).
        var document = Load("Many Decimals");
        for (int row = 0; row < 10; row++)
        {
            int decimals = document.Cells[row, 3].IntegerValue;
            string text = "0." + new string('0', decimals);
            var excel = document.GetNumberFormat(document.Cells[row, 2].NumberFormat)!;
            Assert.Equal(ExcelNumberFormat.Parse(text), excel);
            Assert.False(excel.RequiresVersion18);
            Assert.Equal(text, excel.ToText());
            Assert.Equal("1." + "5".PadRight(decimals, '0'), document.Cells[row, 1].Value.Text);
            Assert.Equal(document.Cells[row, 1].Value.Text, ExcelNumberFormat.Parse(text).Format(1.5));
        }

        Assert.Equal(ExcelFormatVersion.Version17, document.Version);
    }

    [Theory]
    [InlineData("Constant Limits")]
    [InlineData("Constant Range")]
    [InlineData("Constant Edges")]
    public void ConstantRange_ExcelTakesConstantsFrom1EMinus307ToBelow1E306(string name)
    {
        // Column B holds what FORMULA entered in column A. Excel 1.00 refused every formula with a constant outside
        // 1E-307 to below 1E+306, however it was written, and took such a constant on its own as text (spec §15.4).
        var document = Load(name);
        int checkedRows = 0;
        for (int row = 1; row < document.Cells.RowCount; row++)
        {
            if (document.Cells.Find(row, 1) is not { Type: ExcelCellType.Text } entry)
            {
                continue;
            }

            string text = entry.TextValue;
            var cell = document.Cells.Find(row, 0);
            if (text.StartsWith('='))
            {
                bool excelTookIt = cell?.Formula is not null;
                bool libraryTakesIt = true;
                try
                {
                    _ = ExcelFormula.Parse(text);
                }
                catch (ExcelFormulaSyntaxException)
                {
                    libraryTakesIt = false;
                }

                Assert.True(excelTookIt == libraryTakesIt, $"{text}: Excel {excelTookIt}, library {libraryTakesIt}");
            }
            else
            {
                double value = double.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
                bool inRange = value >= 1E-307 && value < 1E306;
                Assert.Equal(inRange ? ExcelCellType.Number : ExcelCellType.Text, cell!.Type);
            }

            checkedRows++;
        }

        Assert.True(checkedRows >= 10);
    }

    [Fact]
    public void NaNWide_GeneralShowsTheNumberAndOtherFormatsHashes()
    {
        // A NaN and an infinity in General, then in 0.00 and a date, in a column three times the standard width:
        // Excel 1.00 drew -2.6965397022935E+308 and 1.7976931348623E+308, then two rows of # (read off the screen).
        var document = Load("NaN Wide");
        Assert.Equal("-2.6965397022935E+308", ExcelValueFormatter.FormatCell(document.Cells["A1"], document, 0));
        Assert.Equal("1.7976931348623E+308", ExcelValueFormatter.FormatCell(document.Cells["A2"], document, 0));
        Assert.Matches("^#+$", ExcelValueFormatter.FormatCell(document.Cells["A3"], document, 0));
        Assert.Matches("^#+$", ExcelValueFormatter.FormatCell(document.Cells["A4"], document, 0));
    }
}
