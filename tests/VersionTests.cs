using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// Behaviour that differs between Excel 1.00–1.03, 1.04 and 1.5, and features first seen in the 1.00 and 1.5 samples.
/// </summary>
public class VersionTests
{
    [Fact]
    public void FunctionCatalogue_CoversEveryRelease()
    {
        Assert.Equal(208, ExcelFunctions.Count);
        Assert.Equal(291, ExcelFunctions.MaxNumber);
        Assert.Equal("CHANGE.LINK", ExcelFunctions.Get(208)!.Name);
        Assert.Equal(ExcelVersion.Excel104, ExcelFunctions.Get(208)!.IntroducedIn);
        Assert.Equal("GET.DOCUMENT", ExcelFunctions.Get(238)!.Name);
        Assert.Equal(ExcelVersion.Excel150, ExcelFunctions.Get(238)!.IntroducedIn);
        Assert.Null(ExcelFunctions.Get(233));
        Assert.Null(ExcelFunctions.Get(238, ExcelVersion.Excel100));
        Assert.Equal(-1, ExcelFunctions.FindNumber("GET.DOCUMENT", ExcelVersion.Excel104));
        Assert.Equal(238, ExcelFunctions.FindNumber("get.document"));
    }

    [Fact]
    public void Excel15_ChangedArgumentRules()
    {
        Assert.Equal(1, ExcelFunctions.Get(57, ExcelVersion.Excel100)!.MaxArguments);   // FORMULA
        Assert.Equal(2, ExcelFunctions.Get(57, ExcelVersion.Excel150)!.MaxArguments);
        Assert.Equal("4139", Convert.ToHexString(ExcelFormula.Parse("=FORMULA(A1)").Expression)[^4..]);

        // Excel 1.5 keeps the old fixed-count form when the old number of arguments is used, so 1.0x can read it.
        Assert.Equal("4139", Convert.ToHexString(ExcelFormula.Parse("=FORMULA(A1)", null, ExcelVersion.Excel150).Expression)[^4..]);
        Assert.Equal("420239", Convert.ToHexString(ExcelFormula.Parse("=FORMULA(A1,B1)", null, ExcelVersion.Excel150).Expression)[^6..]);
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=FORMULA(A1,B1)"));
    }

    [Fact]
    public void Excel15_WritesFixedCountFunctionsWithoutTheDialogBit()
    {
        // DATA.DELETE takes no arguments: 1.00 keeps DATA.DELETE?() as 42 80 82, Excel 1.5's Save writes 41 82
        // (47.19 in 1.5; formulas-1.5 sheet 5 re-entered in Excel 1.5).
        Assert.Equal("428082", Convert.ToHexString(ExcelFormula.Parse("=DATA.DELETE?()").Expression));
        Assert.Equal("4182", Convert.ToHexString(ExcelFormula.Parse("=DATA.DELETE?()", null, ExcelVersion.Excel150).Expression));
        Assert.Equal("428091", Convert.ToHexString(ExcelFormula.Parse("=RUN?()", null, ExcelVersion.Excel150).Expression));
    }

    [Fact]
    public void FunctionsInsideIf_ChangeClassInExcel15()
    {
        // GOTO inside IF: reference class (21) in 1.00–1.03, value class (41) in 1.5.
        Assert.Contains("2135", Convert.ToHexString(ExcelFormula.Parse("=IF(A1,GOTO(B1))").Expression));
        Assert.Contains("4135", Convert.ToHexString(ExcelFormula.Parse("=IF(A1,GOTO(B1))", null, ExcelVersion.Excel150).Expression));
    }

    [Fact]
    public void FunctionsAbove255_UseTokens39And3A()
    {
        // LEFT is file number 256: 0x3A (value class 0x5A), two arguments, 256 - 256. CELL is 291.
        var sheet = ExcelDocument.CreateWorksheet();
        var left = ExcelFormula.Parse("=LEFT(\"abc\",2)", sheet, ExcelVersion.Excel150);
        Assert.EndsWith("5A0200", Convert.ToHexString(left.Expression));
        Assert.Equal("=LEFT(\"abc\",2)", left.ToText());
        Assert.Equal(257, sheet.FunctionCount);
        var trim = ExcelFormula.Parse("=TRIM(\"a\")", null, ExcelVersion.Excel150);
        Assert.EndsWith("5903", Convert.ToHexString(trim.Expression));
        Assert.Equal("=TRIM(\"a\")", trim.ToText());
    }

    [Fact]
    public void NewerFunctions_RaiseTheDocumentsFunctionCount()
    {
        var sheet = ExcelDocument.CreateMacroSheet();
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=GET.DOCUMENT(1)", sheet));
        var formula = ExcelFormula.Parse("=GET.DOCUMENT(1)", sheet, ExcelVersion.Excel150);
        Assert.Equal(239, sheet.FunctionCount);
        Assert.Equal("=GET.DOCUMENT(1)", formula.ToText());
    }

    [Fact]
    public void Excel15_SampleDecompilesNewFunctions()
    {
        var document = Load("INVOICE MACROS");
        Assert.Equal(239, document.FunctionCount);
        Assert.Equal("=\"='\"&GET.DOCUMENT(1)&\"'\"", document.Cells["A7"].Formula!.ToText(document));
    }

    [Fact]
    public void Version18_ReadsAndRoundTrips()
    {
        var bytes = Sample("SIMUL EQUATIONS");
        var document = ExcelDocument.Read(bytes);
        Assert.Equal(18, document.Version);
        Assert.Equal(3, document.NumberFormats.Count);
        Assert.All(document.NumberFormats, f => Assert.Equal(9, f.Descriptor.Length));
        Assert.Equal(ExcelWindow.Version18StateSize, Assert.Single(document.Windows).State.Length);
        Assert.Equal(bytes, document.ToArray());
    }

    [Fact]
    public void ArrayFormulas_DecompileFromTheirAnchor()
    {
        var document = Load("SIMUL EQUATIONS");
        var anchor = document.Cells["B3"].Formula!;
        Assert.True(anchor.IsArrayFormula);
        Assert.True(anchor.IsSharedAnchor);
        Assert.Equal("{=SET.VALUE(solution,data)}", anchor.ToText(document));
        Assert.Equal("{=AVERAGE(D4-B9:B18)}", Load("ACCOUNTS DATA").Cells["D5"].Formula!.ToText(Load("ACCOUNTS DATA")));
    }

    [Fact]
    public void SetArrayFormula_RoundTrips()
    {
        var sheet = ExcelDocument.CreateWorksheet();
        sheet.SetArrayFormula(new ExcelRange(0, 2, 1, 1), "{=A1:A3*2}");
        var reread = RoundTrip(sheet);
        foreach (var address in new[] { "B1", "B2", "B3" })
        {
            var formula = reread.Cells[address].Formula!;
            Assert.True(formula.IsArrayFormula);
            Assert.Equal("{=A1:A3*2}", formula.ToText(reread));
        }

        Assert.True(reread.Cells["B1"].Formula!.IsSharedAnchor);
        Assert.Equal(new ExcelRange(0, 2, 1, 1), reread.Cells["B1"].Formula!.SharedRange);
        Assert.Equal(new ExcelCellAddress(0, 1), reread.Cells["B3"].Formula!.SharedAnchor);
    }

    [Fact]
    public void ArrayConstants_InNamesAndFormulas()
    {
        var matrix = Load("MATRIX");
        var name = matrix.Names.Single(n => n.Name == "col.tmp");
        Assert.Equal("{0;2;-3;1;0;0}", name.GetDefinitionText(matrix));

        var sheet = ExcelDocument.CreateWorksheet();
        var defined = new ExcelDefinedName("Table");
        defined.SetDefinition("={1,2;\"a\",-4}", sheet);
        sheet.Names.Add(defined);
        var formula = ExcelFormula.Parse("=SUM({1,2,3})");
        sheet.Cells["A1"].SetFormula(formula);
        var reread = RoundTrip(sheet);
        Assert.Equal("{1,2;\"a\",-4}", reread.Names.Single(n => n.Name == "Table").GetDefinitionText(reread));
        Assert.Equal("=SUM({1,2,3})", reread.Cells["A1"].Formula!.ToText(reread));
    }

    [Fact]
    public void ExternalMacroCalls_Parse()
    {
        var document = Load("COMMISSION TEMPLATE");
        var text = document.Cells["H12"].Formula!.ToText(document);
        Assert.Equal("=IF(RATE=\"bonus rate\",'COMMISSION MACROS'!Bonus(SALES),0)", text);
        Assert.Equal(text, ExcelFormula.Parse(text, document, ExcelVersion.Excel150).ToText(document));
    }

    [Fact]
    public void NameDefinitions_CanBeUnions()
    {
        var document = Load("INVOICE TEMPLATE");
        var name = document.Names.Single(n => n.Name == "DataInput");
        Assert.StartsWith("$C$9,$E$9,", name.GetDefinitionText(document));
    }

    [Theory]
    [InlineData("CLOCK.CH", 2, 1)]
    [InlineData("BREAK CHART", 1, 5)]
    [InlineData("PROFIT CHART", 1, 1)]
    public void RealCharts_Read(string name, int charts, int series)
    {
        var chart = ChartDocument.Read(Sample(name));
        Assert.Equal(charts, chart.Charts.Count);
        Assert.Equal(series, chart.Series.Count);
        Assert.All(chart.Series, s => Assert.Equal(s.PointCount, s.Values.Count));
    }

    [Fact]
    public void ClockChart_HasUnattachedHourLabels()
    {
        var chart = ChartDocument.Read(Sample("CLOCK.CH"));
        var labels = chart.Charts.SelectMany(c => c.Texts).Select(t => t.Text).ToList();
        Assert.Contains("XII", labels);
        Assert.Contains("VI", labels);
    }
}
