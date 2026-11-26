namespace MacExcelReader.Tests;

public class ExcelDocumentTests
{
    public static TheoryData<string> Samples => TestData.AllSheets();

    public static TheoryData<string> Charts => TestData.AllCharts();

    private static string SamplePath(string name) => name.Contains('/') ? TestData.SampleFile(name) : TestData.SamplePath(name);

    [Theory]
    [MemberData(nameof(Charts))]
    public void Read_Chart_RoundTripsByteForByte(string name)
    {
        var original = File.ReadAllBytes(SamplePath(name));
        Assert.Equal(original, MacExcelReader.Charts.ChartDocument.Read(original).ToArray());
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void Read_Sample_RoundTripsByteForByte(string name)
    {
        var original = File.ReadAllBytes(SamplePath(name));
        var document = ExcelDocument.Read(original);

        var written = document.ToArray();

        Assert.Equal(original, written);
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void Read_Sample_DecompilesEveryFormula(string name)
    {
        var document = ExcelDocument.Load(SamplePath(name));
        foreach (var entry in document.Cells.NonEmptyCells)
        {
            var cell = entry.GetCell();
            if (cell.Formula is { } formula)
            {
                var text = formula.ToText(document);
                Assert.True(text.StartsWith('=') || text.StartsWith("{=", StringComparison.Ordinal), text);
                Assert.DoesNotContain("FUNCTION", text);
                Assert.DoesNotContain("{array", text);
            }
        }

        foreach (var definedName in document.Names)
        {
            _ = definedName.GetDefinitionText(document);
        }
    }

    [Fact]
    public void Read_Worksheet1_IsEmptyWorksheet()
    {
        var document = ExcelDocument.Load(SamplePath("Worksheet1.xls"));

        Assert.Equal(ExcelDocumentKind.Worksheet, document.Kind);
        Assert.Equal(0xAB27, document.Signature);
        Assert.Equal(208, document.FunctionCount);
        Assert.Equal(0, document.Cells.RowCount);
        Assert.Equal(["Criteria", "Database", "Print_Area", "Print_Titles", "Recorder"], document.Names.Select(n => n.Name));
        Assert.Equal(100, document.State.MaxIterations);
        Assert.Equal(0.001, document.State.MaxChange);
        Assert.Equal(75, document.State.GetColumnWidth(0));
        var window = Assert.Single(document.Windows);
        Assert.Equal(3, window.ActiveColumn);
        Assert.Equal(4, window.ActiveRow);
        Assert.Equal(ExcelRange.Parse("D5"), Assert.Single(window.Selection));
    }

    [Fact]
    public void Read_Expenses_ReadsValuesAndFormats()
    {
        var document = ExcelDocument.Load(SamplePath("Expenses.xls"));

        Assert.Equal(257, document.Cells.RowCount);
        Assert.Equal(5, document.Cells.ColumnCount);
        Assert.Equal("Date", document.Cells[9, 0].TextValue);
        Assert.Equal("Expense ", document.Cells[9, 1].TextValue);
        var date = document.Cells[10, 0];
        Assert.Equal(ExcelCellType.Integer, date.Type);
        Assert.Equal(10, date.NumberFormat);
        Assert.Equal(110, document.State.GetColumnWidth(3));
        Assert.Equal("$A$10:$D$256", document.Names.Single(n => n.Name == "Database").GetDefinitionText(document));
    }

    [Fact]
    public void Read_Replace_DecompilesMacroFormulas()
    {
        var document = ExcelDocument.Load(SamplePath("REPLACE.xlm"));

        Assert.Equal(ExcelDocumentKind.MacroSheet, document.Kind);
        Assert.StartsWith("This macro sheet contains", document.Cells[0, 1].TextValue);
        Assert.Equal("=ROW(ACTIVE.CELL())", document.Cells[27, 1].Formula!.ToText(document));

        var replace = document.Names.Single(n => n.Name == "REPLACE");
        Assert.Equal(ExcelMacroKind.Command, replace.MacroKind);
        Assert.Equal('r', replace.ShortcutKey);
        Assert.Equal("$B$19", replace.GetDefinitionText(document));
    }

    [Fact]
    public void Read_DbForm_HasExternalDocumentTableAndProtection()
    {
        var document = ExcelDocument.Load(SamplePath("DB.FORM.xlm"));

        Assert.True(document.HasExternalDocumentTable);
        Assert.Single(document.ExternalDocuments);
        Assert.Equal(["DATABASE", "FORM"], document.ExternalDocuments[0].Names);
        Assert.True(document.State.IsProtected);
    }

    [Fact]
    public void ToArray_KeepsABlankMarginRowAndColumn()
    {
        // Excel 1.03 hangs on a one-row table and ignores values in the table's last row and column, so the
        // writer always leaves them blank. Writing again does not grow the table further.
        var document = ExcelDocument.CreateMacroSheet();
        document.Cells[0, 0].SetText("A1");
        document.Cells[0, 5].SetText("F1");

        var reread = ExcelDocument.Read(document.ToArray());
        Assert.Equal(2, reread.Cells.RowCount);
        Assert.Equal(7, reread.Cells.ColumnCount);
        Assert.Equal("F1", reread.Cells[0, 5].TextValue);

        var again = ExcelDocument.Read(reread.ToArray());
        Assert.Equal(2, again.Cells.RowCount);
        Assert.Equal(7, again.Cells.ColumnCount);
    }

    [Fact]
    public void CreateWorksheet_WriteAndRead_PreservesCells()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells[0, 0].SetText("Hello");
        document.Cells[1, 0].SetValue(42);
        document.Cells[2, 0].SetValue(3.25);
        document.Cells[2, 0].Bold = true;
        document.Cells[2, 0].NumberFormat = 2;

        // =A2*2
        var formula = new ExcelFormula([0x44, 0x80 | 0x40, 0x01, 0x00, 0x1E, 0x00, 0x02, 0x05]);
        formula.SetResult(84);
        document.Cells[3, 1].SetFormula(formula);

        var bytes = document.ToArray();
        var reread = ExcelDocument.Read(bytes);

        // Four rows and two columns in use, plus Excel's blank margin row and column.
        Assert.Equal(5, reread.Cells.RowCount);
        Assert.Equal(3, reread.Cells.ColumnCount);
        Assert.Equal("Hello", reread.Cells[0, 0].TextValue);
        Assert.Equal(42, reread.Cells[1, 0].IntegerValue);
        Assert.Equal(3.25, reread.Cells[2, 0].NumberValue);
        Assert.True(reread.Cells[2, 0].Bold);
        Assert.Equal(2, reread.Cells[2, 0].NumberFormat);
        Assert.Equal("=A2*2", reread.Cells[3, 1].Formula!.ToText(reread));
        Assert.Equal(84.0, reread.Cells[3, 1].Value);
        Assert.Equal(bytes, reread.ToArray());
    }

    [Fact]
    public void CreateWorksheet_MatchesSizeOfExcelEmptySheet()
    {
        var created = ExcelDocument.CreateWorksheet().ToArray();
        var excel = ExcelDocument.Read(File.ReadAllBytes(SamplePath("Worksheet1.xls")));
        excel.CodeSegment1 = default;
        excel.CodeSegment2 = default;
        excel.Windows[0].ActiveColumn = 0;
        excel.Windows[0].ActiveRow = 0;
        excel.Windows[0].Selection[0] = new ExcelRange(0, 0, 0, 0);

        // Only the stale pointers in the state block differ.
        var expected = excel.ToArray();
        Assert.Equal(expected.Length, created.Length);
    }

    [Fact]
    public void Password_Hash_EmptyIsZero()
    {
        Assert.Equal(0, ExcelPassword.Hash(string.Empty));
        Assert.NotEqual(0, ExcelPassword.Hash("secret"));
    }

    [Fact]
    public void MacRoman_RoundTripsAllBytes()
    {
        var bytes = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray();
        Assert.Equal(bytes, MacRoman.GetBytes(MacRoman.GetString(bytes)));
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(255, "IV")]
    public void ColumnName_Formats(int column, string expected)
    {
        Assert.Equal(expected, FormulaDecompiler.ColumnName(column));
    }

    [Fact]
    public void Functions_TableMatchesExcel()
    {
        Assert.Equal(208, ExcelFunctions.Count);
        Assert.Equal("IF", ExcelFunctions.Get(1)!.Name);
        Assert.Equal("ACTIVE.CELL", ExcelFunctions.Get(169)!.Name);
        Assert.True(ExcelFunctions.TryFindNumber("sum", out int sum));
        Assert.Equal(4, sum);
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public void EverySample_HasItsNamesInExcelsOrder(string name)
    {
        // Excel keeps the names sorted, ignoring case and accents (spec §11); only the experiment that wrote them out
        // of order has names elsewhere.
        var document = ExcelDocument.Read(File.ReadAllBytes(TestData.SampleFile(name)));
        var before = document.Names.ToList();
        bool unsorted = name.EndsWith("Names Unsorted", StringComparison.Ordinal);
        Assert.Equal(!unsorted, document.NamesAreSorted);
        document.SortNames();
        Assert.True(document.NamesAreSorted);
        if (unsorted)
        {
            Assert.NotEqual(before, document.Names);
        }
        else
        {
            Assert.Equal(before, document.Names);
        }
    }

    private static (int Rows, int Columns) SavedTable(ExcelDocument document)
    {
        var saved = ExcelDocument.Read(document.ToArray());
        return (saved.Cells.RowCount, saved.Cells.ColumnCount);
    }

    [Theory]
    [InlineData("=D20", 21, 5)]
    [InlineData("=SUM(C3:E9)", 10, 6)]
    [InlineData("=SUM(B2:B16384)", 2, 3)]
    [InlineData("=SUM(B2:IV2)", 3, 2)]
    [InlineData("=Other!K99", 2, 2)]
    public void Save_GrowsAWorksheetsTableToCoverItsFormulasReferences(string text, int rows, int columns)
    {
        // Spec §6.1: a worksheet's table covers every cell its formulas refer to, plus the blank margin, except an
        // area reaching the last row (for rows) or column (for columns) and other documents' cells.
        var document = ExcelDocument.CreateWorksheet();
        document.ExternalDocuments.Add(new ExcelExternalDocument("Other"));
        document.Cells[0, 0].SetFormula(ExcelFormula.Parse(text, document));
        Assert.Equal((rows, columns), SavedTable(document));
    }

    [Theory]
    [InlineData("C7", "E10", "E10", 11, 6)]
    [InlineData("A1", "A2", "A1", 16385, 2)]
    [InlineData("A1", "B1", "A1", 2, 257)]
    [InlineData("C7", "E10", "A1", 2, 2)]
    public void Save_FollowsARelativeNameFromTheCellThatUsesIt(string definition, string definedAt, string usedAt, int rows, int columns)
    {
        // A relative name refers to cells from the cell that uses it, wrapping round the sheet (spec §11): C7 defined
        // with E10 active is three rows up and two columns left, so from E10 the table covers E10 (11 x 6); A1
        // defined from A2 is one row up, which from A1 is A16384. From A1 the first name is IU16382, which would need a
        // table reaching the last row and column; Excel refuses that formula, and the library keeps the table small.
        var document = ExcelDocument.CreateWorksheet();
        var name = new ExcelDefinedName("Near");
        name.SetDefinition("=" + definition, ExcelCellAddress.Parse(definedAt), document);
        document.Names.Add(name);
        var at = ExcelCellAddress.Parse(usedAt);
        document.Cells[at.Row, at.Column].SetFormula(ExcelFormula.Parse("=Near", document));
        Assert.Equal((rows, columns), SavedTable(document));
    }

    [Fact]
    public void Save_LeavesAMacroSheetsTableAtItsCells()
    {
        var document = ExcelDocument.CreateMacroSheet();
        document.Cells[0, 0].SetFormula(ExcelFormula.Parse("=D20", document));
        Assert.Equal((2, 2), SavedTable(document));
    }
}
