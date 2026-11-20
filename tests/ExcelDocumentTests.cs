namespace MacExcelReader.Tests;

public class ExcelDocumentTests
{
    private static string SamplePath(string name) => Path.Combine(AppContext.BaseDirectory, "Samples", name);

    public static TheoryData<string> Samples =>
    [
        "Worksheet1.xls",
        "Expenses.xls",
        "COMMANDS.xlm",
        "DB.FORM.xlm",
        "REPLACE.xlm",
    ];

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
        foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
        {
            if (cell.Formula is { } formula)
            {
                var text = formula.ToText(document);
                Assert.StartsWith("=", text);
                Assert.DoesNotContain("FUNCTION", text);
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
        Assert.Equal(new ExcelRange(4, 4, 3, 3), Assert.Single(window.Selection));
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
        Assert.True(replace.IsCommand);
        Assert.Equal((byte)'r', replace.ShortcutKey);
        Assert.Equal("$B$19", replace.GetDefinitionText(document));
    }

    [Fact]
    public void Read_DbForm_HasExternalDocumentTableAndProtection()
    {
        var document = ExcelDocument.Load(SamplePath("DB.FORM.xlm"));

        Assert.True(document.HasExternalDocumentTable);
        Assert.Single(document.ExternalDocuments);
        Assert.Equal(["DATABASE", "FORM"], document.ExternalNames);
        Assert.True(document.State.IsProtected);
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

        Assert.Equal(4, reread.Cells.RowCount);
        Assert.Equal(2, reread.Cells.ColumnCount);
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
        excel.CodeSegments.Clear();
        excel.CodeSegments.Add([]);
        excel.CodeSegments.Add([]);
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
        Assert.Equal(4, ExcelFunctions.FindNumber("sum"));
    }
}
