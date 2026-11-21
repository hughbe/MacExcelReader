namespace MacExcelReader.Tests;

public class CellTests
{
    [Fact]
    public void NewCell_IsEmptyAndLocked()
    {
        var cell = new ExcelCell();
        Assert.Equal(ExcelCellType.Empty, cell.Type);
        Assert.False(cell.HasValue);
        Assert.Null(cell.Value);
        Assert.True(cell.IsLocked);
        Assert.Equal(0x4000, cell.RawAttributes);
        Assert.False(cell.TryGetNumber(out _));
    }

    [Theory]
    [InlineData(5.0, ExcelCellType.Integer)]
    [InlineData(-32768.0, ExcelCellType.Integer)]
    [InlineData(32768.0, ExcelCellType.Number)]
    [InlineData(2.5, ExcelCellType.Number)]
    [InlineData(-0.0, ExcelCellType.Number)]
    public void SetValue_ChoosesCompactForm(double value, ExcelCellType expected)
    {
        var cell = new ExcelCell();
        cell.SetValue(value);
        Assert.Equal(expected, cell.Type);
        Assert.True(cell.TryGetNumber(out double read));
        Assert.Equal(value, read);
    }

    [Fact]
    public void TypedAccessors_ThrowForOtherTypes()
    {
        var cell = new ExcelCell();
        cell.SetText("x");
        Assert.Throws<InvalidOperationException>(() => cell.IntegerValue);
        Assert.Throws<InvalidOperationException>(() => cell.NumberValue);
        Assert.Null(cell.Formula);
        cell.SetNumber(1.5);
        Assert.Throws<InvalidOperationException>(() => cell.TextValue);
    }

    [Fact]
    public void SetText_RejectsLongText()
    {
        Assert.Throws<ArgumentException>(() => new ExcelCell().SetText(new string('x', 256)));
        new ExcelCell().SetText(new string('x', 255));
    }

    [Fact]
    public void Attributes_MapToTheDocumentedBits()
    {
        var cell = new ExcelCell { IsLocked = false };
        cell.Bold = true;
        Assert.Equal(0x0001, cell.RawAttributes);
        cell.Italic = true;
        cell.LeftBorder = true;
        cell.RightBorder = true;
        cell.TopBorder = true;
        cell.BottomBorder = true;
        Assert.Equal(0x003F, cell.RawAttributes);
        cell.IsHidden = true;
        cell.IsLocked = true;
        cell.UsesArrayFunction = true;
        Assert.Equal(0xE03F, cell.RawAttributes);
    }

    [Fact]
    public void FormatAndAlignment_SurviveValueChanges()
    {
        var cell = new ExcelCell { NumberFormat = 10, Alignment = ExcelAlignment.Right, Bold = true };
        cell.SetValue(29220);
        cell.SetText("now text");
        cell.Clear();
        Assert.Equal(10, cell.NumberFormat);
        Assert.Equal(ExcelAlignment.Right, cell.Alignment);
        Assert.True(cell.Bold);
        Assert.Equal(ExcelCellType.Empty, cell.Type);
    }

    [Fact]
    public void NumberFormat_IsRangeChecked()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelCell { NumberFormat = 64 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelCell { NumberFormat = -1 });
    }

    [Fact]
    public void FormulaResultKinds()
    {
        var formula = new ExcelFormula([0x1E, 0, 1]);
        var cell = new ExcelCell();
        cell.SetFormula(formula);
        Assert.Equal(ExcelCellType.NumericFormula, cell.Type);
        Assert.Equal(0.0, cell.Value);

        formula.SetResult("hello");
        cell.SetFormula(formula);
        Assert.Equal(ExcelCellType.Formula, cell.Type);
        Assert.Equal("hello", cell.Value);
        Assert.Equal(ExcelFormulaResultKind.Text, formula.Result.Kind);

        formula.SetResult(false);
        Assert.Equal(false, formula.Result.Value);

        formula.SetErrorResult(15);
        Assert.Equal("#VALUE!", formula.Result.Value);

        formula.SetResult(2.5);
        Assert.Equal(2.5, formula.Result.Value);
    }

    [Theory]
    [InlineData((byte)0, "#NULL!")]
    [InlineData((byte)7, "#DIV/0!")]
    [InlineData((byte)15, "#VALUE!")]
    [InlineData((byte)23, "#REF!")]
    [InlineData((byte)29, "#NAME?")]
    [InlineData((byte)36, "#NUM!")]
    [InlineData((byte)42, "#N/A")]
    [InlineData((byte)99, "#ERR99")]
    public void ErrorText(byte code, string expected)
    {
        Assert.Equal(expected, ExcelFormulaResult.ErrorText(code));
    }

    [Fact]
    public void Grid_GrowsAndNeverShrinks()
    {
        var grid = new ExcelCellGrid();
        Assert.Null(grid.TryGet(0, 0));
        _ = grid[4, 2];
        Assert.Equal(5, grid.RowCount);
        Assert.Equal(3, grid.ColumnCount);
        grid.EnsureSize(1, 1);
        Assert.Equal(5, grid.RowCount);
        Assert.NotNull(grid.TryGet("C5"));
        Assert.Null(grid.TryGet("D5"));
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.EnsureSize(16385, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => grid.EnsureSize(1, 257));
    }

    [Fact]
    public void Grid_EnumerationsAgree()
    {
        var grid = new ExcelCellGrid();
        grid["B2"].SetValue(1);
        grid["A3"].SetText("x");
        grid["C1"].Bold = true;

        Assert.Equal(2, grid.CountNonEmpty());
        Assert.Equal(["B2", "A3"], grid.NonEmptyCells.Select(c => new ExcelCellAddress(c.Row, c.Column).ToString()));
        var refs = new List<string>();
        foreach (var cell in grid.EnumerateNonEmpty())
        {
            refs.Add(cell.Address.ToString());
        }

        Assert.Equal(["B2", "A3"], refs);
        Assert.Equal(ExcelCellType.Integer, grid.GetCellType(1, 1));
        Assert.Equal(ExcelCellType.Empty, grid.GetCellType(100, 100));
    }
}
