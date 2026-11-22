namespace MacExcelReader.Tests;

public class CellTests
{
    [Fact]
    public void NewCell_IsEmptyAndLocked()
    {
        var cell = new ExcelCell();
        Assert.Equal(ExcelCellType.Empty, cell.Type);
        Assert.False(cell.HasValue);
        Assert.True(cell.Value.IsEmpty);
        Assert.True(cell.IsLocked);
        Assert.Equal(0x4000, cell.RawAttributes);
        Assert.False(cell.TryGetNumber(out _));
    }

    [Theory]
    [InlineData(5.0, ExcelCellType.Integer)]
    [InlineData(0.0, ExcelCellType.Integer)]
    [InlineData(65535.0, ExcelCellType.Integer)]
    [InlineData(40000.0, ExcelCellType.Integer)]
    [InlineData(65536.0, ExcelCellType.Number)]
    [InlineData(-7.0, ExcelCellType.Number)]
    [InlineData(-32768.0, ExcelCellType.Number)]
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
        Assert.Equal("hello", cell.Value.Text);
        Assert.Equal(ExcelValueKind.Text, formula.Result.Kind);

        formula.SetResult(false);
        Assert.Equal(ExcelValue.FromBoolean(false), formula.Result);

        formula.SetResult(ExcelValue.FromError(ExcelError.Value));
        Assert.Equal(ExcelError.Value, formula.Result.Error);
        Assert.Throws<ArgumentException>(() => formula.SetResult(ExcelValue.Empty));

        formula.SetResult(2.5);
        Assert.Equal(2.5, formula.Result.Number);
    }

    [Theory]
    [InlineData(ExcelError.Null, "#NULL!")]
    [InlineData(ExcelError.DivideByZero, "#DIV/0!")]
    [InlineData(ExcelError.Value, "#VALUE!")]
    [InlineData(ExcelError.Reference, "#REF!")]
    [InlineData(ExcelError.Name, "#NAME?")]
    [InlineData(ExcelError.Number, "#NUM!")]
    [InlineData(ExcelError.NotAvailable, "#N/A")]
    [InlineData((ExcelError)99, "#ERR99")]
    public void ErrorText(ExcelError error, string expected)
    {
        Assert.Equal(expected, error.ToText());
        if (Enum.IsDefined(error))
        {
            Assert.True(ExcelErrorExtensions.TryParse(expected.ToLowerInvariant(), out var parsed));
            Assert.Equal(error, parsed);
        }
        else
        {
            Assert.False(ExcelErrorExtensions.TryParse(expected, out _));
        }
    }

    [Fact]
    public void Grid_GrowsAndNeverShrinks()
    {
        var grid = new ExcelCellGrid();
        Assert.Null(grid.Find(0, 0));
        _ = grid[4, 2];
        Assert.Equal(5, grid.RowCount);
        Assert.Equal(3, grid.ColumnCount);
        grid.EnsureSize(1, 1);
        Assert.Equal(5, grid.RowCount);
        Assert.NotNull(grid.Find("C5"));
        Assert.Null(grid.Find("D5"));
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
        foreach (var cell in grid.NonEmptyCells)
        {
            refs.Add(cell.Address.ToString());
        }

        Assert.Equal(["B2", "A3"], refs);
        Assert.Equal(ExcelCellType.Integer, grid.GetCellType(1, 1));
        Assert.Equal(ExcelCellType.Empty, grid.GetCellType(100, 100));
    }

    [Fact]
    public void CellRefs_ReadLazyValuesAndMetadata()
    {
        var source = ExcelDocument.CreateWorksheet();
        source.Cells["B2"].SetNumber(2.5);
        var number = source.Cells["B2"];
        number.NumberFormat = 6;
        number.Alignment = ExcelAlignment.Center;
        number.Bold = true;
        number.Italic = true;
        source.Cells["C2"].SetText("hello");
        var document = ExcelDocument.Read(source.ToArray());

        var refs = document.Cells.NonEmptyCells.GetEnumerator();
        Assert.True(refs.MoveNext());
        var numberRef = refs.Current;
        Assert.Equal("B2", numberRef.Address.ToString());
        Assert.Equal(ExcelCellType.Number, numberRef.Type);
        Assert.Equal(6, numberRef.NumberFormat);
        Assert.Equal(ExcelAlignment.Center, numberRef.Alignment);
        Assert.True(numberRef.Bold);
        Assert.True(numberRef.Italic);
        Assert.NotEqual((ushort)0, numberRef.RawAttributes);
        Assert.Equal("B2 (Number)", numberRef.ToString());
        Assert.True(numberRef.TryGetNumber(out double value));
        Assert.Equal(2.5, value);

        Assert.True(refs.MoveNext());
        Assert.Equal("hello", refs.Current.Value.Text);
        Assert.True(source.Cells["A1"].Value.IsEmpty);
        Assert.Equal(2.5, numberRef.Value.Number);
        Assert.False(refs.MoveNext());
    }

    [Fact]
    public void Cells_ExposeLazyAndStringRepresentations()
    {
        var lazy = new ExcelCell();
        lazy.SetNumber(0);
        Assert.Equal(0, lazy.NumberValue);
        Assert.Equal("0", lazy.ToString());

        lazy.SetText("text");
        Assert.Equal("text", lazy.ToString());
        lazy.Clear();
        Assert.Equal(string.Empty, lazy.ToString());
        Assert.Throws<FormatException>(() => ExcelCellAddress.Parse("invalid"));

        var integer = new ExcelCell();
        integer.SetInteger(3);
        Assert.Equal("3", integer.ToString());
        var formula = new ExcelCell();
        var expression = new ExcelFormula([0x1D, 0]);
        expression.SetResult("result");
        formula.SetFormula(expression);
        Assert.Equal("result", formula.ToString());
    }

    [Fact]
    public void LazyText_CorruptionThrowsFormatException()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetText("hello");
        var bytes = document.ToArray();
        var layout = ExcelFileLayout.Read(bytes);
        bytes[layout.GetSectionOffset(ExcelSection.CellData)] = 255;
        var corrupt = ExcelDocument.Read(bytes);
        Assert.Throws<ExcelFormatException>(() => corrupt.Cells["A1"].TextValue);
    }

    [Fact]
    public void BooleanAndErrorCells_RoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetBoolean(true);
        document.Cells["A2"].SetBoolean(false);
        document.Cells["A3"].SetError(ExcelError.NotAvailable);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Cells["A4"].SetError((ExcelError)1));

        var reread = ExcelDocument.Read(document.ToArray());
        Assert.Equal(true, reread.Cells["A1"].Value);
        Assert.False(reread.Cells["A2"].BooleanValue);
        Assert.Equal(ExcelError.NotAvailable, reread.Cells["A3"].Error);
        Assert.Equal("#N/A", ExcelValueFormatter.FormatCell(reread.Cells["A3"], reread));
        Assert.Throws<InvalidOperationException>(() => reread.Cells["A3"].BooleanValue);
    }

    [Fact]
    public void RelativeNameDefinition_IsRelativeToTheGivenCell()
    {
        // The usage guide's example: =B1 defined with B2 active is the cell above wherever it is used.
        var document = ExcelDocument.CreateWorksheet();
        var above = new ExcelDefinedName("Above");
        above.SetDefinition("=B1", new ExcelCellAddress(1, 1), document);
        Assert.Equal([0x2C, 0xFF, 0xFF, 0x00], above.Definition);   // row −1 and column 0, both relative
        Assert.Equal("D9", above.GetDefinitionText(document, new ExcelCellAddress(9, 3)));
    }
}
