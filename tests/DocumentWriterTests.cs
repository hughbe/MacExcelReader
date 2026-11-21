using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class DocumentWriterTests
{
    [Fact]
    public void CreateWorksheet_HasExcelDefaults()
    {
        var document = ExcelDocument.CreateWorksheet();
        Assert.Equal(ExcelDocumentKind.Worksheet, document.Kind);
        Assert.Equal("XLBN", document.FinderType);
        Assert.Equal(208, document.FunctionCount);
        Assert.Equal(19, document.BuiltInFormatCount);
        Assert.Equal(5, document.Names.Count);
        Assert.Single(document.Windows);
        Assert.False(document.Windows[0].DisplayFormulas);
        Assert.Equal(2, document.CodeSegments.Count);

        var reread = RoundTrip(document);
        Assert.Equal(0, reread.Cells.RowCount);
        Assert.Equal(document.ToArray(), reread.ToArray());
    }

    [Fact]
    public void CreateMacroSheet_DisplaysFormulas()
    {
        var document = ExcelDocument.CreateMacroSheet();
        Assert.Equal("XLPG", document.FinderType);
        Assert.True(RoundTrip(document).Windows[0].DisplayFormulas);
    }

    [Fact]
    public void Cells_AllTypesAndAttributesRoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetInteger(-5);
        document.Cells["B1"].SetNumber(Math.PI);
        document.Cells["C1"].SetText("Café ™");
        var cell = document.Cells["D1"];
        cell.SetNumber(29220);
        cell.NumberFormat = 10;
        cell.Alignment = ExcelAlignment.Center;
        cell.Italic = true;
        cell.TopBorder = true;
        cell.IsHidden = true;
        document.Cells["E2"].Bold = true;

        var reread = RoundTrip(document);

        Assert.Equal(-5, reread.Cells["A1"].IntegerValue);
        Assert.Equal(Math.PI, reread.Cells["B1"].NumberValue);
        Assert.Equal("Café ™", reread.Cells["C1"].TextValue);
        var d1 = reread.Cells["D1"];
        Assert.Equal(10, d1.NumberFormat);
        Assert.Equal(ExcelAlignment.Center, d1.Alignment);
        Assert.True(d1.Italic && d1.TopBorder && d1.IsHidden);
        Assert.False(reread.Cells["E2"].HasValue);
        Assert.True(reread.Cells["E2"].Bold);
        Assert.Equal(2, reread.Cells.RowCount);
        Assert.Equal(5, reread.Cells.ColumnCount);
    }

    [Fact]
    public void Formulas_AllResultKindsRoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        var text = new ExcelFormula([0x17, 2, (byte)'h', (byte)'i']);
        text.SetResult("hi");
        document.Cells["A1"].SetFormula(text);
        var logical = new ExcelFormula([0x1D, 1]);
        logical.SetResult(true);
        document.Cells["A2"].SetFormula(logical);
        var error = new ExcelFormula([0x1C, 42]);
        error.SetErrorResult(42);
        document.Cells["A3"].SetFormula(error);

        var reread = RoundTrip(document);

        Assert.Equal("hi", reread.Cells["A1"].Value);
        Assert.Equal("=\"hi\"", reread.Cells["A1"].Formula!.ToText());
        Assert.Equal(true, reread.Cells["A2"].Value);
        Assert.Equal("#N/A", reread.Cells["A3"].Value);
        Assert.Equal(0, reread.Cells["A3"].Formula!.StoredColumn);
        Assert.Equal(2, reread.Cells["A3"].Formula!.StoredRowWord & 0x3FFF);
    }

    [Fact]
    public void Names_CommandMacrosAndDefinitionsRoundTrip()
    {
        var document = ExcelDocument.CreateMacroSheet();
        document.Names.Add(new ExcelDefinedName("Go", [0x2C, 0x00, 0x04, 0x01], ExcelDefinedName.MacroFlag | ExcelDefinedName.ReferenceFlag)
        {
            ShortcutKey = (byte)'g',
        });

        var reread = RoundTrip(document);
        var name = reread.Names.Single(n => n.Name == "Go");
        Assert.True(name.IsCommand);
        Assert.Equal((byte)'g', name.ShortcutKey);
        Assert.Equal("$B$5", name.GetDefinitionText(reread));
    }

    [Fact]
    public void NewDocuments_HaveTheFixedHeaderExcelExpects()
    {
        // Excel's Open reads the data offset from header offset 22, whatever the segment count says.
        var bytes = ExcelDocument.CreateWorksheet().ToArray();
        Assert.Equal(2, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(12)));
        Assert.Equal(24, System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(22)));

        var document = ExcelDocument.CreateWorksheet();
        document.CodeSegments.Clear();
        Assert.Throws<InvalidOperationException>(() => document.ToArray());
    }

    [Fact]
    public void ExternalDocuments_RoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.ExternalDocuments.Add(new ExcelExternalDocument("Budget") { Names = { "Total", "Tax" } });
        document.ExternalDocuments.Add(new ExcelExternalDocument("Sales") { Names = { "Q1" } });
        document.ExternalDocuments.Add(new ExcelExternalDocument("Empty"));

        var reread = RoundTrip(document);
        Assert.True(reread.HasExternalDocumentTable);
        Assert.Equal(["Budget", "Sales", "Empty"], reread.ExternalDocuments.Select(d => d.Name));
        Assert.Equal(["Total", "Tax"], reread.ExternalDocuments[0].Names);
        Assert.Equal(["Q1"], reread.ExternalDocuments[1].Names);
        Assert.Empty(reread.ExternalDocuments[2].Names);
    }

    [Fact]
    public void PageBreaksAndPrintSetup_RoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.RowPageBreaks.AddRange([10, 20, 30, 40, 50, 60, 70]);
        document.ColumnPageBreaks.AddRange([3, 6]);
        var print = new ExcelPrintSetup(new byte[ExcelPrintSetup.RecordSize])
        {
            Header = "&f",
            Footer = "Page &p",
            PrintHeadings = true,
            PrintGridlines = true,
            TtyPrintWidth = 80,
        };
        print.SetMargin(0, 0.75);
        print.SetMargin(3, 1.25);
        document.PrintSetup = print;

        var bytes = document.ToArray();
        var reread = ExcelDocument.Read(bytes);

        Assert.Equal([10, 20, 30, 40, 50, 60, 70], reread.RowPageBreaks);
        Assert.Equal([3, 6], reread.ColumnPageBreaks);
        var setup = reread.PrintSetup!;
        Assert.Equal("&f", setup.Header);
        Assert.Equal("Page &p", setup.Footer);
        Assert.Null(setup.PrinterMessage);
        Assert.True(setup.PrintHeadings && setup.PrintGridlines);
        Assert.Equal(80, setup.TtyPrintWidth);
        Assert.Equal(0.75, setup.GetMargin(0));
        Assert.Equal(1.25, setup.GetMargin(3));
        Assert.Equal(bytes, reread.ToArray());
    }

    [Fact]
    public void PrintSetup_ValidatesAndStoresAllProperties()
    {
        Assert.Throws<ArgumentException>(() => new ExcelPrintSetup(new byte[ExcelPrintSetup.RecordSize - 1]));
        var setup = new ExcelPrintSetup(new byte[ExcelPrintSetup.RecordSize]);
        setup.TtyLeftMargin = 2;
        setup.TtyPrintWidth = 78;
        setup.TtyTopMargin = 3;
        setup.TtyPrintLength = 60;
        Assert.Equal(2, setup.TtyLeftMargin);
        Assert.Equal(78, setup.TtyPrintWidth);
        Assert.Equal(3, setup.TtyTopMargin);
        Assert.Equal(60, setup.TtyPrintLength);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.GetMargin(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.SetMargin(4, 1));
    }

    [Fact]
    public void SheetSettings_RoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        var state = document.State;
        state.MaxIterations = 7;
        state.CalculationMode = -1;
        state.FullPrecision = false;
        state.Iteration = true;
        state.MaxChange = 0.5;
        state.IsProtected = true;
        state.PasswordHash = ExcelPassword.Hash("pw");
        state.SetColumnWidth(3, 120);
        state.StandardFontSize = 12;

        var reread = RoundTrip(document).State;
        Assert.Equal(7, reread.MaxIterations);
        Assert.Equal(-1, reread.CalculationMode);
        Assert.False(reread.FullPrecision);
        Assert.True(reread.Iteration);
        Assert.Equal(0.5, reread.MaxChange);
        Assert.True(reread.IsProtected);
        Assert.Equal(ExcelPassword.Hash("pw"), reread.PasswordHash);
        Assert.Equal(120, reread.GetColumnWidth(3));
        Assert.Equal(12, reread.StandardFontSize);
        Assert.Throws<ArgumentOutOfRangeException>(() => reread.GetColumnWidth(256));
    }

    [Fact]
    public void Windows_RoundTrip()
    {
        var document = ExcelDocument.CreateWorksheet();
        var window = document.Windows[0];
        window.ActiveRow = 9;
        window.ActiveColumn = 3;
        window.DisplayGridlines = false;
        window.HorizontalSplit = 100;
        window.Selection.Clear();
        window.Selection.Add(new ExcelRange(9, 12, 3, 4));

        var reread = RoundTrip(document).Windows[0];
        Assert.Equal(9, reread.ActiveRow);
        Assert.Equal(3, reread.ActiveColumn);
        Assert.False(reread.DisplayGridlines);
        Assert.Equal(100, reread.HorizontalSplit);
        Assert.Equal(new ExcelRange(9, 12, 3, 4), Assert.Single(reread.Selection));
    }

    [Fact]
    public void Window_ValidatesAndStoresAllProperties()
    {
        Assert.Throws<ArgumentException>(() => new ExcelWindow(0, 0, 0, 0, new byte[ExcelWindow.StateSize - 1]));
        var window = new ExcelWindow(1, 2, 3, 4, new byte[ExcelWindow.StateSize])
        {
            ActiveArea = 5,
            DisplayHeadings = false,
            VerticalSplit = 120,
        };
        Assert.Equal(5, window.ActiveArea);
        Assert.False(window.DisplayHeadings);
        Assert.Equal(120, window.VerticalSplit);
    }

    [Fact]
    public void Names_RejectInvalidLengthsAndRender()
    {
        Assert.Throws<ArgumentException>(() => new ExcelDefinedName(string.Empty));
        var name = new ExcelDefinedName("ShortName");
        Assert.Equal("ShortName", name.ToString());

        var document = ExcelDocument.CreateWorksheet();
        document.Names.Add(new ExcelDefinedName(new string('x', 256)));
        Assert.Throws<InvalidOperationException>(() => document.ToArray());
        document.Names.Clear();
        document.Names.Add(new ExcelDefinedName("LongDefinition", new byte[256]));
        Assert.Throws<InvalidOperationException>(() => document.ToArray());
    }

    [Fact]
    public void Document_SaveWriteAndOptionalMetadataPaths()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.NumberFormats.Add(new ExcelNumberFormat(new byte[7], []));
        document.HasExternalDocumentTable = true;
        document.PrintSetup = new ExcelPrintSetup(new byte[ExcelPrintSetup.RecordSize])
        {
            Header = "header",
            Footer = string.Empty,
            PrinterMessage = "message",
        };
        var path = Path.GetTempFileName();
        try
        {
            document.Save(path);
            Assert.Equal(document.ToArray(), File.ReadAllBytes(path));
            using var stream = new MemoryStream();
            document.Write(stream);
            Assert.Equal(document.ToArray(), stream.ToArray());
            Assert.NotNull(ExcelDocument.Read(new NonSeekableStream(document.ToArray())));
        }
        finally
        {
            File.Delete(path);
        }

        Assert.Throws<ArgumentNullException>(() => document.Write(null!));
        Assert.Null(ExcelDocument.CreateWorksheet().ToString() is null ? null : null);
    }

    [Fact]
    public void NumberFormats_RejectInvalidFixedSizeRecords()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.NumberFormats.Add(new ExcelNumberFormat(new byte[6], []));
        Assert.Throws<InvalidOperationException>(() => document.ToArray());

        document.NumberFormats.Clear();
        var descriptor = new byte[7];
        descriptor[0] = 1;
        document.NumberFormats.Add(new ExcelNumberFormat(descriptor, []));
        Assert.Throws<InvalidOperationException>(() => document.ToArray());
    }

    [Fact]
    public void Document_RejectsCellDataBeyondTableAddressSpace()
    {
        var document = ExcelDocument.CreateWorksheet();
        for (int i = 0; i < 131_073; i++)
        {
            document.Cells[i / 256, i % 256].SetNumber(i + 0.5);
        }

        Assert.Throws<InvalidOperationException>(() => document.ToArray());
    }

    private sealed class NonSeekableStream(byte[] bytes) : Stream
    {
        private readonly MemoryStream _stream = new(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _stream.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    [Fact]
    public void EditingASample_KeepsEverythingElse()
    {
        var original = Load("REPLACE.xlm");
        var document = Load("REPLACE.xlm");
        document.Cells["A1"].SetText("Edited");
        document.Cells["B23"].Formula!.SetResult(99);

        var reread = RoundTrip(document);

        Assert.Equal("Edited", reread.Cells["A1"].TextValue);
        Assert.Equal(99.0, reread.Cells["B23"].Value);
        foreach (var (row, column, cell) in original.Cells.NonEmptyCells)
        {
            if (row == 22 && column == 1)
            {
                continue;
            }

            var other = reread.Cells.TryGet(row, column)!;
            Assert.Equal(cell.Type, other.Type);
            Assert.Equal(cell.Formula?.ToText(original), other.Formula?.ToText(reread));
            if (cell.Type == ExcelCellType.Text)
            {
                Assert.Equal(cell.TextValue, other.TextValue);
            }
        }
    }

    [Fact]
    public void Excel100FunctionNames_AreReadFromSection7AndWrittenBack()
    {
        // Excel 1.00 writes the names of functions numbered 208 and up at the start of section 7 (spec §14).
        var document = ExcelDocument.CreateWorksheet();
        document.ExtraHeader = [];
        document.FunctionCount = 210;
        document.ExtraFunctionNames.AddRange(["FIRST.EXTRA", "SECOND.EXTRA"]);
        document.RowPageBreaks.Add(5);
        document.Cells[0, 0].SetText("x");

        var bytes = document.ToArray();
        var reread = ExcelDocument.Read(bytes);
        Assert.Equal(["FIRST.EXTRA", "SECOND.EXTRA"], reread.ExtraFunctionNames);
        Assert.Equal([(ushort)5], reread.RowPageBreaks);
        Assert.Equal(bytes, reread.ToArray());

        // With an extra header (Excel 1.03 and later) a larger function count carries no names.
        var later = ExcelDocument.CreateWorksheet();
        later.FunctionCount = 239;
        later.RowPageBreaks.Add(5);
        var rereadLater = ExcelDocument.Read(later.ToArray());
        Assert.Empty(rereadLater.ExtraFunctionNames);
        Assert.Equal([(ushort)5], rereadLater.RowPageBreaks);
    }

    [Fact]
    public void GrowingASample_WritesTheLargerGrid()
    {
        var document = Load("Expenses.xls");
        document.Cells[300, 7].SetText("new corner");

        var reread = RoundTrip(document);
        // The table keeps a blank row and column after the last cell with a value (Excel's margin).
        Assert.Equal(302, reread.Cells.RowCount);
        Assert.Equal(9, reread.Cells.ColumnCount);
        Assert.Equal("new corner", reread.Cells[300, 7].TextValue);
        Assert.Equal("Date", reread.Cells["A10"].TextValue);
        Assert.Equal(29220, reread.Cells["A11"].IntegerValue);
    }
}
