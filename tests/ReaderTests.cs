using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class ReaderTests
{
    public static TheoryData<string, int, int, int, int> SampleShapes => new()
    {
        { "Worksheet1.xls", 0, 0, 0, 0 },
        { "Expenses.xls", 257, 5, 988, 0 },
        { "COMMANDS.xlm", 63, 3, 114, 39 },
        { "DB.FORM.xlm", 162, 4, 433, 275 },
        { "REPLACE.xlm", 49, 3, 60, 30 },
    };

    [Theory]
    [MemberData(nameof(SampleShapes))]
    public void Samples_HaveExpectedShape(string name, int rows, int columns, int nonEmpty, int formulas)
    {
        var document = Load(name);
        Assert.Equal(rows, document.Cells.RowCount);
        Assert.Equal(columns, document.Cells.ColumnCount);
        Assert.Equal(nonEmpty, document.Cells.CountNonEmpty());
        Assert.Equal(nonEmpty, document.Cells.NonEmptyCells.Count());
        Assert.Equal(formulas, document.Cells.NonEmptyCells.Count(c => c.Cell.Formula is not null));
    }

    [Theory]
    [MemberData(nameof(SampleShapes))]
    public void Samples_AllReadPathsAgree(string name, int rows, int columns, int nonEmpty, int formulas)
    {
        _ = (rows, columns, formulas);
        var bytes = Sample(name);
        var fromBytes = ExcelDocument.Read(bytes);
        var fromSpan = ExcelDocument.Read((ReadOnlySpan<byte>)bytes);
        var fromStream = ExcelDocument.Read(new MemoryStream(bytes));
        var noCopy = ExcelDocument.Read(bytes.ToArray(), copyData: false);

        foreach (var document in new[] { fromSpan, fromStream, noCopy })
        {
            Assert.Equal(bytes, document.ToArray());
        }

        int count = 0;
        foreach (var cell in fromBytes.Cells.EnumerateNonEmpty())
        {
            count++;
            var full = cell.GetCell();
            Assert.Equal(full.Type, cell.Type);
            Assert.Equal(full.NumberFormat, cell.NumberFormat);
            Assert.Equal(full.TryGetNumber(out double a), cell.TryGetNumber(out double b));
            Assert.Equal(a, b);
            if (cell.Type == ExcelCellType.Text)
            {
                Assert.Equal(full.TextValue, cell.GetText());
            }
        }

        Assert.Equal(nonEmpty, count);
    }

    [Fact]
    public void Read_DoesNotDependOnCallersArray()
    {
        var bytes = Sample("REPLACE.xlm");
        var document = ExcelDocument.Read(bytes);
        Array.Clear(bytes);
        Assert.StartsWith("This macro sheet", document.Cells["B1"].TextValue);
    }

    [Fact]
    public void Read_IsLazyUntilCellsAreTouched()
    {
        var document = Load("DB.FORM.xlm");
        Assert.Equal(ExcelCellType.Formula, document.Cells.GetCellType(27, 0));

        // Saving without touching any cell copies them unchanged.
        Assert.Equal(Sample("DB.FORM.xlm"), document.ToArray());
    }

    [Theory]
    [InlineData(new byte[] { })]
    [InlineData(new byte[] { 0x53, 0x27 })]
    public void Read_TooShort_Throws(byte[] data)
    {
        Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(data));
    }

    [Fact]
    public void Read_BadMagic_Throws()
    {
        var bytes = Sample("Worksheet1.xls");
        bytes[0] = 0;
        var ex = Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(bytes));
        Assert.Contains("magic", ex.Message);
    }

    [Fact]
    public void Read_WrongVersion_Throws()
    {
        var bytes = Sample("Worksheet1.xls");
        bytes[3] = 19;
        var ex = Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(bytes));
        Assert.Contains("version", ex.Message);
    }

    [Fact]
    public void Read_WrongCodeSegmentCount_Throws()
    {
        var bytes = Sample("Worksheet1.xls");
        bytes[13] = 1;
        Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(bytes));
    }

    [Fact]
    public void Read_TruncatedFile_Throws()
    {
        var bytes = Sample("Expenses.xls");
        Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(bytes.AsSpan(0, 2000)));
    }

    [Fact]
    public void Read_CorruptFormula_ThrowsOnAccess()
    {
        var bytes = Sample("REPLACE.xlm");
        var document = ExcelDocument.Read(bytes);
        var formula = document.Cells["B20"].Formula!;

        // Find the formula's record in the file and break its expression length.
        int recordOffset = IndexOf(bytes, formula.RawValue.ToArray().Concat(formula.Reserved.ToArray()).Concat(new[] { formula.StoredColumn }).ToArray());
        Assert.True(recordOffset > 0);
        bytes[recordOffset + 14] = 0xFF;

        var corrupt = ExcelDocument.Read(bytes);
        Assert.Throws<ExcelFormatException>(() => corrupt.Cells["B20"].Formula);
        Assert.Throws<ExcelFormatException>(() => corrupt.ToArray());
    }

    [Fact]
    public void Read_TruncatedNumber_ThrowsOnEveryAccessPath()
    {
        var source = ExcelDocument.CreateWorksheet();
        source.Cells["A1"].SetNumber(42.5);
        var bytes = source.ToArray();
        var tableOffset = ExcelFileLayout.Read(bytes).GetSectionOffset(ExcelSection.CellTable);
        var relativeOffset = bytes.Length - 4 - ExcelFileLayout.Read(bytes).GetSectionOffset(ExcelSection.CellData);

        bytes[tableOffset] = (byte)(relativeOffset >> 12);
        bytes[tableOffset + 1] = (byte)(relativeOffset >> 4);
        bytes[tableOffset + 2] = (byte)((bytes[tableOffset + 2] & 0x0F) | ((relativeOffset & 0xF) << 4));

        var document = ExcelDocument.Read(bytes);
        Assert.Throws<ExcelFormatException>(() => document.Cells["A1"].NumberValue);
        Assert.Throws<ExcelFormatException>(() => document.Cells["A1"].TryGetNumber(out _));
        var references = document.Cells.EnumerateNonEmpty();
        Assert.True(references.MoveNext());
        var reference = references.Current;
        Assert.Throws<ExcelFormatException>(() => reference.TryGetNumber(out _));
        Assert.Throws<ExcelFormatException>(() => document.Cells["A1"].Value);
        Assert.Throws<ExcelFormatException>(() => document.ToArray());
    }

    [Fact]
    public void FileLayout_DescribesSample()
    {
        var bytes = Sample("Expenses.xls");
        var layout = ExcelFileLayout.Read(bytes);

        Assert.Equal(bytes.Length, layout.FileLength);
        Assert.Equal(0x3B8, layout.DataOffset);
        Assert.Equal(layout.DataOffset + 2, layout.StateOffset);
        Assert.Equal(layout.StateOffset + ExcelSheetState.Size, layout.KindOffset);
        Assert.Equal(layout.KindOffset + 1, layout.SectionTableOffset);
        Assert.Equal(0x653, layout.ExtraHeaderOffset);
        Assert.Equal(0, layout.ExtraHeaderLength);
        Assert.Equal([(0x18, 906), (0x3A2, 22)], layout.CodeSegments);

        int previousEnd = 0;
        foreach (var section in Enum.GetValues<ExcelSection>())
        {
            int offset = layout.GetSectionOffset(section);
            int length = layout.GetSectionLength(section);
            Assert.InRange(offset, 0, bytes.Length);
            Assert.InRange(length, 0, bytes.Length - offset);
            Assert.True(offset >= previousEnd);
            previousEnd = offset + length;
        }

        Assert.Equal(bytes.Length, previousEnd);
    }

    [Fact]
    public void FileLayout_RejectsMalformedMetadata()
    {
        var tooShort = Sample("Expenses.xls").AsSpan(0, 24).ToArray();
        Assert.Throws<ExcelFormatException>(() => ExcelFileLayout.Read(tooShort));

        var badMagic = Sample("Expenses.xls");
        badMagic[0] = 0;
        Assert.Throws<ExcelFormatException>(() => ExcelFileLayout.Read(badMagic));

        var badOffset = Sample("Expenses.xls");
        badOffset[22] = badOffset[23] = 0xFF;
        Assert.Throws<ExcelFormatException>(() => ExcelFileLayout.Read(badOffset));
    }

    [Fact]
    public void Read_InvalidWindowMarker_ThrowsFormatException()
    {
        var bytes = ExcelDocument.CreateWorksheet().ToArray();
        var layout = ExcelFileLayout.Read(bytes);
        bytes[layout.GetSectionOffset(ExcelSection.Windows)] = 1;
        Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(bytes));
    }

    [Fact]
    public void Read_ChartFile_Throws()
    {
        var chart = Charts.ChartDocument.Create(Charts.ChartType.Line).ToArray();
        Assert.Throws<ExcelFormatException>(() => ExcelDocument.Read(chart));
    }

    [Fact]
    public void Expenses_DateAndCurrencyColumns()
    {
        var document = Load("Expenses.xls");
        for (int row = 10; row < 30; row++)
        {
            var date = document.Cells[row, 0];
            if (!date.HasValue)
            {
                continue;
            }

            Assert.Equal(10, date.NumberFormat);
            Assert.True(date.TryGetNumber(out double serial));
            Assert.Equal(1984, ExcelValueFormatter.ToDateTime(serial).Year);
            Assert.Equal(5, document.Cells[row, 2].NumberFormat);
        }
    }

    [Fact]
    public void DbForm_NamesAndMacroCalls()
    {
        var document = Load("DB.FORM.xlm");
        Assert.Equal(40, document.Names.Count);
        var add = document.Names.Single(n => n.Name == "ADD");
        Assert.True(add.IsCommand);
        Assert.Equal((byte)'a', add.ShortcutKey);
        Assert.Equal("$A$34", add.GetDefinitionText(document));
        Assert.Equal("=GET.SELECTION()", document.Cells["A28"].Formula!.ToText(document));
        Assert.Equal("=ECHO(TRUE)", document.Cells["B34"].Formula!.ToText(document));
        Assert.Equal(145, document.State.GetColumnWidth(1));
    }

    [Fact]
    public void Commands_HasExpectedHeaderAndNames()
    {
        var document = Load("COMMANDS.xlm");
        Assert.Equal(ExcelDocumentKind.MacroSheet, document.Kind);
        Assert.Empty(document.ExtraHeader);
        Assert.All(document.Names, n => Assert.False(string.IsNullOrEmpty(n.Name)));
        Assert.True(document.Windows[0].DisplayFormulas);
    }

    private static int IndexOf(byte[] haystack, byte[] needle)
    {
        for (int i = 0; i <= haystack.Length - needle.Length; i++)
        {
            if (haystack.AsSpan(i, needle.Length).SequenceEqual(needle))
            {
                return i;
            }
        }

        return -1;
    }
}
