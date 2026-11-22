using System.Buffers.Binary;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class FormulaDecompilerTests
{
    private static string Decompile(byte[] expression, ExcelDocument? document = null, params ushort[] fixups)
        => FormulaDecompiler.Decompile(expression, fixups, document, fixupsInPlace: false);

    [Theory]
    [InlineData(0x03, "A1+2")]
    [InlineData(0x04, "A1-2")]
    [InlineData(0x05, "A1*2")]
    [InlineData(0x06, "A1/2")]
    [InlineData(0x07, "A1^2")]
    [InlineData(0x08, "A1&2")]
    [InlineData(0x09, "A1<2")]
    [InlineData(0x0A, "A1<=2")]
    [InlineData(0x0B, "A1=2")]
    [InlineData(0x0C, "A1>=2")]
    [InlineData(0x0D, "A1>2")]
    [InlineData(0x0E, "A1<>2")]
    public void BinaryOperators(byte op, string expected)
    {
        Assert.Equal(expected, Decompile(Concat(Reference(0, 0), Integer(2), [op])));
    }

    [Fact]
    public void UnaryOperatorsPercentAndParentheses()
    {
        Assert.Equal("-A1", Decompile(Concat(Reference(0, 0), [0x13])));
        Assert.Equal("+A1", Decompile(Concat(Reference(0, 0), [0x12])));
        Assert.Equal("A1%", Decompile(Concat(Reference(0, 0), [0x14])));
        Assert.Equal("(A1+2)*3", Decompile(Concat(Reference(0, 0), Integer(2), [0x03, 0x15], Integer(3), [0x05])));
    }

    [Fact]
    public void Constants()
    {
        Assert.Equal("\"say \"\"hi\"\"\"", Decompile([0x17, 8, .."say \"hi\""u8]));
        Assert.Equal("TRUE", Decompile([0x1D, 1]));
        Assert.Equal("FALSE", Decompile([0x1D, 0]));
        Assert.Equal("#DIV/0!", Decompile([0x1C, 7]));
        Assert.Equal("#N/A", Decompile([0x1C, 42]));
        Assert.Equal("65535", Decompile(Integer(65535)));

        var number = new byte[9];
        number[0] = 0x1F;
        BinaryPrimitives.WriteDoubleBigEndian(number.AsSpan(1), 3.5);
        Assert.Equal("3.5", Decompile(number));
        Assert.Equal("{array}", Decompile([0x20, 0, 0, 0, 0, 0, 0]));
    }

    [Theory]
    [InlineData(true, true, "C5")]
    [InlineData(false, true, "C$5")]
    [InlineData(true, false, "$C5")]
    [InlineData(false, false, "$C$5")]
    public void References_ShowAbsoluteParts(bool rowRelative, bool columnRelative, string expected)
    {
        Assert.Equal(expected, Decompile(Reference(4, 2, rowRelative, columnRelative)));
    }

    [Fact]
    public void Areas()
    {
        // Area A1:B10, both corners absolute.
        Assert.Equal("$A$1:$B$10", Decompile([0x25, 0x00, 0x00, 0x00, 0x09, 0x00, 0x01]));
        Assert.Equal("#REF!", Decompile([0x2A, 0, 0, 0]));
        Assert.Equal("#REF!", Decompile([0x2B, 0, 0, 0, 0, 0, 0]));
    }

    [Fact]
    public void Functions_FixedAndVariableArguments()
    {
        // LEN(A1): fixed-argument function 32.
        Assert.Equal("LEN(A1)", Decompile(Concat(Reference(0, 0), [0x41, 32])));

        // SUM(A1,2,3): variable-argument SUM with 3 arguments.
        Assert.Equal("SUM(A1,2,3)", Decompile(Concat(Reference(0, 0), Integer(2), Integer(3), [0x42, 3, 4])));

        // NA() and PI(): no arguments.
        Assert.Equal("NA()", Decompile([0x21, 10]));
    }

    [Fact]
    public void Functions_MissingArgumentIsEmpty()
    {
        // IF(A1,,2)
        Assert.Equal("IF(A1,,2)", Decompile(Concat(Reference(0, 0), [0x16], Integer(2), [0x42, 3, 1])));
    }

    [Fact]
    public void UserMacroCall_UsesFirstArgumentAsFunction()
    {
        var document = ExcelDocument.CreateMacroSheet();
        document.Names.Add(new ExcelDefinedName("MYMACRO", [0x2C, 0x00, 0x00, 0x00], ExcelDefinedName.ReferenceFlag));
        ushort index = (ushort)document.Names.Count;

        // MYMACRO(5): name token, argument, call function 255 with 2 "arguments".
        var expression = Concat([0x23, 0, 0, 0, 0, 0, 0, 0], Integer(5), [0x42, 2, 255]);
        Assert.Equal("MYMACRO(5)", Decompile(expression, document, index));
    }

    [Fact]
    public void Names_ResolveThroughFixups()
    {
        var document = ExcelDocument.CreateWorksheet();
        Assert.Equal("Database", Decompile([0x23, 0, 0, 0, 0, 0, 0, 0], document, 2));
        Assert.Equal("NAME99", Decompile([0x23, 0, 0, 0, 0, 0, 0, 0], document, 99));
    }

    [Fact]
    public void NameDefinitions_ReadIndicesInPlace()
    {
        var document = ExcelDocument.CreateWorksheet();
        var definition = new byte[] { 0x23, 0x00, 0x03, 0, 0, 0, 0, 0 };
        Assert.Equal("Print_Area", FormulaDecompiler.Decompile(definition, [], document, fixupsInPlace: true));
    }

    [Fact]
    public void ExternalReferences_PrefixTheDocumentName()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.ExternalDocuments.Add(new ExcelExternalDocument("Budget"));
        var expression = Concat([0x1A, 0, 0, 0, 0, 0, 0, 0], Reference(0, 0), [0x1B, 0, 0, 0]);
        Assert.Equal("Budget!A1", Decompile(expression, document, 1));

        document.ExternalDocuments[0].Name = "Budget Plan's";
        Assert.Equal("'Budget Plan''s'!A1", Decompile(expression, document, 1));
    }

    [Fact]
    public void PrivateFallbacksRemainDefined()
    {
        var method = typeof(FormulaTokens).GetMethod(
            "BinaryOperatorText", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        Assert.Equal("?", method.Invoke(null, [(byte)0]));
        Assert.Equal(" ", method.Invoke(null, [(byte)0x0F]));
        Assert.Equal(",", method.Invoke(null, [(byte)0x10]));
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    [InlineData(255, "IV")]
    public void ColumnName_Formats(int column, string expected)
    {
        Assert.Equal(expected, FormulaDecompiler.ColumnName(column));
    }

    [Fact]
    public void Replace_EveryFormulaMatchesExpectedText()
    {
        var document = Load("REPLACE.xlm");
        var expected = new Dictionary<string, string>
        {
            ["B20"] = "=IF(NOT(B19),RETURN())",
            ["B23"] = "=LEN(B19)",
            ["B26"] = "=IF(B25=\"Y\",TRUE,IF(B25<>\"N\",RETURN()))",
            ["B30"] = "=SET.VALUE(B33,-B24+1)",
            ["B33"] = "=SEARCH(B19,B32,B33+B24)",
            ["B34"] = "=IF(NOT(ISERROR(B33)),IF(B26,GOTO(SWAP),GOTO(ASK)))",
            ["B42"] = "=B42+1",
            ["B45"] = "=MID(B32,1,B33-1)&B21&MID(B32,B33+B23,B44-B33-B23+1)",
            ["B46"] = "=FORMULA(B45)",
            ["B48"] = "=RETURN()",
        };

        foreach (var (address, text) in expected)
        {
            Assert.Equal(text, document.Cells[address].Formula!.ToText(document));
        }
    }
}
