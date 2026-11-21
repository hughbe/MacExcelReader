using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class FormulaTokensTests
{
    [Theory]
    [InlineData(0x24, 0x24)]
    [InlineData(0x44, 0x24)]
    [InlineData(0x64, 0x24)]
    [InlineData(0x41, 0x21)]
    [InlineData(0x03, 0x03)]
    public void Normalize_FoldsClassBits(byte raw, byte expected)
    {
        Assert.Equal(expected, FormulaTokens.Normalize(raw));
    }

    [Fact]
    public void Enumerate_ReturnsOffsetsAndLengths()
    {
        var expression = Concat(Reference(0, 0), [0x17, 3, (byte)'a', (byte)'b', (byte)'c'], [0x08], [0x19, 0x04, 1, 9, 9], [0x1D, 1]);
        var tokens = new List<FormulaToken>();
        foreach (var token in FormulaTokens.Enumerate(expression))
        {
            tokens.Add(token);
        }

        Assert.Equal(
            [(0, 0x24, 4), (4, 0x17, 5), (9, 0x08, 1), (10, 0x19, 5), (15, 0x1D, 2)],
            tokens.Select(t => (t.Offset, (int)t.Code, t.Length)));
        Assert.Equal(5, FormulaTokens.CountTokens(expression));
    }

    [Theory]
    [InlineData(new byte[] { 0x44, 0xC0 })]
    [InlineData(new byte[] { 0x17, 5, 1, 2 })]
    [InlineData(new byte[] { 0x1F, 0, 0, 0 })]
    public void Enumerate_TruncatedToken_Throws(byte[] expression)
    {
        Assert.Throws<ExcelFormatException>(() => FormulaTokens.CountTokens(expression));
    }

    [Fact]
    public void ExcelFormula_RejectsInvalidExpression()
    {
        Assert.Throws<ExcelFormatException>(() => new ExcelFormula([0x25, 0, 0]));
        Assert.Throws<ArgumentException>(() => new ExcelFormula(new byte[256]));
    }

    [Fact]
    public void MeasureFixups_CountsNameAndArrayBlocks()
    {
        // Two names, then an array constant (1 column × 2 rows: a number and a string).
        var expression = Concat([0x23, 0, 0, 0, 0, 0, 0, 0], [0x23, 0, 0, 0, 0, 0, 0, 0], [0x20, 0, 0, 0, 0, 0, 0], [0x03, 0x03]);
        byte[] fixups = [0, 1, 0, 2, 1, 0, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, (byte)'h', (byte)'i'];
        Assert.Equal(fixups.Length, FormulaTokens.MeasureFixups(expression, fixups.Concat(new byte[] { 99, 99 }).ToArray()));
    }

    [Fact]
    public void MeasureFixups_Truncated_Throws()
    {
        Assert.Throws<ExcelFormatException>(() => FormulaTokens.MeasureFixups([0x23, 0, 0, 0, 0, 0, 0, 0], [0]));
    }

    [Fact]
    public void GetName_NamesEveryToken()
    {
        Assert.Equal("Reference", FormulaTokens.GetName(0x24));
        Assert.Equal("FunctionVarArgs", FormulaTokens.GetName(0x22));
        Assert.Equal("MarkedName", FormulaTokens.GetName(0x30));
        Assert.Equal("MarkedReference", FormulaTokens.GetName(0x37));
        for (byte code = 1; code < 0x38; code++)
        {
            Assert.False(string.IsNullOrEmpty(FormulaTokens.GetName(code)));
        }
    }

    [Fact]
    public void GetFixupWords_ReturnsNameIndices()
    {
        var formula = new ExcelFormula(Concat([0x23, 0, 0, 0, 0, 0, 0, 0], [0x23, 0, 0, 0, 0, 0, 0, 0], [0x03]), [0, 3, 0, 4]);
        Assert.Equal([3, 4], formula.GetFixupWords());
    }
}
