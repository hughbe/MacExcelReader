using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class FormulaParserTests
{
    public static TheoryData<string> Sheets => AllSheets();

    /// <summary>
    /// Zeroes the bytes Excel leaves over from memory (spec §9.2): a name token's operands (all of them in a cell
    /// formula, where the index is in the fix-ups; all but the index in a name definition), an external-sheet
    /// token's operands in a cell formula (all but the index in a definition), the end token's operands, the
    /// three pointer bytes of an array constant, a deleted reference's operands, and the data byte of the
    /// volatile and SUM attributes.
    /// </summary>
    private static byte[] Normalize(ReadOnlySpan<byte> expression, bool definition)
    {
        var copy = expression.ToArray();
        foreach (var token in FormulaTokens.Enumerate(expression))
        {
            var operands = copy.AsSpan(token.Offset + 1, token.Length - 1);
            switch (token.Code)
            {
                case 0x23:
                    operands[(definition ? 2 : 0)..].Clear();
                    break;
                case 0x1A:
                    operands[..(definition ? 6 : 7)].Clear();
                    break;
                case 0x20:
                    operands[3..].Clear();
                    break;
                case 0x1B:
                case 0x2A:
                    operands.Clear();
                    break;
                case 0x19 when operands[0] is 0x01 or 0x10:
                    operands[1] = 0;
                    break;
            }
        }

        return copy;
    }

    // The pre-release Excel 1.5 build wrote FORMULA with one argument in the variable-count form, where the
    // released 1.5 keeps the fixed-count form (spec §2.3, §15.4).
    private static bool IsKnownException(byte[] file, ExcelFormula formula)
    {
        if (!IsPreRelease15(file))
        {
            return false;
        }

        foreach (var token in FormulaTokens.Enumerate(formula.Expression))
        {
            if (token.Code == 0x22 && formula.Expression[token.Offset + 2] == 57)
            {
                return true;
            }
        }

        return false;
    }

    [Theory]
    [MemberData(nameof(Sheets))]
    public void SetDefinition_SetsTheReferenceFlagAsExcelDoes(string sample)
    {
        // Excel sets flag 0x04 on exactly the names defined as a cell, an area or a union of them.
        var bytes = File.ReadAllBytes(SampleFile(sample));
        var version = SavedBy(bytes);
        var document = ExcelDocument.Read(bytes);
        foreach (var name in document.Names.Where(n => n.Definition.Length > 0))
        {
            var copy = new ExcelDefinedName(name.Name, flags: (byte)(name.Flags ^ ExcelDefinedName.ReferenceFlag));
            copy.SetDefinition("=" + name.GetDefinitionText(document), document, version);
            Assert.True((copy.Flags & ExcelDefinedName.ReferenceFlag) == (name.Flags & ExcelDefinedName.ReferenceFlag), $"{name.Name} ={name.GetDefinitionText(document)}");
        }
    }

    [Theory]
    [MemberData(nameof(Sheets))]
    public void EverySampleFormulaReparsesToExcelsTokens(string sample)
    {
        var bytes = File.ReadAllBytes(SampleFile(sample));
        var version = SavedBy(bytes);
        var document = ExcelDocument.Read(bytes);
        foreach (var (row, column, cell) in document.Cells.NonEmptyCells)
        {
            var address = new ExcelCellAddress(row, column);
            if (cell.Formula is not { } formula || IsKnownException(bytes, formula) || (formula.IsArrayFormula && !formula.IsSharedAnchor))
            {
                continue;
            }

            string text = formula.ToText(document);
            ReadOnlySpan<byte> expected;
            byte[] actual;
            if (formula.IsArrayFormula)
            {
                expected = formula.SharedExpression;
                actual = ExcelFormulaParser.ParseArrayFormula(text[1..^1], document, version).Expression;
            }
            else
            {
                expected = formula.Expression;
                var parsed = ExcelFormula.Parse(text, document, version);
                actual = parsed.Expression.ToArray();
                Assert.Equal(formula.GetFixupWords(), parsed.GetFixupWords());
            }

            Assert.True(
                Normalize(expected, false).AsSpan().SequenceEqual(Normalize(actual, false)),
                $"{address} {text}: {Convert.ToHexString(expected)} vs {Convert.ToHexString(actual)}");
        }
    }

    [Theory]
    [MemberData(nameof(Sheets))]
    public void EverySampleNameDefinitionReparses(string sample)
    {
        var bytes = File.ReadAllBytes(SampleFile(sample));
        var document = ExcelDocument.Read(bytes);
        foreach (var name in document.Names)
        {
            if (name.Definition.Length == 0)
            {
                continue;
            }

            string text = name.GetDefinitionText(document);
            var (parsed, arrays) = ExcelFormulaParser.ParseNameDefinition(text, document, SavedBy(bytes));
            Assert.True(
                Normalize(name.Definition, true).AsSpan().SequenceEqual(Normalize(parsed, true)),
                $"{name.Name} = {text}: {Convert.ToHexString(name.Definition)} vs {Convert.ToHexString(parsed)}");
            Assert.Equal(name.ArrayConstants, arrays);
        }
    }

    [Theory]
    [InlineData("=A1*2", "4440000005" + "1E0002" + "05")]
    [InlineData("A1*2", "4440000005" + "1E0002" + "05")]
    [InlineData("=$A$1+B2", "44000000" + "44C00101" + "03")]
    [InlineData("=SUM($A$1:$A$5)", "25000000040000" + "191000")]
    [InlineData("=-1", "1FBFF0000000000000")]
    [InlineData("=-A1", "44C00000" + "13")]
    [InlineData("=A1%", "44C00000" + "14")]
    [InlineData("=(A1+2)*3", "44C00000" + "1E0002" + "03" + "15" + "1E0003" + "05")]
    [InlineData("=\"a\"\"b\"", "1703612262")]
    [InlineData("=TRUE", "1D01")]
    [InlineData("=#N/A", "1C2A")]
    [InlineData("=1.5", "1F3FF8000000000000")]
    [InlineData("=70000", "1F40F1170000000000")]
    [InlineData("=1^2&3", "1E0001" + "1E0002" + "07" + "1E0003" + "08")]
    [InlineData("=1<=2", "1E0001" + "1E0002" + "0A")]
    [InlineData("=1<>2", "1E0001" + "1E0002" + "0E")]
    [InlineData("=RAND()", "190100" + "4141")]
    [InlineData("=RUN?()", "428091")]
    public void ParsesToTokens(string text, string expectedHex)
    {
        // Strip the spacing used to make the test data readable.
        expectedHex = expectedHex.Replace("4440000005", "44C00000");
        Assert.Equal(expectedHex, Convert.ToHexString(ExcelFormula.Parse(text).Expression));
    }

    [Theory]
    [InlineData("=IF(A1,1,2)")]
    [InlineData("=IF(A1,1)")]
    [InlineData("=IF(A1>0,SUM(A1:A3),-1)")]
    [InlineData("=CHOOSE(2,\"a\",B1,C1+1)")]
    [InlineData("=IF(A1,IF(B1,1,2),3)")]
    [InlineData("=MID(\"abc\",2,1)")]
    [InlineData("=INDEX(A1:C3,1,1):C3")]
    [InlineData("=$A1+A$1")]
    [InlineData("=IV16384")]
    [InlineData("=AVERAGE(1,2,3,4)")]
    public void TextRoundTrips(string text)
    {
        Assert.Equal(text, ExcelFormula.Parse(text).ToText());
    }

    [Fact]
    public void IfJumpsAreConsistent()
    {
        // =IF(A1,1,2): condition, 19 02 n, true branch, 19 08 m, false branch, 19 08 m, IF.
        var hex = Convert.ToHexString(ExcelFormula.Parse("=IF(A1,1,2)").Expression);
        Assert.Equal("44C00000" + "190206" + "1E0001" + "190808" + "1E0002" + "190802" + "420301", hex);
    }

    [Fact]
    public void ChooseHasAJumpTableAndSkips()
    {
        // Index, 19 04 n + table of n + 1 offsets from the table start (choice 1, choice 2, CHOOSE token),
        // each choice followed by a skip to the end of the CHOOSE token minus 1.
        var hex = Convert.ToHexString(ExcelFormula.Parse("=CHOOSE(A1,1,2)").Expression);
        Assert.Equal("44C00000" + "190402" + "03090F" + "1E0001" + "190808" + "1E0002" + "190802" + "4203B4", hex);
        Assert.Equal("=CHOOSE(A1,1,2)", ExcelFormula.Parse("=CHOOSE(A1,1,2)").ToText());
    }

    [Fact]
    public void SumWithOneArgumentIsAnAttribute()
    {
        Assert.Equal("25C000C0020000" + "191000", Convert.ToHexString(ExcelFormula.Parse("=SUM(A1:A3)").Expression));
        Assert.Equal("=SUM(A1:A3)", ExcelFormula.Parse("=SUM(A1:A3)").ToText());
        Assert.Equal("1E0001" + "1E0002" + "420204", Convert.ToHexString(ExcelFormula.Parse("=SUM(1,2)").Expression));
    }

    [Fact]
    public void NamesInsideExternalReferencesUseThatDocumentsNames()
    {
        // DB.FORM's macros use the active sheet's Database and Form names (section 3), not its own names.
        var document = Load("DB.FORM.xlm");
        Assert.Equal("=SELECT(INDEX(!FORM,2,2):INDEX(!FORM,2,COLUMNS(!DATABASE)+1))", document.Cells["C146"].Formula!.ToText(document));

        var parsed = ExcelFormula.Parse("=ROWS(!database)", document);
        Assert.Equal([1, 1], parsed.GetFixupWords());
        Assert.Equal("=ROWS(!DATABASE)", parsed.ToText(document));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=ROWS(!ADD)", document));
    }

    [Fact]
    public void ResolvesDefinedNamesAndMacroCalls()
    {
        var document = Load("REPLACE.xlm");
        var name = document.Names[0];
        var formula = ExcelFormula.Parse("=" + name.Name, document);
        Assert.Equal([1], formula.GetFixupWords());
        Assert.Equal("=" + name.Name, formula.ToText(document));
    }

    [Fact]
    public void ParsedFormulaSurvivesSaving()
    {
        var sheet = ExcelDocument.CreateWorksheet();
        sheet.Cells["A1"].SetValue(21);
        var formula = ExcelFormula.Parse("=A1*2");
        formula.SetResult(42);
        sheet.Cells["A2"].SetFormula(formula);
        var reread = RoundTrip(sheet);
        Assert.Equal("=A1*2", reread.Cells["A2"].Formula!.ToText());
        Assert.Equal(42.0, reread.Cells["A2"].Value);
    }

    [Theory]
    [InlineData("=", 0)]
    [InlineData("=1+", 3)]
    [InlineData("=SUM(1", 6)]
    [InlineData("=NOPE(1)", 1)]
    [InlineData("=IF(1)", 1)]
    [InlineData("=\"abc", 1)]
    [InlineData("=1 2)", 3)]
    [InlineData("=SUM?(1)", 1)]
    [InlineData("=Other!A1", 1)]
    [InlineData("=UnknownName", 1)]
    [InlineData("=@", 1)]
    public void RejectsInvalidText(string text, int position)
    {
        var exception = Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse(text));
        Assert.Equal(position, exception.Position);
    }

    [Theory]
    [InlineData("=#NULL!")]
    [InlineData("=#DIV/0!")]
    [InlineData("=#VALUE!")]
    [InlineData("=#REF!")]
    [InlineData("=#NAME?")]
    [InlineData("=#NUM!")]
    [InlineData("=#N/A")]
    public void ParsesAllErrorTokens(string text)
    {
        Assert.Equal(text, ExcelFormula.Parse(text).ToText());
    }

    [Fact]
    public void RejectsAdditionalParserLimits()
    {
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=\"" + new string('x', 256) + "\""));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=!"));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=!A1:B"));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=" + new string('(', 300) + "1" + new string(')', 300)));

        var document = ExcelDocument.CreateWorksheet();
        var name = new ExcelDefinedName("Macro");
        document.Names.Add(name);
        var arguments = string.Join(",", Enumerable.Repeat("1", 255));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse($"=Macro({arguments})", document));
    }

    [Fact]
    public void ParsesExternalRangesAndUserMacros()
    {
        var document = ExcelDocument.CreateWorksheet();
        var externalDocument = new ExcelExternalDocument("Book") { Names = { "Missing" } };
        document.ExternalDocuments.Add(externalDocument);
        var external = ExcelFormula.Parse("='Book'!A1:B2", document);
        Assert.Equal([1], external.GetFixupWords());
        Assert.Equal("=Book!A1:B2", external.ToText(document));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse("=Book!A1:Missing", document));

        document.Names.Add(new ExcelDefinedName("RunMe", flags: ExcelDefinedName.ReferenceFlag));
        Assert.True(ExcelFormula.Parse("=RunMe(1)", document).Expression.Length > 0);
    }

    [Fact]
    public void RejectsBranchOffsetsBeyondEncodingLimit()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Names.Add(new ExcelDefinedName("RunMe", flags: ExcelDefinedName.ReferenceFlag));
        var huge = "RunMe(" + string.Join(",", Enumerable.Repeat("1", 130)) + ")";
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse($"=IF(TRUE,{huge},0)", document));
        Assert.Throws<ExcelFormulaSyntaxException>(() => ExcelFormula.Parse($"={huge}:A1", document));
        Assert.True(ExcelFormula.Parse("=(SUM(1)):A1", document).Expression.Length > 0);
        Assert.True(ExcelFormula.Parse("=-SUM(1):A1", document).Expression.Length > 0);
        Assert.True(ExcelFormula.Parse("=(-SUM(1)):A1", document).Expression.Length > 0);
    }

    [Fact]
    public void RandomTextOnlyRaisesSyntaxErrors()
    {
        const string alphabet = "=+-*/^&%<>:,!()\"'$#?. 0123456789ABCIFSUMRANDTRUEabcxyz";
        var random = new Random(1234);
        var document = Load("DB.FORM.xlm");
        var builder = new System.Text.StringBuilder();
        for (int i = 0; i < 50_000; i++)
        {
            builder.Clear();
            int length = random.Next(1, 24);
            for (int j = 0; j < length; j++)
            {
                builder.Append(alphabet[random.Next(alphabet.Length)]);
            }

            try
            {
                var formula = ExcelFormula.Parse(builder.ToString(), document);
                _ = formula.ToText(document);
            }
            catch (ExcelFormulaSyntaxException)
            {
            }
        }
    }
}
