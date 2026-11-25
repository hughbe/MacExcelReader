using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// Compiles every sample formula and definition, in every release, many mutations of them and many generated
/// formulas: the compiler may only report <see cref="ExcelFormulaSyntaxException"/>, and whatever it compiles must
/// decompile. (These inputs also showed the rewritten compiler giving the old one's bytes, apart from the
/// empty-intersection rule Excel showed in `Empty Intersection Ranges`.)
/// </summary>
public class FormulaParserFuzzTests
{
    public static TheoryData<string> Sheets => AllSheets();

    private static readonly ExcelVersion[] Versions = [ExcelVersion.Excel100, ExcelVersion.Excel104, ExcelVersion.Excel150];

    private enum Kind { Cell, Definition, Array }

    private static int s_successes;
    private static int s_failures;

    // Compiles the text; it must compile, or fail with a syntax error, and whatever compiles must decompile.
    private static void Compare(string text, ExcelDocument? document, ExcelVersion version, Kind kind)
    {
        ushort start = document?.FunctionCount ?? 0;
        try
        {
            switch (kind)
            {
                case Kind.Cell:
                    _ = ExcelFormulaParser.Parse(text, document, version).ToText(document);
                    break;
                case Kind.Definition:
                    var (definition, arrays) = ExcelFormulaParser.ParseNameDefinition(text, document, version);
                    var name = new ExcelDefinedName("Fuzz", definition);
                    name.ArrayConstants.AddRange(arrays);
                    _ = name.GetDefinitionText(document);
                    break;
                default:
                    _ = ExcelFormulaParser.ParseArrayFormula(text, document, version);
                    break;
            }

            Interlocked.Increment(ref s_successes);
        }
        catch (ExcelFormulaSyntaxException)
        {
            Interlocked.Increment(ref s_failures);
        }
        finally
        {
            if (document is not null)
            {
                document.FunctionCount = start;
            }
        }
    }

    // The formula and definition texts of a sample, as the decompiler gives them.
    private static List<(string Text, Kind Kind)> Texts(ExcelDocument document)
    {
        var texts = new List<(string, Kind)>();
        foreach (var entry in document.Cells.NonEmptyCells)
        {
            if (entry.GetCell().Formula is not { } formula)
            {
                continue;
            }

            string text = formula.ToText(document);
            if (text.StartsWith("{=", StringComparison.Ordinal) && text.EndsWith('}'))
            {
                texts.Add((text[1..^1], Kind.Array));
            }
            else
            {
                texts.Add((text, Kind.Cell));
            }
        }

        foreach (var name in document.Names)
        {
            if (name.Definition.Length > 0)
            {
                texts.Add(("=" + name.GetDefinitionText(document), Kind.Definition));
            }
        }

        return texts;
    }

    [Theory]
    [MemberData(nameof(Sheets))]
    public void SampleFormulas_Compile(string name)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(name)));
        foreach (var (text, kind) in Texts(document))
        {
            foreach (var version in Versions)
            {
                Compare(text, document, version, kind);
                Compare(text, null, version, kind);
            }
        }
    }

    // Characters formulas are made of, to insert at random.
    private const string Alphabet = "=+-*/^&%:,;!(){}\"'#$ .0123456789AaBbCcEeIiNnRrSsTtUuVv?_\\<>";

    [Theory]
    [MemberData(nameof(Sheets))]
    public void MutatedFormulas_CompileOrFailCleanly(string name)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(name)));
        var texts = Texts(document);
        var random = new Random(name.GetHashCode(StringComparison.Ordinal));
        foreach (var (text, kind) in texts)
        {
            for (int i = 0; i < 40; i++)
            {
                var chars = new System.Text.StringBuilder(text);
                int changes = 1 + random.Next(3);
                for (int c = 0; c < changes; c++)
                {
                    int at = random.Next(chars.Length + 1);
                    switch (random.Next(4))
                    {
                        case 0 when chars.Length > 0 && at < chars.Length:
                            chars.Remove(at, 1);
                            break;
                        case 1:
                            chars.Insert(at, Alphabet[random.Next(Alphabet.Length)]);
                            break;
                        case 2 when chars.Length > 1 && at < chars.Length:
                            int length = random.Next(1, Math.Min(12, chars.Length - at) + 1);
                            chars.Insert(random.Next(chars.Length + 1), chars.ToString(at, length));
                            break;
                        default:
                            if (at < chars.Length)
                            {
                                chars[at] = Alphabet[random.Next(Alphabet.Length)];
                            }

                            break;
                    }
                }

                var version = Versions[random.Next(Versions.Length)];
                Compare(chars.ToString(), document, version, kind);
            }
        }
    }

    // Fragments to glue together into formulas that exercise every rule: references of every form, names, external
    // references, functions with every argument shape, operators, array constants, strings with quotes, errors.
    private static readonly string[] Fragments =
    [
        "A1", "$A$1", "A$1", "$A1", "B2:C3", "IV16384", "1:3", "$1:$2", "A:B", "$A:$C", "A1:B2 B1:C3", "(A1,B2)", "(A1:B2,C3)",
        "1", "0", "65535", "65536", "1.5", "-1", "+1", "1E10", ".5", "2%", "\"text\"", "\"a\"\"b\"", "\"\"", "TRUE", "false",
        "#N/A", "#REF!", "#DIV/0!", "#NULL!", "{1,2;3,4}", "{\"a\",TRUE;#N/A,-2}", "{1}", "IF(A1,1,2)", "IF(A1,B1)",
        "CHOOSE(A1,B1,C1,D1)", "SUM(A1)", "SUM(A1,B2)", "SUM(A1:B2*2)", "NOW()", "RAND()", "INDEX(A1:C3,1,1)", "INDEX(A1:C3,1,1):C3",
        "ROW()", "COLUMN(A1)", "TRANSPOSE(A1:B2)", "LINEST(A1:A3)", "DATA.DELETE?()", "RUN?()", "ARGUMENT(\"x\")", "OFFSET(A1,1,1)",
        "Total", "Area", "TwoAreas", "Macro1(1)", "Budget!A1", "Budget!$A$1:$B$2", "Budget!Rate", "'My Budget'!Q1", "!Database",
        "!A1", "!1:2", "LEFT(\"x\",1)", "FORMULA(\"=1\")", "ACTIVATE(\"x\")", "DCOUNT(A1:B2,,C1:C2)", "TABLE(A1,B1)",
    ];

    private static readonly string[] Operators = ["+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">=", ":", " ", ","];

    [Fact]
    public void GeneratedFormulas_CompileOrFailCleanly()
    {
        // A sheet whose names and external documents the fragments use, like the formula lists' setup.
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(AllSamples().First(s => s.EndsWith("Formulas 1.0 Sheet 2", StringComparison.Ordinal)))));
        var random = new Random(1234);
        for (int i = 0; i < 20000; i++)
        {
            var text = new System.Text.StringBuilder("=");
            int terms = 1 + random.Next(5);
            for (int t = 0; t < terms; t++)
            {
                if (t > 0)
                {
                    text.Append(Operators[random.Next(Operators.Length)]);
                }

                string fragment = Fragments[random.Next(Fragments.Length)];
                if (random.Next(6) == 0)
                {
                    fragment = "(" + fragment + ")";
                }

                if (random.Next(8) == 0)
                {
                    fragment = Fragments[random.Next(Fragments.Length)].Split('(')[0] + "(" + fragment + ")";
                }

                text.Append(fragment);
            }

            var version = Versions[random.Next(Versions.Length)];
            var kind = (Kind)random.Next(3);
            Compare(text.ToString(), random.Next(4) == 0 ? null : document, version, kind);
        }

        // Both outcomes must be well represented for the comparison to mean anything.
        Assert.True(s_successes > 3000 && s_failures > 3000, $"{s_successes} compiled, {s_failures} failed");
    }

    [Theory]
    [InlineData("=" + "(" + "(" + "1" + ")" + ")")]
    [InlineData("=(A1 B2):C3")]
    [InlineData("=(A1 B2,C3)")]
    [InlineData("=\"unterminated")]
    [InlineData("='unterminated")]
    [InlineData("=")]
    [InlineData("")]
    [InlineData("   =  1 +  2 ")]
    [InlineData("={1,2;3}")]
    [InlineData("={1,2;3,x}")]
    [InlineData("=SUM(,)")]
    [InlineData("=1E400")]
    [InlineData("=#ERR")]
    [InlineData("=A1:Name")]
    [InlineData("=''!A1")]
    [InlineData("='My ''Budget'!A1")]
    public void EdgeCases_CompileOrFailCleanly(string text)
    {
        var document = ExcelDocument.Read(File.ReadAllBytes(SampleFile(AllSamples().First(s => s.EndsWith("Formulas 1.0 Sheet 2", StringComparison.Ordinal)))));
        foreach (var version in Versions)
        {
            foreach (var kind in Enum.GetValues<Kind>())
            {
                Compare(text, document, version, kind);
                Compare(text, null, version, kind);
            }
        }
    }

    private static readonly string[] AllocationFormulas =
        ["=1", "=A1", "=$A$1:$B$2", "=IF(A1>0,SUM($A$1:$A$5)*2,-1)", "=\"a\"\"b\"", "=(A1:B2,C3)", "={1,2;3,4}", "=CHOOSE(A1,1,2,3)", "=NOW()"];

    [Fact]
    public void Parse_AllocatesOnlyTheFormula()
    {
        // The formula and its record (about 150 bytes for these); nothing for the compiler itself.
        foreach (string text in AllocationFormulas)
        {
            ExcelFormula.Parse(text);
            ExcelFormula.Parse(text);
            long before = GC.GetAllocatedBytesForCurrentThread();
            ExcelFormula.Parse(text);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            Assert.True(allocated <= 160, $"{text}: {allocated} bytes");
        }
    }
}
