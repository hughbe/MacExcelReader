using System.Buffers;
using System.Text;

namespace MacExcelReader.Tests;

/// <summary>
/// Checks the SYLK export against the files Excel 1.00 itself exported from the samples (spec Appendix C).
/// </summary>
public class SylkTests
{
    private static readonly string SylkDirectory = Path.Combine(AppContext.BaseDirectory, "Data", "Sylk");

    public static TheoryData<string> Exports()
        => new(Directory.GetFiles(SylkDirectory, "*.slk").Select(Path.GetFileNameWithoutExtension).Order(StringComparer.Ordinal)!);

    [Theory]
    [MemberData(nameof(Exports))]
    public void ExportMatchesExcel100(string sample)
    {
        var document = ExcelDocument.Load(TestData.SamplePath(sample));

        // The reference style and calculation mode are global settings (spec §6), and Excel exported every sample
        // with A1 references and automatic calculation, whatever the file recorded.
        document.State.ReferenceStyle = ExcelReferenceStyle.A1;
        document.State.CalculationMode = ExcelCalculationMode.Automatic;
        var output = new ArrayBufferWriter<byte>();
        document.ExportSylk(output);

        byte[] expected = File.ReadAllBytes(Path.Combine(SylkDirectory, sample + ".slk"));
        if (!output.WrittenSpan.SequenceEqual(expected))
        {
            var ours = Lines(output.WrittenSpan, sample);
            var excel = Lines(expected, sample);
            int line = 0;
            while (line < ours.Length && line < excel.Length && ours[line] == excel[line])
            {
                line++;
            }

            if (line == ours.Length && line == excel.Length)
            {
                return;
            }

            Assert.Fail($"Line {line + 1}: Excel {(line < excel.Length ? excel[line] : "(end)")} | ours {(line < ours.Length ? ours[line] : "(end)")}");
        }
    }

    [Theory]
    [InlineData("caf\u00E9", "K\"caf\u001BNBe\"")]        // é: ESC N, acute, e
    [InlineData("\u00C6on", "K\"\u001BNaon\"")]          // Æ: ESC N a
    [InlineData("\u00A35", "K\"\u001B*35\"")]           // £ (A3): ESC and its nibbles
    [InlineData("a;b", "K\"a;;b\"")]                     // a field's ';' is doubled
    [InlineData("say \"hi\"", "K\"say \"hi\"\"")]       // quotes are not
    public void TextIsEscapedAsExcelDoes(string text, string expected)
    {
        var sheet = ExcelDocument.CreateWorksheet();
        sheet.Cells["A1"].SetText(text);
        // B;Y1;X1 has already set the position (spec Appendix C).
        Assert.Contains("\rB;Y1;X1\rO;L\rC;" + expected + "\r", Export(sheet));
    }

    [Fact]
    public void EditedCellsAreWrittenFromTheirObjects()
    {
        var sheet = ExcelDocument.Load(TestData.SamplePath("Expenses"));
        sheet.Cells["B11"].SetText("rent");
        sheet.Cells["C11"].SetValue(-2.5);
        var formula = ExcelFormula.Parse("=C11*2+$C$12");
        formula.SetResult("text result");
        sheet.Cells["E11"].SetFormula(formula);
        sheet.Cells["E11"].Bold = true;
        string text = Export(sheet);
        Assert.Contains("C;X2;N;K\"rent\"\r", text);
        Assert.Contains("C;N;K-2.5\r", text);
        Assert.Contains("F;SD;X5\rC;K\"text result\";ERC[-2]*2+R12C3\r", text);
    }

    [Fact]
    public void NewSheetHasExcelsDefaults()
    {
        var sheet = ExcelDocument.CreateWorksheet();
        sheet.Cells["B2"].SetValue(1);
        string text = Export(sheet);
        Assert.StartsWith("ID;PXL;N;E\rP;PGeneral\r", text);
        Assert.Contains("\rF;P0;DG0G10;N3 10\rB;Y2;X2\rO;L\r", text);
        Assert.EndsWith("C;K1\rE\r", text);
    }

    private static string Export(ExcelDocument document)
    {
        var output = new ArrayBufferWriter<byte>();
        document.ExportSylk(output);
        return Encoding.Latin1.GetString(output.WrittenSpan);
    }

    // Excel recalculated these samples' formulas while exporting them: they use NOW() or documents that were not
    // open, which gave #REF!, and the cells that depend on those. Their formulas' values are left out.
    private static readonly string[] s_recalculated =
        ["BREAKEVEN", "CLOCK.WS", "COMMISSION TEMPLATE", "INVOICE TEMPLATE", "QUARTER TEMPLATE", "QUARTER3"];

    private static string[] Lines(ReadOnlySpan<byte> bytes, string sample)
    {
        var lines = Encoding.Latin1.GetString(bytes).Split('\r');
        if (!s_recalculated.Contains(sample))
        {
            return lines;
        }

        for (int i = 0; i < lines.Length; i++)
        {
            int k = lines[i].IndexOf(";K", StringComparison.Ordinal);
            int e = lines[i].IndexOf(";E", StringComparison.Ordinal);
            if (lines[i].StartsWith("C;", StringComparison.Ordinal) && k >= 0 && e > k)
            {
                lines[i] = lines[i][..k] + lines[i][e..];
            }
        }

        return lines;
    }
}
