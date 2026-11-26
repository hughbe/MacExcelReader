using System.Buffers;
using System.Text;

namespace MacExcelReader.Tests;

/// <summary>
/// Checks the SYLK export against the files Excel 1.00 itself exported from the samples (docs/Sylk.md).
/// </summary>
public class SylkTests
{
    private static readonly string DataDirectory = Path.Combine(AppContext.BaseDirectory, "Data");

    public static TheoryData<string> Exports() => Samples("Sylk");

    public static TheoryData<string> Exports15() => Samples("Sylk 1.5");

    private static TheoryData<string> Samples(string folder)
        => new(Directory.GetFiles(Path.Combine(DataDirectory, folder), "*.slk").Select(Path.GetFileNameWithoutExtension).Order(StringComparer.Ordinal)!);

    [Theory]
    [MemberData(nameof(Exports))]
    public void ExportMatchesExcel100(string sample) => AssertMatches(sample, "Sylk", ExcelVersion.Excel100);

    [Theory]
    [MemberData(nameof(Exports15))]
    public void ExportMatchesExcel15(string sample) => AssertMatches(sample, "Sylk 1.5", ExcelVersion.Excel150);

    [Theory]
    [MemberData(nameof(Exports))]
    public async Task ExportAsync_WritesTheSameBytes(string sample)
    {
        var document = ExcelDocument.Load(TestData.SamplePath(sample));
        var expected = new MemoryStream();
        document.ExportSylk(expected, ExcelVersion.Excel150);
        var actual = new MemoryStream();
        await document.ExportSylkAsync(actual, ExcelVersion.Excel150);
        Assert.Equal(expected.ToArray(), actual.ToArray());
    }

    [Fact]
    public async Task ExportAsync_WritesALargeSheetInPieces()
    {
        var document = ExcelDocument.CreateWorksheet();
        for (int row = 0; row < 600; row++)
        {
            for (int column = 0; column < 10; column++)
            {
                document.Cells[row, column].SetText($"Row {row} column {column}");
            }
        }

        var expected = new MemoryStream();
        document.ExportSylk(expected);
        var actual = new WriteCountingStream();
        await document.ExportSylkAsync(actual);
        Assert.Equal(expected.ToArray(), actual.ToArray());

        // The export goes out in writes of roughly 16 KB, not one write of the whole of it.
        Assert.True(actual.Writes >= expected.Length / 32768, $"{actual.Writes} writes of {expected.Length} bytes");
        Assert.True(actual.LargestWrite < 32768, $"largest write {actual.LargestWrite}");
    }

    private sealed class WriteCountingStream : MemoryStream
    {
        public int Writes { get; private set; }

        public int LargestWrite { get; private set; }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            Writes++;
            LargestWrite = Math.Max(LargestWrite, buffer.Length);
            return base.WriteAsync(buffer, cancellationToken);
        }
    }

    private static void AssertMatches(string sample, string folder, ExcelVersion version)
    {
        var document = ExcelDocument.Load(TestData.SamplePath(sample));

        // The reference style and calculation mode are global settings (spec §6), and Excel exported every sample
        // with A1 references and automatic calculation, whatever the file recorded.
        document.State.ReferenceStyle = ExcelReferenceStyle.A1;
        document.State.CalculationMode = ExcelCalculationMode.Automatic;
        var output = new ArrayBufferWriter<byte>();
        document.ExportSylk(output, version);

        byte[] expected = File.ReadAllBytes(Path.Combine(DataDirectory, folder, sample + ".slk"));

        // Through a stream, and through a buffer small enough to flush many times, the bytes are the same.
        var stream = new MemoryStream();
        document.ExportSylk(stream, version);
        Assert.Equal(output.WrittenSpan.ToArray(), stream.ToArray());
        var small = new MemoryStream();
        using (var writer = new MacExcelReader.Internal.StreamBufferWriter(small, 16))
        {
            document.ExportSylk(writer, version);
            writer.Flush();
        }

        Assert.Equal(output.WrittenSpan.ToArray(), small.ToArray());
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
        // B;Y1;X1 has already set the position (docs/Sylk.md).
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
