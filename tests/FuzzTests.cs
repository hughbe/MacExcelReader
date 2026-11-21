using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// Corrupts files at random and checks that the library only ever reports <see cref="ExcelFormatException"/>.
/// </summary>
public class FuzzTests
{
    public static TheoryData<string> Samples => ["Worksheet1.xls", "Expenses.xls", "COMMANDS.xlm", "DB.FORM.xlm", "REPLACE.xlm"];

    [Theory]
    [MemberData(nameof(Samples))]
    public void CorruptedSamples_OnlyThrowFormatExceptions(string name)
    {
        var original = Sample(name);
        var random = new Random(name.GetHashCode(StringComparison.Ordinal) & 0x7FFF);
        for (int iteration = 0; iteration < 400; iteration++)
        {
            var bytes = (byte[])original.Clone();
            int changes = 1 + random.Next(4);
            for (int i = 0; i < changes; i++)
            {
                bytes[random.Next(bytes.Length)] = (byte)random.Next(256);
            }

            if (random.Next(10) == 0)
            {
                Array.Resize(ref bytes, random.Next(bytes.Length));
            }

            Exercise(bytes);
        }
    }

    [Fact]
    public void CorruptedCharts_OnlyThrowFormatExceptions()
    {
        var chart = ChartDocument.Create(ChartType.ColumnOrBar);
        chart.MainChart.Texts.Add(ChartText.Create(ChartTextKind.ChartTitle, "Title"));
        chart.Series.Add(ChartSeries.Create([1, 2, 3], "S", "=R1C1:R1C3"));
        var original = chart.ToArray();
        var random = new Random(1234);
        for (int iteration = 0; iteration < 400; iteration++)
        {
            var bytes = (byte[])original.Clone();
            for (int i = 0; i < 1 + random.Next(3); i++)
            {
                bytes[random.Next(bytes.Length)] = (byte)random.Next(256);
            }

            try
            {
                var document = ChartDocument.Read(bytes);
                _ = document.ToArray();
            }
            catch (ExcelFormatException)
            {
            }
            catch (InvalidOperationException)
            {
                // A document that reads but cannot be written as-is (for example three charts).
            }
        }
    }

    private static void Exercise(byte[] bytes)
    {
        try
        {
            var document = ExcelDocument.Read(bytes);
            foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
            {
                _ = cell.Value;
                if (cell.Formula is { } formula)
                {
                    _ = formula.ToText(document);
                }
            }

            foreach (var name in document.Names)
            {
                _ = name.GetDefinitionText(document);
            }

            _ = document.ToArray();
        }
        catch (ExcelFormatException)
        {
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("too long") || ex.Message.Contains("1 MB"))
        {
            // A corrupted document can read but exceed a limit when written.
        }
    }
}
