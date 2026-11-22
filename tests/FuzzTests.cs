using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

/// <summary>
/// Corrupts files at random and checks that the library only ever reports <see cref="ExcelFormatException"/>.
/// </summary>
public class FuzzTests
{
    // Every sheet and chart, at a lower iteration count than before because there are many more.
    public static TheoryData<string> Samples => new(AllSamples());

    [Theory]
    [MemberData(nameof(Samples))]
    public void CorruptedSamples_OnlyThrowFormatExceptions(string name)
    {
        var original = File.ReadAllBytes(SampleFile(name));
        var random = new Random(name.GetHashCode(StringComparison.Ordinal) & 0x7FFF);
        for (int iteration = 0; iteration < 150; iteration++)
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
        chart.Series.Add(ChartSeries.Create("S", [1, 2, 3], "Data", "R1C1:R1C3"));
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
        // Workspace ("Resume Excel") files start with a zero word.
        if (bytes.Length > 1 && bytes[0] == 0 && bytes[1] == 0)
        {
            try
            {
                _ = ExcelWorkspace.Read(bytes).ToArray();
            }
            catch (ExcelFormatException)
            {
            }
            catch (InvalidOperationException)
            {
            }

            return;
        }

        if (bytes.Length > 0 && bytes[0] == 0x01)
        {
            try
            {
                _ = ChartDocument.Read(bytes).ToArray();
            }
            catch (ExcelFormatException)
            {
            }
            catch (InvalidOperationException)
            {
                // A corrupted chart can read but not be writable as-is.
            }

            return;
        }

        try
        {
            var document = ExcelDocument.Read(bytes);
            foreach (var entry in document.Cells.NonEmptyCells)
            {
                var cell = entry.GetCell();
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
