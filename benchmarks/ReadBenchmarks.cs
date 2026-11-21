using BenchmarkDotNet.Attributes;

namespace MacExcelReader.Benchmarks;

/// <summary>
/// Benchmarks for reading documents: two of the Excel 1.03 samples and a large generated worksheet.
/// </summary>
[MemoryDiagnoser]
[SimpleJob]
public class ReadBenchmarks
{
    private byte[] _expenses = null!;
    private byte[] _dbForm = null!;
    private byte[] _large = null!;
    private ExcelDocument _dbFormDocument = null!;
    private string[] _dbFormTexts = null!;

    [GlobalSetup]
    public void Setup()
    {
        string samples = Path.Combine(AppContext.BaseDirectory, "Samples");
        _expenses = File.ReadAllBytes(Path.Combine(samples, "Expenses"));
        _dbForm = File.ReadAllBytes(Path.Combine(samples, "DB.FORM"));
        _large = CreateLargeWorksheet(2000, 20);
        _dbFormDocument = ExcelDocument.Read(_dbForm);
        _dbFormTexts = _dbFormDocument.Cells.NonEmptyCells
            .Where(c => c.Cell.Formula is not null)
            .Select(c => c.Cell.Formula!.ToText(_dbFormDocument))
            .ToArray();
    }

    [Benchmark]
    public ExcelFormula ParseFormula() => ExcelFormula.Parse("=IF(A1>0,SUM($A$1:$A$5)*2,-1)");

    [Benchmark]
    public int ParseDbFormFormulas()
    {
        int length = 0;
        foreach (var text in _dbFormTexts)
        {
            length += ExcelFormula.Parse(text, _dbFormDocument).Expression.Length;
        }

        return length;
    }

    /// <summary>
    /// Creates a worksheet with a mix of integers, numbers, text and formulas in every cell.
    /// </summary>
    internal static byte[] CreateLargeWorksheet(int rows, int columns)
    {
        var document = ExcelDocument.CreateWorksheet();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                var cell = document.Cells[r, c];
                switch ((r + c) % 4)
                {
                    case 0:
                        cell.SetInteger((short)(r + c));
                        break;
                    case 1:
                        cell.SetNumber(r * 1.5 + c);
                        break;
                    case 2:
                        cell.SetText($"Row {r} column {c}");
                        break;
                    default:
                        // =A1+1 (relative reference to column 0 of the same row) + SUM(...)
                        var formula = new ExcelFormula([0x44, 0xC0, 0x00, 0x00, 0x1E, 0x00, 0x01, 0x03, 0x25, 0xC0, 0x00, 0xC0, 0x05, 0x00, 0x03, 0x41, 0x04, 0x03]);
                        formula.SetResult(r);
                        cell.SetFormula(formula);
                        break;
                }
            }
        }

        return document.ToArray();
    }

    [Benchmark]
    public ExcelDocument ReadExpenses() => ExcelDocument.Read(_expenses);

    [Benchmark]
    public ExcelDocument ReadDbForm() => ExcelDocument.Read(_dbForm);

    [Benchmark]
    public ExcelDocument ReadLarge() => ExcelDocument.Read(_large);

    [Benchmark]
    public int ReadLargeAndSumValues()
    {
        var document = ExcelDocument.Read(_large);
        int count = 0;
        foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
        {
            if (cell.Value is not null)
            {
                count++;
            }
        }

        return count;
    }

    [Benchmark]
    public double ReadLargeAndSumNumbers()
    {
        var document = ExcelDocument.Read(_large);
        double total = 0;
        foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
        {
            if (cell.TryGetNumber(out double value))
            {
                total += value;
            }
        }

        return total;
    }

    [Benchmark]
    public double ReadLargeAndSumNumbersNoAllocation()
    {
        var document = ExcelDocument.Read(_large, copyData: false);
        double total = 0;
        foreach (var cell in document.Cells.EnumerateNonEmpty())
        {
            if (cell.TryGetNumber(out double value))
            {
                total += value;
            }
        }

        return total;
    }

    [Benchmark]
    public ExcelDocument ReadLargeNoCopy() => ExcelDocument.Read(_large, copyData: false);

    [Benchmark]
    public int ReadLargeCountCells() => ExcelDocument.Read(_large).Cells.CountNonEmpty();

    [Benchmark]
    public object? ReadLargeSingleCell() => ExcelDocument.Read(_large).Cells.TryGet(1500, 10)?.Value;

    [Benchmark]
    public byte[] RoundTripLarge() => ExcelDocument.Read(_large).ToArray();

    [Benchmark]
    public byte[] RoundTripExpenses() => ExcelDocument.Read(_expenses).ToArray();

    [Benchmark]
    public int DecompileDbFormFormulas()
    {
        var document = ExcelDocument.Read(_dbForm);
        int length = 0;
        foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
        {
            if (cell.Formula is { } formula)
            {
                length += formula.ToText(document).Length;
            }
        }

        return length;
    }
}
