# MacExcelReader

MacExcelReader is a fast, allocation-conscious .NET library for reading and writing Microsoft Excel 1.x for Macintosh documents (Excel 1.00 to 1.5): worksheets (Finder type `XLBN`), macro sheets (`XLPG`) and charts (`MCBN`).

The file format was reverse-engineered from the code of Excel 1.00, 1.01, 1.03, 1.04 and 1.5 and is documented in [docs/Spec.md](docs/Spec.md). The full API is described in the [usage guide](docs/Usage.md).

---

## Features

- **Exact round-trips.** All 37 sample documents from the Excel 1.00, 1.03 and 1.5 disks, and 33 files Excel itself saved in experiments, read and write back byte for byte, sheets and charts, including Excel 1.5's version-18 files. Fields the library does not model are preserved untouched.
- **Cells.** Integers, numbers, text and formulas with their current values, plus number formats, alignment, bold and italic, borders, and protection flags.
- **Formulas, both ways.** Decompile tokens to A1-style text (`=IF(NOT(FORMULA.FIND(B19,1,2,1)),RETURN())`) and compile text back to tokens with `ExcelFormula.Parse`. The parser reproduces Excel's own encoding byte for byte for 1,281 of the samples' 1,282 formulas and all 206 name definitions, for whichever release saved them: token classes, `IF` and `CHOOSE` jumps, the one-argument `SUM` attribute, volatile and memory-function tokens, array formulas and array constants, macro commands, dialog-box forms (`RUN?()`), names, external references and calls to macros on this or another sheet. Every function Excel 1.00–1.5 knows, with each release's argument rules. Allocation-free token walking.
- **Display formatting.** All 19 built-in number formats, with dates in the 1904 date system. User-defined formats are read and written.
- **Everything else in the file.** Calculation settings, column widths, fonts, protection and the password hash, defined names and command-macro shortcuts, external documents, windows and selections, page breaks and Page Setup.
- **Writing.** Create new worksheets and macro sheets with Excel 1.03's defaults, or edit existing ones; only the cells you touch are re-encoded.
- **Charts.** Chart type and options, axes and scales, legend, titles, data labels, free text and arrows, point formats (markers, patterns, exploded pie slices, label settings), overlay charts, the layout Excel last drew, and the series data. Numbers use Microsoft's 8-byte decimal format.
- **Robust.** Malformed input only ever raises `ExcelFormatException`. This is checked by fuzzing the samples with hundreds of thousands of random corruptions.

---

## Quick start

```sh
dotnet add reference path/to/MacExcelReader/src/MacExcelReader.csproj
```

### Read

```csharp
using MacExcelReader;

var document = ExcelDocument.Load("Expenses");   // the file's data fork

foreach (var (row, column, cell) in document.Cells.NonEmptyCells)
{
    var address = new ExcelCellAddress(row, column);
    Console.WriteLine($"{address}: {ExcelValueFormatter.FormatCell(cell)}  {cell.Formula?.ToText(document)}");
}

Console.WriteLine(document.Cells["C11"].Value);   // 1000
```

### Scan a large sheet without allocating

```csharp
double total = 0;
foreach (var cell in document.Cells.EnumerateNonEmpty())
{
    if (cell.TryGetNumber(out double value))
    {
        total += value;
    }
}
```

### Write

```csharp
var sheet = ExcelDocument.CreateWorksheet();
sheet.Cells["A1"].SetText("Total");
sheet.Cells["A2"].SetValue(42);
sheet.Cells["A2"].Bold = true;

var formula = ExcelFormula.Parse("=A2*2");
formula.SetResult(84);   // the value Excel shows until it recalculates
sheet.Cells["A3"].SetFormula(formula);

sheet.Save("Totals");   // set Finder type sheet.FinderType ("XLBN") and creator "XCEL"
```

### Charts

```csharp
using MacExcelReader.Charts;

var chart = ChartDocument.Load("Chart1");
Console.WriteLine($"{chart.MainChart.Type}, legend {chart.MainChart.LegendPosition}");
foreach (var series in chart.Series)
{
    Console.WriteLine($"{series.Name}: {string.Join(", ", series.SecondValues)}");
}
```

---

## Performance

Reading is lazy:
- `Read` builds no cell objects; the grid is attached to the file's cell table.
- A cell is created when first accessed, and its text or formula is decoded when first read.
- Saving copies untouched cells byte for byte.

Measured with BenchmarkDotNet (`benchmarks/`) on an Apple Silicon Mac with .NET 9. "Large" is a generated 2,000 × 20 sheet: 40,000 cells mixing integers, numbers, text and formulas.

| Operation | Time | Allocated |
|-----------|-----:|----------:|
| Read Expenses (15 KB, 988 cells) | 3.1 µs | 23 KB |
| Read DB.FORM (24 KB, 433 cells, 275 formulas) | 3.9 µs | 35 KB |
| Read large sheet | 77 µs | 841 KB |
| Read large sheet, `copyData: false` | 10 µs | 26 KB |
| Read large sheet and sum every number with `EnumerateNonEmpty` | 145 µs | 26 KB |
| Read large sheet and sum every number via `NonEmptyCells` | 759 µs | 2.7 MB |
| Read large sheet and box every value via `Value` | 2.8 ms | 4.9 MB |
| Read and re-save Expenses unchanged | 11 µs | 54 KB |
| Read and re-save the large sheet unchanged | 705 µs | 2.4 MB |
| Read DB.FORM and decompile all 275 formulas | 46 µs | 206 KB |
| Parse `=IF(A1>0,SUM($A$1:$A$5)*2,-1)` | 0.71 µs | 2.3 KB |
| Parse all 275 DB.FORM formulas from text | 118 µs | 282 KB |

The 841 KB when reading the large sheet is the copy of the input; pass `copyData: false` to avoid it.

Before the lazy design, reading Expenses took 36 µs (192 KB) and the large sheet 3.4 ms (7.6 MB).

Run them yourself:

```sh
dotnet run -c Release --project benchmarks -- --filter '*'
```

---

## Dumper

`dumper` is a command-line tool for inspecting files. It accepts worksheets, macro sheets and charts.

| Command | What it shows |
|---------|---------------|
| `info <file>` | A summary: kind, size, cell counts, names, protection |
| `cells <file> [-r A1:D20] [--all]` | Every non-empty cell with its displayed value, format, style and formula |
| `cell <file> <address>` | Everything about one cell: value, date, formats, flags, formula, result, raw bytes and a token-by-token breakdown |
| `parse <formula> [-f file] [--name] [--array] [--excel 1.5]` | Compile formula text and show the tokens Excel would store |
| `sheet <list> <output> [--split N] [--excel 1.5]` | Build sheets from a formula list (`tests/Data`), for testing formulas in Excel |
| `excel-test pack <disk> <files...>` / `excel-test check <disk>` | Test files in a real Excel under an emulator: see [Testing in a real Excel](docs/Testing.md) |
| `names <file>` | Defined names, command shortcuts and definitions |
| `export <file> [-o out.csv] [--tsv] [--formulas] [--raw] [-r range]` | The sheet as CSV or TSV: displayed values, raw values or formulas |
| `json <file> [-o out.json] [--empty]` | The whole document as JSON |
| `dump <file>` | Everything: header, settings, column widths, external references, names, windows, page breaks, print setup and cells |
| `workspace <file>` | A workspace ("Resume Excel", type XLIN) file: settings and the windows it reopens |
| `layout <file>` | Where each part of the file lies: header, code segments, state block and all eight sections |
| `chart <file>` | A chart file: charts, axes, scales, text objects and series data |
| `newchart <type> <output>` | Build a chart with the library, for testing in Excel |

```sh
dotnet run --project dumper -- cell "tests/Samples/Excel 1.00/REPLACE" B27
dotnet run --project dumper -- export "tests/Samples/Excel 1.00/Expenses" -r A10:D20
dotnet run --project dumper -- parse '=IF(A1>0,SUM(A1:A3),-1)'
```

---

## Limitations

- New documents carry empty embedded p-code segments. Excel 1.03 never uses a file's own copy, but a different build of Excel linking to the document while it is closed would; copy the segments from an existing document if that matters.
- The formula parser does not calculate. A new formula's cached value is whatever you set with `SetResult`, and Excel keeps it until something the formula depends on changes. Set `document.State.RecalculationState = 1` to have Excel recalculate the sheet when it opens it.
- The parser's rules come from Excel's code and tables and were checked against every sample formula; formulas unlike any in the samples could still encode differently. Excel 1.5's functions numbered above 255 follow its Save code, as no sample uses one; Excel 1.5 accepts them.
- Files written by the library have been checked in Excel 1.00, 1.03 and 1.5 under Mini vMac: every formula in `tests/Data` in 1.03, part of the 1.5 list including functions above 255, and the experiments listed in [Testing in a real Excel](docs/Testing.md), including charts the library builds from scratch, which Excel 1.00 draws and re-saves with no change to their content.

---

## Building and testing

```sh
dotnet build
dotnet test
```

The tests run over the sample documents in `tests/Samples`, grouped by the release that saved them: `Excel 1.00` (the 1.00 data disk, whose Expenses, REPLACE, COMMANDS and DB.FORM are identical to the 1.03 disk's), `Excel 1.01` (files the French 1.01 saved in experiments), `Excel 1.03` (Worksheet1) and `Excel 1.5` (the 1.5 sampler, including two charts). Each folder also holds the files that release saved in experiments, named for what they show; `tests/ExperimentTests.cs` checks what each one showed.

---

## License

MIT, see [LICENSE](LICENSE).
