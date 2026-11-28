# MacExcelReader

[![NuGet](https://img.shields.io/nuget/v/MacExcelReader.svg)](https://www.nuget.org/packages/MacExcelReader)
[![CI](https://github.com/hughbe/MacExcelReader/actions/workflows/ci.yml/badge.svg)](https://github.com/hughbe/MacExcelReader/actions/workflows/ci.yml)

A fast .NET library for reading and writing Microsoft Excel 1.x for Macintosh documents, from Excel 1.00 (1985) to
Excel 1.5 (1988): worksheets (Finder type `XLBN`), macro sheets (`XLPG`) and charts (`MCBN`).

The file format was reverse-engineered from the code of Excel 1.00, 1.01, 1.03, 1.04, 1.06 and 1.5, and checked
by running files in the real thing under an emulator. It is documented in full in the
[format specification](https://github.com/hughbe/MacExcelReader/blob/main/docs/Spec.md).

---

## Install

```sh
dotnet add package MacExcelReader
```

The library targets .NET 9 and has no dependencies.

---

## Features

- **Exact round-trips.** Every sample, the 37 documents from Microsoft's Excel 1.00, 1.03 and 1.5 disks and 147
  files Excel saved in experiments, reads and writes back byte for byte: sheets and charts, version-17 and
  Excel 1.5's version-18 files. Fields the library does not model are kept as they are.
- **Cells.** Integers, numbers, text and formulas with their current values, number formats, alignment, bold
  and italic, borders and protection.
- **Formulas, both ways.** Decompile tokens to A1-style text and compile text back to tokens with
  `ExcelFormula.Parse`, as each release encodes them: token classes, `IF` and `CHOOSE` jumps, array formulas and
  constants, macro commands and dialog-box forms (`RUN?()`), names, external references, and every function
  Excel 1.00–1.5 knows with each release's argument rules. All 1,274 test formulas compile to the bytes Excel
  1.00, 1.03 or 1.5 gave them when re-entering them, except a whole-sheet range that Excel re-entered from
  shortened text, and so do all 206 name definitions in Microsoft's samples.
- **Number formats.** Decompile any format to its text, format values with it as Excel does (dates in the 1904
  date system, General, `###` in a narrow column), and compile text into a new format with
  `ExcelDocument.AddNumberFormat`, byte for byte as Excel 1.00 or 1.5 does.
- **Everything else in the file.** Calculation settings, column widths, fonts, protection and its password
  hash, defined names and macro shortcuts, external documents, windows and selections (frozen panes and
  gridline colours in Excel 1.5), page breaks and Page Setup.
- **Writing.** Create worksheets and macro sheets with Excel's defaults, or edit existing ones; only the cells
  you change are re-encoded. Values are checked as they are set against what a file can hold and what Excel
  itself would store, and defined names are kept in the order Excel searches them.
- **Charts.** Type and options, axes and scales, legend, titles, data labels, free text and arrows, point
  formats, overlay charts and the series data, in Microsoft Chart's 8-byte decimal numbers. Automatic axis
  scales are worked out as Excel 1.00 draws them, to the bit.
- **Drawing.** `ExcelRenderer` draws a sheet's window as Excel 1.00 shows it, as a one-bit picture or PNG, with
  the bitmap fonts of a System file you supply: headings, gridlines, cells laid out as Excel lays them out,
  borders, the selection and the scroll bars, and sizes the System has no bitmap for stretched as the Mac Plus
  ROM stretches them. It matches all 27 windows captured from Excel 1.00's screen pixel for pixel, and draws a
  window in about 0.2 ms whatever the sheet's size. `ChartRenderer` draws a chart's window the same way: axes,
  ticks, gridlines, labels, titles, markers, columns, bars, areas and pies, with QuickDraw's lines, polygons and
  arcs. It matches ten of twelve chart windows captured from Excel 1.00 pixel for pixel; legends are not drawn
  yet.
- **Workspaces.** Read and write "Resume Excel" files: the settings and the windows Excel reopens.
- **SYLK export.** `ExportSylk` writes a sheet as SYLK, byte for byte as Excel 1.00's or 1.5's Save As does,
  checked against 178 files the two exported. It allocates nothing per cell.
- **Robust.** Malformed input only ever raises `ExcelFormatException`, checked by fuzzing the samples with
  hundreds of thousands of random corruptions.

---

## Quick start

### Read a sheet

```csharp
using MacExcelReader;

var document = ExcelDocument.Load("Expenses");   // the file's data fork

// Cells are read straight from the file: walking them allocates nothing.
foreach (var cell in document.Cells.NonEmptyCells)
{
    Console.WriteLine($"{cell.Address}: {ExcelValueFormatter.FormatCell(cell, document)}");
}

ExcelValue value = document.Cells["C11"].Value;   // a number, text, TRUE/FALSE, an error, or empty
Console.WriteLine(value);                         // 1000
Console.WriteLine(document.Cells["B19"].Formula?.ToText(document));
```

### Sum a large sheet without allocating

```csharp
double total = 0;
foreach (var cell in document.Cells.NonEmptyCells)
{
    if (cell.TryGetNumber(out double value))
    {
        total += value;
    }
}
```

### Write a sheet

```csharp
var sheet = ExcelDocument.CreateWorksheet();
sheet.Cells["A1"].SetText("Total");
sheet.Cells["A2"].SetValue(1234.5);
sheet.Cells["A2"].NumberFormat = sheet.AddNumberFormat("#,##0.00 \"DM\"");
sheet.Cells["A2"].Bold = true;

var formula = ExcelFormula.Parse("=A2*2");
formula.SetResult(2469);   // the value Excel shows until it recalculates
sheet.Cells["A3"].SetFormula(formula);

sheet.Save("Totals");   // give the file Finder type sheet.FinderType ("XLBN") and creator "XCEL"
```

### Read a chart

```csharp
using MacExcelReader.Charts;

var chart = ChartDocument.Load("Chart1");
Console.WriteLine($"{chart.MainChart.Type}, legend {chart.MainChart.LegendPosition}");
foreach (var series in chart.Series)
{
    Console.WriteLine($"{series.Name}: {string.Join(", ", series.Values)}");
}
```

The [usage guide](https://github.com/hughbe/MacExcelReader/blob/main/docs/Usage.md) covers the whole API.

---

## Documentation

| Document | What it covers |
|----------|----------------|
| [Usage guide](https://github.com/hughbe/MacExcelReader/blob/main/docs/Usage.md) | The library's API, with examples |
| [Format specification](https://github.com/hughbe/MacExcelReader/blob/main/docs/Spec.md) | The file format, byte by byte, with the code of each Excel release it comes from |
| [How Excel 1.x displays a sheet](https://github.com/hughbe/MacExcelReader/blob/main/docs/Display.md) | How Excel formats values, fits them into cells and draws the window, pixel by pixel |
| [How Excel 1.x exports SYLK](https://github.com/hughbe/MacExcelReader/blob/main/docs/Sylk.md) | What Excel's Save As writes for the SYLK type, byte for byte |
| [Testing in a real Excel](https://github.com/hughbe/MacExcelReader/blob/main/docs/Testing.md) | The harness that runs files in Excel, headless under MAME or by hand in Mini vMac, and every experiment's results |
| [Excel bugs and quirks](https://github.com/hughbe/MacExcelReader/blob/main/docs/ExcelBugs.md) | Excel's own bugs and oddities found along the way |
| [Excel 1.x and BIFF2](https://github.com/hughbe/MacExcelReader/blob/main/docs/Biff2.md) | How the format differs from BIFF2, Excel 2.x's format |
| [How Excel 1.x calculates](https://github.com/hughbe/MacExcelReader/blob/main/docs/Calculation.md) | What Excel 1.00 calculates, measured: arithmetic, functions, dates, text and errors |

---

## Performance

Reading is lazy:
- `Read` builds no cell objects; the grid is attached to the file's cell table.
- A cell is created when first accessed, and its text or formula is decoded when first read.
- Saving copies untouched cells byte for byte.

Measured with BenchmarkDotNet (`benchmarks/`) on an Apple Silicon Mac with .NET 9. "Large" is a generated 2,000 × 20 sheet: 40,000 cells mixing integers, numbers, text and formulas.

| Operation | Time | Allocated |
|-----------|-----:|----------:|
| Read Expenses (15 KB, 988 cells) | 0.82 µs | 17 KB |
| Read DB.FORM (24 KB, 433 cells, 275 formulas) | 2.4 µs | 30 KB |
| Read large sheet | 42 µs | 849 KB |
| Read large sheet, `copyData: false` | 0.36 µs | 1.8 KB |
| Read large sheet, `copyData: false`, and sum every number via `NonEmptyCells` | 119 µs | 1.8 KB |
| Read large sheet and create every cell object (`GetCell`) | 1.4 ms | 2.7 MB |
| Read and re-save Expenses unchanged | 11.3 µs | 32 KB |
| Read Expenses with `copyData: false` and write it to a stream | 10.7 µs | 1.8 KB |
| Read and re-save the large sheet unchanged | 626 µs | 1.7 MB |
| Read DB.FORM and decompile all 275 formulas | 51 µs | 95 KB |
| Format every DB.FORM cell for display, as strings | 11.6 µs | 20 KB |
| Format every DB.FORM cell for display, into a span (`TryFormatCell`) | 14.8 µs | 0 B |
| Export DB.FORM as SYLK to a stream | 68 µs | 120 B |
| Parse `=IF(A1>0,SUM($A$1:$A$5)*2,-1)` | 0.46 µs | 144 B |
| Parse all 275 DB.FORM formulas from text | 58 µs | 31 KB |

The 849 KB when reading the large sheet is the copy of the input; pass `copyData: false` to avoid it. Parsing allocates only the formula it returns.

Displaying values allocates only the strings it returns, and nothing through the span forms
(`ExcelValueFormatter.TryFormatCell`, `ExcelNumberFormat.TryFormat`, `ExcelValue.TryFormat`,
`ExcelFormula.TryFormatText`): formatting every cell of DB.FORM that way allocates 0 bytes.

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
| `sheet <list> <output> [--split N] [--excel 1.5] [--type]` | Build sheets from a formula list (`tests/Data`), for testing formulas in Excel; `--type` has Excel 1.00 type each formula in A1 text instead |
| `excel-test experiment <files...> -o <dir>` / `pack` / `run` / `check` / `files` | Test files in a real Excel under an emulator (`experiment` and `run`: Excel 1.00 to 1.5 in MAME, headless, with `--excel`; `experiment` packs, runs and checks any number of files and reports each command's result; a `FILE.type` has `run` type formulas into a file as a person would; `files` lists or copies out a floppy image's files): see [Testing in a real Excel](https://github.com/hughbe/MacExcelReader/blob/main/docs/Testing.md) |
| `names <file>` | Defined names, command shortcuts and definitions |
| `export <file> [-o out.csv] [--tsv] [--formulas] [--raw] [-r range]` | The sheet as CSV or TSV: displayed values, raw values or formulas |
| `json <file> [-o out.json] [--empty]` | The whole document as JSON |
| `sylk <file> [-o out.slk] [--excel 1.5]` | The sheet as SYLK, as Excel 1.00 (or 1.5) exports it |
| `render <file> --system <disk or rsrc> [-o out.png] [--no-frame] [--scale N]` | The sheet's or chart's window drawn as Excel 1.00 shows it, as a PNG, with the fonts of a System file (Excel 1.00's Data disk image works) |
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

- **No calculation.** A new formula's cached value is whatever you set with `SetResult`, and Excel keeps it
  until something the formula depends on changes. Set
  `document.State.RecalculationState = ExcelRecalculationState.Needed` to have Excel recalculate the sheet when
  it opens it.
- **Formulas from pre-release builds.** Of the 1,282 formulas in Microsoft's samples, 1,243 recompile to the
  stored bytes. The other 39 were compiled by builds before the releases (the 1.00 disk's macro sheets and two
  of the 1.5 samples) and encode a few things differently; they still decompile correctly
  ([spec §2.3 and §15.4](https://github.com/hughbe/MacExcelReader/blob/main/docs/Spec.md)).
- **Display.** `ExcelValueFormatter.FormatCell(cell, document)` gives the text Excel turns a value into.
  `FormatCell(cell, document, column)`, `LayoutCell` and `GetRowMetrics` follow how Excel 1.00 and 1.5 draw it
  (pass `ExcelVersion.Excel103` for 1.03's one difference), in
  pixels of the sheet's font at any size, for the font families in Excel 1.00's System (every sample uses Geneva
  10). For other families they count every character as one digit width, so a value at the edge of its column
  may differ. Colours are not applied, and format text uses the US letters and separators.
- **Not thread-safe.** Reading a value can decode and cache it, and the grid creates cells on first access, so
  share a document between threads only behind a lock, even for reading.
- **Empty embedded code.** New documents carry empty copies of the code segments every Excel document embeds.
  No released Excel reads them ([spec §4.1](https://github.com/hughbe/MacExcelReader/blob/main/docs/Spec.md)), so this only matters to a build with a
  different code signature reading the file through a link.

Files the library writes have been opened and re-saved in Excel 1.00, 1.03, 1.04, 1.06 and 1.5 in the
[experiments](https://github.com/hughbe/MacExcelReader/blob/main/docs/Testing.md), including charts built from scratch.

---

## Building and testing

```sh
dotnet build
dotnet test
```

The tests run over the samples in `tests/Samples`, in a folder for each release that saved them, and over the
formula and number-format lists in `tests/Data`. Files Excel saved in experiments are named for what they show,
and `tests/ExperimentTests.cs` checks each one.

### Releasing

Push a version tag, such as `v1.0.0`. The Release workflow builds and tests that commit, packs it with the tag
as the version and publishes it to nuget.org with NuGet Trusted Publishing.

---

## License

MIT; see [LICENSE](https://github.com/hughbe/MacExcelReader/blob/main/LICENSE).
