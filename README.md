# MacExcelReader

MacExcelReader is a .NET library for reading and writing Microsoft Excel 1.x for Macintosh documents: worksheets (Finder type `XLBN`), macro sheets (`XLPG`) and charts (`MCBN`). The format is documented in [docs/Spec.md](docs/Spec.md), which was reverse-engineered from Excel 1.03.

---

## Features

- Read and write worksheets and macro sheets, byte-for-byte: every sample document from the Excel 1.03 disk round-trips unchanged.
- Cell values (integers, numbers, text), formulas with their current values, number formats, alignment, bold/italic, borders and protection flags.
- Decompile formulas and name definitions to A1-style text (`=IF(NOT(B19),RETURN())`), including macro commands and calls to user-defined macros.
- Calculation settings, column widths, standard font, document protection and the protection password hash.
- Defined names (including command-macro shortcut keys), external documents and names, windows and selections, page breaks and Page Setup.
- Create new worksheets and macro sheets with the defaults Excel 1.03 writes.
- Read and write chart files: chart type and options, axis scales, legend, text objects, and the series data (including Microsoft's 8-byte decimal number format).

---

## Usage

### Reading a worksheet

```csharp
using MacExcelReader;

var document = ExcelDocument.Load("Expenses");   // the file's data fork

foreach (var (row, column, cell) in document.Cells.NonEmptyCells)
{
    string address = FormulaDecompiler.ColumnName(column) + (row + 1);
    Console.WriteLine($"{address}: {cell.Value}  {cell.Formula?.ToText(document)}");
}

foreach (var name in document.Names)
{
    Console.WriteLine($"{name.Name} = {name.GetDefinitionText(document)}");
}
```

### Writing a worksheet

```csharp
var document = ExcelDocument.CreateWorksheet();
document.Cells[0, 0].SetText("Total");
document.Cells[1, 0].SetValue(42);
document.Cells[2, 0].SetValue(3.25);
document.Cells[2, 0].Bold = true;

// =A2*2, as parsed-expression tokens (spec §9): reference A2 (relative), integer 2, multiply.
var formula = new ExcelFormula([0x44, 0xC0, 0x01, 0x00, 0x1E, 0x00, 0x02, 0x05]);
formula.SetResult(84);
document.Cells[3, 0].SetFormula(formula);

document.Save("Worksheet");
```

The library writes only the data fork. To open the file in Excel, give it Finder type `XLBN` (or `XLPG` for a macro sheet, see `ExcelDocument.FinderType`) and creator `XCEL`.

Formulas are written as parsed expressions; the library does not include a formula parser. Excel recalculates on open, so stored results only matter for display before recalculation.

### Charts

```csharp
using MacExcelReader.Charts;

var chart = ChartDocument.Load("Chart1");
Console.WriteLine(chart.MainChart.Type);
foreach (var series in chart.Series)
{
    Console.WriteLine($"{series.Name}: {string.Join(", ", series.SecondValues)}");
}
```

Chart support is derived from Excel's chart reader and writer code; no real chart file was available to test against.

---

## Dumper

The `dumper` project is a command-line tool for inspecting files:

```sh
dotnet run --project dumper -- info tests/Samples/REPLACE.xlm
dotnet run --project dumper -- cells tests/Samples/Expenses.xls
dotnet run --project dumper -- chart Chart1
```

---

## Limitations

- New documents carry no embedded p-code segments. Those are only used when a different build of Excel links to the document while it is closed.
- Fields the spec marks as not identified are preserved as raw bytes rather than exposed.
- Array formulas, tables, array constants, user number formats, page breaks, Page Setup and charts are implemented from the spec alone; none of the sample files contain them.

---

## Building and testing

```sh
dotnet build
dotnet test
```

The tests use the five documents from the Excel 1.03 disk in `tests/Samples`.

---

## License

MIT, see [LICENSE](LICENSE).
