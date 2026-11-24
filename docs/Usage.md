# MacExcelReader usage guide

This guide walks through the library's API. For the file format itself, see [Spec.md](Spec.md).

- [Files and Finder types](#files-and-finder-types)
- [Opening a document](#opening-a-document)
- [Reading cells](#reading-cells)
- [Displaying values](#displaying-values)
- [Formulas](#formulas)
- [Names and external references](#names-and-external-references)
- [Sheet settings, protection and printing](#sheet-settings-protection-and-printing)
- [Creating and editing documents](#creating-and-editing-documents)
- [Charts](#charts)
- [Performance](#performance)
- [Errors](#errors)
- [Inspecting files](#inspecting-files)

---

## Files and Finder types

Excel 1.x documents are classic Macintosh files. Everything the library needs is in the **data fork**; the
resource fork is empty. The Finder type says what the file is:

| Finder type | Creator | Contents | Class |
|-------------|---------|----------|-------|
| `XLBN` | `XCEL` | Worksheet | `ExcelDocument` |
| `XLPG` | `XCEL` | Macro sheet | `ExcelDocument` |
| `MCBN` | | Chart | `Charts.ChartDocument` |

If your files are inside disk images, extract the data forks first (for example with
[HfsReader](https://github.com/hughbe/HfsReader)). When writing, give the file the Finder type from
`ExcelDocument.FinderType` and creator `XCEL` so the Finder opens it in Excel.

You can tell the two formats apart from the first two bytes: `53 27` for worksheets and macro sheets,
`01 8C` for charts.

## Opening a document

```csharp
using MacExcelReader;

var document = ExcelDocument.Load("Expenses");          // from a path
var fromBytes = ExcelDocument.Read(bytes);              // copies the array
var fromStream = ExcelDocument.Read(stream);
var noCopy = ExcelDocument.Read(bytes, copyData: false); // you hand the array over
var later = await ExcelDocument.LoadAsync("Expenses");   // ReadAsync, SaveAsync and WriteAsync too
```

`Read(byte[])` copies the bytes so you can reuse the array. Pass `copyData: false` to skip the copy; the
document reads cells from the array lazily, so you must not change it afterwards.

`document.Kind` is `Worksheet` or `MacroSheet`. `ChartDocument` and `ExcelWorkspace` have the same `Load`, `Read`
(bytes, span or stream), `ReadAsync`, `Save`, `Write` and `WriteAsync` methods.

## Reading cells

The grid is dense: every position inside `Cells.RowCount` × `Cells.ColumnCount` is a cell. Rows and columns
are zero-based; addresses use A1 notation.

```csharp
var cell = document.Cells["B19"];          // same as document.Cells[18, 1]
var maybe = document.Cells.Find("Z99");  // null outside the grid; never grows it

switch (cell.Type)
{
    case ExcelCellType.Integer: ushort i = cell.IntegerValue; break;
    case ExcelCellType.Number:  double d = cell.NumberValue; break;
    case ExcelCellType.Text:    string s = cell.TextValue; break;
    case ExcelCellType.BooleanOrError:
        if (cell.IsError) { ExcelError e = cell.Error; } else { bool b = cell.BooleanValue; }
        break;
    case ExcelCellType.NumericFormula:
    case ExcelCellType.Formula: var f = cell.Formula!; break;
}

ExcelValue value = cell.Value;   // a number, text, TRUE or FALSE, an error, or empty
switch (value.Kind)
{
    case ExcelValueKind.Number:  double n = value.Number; break;
    case ExcelValueKind.Text:    string t = value.Text; break;
    case ExcelValueKind.Boolean: bool b = value.Boolean; break;
    case ExcelValueKind.Error:   ExcelError e = value.Error; break;
}

Console.WriteLine(value);   // as Excel shows it with General: 12.5, TRUE, #N/A
```

`ExcelValue` is a struct, so reading a value does not allocate, apart from decoding text. Numbers, text and TRUE or
FALSE convert to it implicitly, and errors through `ExcelValue.FromError`:

```csharp
cell.SetValue(12.5);
cell.SetValue("Total");
cell.SetValue(ExcelValue.FromError(ExcelError.NotAvailable));
```

`Integer` cells hold whole numbers from 0 to 65535; Excel stores a typed whole number in that range this way to
save space, and a negative one as a number. `BooleanOrError` cells hold a TRUE, FALSE or error typed into the
cell. Use `TryGetNumber` to treat
integers, numbers and numeric formulas alike, without allocating:

```csharp
if (cell.TryGetNumber(out double number)) { ... }
```

Each cell also has its formatting: `NumberFormat` (0–63; 0–18 are built in), `Alignment`, `Bold`, `Italic`,
`LeftBorder`/`RightBorder`/`TopBorder`/`BottomBorder`, and protection flags `IsLocked` and `IsHidden`.

### Iterating

`NonEmptyCells` gives each cell with a value as an `ExcelCellRef`: a lightweight view read straight from the
file, with the cell's address, type, value and formatting. A `foreach` over it allocates nothing; LINQ works too.

```csharp
foreach (var c in document.Cells.NonEmptyCells)
{
    if (c.TryGetNumber(out double v)) total += v;
    ExcelValue value = c.Value;
    ExcelCell full = c.GetCell();   // to change the cell, or read its formula
}

int count = document.Cells.CountNonEmpty();
ExcelCellType type = document.Cells.GetCellType(5, 2);   // without creating the cell

Span<char> buffer = stackalloc char[255];                // text is at most 255 characters
foreach (var c in document.Cells.NonEmptyCells)
{
    if (c.TryGetText(buffer, out int length)) Use(buffer[..length]);   // text without a string
}
```

The indexer `document.Cells[row, column]` takes rows 0–16383 and columns 0–255 and creates the cell; `Find`
returns null outside the grid instead.

## Displaying values

`ExcelValueFormatter` formats values the way Excel shows them, through the tokens of their number format. Pass the
document so it can use the sheet's own formats (index 19 and up):

```csharp
string shown = ExcelValueFormatter.FormatCell(cell, document); // "1/1/84", "$1,000 ", "22-17-023"…
string drawn = ExcelValueFormatter.FormatCell(cell, document, column: 2); // fitted to column C: "1.2E+08", "####"…
ExcelNumberFormat format = document.GetNumberFormat(cell.NumberFormat)!;
string text = format.ToText();                                 // "00-00-000", as the Format Number dialog shows it
string fmt = ExcelNumberFormat.BuiltIn[10].ToText();         // "m/d/yy"
DateTime date = ExcelValueFormatter.ToDateTime(29220);        // 1984-01-01
double serial = ExcelValueFormatter.ToSerial(DateTime.Today);
```

Each has a `Try` form that writes into a span instead of returning a string, and allocates nothing:
`ExcelValueFormatter.TryFormatCell` and `TryFormat`, `ExcelNumberFormat.TryFormat`, `ExcelValue.TryFormat`,
`ExcelFormula.TryFormatText` and `ExcelDefinedName.TryFormatDefinitionText`. `ExcelValue`, `ExcelCellAddress` and
`ExcelRange` are `ISpanFormattable`, so interpolating them into a span allocates nothing either.

```csharp
Span<char> buffer = stackalloc char[512];
foreach (var c in document.Cells.NonEmptyCells)
{
    if (ExcelValueFormatter.TryFormatCell(c, document, buffer, out int length)) Write(buffer[..length]);
}
```

To draw a sheet as Excel 1.00 did, `LayoutCell` gives each cell's text, where it starts and the stretch of the
row it may cover, in pixels from the left edge of column A (spec Appendix D):

```csharp
if (ExcelValueFormatter.LayoutCell(document, row: 0, column: 1) is { } display)
{
    // For "General text spills right" in B1 of a sheet with 75-pixel columns: X 77, Width 121,
    // ClipLeft 75, ClipRight 224 — left aligned, 2 pixels in, running on over a blank C1.
}
```

Excel 1.x for Macintosh uses the **1904 date system**: serial 0 is 1 January 1904.

## Formulas

```csharp
var formula = cell.Formula!;
string text = formula.ToText(document);     // "=IF(NOT(B19),RETURN())"
ExcelValue result = formula.Result;         // the value Excel last calculated
```

Passing the document lets the decompiler resolve defined names and external documents. Formulas are stored
as **parsed expressions** (spec §15): reverse-Polish tokens closely related to BIFF2. You can walk them
without allocating:

```csharp
foreach (var token in FormulaTokens.Enumerate(formula.Expression))
{
    Console.WriteLine($"{token.Offset}: {FormulaTokens.GetName(token.Code)} ({token.Length} bytes)");
}
```

Array formulas decompile from the cell that holds the shared formula, as `{=TREND(B1:B5)}`; `IsArrayFormula`,
`IsSharedAnchor`, `SharedAnchor` and `SharedRange` describe the range.

Functions are identified by number (spec Appendix A). `ExcelFunctions.Get(n)` gives the name, argument
counts and types, flags and the release that introduced it; `ExcelFunctions.Get(n, ExcelVersion.Excel150)`
gives a particular release's rules, and `ExcelFunctions.TryFindNumber("SUM", out int number)` goes the other way.

### Writing a formula

`ExcelFormula.Parse` compiles formula text into the tokens Excel would store:

```csharp
var formula = ExcelFormula.Parse("=IF(A1>0,SUM($A$1:$A$5)*2,-1)");
formula.SetResult(0);            // the value Excel shows until it recalculates
document.Cells["B1"].SetFormula(formula);
document.State.RecalculationState = ExcelRecalculationState.Needed;   // have Excel recalculate the sheet when it opens it
```

Excel does not recalculate the formulas it loads until something they depend on changes, so set the
recalculation state when you cannot work out the results yourself.

The leading `=` is optional. It supports:

- numbers, strings (`"say ""hi"""`), `TRUE`/`FALSE` and error values (`#N/A`, `#DIV/0!`, …);
- A1 references and areas with `$` for absolute parts, up to `IV16384`;
- the operators `: , % ^ * / + - & = <> < > <= >=`, with Excel's precedence, and parentheses;
- every Excel 1.x function and macro command, with argument counts checked against Excel's function
  table, and dialog-box forms of commands such as `RUN?()`;
- array constants (`{1,2;3,"x"}`);
- defined names and calls to user-defined function macros (`=MYMACRO(1,2)`, `='Other Sheet'!MACRO(1)`);
- external references (`Budget!A1`, `'My Budget'!A1`, `Budget!Total`) and the active sheet (`!A1`), which
  need a matching entry in `document.ExternalDocuments`, with the names used from it in that entry's
  `Names`.

**Excel versions.** The encoding differs slightly between releases, and Excel 1.5 adds functions. Pass the
release the file is for; the default, `ExcelVersion.Excel100`, produces formulas every Excel 1.x reads:

```csharp
var formula = ExcelFormula.Parse("=LEFT(A1,3)", document, ExcelVersion.Excel150);
```

A formula that uses a function added after 1.03 raises the document's `FunctionCount`, which tells older
releases they cannot read it.

**Array formulas** cover a range of cells:

```csharp
document.SetArrayFormula(ExcelRange.Parse("C1:C5"), "{=B1:B5*2}");
```

Pass the document whenever the formula uses names or external references; otherwise `null` is fine.
Invalid text raises `ExcelFormulaSyntaxException`, whose `Position` is the zero-based offset of the problem.

Name definitions store their indices inside the tokens, so they have their own entry point. They can also be
unions without parentheses (`=$A$1,$C$3`) and array constants:

```csharp
var name = new ExcelDefinedName("Totals");
name.SetDefinition("=$A$1:$A$5", document);
```

A relative reference in a name is stored as an offset from the cell that uses the name, so it means something
only against a cell, as in Excel's Define Name dialog, which uses the active cell. Give that cell when setting or
reading such a definition; without one, A1 is used:

```csharp
var above = new ExcelDefinedName("Above");
above.SetDefinition("=B1", new ExcelCellAddress(1, 1), document);   // B2 active: the cell above
Console.WriteLine(above.GetDefinitionText(document, new ExcelCellAddress(9, 3)));   // D9
```

The rules the parser follows (token classes, `IF` and `CHOOSE` jumps, the `SUM` attribute, volatile and
memory-function tokens, and how they differ between releases) are in
[spec §15.4](Spec.md#154-how-excel-compiles-formula-text). To see what it produces:

```sh
dotnet run --project dumper -- parse '=COLUMNS(!DATABASE)' --file "tests/Samples/Excel 1.00/DB.FORM"
dotnet run --project dumper -- parse '=GET.DOCUMENT(1)' --excel 1.5
```

### Building tokens by hand

You can also build an expression yourself. The common tokens:

| Token | Bytes | Meaning |
|-------|-------|---------|
| `0x03`…`0x0E` | 1 | `+ - * / ^ & < <= = >= > <>` |
| `0x17 len text` | 2+ | String constant |
| `0x1E nn nn` | 3 | Integer constant |
| `0x1F` + 8 bytes | 9 | Number constant (big-endian double) |
| `0x23` + 7 bytes | 8 | Defined name; its 1-based index goes in the fix-ups |
| `0x25 r1 r1 r2 r2 c1 c2` | 7 | Area |
| `0x41 ff` | 2 | Function `ff` with its fixed argument count |
| `0x42 nn ff` | 3 | Function `ff` with `nn` arguments |
| `0x44 rr rr cc` | 4 | Cell reference. Row word: bit 15 relative row, bit 14 relative column. |

```csharp
// =AVERAGE($A$1:$A$5)*2: area rows 0–4 of column 0, AVERAGE (5) with 1 argument, integer 2, multiply.
var expression = new byte[] { 0x25, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00, 0x42, 1, 5, 0x1E, 0, 2, 0x05 };
var formula = new ExcelFormula(expression);
```

A formula that uses a defined name passes the name's 1-based index as a fix-up word:
`new ExcelFormula(expression, fixups: [0x00, 0x03])`.

`ExcelFunctions.Get(n)` describes each function: its argument counts, the types it accepts and returns
(`ExcelValueTypes`), and whether it is macro-only, volatile or has a dialog-box form.

## Names and external references

```csharp
foreach (var name in document.Names)
{
    Console.WriteLine($"{name.Name} = {name.GetDefinitionText(document)}");
    if (name.MacroKind == ExcelMacroKind.Command) Console.WriteLine($"  Option-Command-{name.ShortcutKey}");
}
```

Name definitions are parsed expressions too; name tokens inside them refer to other names by 1-based index
into `document.Names`. Excel keeps the names in alphabetical order, ignoring case and accents, and finds a name
typed in a formula by searching that order, so saving sorts names that changed and renumbers the formulas and
definitions that use them. `document.SortNames()` does it on demand. `document.ExternalDocuments` lists the other documents
the sheet links to; each entry's `Names` are the names the sheet uses from that document.

Excel 1.5 can save **version-18** files (`document.Version` is 18) when a document uses something older
releases cannot read; the library reads and writes them, including their longer number-format and window
records. Setting `document.Version` converts the number formats to that version's form.

## Sheet settings, protection and printing

```csharp
var state = document.State;
state.CalculationMode;    // ExcelCalculationMode.Automatic, AutomaticExceptTables or Manual
state.FullPrecision;      // false = precision as displayed
state.Iteration; state.MaxIterations; state.MaxChange;
state.GetColumnWidth(0); state.SetColumnWidth(0, 100);   // 0 hides a column
state.StandardFont; state.StandardFontSize;

document.Protect("secret");     // as Protect Document does; the password is optional
document.Unprotect("secret");   // false, leaving it protected, for a wrong password
```

`document.Windows` holds the saved windows (position, active cell, selection, display options and splits),
`RowPageBreaks` and `ColumnPageBreaks` the manual page breaks, and `PrintSetup` the header, footer, margins
and print options.

## Creating and editing documents

```csharp
var sheet = ExcelDocument.CreateWorksheet();      // or CreateMacroSheet()
sheet.Cells["A1"].SetText("Total");
sheet.Cells["A2"].SetValue(42);                   // stored as Integer
sheet.Cells["A3"].SetValue(3.25);                 // stored as Number
sheet.Cells["A3"].NumberFormat = 2;               // 0.00
sheet.Cells["A4"].SetValue(1234.5);
sheet.Cells["A4"].NumberFormat = sheet.AddNumberFormat("#,##0.00 \"DM\"");   // a user format: index 19
sheet.Cells["A3"].Bold = true;
sheet.Save("Totals");
```

New documents use the defaults Excel 1.03 writes for an empty sheet. The grid grows as you assign cells.

Editing a file you read is safe: only the cells you touch are re-encoded, and everything the library does not
model (unidentified fields, bytes left over from memory, embedded code) is written back exactly as it was read. A
document read and saved without changes is byte-for-byte identical to the original.

## Exporting SYLK

`ExportSylk` writes a worksheet or macro sheet as SYLK, the text format Excel's Save As offers, with the bytes
Excel 1.00–1.06 write, or Excel 1.5 with `ExcelVersion.Excel150` (spec Appendix C): Mac OS Roman text with carriage returns, R1C1 formulas with their current
values, number formats, styles, column widths, names and links.

```csharp
var document = ExcelDocument.Load("Expenses");
document.ExportSylk("Expenses.slk");          // give it Finder type TEXT
document.ExportSylk("Expenses 1.5.slk", ExcelVersion.Excel150);

var output = new ArrayBufferWriter<byte>();   // or into any IBufferWriter<byte>, without allocating per cell
document.ExportSylk(output);
```

The `O` record carries the document's reference style and calculation mode. These are global settings in Excel,
so its own export shows whatever was in effect when it ran.

## Charts

```csharp
using MacExcelReader.Charts;

var chart = ChartDocument.Load("Chart1");
var main = chart.MainChart;
Console.WriteLine($"{main.Type}, stacked: {main.IsStacked}, legend: {main.LegendPosition}");
var scale = main.ValueScale;
Console.WriteLine(scale.AutomaticMaximum ? "auto" : scale.Maximum.ToString());

foreach (var series in chart.Series)
{
    Console.WriteLine($"{series.Name} ({series.CategoriesReference}): {string.Join(", ", series.Values)}");
}

var format = main.DefaultSeriesFormat;      // and main.GetPointFormat(i) for per-series or per-point overrides
Console.WriteLine($"marker {format.Marker}, pie slices pulled out {format.Explosion}%, point labels {format.PointLabels}");
if (chart.OverlayChart is { } overlay)
{
    Console.WriteLine($"overlay {overlay.Type}, shares the main value axis: {overlay.SharesValueAxisScale}");
}
```

To build a chart, give each series its source document and references, as Excel 1.x always has them:

```csharp
var newChart = ChartDocument.Create(ChartType.ColumnOrBar);
newChart.Series.Add(ChartSeries.Create("Sales", [1, 3, 2], "Data", "R2C1:R2C3", categoriesReference: "R1C1:R1C3"));
newChart.Save("Sales Chart");
```

Patterns, markers, line weights and arrowheads are enums (`ChartPattern`, `ChartMarker`, `ChartLineWeight`,
`ArrowHeadWidth`, `ArrowHeadLength`, `ArrowHeadStyle`) named after Excel 1.00's Patterns dialogs.

A chart file stores a copy of the plotted data, so the values are available without the worksheet. Numbers
in chart files use Microsoft's 8-byte decimal format; `MicrosoftDecimal` converts them. `CategoryCount`,
`DrawingMargins`, the area formats' `Box` and the axes' `TickLength`, `TickLabelsAt` and `LineAt` are the layout Excel worked out the
last time it drew the chart; Excel recomputes them, so a writer can leave them as they are. A chart Excel 1.5 saved last
holds the category count as a word instead, which `Excel15CategoryCount` reads. Chart support was
derived from Excel's code and is checked against Microsoft's three sample charts and the charts Excel saved in
experiments, which all read and write back byte for byte.

## Performance

Reading is lazy:

- **Read** parses the header, settings, names and windows, and attaches the grid to the file's cell table.
  It creates no cell objects.
- A **cell** is created the first time you access it; its **text or formula** is decoded the first time you
  read its value. Numbers are read straight from the buffer each time.
- **`NonEmptyCells`** scans the cell table without creating anything.
- **Saving** copies untouched cells byte-for-byte instead of re-encoding them.
- **Saving to a file or stream** (`Save`, `Write`, `ExportSylk`) builds the bytes in a buffer rented from the shared
  pool, so only `ToArray` allocates the file's size.
- The **embedded code segments** (`CodeSegment1`, `CodeSegment2`) and **name definitions**
  (`ExcelDefinedName.Definition`) are `ReadOnlyMemory<byte>` over the file's bytes rather than copies.

To get the most out of it: walk `NonEmptyCells` and read `Value` or `TryGetNumber` from the `ExcelCellRef` rather
than calling `GetCell()`, and use `Read(bytes, copyData: false)` when you own the buffer. See [the README](../README.md#performance) for
benchmark numbers.

## Errors

Every problem with a file's contents is reported as `ExcelFormatException`: a wrong magic number or
version, sections or records that run past the end of the file, unknown formula tokens, and so on. Because
cells are decoded lazily, a damaged cell surfaces when you read it (or when you save), not from `Read`.
Programming errors such as an out-of-range column, a text value over 255 characters or a number that is
infinite or not a number throw the usual `ArgumentException` family, when the value is set rather than when the
document is saved. The test suite corrupts the sample files at random and checks that nothing
else is ever thrown.

## Inspecting files

`ExcelFileLayout.Read(bytes)` reports where each part of a file lies (header, code segments, state block,
section table and the eight sections) without parsing them. The `dumper` tool exposes this and much more:

```sh
dotnet run --project dumper -- layout Expenses
dotnet run --project dumper -- cell REPLACE B27     # every field of one cell, with its tokens
dotnet run --project dumper -- parse "=SUM(A1:A5)"   # the tokens for some formula text
```
