# Testing in a real Excel

The unit tests check the library against Excel's own files. This page covers the other direction: putting
files the library wrote, or any others, into a real Excel 1.x and checking what Excel made of them. It works
for worksheets, macro sheets and charts.

You need [Mini vMac](https://www.gryphel.com/c/minivmac/) (or another Mac emulator), a Mac ROM, a bootable
System disk and an Excel disk image. None are included.

## How it works

```bash
dotnet run --project dumper -- excel-test pack Test.dsk FILE...
```

This writes a 400K floppy image (800K if needed) holding the files and a generated macro sheet, **Driver**.
For each file, Driver:

1. opens it, without updating links, so linked files don't stop it with a question;
2. for each formula cell, records `GET.FORMULA` of it in one of its own cells, which is Excel's own text for
   the formula. It refers to the cell as `!B4`, meaning "on the active sheet", because Excel 1.00 gives
   `#REF!` for a reference to a document by name that was opened after Driver;
3. saves it with Save As as `<name> saved` and closes it.

Finally Driver saves itself, so its cells keep Excel's answers.

| Excel | Running Driver |
|---|---|
| 1.00, 1.01, 1.03, 1.04, 1.06 | Open Driver, then choose Macro > Run > `Driver!Run` (Option-Command-R; Macro > Exécuter in the French 1.01). `check` cannot turn the French 1.01's GET.FORMULA text into A1, so read it from Driver. |
| 1.5 | Open Driver. It runs by itself (Auto_Open). Pack with `--excel 1.5`. |

Start Excel from its own disk, not by opening a document from the Finder, which may pick a newer Excel. To
open a workspace (`XLIN`) file, which only the Finder can do, boot from a disk whose only Excel is the one under
test: Excel 1.00's data disk has its own System. Then:

```bash
dotnet run --project dumper -- excel-test check Test.dsk [-o extracted/] [--all]
```

For each file, `check` reports:

- whether Excel saved it at all. A missing `<name> saved` means Excel couldn't open it or Driver stopped;
- **GET.FORMULA**: every formula whose Excel text, turned from R1C1 into A1, differs from the library's
  decompiled text;
- **re-save**: every difference between the JSON models (`dumper json`) of the original and Excel's re-save,
  such as cells lost or changed, names, number formats, settings and chart contents;
- the byte counts, and how many bytes differ.

By default it skips the changes Excel is expected to make:
- automatic chart scale limits, which Excel works out when it draws;
- recalculated formula results;
- windows;
- the code signature and chart version, which say which Excel saved the file;
- the table size;
- external names no formula uses.

`--all` shows them too. The exit code is 0 when nothing unexpected differs.

To test charts, `dumper newchart TYPE OUT [--series 1,2,3]... [--title T] [--bar]` builds one with the library.

Options for `pack`:
- `--volume NAME`: the disk's name. Driver opens files as `Volume:Name`, so each mounted test disk needs its
  own.
- `--keep-open`: leave every file open, to look at charts and sheets on screen.
- `--no-probe`: open and re-save only.
- `--recompile`: after probing, re-enter each formula from Excel's own text, so the re-saved file holds Excel's
  compilation of it. `check` then shows every formula whose tokens differ from the library's. Formulas that
  refer to other documents are skipped, because re-entering one makes Excel ask for the file. Recompile one
  sheet of about 100 formulas per disk: the extra Driver rows can run Excel 1.03 out of memory.
- `--each FORMULA`: a macro command to run on every file before it is saved. A file can also have its own
  commands in `FILE.commands` next to it, one per line.
- `--end FORMULA`: a macro command to run after Driver has saved itself, for example before quitting.

If OPEN fails, for example because Excel is out of memory, Driver skips that file rather than run its commands on
Driver itself.

Options for `check`:

- `--all`: also show the differences Excel is expected to make.
- `--bytes`: also list raw byte differences the JSON leaves out: the state block, window state, name flags and
  macro-kind bytes, and external-document entries.

`check` always compares each formula's tokens, with the bytes left over from memory cleared.

## Experiments

To find out what a byte means, have Excel change it: build a small sheet, give it a `.commands` file and pack it.

```bash
printf '%s\n' "=1" > calc.txt
dotnet run --project dumper -- sheet calc.txt out/calc
printf '%s\n' '=CALCULATION(3)' > out/calc.commands
dotnet run --project dumper -- excel-test pack Exp.dsk out/calc
# run Driver in Excel, then
dotnet run --project dumper -- excel-test check Exp.dsk --bytes
```

Things to know when writing commands:

- The macro command FORMULA takes **R1C1** text: `=FORMULA("=SUM(R1C1:R2C2)")`. A1 text gives `#VALUE!`.
- CALCULATION and A1.R1C1 are global, so they carry on into every later file on the disk.
- In Excel 1.00 a window's title includes its volume (`Ws:F 01 saved`), which ACTIVATE needs.
- Excel does not recalculate the formulas it loads, so a sheet the library wrote keeps its stored results. Add
  `recalc 1` to the list to have Excel recalculate it on Open.
- Excel 1.00 writes a workspace (`Resume Excel`, next to the application) on every Quit, so leaving files open
  with `--keep-open`, running `--end` commands and quitting shows what the workspace records.

## Formula lists

`tests/Data/formulas-1.0.txt` and `formulas-1.5.txt` hold every formula feature and every function, with
their fewest, one more and most arguments, and dialog-box forms. Build sheets from them:

```bash
dotnet run --project dumper -- sheet tests/Data/formulas-1.0.txt out/F --split 100
dotnet run --project dumper -- sheet tests/Data/formulas-1.5.txt out/G --split 100 --excel 1.5
```

`--excel 1.04` compiles for Excel 1.04, which adds CHANGE.LINK and a third argument to OPEN; pack with the same
option. Excel 1.04 and 1.06 ask for a registration name the first time they start. The 1.06 disk has no System Folder: boot from another system disk and start Excel from the 1.06 disk.

A list is one formula per line, with `## Heading` lines and these setup lines:

```text
width A 220                     column width
external Budget Rate Total      an external document and the names used from it ("!" is the active sheet)
name Total =$B$2                a defined name; add "macro", or "command K" for a command macro with key K
cell D1 =ARGUMENT("x")          a cell outside columns A and B, such as the macro a name points at
{=TRANSPOSE(A1:B1)} 2           an array formula over 2 rows
recalc 1                        state word 0x018 (1 marks the sheet as needing recalculation)
```

The same format works for any new experiment: write a list and build it.

## Running it in Mini vMac

`tools/minivmac/start.sh SYSTEM.dsk EXCEL.dsk Test.dsk` restarts the emulator and mounts the disks in order.
Things to know:

- **Focus:** Mini vMac pauses whenever its window is not in front.
- **Disk limit:** Mini vMac mounts at most six disk images.
- **Ejecting:** ejecting a disk whose file was recently open can leave the Mac asking for it forever. Restart
  instead.
- **Saved files:** Excel's new files reach the disk image while Excel is running, so `check` can be run without
  shutting down. Copy the image first if Mini vMac still has it open.
- **Memory:** Excel 1.03 runs out of memory on a macro sheet of about 500 formulas, and Driver grows with every
  probe. Pack sheets of about 100 formulas, two or three at a time.

## Results so far

| Date | Excel | Files | Result |
|---|---|---|---|
| Nov 2026 | 1.03 | `formulas-1.0` in six sheets (578 formulas) | All formulas match. Found the cell-table margin rule (spec §6.1). |
| Nov 2026 | 1.03 | `F 01`, `CLOCK.CH` through Driver | Formulas match. Excel sets name flag `0x04` on reference names (spec §11). |
| Nov 2026 | 1.5 | `formulas-1.5` sheets 1 and 7, including functions 256 and up, and `CLOCK.CH` | No unexpected differences |
| Nov 2026 | 1.00 | `F 01`, `CLOCK.CH` | No unexpected differences. Excel 1.00 re-saves the version-3 chart as version 12 (spec §16.2). |
| Nov 2026 | 1.00 | A workspace written by `ExcelWorkspace` | Reopened both windows and restored the calculation options (spec Appendix B) |
| Nov 2026 | 1.03 | `F 01` with `--recompile` | Found four parser differences, now fixed: unions of plain references (`26` with cached areas), the class of memory tokens, `=+1` and `=#REF!` (spec §15.4) |
| Nov 2026 | 1.03 | Experiments: CALCULATION, A1.R1C1, SET.PAGE.BREAK, TABLE, DEFINE.NAME, links, intersections | Calculation-mode values, reference style, page-break numbering, the macro-kind byte, state words 0x024 and 0x03E, intersection encoding (spec §6, §11, §14, §15). Tables do not fit the inferred layout (§12.6). |
| Nov 2026 | 1.00 | FULL, PRINTER.SETUP, Set Preferred Format, then Quit | All three unknown workspace fields (spec Appendix B) |
| Nov 2026 | 1.5 | FORMAT.NUMBER with `@` and colours, FREEZE.PANES, DISPLAY colour | Tokens `23`, `24`, window bytes 121, 127, 128 and the version-18 rule (spec §7, §13) |
| Nov 2026 | 1.03 | Every Gallery format and COMBINATION on `PROFIT CHART` (`--bytes`) | Explosion and data-label bits of point formats, overlay option flags, the gallery table (spec §16.5, §16.8, §16.10) |
| Nov 2026 | 1.03 | COLUMN.WIDTH, error formulas, `recalc 1`, `=(R1C1,R2C2)`, TABLE with one input | Column widths in pixels, all seven error codes, state word 0x018 makes Open recalculate, memory error `27` with #NULL! or #VALUE!, table byte 14 (spec §6, §12.4, §12.6, §15.4) |
| Nov 2026 | 1.5 | Re-saved version-18 sheets as samples | A stream position on a block boundary is stored as offset 512 (spec §1.2) |
| Nov 2026 | 1.00 | Chart > Axes on an overlay chart, with the chart code decompiled | Overlay axes are their axis line patterns; the overlay option flags share the main chart's scales; data-label bits and the two built-in text objects (spec §16.5, §16.7, §16.8) |
| Nov 2026 | 1.00 | A sheet linked to `Src`, saved while `Src` was open | External-document entries still start with 6 zero bytes (spec §8) |
| Nov 2026 | 1.00 | Number and date constants, intersections, array constants, `=#REF!`, `*` fill formats, deleting a table input | Integer limit, 1904 dates, array-constant classes (a parser fix), `#REF!` as `2A`, token `0A`, the table's deleted-input bit (spec §7, §12, §15.4) |
| Nov 2026 | 1.00 | Chart code decompiled, checked against the gallery samples | Axis words 8 and 10, the text objects' positions, how point formats are keyed, and that overlay flags 11–10 have no effect (spec §16.5–16.8) |
| Nov 2026 | 1.00 | Charts built with `dumper newchart`, and Chart > Add Arrow | Excel re-saves library charts unchanged apart from its code segment and memory leftovers; fixed two series-header constants and the label strings; arrows are text objects (spec §16.7, §16.9) |
| Nov 2026 | 1.00 | A table saved with a non-zero list link | Open relinks it but does not calculate it, and Save writes 0 (spec §12.6) |
| Nov 2026 | 1.00 | Format > Patterns on an arrow, twice | Arrow pattern, shaft weight, head width, length and style (spec §16.7) |
| Nov 2026 | 1.00 | Format > Patterns on two markers of a line series | Marker values and point-keyed overrides (spec §16.8); the 15 fill patterns came from `PCOD` 0 |
| Nov 2026 | 1.00 | Library charts opened on screen, and variants mixing Excel's and the library's records | Library charts drew nothing: `ChartSeries.Create` left series bit 10 clear and `Chart.Create` wrote an invisible default series format. Fixed; all four library charts now draw (spec §16.9) |
| Nov 2026 | 1.00 | Format > Patterns on a column series; series with and without a source document | Point-format border and area; Excel 1.x series need a source document, references without `=` (spec §16.8, §16.9) |
| Nov 2026 | 1.00 | Two sheets of overflowing text, state word 0x01A 0 and 1 | No visible difference; Excel ran its whole-sheet pass and saved 0 (spec §6) |
| Nov 2026 | 1.00 | Empty arguments and SERIES entered with FORMULA; three more arrowhead settings | Type bit 0x80 and flag 0x02 (spec §15.3); arrow head length and style corrected (§16.7) |
| Nov 2026 | 1.00 | Chart > Add Legend on a library chart | Chart header bit 15 is the legend (spec §16.5) |
| Nov 2026 | 1.00 | Sheets with overflowing text and Fill-aligned runs, saved with state word 0x01A at 0 and 1 | The word has no visible effect (spec §6) |
| Nov 2026 | 1.00 | DCOUNT with field 1, 0, empty and 2, re-entered with `--recompile` | An empty argument whose type has bit 0x80 is "not given" (spec §15.3) |
| Nov 2026 | 1.01 (French) | A formula sheet probed with GET.FORMULA, then NEW(1) and SAVE.AS | One-byte extra header; French built-in names, function names and R1C1 letters; same tokens (spec §2.1, §11, §15.4) |
| Nov 2026 | 1.04 | CHANGE.LINK and a three-argument OPEN with `--excel 1.04 --recompile`, and a sheet with no new functions whose count was set to 209 | Import word 0x740A; function count from the formulas (209, 208); same tokens as the library (spec §2.1, §2.3, §5) |
| Nov 2026 | 1.06 | The same two sheets as the 1.04 run | Same tokens, import word 0x740A and function counts as 1.04 (spec §2.1, §2.3) |
| Nov 2026 | 1.5 | Library line charts with 102 and 103 points, opened and saved | Chart version 1 for more than 102 categories (spec §16.2) |
| Nov 2026 | 1.5 | `SIMUL EQUATIONS` opened and saved by hand; FORMAT.NUMBER with a text section, a trailing comma and two plain sections | The released 1.5 saves `SIMUL EQUATIONS` as version 17; text sections and trailing-comma scaling make version 18 (spec §2.1, §7) |
| Nov 2026 | 1.5 | Functions 256–291 from `formulas-1.5`, with `--recompile` | Excel 1.5's tokens match the library's; saved as the `Functions 256 And Up` sample (spec §15.2) |
| Nov 2026 | 1.00 | `Literals` written by the library with function count 209 and one extra function name, and with 208 | 1.00 refuses a count above 208 ("File is not a saved worksheet") before reading the names; 208 opens (spec §5, §14) |
| Nov 2026 | 1.5 | `formulas-1.5` sheets 2–6 with `--recompile`, and `BREAK CHART` and `PROFIT CHART` | Charts and sheets 3, 4 and 6 match. Array-formula operands take array class only inside array arguments, and 1.5's Save writes `DATA.DELETE?()` as `41 82`; both now in the parser (spec §15.4). GET.FORMULA gives every cell of an array range the anchor's text; the check now reads it from the anchor. |
| Nov 2026 | 1.00 | Six array formulas re-entered with `--recompile` | An operator at the top of an array formula leaves its operands in value class, as in 1.5 (spec §15.4) |
| Nov 2026 | 1.00 | `formulas-1.0` sheets 2–6 with `--recompile`, and IF and CHOOSE probes, some typed | Sheets 3–6 match. Function calls inside IF and CHOOSE follow the ordinary argument rules, typed or through FORMULA; the 1.00 disk's macro sheets carry an earlier compiler's reference classes (spec §15.4) |
| Nov 2026 | 1.5 | `formulas-1.5` sheet 1 with `--recompile` | Matches, apart from GET.FORMULA's whole-row text for `=A1:IV16384` |

The files Excel saved in these experiments are samples in `tests/Samples` (named for what they show, such as
`Errors`, `Column Widths` and `Pie Exploded`), and `tests/ExperimentTests.cs` checks what each one showed. Every
sample is also read, written back byte for byte and, for formulas, recompiled by the unit tests.

Bugs and quirks of Excel found this way are in [Excel bugs and quirks](ExcelBugs.md).
