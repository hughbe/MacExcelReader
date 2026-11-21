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
| 1.00, 1.03 | Open Driver, then choose Macro > Run > `Driver!Run` (Option-Command-R) |
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
- recalculated formula results;
- windows;
- the code signature and chart version, which say which Excel saved the file;
- the table size;
- external names no formula uses.

`--all` shows them too. The exit code is 0 when nothing unexpected differs.

Options for `pack`:
- `--volume NAME`: the disk's name. Driver opens files as `Volume:Name`, so each mounted test disk needs its
  own.
- `--keep-open`: leave every file open, to look at charts and sheets on screen.
- `--no-probe`: open and re-save only.

## Formula lists

`tests/Data/formulas-1.0.txt` and `formulas-1.5.txt` hold every formula feature and every function, with
their fewest, one more and most arguments, and dialog-box forms. Build sheets from them:

```bash
dotnet run --project dumper -- sheet tests/Data/formulas-1.0.txt out/F --split 100
dotnet run --project dumper -- sheet tests/Data/formulas-1.5.txt out/G --split 100 --excel 1.5
```

A list is one formula per line, with `## Heading` lines and these setup lines:

```text
width A 220                     column width
external Budget Rate Total      an external document and the names used from it ("!" is the active sheet)
name Total =$B$2                a defined name; add "macro", or "command K" for a command macro with key K
cell D1 =ARGUMENT("x")          a cell outside columns A and B, such as the macro a name points at
{=TRANSPOSE(A1:B1)} 2           an array formula over 2 rows
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

Bugs and quirks of Excel found this way are in [Excel bugs and quirks](ExcelBugs.md).
