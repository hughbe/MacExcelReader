# Testing in a real Excel

The unit tests check the library against Excel's own files. This page covers the other direction: putting
files the library wrote, or any others, into a real Excel 1.x and checking what Excel made of them. It works
for worksheets, macro sheets and charts.

It runs headless, from the command line or the tests, in [MAME](https://www.mamedev.org/)'s Mac Plus
(`brew install mame`), with no window and nobody at the keyboard, for Excel 1.00, 1.01, 1.03, 1.04, 1.06 and 1.5.
You need the Mac Plus ROM and each release's disks; none are included.
[Mini vMac](https://www.gryphel.com/c/minivmac/), by hand, remains the backup.

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

Finally Driver saves itself, so its cells keep Excel's answers, and shows the alert "Driver finished".

`excel-test run Test.dsk --excel VERSION` runs Driver in MAME ([below](#running-it-in-mame)); pack with the same
`--excel`. By hand:

| Excel | Running Driver by hand |
|---|---|
| 1.00, 1.01, 1.03, 1.04, 1.06 | Open Driver, then choose Macro > Run > `Driver!Run` (Option-Command-R; Macro > Exécuter in the French 1.01). `check` cannot turn the French 1.01's GET.FORMULA text into A1, so read it from Driver; it also adds French built-in names (`Base_de_données` and so on) to every sheet it saves. |
| 1.5 | Open Driver. It runs by itself (Auto_Open). Pack with `--excel 1.5`. |

By hand, start Excel from its own disk, not by opening a document from the Finder, which may pick a newer Excel.
To open a workspace (`XLIN`) file, which only the Finder can do, boot from a disk whose only Excel is the one
under test: Excel 1.00's data disk has its own System. Then:

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
- automatic chart scale values, which Excel works out when it draws ([Display](Display.md));
- recalculated formula results;
- windows;
- the code signature and chart version, which say which Excel saved the file;
- the table size;
- external names no formula uses.

`--all` shows them too. The exit code is 0 when nothing unexpected differs.

To test charts, `dumper newchart TYPE OUT [--series 1,2,3]... [--title T] [--bar] [--log] [--scale min=2]... [--overlay TYPE [--overlay-from N] [--share] [--share-x] [--hide-overlay-axes]] [--hide-value-axis] [--axis-font 3:9]` builds one with the library.

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
- `excel-test junk <input> <output> [--only CATEGORIES]` (a separate command): writes a copy of a sheet with junk in
  every field the spec says Excel ignores or rebuilds, by category, to test those claims in Excel.
- `--save-as TYPE`: save each file with `SAVE.AS` type `TYPE` instead of Normal (2 SYLK, 3 Text, 4 WKS; [SYLK](Sylk.md)). A 400K disk holds about 21 SYLK exports of the samples.

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

## Running an experiment in one go

```bash
dotnet run --project dumper -- excel-test experiment out/Sheet1 out/Sheet2 ... -o results/ [--no-probe] [--recompile] [--each FORMULA]... [--batch N] [--disk-size 400|800]
```

`experiment` does `pack`, `run` and `check` for a whole experiment in Excel 1.00 under MAME (below), however many
files it has, and is the quickest way to ask Excel something:

- **Disks.** It puts at most `--batch` files (30) on a disk and halves a batch whose disk overflows, so any number
  of files can go in one command: 128 one-name sheets ran on nine disks in about 40 seconds.
- **Commands.** Give each file a `FILE.commands`, one macro command per line, in UTF-8 (`pack` refuses anything
  else rather than mangle accented letters). A command that gives an error does not stop the others.
  `results/commands.tsv` lists every command with what it returned, by disk, file and Driver row, and the summary
  shows the ones that gave an error. Wrap a command in `ISERROR(…)` to turn its failure into TRUE.
- **Saves.** Excel's saves are collected in `results/saved/`, each disk's files as Excel left them in
  `results/DISK/`, with `check`'s report in `results/DISK/check.txt` (`--verbose` shows it as it runs). Under
  MAME a save sometimes reaches the disk without its data, or with one 512-byte sector of it as zeros (a formula
  then cannot be read): 2 of 76 disks of about 20 KB saves, on 400K and 800K disks alike, at unrelated places
  (track 3 sector 11, track 20 sector 5). The test disk is ejected, which flushes it, before MAME stops, so the
  cause is **(unknown)**; MAME's conversion of a sector it cannot decode back to the raw image is one candidate,
  and checking it needs MAME's `floptool`, which Homebrew's MAME does not include. `check` treats a sheet as
  readable only if every formula and name in it decodes, and a chart only if no sector of its series lists is all
  zeros (a chart that lost one still reads, its last series cut short; 1 of the 123 scale charts did). It names
  any zeroed sectors, and `experiment` runs those files again on a disk of their own. A check
  that fails on one disk is reported with the rest instead of ending the experiment.
- **Crashes.** A disk that does not finish is split in halves until the file that stops Excel is found; its
  last screen is in `results/DISK/screen.png`. Excel 1.00 crashes on some formulas it should refuse and on some
  number formats ([Excel bugs](ExcelBugs.md)), so give a suspect formula or format a sheet of its own.
- **Recompiling.** With `--recompile`, Driver enters each formula again with FORMULA, from the text GET.FORMULA
  gave. `check` lists the ones Excel refused (FORMULA gave an error, so the cell kept the library's tokens and
  would otherwise pass as a match) under "not recompiled", and `commands.tsv` has a row for each.
- **Typing.** Give a file a `FILE.type` to have Excel 1.00 compile formulas from A1 text, as a person types them:
  one formula per line, the cell, then spaces and the text (`B2 =D:E`), in UTF-8; `#` lines are comments. Driver
  saves and leaves that file open, and once Driver has finished `run` types each formula: it brings the window to
  the front, chooses Formula > Goto (Command-G) and types the cell, then types the text and Return. An alert that
  follows is answered with Return. A formula Excel refuses ("Error in formula") stays in the formula bar, which
  Command-period does not cancel in Excel 1.00, so `run` clicks the formula bar's cancel box; it tells a refusal
  from an entry by whether the formula bar's enter box is still drawn. After the last formula it saves the file
  with Command-S. `pack` puts the list on the disk as a text file, `Typing`, and `run` writes back `Typed`,
  saying for each formula whether Excel took it, which `check` reads: a typed formula Excel refused in a cell our
  copy has a formula for is reported "refused when typed", and in a cell our copy leaves empty `check` compiles
  the text itself and compares the tokens with Excel's. `sheet LIST OUT --type` builds such a sheet from a formula
  list: column A holds each text, column B is left empty and `OUT.type` types each formula into it, so Excel's
  save holds its own compilation of every formula, or nothing where it refused one (the `Typed` samples).
  Typing takes about two seconds of emulated time a formula. Array formulas are not typed.

Things to keep in mind when designing an experiment. `pack` warns about the ones marked *(warned)*, from each
`.commands` line and `--each` and `--end`:

- FORMULA and the other macro commands that take formula text read it as R1C1, whatever `A1.R1C1` says:
  `=FORMULA("=A1+1")` gives `#VALUE!`. Write `=FORMULA("=R1C1+1")`, or `R[-1]C` for relative references
  *(warned)*.
- Some formulas and formats crash Excel 1.00 where it should refuse them ([Excel bugs](ExcelBugs.md)): a number
  format with two exponents in a section after the first, given to TEXT (two in the first section give `#VALUE!`);
  `:` followed by a minus sign; and a function joined by `:` or a space to a union ending in a relative reference
  *(warned)*. A crash costs a bisection of its disk, so give anything suspect a sheet of its own.
- `--recompile` goes through R1C1 text, which cannot say everything A1 text can: whole rows or columns with a
  relative end compile to two areas joined by `:` ([Excel bugs](ExcelBugs.md#getformulas-whole-row-text-does-not-compile-back-the-same-excel-100-103)),
  so token differences there are not differences in how Excel compiles typed A1 text. `check` shows those in grey
  and does not count them; type such formulas instead (`FILE.type` above).
- A typed formula that refers to its own cell makes Excel show "Can't resolve circular references"; the formula is
  still entered. Keep typed formulas' references clear of the cells they are typed into.
- Open grows a worksheet's table to every cell its formulas refer to (spec §6.1). A sheet referring to both
  row 16384 and column IV (`IV16384`, or `IV2` and `B16384` in different formulas) needs about 25 MB, so Excel
  will not open it, and a far reference can make the save too big for the disk. `pack` warns about a sheet whose
  table would pass 1 MB; generated formulas should refer to cells near the top left.
- Excel keeps one name for names that compare equal (spec §11), so a later `DEFINE.NAME` of an equal name
  renames the earlier one rather than adding a second.
- `dumper cell FILE ADDRESS` shows a formula's raw record when the library cannot decode it, to tell a lost
  sector from a token the library does not know. `excel-test files DISK [-o DIR]` lists the files on a floppy
  image, MFS or HFS, and copies them out with their resource forks (`NAME.rsrc`), since macOS no longer mounts
  these disks.

## Running it in MAME

```bash
export MACEXCEL_SYSTEM=~/Excel/Data.img MACEXCEL_PROGRAM=~/Excel/Program.img MACEXCEL_ROM=~/Mac/MacPlus.ROM
dotnet run --project dumper -- excel-test run Test.dsk [--excel 1.03] [--capture "Volume:Name saved"]... [-o screens/]
```

`run` boots Excel in MAME's Mac Plus, in the background (no window, no Dock icon, no focus taken: it starts MAME
with SDL's dummy video and audio drivers and no input providers), runs Driver, writes the disk back with Excel's
saves for `check`, and exits 0 if Driver finished. MAME runs the Mac about 30 times faster than real time: a disk
of three 100-formula sheets takes about five seconds, in any release. It needs:

- `mame` on the `PATH` (or `--mame`);
- the Mac Plus ROM (128K, checksum `4D1F8172`; Mini vMac's `vMac.ROM` is one, with bytes after it that `run`
  ignores), `--rom` or `MACEXCEL_ROM`. `run` splits it into the even and odd halves MAME's `macplus` set wants.
  MAME also wants the keyboard's microcontroller ROM, which `run` does not need: it writes an empty stand-in, so
  MAME warns that `341-0332-a.bin` has the wrong checksum;
- the release's disks, `--program` and `--system` or the variables in the table: the disk (or folder) with Excel
  on it, and a startup disk for its System and Finder where that disk has none.

| `--excel` | Program (`--program`) | System (`--system`) |
|---|---|---|
| `1.0` (default) | `MACEXCEL_PROGRAM`: Excel 1.00's Program disk, the archive.org copy, whose copy protection passes in an emulator | `MACEXCEL_SYSTEM`: Excel 1.00's Data disk (which the renderer also uses for fonts) |
| `1.01` | `MACEXCEL_PROGRAM_1_01`: the French 800K disk (WinWorld) | `MACEXCEL_SYSTEM_1_01`: any startup disk, such as Excel 1.04's |
| `1.03` | `MACEXCEL_PROGRAM_1_03`: the 800K Excel Program disk, which has its own System Folder | not needed |
| `1.04` | `MACEXCEL_PROGRAM_1_04`: the 800K disk (WinWorld), with its own System Folder | not needed |
| `1.06` | `MACEXCEL_PROGRAM_1_06`: the 800K disk (WinWorld) | `MACEXCEL_SYSTEM_1_06`: as for 1.01 |
| `1.5` | `MACEXCEL_PROGRAM_1_5`: a folder holding `Microsoft Excel` with its resource fork and Finder info (macOS only), or a disk with it | `MACEXCEL_SYSTEM_1_5`: as for 1.01 |

The 1.03 and 1.04 disks both have System 3.2 and Finder 5.3, so every release after 1.00 runs under that System.
The Mac Plus has 4MB, MAME's default, which is enough for 1.5.

How it works:

- **Boot disk.** Excel 1.00's copy protection checks the disk Excel runs from, so for 1.00 `run` keeps the
  Program disk as it is, grows it to 800K and adds the System and Finder, then copies in the Data disk's boot
  blocks with Excel as the startup application (`bbHelloName`). The later releases came on HFS disks, whose
  System looks for the startup application only in the System Folder, so `run` lifts Excel, its help file, the
  System and the Finder onto a new 800K MFS disk, which has no folders, with the boot blocks of the disk the
  System came from. None of them checks for its master disk. The Mac boots straight into Excel, with the test
  disk in the second drive. Excel 1.01's application, `Microsoft Excel 1.01F`, is renamed `Microsoft Excel` there,
  since the boot blocks hold 15 characters.
- **Starting.** Excel 1.04, 1.06 and 1.5 ask for a name to personalize the copy every time, since the boot disk
  is new each run: the script types one and clicks OK, which Return does not choose. Other dialogs at startup,
  such as 1.5's opening screen, close by themselves.
- **Keys.** MAME's Mac Plus keyboard cannot run without its microcontroller's ROM, so `dumper/excel-test.lua`,
  MAME's autoboot script, posts key and mouse events into the Event Manager's queue itself, as PostEvent does. It
  opens Driver with Command-O, typing its name to select it in the Open dialog; runs it with Option-Command-R
  (Excel 1.5 runs it by itself, Auto_Open); waits for the "Driver finished" alert, answering and logging any
  other dialog with Return; types any `FILE.type` formulas ([above](#running-an-experiment-in-one-go)); and
  finally ejects the test disk with Command-Shift-2, which flushes it. A click is a mouse-down and a mouse-up event
  posted the same way, at the pointer (its three low-memory copies moved to the spot) or at a dialog item's middle,
  with the button state (`MBState`, `$172`) set to down and back to up.
- **Windows.** The script skips Excel's own windows (kind 8): the formula bar, which 1.00 leaves untitled and
  1.5 titles with its font's name (`Times`), and the Clipboard.
- **Captures.** For each `--capture` title, the script brings windows to the front with Driver's second macro,
  `Next` (Option-Command-W, `ACTIVATE.NEXT`), until that one is in front, then copies the screen buffer and
  crops it to the window, title bar, frame and shadow, as `tests/Data/Screens` holds them. The pointer stays in
  the top left corner, over the Apple menu. `TITLE=FILE.png` names the PNG. Pack with `--keep-open`. Excel 1.00
  titles the window of a file on another disk with its volume (`Render:R Basics saved`); 1.03 and 1.5 do not
  (`R Basics saved`), so a title with a volume also matches the window without it. Driver cannot bring a window
  to the front with `ACTIVATE`, which gave an error for every saved file's title (`Render:R Basics saved`,
  `R Basics saved`, `Render:R Basics`, `R Basics`) when run from a second macro in 1.00.
- **Failures.** If Driver does not finish within `--timeout` seconds of emulated time (default 3600), `run`
  leaves the disk unchanged and keeps MAME's working folder, with the last screen as `screen.png`. `--work DIR`
  keeps it anyway, `--screen FILE` saves the last screen, and `--show` opens MAME's window at real speed to
  watch.

`excel-test experiment` takes the same `--excel`, `--program` and `--system`. `ExcelRunTests` runs Render Basics
through `pack`, `run` and `check` and compares the capture with its golden for 1.00, and runs it in each later
release whose variables are set (only that the run finishes for the French 1.01). Each test passes without
checking anything when its variables are not set or MAME is not installed.

## Running it in Mini vMac

Mini vMac is the backup, for watching a run by hand or anything the MAME script cannot do.
`tools/minivmac/start.sh SYSTEM.dsk EXCEL.dsk Test.dsk` restarts the emulator and mounts the disks in order. Things
to know:

- **Focus:** Mini vMac pauses whenever its window is not in front.
- **Disk limit:** Mini vMac mounts at most six disk images.
- **Ejecting:** ejecting a disk whose file was recently open can leave the Mac asking for it forever. Restart
  instead.
- **Saved files:** Excel's new files reach the disk image while Excel is running, so `check` can be run without
  shutting down. Copy the image first if Mini vMac still has it open.
- **Memory:** Excel 1.03 runs out of memory on a macro sheet of about 500 formulas, and Driver grows with every
  probe. Pack sheets of about 100 formulas, two or three at a time.

## Comparing drawing with Excel's screen

To check `ExcelRenderer` against Excel, pack sheets with `--keep-open` and capture their windows with
`excel-test run --capture`. `tests/Data/Screens` holds each capture as a one-bit PNG, with `index.txt` naming the
sample Excel saved and the window's title, and `ExperimentTests.Rendering_MatchesExcelsScreen` draws Excel's save
of each and compares every pixel. It needs the System's fonts and passes without checking anything otherwise:
set `MACEXCEL_SYSTEM` to Excel 1.00's Data disk image (or a System file's resource fork) to run it. To capture
them again, pack each volume's samples under the names in their titles:

```bash
cp "tests/Samples/Excel 1.00/Render Basics" "out/R Basics"
dotnet run --project dumper -- excel-test pack Render.dsk "out/R Basics" --volume Render --keep-open --no-probe
dotnet run --project dumper -- excel-test run Render.dsk --capture "Render:R Basics saved=Render Basics.png" -o tests/Data/Screens
MACEXCEL_SYSTEM=~/Excel/Data.img dotnet test --filter Rendering_MatchesExcelsScreen
```

Without MAME, capture them in Mini vMac: pack with `--keep-open`, run Driver, bring each window to the front from
the Window menu and capture Mini vMac's window at its exact pixels with macOS's `screencapture -o -l<window id>`
(from a terminal allowed to record the screen). Every emulated pixel is a 2 × 2 block in the capture, below a
28-pixel title bar: halve it and crop it to the window, title bar, frame and shadow, as one-bit PNGs like those in
`tests/Data/Screens`. Captures made this way include the mouse pointer, which `run --capture` keeps out of every
window, so they differ from the renderer where it lies.

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
| Nov 2026 | 1.00, 1.5 | TEXT of every sample format for ten values; dates around 2040; `=""&A1` for General | The library matches every result after fixes: the zero section for values that round to zero, `-` before single-section dates, rounding to the second, the 2040 date limit, and General's 14 digits and range ([Display](Display.md)). TEXT will not take General or fills. |
| Nov 2026 | 1.00, 1.5 | FORMAT.NUMBER with the formats in `tests/Data/number-formats-1.0.txt` and `-1.5.txt` | Every one compiles to the bytes `ExcelNumberFormat.Parse` gives (spec §7) |
| Nov 2026 | 1.00, 1.5 | Questions from comparing with BIFF2: TRUE, FALSE, errors and whole numbers typed with FORMULA; TRUE, FALSE and errors in array constants; DEFINE.NAME with relative references, used from other cells | Constants are cell type 4 (a Boolean-or-error flag and a value); integer cells are unsigned, so 40000 is an integer and −7 a number; array elements of types 4 and 16; a relative part of a name's reference is an offset from the cell using the name, as in BIFF2. Both releases save the same bytes (spec §9, §11, §12.5; `Constant Cells`, `Relative Names`, `Array Constant Types`) |
| Nov 2026 | 1.00 | Every sample exported with `--save-as 2` (SYLK), compared with the library's reading, then with its own SYLK export; cell-type switches decompiled; whole rows and columns and long constants typed | Values, formats, styles, protection and names all agree. Found the 1.00 format-text rule (`;;` shows as `;`), whole-row and whole-column text, 14-digit constants and that cell type 7 is never written. `ExcelDocument.ExportSylk` now writes the same bytes for all 84 (`tests/Data/Sylk`) (spec §7, §9, §15.4, [SYLK](Sylk.md)) |
| Nov 2026 | 1.5 | 94 samples exported with `--save-as 2 --excel 1.5` (all but `Worksheet1`, which failed to export), and 1.5's writer compared with 1.00's | Three changes: export-form format text with every literal escaped, a whole `F` record for every cell of a grid with a margin, and the gridline colour. `ExportSylk(…, ExcelVersion.Excel150)` writes the same bytes for all 94 (`tests/Data/Sylk 1.5`). A format of only a text section shows as `@` (spec §7, [SYLK](Sylk.md)) |
| Nov 2026 | 1.00 | Every sample with junk (`A5`) in every field the spec called ignored or rebuilt (`excel-test junk`), opened and saved by Driver; categories bisected on `Table Relinked` and `COMMISSION TEMPLATE` | 61 files across four disks open and save back with only the expected changes (name flag bits cleared). Formula bytes 8–10 are not junk: junk there makes Excel recalculate and can hang it. Open resets most state-block pointers (`10.00`) but keeps 0x028, 0x030, 0x03C and the print-title values, and does not clear the margin row's transient marks. A long Driver run of fully junk-filled files (disks A2 and A3) still crashes or hangs Excel saving a large sheet (`TEXT Formats 1`, `COMMISSION TEMPLATE`, system errors 3 and 25), while the same runs with the clean files finish. Every category on its own passes over the same run, including the window pointers and the name flags; names and windows together fail, so it takes junk in more than one place, building up over a session (spec §6, §9, §12.3). The window pointers are not the cause: Open reads each window record and its selection list, then overwrites offsets 23 and 29 with the new selection list and the document before anything uses them (`46.06`). The cause, found later: junk in the row index (§8), which Open treats as a size and allocates from, about 84 KB per row at `A5 A5`. With it Excel runs out of memory after a few files and hangs saving a large sheet (`COMMISSION TEMPLATE` after three others); the same files with `01 01` there save. No other single group of fields triggers it, and the junk command now leaves the row index out by default |
| Nov 2026 | 1.00 | 181 formulas probing arithmetic, rounding, dates, text, comparisons and errors, on worksheets whose formulas were marked as waiting to be recalculated (`FF FF` at offset 8) | Excel calculated them on open; the results are in `tests/Data/calc-1.00.txt` and described in docs/Calculation.md. IEEE double arithmetic, its own elementary functions, SIGN(-0) = −1 and wrong DATE results (`Calc 1`, `Calc 2`) |
| Nov 2026 | 1.00 | `Expenses`, `DB.FORM` and `LOAN1` with junk in the code-segment count (header word 12) or in the section table | All six open and save normally: Open reads neither. The library now reads them too, finding the sections by reading through them (spec §4, §5) |
| Nov 2026 | 1.00 | `Dirty 0000`, `Dirty FFFF`, `Dirty FFFE` and `Dirty 1234`: =1+1 and =A1*10 saved with the value 99 and those bytes at formula offset 8 | `FFFF` and `1234` recalculate on open (2 and 50); `0000` and `FFFE` keep 99. Offset 8 is the formula's link in the list of formulas waiting to be recalculated (spec §12.3) |
| Nov 2026 | 1.00 | GET.FORMULA of nine long number constants | Constants show as General does: 14 digits, rounded, a tie rounding down; all nine match the library (spec §15.4) |
| Nov 2026 | 1.00 | Drawing code decompiled; `Fit` (19 General numbers in 13 column widths, plain, bold and italic), `Spill` and `More` (alignments, neighbours, fills, hidden columns) and `Zero` (negatives that round to nothing) built by the library and read off the screen, with pixel positions measured from a screenshot | The room is the column less 4 pixels and the italic allowance; General's digits, the `#` count and centring, alignment offsets, where text runs on and stops, and fill counts, all in pixels of Geneva 10. A negative number that rounds to nothing shows as `0` ([Display](Display.md); `tests/Data/display-1.00.txt`, `DisplayTests`) |
| Nov 2026 | 1.00 | `Order` and `Circ` (every formula marked waiting) and six `It` sheets (iteration set with `CALCULATION`, formulas entered with `FORMULA`); the calculate loop and Open's reading of the state block decompiled | Recalculation follows dependencies; circular references without iteration keep their values; iteration stops at the first iteration whose every change is strictly below the maximum change, or at the maximum count; the entered cell goes first. Open takes the calculation settings from a file only when no other document is open (docs/Calculation.md, spec §6). Excel's saves are the samples `Recalculation Order`, `Circular No Iteration`, `Iteration Count`, `Iteration Five`, `Iteration Converge`, `Iteration Big Change`, `Iteration Pair` and `Iteration Settings Kept` |
| Nov 2026 | 1.03, 1.5 | The junk test (`excel-test junk`, default categories) on 13 samples in Excel 1.03 (`Expenses`, `DB.FORM`, `MATRIX`, `Array Formula Classes`, `Relative Names`, `Table Deleted Input`, `Column Widths`, `Errors`, `Intersections`, `Page Breaks`, `Table`, `Macro Names`, `Recalculated`) and 12 in Excel 1.5 (`ACCOUNTS MACROS`, `BEARINGS`, `CASHBUDGET`, `COMMISSION TEMPLATE`, `Colour Formats`, `Constant Cells 1.5`, `Format Text Section`, `Formulas 1.5 Sheet 1`, `Frozen Panes`, `INVENTORY`, `INVOICE MACROS`, `LOAN1`), opened and re-saved with `--no-probe` | Every file opened and saved back with the same content, as in 1.00. The only differences: the junk in the names' flag bytes, which Excel clears; the calculation mode, which Excel takes from its own settings (spec §6); and Excel 1.5's usual rewrite of the pre-release `42 01 39` calls in `INVOICE MACROS` (Excel bugs) |
| Nov 2026 | 1.00, 1.03, 1.5 | An empty worksheet built by the library (825 bytes), and cut-down copies: without its names, without its window record, without both (701 bytes), and cut off before section 7, after the cell data, after the cell table, after the section table, after the state block and after the header | All three releases open the first four and refuse the rest; Excel 1.03 crashes on the one missing only section 7. Saved back, a sheet gets Excel's five built-in names; one with no window record opens with no window and is saved without one, and the Driver's `CLOSE()` then closes the Driver itself (spec §17.2). Excel's saves are the samples `Minimal Bare` (1.00, 1.03, 1.5) and `Minimal No Window` (1.00) |
| Nov 2026 | 1.03, 1.5 | The drawing sheets `Fit`, `Spill`, `More`, `Zero` and `Gen11b` opened in Excel 1.03 and 1.5 and compared with Excel 1.00's screenshots pixel by pixel | In Geneva 10 both draw every cell exactly as 1.00 does. In Geneva 11 (`Gen11b`), 1.5 draws the cells as 1.00 does, while 1.03 adds a pixel per character when checking whether text fits. The heading letters differed because 1.03 and 1.5 ran under a later System than 1.00: Excel 1.00 under that System matches 1.5 pixel for pixel, so the headings come from the System's Font Manager ([Display](Display.md); `DisplayTests`) |
| Nov 2026 | 1.00 (patched) | A copy of Excel 1.00 with its base function count lowered from 208 to 200 (one word of start-up p-code); FORMULA entered `=FULL()` in a macro sheet with count 200, saved, then reopened, and reopened with one name changed | Save wrote count 208 and the names of functions 200–207 at the start of section 7, as the spec describes from the code; Open accepts the file and shows "File error" for the changed name, leaving the document open (spec §14). The file is `tests/Data/Function Names 1.00 Patched` |
| Nov 2026 | 1.00 | The 18 single-chart samples with `A5` in every chart field the spec called unused or left over (chart header byte 5, the scale pointer, byte 137, the drop-line and hi-lo-line unused bytes, the pointers at 262, 358 and 362, the built-in text objects' string pointers and font byte 0, and the default point format's bytes 0 and 5 and unused bits), opened and saved by Driver; the unjunked charts as a control; then one field group at a time on `Area Drop Lines` | All ignored except record offset 362, the head of the in-memory list of data-point formats: with junk there and no formats, Excel lost the series and wrote garbage records (`tests/Data/Point Format Head Junk 1.00`). The other 17 fields re-saved cleanly on all 18 charts (spec §16.5) |
| Nov 2026 | 1.00 | The rare parts of the format: values at A16384, IV1 and IV16384 and arrays at the corner (`Edge`, then `Edge Column` and `Edge Row Array` separately), NEW.WINDOW twice with SELECT, MOVE and SIZE (`Three Windows`), PAGE.SETUP on a chart (`Chart Page Setup`), and quitting with those open (`Resume Excel 3`) | A value in the last row or column gives a margin row or column past the sheet (16385 rows, 257 columns), and arrays there save and reload; both edges at once run Excel out of memory. Several windows save newest first; the chart page setup and a many-window workspace are as the spec describes (spec §6.1, §13, §16.6, Appendix B). `Edge Row Array` (329 KB) is not kept |
| Nov 2026 | 1.03, 1.5 | The same three-window test (`Three Windows 1.03`; in 1.5 with frozen panes in one window and red gridlines in another, `Three Windows 1.5`), then Quit with the windows open | Both save the windows newest first, as 1.00 does; 1.5 saves version 18 with each window's panes and colour in its own state. Neither disk showed a new workspace after Quit, because the disks were copied before the emulated Mac flushed its cache (see the next row) (spec §13, Appendix B) |
| Nov 2026 | 1.03, 1.5 | Started Excel, opened one document (`Expenses` in 1.03, `Formulas 1.5` in 1.5), chose Quit, then Shut Down in the Finder before copying the disks | Both wrote a workspace: 1.03 to its program disk (`Resume Excel 1.03`), 1.5 to the startup disk's System Folder (`Resume Excel 1.5`), in 1.00's layout. Before Shut Down, 1.03's disk had the new data but the old length and date (spec Appendix B) |
| Nov 2026 | 1.00 | QA of the library: number cells holding NaN and infinities with formulas on them; names written out of order, then typed formulas and DEFINE.NAME; FORMAT.NUMBER and TEXT with 15 to 200 decimal places; Mac Roman characters on screen; number constants entered with FORMULA, bisected | Excel keeps NaN and infinities and calculates with them but never makes one (overflow is `#NUM!`); General shows them as ±1.7976931348623E+308 and -2.6965397022935E+308, other formats as `#`. Out-of-order names break typed lookups (`#NAME?`), so the library sorts changed names on save. 128 or more decimals are plain decimals (the header difference is 8-bit). `0xDB`, `0xDE`, `0xDF` and `0xF0` have no glyph in 1.00's System. Constants are taken from 1E-307 to below 1E+306 by value; outside, a formula is refused and a constant becomes text (spec §1, §7, §11, §12.1, §15.4; `NaN Cells`, `NaN Wide`, `Names Unsorted`, `Many Decimals`, `Huge Constants`, `Constant Limits`, `Constant Range`, `Constant Edges`). Excel 1.00's master disk image fails its key-disk check under Mini vMac; use the archive.org copy |
| Nov 2026 | 1.00 | Ranges with an intersection of cells that do not overlap on a side (`=(R1C1 R2C2):R3C3`, both sides, inside SUM, and inside a union), entered with FORMULA, to settle a case the old compiler crashed on | A side with no area makes the range stand for none, and no area is a #NULL! memory error (`27`) in every class, not only value class. The rewritten compiler follows both rules (spec §15.4; `Empty Intersection Ranges`) |
| Nov 2026 | 1.00 | TEXT of 27 values in 11 formats (297 cases), chosen where the double's binary expansion, Excel's 14-digit conversion and rounding ties give different text | Every number format rounds Excel's 14-digit conversion half up to its places, not the double itself: 50 cases the library got wrong before, all now matching. TEXT gives #VALUE! past 255 characters ([Display](Display.md); `TEXT Digits 1`–`3`) |
| Nov 2026 | 1.00 | Sheets built to show text, bold, italic, alignments, fills, borders, a selection, scrolling, no gridlines, no headings, formulas shown, odd and hidden column widths, `Expenses`, Geneva 11 and 12, New York 10 and Monaco 9 as the standard font, and 17 scroll positions to row 16383, opened with `--keep-open` and each window captured from the screen | `ExcelRenderer` draws all but Geneva 11 pixel for pixel, apart from the mouse pointer: headings, gridlines and their pattern phase, clearing under spilled text, QuickDraw's italic, the selection and active cell, Excel's split boxes, the thumbs, and the grey past row 16384. The vertical thumb fits the 17 positions but its rule is inferred; Geneva 11, stretched from 12, differs in about 1,000 pixels (spec §13, [Display](Display.md); the `Render` samples, `tests/Data/Screens`) |
| Nov 2026 | 1.00 | The same captures, with the vertical scroll bar's code decompiled (`18.0c`, `18.12`) and the Mac Plus ROM's text drawing read (`FMSwapFont`, `DrText`, `StretchBits`) | `ExcelRenderer` draws every capture pixel for pixel, Geneva 11 included, apart from the mouse pointer. The vertical scroll bar runs to 17982, with the first and last 20 rows taking 819 each; a scaled size is 234/256 of Geneva 12, not 11/12, which fixes the text widths, and StretchBits ORs together the rows and columns it shrinks ([Display](Display.md)) |
| Nov 2026 | 1.00 | The headless harness (`excel-test run`, MAME's Mac Plus): the 27 `Render` windows captured again, `TEXT Digits 1`–`3` run again from their sheets and checked with `check` | Every window equals Excel 1.00's screen under Mini vMac outside the pointer, which the old captures held, and `ExcelRenderer` draws every pixel of each, frame and shadow included; the screens are now these captures. The TEXT sheets' saves have the same cells as the Mini vMac run and the samples, and `check` gives the same result. Excel 1.00's copy protection checks the disk Excel runs from |
| Nov 2026 | 1.00 | `DEFINE.NAME` with every Mac OS Roman character from `0x80` to `0xFF` in a name, one sheet each after 31 plain names, run with `excel-test experiment` (128 sheets on nine disks in about 40 seconds); then twelve names and their equivalents (`Name Equivalents`) | Names compare ignoring case and accents, with `Æ`/`æ` as AE, `Œ`/`œ` as OE and `Ø`/`ø` as O; equal names are one name, which a later definition renames; `ß`, symbols and the capitals `0xE5`–`0xF4` are refused. The library's name order and formula name lookup now follow this (spec §11; `tests/Data/name-characters-1.00.txt`). About one save in twenty under MAME reached the directory without its data; `experiment` runs those again |
| Nov 2026 | 1.00 | `Referenced Cells` (seven formulas referring to cells, whole rows and columns, areas to the last row and column, a name and another document) and its macro-sheet twin, opened and saved; then names defined with R1C1 offsets and used from other cells, typed by FORMULA in new sheets (`Name Relative Down`, `Off`, `Wrap Row`, `Wrap Column`) | Open grows a worksheet's table to every cell its formulas refer to, names' definitions included, plus the margin; areas reaching the last row or column don't count in that direction, other documents and macro sheets don't count at all. A relative name counts from the cell that uses it and wraps round the sheet; a formula whose table would not fit in memory gives `#VALUE!`. Typing a formula grows a new sheet's table in the count form, without a margin. The library sizes the tables it writes the same way (spec §6.1, §11) |
| Nov 2026 | 1.00 | TEXT of three values in ten scientific formats, among them a run after the exponent with and without a `.` before it (`0E+00 0`, `0E+00.0`) | The exponent fills the digit runs after the `E` up to a `.`, padded to the placeholders after the `E`; runs after the `.` show only their 0s. The library showed `0E+00.0` as `1E+03.0` and `0E+00 0` as `1E+003 0` before; all ten now match ([Display](Display.md)) |
| Nov 2026 | 1.00 | 960 random formulas on 12 sheets, recompiled by Excel (`--recompile`); then the refused and crashing ones narrowed down with FORMULA in R1C1, one per sheet | Excel's compiler takes only references (or IF, CHOOSE, INDEX and other functions that can return one) as operands of `:`, intersection and union, and as arguments typed R alone; 62 formulas the library compiled were refused, and the library now refuses them too. Excel pairs `cell:cell` into an area before other ranges (`=B$2:$C3:B2:C3`), which the library now does. FORMULA reads R1C1 only, and R1C1 text turns relative whole rows and columns into two areas, so those differences were not real. Two refused shapes crash Excel instead (spec §15.4, [Excel bugs](ExcelBugs.md)) |
| Nov 2026 | 1.00 | The 1,042 TEXT results of the earlier random formats again, with Excel's display routine (`19.0d`) and format compiler (`07.09`) read from their code; then 1,200 new random formats (`TEXT Formats 01`–`23`; seven sheets crashed Excel) | The library's display now follows `19.0d` token by token: a second `.` prints and the decimals go on, each decimal run pads on its own, integer digits no run took come out at the point, the exponent's digits spread over the runs after the `E`, percent signs do not scale a scientific number, an empty section still shows `-`, and only the hour just before AM/PM is 12-hour. 237 differences became 5, and on the new formats 810 of 815 agree; every remaining one is a format whose exponent count Excel's compiler writes onto the `E` when it pads a thousands format, which the library now does too, leaving a count byte Excel fills from memory (spec §7, [Display](Display.md)) |
| Nov 2026 | 1.01, 1.03, 1.04, 1.06, 1.5 | `excel-test run` and `experiment` in MAME: Render Basics in each release, captured; eight 1.03 samples (`Column Widths`, `Errors`, `Intersections`, `Macro Names`, `Table`, `Combination`, `Pie Exploded`, `Recalculated`) and seven 1.5 samples (`Colour Formats`, `Frozen Panes`, `Format Text Section`, `Functions 256 And Up`, `BEARINGS`, `LOAN1`, `PROFIT CHART`) on two disks each | Every release runs headless from its own disks under System 3.2, about four seconds a disk. Excel's saves match, apart from `Macro Names`' calculation mode (Excel's own setting) and the French built-in names 1.01 adds. 1.03 and 1.5 leave the volume out of a window's title; 1.5 titles its formula bar with its font's name |
| Nov 2026 | 1.00 | 116 formulas typed in A1 text with `sheet --type`: whole columns and rows with every placing of `$`, ranges, intersections and unions of them, §15.4's range and reference-operand cases, signs before whole rows, and unions without parentheses mixed with other operators | Typed `=D:E`, `=$D:E` and `=30:31` are one area, unlike their R1C1 text. Three parser fixes: a comma outside a function's arguments is a union binding more tightly than every operator but range and intersection (`=1+D30,E31` is `1+(D30,E31)`), whole rows can follow an intersection's space, and a sign against digits makes a number (`=-30:31` is refused); a union with an empty side is a #NULL! memory error. The library now compiles every typed text to Excel's tokens and refuses the 20 Excel refuses (spec §15.4; `Typed Columns`, `Typed Rows`, `Typed Ranges`, `Typed Signs`, `Typed Unions`) |
| Nov 2026 | 1.00 | The seven random-format sheets that crashed Excel, split into 280 formats with TEXT and FORMAT.NUMBER on sheets of their own | Only TEXT crashes, and only for a format with two exponents in a section after the first, whichever section the value uses; with them in the first section TEXT gives `#VALUE!` (`TEXT Two Exponents`), and FORMAT.NUMBER refuses both. `pack` now warns about these and the other crashing shapes ([Excel bugs](ExcelBugs.md)) |
| Nov 2026 | 1.00 | `Calc 3`: 166 function results, among them square roots, exponentials, logarithms and trigonometry up to 1E9, MOD and ROUND at their limits, financial functions, statistics, dates and text, compared with correctly rounded values | `LN` and `LOG10` were exact, `EXP` and `SQRT` at most a unit out, the trigonometric functions within 2 units up to 3 and up to 189 at 1E8, then #NUM! at 1E9; `MOD` fails for large quotients, `ROUND` works on the 14-digit conversion, times round to the second, and `RATE` without a guess fails over 360 periods ([Calculation](Calculation.md); `tests/Data/calc-1.00.txt`, now checked against all three Calc samples) |
| Nov 2026 | 1.00, 1.04 | The System files of Excel 1.00's Data disk and of Excel 1.04's disk (System 3.2), read with `excel-test files` | The Data disk's System has no `FOND` resources; System 3.2 has five, and Geneva's gives bold an extra width of 0, which is why a scaled size's bold headings are lighter under it ([Display](Display.md)) |
| Nov 2026 | 1.00 | 123 library charts with chosen data, fixed scale fields, stacked, 100%, logarithmic and scatter formats, and windows from 70 to 330 points (`newchart --scale`, `--log`, `SIZE`), with the scale code (`37.0a`, `37.16`, `37.2d`) decompiled | Excel saves the automatic scale it drew with; the library now works it out to the bit for every chart sample, Excel's own from 1.03 and 1.5 included: the 8% rule for starting at zero, the major units 10, 20, 50 × a power of ten, stacked totals counted positive, logarithmic scales as exponents, and the label count that fits the plot area (`Scale` samples; [Display](Display.md)) |
| Nov 2026 | 1.00 | 30 more library charts: fixed ends beyond the data, and overlays with their axes shown and hidden, shared and not (`newchart --overlay`, `--share`, `--hide-overlay-axes`); four chart windows captured at heights from 130 to 250 points | A fixed end beyond the data gives a negative range, whose major unit Excel takes from its size, and the zero test that divides by 0 never starts at zero. Excel works out the overlay's shared-scale flags again at each drawing (`37.34`) and plots both charts on their combined data (`37.2f`); the library now does both, matching all 139 chart saves. On screen the labels keep their size in every window, and the value axes' interval counts fit the plot height ÷ the label height ([Display](Display.md)) |
| Nov 2026 | 1.00 | 495 library charts in threshold sweeps: line charts at 17 window heights in the System font and 9 each in Geneva 9 and New York 9, line charts at 7 widths, and bar charts at 13 widths, each with 9 data ranges that change the major unit at one interval count (`newchart --axis-font`); `37.08`, `37.10` and `37.0e` decompiled | `37.0a` counts labels against the plot area before the labels are laid out. Each window size gives one interval count, the same for every data range; width makes no difference to a vertical axis. The counts first seemed to fit labels a pixel shorter than their fonts, until the label layout (next row) explained them ([Display](Display.md)) |
| Nov 2026 | 1.00 | The label layout decompiled (`37.09`, `37.0e`, `37.0f`, `37.2a`, `37.2b`) and checked against the 495 sweep charts | The room the labels take is their extent (text plus a pixel of margin round it; 100 twips more beside the value labels across), half the other axis's labels overhanging, and two tick lengths, on each side; the count divides by the font's ascent and descent. Worked back from the saved layout, it reproduces the scales of all 728 chart saves (`ChartDocument.GetLabelSpaces`; [Display](Display.md)) |
| Nov 2026 | 1.00 | 21 column and bar charts Excel made with NEW from text categories in a worksheet (`sheet --worksheet` with `cell` lines), long labels, single long words, hyphens, ten narrow categories and small windows; the label layout's line breaker decompiled (`33.1d`, `33.1e`, `33.1f`, `33.21`) | Category labels wrap to their slot a word at a time, a word running to the next space or hyphen, and a word too wide breaks between characters; each line takes 240 twips of Geneva 9 off the plot area, exactly as the library now works out (`Labels` samples, `ChartDocument.GetTickLabelSize`). Charts Excel makes label in font 1, the application font, which is Geneva ([Display](Display.md)) |
| Nov 2026 | 1.00 | The chart layout decompiled (`37.03`, `37.04`, `37.07`, `37.0f`, `37.2c`, `38.02`) and run forward from each save's chart area | The chart area, border, title (its box plus 5% of the chart's height), tick length (the smaller side ÷ 70) and the labels' room give the plot area Excel saved for all 716 charts without a legend, axis titles or overlay; number labels are formatted as General in their slot's width, only the ends (and ends ± a unit) measured. A library chart never drawn gets Excel's layout and scales ([Display](Display.md), `ChartDocument.LayOut`) |

The files Excel saved in these experiments are samples in `tests/Samples` (named for what they show, such as
`Errors`, `Column Widths` and `Pie Exploded`), and `tests/ExperimentTests.cs` checks what each one showed. Every
sample is also read, written back byte for byte and, for formulas, recompiled by the unit tests.

Bugs and quirks of Excel found this way are in [Excel bugs and quirks](ExcelBugs.md).
