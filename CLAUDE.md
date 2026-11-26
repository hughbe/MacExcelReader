# MacExcelReader

A .NET library that reads and writes Excel 1.x for Macintosh files, and `docs/Spec.md`, the format
specification it is built from. The library is unreleased, so its API can change.

## Layout

- `src/`: the library. `dumper/`: the `dumper` command-line tool, including `sheet` (build test sheets) and
  `excel-test` (the harness for running files in a real Excel). `tests/`: xUnit tests over `tests/Samples`.
- `tools/`: `pcode.py`, the p-code disassembler (`python3 -I tools/pcode.py EXCEL.rsrc 18.12` lists a procedure
  from the application's resource fork), and `minivmac/` for running the harness by hand in releases other than 1.00
  (`excel-test run` runs it headless in MAME for 1.00).
- `docs/Spec.md`: the format. `docs/Display.md`: how Excel formats values and draws cells and windows.
  `docs/Calculation.md`: what Excel calculates. `docs/Sylk.md`: Excel's SYLK export. `docs/ExcelBugs.md`: Excel's
  own bugs and quirks. `docs/Testing.md`: the harness and the results of every experiment run.

## Working on the format

- The spec describes the file format and how it changes across releases: the bytes, what they mean, and the rules
  that decide what Excel writes. What Excel does beyond the bytes goes in a document of its own, not the spec:
  display and drawing in `docs/Display.md`, calculation in `docs/Calculation.md`, SYLK export in `docs/Sylk.md`,
  and a new document for anything else of that kind. Those documents follow the spec's conventions and cite its
  sections as "spec §7".
- Work from Excel 1.00 first: decompile 1.00's p-code and run experiments in Excel 1.00, then describe what
  1.01, 1.03 and 1.5 changed. The only 1.01 is the French release (`Excel 1.01/` beside the other Excel
  folders); it keeps 1.00's procedure numbers except where it inserts procedures, so give its number only where
  it differs ("`46.06` in 1.00, `46.07` in 1.01").
- Excel 1.04 (English, from WinWorld) sits between 1.03 and 1.5. It inserts segments `0x47` and `0x48`, so 1.03's
  `0x47`–`0x73` are its `0x49`–`0x75`; compare it with 1.03. Excel 1.06 keeps 1.04's numbering and writes the same files. Cite code as `segment.procedure` with the release (`46.06` in 1.00), and say what the
  procedure is unless the sentence already does ("FREEZE.PANES, `0f.3e` in 1.5"). Where a fact was read from
  1.03 or 1.5, find the 1.00 counterpart (procedures keep their frame size and opcode sequence across releases,
  only the numbers move) and cite it first: "`47.0b` in 1.00, `45.0c` in 1.03".
- The spec states facts plainly. Mark only what is not established: **(inferred)**, **(assumed)** or
  **(unknown)**. Never label things "confirmed".
- List every set of values a field can take (chart types, tick marks, alignments, arrow heads and so on) as a
  value/meaning table under the record table, not inline in a sentence. A field that packs several flags or
  bit fields gets its own table too (a `Bits` column, highest first, or a `Flag` column for masks such as
  `0x04`), and its row in the record table says "(below)". Only a field holding a single bit field describes
  it in its row ("bits 15–12: weight").
- Keep every table in order: record tables by offset, bit tables highest bit first. An unused byte or bit gets
  its own row in its place ("Not used"), not a combined row at the end.
- Bytes Excel copies straight from memory are "left over from memory" (defined in the spec's Conventions); do not
  call them junk, garbage or stale.
- Write version differences from earlier to later ("1.03 adds…"), not the other way round.
- Every file Excel saves in an experiment that shows something goes into `tests/Samples/<release>/`, named for
  what it shows, with a test in `tests/ExperimentTests.cs`. The round-trip and recompile tests then cover it.
- A finding from the decompiled code that the samples can show (a field's value, a byte always 0) also gets a
  test over those samples, in `tests/ExperimentTests.cs` or, for a rule every sample follows, a theory over
  `TestData.AllSheets()` or `AllCharts()`.
- When a finding adds a field, expose it in the library and show it in the `dumper` output.
- A later release's code can explain an earlier release's bytes: Excel 1.5 widening a field leaves values that
  1.00 and 1.03 carry along. Check Excel 1.5's code when Excel 1.00 never writes a byte.
- Keep the spec clear and plain: no duplication between sections, the right level of detail, no filler.
  Update §18 (open questions) or the open questions of the document the finding is in, `docs/Testing.md`
  (results table) and the README when findings change them.

## Commands

```sh
dotnet build
dotnet test
dotnet run --project dumper -- --help
```
