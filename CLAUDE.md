# MacExcelReader

A .NET library that reads and writes Excel 1.x for Macintosh files, and `docs/Spec.md`, the format
specification it is built from. The library is unreleased, so its API can change.

## Layout

- `src/`: the library. `dumper/`: the `dumper` command-line tool, including `sheet` (build test sheets) and
  `excel-test` (the harness for running files in a real Excel). `tests/`: xUnit tests over `tests/Samples`.
- `docs/Spec.md`: the format. `docs/ExcelBugs.md`: Excel's own bugs and quirks. `docs/Testing.md`: the harness
  and the results of every experiment run.

## Working on the format

- Work from Excel 1.00 first: decompile 1.00's p-code and run experiments in Excel 1.00, then describe what 1.03
  and 1.5 changed. Cite code as `segment.procedure` with the release (`46.06` in 1.00).
- The spec states facts plainly. Mark only what is not established: **(inferred)**, **(assumed)** or
  **(unknown)**. Never label things "confirmed".
- List every set of values a field can take (chart types, tick marks, alignments, arrow heads and so on) as a
  value/meaning table under the record table, not inline in a sentence. A field that packs several flags or
  bit fields gets its own table too (a `Bits` column, highest first, or a `Flag` column for masks such as
  `0x04`), and its row in the record table says "(below)". Only a field holding a single bit field describes
  it in its row ("bits 15–12: weight").
- Bytes Excel copies straight from memory are "left over from memory" (defined in the spec's Conventions); do not
  call them junk, garbage or stale.
- Write version differences from earlier to later ("1.03 adds…"), not the other way round.
- Every file Excel saves in an experiment that shows something goes into `tests/Samples/<release>/`, named for
  what it shows, with a test in `tests/ExperimentTests.cs`. The round-trip and recompile tests then cover it.
- When a finding adds a field, expose it in the library and show it in the `dumper` output.
- Keep the spec clear and plain: no duplication between sections, the right level of detail, no filler.
  Update §18 (open questions), `docs/Testing.md` (results table) and the README when findings change them.

## Commands

```sh
dotnet build
dotnet test
dotnet run --project dumper -- --help
```
