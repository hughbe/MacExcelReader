# Checking files in a real Excel under Mini vMac

These scripts take sheets the library wrote into Excel 1.x running in [Mini vMac](https://www.gryphel.com/c/minivmac/), and bring Excel's saved
copies back out to compare. You need the emulator, a Mac ROM, a bootable System disk and an Excel disk image;
none are included.

| Script | Does |
|---|---|
| `mkmfs.py OUT.dsk SIZE_K VOLUME name=data[,rsrc]:TYPE:CRTR ...` | Builds a 400K (or 800K) MFS floppy image holding the given files. Macro sheets are `XLPG:XCEL`, worksheets `XLBN:XCEL`. |
| `mfs.py IMAGE.dsk OUTDIR` | Lists an MFS image and extracts every file (data fork, plus `.rsrc` for a resource fork). |
| `start.sh SYSTEM.dsk [DISK.dsk ...]` | Kills Mini vMac, starts it, boots the system disk, then mounts the others. `MINIVMAC` sets the app path. |
| `check_coverage.py EXCEL.csv OURS.csv ...` | Reads the Check macro's results and reports the rows that still differ after converting Excel's R1C1 text to A1. |

## Formula round trip

```bash
dotnet run --project dumper -- coverage out/R --split 100
```

This writes `out/R 01` … `out/R 06`, about 100 formulas each. Column A holds each formula as text, and column
B holds the real formula. Each sheet defines a command macro, `Check` (Option-Command-K), that compares
`GET.FORMULA` of column B with column A and lists the rows that differ: their row numbers in column G and
Excel's text in column C.

```bash
args=(); for i in 01 02 03 04 05 06; do args+=("R$i=out/R $i:XLPG:XCEL"); done
python3 tools/minivmac/mkmfs.py Run.dsk 400 Run "${args[@]}"
tools/minivmac/start.sh System.dsk "Microsoft Excel 1.03.dsk" Run.dsk
```

In the emulator:

1. Start Excel from its own disk, not by opening a document from the Finder. The Finder may pick a newer Excel.
2. Open each sheet.
3. Run Macro > Run > Check.
4. Note the count it reports, then save with Command-S.

Then read the results back:

```bash
python3 tools/minivmac/mfs.py Run.dsk saved
for i in 01 02 03 04 05 06; do
  dotnet run --project dumper -- export "saved/R$i" > "saved/R$i.csv"
  dotnet run --project dumper -- export "out/R $i" > "saved/ours$i.csv"
done
python3 tools/minivmac/check_coverage.py saved/R01.csv saved/ours01.csv ...   # each pair in turn
```

A clean run reports 0 unexplained rows and no lost rows. Excel's own saved copy is also useful to diff byte by
byte against the library's output (`dumper dump`, `dumper layout`).

[Excel bugs and quirks](../../docs/ExcelBugs.md) lists what Excel does wrong and what only looks wrong.
