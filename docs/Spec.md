# Microsoft Excel 1.x for Macintosh — document file format

This describes the files Microsoft Excel 1.00 to 1.5 for the Macintosh save: worksheets, macro sheets and
charts. It was reverse-engineered from the applications' code and checked against every sample file available.

## Sources

**Code.** The p-code and 68000 code in the applications' `PCOD` resources, disassembled:

| Release | Disk | Used for |
|---------|------|----------|
| Excel 1.00 (1985) | "Excel Program" disk (400K MFS), in two copies whose code differs in 4 bytes (the copy protection) | Save (`47.0b`), Open (`46.06`), the formula compiler (`11.00`), recalculation (`25.01`), the TABLE command (`4b`), chart writer, reader and drawing code (`5d`, `5e`, `37`), workspace files (`64.00`, `64.01`), function table and names. Each was compared with its 1.01 and 1.03 counterparts instruction by instruction; the differences are noted where they matter. |
| Excel 1.01 (French) | `Microsoft Excel 1.01 [French].dsk` (800K HFS), the only 1.01 available | Every procedure, matched to its 1.00 counterpart. 1.01 keeps 1.00's segment and procedure numbers apart from a few procedures it inserts; where a number differs the spec gives both (`46.06` in 1.00, `46.07` in 1.01). |
| Excel 1.03 | `Microsoft Excel 1.03.dsk` (800K HFS) | Save (`45.0c`), Open (`44.08`), the external-link reader (`72.01`), the formula compiler (`0f.00`), chart reader and writer (`5d`, `5c`), the counterparts of 1.00's `47.0b`, `46.06`, `73.01`, `11.00`, `5e` and `5d`. The spec cites 1.00's numbers first. |
| Excel 1.04 | `Microsoft Excel 1.04.dsk` (800K HFS, from WinWorld) | Every procedure, matched to its 1.03 counterpart. 1.04 inserts two segments (`0x47`, `0x48`), so 1.03's `0x47`–`0x73` are `0x49`–`0x75` in 1.04 and lower numbers are unchanged apart from a few inserted procedures; the spec gives 1.04's numbers where it describes a 1.04 change. |
| Excel 1.06 | `Microsoft Excel 1.06.dsk` (800K HFS) | Every procedure, matched to its 1.04 counterpart. 1.06 keeps 1.04's segment and procedure numbers apart from a few procedures it inserts or drops. |
| Excel 1.5 (May 1988) | Excel 1.5 folder | Save (`47.0d`), format and token translation (`47.18`), function table and file-number table, version-18 rules. Each 1.5 change the spec names was checked against 1.06's code, the release before it. |

Code is cited as `segment.procedure` in hexadecimal, for example `0f.00`.

**Samples.** 37 documents and one workspace file from Microsoft's disks, and 60 files Excel saved in
experiments. A reader built from this spec reads all of them, and a writer rebuilds them byte for byte.

| Source | Files |
|--------|-------|
| Excel 1.00 data disk | `Expenses`, `CLOCK.WS` (worksheets); `COMMANDS`, `DB.FORM`, `REPLACE`, `MATRIX`, `CLOCK.MS` (macro sheets); `CLOCK.CH` (chart). Both copies of the disk hold the same bytes. |
| Excel 1.00 program disk | `Resume Excel`, a workspace file (Appendix B) |
| Excel 1.03 disk | `Worksheet1` (empty worksheet). The disk's other samples are byte-identical to the 1.00 disk's. |
| Excel 1.04 disk | `Expenses`, `COMMANDS`, `DB.FORM` and `REPLACE`, byte-identical to the 1.00 disk's, and a `Resume Excel` with no windows. |
| Excel 1.06 disk | `COMMANDS`, `DB.FORM` and `REPLACE`, byte-identical to the 1.00 disk's. |
| Excel 1.5 sampler | 26 worksheets and macro sheets, including one version-18 file (`SIMUL EQUATIONS`), and 2 charts (`BREAK CHART`, `PROFIT CHART`) |
| Experiments | Files Excel 1.00, 1.01, 1.03, 1.04, 1.06 and 1.5 saved in the test harness (docs/Testing.md), each named for what it shows, such as `Literals`, `Errors`, `Column Widths`, `Table Deleted Input`, `Arrow Heads` and `Legend`. They are in `tests/Samples`, next to the release that saved them. |

Microsoft's samples hold 1,282 formulas and 206 name definitions. §2.3 explains how to tell which build saved
each file.

Everything else is read from code, seen in the samples, or shown by Excel itself (docs/Testing.md). The rest
is marked:

* **(inferred)**: a conclusion from indirect evidence, such as dialog item order or consistency with later
  formats;
* **(assumed)**: a rule chosen because nothing contradicts it, with no evidence for it;
* **(unknown)**: not identified.

§18 lists everything that is still open. [Excel bugs and quirks](ExcelBugs.md) lists how Excel itself
misbehaves, for example with a file that breaks a rule given here.

---

## Contents

1. [Conventions](#1-conventions)
2. [Versions](#2-versions)
3. [Overall layout](#3-overall-layout)
4. [File header (24 bytes)](#4-file-header-24-bytes)
5. [Data area preamble](#5-data-area-preamble)
6. [Sheet state block (630 bytes)](#6-sheet-state-block-630-bytes)
7. [Section 0 — number formats](#7-section-0--number-formats)
8. [Section 1 — external documents and row index](#8-section-1--external-documents-and-row-index)
9. [Section 2 — cell table](#9-section-2--cell-table)
10. [Section 3 — external names](#10-section-3--external-names)
11. [Section 4 — defined names](#11-section-4--defined-names)
12. [Section 5 — cell data](#12-section-5--cell-data)
13. [Section 6 — windows](#13-section-6--windows)
14. [Section 7 — page breaks and print settings](#14-section-7--page-breaks-and-print-settings)
15. [Parsed expressions](#15-parsed-expressions)
16. [Chart files](#16-chart-files)
17. [Recipes](#17-recipes)
18. [Open questions](#18-open-questions)
* [Appendix A. Functions](#appendix-a-functions)
* [Appendix B. Workspace files](#appendix-b-workspace-files)

---

## 1. Conventions

* All multi-byte integers are **big-endian**.
* `byte` = 8 bits, `word` = 16 bits, `long` = 32 bits, `double` = IEEE 754 64-bit.
* *Pascal string* = one length byte followed by that many characters (Mac Roman).
* *Rectangle* = four words, top, left, bottom, right: the Macintosh Toolbox's (QuickDraw's) `Rect`, in pixels
  unless a field says otherwise.
* Bits are numbered from 0, the least significant: bit 15 is the top bit of a word. A field that packs several
  flags or bit fields has its own table under the record table: a **Bits** column (highest first) where the
  spec numbers bits, or a **Flag** column where Excel's code tests masks (`0x04`). A field holding one bit
  field says so in its row: "bits 15–12: weight".
* Rows and columns are **zero-based** on disk. Row limit 16,384 (14 bits), column limit 256.
* The whole document lives in the **data fork**. The resource fork is empty.
* **Left over from memory.** Excel saves many structures by copying them straight out of memory, so some bytes
  in a file are whatever they held while Excel was running, often a pointer into Excel's memory. They mean
  nothing on disk: Open ignores them or rebuilds them, and a writer should write zeros.

### 1.1 Finder types

| Type   | Creator | Meaning                         |
|--------|---------|---------------------------------|
| `XLBN` | `XCEL`  | Worksheet                       |
| `XLPG` | `XCEL`  | Macro sheet                     |
| `XLIN` | `XCEL`  | Workspace file ("Resume Excel"): the settings and open windows Excel restores. Not a document; see Appendix B. |
| `MCBN` |         | Chart, saved in Microsoft Chart format (§16) |

Save dispatches on the document kind: charts go to the chart writer; otherwise the file type is `XLPG` when
the document kind byte (§5) is `0x40`, else `XLBN`.

### 1.2 Stream positions

The writer is a 512-byte buffered stream. A stream position is stored as a long made of two words, **(block
number, offset within block)**:

```
absolute offset = hi_word * 512 + lo_word
```

e.g. `00 02 01 F0` → 2 × 512 + 0x1F0 = 0x5F0.

The offset runs from 1 to 512, not 0 to 511: the buffer is written out only when the next byte arrives, so a
position on a block boundary is stored as the end of the block before it. Section 2 of Excel 1.5's `Frozen
Panes` starts at 0x600 and is stored as `00 02 02 00`, not `00 03 00 00`. Only position 0 has offset 0.

### 1.3 Packed pointers ("seg:off")

Excel 1.x's p-code uses 8086-style 20-bit pointers: `address = seg * 16 + off` with `off` in 0…15 once
normalised. Several structures were dumped from memory with such pointers still in them. They are left over
from memory (above), except in the cell table, where the writer replaces them with section-relative offsets
(§9).

---

## 2. Versions

### 2.1 Releases

| Release | Format versions it writes | Functions | What changed |
|---------|--------------------------|-----------|--------------|
| 1.00 (September 1985) | 17 | 0–207 | — |
| 1.01 | 17 | 0–207 | Adds a one-byte extra header (§5) and drops the extra function names (§14), keeping the function-count flag in state word `0x01C` instead (§5, §6). Takes the argument separator and other characters from the system's international settings (§15.4), and moves its words into resources so that it can be translated: function names (§15.3), the built-in names of a new sheet (§11), built-in number formats and format letters (§7), and the workspace file's name (Appendix B). Its code asks for a string by number (`0d.0d` in 1.01): below `0x2000` the `STR` resource with that ID, from `0x2000` up item number − `0x2000`, counting from 0, of `STR#` 256. Function table and tokens identical to 1.00's. |
| 1.03 | 17 | 0–207 | Writes a zero-length name for an external entry that refers to the document itself (§8), and makes its in-memory external entries 16 bytes (§8). Function table and tokens identical to 1.01's. |
| 1.04 | 17 | 0–208 | Adds CHANGE.LINK (208, a command with up to two text arguments, carried out by `44.00` in 1.04) and a third argument to OPEN, matching the Read Only box its Open dialog gains (Appendix A). Its Save writes the function count as one more than the highest function number a formula uses (§5): 209 for a sheet calling CHANGE.LINK, and 208 for one that calls nothing newer even when the file it opened said 209 (`Changes 1.04`, `No New Functions 1.04` samples). It keeps a document's file reference between saves and flushes the volume after saving (`44.0c`, `44.0d` in 1.04), which does not change the bytes written. Embedded code identical to 1.03's apart from the import word (§2.3). Compiler and tokens identical to 1.03's. |
| 1.06 (December 1987) | 17 | 0–208 | Writes the same files as 1.04, with byte-identical embedded code, so its files cannot be told from 1.04's (§2.3); compiler, function table and tokens unchanged (`Changes 1.06` and `No New Functions 1.06` samples). Its other changes are inside the application: a `SIZE` resource for MultiFinder, one routine (`01.1a` in 1.06) for all its checks for pending events, and no longer reserving master pointers around Open (1.04's `15.2b`, `15.2c`). |
| 1.5 (May 1988) | 17, or 18 when needed | 0–291 | Adds functions 209–291 (tokens `0x39`, `0x3A` for 256 and up, §15.2). Changes the argument rules of eleven more functions (Appendix A; OPEN's had changed in 1.04) and the token class of function calls inside IF and CHOOSE (§15.4). Marks charts with more than 102 categories as chart version 1, which earlier releases refuse (§16.2). Adds frozen panes and gridline colours to windows (§13) and text sections, scaling by trailing commas, `@` and colours to number formats (§7). |

The format version is header word 2 (§4). Excel 1.00's Open accepts only version 17 (`46.06`), and so do
1.01's (`46.07`), 1.03's, 1.04's and 1.06's (`44.0e` in 1.04 and 1.06). Excel 1.5 writes 17 unless the
document needs something only it can read, then 18 and the warning "This file will only be readable by Excel
1.5 or greater" (`47.0d`). Save starts at 17 and switches to 18 when one of the document's windows has frozen
panes or a gridline and heading colour other than automatic (state bytes 121 and 127, §13), or when the
number-format check (`47.10` in 1.5) flags a user-defined format. The check turns each format back into text
(`0d.19`, with `0d.1a` for each section), and the decompiler raises a flag when the format has any of:

| Feature | Example | Where the flag is set |
|---------|---------|-----------------------|
| A text section that is not General and is not the positive section's bytes | `0;-0;0;"t"` | `0d.19` |
| A trailing thousands separator, which scales the number by 1000 (number header `c` below `b`, §7) | `#,##0,` | `0d.1a`, token `03` |
| `@` (token `23`) | `@` | `0d.1a` |
| A colour (token `24`) | `[RED]0.00` | `0d.1a` |

A format using only what earlier releases read saves as 17 (`Format Two Sections`). `SIMUL EQUATIONS`, the one
version-18 file among Microsoft's samples, has none of these: the released 1.5 opens it and saves it as version
17 without the warning (`SIMUL EQUATIONS Saved`). Its embedded code is the released build's (§2.3), so it was
saved by a build with the same segment 1 but a different version rule, or in a state its file does not keep
**(inferred)**.

A version-17 document can also need a later release without changing its version, because it uses a function
the reader lacks. The function count at the start of the data area (§5) records that, and Excel 1.5 warns
"readable by Excel 1.04 or greater" for function 208 and "1.5 or greater" for 209 and up.

### 2.2 Function numbers across releases

Function numbers in files are stable: 0–207 mean the same in every release, and later releases only add
numbers. Excel 1.5 renumbered its functions internally (into the order BIFF later used: PV is 56 internally
but 58 in files), and translates in both directions on Save and Open with a sorted table of 279 (internal
number, file number) pairs. Appendix A lists every file number.

### 2.3 Identifying the build that saved a file

Each document embeds a copy of part of the saving build's code (§4.1), so the code fingerprints the build. The
segment's import word, which names the host procedure in the saving build's own segment numbering, tells the
builds apart even where the lengths agree:

| Segment 1 length | Import | Matches | Samples |
|-----------------:|-------:|---------|---------|
| `0x038A` | `0x720A` | A build earlier than 1.00, without the p-code compiler's shared-sequence optimisation | `Expenses` |
| `0x0326` | `0x730A` | Excel 1.00 or 1.01: byte-identical to segment `0x70` of both. A 1.01 file has the one-byte extra header (§5), a 1.00 file none. | the other six sheets of the 1.00 disk |
| `0x0326` | `0x720A` | Excel 1.03: byte-identical to its segment `0x6F` | `Worksheet1` |
| `0x0326` | `0x740A` | Excel 1.04 or 1.06: segment `0x71` of both, byte-identical, which differs from 1.03's only in the import word and one byte near the end | `Changes 1.04`, `No New Functions 1.04`, `Changes 1.06`, `No New Functions 1.06` |
| `0x032A` | `0x7A0A` | Excel 1.5: byte-identical to the 1.5 application's code | ten of the 1.5 samples |
| `0x03FC` | `0x7C0A` | Another Excel 1.5 build, encoding formulas exactly as the 1.5 application's code does | `ACCOUNTS MACROS`, `COMMISSION MACROS` |
| `0x032A` | `0x7C0A` | A pre-release Excel 1.5 build (different bytes) | the other 14 of the 1.5 samples |

Charts cannot be dated this way: their embedded code (§16.2) is the same in every release. `CLOCK.CH` is
format version 3, which no released build writes (§16.2), so it comes from a build before 1.00.

The pre-release build differs from the released one in one respect the samples show: it writes FORMULA with
one argument in the variable-count form (`42 01 39`), where the released 1.5 keeps the fixed-count form (`41
39`, §15.4). It treats ACTIVATE like the released 1.5.

---

## 3. Overall layout

```
+0x00  File header                     24 bytes
+0x18  Embedded p-code segment 1       header.len1 bytes
       Embedded p-code segment 2       header.len2 bytes
       Function count                  2 bytes          <- header.dataOffset
       Sheet state block               630 bytes
       Document flags                  1 byte
       Section table                   8 longs (32 bytes)
       Extra-header length             word
       Extra header                    that many bytes
       Section 0  Number formats
       Section 1  External documents + row index
       Section 2  Cell table
       Section 3  External names
       Section 4  Defined names
       Section 5  Cell data
       Section 6  Windows
       Section 7  Page breaks and print settings
```

Sections are contiguous and in this order; the section table gives the start of each.

---

## 4. File header (24 bytes)

| Offset | Size | Value in samples | Meaning |
|-------:|-----:|------------------|---------|
| 0  | word | `0x5327` | Magic. Open rejects anything else. |
| 2  | word | 17, or 18 | Format version (§2.1). Open rejects anything but 17 before Excel 1.5. |
| 4  | word | `0xAB27` | Code signature of the writing build (see §4.1). |
| 6  | 3 words | 0, 0, 0 | Written as zero; never read. |
| 12 | word | 2 | Number of embedded code segments. Always 2; Open does not read it (see below). |
| 14 | word | `0x0326`, `0x032A`, … | Length of segment 1; identifies the saving build (§2.3). |
| 16 | word | `0x0018` | Offset of segment 1 (always 24). |
| 18 | word | `0x0016` | Length of segment 2. |
| 20 | word | 24 + len1 | Offset of segment 2. |
| 22 | word | 24 + len1 + len2 | **Offset of the data area.** |

Open checks the magic and the version, then seeks to offset 22 and reads the data offset from there. It
ignores the segment count and the segment table. So the header is always 24 bytes. A writer that has no code
to embed must still write two segment entries (Excel's own layout, with lengths of 0 and offsets of 24), so
that the data offset lands at offset 22.

### 4.1 Embedded p-code segments

Each Excel 1.x document carries two segments of the application's own interpreted code, copied verbatim out of
memory: segments `0x70` and `0x71` in Excel 1.00, which 1.03 renumbers `0x6F` and `0x70`.

They exist for **external references to closed documents**. When Excel needs a value from a file it has not
opened, the link reader (`73.01` in 1.00) looks at header word 4:

* `0xAB27` (same build) → use the application's own copy of the segment (`0x70` in 1.00).
* `0x13E8`, or any other `0xABxx` → load segment 1 *from the file*, patch its single import to point back at
  the host (segment `0x73`, procedure `0x0A` in 1.00; `0x72` in 1.03), and call its procedure 0 with message
  25 to fetch the value.

So a document is readable for linking by Excel builds that do not otherwise know its exact layout. A reader
that only wants the data can **skip both segments**.

Segment image layout:

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 | word | Offset of the constant/import table end (`2 × nProcs + 8`: exactly one import word precedes it) |
| 2 | word | The p-code interpreter's least-recently-used stamp for the segment, which it uses to choose which code segments to unload when memory runs short. Left over from memory. |
| 4 | word | Number of procedures (3 in segment 1, 1 in segment 2) |
| 6 | word × n | Offset of each procedure |
| … | word | Import: the host procedure, in the saving build's numbering: `0x730A` (segment `0x73`, procedure `0x0A`) in 1.00, `0x720A` in 1.03 (§2.3) |
| … | | p-code |
| last 10 bytes | | Header of whatever heap block followed in memory: left over from memory |

The stored length is the heap block size, which includes a 10-byte block header that is *not* written, so each
image over-runs 10 bytes into the next block.

The segment is a copy of the saving build's own code, so its length and bytes identify that build (§2.3). The
0x38A-byte segment in `Expenses` holds the same three procedures as 1.00's 0x326-byte one, with the same
logic, but without the p-code compiler's shared-sequence opcodes that fold repeated instruction runs. So
`Expenses` was saved by an earlier build compiled without that size optimisation, which nevertheless used the
same code signature `0xAB27`. Chart files carry signature `0x13E8` (§16.2).

---

## 5. Data area preamble

At `header.dataOffset`:

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 | word | **Function count**: one more than the highest function number the document may use (below) |
| 2 | 630 bytes | Sheet state block (§6) |
| 632 | byte | **Document kind** (below) |
| 633 | 8 longs | Section table: stream positions (§1.2) of sections 0…7. Written last, by seeking back. Open *skips* this table and reads sequentially; it is redundant but reliable. |
| 665 | word | `n` = number of extra header bytes that follow: 0 from Excel 1.00, 1 from 1.01 and later. Excel 1.00's Open reads `n` and skips that many bytes, so it reads later files' extra header without trouble. |
| 667 | n bytes | Extra header. Byte 0, when present, is a system/locale code (high byte of an application global; 0 in every sample, and in the French 1.01's own data); from 1.01 on, Open warns with alert 0x106 if it differs from the reader's (`46.07` in 1.01). Open skips any further bytes, so this is the format's forward-compatibility hook. |

Document kind:

| Value | Kind |
|---:|---|
| `0x10` | Worksheet |
| `0x40` | Macro sheet |
| `0x20` | Chart: the in-memory kind only, since charts never reach this writer (§16) |

Section 0 begins immediately after the extra header.

**Function count.** Open refuses a file whose count is larger than its function table (208 entries in
1.00–1.03), since its formulas could use numbers the reader lacks: Excel 1.00 gives "File is not a saved
worksheet" (alert `0x0159`, `46.06`). Excel 1.00 also has a base count fixed at start-up, and its Save writes
the table length if a global flag is set, and otherwise the base count; the formula compiler sets the flag
when a formula uses a number at or above the base count, and Open sets it after reading a file with extra
functions. Between the base count and the count, 1.00's files hold the names of the extra functions at the
start of section 7 (§14). In released 1.00 the base count and the table length are both 208, so it always
writes 208 and never a name. Excel 1.01 keeps the flag per document, in state word `0x01C` (§6), and drops
the names; 1.03 does the same. Excel 1.04's Save (`45.0d` in 1.04 and 1.06) and 1.5's write 208, or one more
than the highest function number any formula uses if that is higher. Every sample has 208 except
`INVOICE MACROS` (GET.DOCUMENT, 238, so 239), `Changes 1.04` and `Changes 1.06` (CHANGE.LINK, 208, so 209) and
`Functions 256 And Up` (292).

---

## 6. Sheet state block (630 bytes)

A raw dump of a contiguous run of application globals. Offsets are relative to the start of the block.

| Offset | Size | Samples | Meaning |
|-------:|-----:|---------|---------|
| 0x000 | word | 100 | Maximum iterations (Calculation dialog; 1…32767) |
| 0x002 | word | 1 | Calculation mode (values below; `CALCULATION(2)` saves −1, `CALCULATION(3)` saves 0). Like 0x006 it is a global setting: every sheet Excel saves gets the current one. |
| 0x004 | word | 1 | Precision: 1 = full precision, 0 = precision as displayed (results are rounded through the cell format when 0) |
| 0x006 | word | 1 | Reference style: 1 = A1, 0 = R1C1 (`A1.R1C1(FALSE)` saves 0). It is saved and restored around formula parsing, and is global like 0x002. |
| 0x008 | word | 0 | Unused: no code in the application reads or writes it |
| 0x00A | double | 0.001 | Maximum change (iteration) |
| 0x012 | word | 0 | Iteration enabled (0 / 1) |
| 0x014 | word | | **Last row** of the grid; see §6.1 |
| 0x016 | word | 255 | Index of the last column whose width is not zero. Starts at 255; the column-width routine lowers it when the last such column is set to width 0 and raises it when a higher column is given a width. (A column width of 0 is how a column is hidden.) |
| 0x018 | word | 0 / 1 | Recalculation state: 0 clean, 1 needs recalculation, 2 a pending state the code never lowers to 1. Open loads it as it is and the recalculation driver (`25.0e`) acts on it, so Excel recalculates a file saved with 1 or 2 as soon as it opens it (`Recalculated` sample). With 0 it keeps the stored results, even after CALCULATE.NOW, until something they depend on changes. The external-link reader (`73.0d` in 1.00, `72.0d` in 1.03) checks it before trusting a closed file's stored values; when it marks the sheet for recalculation, 1.01 and later leave a 2 as it is where 1.00 sets 1. |
| 0x01A | word | 0 | Pending-work flag. Excel 1.00 never sets it: a new sheet starts with 0 (`10.00`), so only a file can. When it is set, `22.03` runs the range routine `23.01` over the whole sheet (rows 0–16383, columns 0–255) with operation code `0xD01F` and clears the word. Operation 31 does nothing to a cell; the flag bits make the routine rescan each row (`23.02`), checking for each cell whether its neighbours are empty and whether they are Fill-aligned: what Excel needs to draw text that spills into the next cells. Excel 1.00 works that out as it draws anyway, so the word has no visible effect: sheets with overflowing text and with runs of Fill-aligned cells look the same saved with 0 or 1, and Excel saves them back with 0. |
| 0x01C | word | 0 | Set to 1 by the formula compiler when a formula uses a function at or above the base function count. It selects which count Save writes at the start of the data area (§5). Cleared for a new sheet. |
| 0x01E | word | 0 / 1 | **Document protected.** Selects the "Protect Document…" / "Unprotect Document…" menu text and is checked by ~50 editing routines. 1 only in `DB.FORM`. |
| 0x020 | word | 0 / 1 | Which form 0x014 and 0x274 are in: 1 = last indices, 0 = counts (§6.1) |
| 0x022 | word | | Number of defined names that have a definition (0 for a new sheet, 6 in `REPLACE`). Define Name counts it up; Open and Save carry it through unchanged, so a file written with 0 keeps 0 when Excel re-saves it. Every sample has the true count. |
| 0x024 | word | | Number of names whose definition is a reference (flag `0x04`, §11; 5 in `REPLACE`). Excel recomputes it when it saves. |
| 0x026 | long | | Pointer, left over from memory |
| 0x02A | long | | Head of the in-memory name list, left over from memory |
| 0x02E | 4 longs | | Pointers and `FFFF` markers, left over from memory |
| 0x03E | word | 0 / 1 / 2 | 1-based position of the active-sheet external entry (the one named `\0`, §8), or 0 if there is none. Link refresh skips that entry. Excel 1.00 and 1.03 recompute it when they save. |
| 0x040 | word | 1 | Window counter for the document: set to 1 for a new sheet and updated as windows are opened and closed |
| 0x042 | 256 words | 75 | **Column widths** in pixels, columns 0…255. Default 75. COLUMN.WIDTH(*w*) stores *w* × the width of `0` in the standard font (`0x254`, 7) + 5, rounded: 1 → 12, 2.5 → 23, 10 → 75, 20 → 145; 0 hides the column (`Column Widths` sample). |
| 0x242 | word | 0 | **Protection password hash**; 0 = none. Unprotect hashes the typed password and refuses on mismatch. See §6.2. |
| 0x244 | word | 3 | Standard font number (passed to `TextFont`; 3 = Geneva) |
| 0x246 | word | 10 | Standard font size (passed to `TextSize`) |
| 0x248 | word | 5 | Cached: half the font ascent |
| 0x24A | word | 12 | Cached: ascent + descent |
| 0x24C | word | 13 | Cached: default row height in pixels |
| 0x24E | word | 10 | Cached: font ascent |
| 0x250 | word | | Cached: packed font metric |
| 0x252 | word | 8 | Cached: width of `0` in bold |
| 0x254 | word | 7 | Cached: width of `0` in plain |
| 0x256 | long | | Pointer, left over from memory |
| 0x25A | long | 0 | Head of the row page-break list in memory (rebuilt from §14) |
| 0x25E | long | 0 | Head of the column page-break list in memory (rebuilt from §14) |
| 0x262 | word | 0 | Cached: 1 if a print area is defined |
| 0x264 | word | 0 | Cached: pixel width of the print-title columns |
| 0x266 | long | | Pointer, left over from memory |
| 0x26A | word | 0 | Cached: number of print-title rows |
| 0x26C | word | 0 | Cached print-titles value |
| 0x26E | 6 bytes | 0 | Cached print-titles range (first row word, last row word, first column byte, last column byte) |
| 0x274 | word | | **Last column** of the grid; see §6.1 |

Calculation mode (`0x002`):

| Value | Meaning |
|---:|---|
| 1 | Automatic |
| −1 | Automatic except tables |
| 0 | Manual |

Offsets 0x262–0x273 are all derived from the `Print_Area` and `Print_Titles` names by one routine, which is
re-run whenever a name or a column width changes; a writer can leave them zero.

The cached font values at 0x248–0x254 are recomputed from the font number and size by a single routine
(`GetFontInfo` / `CharWidth`); a writer can copy them from a sample.

### 6.1 Grid dimensions

```
nRows = word[0x014] + word[0x020]
nCols = word[0x274] + word[0x020]
```

These size the row index (§8) and the cell table (§9). Excel itself uses only rows `0 … word[0x014] − 1` and
columns `0 … word[0x274] − 1`, so in the last-index form (0x020 = 1) the table's last row and column are a
**blank margin**. Every sample in that form has no values there, and a writer using it must leave one blank
row and column after the last value: Excel 1.03 loses, misdraws or hangs on values in the margin ([Excel bugs
and quirks](ExcelBugs.md#files-that-break-the-cell-table-margin-excel-103)). A new sheet starts with all three
words 0, so 0x014 and 0x274 hold the row and column counts. The first time the grid grows, the grid-sizing
routine (`22.08`) converts them: it sets 0x020 to 1 and subtracts 1 from both, so they hold the last row and
column index from then on. Most samples are in the second form; `CLOCK.WS`, `CLOCK.MS` (1.00) and `LOAN1`
(1.5) are in the first. An empty sheet has `nRows = nCols = 0`.

### 6.2 Password hash

From the password dialog's handler (`24.01` in 1.00, `22.01` in 1.03). For a password of `n` characters
`c[1…n]`:

```
h = 0
if n > 0:
    for b in c[n], c[n-1], …, c[1], n:        # characters last to first, then the length
        h = ((h << 1) & 0x7FFF) | ((h >> 14) & 1)   # rotate left within 15 bits
        h = h ^ b
    h = h ^ 0xCE4B
```

An empty password gives 0. This is the same hash later Excel versions use for sheet protection.

---

## 7. Section 0 — number formats

| Size | Meaning |
|-----:|---------|
| byte | Number of built-in formats in the writing build (19 in all samples). If this differs from the reader's, Open renumbers cell format indices. |
| byte | `n` = number of user-defined formats |
| n × … | User-defined format entries |

User-defined formats are numbered from 19, in this order; a cell's format index (§9) refers to them. Twenty of
the Excel 1.5 samples have between one and five.

Each entry is a descriptor and a tokenised body. A format has a section for negative numbers, one for zero and
one for positive numbers, in that order; version-18 files add a fourth, for text:

| Size (v17) | Size (v18) | Meaning |
|-----------:|-----------:|---------|
| 3 bytes | 4 bytes | Offset of each section within the body: negative, zero, positive (, text) |
| 3 bytes | 4 bytes | Length of each section |
| byte | byte | 1 when the format has a single section (every slot is the same bytes at offset 0), else 0 |
| n bytes | n bytes | Body, `n = max(offset[i] + length[i])` |

Sections can share bytes: `$#,##0 ;($#,##0)` (built-in 5) has offsets `0C 00 00` and lengths `0E 0C 0C`, so
zero and positive both use the 12 bytes at offset 0 and negative the 14 at offset 12. A section of length 0
displays nothing (`INVOICE LIST`: `$#,##0.00 ;($#,##0.00);` hides zeros). A format given without a text section
gets `01` (General) there, as in Microsoft's version-18 sample; `0;-0;0;"t"` stores `09 01 74` (`Format Text
Section`). Excel 1.5 writes the 3-slot form to version-17 files (`47.11`).

Built-in formats are not stored in documents; they live in the application's `FRMT` 256 resource in the same
descriptor-plus-body form. Total formats are capped at 64. The 19 built-ins, by index, in Excel 1.00. A file
stores only the index, so a translated Excel shows its own built-ins: the French 1.01's 10 and 18 put the day
first (`d/m/yy`), and its currency formats 5 and 6 put the symbol after the number and show negatives with a
minus sign. From 1.01 the format compiler also takes the section separator and the format letters from a table
of characters (`07.09` and its new helper `07.26` in 1.01), so a translated Excel reads and shows its own
letters for the same tokens.

| # | Format | # | Format |
|--:|--------|--:|--------|
| 0 | `General` | 10 | `m/d/yy` |
| 1 | `0` | 11 | `d-mmm-yy` |
| 2 | `0.00` | 12 | `d-mmm` |
| 3 | `#,##0` | 13 | `mmm-yy` |
| 4 | `#,##0.00` | 14 | `h:mm AM/PM` |
| 5 | `$#,##0 ;($#,##0)` | 15 | `h:mm:ss AM/PM` |
| 6 | `$#,##0.00 ;($#,##0.00)` | 16 | `h:mm` |
| 7 | `0%` | 17 | `h:mm:ss` |
| 8 | `0.00%` | 18 | `m/d/yy h:mm` |
| 9 | `0.00E+00` | | |

The format names above are the standard Excel 1.x list (they are not stored as text anywhere). The token
meanings below come from the routine that turns a tokenised format back into text for the Format Number dialog
(`07.17` in 1.00, `05.17` in 1.03). They agree with every built-in body and decode all 51 user formats in the
samples, from `00-00-000"E"` to `0.000 "mgal"` and `mmmm-yy`. Letters and separators are taken from an
international character table at run time, so the ones shown are the US defaults.

| Token | Operands | Text produced |
|------:|----------|---------------|
| `01` | | `General` |
| `02` | 1 byte `c` | Literal character `c`. A backslash is written before it unless it is one of the characters that may appear unescaped. |
| `03` | 3 bytes `a b c` | Starts a number section; produces no text. `a` = number of integer digit positions (drives thousands-separator placement), `b` = decimal places, `c` a signed byte: `(c − b) / 2` = number of percent signs. Excel 1.5: when `c` is below `b`, `(b − c) / 3` is the number of trailing thousands separators, each dividing the number by 1000 (`#,##0,` is `03 04 00 FD`, `Format Scaling`). |
| `1C` | 3 bytes | Same, for a scientific section |
| `04` | 1 byte `n` | `n` × `#` |
| `06` | 1 byte `n` | `n` × `0` |
| `05` | | Decimal point; digit runs after it are written without separators |
| `07` | | Turns on the thousands separator, inserted every three integer positions |
| `08`, `1B` | | Reserved. The format compiler (`07.09` in 1.00, `05.09` in 1.03) never emits them and the routines that read formats skip them. |
| `09` | 1 byte `n`, `n` chars | Quoted string `"…"` |
| `0A` | 1 byte `c` | `*c`: fill the cell with `c`. Excel 1.00 compiles `*-0` to `03 01 00 00 0A 2D 06 01`, and `\-0` to `… 02 2D …` (`Fill Format`). The `*` comes from the international table. |
| `0B`–`0E` | | `d`, `dd`, `ddd`, `dddd` |
| `0F`–`12` | | `m`, `mm`, `mmm`, `mmmm` (month) |
| `13`, `14` | | `h`, `hh` (24-hour) |
| `1F`, `20` | | `h`, `hh` (12-hour; used when the section has an AM/PM token) |
| `15`, `16` | | `m`, `mm` (minutes) |
| `17`, `18` | | `s`, `ss` |
| `19`, `1A` | | `yy`, `yyyy` |
| `1D`, `1E` | 1 byte | Exponent sign `+` / `-`; the byte is the exponent digit count |
| `21` | 2 bytes `n1 n2`, `n1` chars, `n2` chars | AM/PM designator, written as the two strings joined by `/` (`AM/PM`, `A/P`) |
| `22` | | Currency symbol |
| `23` | | Excel 1.5: `@`, the text placeholder. The character is the 28th of the format-character string (below). |
| `24` | 1 byte `n` | Excel 1.5: a colour, written `[name]` with the name `n` from `STR#` 263 (`black`, `white`, `red`, `green`, `blue`, `yellow`, `magenta`, `cyan`, from 0). For example, `[RED]0.00` holds `24 02`, and `[BLUE]#,##0;[RED]-#,##0` holds `24 04` and `24 02`. |

A format that uses `23` or `24`, a text section or trailing-comma scaling makes Excel 1.5 save the file as
version 18 (§2.1).

Excel 1.5 reads its format characters from `STR` 479, `;#0.%,:/ -+().\*ymdhmsE/"[]@`, at startup (`01.0b`
loads it, or the identical `STR` 480 when called with 1). It then replaces the decimal point (character 4),
thousands separator (6), time separator (7) and date separator (8) with those of `INTL` 0. The decompiler takes
`;` from character 1, the fill `*` from 16, the brackets of a colour from 26 and 27, and `@` from 28.

So `#,##0` is `03 04 00 00 | 07 | 04 03 | 06 01`: header, thousands on, `###`, `0`. Literals and strings sit
wherever they print: `0.0" S"` is `03 01 01 01 | 06 01 | 05 | 06 01 | 09 02 " S"`, and `### ##0.00` uses `02
20` for the space.

The compiler shows what the header bytes are: it writes the integer-position count, the decimal count, and
decimals plus twice the number of percent signs. It also shows how the hour tokens are chosen: `h`/`hh`
compile to `13`/`14`, and are bumped by 12 to `1F`/`20` when an AM/PM designator follows. And it shows that
`m` is compiled as minutes rather than month when it directly follows an hour token.

---

## 8. Section 1 — external documents and row index

**External document table**

| Size | Meaning |
|-----:|---------|
| word | 0 if the sheet has no external-document table. Otherwise `n × 12 + 8`: the size of Excel 1.00's in-memory table, an 8-byte header and 12-byte entries, whose first 4 bytes 1.00 writes as they are. Excel 1.03 makes the entries 16 bytes but still computes and writes the 1.00 value. |
| word | `n` = number of external documents (present only if the previous word is non-zero) |
| n × … | One entry per external document |

Entry:

| Size | Meaning |
|-----:|---------|
| 6 bytes | Not used; 0. These are the first 6 bytes of Excel 1.00's 12-byte in-memory entry. Excel clears them when it makes an entry (`16.00` in 1.00), Open reads them into the entry (`16.09`) and Save writes them back (`47.0b`), and nothing else reads or writes them. The rest of the entry is a pointer to the document's name (offset 6), which Save writes out as the string below, and a pointer into the list of formulas that use the link (offset 9), which Open rebuilds. |
| Pascal string | Document name: usually a file name (`CASHBUDGET`, `GREAT BRITAIN`), but it can be a full path (`::sampler files - current 1/13:Files:DATA`). Excel 1.00's Save always writes the entry's name; Excel 1.03's writes a zero-length name for an entry that refers to this document itself. An entry named with the single character `0x00` stands for whichever sheet is active: `DB.FORM`'s macros use it as `!DATABASE` and `!FORM`, Excel 1.03 makes one for `=!A1` (`Links` sample), and the test harness's `GET.FORMULA(!B4)` reads B4 of the active document. |

An external reference names its document by the entry's 1-based position in this table (§15.2, token `1A`).

**Row index.** Immediately after, `nRows` words, one per row. Mostly zero; 35 rows across the samples have a
small non-zero value.

In memory each row has a 4-byte slot that may point at a per-row auxiliary table. Save writes that table's
entry count (low 15 bits of its first word), or 0 if the row has none. Open, for a value greater than 2,
allocates an *empty* table with that capacity (`2n + 2` bytes) and does not read any contents. So the word is
only a size hint for a structure Excel rebuilds at run time; a writer can emit 0 and a reader can ignore it.

The structure is the recalculation dependency index. A row's slot records which rows hold formulas that refer
to cells in this row; its top two bits tag the form: 1 for a single row held inline, 2 for an inline pair, 3
for a pointer to a table (count word, capacity word, then row numbers whose own top two bits are flags). Open
rebuilds it by calling the dependency routine once per formula as the cell data is read. The transient bits
12–6 of cell word 2 (§9) belong to the same machinery: the routine that walks dependents tests bits 12, 11,
10, 9 and 8 of each cell as it goes.

---

## 9. Section 2 — cell table

A dense, row-major grid: `nRows × nCols` entries of 6 bytes. The entry for (row, col) is at index `row × nCols
+ col`. Row and column are therefore implicit.

| Field | Meaning |
|-------|---------|
| word 0 | Value or offset, depending on type |
| word 1, bits 15–12 | Low 4 bits of the data offset (types 2, 3, 5, 6) |
| word 1, bits 11–9 | **Cell type** |
| word 1, bits 8–6 | Alignment (values below; from the importer's mapping of alignment letters C, G, L, R, X to 2, 0, 1, 3, 4; all five occur in the samples). |
| word 1, bits 5–0 | Number format index: 0–18 built in, 19 onwards the user formats of §7 in order |
| word 2, bit 15 | Hidden (CELL.PROTECTION) |
| word 2, bit 14 | Locked (CELL.PROTECTION) |
| word 2, bit 13 | **Formula uses an array-valued function.** The parser sets it when the formula calls a function whose table entry is marked as returning an array — LINEST, TREND, LOGEST, GROWTH, TRANSPOSE, WINDOWS, DOCUMENTS, LINKS — or ROW or COLUMN in the parser's array mode, or a name whose flags byte has `0x20`. Preserved by Open. |
| word 2, bits 12–6 | Transient dependency and recalculation marks; Open clears them (`AND 0xE03F`) |
| word 2, bit 5 | Bottom border |
| word 2, bit 4 | Top border |
| word 2, bit 3 | Right border |
| word 2, bit 2 | Left border |
| word 2, bit 1 | Italic |
| word 2, bit 0 | Bold |

Alignment:

| Value | Meaning |
|---:|---|
| 0 | General |
| 1 | Left |
| 2 | Centre |
| 3 | Right |
| 4 | Fill |

The importer gives the style and border bits: it maps the letters D, I, L, R, T, B to bits 0, 1, 2, 3, 4, 5
respectively.

Cell types:

| Type | Meaning | word 0 / data |
|-----:|---------|---------------|
| 0 | Empty (formatting only) | word 0 is a pointer left over from memory; ignore it |
| 1 | Integer | word 0 = signed 16-bit value, stored inline. Bits 15–12 of word 1 are not meaningful (always `0xD` in the sample). |
| 2 | Number | 8-byte double in cell data |
| 3 | Text | Pascal string in cell data |
| 4 | Not seen in any sample | No data in cell data |
| 5 | Formula, numeric result | Formula record in cell data (§12) |
| 6 | Formula, non-numeric result | Formula record in cell data (§12) |

For types 2, 3, 5 and 6 the cell's data is at

```
section5_start + (word0 << 4) + (word1 >> 12)
```

Data records appear in section 5 in the same order as their cells appear in the table, with no gaps, so the
offsets are strictly increasing.

Dates are serial numbers in the **1904 system**: Excel 1.00 gives `DATE(1984,1,1)` as 29220 (`Literals`).

---

## 10. Section 3 — external names

The names this sheet uses from each external document. There is one list per entry of the external-document
table (§8), in the same order. Each list holds:

| Size | Meaning |
|-----:|---------|
| Pascal strings | The names used from that document, one per name (never empty) |
| byte | `0x00`, ending the list |

Save lists only the names that some formula or definition actually uses. It numbers them from 1 within the
list, and that number is the index a name token inside an external reference holds (§12.5, §11). The entry for
this document itself (zero-length name, §8) always gets an empty list: a single `0x00`. With no
external-document table the section is empty.

`DB.FORM` has one external document, the active sheet, so section 3 is one list: `08 "DATABASE" 04 "FORM" 00`.
Its formulas refer to these as `!DATABASE` (index 1) and `!FORM` (index 2).

---

## 11. Section 4 — defined names

The defined names, then the data of the array constants their definitions contain.

**Names.** A sequence of entries in case-insensitive alphabetical order, ending with a `0x00` byte. A name's
1-based position is the index `0x23` tokens use, so Save numbers the names in this order (`47.0b` in 1.00,
`45.0c` in 1.03).

| Size | Meaning |
|-----:|---------|
| byte | `cch` = name length (never 0; 0 is the terminator) |
| byte | `cce` = definition length |
| byte | Flags (below) |
| 2 bytes | Only if flag `0x40`: the shortcut key (the character typed with Option-Command; 0 for none), then the **macro kind** (below) |
| cce bytes | Definition, a parsed expression (§15) |
| byte | Copy of `cce` |
| cch bytes | Name characters |

Name flags:

| Flag | Meaning |
|-----:|---------|
| `0x04` | The definition is a reference: one cell or area token (`2C`, `2D`) or a union in its definition form (`2E`). Every such name in the samples has it, on worksheets and macro sheets alike, and no other name does. Excel sets it when it saves, so macro names, which point at a cell, have it too. |
| `0x20` | The definition is array-valued, so formulas using the name get cell word 2 bit 13 (§9). |
| `0x40` | A macro, function or command; two more bytes follow. |

Open keeps only `0x64`, so the `0x01`, `0x02` and `0x80` some samples have are in-memory marks that Save
writes out and Open discards.

Macro kind:

| Flag | Meaning |
|-----:|---------|
| `0x40` | Command macro, listed in Macro > Run |
| `0x80` | Function macro, called from formulas |

Open keeps those two bits and discards the low nibble (left over from memory in the samples: `0x40`–`0x4E`,
`0x80`–`0x8A`). For example, `DEFINE.NAME("Cmd",…,2,"q")` gives flags `0x44`, key `q`, kind `0x44`, and
`DEFINE.NAME("Fn",…,1)` gives flags `0x44`, key 0, kind `0x80`. In the samples, `MATRIX`'s INV and MULT and
`BEARINGS`'s Coordinates are function macros; every other macro is a command.

Definitions differ from cell formulas:

* Fix-ups are applied **in place**. A `0x23` token's first two operand bytes hold the name index; a `0x1A`
  token's last three operand bytes are `00 00 <document index>`; and a name token inside an external reference
  holds the index in that document's list in section 3 (§10).
* A definition can be a union written without parentheses, such as `=$C$9,$E$9,$C$10` (`INVOICE TEMPLATE`).
* Token classes follow the rules for array formulas (§15.4).

A new worksheet always has five names with empty definitions: `Criteria`, `Database`, `Print_Area`,
`Print_Titles`, `Recorder`. From 1.01 on they are items 30–34 of the string list `STR#` 256, so a translated
Excel gives its own: the French 1.01's are `Critères`, `Base_de_données`, `Zone_d_impression`,
`Impression_des_titres` and `Enregistrement` (`New Sheet French` sample). Excel adds its five names to any
sheet it opens that lacks them, so a sheet with English names saved by the French Excel has both sets
(`Formulas French`).

Example (`Expenses`): `08 07 04 | 2D 00 09 00 FF 00 03 | 07 | "Database"`, the area rows 9–255, columns 0–3.

**Array data.** After the terminator come the data blocks (§12.5) of every array constant (`0x20` token) in
the definitions, name by name in the order above. `MATRIX` defines `col.tmp` as `={0;2;-3;1;0;0}`: its
definition is `60 01 00 06 77 E4 04` (array constant, one column, six rows, three bytes left over from
memory), and its block, `01 00 06` followed by six elements, follows the names.

---

## 12. Section 5 — cell data

### 12.1 Number (type 2)

8 bytes, IEEE double.

### 12.2 Text (type 3)

Pascal string (up to 255 characters).

### 12.3 Formula record (types 5 and 6)

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0  | 8 bytes | Current value (below) |
| 8  | 2 bytes | Left over from memory: values such as `FF FF` and `77 F2` occur in the samples. Ignore; write zero. |
| 10 | byte | Recalculation scratch; varies. Ignore; write zero. |
| 11 | byte | Column |
| 12 | word | Row and two marks (below) |
| 14 | byte | `cce` = length of the parsed expression |
| 15 | cce bytes | Parsed expression (§15) |
| … | Pascal string | Present only if the value is a string |
| … | | Trailing fix-ups (§12.5) |

Row word (offset 12):

| Bits | Meaning |
|-----:|---------|
| 15 | A transient "needs fix-up" mark, set while importing and cleared by a later pass; 0 in the samples |
| 14 | Mirrors "alignment is Fill" for the cell, set from bit 2 of the alignment code whenever alignment is applied; 0 in the samples |
| 13–0 | Row |

### 12.4 Current value

* If bytes 0–1 are **not** `FF FF`: an IEEE double (cell type 5).
* If bytes 0–1 are `FF FF` (cell type 6), byte 2 selects:

| Byte 2 | Meaning | Where the value is |
|-------:|---------|--------------------|
| 0 | String | Bytes 4–7 are a pointer left over from memory. The string follows the expression as a Pascal string. |
| 1 | Boolean | Byte 4: 0 = FALSE, 1 = TRUE |
| 2 | Error | Byte 4 = error code, the same as the later BIFF codes (below; the `Errors` sample has all seven) |

Error codes:

| Code | Error |
|---:|---|
| 0 (`0x00`) | #NULL! |
| 7 (`0x07`) | #DIV/0! |
| 15 (`0x0F`) | #VALUE! |
| 23 (`0x17`) | #REF! |
| 29 (`0x1D`) | #NAME? |
| 36 (`0x24`) | #NUM! |
| 42 (`0x2A`) | #N/A |

### 12.5 Trailing fix-ups

In memory, some tokens hold pointers. Save writes the expression unchanged (pointer bytes and all) and then
appends replacements, in token order:

| Token | Appended | Meaning |
|-------|----------|---------|
| `0x23` name | word | 1-based index of the name in section 4. Inside an external reference (between `1A` and `1B`) it is instead the 1-based index in that document's list in section 3. |
| `0x1A` external sheet | word | 1-based index of the external document in §8. An entry named `\0` stands for the active sheet, written `!A1` (§8) |
| `0x20` array constant | array block (below) | The constant's elements |
| `0x26` memory area | word `n`, then `n` × 6-byte areas (first row word, last row word, first column byte, last column byte) | The areas the wrapped union or intersection stands for, as the compiler worked them out (§15.4). |

Array block: byte column count (0 means 256), word row count, then the `rows × cols` elements row by row, each
a type byte and a value: type 1 is a number (a double) and type 2 a string (a Pascal string). Readers treat
any other type as 8 bytes; no other type occurs. The same blocks hold the array constants of name definitions
(§11). For example, in `MATRIX`, `{0,1,3;2,-1,2;-3,2,-2}` is `09 00 01` and nine type-1 elements.

### 12.6 Array formulas and tables

If the first token of the expression is `0x01` (array formula) or `0x02` (TABLE), the cell does not own its
expression: the token's three operand bytes are a pointer to a shared record, left over from memory, and the
trailer replaces the fix-ups of §12.5. Every cell of the range carries the token; the top-left cell, the
*anchor*, holds the shared record. Array formulas occur in `MATRIX`, `SIMUL EQUATIONS` and `ACCOUNTS DATA`;
tables in the experiment samples `Table`, `Table Row Input`, `Table One Argument`, `Table Deleted Input` and
`Table Relinked`.

*Anchor cell* of an array formula:

| Size | Meaning |
|-----:|---------|
| word | `0x4000`, a marker that cannot be a row number |
| byte | 0 |
| 10 bytes | Shared record header: first row word, last row word, first column, last column, three bytes left over from memory, then (offset 9) the length `n` of the shared expression |
| n bytes | The shared expression |
| … | Its fix-ups, as in §12.5 |

*Anchor cell* of a table: the marker word and byte, then the table's 18-byte in-memory record, which Open
reads as it is (`46.06` in 1.00, `46.07` in 1.01) and the TABLE command builds (`4b.06` in 1.00):

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 | word | First row of the table's cells |
| 2 | word | Last row |
| 4 | byte | First column |
| 5 | byte | Last column |
| 6 | 3 bytes | Memory pointer linking the record into one of Excel's table lists, or 0. A table waiting to be calculated goes on a list that the table calculation (`3f.00` in 1.00) takes and empties, and `1d.06` (`1d.07` in 1.01) unlinks a record. Open reads the saved value: when it is not 0, Open links the table into a list again (`46.06`; `46.07` in 1.01). A table saved with a non-zero value in Excel 1.00 was not calculated on Open and was saved back with 0. Write 0, as every saved table has. |
| 9 | byte | Column of the first input cell |
| 10 | word | Row of the first input cell, with two flags (input word, below) |
| 12 | word | With bits 5–0 of byte 14, a pointer linking the record into Excel's list of all tables; Open links every table it reads (`3f.04`, `46.06` in 1.00; `46.07` in 1.01). Left over from memory (`FF FF` in every table seen). |
| 14 | byte | Input flags and the top of the pointer at offset 12 (below) |
| 15 | byte | Column of the second input cell |
| 16 | word | Row of the second input cell, with the same flags as offset 10 |

Input word (offsets 10 and 16):

| Bits | Meaning |
|-----:|---------|
| 15–2 | Row of the input cell |
| 1 | The input was deleted by the last command, which can still be undone. The move and delete code sets bits 1 and 0 together and keeps the old record for Undo (`4f.0e`); Undo clears bit 0 of an input with bit 1 set (`0e.01`), and the next command clears bit 1 (`4f.0b`). 0 in every saved file seen, `Table Deleted Input` included. |
| 0 | The input cell has been deleted, and the formula shows `#REF!` in its place (`1d.01` in 1.00) |

Byte 14:

| Bits | Meaning |
|-----:|---------|
| 7 | There is a second input, the column input |
| 6 | The first input is the row input (without it, the column input) |
| 5–0 | The top of the pointer at offset 12, left over from memory, so they differ between saves of the same table (`C0`, `C6`, `C8`, `CE`) |

`TABLE(Y7,X5)` on B2:C3 is `00 01 00 02 01 02 00 00 00 18 00 18 FF FF C0 17 00 10`. `TABLE(,X1)` has X1 as its
first input and byte 14 = 0; `TABLE(Y7,)` has Y7 as its first input and byte 14 = `40`, and bytes 15–17 are
left over from memory. The macro command `TABLE(Y7)`, with one argument, also sets bit 7 and takes a second
input from memory it never set ([Excel bugs and quirks](ExcelBugs.md); `4b.05` in 1.00).

GET.FORMULA gives the inputs as absolute references, `=TABLE(R7C25,R5C24)`, and a deleted one as `#REF!`:
`=TABLE(#REF!,R5C24)` after Y7 is deleted.

*Any other cell* of the range:

| Size | Meaning |
|-----:|---------|
| word | Anchor row |
| byte | Anchor column |

Example (`SIMUL EQUATIONS` B3, `{=SET.VALUE(solution,data)}`): the cell's expression is `01 77 51 0A`, and its
trailer `40 00 00 | 00 02 00 04 01 01 … 0A | 23 … 63 … 42 02 C2`, a range of rows 2–4 in column B. The shared
expression follows the array-formula class rules of §15.4.

---

## 13. Section 6 — windows

One record per open window on the document, front to back, terminated by a `0x00` byte.

| Size | Meaning |
|-----:|---------|
| byte | `0xFF` marker |
| 4 words | Window rectangle, in global (screen) coordinates |
| 41 bytes (v17), 129 bytes (v18) | Window state, copied from offsets 183–223 (v17) or 183–311 (v18) of the window record |
| word | `n` = number of selected areas |
| n × 6 bytes | Each: first row word, last row word, first column byte, last column byte |

On Open the rectangle is first adjusted to fit the current screen.

Window state block, by offset within the 41 bytes. Taken from the routine that initialises a new window, the
one that stores the cursor, and the DISPLAY command:

| Offset | Size | New window | Meaning |
|-------:|-----:|-----------|---------|
| 0 | byte | 0 | **Active cell column** |
| 1 | byte | | Width of the row-heading strip in pixels, capped at 255 (29 in every sample). Used as the left offset when headings are shown. |
| 2 | byte | 0 | **Active pane** (0–3); it indexes the per-pane bytes at offset 19 |
| 3 | word | 0 | Top row, upper panes |
| 5 | word | 0 | Top row, lower panes |
| 7 | word | rows − 2 | Bottom row, upper panes |
| 9 | word | rows − 2 | Bottom row, lower panes |
| 11 | byte | 0 | Left column, left panes |
| 12 | byte | 0 | Left column, right panes |
| 13 | byte | cols − 1 | Right column, left panes |
| 14 | byte | cols − 1 | Right column, right panes |
| 15 | word | cols − 1 | Word companion of offset 13; compared against the last-non-hidden-column bound (§6, 0x016) when drawing |
| 17 | word | cols − 1 | Word companion of offset 14 |
| 19 | 4 bytes | 0 | One byte per pane: the pattern phase (0–7). Excel rotates its 8-byte QuickDraw pattern by this many bytes before drawing in the pane, so that patterned drawing stays aligned as the pane scrolls. |
| 23 | 3 bytes | | Pointer to the selection list, left over from memory |
| 26 | byte | 1 on a macro sheet | **Display formulas** |
| 27 | byte | 1 | **Display gridlines** |
| 28 | byte | 1 | **Display row and column headings** |
| 29 | 3 bytes | | Pointer to the document, left over from memory |
| 32 | byte | | Padding, never initialised or read: left over from memory |
| 33 | word | 0 | **Index of the active area** within the selection list that follows (the area containing the active cell) |
| 35 | word | 0 | **Active cell row** |
| 37 | word | 0 | Vertical pixel position of the horizontal split (between upper and lower panes); 0 = not split |
| 39 | word | 0 | Horizontal pixel position of the vertical split (between left and right panes); 0 = not split |

"rows" and "cols" are the counts that fit the window when it is created. A window has up to four panes: the
first value of each pair describes the upper (or left) panes, the second the lower (or right) panes. This
comes from the print code, which builds a temporary window and stores the title rows in offsets 3 and 7, the
body rows in 5 and 9, and the boundary in 37, and does the same for columns with 11/13, 12/14 and 39.

None of the samples is split, so their second values are leftovers (for example bottom rows 17 and 20 in
`Expenses`). When formulas are displayed, columns are drawn at double width, which is why the macro sheets
show a right column of 2.

Check against `Worksheet1`: offset 0 is 3 and offset 35 is 4, and its single selection is row 4, column 3.
Every sample has exactly one window.

**Version 18.** Excel 1.5's window record is longer, and version-18 files save 88 more bytes of it after the
41, giving a 129-byte state. By state offset (window-record offset − 183):

| Offset | Size | New window | Meaning |
|-------:|-----:|-----------|---------|
| 41 | 80 bytes | 0 | **Rows waiting to be redrawn**, one byte per row on screen. Excel 1.5 marks a row in every window of the document that shows it (`22.02`, native code); `22.03` redraws the marked rows and clears the marks, and scrolling shifts the bytes with the rows (`18.01`). They are 0 in every sample. |
| 121 | byte | 0 | **Panes frozen.** FREEZE.PANES toggles it; SPLIT does nothing while it is set (FREEZE.PANES is `0f.3e` and SPLIT `0f.47` in 1.5). |
| 122 | byte | | Not used: never written, so left over from memory |
| 123 | word | | **Row of the frozen split**: the first row of the lower panes, counting from 0. FREEZE.PANES stores it (`0f.40`), and while the panes are frozen scrolling stops there (`18.01`). Left over from memory until the window's panes are first frozen. |
| 125 | word | | **Column of the frozen split**: the first column of the right panes, counting from 0, stored and used the same way |
| 127 | byte | 1 | **Gridline and heading colour is automatic** (the DISPLAY dialog, `0d.20` in 1.5, shows "automatic" when it is set) |
| 128 | byte | 0 | Gridline and heading colour, used when offset 127 is 0: 0–7 in the order of `STR#` 263 (§7). DISPLAY's colour argument counts from 1: `DISPLAY(FALSE,TRUE,TRUE,TRUE,3)` saves 2, red. |

`SPLIT(1,2)` then `FREEZE.PANES(TRUE)` saves 1, 2 and 1 at offsets 121, 123 and 125 (`Frozen Panes`). A window
with panes frozen (121) or a colour that is not automatic (127 = 0) makes Excel 1.5 save version 18 (§2.1),
since a version-17 file cannot hold either.

---

## 14. Section 7 — page breaks and print settings

Three parts, each self-terminating. Only the `Page Breaks` experiment sample has page breaks; three of
Microsoft's samples have a print record (`MATRIX`, `BEARINGS`, `JAPAN`) and the rest a single `0x00` for it.

**Excel 1.00: extra function names.** In a file without an extra header (§5, `n = 0`), when the function count
is above the reader's base count, section 7 starts with one Pascal string per function from the base count up
to the function count. Excel 1.00's Save writes one name per table entry from its base count to the end of its
table (`47.0b`), and Open reads them before the page breaks and compares each with its own table entry for
that number, refusing the file ("File error", alert `0x2038`) on any difference (`46.06`). So a build
whose table holds functions beyond its base set checks that a file numbers them the same way. Released
1.00's base count and table length are both 208, so it never writes the list, and it refuses a file whose count
is above 208 before reaching it (§5): a sheet the library wrote with count 209 and one name gets "File is not a
saved worksheet", and the same sheet with count 208 opens. No sample has the list. Excel 1.01 dropped it in
both Save and Open (`47.0b`, `46.07` in 1.01), and later files carry no names even when their count is higher.

1. **Column page breaks**: a sequence of chunks, each a count byte `n` (1…5) followed by `n` bytes, one column
   number per byte. Terminated by `0x00`.
2. **Row page breaks**: a sequence of chunks, each a count byte `n` (1…5) followed by `n` words, one row
   number per word. Terminated by `0x00`.

   In memory both are linked lists of 14-byte nodes holding five word slots; every node but the last is full,
   which is why a chunk never exceeds five. That the byte-sized list is columns and the word-sized list is
   rows follows from the value ranges. `SET.PAGE.BREAK` with C5 selected saves column break 1 and row break 3:
   each number is the 0-based index of the last column or row **before** the break.
3. **Print record**: `0x00` if the document has none. Otherwise `0xFF`, then a 122-byte page-setup structure,
   then three strings, each either `0x00` (absent) or a Pascal string.

   The three strings are, in order, the **page header**, the **page footer** and the **printer message** (the
   set-up string sent to a serial "TTY" printer). Default header and footer are `&f` and `Page &p`.

   The 122-byte structure is Excel's Page Setup record:

   | Offset | Size | Meaning |
   |-------:|-----:|---------|
   | 0 | 62 bytes | The first 62 bytes of the Toolbox print record (`TPrint`: version, `TPrInfo`, paper rectangle, `TPrStl`, `prInfoPT`, `TPrXInfo`), copied in and out around the system print dialogs |
   | 62 | 4 doubles | Left, right, top and bottom margins in inches (Excel's defaults 0.75, 0.75, 1, 1 in the 1.5 samples; 0.25 all round in `MATRIX`) |
   | 94 | 3 bytes | Pointer to the page header, left over from memory |
   | 97 | byte | Print row and column headings |
   | 98 | 3 bytes | Pointer to the page footer, left over from memory |
   | 101 | byte | Print gridlines |
   | 102 | 4 words | Printable-area rectangle; Excel derives print width and rows per page from it |
   | 110 | word | TTY left margin, in characters (0…256) |
   | 112 | word | TTY print width, in characters |
   | 114 | word | TTY top margin, in lines |
   | 116 | word | TTY print length, in lines |
   | 118 | 3 bytes | Pointer to the printer message, left over from memory |
   | 121 | byte | Padding after the 3-byte pointer, bringing the record to an even 122 bytes |

   Offsets 97, 101, 110–118 and the three pointers are read directly from the dialog code (dialog items "Print
   Row & Column Headings", "Page Header", "Left Margin", "Print Width", "Top Margin", "Print Length", "Printer
   Message"). The samples agree: TTY settings 9, 80, 4 and 57, printable rectangles such as (42, 23, 510,
   707), header `&f` and footer `Page &p` (or all three strings absent in `MATRIX`), and `TPrint` versions 2
   (1.00) and 3 (1.5).

The file ends after this section.

---

## 15. Parsed expressions

Reverse-Polish token stream. This is the direct ancestor of the BIFF2 `ptg` encoding: token numbers and
operand sizes coincide wherever they could be checked.

Tokens `0x01`–`0x1F` are plain bytes. Tokens from `0x20` up carry a **class** in bits 6–5:

| Bits 6–5 | Class | Example |
|---------:|-------|---------|
| `01` | Reference: the operand is used as a reference | `0x24` |
| `10` | Value: the operand's value is used | `0x44` |
| `11` | Array: the operand is used as an array | `0x64` |

The application normalises a token byte with: *if bit 6 is set, set bit 5; keep the low 6 bits*. So `0x24`,
`0x44` and `0x64` are all the cell-reference token `0x24`, and this spec names tokens by that normalised
value. §15.4 says which class the compiler chooses.

### 15.1 Tokens with no operands (`0x03`–`0x16`)

All occur in Microsoft's samples except `07`, `0F`, `12` and `14`, which Excel compiled in experiments. `01`
and `02` are described in §12.6. The binary operators come from Excel's own operator table (`PCOD` 0 at
`0xCFD` in 1.00, `0xCD5` in 1.01, `0xCE9` in 1.03), which pairs tokens `03 04 0B 0D 07 06 05 08 11 09 10 0E 0A
0C 0F` with the characters `+ - = > ^ / * & : < , < < >` and space (the second character of `<>`, `<=` and
`>=` comes from elsewhere). So `07` is `^` and `0F` is the space, Excel's intersection operator. Excel
compiles `=2^3` to `07`, `=A1:B2 B1:C3` to `0F`, `=+A1` to `12` and `=A1%` to `14`, the same meanings as in
BIFF.

Excel 1.00's operator-character table holds `+-^*/,&:<>=!`; Excel 1.01's, 1.03's and 1.5's hold `+-^*/,&:<>=
(){};!`. The difference is in the table only: Excel 1.00 compiles `=R1C1 R1C2` (an intersection) and
`={1,2;3,4}` like 1.03 does (`Literals` and `Array Constants` samples).

| Token | Meaning | Token | Meaning |
|------:|---------|------:|---------|
| 03 | add | 0D | greater than |
| 04 | subtract | 0E | not equal |
| 05 | multiply | 0F | intersection |
| 06 | divide | 10 | union |
| 07 | power | 11 | range |
| 08 | concatenate | 12 | unary plus |
| 09 | less than | 13 | unary minus |
| 0A | less or equal | 14 | percent |
| 0B | equal | 15 | parentheses |
| 0C | greater or equal | 16 | missing argument |

### 15.2 Tokens with operands

Operand sizes for `0x19`–`0x37` come from the applications' size tables and were verified by walking every
formula in the samples; `0x38`–`0x3A` are Excel 1.5's. Tokens not listed (`0x00`, `0x18`, `0x3B` and up) do
not exist.

| Token | Operand bytes | Meaning |
|------:|--------------:|---------|
| 17 | 1 + n | String constant: length byte, characters |
| 19 | 2 (+ n + 1) | Attribute: flags byte, data byte; see the table below. If flag `0x04` is set, `data + 1` further bytes follow. |
| 1A | 7 | External sheet reference start. In a cell formula all seven bytes are left over from memory and the 1-based external-document index is in the fix-ups (§12.5). In a name definition bytes 5–6 are `00 00` and byte 7 is the index (§11). |
| 1B | 3 | External sheet reference end. The three bytes are left over from memory. |
| 1C | 1 | Error constant |
| 1D | 1 | Boolean constant |
| 1E | 2 | Integer constant (word) |
| 1F | 8 | Number constant (double) |
| 20 | 6 | Array constant: column count, row count (word), three bytes left over from memory. The elements follow in the fix-ups of a cell formula (§12.5) or after the names for a definition (§11). |
| 21 | 1 | Function, fixed argument count: function number |
| 22 | 2 | Function, variable argument count: argument count, function number. Bit 7 of the count marks a command's dialog-box form (`RUN?()`, `42 80 91`) |
| 23 | 7 | Name. In a cell formula all seven bytes are left over from memory (the first three look like a pointer) and the index is in the fix-ups (§12.5). In a name definition the first two bytes are the 1-based name index (§11) and the other five are left over. |
| 24 | 3 | Cell reference: row word, column byte |
| 25 | 6 | Area reference: first row word, last row word, first column byte, last column byte |
| 26 | 4 | Memory area: three bytes left over from memory, then the length in bytes of the subexpression that follows. Wraps a union or intersection of plain cell references; the areas it stands for are in the fix-ups (§12.5, §15.4). At recalculation Excel rewrites a memory token by its result (`25.01`): `26` when the areas were worked out, `27` when the expression gave an error, `28` when there was no memory to hold the list. |
| 27, 28 | 4 | Memory error and memory with no memory: forms of `26` that the compiler and recalculation write (§15.4). Operand bytes: 0, the error code (§12.4), a byte left over from memory, then the length as in `26`. The compiler writes `27` in place of a memory area that does not stand for exactly one area when a single value is needed (`11.09` in 1.00, `11.0a` in 1.01): #NULL! (`00`) for no area, #VALUE! (`0F`) for more than one. Recalculation writes `28` when it cannot get memory for the area list (`25.01`). |
| 29 | 1 | Memory function: the length in bytes of the subexpression that follows. Wraps a range or union that involves a function result or a name (§15.4). |
| 2A | 3 | Deleted cell reference. The bytes look like the former row word and column (`2A 00 84 02`); nothing reads them. |
| 2B | 6 | Deleted area reference |
| 2C | 3 | Cell reference inside a name definition |
| 2D | 6 | Area reference inside a name definition |
| 2E | 1 | Memory function inside a name definition (the definition form of `29`; `INVOICE TEMPLATE`) |
| 2F | 1 | Memory function with no memory, in a name definition: what `2E` becomes when recalculation runs out of memory while evaluating it. The evaluator rewrites the token's low five bits to `0F` (`25.01` in 1.00), as it turns `26`–`28` into `28`, and Open turns `2F` back into `2E` (`46.06`; `46.07` in 1.01) so that it is evaluated again. |
| 30 | 7 | Marked copy of `23` (name) |
| 31 | 3 | Marked copy of `24` (cell reference) |
| 32 | 6 | Marked copy of `25` (area reference) |
| 33, 34, 35 | 4 | Marked copies of `26`, `27`, `28` |
| 36 | 1 | Marked copy of `29` |
| 37 | 3 | Second marked copy of `24` |
| 38 | 2 | Excel 1.5's in-memory form of a call with an argument count to a function it numbers internally from 178. Its Save translates it; it does not appear in files. |
| 39 | 1 | Excel 1.5: function with a fixed argument count, numbered 256 or more: the number minus 256 (Save, `47`) |
| 3A | 2 | Excel 1.5: function with an argument count, numbered 256 or more: argument count, then the number minus 256 |

**Attributes** (token `19`). The flags byte says which kind:

| Flags | Name | Data byte and what follows |
|------:|------|----------------------------|
| `01` | Volatile | First token of a formula that calls a volatile function. Data byte is left over from memory. |
| `02` | IF jump | Follows IF's condition. Data = bytes to skip when the condition is false (§15.4). |
| `04` | CHOOSE jump table | Follows CHOOSE's index. Data = `n`, the number of choices; `n + 1` offset bytes follow (§15.4). |
| `08` | Skip | Follows each IF branch and each CHOOSE choice. Data = bytes to skip, minus 1, to reach the end of the IF or CHOOSE token. |
| `10` | SUM | Replaces the function token of `SUM` with one argument. Data byte is left over from memory. |

All five occur in the samples. A reader that only wants the formula text can ignore every attribute except
`10`, which stands for `SUM(…)`.

Tokens `30`–`37` are working forms. While references are being adjusted, a routine adds 13 to tokens `23`–`29`
that involve a relative reference (and 19 to a cell reference whose column falls outside a given range, giving
`37`); the matching pass subtracts the same amount to restore them. They have the same operands as the token
they stand for and are not expected in a saved file.

Row words in references:

| Bits | Meaning |
|-----:|---------|
| 15 | The row is relative |
| 14 | The column is relative |
| 13–0 | Row |

`C0 12` is relative row 18 with a relative column.

### 15.3 Function numbers

A function token holds the function's number in one byte. Appendix A lists every number with its name,
argument rules and the release that introduced it; §2.2 explains why the numbers are stable across releases.
Numbers 0–55 agree with later BIFF; from 56 on they do not (Excel 1.x inserts COLUMN.WIDTH and FORMULA at
56–57). Function 255 calls a user-defined function (§15.4).

**Names.** Excel 1.00 keeps the names as Pascal strings inside `PCOD` 0; Excel 1.01 and later move the list to
`STR#` 257, so a translated release has its own names (the French 1.01's begin `VRAI`, `FAUX`). The numbers do
not change, so a file opens the same in any language. With `list` the list, zero-based, in 1.00–1.03:

```
number 0…33   -> list[number + 2]     (COUNT, IF, ISNA, … VALUE)
number 34, 35 -> TRUE, FALSE          (list[0], list[1])
number 36…    -> list[number]         (AND, OR, NOT, … )
```

Excel 1.5's list (280 entries) uses the same mapping for its *internal* numbers, which differ from file
numbers from 56 on; its translation table (§2.2) maps between them. Functions 176, 203 and 204 are named just
`7`, `8` and `9`, and Excel 1.5 gives them no file number.

**Function table.** Each release has a table of 10-byte entries, one per function, in `PCOD` 0: at `0x1B8C` in
1.00 and `0x1512` in 1.03 (208 entries each, in file order, with identical argument rules), and at `0x1C92` in
1.5 (in internal order).

| Byte | Meaning |
|-----:|---------|
| 0–1 | Word: name offset and flags (below) |
| 2 | Low nibble = minimum arguments; high nibble − 1 = maximum arguments |
| 3 | Return type |
| 4 | Type the first argument accepts |
| 5 | Type the other arguments accept |
| 6–9 | How the function is implemented. Byte 7 selects a family of routines (`0x1C` for commands, `0x11` for SUM-like aggregates, `0x12` for IF and CHOOSE, `0x16` for date and financial functions, the same in 1.00 and 1.03); byte 6 is an argument to it; for commands bytes 8–9 are the segment and procedure that carry them out, in that release's numbering (PROTECT.DOCUMENT is `24.07` and TABLE `4b.05` in 1.00, `22.07` and `4a.05` in 1.03). Excel 1.5's families are renumbered (`0x28` for commands). |

Bytes 0–1:

| Bits | Meaning |
|-----:|---------|
| 15–4 | The offset of the function's name in the name list. In Excel 1.00 the list is the Pascal strings at `PCOD` 0 offset `0x14FA` (COUNT is at `0x00B`, TRUE at 0); Save's function-name list (§14) reads names this way. Excel 1.03 and 1.5 build the list at start-up and use the same field as an offset into it (`g0812` in 1.03), with different values. |
| 3–0 | Flags (below) |

Function flags:

| Flag | Meaning |
|-----:|---------|
| `0x01` | Macro sheets only |
| `0x02` | Cannot be used in a cell formula: set on SERIES, which belongs in charts, and on the unnamed placeholders 176, 203 and 204. Excel 1.00's FORMULA would not enter `=SERIES("a",R1C1,R1C2,1)` in a macro-sheet cell (`Empty Arguments` sample). |
| `0x04` | Has a dialog-box form: the compiler accepts `NAME?(…)` only for functions with this flag (`11.00` in 1.00) |
| `0x08` | Volatile: set on INDEX, RAND, NOW, AREAS, ROWS and COLUMNS; formulas that call these start with `19 01` |

Type bytes are bit masks:

| Flag | Meaning |
|-----:|---------|
| `0x01` | Number |
| `0x02` | Text |
| `0x04` | Logical |
| `0x08` | Reference |
| `0x10` | Error |
| `0x20` | Command |
| `0x40` | Array |
| `0x80` | An empty argument means "not given" |

The compiler accepts an empty argument anywhere (`=ROUND(1,)` is `1E 00 01 16 41 1B`). Where the argument's
type has `0x80`, the function is told the argument was not given; elsewhere it sees 0. In Excel 1.00
`=DCOUNT(D1:D4,,F1:F2)` counts the matching records, as with no field, while `=DCOUNT(D1:D4,0,F1:F2)` gives
#VALUE! (`DCount Empty Field` sample); and `=ROUND(1,)` gives 1 and `=MID("abc",,1)` gives #VALUE!, as with 0
(`Empty Arguments` sample). The bit is on the arguments of the commands whose omitted arguments mean "leave
unchanged" or "use the dialog's setting" (STYLE, BORDER, FONT, PAGE.SETUP, CALCULATION, SORT…), OPEN's second,
INPUT's later ones and DCOUNT's field.

### 15.4 How Excel compiles formula text

A writer that wants Excel's exact bytes must follow these rules. Rules read from Excel's formula compiler
(`11.00` in 1.00, `0f.00` in 1.03) are marked *(code)*; the others come from the function tables and the
samples. They were checked by decompiling every cell formula (including array formulas) and name definition in
the samples, Microsoft's and the ones Excel saved in experiments, and compiling the text again for the release
that saved each file (§2.3): the tokens match byte for byte, apart from the bytes §15.2 describes as left over
from memory, in every case except the pre-release 1.5 build's FORMULA calls (§2.3) and the few formulas the
unit tests list as exceptions, which Excel kept from an earlier version of the library or which A1 text cannot
express.

Excel 1.01's compiler (`11.00`, identical to 1.03's `0f.00`) produces the same tokens as 1.00's; the two
differ only in reading text. 1.00 always separates arguments and array elements with a comma (character 44 in
its tokeniser). From 1.01 on Excel copies the system's international settings (`INTL` resource 0) at start-up
and uses their list separator (byte 2) instead: in the compiler, in the formula text it writes back (the
separator replaces the comma in its operator-character table) and in TABLE's text (`1d.01`). 1.01 forces a
comma when the settings' country code is 0, the United States (`04.0e` in 1.01); 1.03 reads its exceptions
from a string resource instead (`02.11` in 1.03). So the French 1.01 on a French system shows `=SOMME(A1;B1)`
for the same tokens as `=SUM(A1,B1)`. On an English system it keeps the comma, and GET.FORMULA in the French
1.01 gives `=SOMME(L[-3]C[2],L[-2]C[2])`, with `L` for row in R1C1 references, and `={1.2;3.4}` for
`{1,2;3,4}`, a full stop between array columns (`Driver French` sample); the tokens are unchanged. Its own
`INTL` 256, which the chart code loads while it works, has `.`, `,` and `;`.

**Token classes.** For a reference, area or name operand:

* Reference class (`0x2x`) when it is:
  * an operand of a range (`:`) or union (`,`) operator;
  * the whole of a name definition;
  * the macro in a call to a user-defined function;
  * a function argument whose accepted type includes reference (`0x08`).
* Value class (`0x4x`) anywhere else, including the whole of a cell formula and the operands of other
  operators.

For a function call:

* Reference class when it is in one of the first three places above.
* In Excel 1.00–1.03, reference class as any argument of `IF` or `CHOOSE`. In Excel 1.5 these arguments follow
  the next two rules, like any other argument (§2.1).
* Array class (`0x6x`) when it is an argument that accepts arrays and the function returns an array.
* Reference class when it is an argument that accepts references and the function can return a reference.
* Value class everywhere else.

**Array formulas and name definitions.** In a name definition, operands of operators and arguments that accept
arrays (type bit `0x40`) take array class where the rules above give value class. In the shared expression of
an array formula (§12.6), arguments that accept arrays do too, and operands of operators only where the
operator is itself inside such an argument: Excel 1.5 compiles `{=A1:A2*2}` with the area in value class (`45
…`) but `{=SUM(A1:A2*B1:B2)}` with both areas in array class (`Formulas 1.5 Sheet 2`). The expression as a
whole keeps the class the rules give.
`{=AVERAGE(D4-B9:B18)}` is `64 C0 03 03  65 C0 08 C0 11 01 01  04  42 01 05` (both operands of the subtraction
in array class, AVERAGE itself in value class); the definition `rads = PI()/180` is `61 13  1E 00 B4  06`.

**Function tokens** *(code)*. `21 fn` when the function's minimum and maximum argument counts are equal and it
is not written in its dialog-box form; otherwise `22 argc fn`. Functions numbered 256 and up use `39 fn` and
`3A argc fn`, with *fn* the number minus 256 (Excel 1.5's Save, `47`). The counts are those of the release
saving the file, with one exception: Excel 1.5 gave FORMULA, ACTIVATE, CLOSE, VSCROLL, HSCROLL, FORMULA.ARRAY
and FORMULA.FILL an optional argument, but its Save writes a call with only the old minimum number of
arguments in the fixed-count form, so older releases can still read it. `FORMULA(A1)` is `41 39` in every
release and `FORMULA(A1,B1)` is `42 02 39`, readable only by 1.5 (Save warns, and records that in the function
count, §5). Its Save also warns for DISPLAY with four or more arguments, PRINT with six or more, and OPEN with
three or more (Excel 1.04 already accepts a third OPEN argument; 1.03 and earlier do not). The pre-release 1.5
build in the samples lacks the exception (§2.3). Excel 1.5's Save (`47.19`) also writes every call of a
function whose minimum and maximum argument counts are equal in the fixed-count form, whatever its argument
count byte holds, so it drops the dialog-box bit: `DATA.DELETE?()`, the only dialog-box function with a fixed
count, is `42 80 82` from Excel 1.00 and `41 82` from 1.5 (`Formulas 1.5 Sheet 5`).

**Dialog-box forms** *(code)*. A `?` directly after a function name, as in `RUN?()`, sets bit 7 of the
argument count. That forces the `22` form: `RUN?()` is `42 80 91`. It is a syntax error unless the function
has table flag `0x04`.

**SUM with one argument** *(code and samples)*. The argument, then `19 10 xx` in place of the function token.
`=SUM(A1:A3)` is `25 C0 00 C0 02 00 00  19 10 00`. SUM with any other number of arguments is an ordinary `22
argc 04`.

**IF** *(code and samples)*. In order:

1. the condition, then `19 02 n`;
2. the true branch, then `19 08 m`;
3. if there is one, the false branch, then `19 08 m`;
4. the `IF` token, `22 argc 01`.

*n* is the number of bytes from just after the `19 02 n` to the start of the false branch, or to the `IF`
token when there is no false branch. Each *m* is the number of bytes from just after that `19 08 m` to the end
of the `IF` token, minus 1. The last *m* is therefore always 2. `=IF(A1,1,2)` is:

```
44 C0 00 00   19 02 06   1E 00 01   19 08 08   1E 00 02   19 08 02   42 03 01
condition     jump       true       skip       false      skip       IF
```

**CHOOSE** *(code and samples)*. In order:

1. the index, then `19 04 n` and a table of `n + 1` bytes, where *n* is the number of choices;
2. each choice, then `19 08 m` with *m* as for `IF`;
3. the `CHOOSE` token, `22 argc B4`.

Table entry *k* is the offset of the start of choice *k* + 1, counted from the first byte of the table. The
last entry is the offset of the `CHOOSE` token, so the first entry is always `n + 1`. The compiler refuses a
CHOOSE whose code would be longer than 255 bytes. `=CHOOSE(A1,1,2)` is:

```
44 C0 00 00   19 04 02   03 09 0F   1E 00 01   19 08 08   1E 00 02   19 08 02   42 03 B4
index         jump       table      choice 1   skip       choice 2   skip       CHOOSE
```

**Numbers.** Whole numbers 0–65535 are `1E` words; other numbers are `1F` doubles (`=65535` is `1E FF FF`,
`=65536` a double, `Literals`). A minus sign directly on a number constant is folded into it: `-1` is `1F BF
F0 00 00 00 00 00 00`, not `1E 00 01 13`. A minus sign on anything else is `13`.

**TRUE and FALSE** without parentheses are `1D` constants.

**Array constants.** `{1,2;3,4}` (rows separated by `;`) is a `20` token followed by the data block of §12.5.
It takes array class (`60`) in a name definition and as a function argument (`SUM({1,2,3})`,
`TRANSPOSE({1,2})`, `INDEX({1,2},1)`, `IF(1,{1,2},0)`), and value class (`40`) at the top of a cell formula or
as an operand of an operator (`={1,2}`, `={1,2}+1`), as Excel 1.00 compiles them (`Array Constants`).

**Volatile functions.** A formula that calls a function with table flag `0x08` (INDEX, RAND, NOW, AREAS, ROWS,
COLUMNS) anywhere starts with `19 01 xx`. Formulas that call none of these never have the attribute.

**Memory tokens.** The outermost union or intersection, and the outermost range that is not a single cell or
area, is preceded by a memory token giving the length *n* of the subexpression that follows, up to and
including its last `0F`, `10` or `11`. Ranges, unions and intersections nested inside it are not wrapped
again.

* Built only from cell references and areas on this sheet, it is a **memory area**, `26 xx xx xx n`. The
  compiler also works out the areas the expression stands for and puts them in the fix-ups (§12.5): the
  operands of a union, the overlaps of an intersection, the bounding box of a range. `=SUM((A1,B2),C3)` is `26
  … 09 24… 24… 10 15 24… 42 02 04` with areas A1 and B2, and `=A1:B2 B1:C3` is `46 … 0F 25… 25… 0F` with one
  area, B1:B2.
* Where a single value is needed (value class) and the areas are not exactly one, it is a **memory error**
  instead, with no areas in the fix-ups: `=(A1,B2)` is `47 00 0F xx 09 24… 24… 10 15` (#VALUE!), and `=A1 B1`,
  whose cells do not overlap, is `47 00 00 xx 09 24… 24… 0F` (#NULL!). Excel 1.03 compiled both (`Memory
  Error` and `Errors` samples).
* Otherwise (a name, a function result, another sheet) it is a **memory function**, `29 n` (`2E n` in a name
  definition).

Each takes the class a reference would in its place: `46` and `49` at the top of a cell formula, `26` and `29`
as a function's reference argument. The `Recompiled` and `Intersections` samples hold formulas Excel 1.03
compiled this way. Examples:

* `INDEX(A1:C3,1,1):C3` is `29 15` and 21 bytes: `25… 1E… 1E… 22 03 1D 24… 11`;
* `matrix:inverse` (two names) is `29 11 23… 23… 11`;
* the definition `=$C$9,$E$9,…,$C$6` (ten areas) is `2E 3A` and one 58-byte union chain.

**Signs.** A plus sign on a number constant is dropped (`=+1` is `1E 00 01`); on anything else it is token
`12`.

**`#REF!`.** A typed `#REF!` compiles to `2A`, a deleted cell reference, not the error constant `1C 17`. It
takes the operand class: `=#REF!` is `4A 00 00 00` (`Literals`, `Errors`), and a name defined as `#REF!` holds
`2A`.

**External references.**

* `1A`, then the operand, then `1B` (operand bytes in §15.2).
* `Budget!A1` names an entry in §8. `!A1` names the entry called `00`, which stands for the active sheet (§8).
* A name after the `!` (`!DATABASE`) is looked up in that document's list in section 3, not in this sheet's
  names (§10).
* A call to a user-defined function on another sheet, `'COMMISSION MACROS'!Bonus(SALES)`, is `1A`, the macro's
  name token, `1B`, the arguments, then `22 (argc + 1) FF`.

**Names.** `23` and seven bytes (§15.2). A call to a user-defined function is:

1. the macro's name token, in reference class;
2. the arguments;
3. `22 (argc + 1) FF`.

**Function count.** A formula that uses a function numbered above 207 needs the document's function count (§5)
raised to at least one more than that number.

---

## 16. Chart files

Charts are not written in the worksheet format. Save hands them to a separate writer (`5d` in Excel 1.00, `5c`
in 1.03) that produces a Microsoft Chart document, Finder type `MCBN`; Open reads them with `5e` (1.00) or
`5d` (1.03). This chapter comes from those routines, the chart dialogs, which show what each field means, and
the drawing code (`37` in 1.00), which works out the layout fields. Microsoft's samples are `CLOCK.CH` (from a
build before Excel 1.00: an overlay and a main chart, with free text for a clock face), `BREAK CHART` and
`PROFIT CHART` (Excel 1.5, line charts with titles, legends and up to five series). The experiment samples are
`PROFIT CHART` put through the Gallery formats (§16.10), Chart > Axes, Chart > Add Arrow, Chart > Add Legend
and Format > Patterns, and charts the library built and Excel 1.00 re-saved (`Library Column` and others).

### 16.1 Layout

```
+0      Header                         128 bytes
+128    Embedded p-code segment        header.segLength bytes
        Data area:                     <- header.dataOffset
          Chart-set header             8 bytes
          For each chart:              overlay first, then main
            Chart record               378 bytes
            X-axis scale               44 bytes, scatter charts only
            Text-object records        46 bytes each
            Text strings
            Series word list
            Data-point formats         12 bytes each
          Window                       12 bytes
          Page setup                   666 bytes
          Series lists                 <- header.seriesOffset
```

### 16.2 Header (128 bytes)

| Offset | Size | Value | Meaning |
|-------:|-----:|-------|---------|
| 0 | word | `0x018C` | Magic. Open rejects anything else. |
| 2 | word | 1, 3 or 12 | Format version. Excel 1.00 to 1.06 write 12 (`push 12` in 1.00's `5d.00` and 1.03's `5c.00`). Excel 1.5 writes 12, or 1 when either chart has more than 102 categories (the word at chart header offset 4, `64.00` in 1.5): `Chart 102 Categories` is saved as 12 and `Chart 103 Categories` as 1. The layout is the same; 1.5's reader treats 1 as 12 (internally 28, whose extra bit nothing tests), while every earlier release refuses it, so the number keeps such charts from older releases. `CLOCK.CH`, on the 1.00 data disk, has 3, so it was saved by a build before the release; Excel 1.00 opens it and saves it back as 12. Excel 1.00's to 1.06's Open reject 1, treat 2 and 3 as 12, and use bit 2 and bit 3 to decide how to read fonts and series headers. |
| 4 | word | `0x13E8` | Code signature (§4.1) |
| 6 | 6 bytes | 0 | |
| 12 | word | 1 | Number of embedded code segments |
| 14 | word | | Length of the segment |
| 16 | word | 128 | Offset of the segment |
| 18 | word | 1 or 2 | **Number of charts**: 2 when there is an overlay chart |
| 20 | word | 128 + segment length | **Offset of the data area** |
| 22 | word | | Largest point count of any series |
| 24 | word | | **Total number of series** in both series lists |
| 26 | word | | **File offset of the series lists** (absolute); 0 if there are none |
| 28 | | 0 | Zero fill to offset 128 |

The embedded segment is the application's `XCPC` 256 resource (598 bytes of p-code). Open skips it.

### 16.3 Numbers

Chart files do not store IEEE doubles. Every number is converted to Microsoft's 8-byte **decimal** format:

| Byte | Meaning |
|-----:|---------|
| 0, bit 7 | Sign (1 = negative) |
| 0, bits 6–0 | Decimal exponent + 64 |
| 1–7 | 14 BCD digits, most significant first, two per byte |

```
value = (−1)^sign × 0.d1 d2 … d14 × 10^(byte0 & 0x7F − 64)
```

Zero is eight zero bytes. A value whose exponent exceeds 63 is written as exponent 63 with all digits 9; one
below −63 as exponent −63 with digits `10 00 …`. So 1.0 is `41 10 00 00 00 00 00 00`.

### 16.4 Common sub-records

**Font descriptor** (4 bytes):

| Offset | Meaning |
|-------:|---------|
| 0 | Not used |
| 1 | Font |
| 2 | Size |
| 3, bit 5 | Bold |
| 3, bit 4 | Italic |

In memory bytes 1 and 2 are a Macintosh font number and point size. On disk they are Microsoft Chart codes:
the font byte is 16 for New York (font 2), 1 for Geneva (font 3) and 32 for anything else, plus `0x40` for
bold and `0x80` for italic; the size byte is an index 0–2, `(size − 9) / 3` for Geneva and New York and `(size
− 12) / 6` otherwise. Excel also hides the real Macintosh font number and size in spare bits nearby (listed
with each record) and restores them on load when the version has bit 2 set; for other versions it converts the
codes back.

**Line format** (4 bytes), for axis lines, gridlines, drop lines, hi-lo lines:

| Offset | Meaning |
|-------:|---------|
| 0 | Not used |
| 1 | Pattern; **0 = line not shown** |
| 2 | Word: bits 15–12: weight (below) |

Line weight:

| Value | Meaning |
|---:|---|
| 0 | Hairline |
| 1 | Medium |
| 2 | Thick |

**Area format** (16 bytes), for the chart area, plot area, legend, text boxes:

| Offset | Meaning |
|-------:|---------|
| 1 | Border pattern (0 = none) |
| 2 | Word: bits 15–12: border weight |
| 5 | Background pattern (0 = none) |
| 6 | Word of flags (below) |
| 8 | Four words: the box as top, left, height and width, in twips, worked out when the chart is drawn (`37.03` in 1.00). The chart area's is (0, 0, height, width) of the drawing rectangle (§16.5); the plot area's is relative to it. `PROFIT CHART`'s chart area is (0, 0, 4784, 9075) and its plot area (996, 1336, 2824, 7213). |

Area format word at offset 6:

| Bits | Meaning |
|-----:|---------|
| 15 | Set when a chart is loaded |
| 14 | Automatic size (text boxes); also set on load |
| 11–8 | Shadow when non-zero |

Pattern bytes index Excel's 15 patterns, in the order the Patterns dialogs show them, with 0 meaning
invisible. "Automatic" is a separate flag where the object supports it. The patterns are 8-byte QuickDraw
patterns at `PCOD` 0 offset `0x254C` in Excel 1.00:

| Value | Pattern | QuickDraw pattern (rows top to bottom, bit 7 leftmost) |
|---:|---|---|
| 0 | Invisible: not drawn |  |
| 1 | Solid black | `FF FF FF FF FF FF FF FF` |
| 2 | 75% grey | `EE BB EE BB EE BB EE BB` |
| 3 | 50% grey | `55 AA 55 AA 55 AA 55 AA` |
| 4 | 25% grey | `88 22 88 22 88 22 88 22` |
| 5 | White | `00 00 00 00 00 00 00 00` |
| 6 | Horizontal stripes | `FF FF 00 00 FF FF 00 00` |
| 7 | Vertical stripes | `CC CC CC CC CC CC CC CC` |
| 8 | Diagonal stripes, down to the right (\\) | `CC 66 33 99 CC 66 33 99` |
| 9 | Diagonal stripes, up to the right (/) | `99 33 66 CC 99 33 66 CC` |
| 10 | Square grid | `80 80 80 FF 80 80 80 80` |
| 11 | Diagonal lattice | `80 41 22 14 08 14 22 41` |
| 12 | Mostly black, thin white lines up to the right | `FC F9 F3 E7 CF 9F 3F 7E` |
| 13 | Thin black lines up to the right | `03 06 0C 18 30 60 C0 81` |
| 14 | Mostly black, thin white lines down to the right | `3F 9F CF E7 F3 F9 FC 7E` |
| 15 | Thin black lines down to the right | `C0 60 30 18 0C 06 03 81` |

### 16.5 Chart record (378 bytes)

Offsets are within the on-disk record; "mem" is the offset in Excel's 392-byte chart structure that the bytes
come from and go back to.

| Offset | mem | Size | Meaning |
|-------:|----:|-----:|---------|
| 0 | 0 | 18 | Chart header (below) |
| 18 | — | 6 | Built for the file: two words, (0, 0) for a lone chart, (1, 0) for the overlay and (0, 2) for a main chart that has an overlay (ignored on load); then a word whose bits 11–8 are the overlay option flags (below) |
| 24 | 20 | 16 | **Chart area** format |
| 40 | 36 | 24 | **Category axis** (axis record, below) |
| 64 | 62 | 24 | **Value axis** (axis record) |
| 88 | 88 | 44 | **Value-axis scale** (scale record, below) |
| 132 | 134 | 44 | Category-axis options, plot area and legend (below) |
| 176 | — | 78 | Zero; skipped on load |
| 254 | 178 | 4 | **Drop lines** line format (line and area charts) |
| 258 | 182 | 4 | **Hi-lo lines** line format (line charts) |
| 262 | 186 | 2 | Pointer, left over from memory |
| 264 | 188 | 2 | **Number of text objects** that follow (§16.7), including the titles |
| 266 | 190 | 46 | Built-in text object −2 (§16.7) |
| 312 | 250 | 46 | Built-in text object −1 |
| 358 | 310 | 2 | Pointer, left over from memory; skipped on load |
| 360 | 312 | 2 | Number of words in the series word list |
| 362 | 314 | 2 | Pointer, left over from memory |
| 364 | 316 | 2 | **Number of data-point format records** |
| 366 | 318 | 12 | **Default series format** (point-format record, §16.8) |

**Overlay option flags** (offset 22, bits 11–8; mem 18, bits 15–12). Excel sets all four when it makes an
overlay chart, and hides the overlay's axes at the same time (`0f.04` in 1.00); they are clear in every main
chart.

| Bits | Meaning |
|-----:|---------|
| 11–10 | No effect: a new chart's record clears them (`0f.06`), the reader and writer copy them, and nothing else in Excel 1.00 reads them |
| 9 | Scale the overlay's value axis over the data of both charts, so the overlay plots against the main chart's axis (`37.0a`) |
| 8 | The same for the X axis of a scatter overlay |

Showing an overlay axis with Chart > Axes sets that axis's line pattern (`Overlay Axis` sample), not these
bits. The word's low byte is 0 from Excel 1.00 and 1.03 and `0x0F` in the 1.5 samples.

The two built-in text objects are addressed as indices −2 and −1 of the chart's text list and are formatted
together with the axis labels by "apply to all". They are the templates for data labels: a new chart makes
both with kind 4, data label (`0f.05` in 1.00), and the label code (`33.10`) copies object −2 for each series
whose point format has bit 11 set (one label per series) and object −1 for each data point of a series with
bit 10 set (one label per point) (§16.8).

**Chart header** (18 bytes):

| Offset | Meaning |
|-------:|---------|
| 0 | Word of flags (below) |
| 2 | Word: **chart type** (below) |
| 4 | Byte: the number of categories, the largest point count of the chart's series, worked out when the chart is drawn (`37.13` in 1.00; 12 for `PROFIT CHART` once Excel has drawn it, 0 in overlays). The drawing code divides it by the categories between tick labels to place them (`37.2b`). |
| 5 | Byte: not used. Excel 1.00 to 1.06 never write it (1.06's `35.13` still stores a byte at offset 4), so it keeps whatever the file had (0 in a new chart, `0f.05` in 1.00). Excel 1.5 widens the category count to a word over bytes 4–5 (its `37.13` stores a word and `37.2b` reads one), so a chart Excel 1.5 saved last has 0 in byte 4 and the count here: `PROFIT CHART` has 0 and 12. Every main chart in the 1.00 and 1.03 samples that began as `PROFIT CHART` still has that 12; charts made by the library and saved by Excel 1.00 have 0. |
| 6 | Word: how much narrower the plot became when Excel laid out the axis and tick labels, the width before layout less the width after (`37.03` in 1.00): 266 for `PROFIT CHART`'s column, line, area and scatter forms, 116 for its bar forms, 300 or 405 for its pie forms. |
| 8 | Word: % overlap (column/bar). Negative means the "Overlapped" option is on; the magnitude is the percentage. |
| 10 | Word: % cluster spacing (column/bar) |
| 12 | Word: angle of the first pie slice, degrees from vertical (0–359) |
| 14 | Two words: the top and left margins, in twips, of the chart's drawing rectangle, which Excel centres in the window's drawing size (§16.6) when it draws the chart (`37.03` in 1.00). The rectangle's height and width are in the chart area format (bytes 8–15). `PROFIT CHART`: margins 208 and 302 around 4784 by 9075, in a 5200 by 9680 drawing size. |

Header word 0:

| Bits | Meaning |
|-----:|---------|
| 15 | The chart has a legend. LEGEND and Chart > Add Legend set it (`30.1b` in 1.00; `Legend` sample), and only then does a click find the legend (`36.02`). |
| 14 | 100% |
| 13 | Stacked |
| 12 | Bars (horizontal) rather than columns |
| 11 | Vary colours by category |
| 10 | Overlay only: automatic series distribution |
| 9–2 | Overlay only: the first series in the overlay chart |

Chart type:

| Value | Type |
|---:|---|
| 1 | Column, or bar with header bit 12 |
| 2 | Line |
| 3 | Pie |
| 4 | Area |
| 5 | Scatter |

The Gallery's "Main Chart" choices map as Area → 4, Bar → 1 with bit 12 set, Column → 1, Line → 2, Pie → 3,
Scatter → 5.

**Axis record** (24 bytes on disk):

| Offset | Meaning |
|-------:|---------|
| 0 | Word: tick marks and tick labels (below) |
| 2 | Font descriptor for the tick labels |
| 6 | Word: tick-mark length in twips, worked out when the chart is drawn as the smaller side of the area being laid out divided by 70 (`37.03` in 1.00) |
| 8 | Word: where the tick labels are drawn, in twips, worked out from the tick-label position when the chart is drawn; 0 when the axis has no tick labels (`37.3e` in 1.00) |
| 10 | Word: where the axis line is drawn, in twips, worked out from the crossing settings (`37.02`, `37.22`). In `PROFIT CHART`, whose value axis crosses before the first category, the value axis's is the plot area's left edge. |
| 12 | **Axis line** format; 0 pattern = axis hidden. Bits 9–2 of its word at +2 hold the saved Macintosh font number. |
| 16 | **Major gridlines** format. Bits 9–2 of its word at +2 hold the saved font size. |
| 20 | **Minor gridlines** format |

Axis word 0:

| Bits | Meaning |
|-----:|---------|
| 15–12 | Major tick marks (values below) |
| 11–8 | Minor tick marks |
| 7–4 | Tick-label position (values below) |
| 3 | Set on load |

Tick marks:

| Value | Meaning |
|---:|---|
| 0 | Invisible |
| 1 | Inside |
| 2 | Outside |
| 3 | Cross |

Tick-label position:

| Value | Meaning |
|---:|---|
| 0 | None |
| 1 | Next to axis |
| 2 | Low |
| 3 | High |

**Scale record** (44 bytes), for the value axis, and for scatter charts also the X axis:

| Offset | Meaning |
|-------:|---------|
| 0 | Minimum (decimal number) |
| 8 | Maximum |
| 16 | Major unit |
| 24 | Minor unit |
| 32 | Category axis crosses at |
| 40 | Word of flags (below) |
| 42 | Pointer, left over from memory; zeroed on load |

Scale word 40:

| Bits | Meaning |
|-----:|---------|
| 15 | Automatic minimum |
| 14 | Automatic maximum |
| 13 | Automatic major unit |
| 12 | Automatic minor unit |
| 11 | Automatic crossing point |
| 10 | Logarithmic scale |
| 9 | Values in reverse order |

**Category-axis options, plot area and legend** (44 bytes, mem 134):

| Offset | mem | Meaning |
|-------:|----:|---------|
| 0 | 134 | Byte: value axis crosses at category number |
| 1 | 135 | Byte: crossing and order flags (below) |
| 2 | 136 | Byte: number of categories between tick labels |
| 3 | 137 | Byte: not used; nothing in Excel 1.00 reads or writes it |
| 4 | 138 | Word: non-zero (a memory pointer) when an X-axis scale record follows the chart record |
| 6 | 140 | **Plot area** format (16 bytes) |
| 22 | 156 | Word: **legend position** and saved font number (below) |
| 24 | 158 | **Legend** font descriptor |
| 28 | 162 | **Legend** area format. Its byte at +7 holds the saved legend font size. |

Byte 1:

| Bits | Meaning |
|-----:|---------|
| 7 | The value axis crosses between categories |
| 6 | The value axis crosses at the maximum category |
| 5 | Categories in reverse order |

Word 22:

| Bits | Meaning |
|-----:|---------|
| 15–13 | Legend position (values below) |
| 8–1 | Saved legend font number |

Legend position:

| Value | Meaning |
|---:|---|
| 0 | Bottom |
| 1 | Corner |
| 2 | Top |
| 3 | Vertical |

For a scatter chart the category axis is a value axis: its scale is the extra 44-byte record after the chart
record, and "reverse order" lives in that record's flags instead of bit 5 above.

### 16.6 Chart-set header and trailer

The 8-byte chart-set header is three words from memory (two pointers and the number of charts) and the word
`0x0011`. Open skips it.

After the last chart:

| Size | Meaning |
|-----:|---------|
| 8 bytes | Window position and size: two 4-byte points, the first stored negated |
| 2 words | The chart's drawing width and height in twips (1/1440 inch): the window's content size less 16 pixels of scroll bar, at 72 pixels to the inch (`30.0f` in 1.00). `CLOCK.CH`'s window is 295 by 243 pixels and these are 5580 and 4540, 279 by 227 pixels. Excel scales the chart's fonts from them (`31.00`). |
| word | 1 if a page setup follows, else 0 |
| 62 bytes | First 62 bytes of the Page Setup record (§14). Byte 32 carries the "print row and column headings" flag. |
| 58 bytes | Zero |
| 4 × 8 bytes | Margins as decimal numbers, in the order top, bottom, left, right |
| 256 bytes | Header: a Pascal string in a fixed 256-byte field |
| 256 bytes | Footer, likewise |

Without a page setup the word is 0 and 664 zero bytes follow, so the page setup is always 666 bytes.

### 16.7 Text objects

A chart's text objects are the chart title, the axis titles, any attached or free text, and arrows. The chart
record gives the count (mem 188); the 46-byte records follow the chart record (or the X-axis scale), titles
first. After all the records come their strings, one Pascal string per object **except** objects with
automatic text, which have none.

Each 46-byte record is four slices of the 60-byte memory structure: bytes 0–5, 8–37, 38–43 and 46–49. In
memory offsets:

| mem | Meaning |
|----:|---------|
| 0 | Word: horizontal position in thousandths of the chart's drawing rectangle; Excel keeps it and the box's twips position at mem 30–33 in step (`33.19` in 1.00). The matching vertical position, mem 6, is not saved; Open works it out again. |
| 2 | Word: **kind** (bits 3–0, values below); bits 15–14 are cleared on load |
| 4 | Byte: series number (data labels) |
| 5 | Byte: point number, 255 = the whole series (data labels) |
| 8 | Pointer to the string, left over from memory |
| 10 | Word: text options (below) |
| 12 | Font descriptor |
| 16 | An arrow's line format (§16.4). Byte 17 is its pattern: 0 for text, so a non-zero byte 17 is what makes the object an **arrow**. The word at 18 holds the shaft weight in bits 13–12 (line weight values). |
| 20 | Text: bytes 20 and 21 hold the saved Macintosh font number and size. Arrow: the head word (below). |
| 22 | Area format of the text box. Bit 14 of its word at +6 is automatic size, and its box (+8) is the text's position and size in twips. |
| 38, 46, 48 | The arrow's other end, held as mem 0 and mem 30–33 hold the first: thousandths at 38 (and 44, not saved), twips at 46 and 48 (`33.19`). The arrowhead is drawn at this end: in the `Arrow` and `Arrow Heads` samples it is the end with the larger x, where Excel 1.00 draws the head. |

Text options (mem 10):

| Bits | Meaning |
|-----:|---------|
| 13 | Automatic text |
| 12 | Vertical orientation |
| 11–9 | Horizontal alignment (values below) |
| 8–6 | Vertical alignment (values below) |
| 5 | Show key |
| 4 | Show value |

Arrow head word (mem 20):

| Bits | Meaning |
|-----:|---------|
| 13–12 | Head width (values below) |
| 9–8 | Head length, 0 meaning no head (values below) |
| 7 | The head is open rather than filled |

An arrow is a free object (kind 6) with automatic text and no string; Chart > Add Arrow makes one, and its
Format > Patterns dialog sets the fields above (`Arrow`, `Arrow Patterns` and `Arrow Heads` samples, Excel
1.00).

Text kind:

| Value | Meaning |
|---:|---|
| 1 | Chart title |
| 2 | Category-axis title |
| 3 | Value-axis title |
| 4 | Data label |
| 6 | Free text, or an arrow when mem 17 is not 0: the hour labels of `CLOCK.CH`, and "Break-even Point" and its arrow in `PROFIT CHART` |

Horizontal alignment:

| Value | Meaning |
|---:|---|
| 0 | Centre |
| 1 | Left |
| 2 | Right |

Vertical alignment:

| Value | Meaning |
|---:|---|
| 0 | Bottom |
| 1 | Centre |
| 2 | Top |

Arrow head width:

| Value | Width |
|---:|---|
| 0 | Narrow |
| 1 | Medium |
| 2 | Wide |

Arrow head length, which is also whether there is a head:

| Value | Length |
|---:|---|
| 0 | No head (the dialog's first head style) |
| 1 | Short |
| 2 | Medium |
| 3 | Long |

With a head, bit 7 chooses the dialog's second head style (open, bit set) or third (filled, bit clear).

On load the titles (kinds 1–3) of the main chart are moved out of the list into Excel's three title slots.

### 16.8 Data-point formats and the series word list

The **series word list** is one word per series of the chart (count at mem 312): the memory pointers to the
chart's series structures, which the drawing code collects into a list (`37.13` in 1.00). Left over from
memory; Open skips it.

**Data-point format records** (12 bytes each, count at mem 316) override the default series format (mem 318,
same layout) for one series, or for one point of a single-series chart such as a pie. Excel finds a series'
record by its byte 6 (`3c.01` in 1.00).

| Offset | Meaning |
|-------:|---------|
| 0 | Not used (the line format's unused byte) |
| 1 | Line or border pattern |
| 2 | Word: line weight and automatic line (below) |
| 4 | Area pattern |
| 5 | Not used |
| 6 | Series number; in a chart with one series, the point (category) number, from 0 (`33.10`; the `Point Markers` sample's overrides are points 8 and 10) |
| 7 | Marker (below) |
| 8 | Word of flags (below) |
| 10 | Word: how far a pie slice is pulled out, as a percentage of the radius: 25 for the first slice of Pie format 3 (a point record), 20 for every slice of Pie format 4 (the default series format). GALLERY.PIE stores it as a word (`35.18` in 1.00), and the code that works out which Pie format a chart has tests the whole word (`35.17`). |

Point format word 2:

| Bits | Meaning |
|-----:|---------|
| 15–12 | Line or border weight (line weight values, §16.4) |
| 10 | Automatic line or border |

Point format word 8:

| Bits | Meaning |
|-----:|---------|
| 15 | Automatic marker |
| 14 | Not used |
| 13 | Set when a format is made (`0f.12`), never read |
| 12 | Automatic area |
| 11 | A data label for the series, made from built-in text object −2 (`33.10` in 1.00; §16.5) |
| 10 | A data label on each data point, made from built-in text object −1 |
| 9 | The labels show percentages (for a pie the label code totals the series, `33.00`); Pie format 6 sets it with bit 10 |
| 8–0 | Not used |

Excel 1.00 clears a point-format record when it makes one (`0f.12`) and never reads or writes bytes 0 and 5 or
the unused bits of word 8, so they are 0.

Marker, in the order of the Marker Pattern dialog (`Point Markers` sample, Excel 1.00):

| Value | Marker |
|---:|---|
| 0 | None |
| 1 | Filled diamond |
| 2 | Hollow diamond |
| 3 | Filled square |
| 4 | Hollow square |
| 5 | Filled triangle |
| 6 | Hollow triangle |
| 7 | X |
| 8 | Asterisk |
| 9 | Short horizontal dash |

### 16.9 Series records

The series lists start at the absolute file offset in header word 26. The first list holds the series flagged
with header bit 10, the second the rest; the header (word 24) gives the total. Every series in the samples has
bit 10 set. A chart whose only series lacks it opens and saves in Excel 1.00, but Excel draws nothing of it,
not even the title or axes. A chart file carries a **copy of the plotted data**, not just references to it.
Each series is:

| Size | Meaning |
|-----:|---------|
| word | Total length of this record, or 0 for the last record of a list (the next record then follows its strings) |
| word | Offset from the start of the record to the second array (4 + 98 + the first array). The first array can hold strings, so this is how a reader finds the second. |
| 98 bytes | Series header (below) |
| n × 8 bytes, or n strings | First array, `n` = point count: decimal numbers, or `n` Pascal strings for a text series (type 3). A numeric array shorter than `n` is padded with 1, 2, 3 … |
| n × 8 bytes | Second array, decimal numbers; the strings follow it |
| Pascal string | Series name |
| 2 strings | Two label strings, each one length byte, the first part, `.`, and as much of the second part as fits in 255. For a series of its own with no second part, Excel 1.00 writes the series name and `.` in both (`Series 1.`). |
| | Source document: `0x00` if none; otherwise the three bytes `02 1B 06` |
| Pascal string | Name of the source document. Excel 1.x plots only data on a document: a series without one opens and draws in Excel 1.00, but its SERIES formula shows bytes left over from memory as the document name. |
| Pascal string | Second reference, the values, **as formula text** in R1C1 form without `=` or the document (`R26C7:R37C7` in `PROFIT CHART`, whose document is `BREAKEVEN`). An empty reference gets a short string left over from memory when Excel saves (5 bytes in every chart Excel 1.00 re-saved in the experiments). |
| Pascal string | First reference, the categories, likewise. Excel shows an empty one held as literal text (header word 14) as just the document name: `=SERIES("S1",Data!,Data!R2C1:R2C3,1)`. |
| 4 bytes | Four empty strings |

A zero word ends the series section.

**Series header** (98 bytes). Open uses only the fields marked †; the rest are constants Excel writes for
Microsoft Chart.

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 † | word | Series type and flags (below) |
| 2 † | word | Format of the first array (below) |
| 4 † | word | Format of the second array, as word 2 |
| 6 | 2 bytes | 10, 10 |
| 8 | 2 words | 60, 60 |
| 14 † | word | How the first reference is held, and the first array's number format (below) |
| 16 † | word | The same for the second reference and array |
| 22 | 2 words | Copied from the series structure |
| 26 | 4 × 8 bytes | Four decimal numbers, all 1.0 |
| 60 † | 2 words | Point count, twice |
| 64 | 4 bytes | 0, 1, 1, 0 (byte 66 † is used for date series) |
| 68 | 11 words | All 2 |

Series word 0:

| Bits | Meaning |
|-----:|---------|
| 15–13 | Series type (values below) |
| 11 | Always 1 |
| 10 | The series belongs to the first list |
| 7 | Always 1 |

Series words 2 and 4:

| Bits | Meaning |
|-----:|---------|
| 14–13 | Always 2 |
| 12–10 | Format class of the array (below) |
| 9–6 | Its sub-code |
| 5–4 | Word 2 only: always 1 |

Series words 14 and 16:

| Bits | Meaning |
|-----:|---------|
| 15–14 | How the reference is held: 2 literal text, 3 formula |
| 13–8 | Excel number-format index of the array |

Series type:

| Value | Meaning |
|---:|---|
| 0 | Dates |
| 1, 2 | Numbers |
| 3 | Text categories |

Format class: 2 for General, 1 for formats 1–4, 4 for the currency formats 5–6, 5 for the percent formats 7–8;
other formats go through a separate mapping. For files from Excel (version bit 3 set) Open takes the format
indices from words 14 and 16 directly; for older Microsoft Chart files it derives them from the class and
sub-code.

### 16.10 Gallery formats

What each Gallery format sets in Excel 1.03, from the line chart `PROFIT CHART` put through every format by
the macro commands GALLERY.AREA, GALLERY.BAR, GALLERY.COLUMN, GALLERY.LINE, GALLERY.PIE, GALLERY.SCATTER and
COMBINATION and saved (samples `Area Drop Lines`, `Area Labels`, `Bar Stacked`, `Combination`, `Line Log
Scale`, `Pie Exploded`, `Pie Exploded All`, `Pie Percentages` and `Scatter`). Each also resets the axes' tick
labels to "next to axis" and the scale to automatic.

| Format | 1 | 2 | 3 | 4 | 5 | 6 |
|--------|---|---|---|---|---|---|
| Area | stacked | stacked 100% | stacked, drop lines | stacked, major gridlines on both axes | stacked, data labels | — |
| Bar, Column | plain | colours vary by category | stacked, overlap 100% | overlap 30% | stacked 100%, overlap 100% | value-axis major gridlines |
| Line | lines and markers | lines only (marker 0) | markers only (line pattern 0, automatic line off) | value-axis major gridlines | major gridlines on both axes | logarithmic value axis with gridlines |
| Pie | colours vary by category | patterns: colours do not vary, line pattern 5, area pattern 1 | first slice pulled out 25% | every slice pulled out 20% | data labels | percentage labels |
| Scatter | markers | automatic line off | major gridlines on both axes | logarithmic value axis with gridlines | — | — |
| Combination | column chart with a line overlay | the same, with the overlay's value axis crossing at the maximum category | line chart with a line overlay, its value axis crossing likewise | area chart with a column overlay | — | — |

Bar formats set header bit 12; stacked formats bit 13 and 100% formats bit 14; "colours vary" is bit 11
(§16.5).

---

## 17. Recipes

### 17.1 Reading

1. Check that the word at offset 0 is `0x5327`, and read the format version at offset 2: 17, or 18 for Excel
   1.5's extended files (§2.1).
2. Seek to the offset in the word at 22. Read the function count (§5); it tells you the newest functions the
   formulas may use.
3. Read the 630-byte state block. Compute `nRows` and `nCols` (§6.1).
4. Read the kind byte and the section table. Convert each entry with `hi × 512 + lo` (§1.2).
5. Skip the extra header (a word count, then that many bytes).
6. Section 0: the number-format count and the user formats, with 3-slot descriptors in version 17 and 4-slot
   ones in version 18 (§7).
7. Section 1: the external-document table, then `nRows` row-index words (§8).
8. Section 2: `nRows × nCols` six-byte cell entries (§9). For each entry of type 2, 3, 5 or 6, read its record
   at `section5 + (word0 << 4) + (word1 >> 12)` (§12).
9. Section 3: if the section is not empty, read one `0x00`-terminated list of names per external document
   (§10).
10. Section 4: names until the `0x00` terminator, then the data of any array constants in their definitions
    (§11).
11. Section 6: windows, with 41-byte states in version 17 and 129-byte ones in version 18 (§13).
12. Section 7: page breaks and the print record (§14).

For formulas, walk the token stream (§15.2) to find how many fix-ups follow the expression (§12.5), and
decompile the shared expression of array formulas from their anchor cell (§12.6). Name function numbers with
Appendix A.

### 17.2 Writing

A writer that copies the samples' choices produces files every Excel 1.x can open, as long as it writes
version 17 and uses only functions 0–207:

1. **Header.** Write the 24-byte header with two code segments and signature `0xAB27`. The data offset at word
   22 must be correct, since Open reads only that. The segments only matter when another document links to
   this one while it is closed (§4.1):
   * Excel 1.00 and 1.03 recognise their own signature and use their own copy of the code, so empty segments
     (length 0, offset 24) work with them.
   * A build with a different signature would load the file's copy instead. To support such builds, copy both
     segments from a sample.
2. **Data area.**
   * The function count: 208, or one more than the highest function number used (§5).
   * The 630-byte state block. Copy a sample's block, set the grid size (§6.1) and anything else you need, and
     leave the cached print values zero (§6).
   * The kind byte: `0x10` or `0x40`.
   * A placeholder for the section table.
   * The extra header: `00 01 00`, as Excel 1.01 and later write, or `00 00` as 1.00 does.
3. **Sections 0–1.** Section 0 is `13 00` (19 built-in formats, no user formats), or the user formats. Section
   1 is `00 00` for no external documents, then the row-index words, which can all be zero.
4. **Section 2.** Reserve `nRows × nCols × 6` bytes and fill them in as you write the cell data. Every cell
   needs an entry; an empty cell is type 0.
5. **Sections 3 and 4.** Section 3 is empty unless there are external documents. Section 4 is the names in
   case-insensitive alphabetical order, `00`, then any array-constant data. Name tokens refer to names by
   position in this order, so sort the names before compiling formulas that use them.
6. **Section 5.** The cell data, in cell-table order. Set each data cell's offset fields as you go: the offset
   from the start of section 5, with bits 19–4 in word 0 and bits 3–0 in bits 15–12 of word 1. Cell data is
   therefore limited to 1 MB.
7. **Sections 6 and 7.** Section 6 is the window records (§13), then `00`; every sample has exactly one.
   Section 7 is `00 00 00` for no page breaks and no print record.
8. Seek back and write the section table as stream positions (§1.2).

Compile formulas with the rules of §15.4 for Excel 1.00–1.03 unless the file is meant for a later release.

This recipe reproduces the samples byte for byte, and files written with it open in Excel 1.00, 1.03 and 1.5
(docs/Testing.md). A formula's stored value is what Excel shows until it recalculates, and Excel does not
recalculate a loaded formula until something it depends on changes. Set state word `0x018` to 1 to have Excel
recalculate the sheet when it opens it (§6; five 1.00 samples and seven 1.5 samples are saved that way).

---

## 18. Open questions

None of these affects reading the samples or writing valid files.

* Which build saved Microsoft's version-18 `SIMUL EQUATIONS`, and why it chose version 18 (§2.1).

---

## Appendix A. Functions

Every function number a file can contain, as Excel's own tables define it: names from `STR#` 257, argument
rules from the function tables (§15.3), and numbers 208–291 from Excel 1.5's translation table (§2.2).

* **Args**: the minimum and maximum number of arguments. A function with a single count uses token `21`; any
  other uses `22` (§15.4).
* **Returns**, **First**, **Others**: what the function returns and what its first and later arguments accept:
  **N** number, **T** text, **L** logical, **R** reference, **E** error, **C** command, **A** array, **?**
  type bit `0x80`: an empty argument means "not given" (§15.3).
* **Flags**: **M** macro sheets only, **D** has a dialog-box form (`NAME?(…)`), **V** volatile (formulas start
  with `19 01`), **X** cannot be used in a cell formula (flag `0x02`, §15.3).
* **Since**: the first release with the function (§2.1).
* **Changed in 1.5**: argument rules Excel 1.5's table gives instead. A file saved by 1.5 follows them, except
  that a call using only the old arguments keeps the old encoding (§15.4); one saved by an earlier release
  follows the main columns.

Numbers marked unused are never assigned. Function 255 is not in the table: it calls a user-defined function
(§15.4).

| No. | Name | Args | Returns | First | Others | Flags | Since | Changed in 1.5 |
|----:|------|-----:|---------|-------|--------|-------|-------|----------------|
| 0 | COUNT | 1–14 | N | NREA | NREA |  | 1.00 |  |
| 1 | IF | 2–3 | NTLRCA | LA | NTLRECA |  | 1.00 |  |
| 2 | ISNA | 1 | L | NTLE | – |  | 1.00 |  |
| 3 | ISERROR | 1 | L | NTLE | – |  | 1.00 |  |
| 4 | SUM | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 5 | AVERAGE | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 6 | MIN | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 7 | MAX | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 8 | ROW | 0–1 | NA | R | – |  | 1.00 |  |
| 9 | COLUMN | 0–1 | NA | R | – |  | 1.00 |  |
| 10 | NA | 0 | NT | – | – |  | 1.00 |  |
| 11 | NPV | 2–14 | N | N | NRA |  | 1.00 |  |
| 12 | STDEV | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 13 | DOLLAR | 1–2 | T | N | N |  | 1.00 |  |
| 14 | FIXED | 1–2 | T | N | N |  | 1.00 |  |
| 15 | SIN | 1 | N | N | – |  | 1.00 |  |
| 16 | COS | 1 | N | N | – |  | 1.00 |  |
| 17 | TAN | 1 | N | N | – |  | 1.00 |  |
| 18 | ATAN | 1 | N | N | – |  | 1.00 |  |
| 19 | PI | 0 | N | – | – |  | 1.00 |  |
| 20 | SQRT | 1 | N | N | – |  | 1.00 |  |
| 21 | EXP | 1 | N | N | – |  | 1.00 |  |
| 22 | LN | 1 | N | N | – |  | 1.00 |  |
| 23 | LOG10 | 1 | N | N | – |  | 1.00 |  |
| 24 | ABS | 1 | N | N | – |  | 1.00 |  |
| 25 | INT | 1 | N | N | – |  | 1.00 |  |
| 26 | SIGN | 1 | N | N | – |  | 1.00 |  |
| 27 | ROUND | 2 | N | N | N |  | 1.00 |  |
| 28 | LOOKUP | 2–3 | NTL | NTL | RA |  | 1.00 |  |
| 29 | INDEX | 2–4 | NTLRA | RA | N | V | 1.00 |  |
| 30 | REPT | 2 | T | T | N |  | 1.00 |  |
| 31 | MID | 3 | T | T | N |  | 1.00 |  |
| 32 | LEN | 1 | N | T | – |  | 1.00 |  |
| 33 | VALUE | 1 | N | T | – |  | 1.00 |  |
| 34 | TRUE | 0 | L | – | – |  | 1.00 |  |
| 35 | FALSE | 0 | L | – | – |  | 1.00 |  |
| 36 | AND | 1–14 | L | LRA | LRA |  | 1.00 |  |
| 37 | OR | 1–14 | L | LRA | LRA |  | 1.00 |  |
| 38 | NOT | 1 | L | L | – |  | 1.00 |  |
| 39 | MOD | 2 | N | N | N |  | 1.00 |  |
| 40 | DCOUNT | 3 | N | R | NTR? |  | 1.00 |  |
| 41 | DSUM | 3 | N | R | NTR |  | 1.00 |  |
| 42 | DAVERAGE | 3 | N | R | NTR |  | 1.00 |  |
| 43 | DMIN | 3 | N | R | NTR |  | 1.00 |  |
| 44 | DMAX | 3 | N | R | NTR |  | 1.00 |  |
| 45 | DSTDEV | 3 | N | R | NTR |  | 1.00 |  |
| 46 | VAR | 1–14 | N | NRA | NRA |  | 1.00 |  |
| 47 | DVAR | 3 | N | R | NTR |  | 1.00 |  |
| 48 | TEXT | 2 | T | N | T |  | 1.00 | first NTL |
| 49 | LINEST | 1–2 | A | RA | RA |  | 1.00 |  |
| 50 | TREND | 1–3 | A | RA | RA |  | 1.00 |  |
| 51 | LOGEST | 1–2 | A | RA | RA |  | 1.00 |  |
| 52 | GROWTH | 1–3 | A | RA | RA |  | 1.00 |  |
| 53 | GOTO | 1 | C | R | – | M | 1.00 |  |
| 54 | HALT | 0 | C | – | – | M | 1.00 |  |
| 55 | RETURN | 0–1 | C | NTLRECA | – | M | 1.00 |  |
| 56 | COLUMN.WIDTH | 0–2 | L | N? | TR? | M D | 1.00 |  |
| 57 | FORMULA | 1 | L | NTLA | – | M | 1.00 | args 1–2, first NTLA?, others TR? |
| 58 | PV | 3–5 | N | N | N |  | 1.00 |  |
| 59 | FV | 3–5 | N | N | N |  | 1.00 |  |
| 60 | NPER | 3–5 | N | N | N |  | 1.00 |  |
| 61 | PMT | 3–5 | N | N | N |  | 1.00 |  |
| 62 | RATE | 3–6 | N | N | N |  | 1.00 |  |
| 63 | MIRR | 3 | N | RA | N |  | 1.00 |  |
| 64 | IRR | 1–2 | N | RA | N |  | 1.00 |  |
| 65 | RAND | 0 | N | – | – | V | 1.00 |  |
| 66 | MATCH | 2–3 | N | NTL | NRA |  | 1.00 |  |
| 67 | DATE | 3 | N | N | N |  | 1.00 |  |
| 68 | TIME | 3 | N | N | N |  | 1.00 |  |
| 69 | DAY | 1 | N | N | – |  | 1.00 |  |
| 70 | MONTH | 1 | N | N | – |  | 1.00 |  |
| 71 | YEAR | 1 | N | N | – |  | 1.00 |  |
| 72 | WEEKDAY | 1 | N | N | – |  | 1.00 |  |
| 73 | HOUR | 1 | N | N | – |  | 1.00 |  |
| 74 | MINUTE | 1 | N | N | – |  | 1.00 |  |
| 75 | SECOND | 1 | N | N | – |  | 1.00 |  |
| 76 | NOW | 0 | N | – | – | V | 1.00 |  |
| 77 | AREAS | 1 | N | R | – | V | 1.00 |  |
| 78 | ROWS | 1 | N | RA | – | V | 1.00 |  |
| 79 | COLUMNS | 1 | N | RA | – | V | 1.00 |  |
| 80 | OFFSET | 3 | R | R | N | M | 1.00 |  |
| 81 | ABSREF | 2 | R | T | R | M | 1.00 |  |
| 82 | RELREF | 2 | T | R | R | M | 1.00 |  |
| 83 | ARGUMENT | 0–3 | L | T? | NR? | M | 1.00 |  |
| 84 | SEARCH | 2–3 | N | T | T |  | 1.00 |  |
| 85 | FORMAT.NUMBER | 0–1 | L | T | – | M D | 1.00 |  |
| 86 | ALIGNMENT | 0–1 | L | N | – | M D | 1.00 |  |
| 87 | STYLE | 0–2 | L | L? | L? | M D | 1.00 |  |
| 88 | BORDER | 0–5 | L | L? | L? | M D | 1.00 |  |
| 89 | CELL.PROTECTION | 0–2 | L | L? | L? | M D | 1.00 |  |
| 90 | CLEAR | 0–1 | L | N | – | M D | 1.00 |  |
| 91 | EDIT.DELETE | 0–1 | L | N | – | M D | 1.00 |  |
| 92 | INSERT | 0–1 | L | N | – | M D | 1.00 |  |
| 93 | CUT | 0 | L | – | – | M | 1.00 |  |
| 94 | COPY | 0 | L | – | – | M | 1.00 |  |
| 95 | PASTE | 0 | L | – | – | M | 1.00 |  |
| 96 | FILL.RIGHT | 0 | L | – | – | M | 1.00 |  |
| 97 | FILL.DOWN | 0 | L | – | – | M | 1.00 |  |
| 98 | TRANSPOSE | 1 | A | NTLA | – |  | 1.00 |  |
| 99 | ERROR | 0–2 | L | L | R | M | 1.00 |  |
| 100 | SELECT | 1–2 | L | TR? | TR? | M | 1.00 |  |
| 101 | STEP | 0 | L | – | – | M | 1.00 |  |
| 102 | DEFINE.NAME | 0–4 | L | T? | NTLREA? | M D | 1.00 |  |
| 103 | TYPE | 1 | N | NTLEA | – |  | 1.00 |  |
| 104 | ECHO | 0–1 | L | L | – | M | 1.00 |  |
| 105 | NEW | 0–1 | L | N | – | M D | 1.00 |  |
| 106 | OPEN | 0–2 | L | T | L? | M D | 1.00 | args 0–3 (from 1.04) |
| 107 | DELETE.NAME | 0–1 | L | T | – | M | 1.00 |  |
| 108 | SAVE | 0 | L | – | – | M | 1.00 |  |
| 109 | A1.R1C1 | 0–1 | L | L | – | M | 1.00 |  |
| 110 | FILE.DELETE | 0–1 | L | T | – | M D | 1.00 |  |
| 111 | PRINTER.SETUP | 0–3 | L | N? | N? | M D | 1.00 |  |
| 112 | NEW.WINDOW | 0 | L | – | – | M | 1.00 |  |
| 113 | SELECT.LAST.CELL | 0 | L | – | – | M | 1.00 |  |
| 114 | SHOW.ACTIVE.CELL | 0 | L | – | – | M | 1.00 |  |
| 115 | DISPLAY | 0–3 | L | L? | L? | M D | 1.00 | args 0–5, others NL? |
| 116 | FONT | 0–2 | L | T? | N? | M D | 1.00 |  |
| 117 | ACTIVATE | 1 | L | T | – | M | 1.00 | args 1–2, first T?, others N |
| 118 | CLOSE | 0 | L | – | – | M | 1.00 | args 0–1, first L |
| 119 | SET.NAME | 1–2 | L | T | NTLREA | M | 1.00 |  |
| 120 | CALLER | 0 | R | – | – | M | 1.00 |  |
| 121 | PASTE.SPECIAL | 0–3 | L | N? | NL? | M D | 1.00 |  |
| 122 | OPEN.LINKS | 0–14 | L | TEA | TEA | M D | 1.00 |  |
| 123 | PAGE.SETUP | 0–8 | L | T? | NTL? | M D | 1.00 |  |
| 124 | PRINT | 0–5 | L | N? | N? | M D | 1.00 | args 0–9, others NL? |
| 125 | UNDO | 0 | L | – | – | M | 1.00 |  |
| 126 | CREATE.NAMES | 0–2 | L | L? | L? | M D | 1.00 |  |
| 127 | FORMULA.FIND | 0–5 | L | T? | N? | M D | 1.00 |  |
| 128 | DATA.FIND | 0–1 | L | L | – | M | 1.00 |  |
| 129 | EXTRACT | 0–1 | L | L | – | M D | 1.00 |  |
| 130 | DATA.DELETE | 0 | L | – | – | M D | 1.00 |  |
| 131 | SET.DATABASE | 0 | L | – | – | M | 1.00 |  |
| 132 | SET.CRITERIA | 0 | L | – | – | M | 1.00 |  |
| 133 | SORT | 0–7 | L | N? | NTR? | M D | 1.00 |  |
| 134 | DATA.SERIES | 0–5 | L | N? | N? | M D | 1.00 |  |
| 135 | TABLE | 0–2 | L | TR? | TR? | M D | 1.00 |  |
| 136 | CALCULATE.NOW | 0 | L | – | – | M | 1.00 |  |
| 137 | PRECISION | 0–1 | L | L | – | M | 1.00 |  |
| 138 | CALCULATION | 0–4 | L | N? | NL? | M D | 1.00 |  |
| 139 | PROTECT.DOCUMENT | 0–1 | L | L | – | M D | 1.00 |  |
| 140 | FORMULA.GOTO | 0–1 | L | TR | – | M D | 1.00 |  |
| 141 | SET.PRINT.AREA | 0 | L | – | – | M | 1.00 |  |
| 142 | SET.PRINT.TITLES | 0 | L | – | – | M | 1.00 |  |
| 143 | SET.PAGE.BREAK | 0 | L | – | – | M | 1.00 |  |
| 144 | DELETE.FORMAT | 0–1 | L | T | – | M D | 1.00 |  |
| 145 | RUN | 0–1 | LR | TR | – | M D | 1.00 |  |
| 146 | DEREF | 1 | NTLEA | R | – | M | 1.00 |  |
| 147 | SHOW.CLIPBOARD | 0 | L | – | – | M | 1.00 |  |
| 148 | WINDOWS | 0 | A | – | – | M | 1.00 |  |
| 149 | MESSAGE | 1–2 | L | L? | T | M | 1.00 |  |
| 150 | HLINE | 1 | L | N | – | M | 1.00 |  |
| 151 | SERIES | 4 | L | T? | NRA? | X | 1.00 |  |
| 152 | REMOVE.PAGE.BREAK | 0 | L | – | – | M | 1.00 |  |
| 153 | SAVE.AS | 0–2 | L | T? | N? | M D | 1.00 |  |
| 154 | DOCUMENTS | 0 | A | – | – | M | 1.00 |  |
| 155 | VSCROLL | 1 | L | N | – | M | 1.00 | args 1–2, others L |
| 156 | ACTIVATE.NEXT | 0 | L | – | – | M | 1.00 |  |
| 157 | ACTIVATE.PREV | 0 | L | – | – | M | 1.00 |  |
| 158 | DATA.FIND.NEXT | 0 | L | – | – | M | 1.00 |  |
| 159 | DATA.FIND.PREV | 0 | L | – | – | M | 1.00 |  |
| 160 | FORMULA.FIND.NEXT | 0 | L | – | – | M | 1.00 |  |
| 161 | FORMULA.FIND.PREV | 0 | L | – | – | M | 1.00 |  |
| 162 | GALLERY.AREA | 0–1 | L | N | – | M D | 1.00 |  |
| 163 | GALLERY.LINE | 0–1 | L | N | – | M D | 1.00 |  |
| 164 | GALLERY.SCATTER | 0–1 | L | N | – | M D | 1.00 |  |
| 165 | GALLERY.COLUMN | 0–1 | L | N | – | M D | 1.00 |  |
| 166 | GALLERY.BAR | 0–1 | L | N | – | M D | 1.00 |  |
| 167 | GALLERY.PIE | 0–1 | L | N | – | M D | 1.00 |  |
| 168 | COMBINATION | 0–1 | L | N | – | M D | 1.00 |  |
| 169 | ACTIVE.CELL | 0 | R | – | – | M | 1.00 |  |
| 170 | SELECTION | 0 | R | – | – | M | 1.00 |  |
| 171 | RESULT | 1 | L | N | – | M | 1.00 |  |
| 172 | MAIN.CHART.TYPE | 0–1 | L | N | – | M D | 1.00 |  |
| 173 | OVERLAY.CHART.TYPE | 0–1 | L | N | – | M D | 1.00 |  |
| 174 | COPY.CHART | 0–1 | L | N | – | M D | 1.00 |  |
| 175 | COPY.PICTURE | 0 | L | – | – | M | 1.00 |  |
| 176 | `7` *(placeholder)* | 1–4 | – | – | – | X | 1.00 | no file number |
| 177 | ATAN2 | 2 | N | N | N |  | 1.00 |  |
| 178 | ASIN | 1 | N | N | – |  | 1.00 |  |
| 179 | ACOS | 1 | N | N | – |  | 1.00 |  |
| 180 | CHOOSE | 2–14 | NTLRCA | NA | NTLRECA |  | 1.00 |  |
| 181 | HLOOKUP | 3 | NTL | NTL | NRA |  | 1.00 |  |
| 182 | VLOOKUP | 3 | NTL | NTL | NRA |  | 1.00 |  |
| 183 | LINKS | 0–1 | A | T | – | M | 1.00 |  |
| 184 | ALERT | 2 | L | T | N | M | 1.00 |  |
| 185 | INPUT | 2–3 | NTLRECA | T | NT? | M | 1.00 |  |
| 186 | BEEP | 0 | L | – | – | M | 1.00 |  |
| 187 | ISREF | 1 | L | NTLREA | – |  | 1.00 |  |
| 188 | GET.FORMULA | 1 | T | R | – | M | 1.00 | first TR |
| 189 | GET.NAME | 1 | T | T | – | M | 1.00 |  |
| 190 | LEGEND | 0–1 | L | L | – | M | 1.00 |  |
| 191 | PREFERRED | 0 | L | – | – | M | 1.00 |  |
| 192 | HSCROLL | 1 | L | N | – | M | 1.00 | args 1–2, others L |
| 193 | VPAGE | 1 | L | N | – | M | 1.00 |  |
| 194 | SET.VALUE | 2 | L | R | NTLEA | M | 1.00 |  |
| 195 | FORMULA.ARRAY | 1 | L | NTLA | – | M | 1.00 | args 1–2, first NTLA?, others TR? |
| 196 | FORMULA.FILL | 1 | L | NTLA | – | M | 1.00 | args 1–2, first NTLA?, others TR? |
| 197 | SELECT.CHART | 0 | L | – | – | M | 1.00 |  |
| 198 | HPAGE | 1 | L | N | – | M | 1.00 |  |
| 199 | VLINE | 1 | L | N | – | M | 1.00 |  |
| 200 | MOVE | 2–3 | L | N? | NT? | M | 1.00 |  |
| 201 | SIZE | 2–3 | L | N? | NT? | M | 1.00 |  |
| 202 | FULL | 0–1 | L | L | – | M | 1.00 |  |
| 203 | `8` *(placeholder)* | 1 | – | – | – | X | 1.00 | no file number |
| 204 | `9` *(placeholder)* | 0–1 | – | – | – | X | 1.00 | no file number |
| 205 | CANCEL.COPY | 0 | L | – | – | M | 1.00 |  |
| 206 | UNLOCKED.NEXT | 0 | L | – | – | M | 1.00 |  |
| 207 | UNLOCKED.PREV | 0 | L | – | – | M | 1.00 |  |
| 208 | CHANGE.LINK | 0–2 | – | T | T | M D | 1.04 |  |
| 209 | DIALOG.BOX | 1 | NL | TR | – | M | 1.5 |  |
| 210 | ADD.MENU | 2 | N | N | R | M | 1.5 |  |
| 211 | ADD.COMMAND | 3 | N | N | NTR | M | 1.5 |  |
| 212 | ADD.BAR | 0 | N | – | – | M | 1.5 |  |
| 213 | SHOW.BAR | 0–1 | L | N | – | M | 1.5 |  |
| 214 | ENABLE.COMMAND | 4 | L | N | NTL | M | 1.5 |  |
| 215 | CHECK.COMMAND | 4 | L | N | NTL | M | 1.5 |  |
| 216 | RENAME.COMMAND | 4 | L | N | NT | M | 1.5 |  |
| 217 | DELETE.MENU | 2 | L | N | NT | M | 1.5 |  |
| 218 | DELETE.COMMAND | 3 | L | N | NT | M | 1.5 |  |
| 219 | GET.BAR | 0 | N | – | – | M | 1.5 |  |
| 220 | SLN | 3 | N | N | N |  | 1.5 |  |
| 221 | SYD | 4 | N | N | N |  | 1.5 |  |
| 222 | DDB | 4 | N | N | N |  | 1.5 |  |
| 223 | IPMT | 4–6 | N | N | N |  | 1.5 |  |
| 224 | PPMT | 4–6 | N | N | N |  | 1.5 |  |
| 225 | REPLACE | 4 | T | T | T |  | 1.5 |  |
| 226 | SUBSTITUTE | 3–4 | T | T | T |  | 1.5 |  |
| 227 | CODE | 1 | N | T | – |  | 1.5 |  |
| 228 | FIND | 2–3 | N | T | T |  | 1.5 |  |
| 229 | MDETERM | 1 | N | A | – |  | 1.5 |  |
| 230 | MINVERSE | 1 | A | A | – |  | 1.5 |  |
| 231 | MMULT | 2 | A | A | A |  | 1.5 |  |
| 232 | LOG | 1–2 | N | N | N |  | 1.5 |  |
| 233 | *(unused)* | | | | | | | |
| 234 | WAIT | 1 | L | N | – | M | 1.5 |  |
| 235 | SPLIT | 0–2 | L | N? | N? | M | 1.5 |  |
| 236 | FREEZE.PANES | 0–1 | L | L | – | M | 1.5 |  |
| 237 | GET.CELL | 1–2 | NTLA | N | R | M | 1.5 |  |
| 238 | GET.DOCUMENT | 1–2 | NTLA | N | T | M | 1.5 |  |
| 239 | GET.WINDOW | 1–2 | NTLA | N | T? | M | 1.5 |  |
| 240 | GET.WORKSPACE | 1 | NTL | N | – | M | 1.5 |  |
| 241 | GET.DEF | 1–2 | T | T | T | M | 1.5 |  |
| 242 | FILES | 0–2 | A | T | T | M | 1.5 |  |
| 243 | DIRECTORY | 0–1 | T | T | – | M | 1.5 |  |
| 244 | QUIT | 0 | L | – | – | M | 1.5 |  |
| 245 | *(unused)* | | | | | | | |
| 246 | *(unused)* | | | | | | | |
| 247 | *(unused)* | | | | | | | |
| 248 | *(unused)* | | | | | | | |
| 249 | *(unused)* | | | | | | | |
| 250 | *(unused)* | | | | | | | |
| 251 | CHAR | 1 | T | N | – |  | 1.5 |  |
| 252 | LOWER | 1 | T | T | – |  | 1.5 |  |
| 253 | UPPER | 1 | T | T | – |  | 1.5 |  |
| 254 | PROPER | 1 | T | T | – |  | 1.5 |  |
| 255 | *(unused)* | | | | | | | |
| 256 | LEFT | 1–2 | T | T | N |  | 1.5 |  |
| 257 | RIGHT | 1–2 | T | T | N |  | 1.5 |  |
| 258 | EXACT | 2 | L | T | T |  | 1.5 |  |
| 259 | TRIM | 1 | T | T | – |  | 1.5 |  |
| 260 | ISERR | 1 | L | NTLE | – |  | 1.5 |  |
| 261 | ISTEXT | 1 | L | NTLRE | – |  | 1.5 |  |
| 262 | ISNUMBER | 1 | L | NTLRE | – |  | 1.5 |  |
| 263 | ISBLANK | 1 | L | NTLRE | – |  | 1.5 |  |
| 264 | T | 1 | T | NTLR | – |  | 1.5 |  |
| 265 | N | 1 | N | NTLR | – |  | 1.5 |  |
| 266 | DATEVALUE | 1 | N | T | – |  | 1.5 |  |
| 267 | TIMEVALUE | 1 | N | T | – |  | 1.5 |  |
| 268 | CLEAN | 1 | T | T | – |  | 1.5 |  |
| 269 | COUNTA | 1–14 | N | NREA | NREA |  | 1.5 |  |
| 270 | PRODUCT | 1–14 | N | NRA | NRA |  | 1.5 |  |
| 271 | FACT | 1 | N | N | – |  | 1.5 |  |
| 272 | DPRODUCT | 3 | N | R | NTR |  | 1.5 |  |
| 273 | ISNONTEXT | 1 | L | NTLRE | – |  | 1.5 |  |
| 274 | TRUNC | 1 | N | N | – |  | 1.5 |  |
| 275 | ISLOGICAL | 1 | L | NTLRE | – |  | 1.5 |  |
| 276 | DCOUNTA | 3 | N | R | NTR? |  | 1.5 |  |
| 277 | STDEVP | 1–14 | N | NRA | NRA |  | 1.5 |  |
| 278 | DSTDEVP | 3 | N | R | NTR |  | 1.5 |  |
| 279 | VARP | 1–14 | N | NRA | NRA |  | 1.5 |  |
| 280 | DVARP | 3 | N | R | NTR |  | 1.5 |  |
| 281 | *(unused)* | | | | | | | |
| 282 | *(unused)* | | | | | | | |
| 283 | DELETE.BAR | 1 | L | N | – | M | 1.5 |  |
| 284 | FOR | 3–4 | C | T | N | M | 1.5 |  |
| 285 | WHILE | 1 | C | L | – | M | 1.5 |  |
| 286 | BREAK | 0 | C | – | – | M | 1.5 |  |
| 287 | NEXT | 0 | C | – | – | M | 1.5 |  |
| 288 | REFTEXT | 1–2 | T | R | L | M | 1.5 |  |
| 289 | TEXTREF | 1–2 | R | T | L | M | 1.5 |  |
| 290 | INDIRECT | 1–2 | NTLEA | T | L | V | 1.5 |  |
| 291 | CELL | 1–2 | NTL | T | R | V | 1.5 |  |

---

## Appendix B. Workspace files

A workspace file (Finder type `XLIN`, creator `XCEL`; on the 1.00 program disk it is called `Resume Excel`;
1.01 takes the name from item 113 of its string list `STR#` 256, `Reprendre Excel` in French) records what
Excel had open, so that opening it reopens the same documents in the same windows with the same calculation
settings. It holds no cell data. Written by `64.00` and read by `64.01` in Excel 1.00 (`63.00`, `63.01` in
1.03, which build the folder string differently but write the same fields).

| Size | Meaning |
|-----:|---------|
| word | 0. Open refuses anything else (alert `0x2038`). |
| 20 bytes | Calculation settings, laid out as offsets 0x000–0x013 of the sheet state block (§6): maximum iterations, calculation mode, precision, reference style, unused word, maximum change (double), iteration. In 1.00 they are global, not per document. |
| word | 1 if a window was **zoomed** to fill the screen (the FULL command or the zoom box; global `g0c5a`), else 0. Open zooms the front window again (`17.0c`). The rectangles below are always the windows' normal size: Excel un-zooms a window before recording it. |
| … | The document windows, front to back, each as below, then a `0x00` byte |
| word | `n`, then `n` bytes: the **preferred chart format** (Chart > Set Preferred Format, applied by Gallery > Preferred). Set Preferred Format copies the active chart's formatting into a data block of its own (below; `5a.13` in 1.00). `n` = 0 when none has been set (1,640 bytes for `CLOCK.CH`'s format). |
| byte | 0 |
| 3 words | **Printer Setup** (dialog 288, PRINTER.SETUP): printer type, baud rate and port (values below). The sample's `0, 5, 0` is a Macintosh printer; `PRINTER.SETUP(2,2,3)` gives `1, 2, 1`. |
| byte | 0 |

Printer type:

| Value | Meaning |
|---:|---|
| 0 | Macintosh printer |
| 1 | TTY (serial) |

Baud rate:

| Value | Rate |
|---:|---|
| 0 | 300 |
| 1 | 600 |
| 2 | 1200 |
| 3 | 2400 |
| 4 | 4800 |
| 5 | 9600 |

Port:

| Value | Meaning |
|---:|---|
| 0 | Printer |
| 1 | Modem |

Window entry:

| Size | Meaning |
|-----:|---------|
| byte | Document kind, as in §5 (`0x20` for a chart) |
| Pascal string | Document name. Excel 1.00 writes the window title, which includes the volume for a document on another disk (`Ws:Clock saved`). |
| Pascal string | Folder: the volume name and any folder names, each followed by `:` |
| word | 0 unless the system has HFS (low-memory `FSFCBLen` above 0). Then it is the low word of the directory ID of the document's folder, which Excel asks the File Manager for (HFS Dispatch, `48.22` in 1.00): 2, the ID of a volume's root directory, for a document at the top of a floppy; 0 for one on the startup disk. |
| 4 words | Window rectangle, in global coordinates. Open fits it to the screen. |
| 41 bytes | Window state, as in §13. Absent for a chart. |
| word + `n` × 6 bytes | Selected areas, as in §13. Absent for a chart. |

Open reopens each document by its Finder type: `MCBN` through the chart reader, `XLBN` and `XLPG` through Open
(§3), anything else through the importer. The sample has one window: `Expenses` in `Excel Data:`, a worksheet,
rectangle 73, 14, 335, 503, selection A1.

**Preferred chart format.** Each chart document keeps its charts in a data block of its own, which the chart
code makes its data segment (`pushG g23cc | setDS`) and addresses by fixed offsets; the document points at it
from offset 202. Excel 1.00 makes the block 1,078 bytes long for a new chart (`0f.00`), sets up its fixed
fields (`0f.02`) and allocates the chart structures, point formats and strings in it as it needs them
(`30.11`), so it grows. Set Preferred Format makes a new block and copies the chart structures and titles of
the active chart into it, but not its series (`5a.14`); that block is the preferred format, and the chart
keeps its own. The block's first word is its own length, as for every block on Excel's heap, which the memory
manager (`0a.03`) moves like any other; the pointers inside are offsets in the block, so they stay valid when
it moves.

| Offset | Meaning |
|-------:|---------|
| 0 | Word: the block's length |
| 320 | A point-format record (§16.8) that the chart code fills for one series at a time (`3c.02`, `35.17`) |
| 334 | Chart title: a text object in its 60-byte in-memory form (§16.7), kind 1 (`0f.02`) |
| 394 | Category-axis title, kind 2 |
| 454 | Value-axis title, kind 3 |
| 514 | Three rectangles: the chart area, plot area and legend as last drawn (`31.03`, `57.00`). A click is tested against them for selection kinds 0–2 (`36.03`). |
| 538 | Word: the **selection**, the kind of chart object selected (below); −1 when nothing is (`0f.02`, `36.06`) |
| 540 | Word: which object of that kind (below) |
| 542 | Word: a second index for some kinds (below) |
| 544 | Byte: the number of selection handles drawn round the selection (`36.17`) |
| 545 | Byte: the style of those handles, from a table indexed by the selection kind |
| 546 | Word: the offset of the structure of the chart the selection belongs to, main or overlay |
| 548 | Rectangle: the outline of the selection. Excel frames it, or for line-like objects draws a line between its corners, depending on the handle style (`36.17`, `36.36`). |
| 556 | 4 bytes for each handle, up to eight: where it is drawn (`36.17`) |
| 652 | Two words: the offsets of the main chart's structure and the overlay's (0 when there is none). Save writes them last first, which is why a file has the overlay first (§16.1; `5d.00`). |
| 656 | Word: the number of charts |
| 658 | Word: the offset of the first series in the first series list (§16.9), made from the worksheet selection when the chart is made (`34.00`, `0f.02`); the series link through their offset 12 |
| 660 | Word: the same for the second series list |
| 662 | Word: the number of series in the first list |
| 664 | Word: the number in the second; Save adds the two for the series count (`5d.00`). Words 658–664 are 0 in a preferred format, which has no series. |

Selection kinds, with what Format > Patterns shows for each (`59.00` in 1.00). The kinds come from the click
code (`36.02`, `36.04`, `36.05`, `31.0a`) and the code that draws each chart type, which tests the click as it
draws (`39`, `53`–`56`, `3a`):

| Value | Selected | Words 540 and 542 | Patterns dialog |
|---:|---|---|---|
| −1 | Nothing | | |
| 0 | Chart area (Chart > Select Chart) | | Area |
| 1 | Plot area (Chart > Select Plot Area) | | Area |
| 2 | Legend | | Area |
| 3 | Value axis | | Axis |
| 4 | Category axis | | Axis |
| 5 | Chart title | | Area |
| 6 | Category-axis title | | Area |
| 7 | Value-axis title | | Area |
| 8 | Free text | Its index in the text list | Area |
| 9 | Data label | Series and point (§16.7, mem 4 and 5) | Area |
| 10 | Arrow | Its index in the text list | Arrow |
| 11 | Gridlines | Which axis's, major or minor (`3a.08`, `3a.0f`) | Line |
| 12 | A data series, or a pie slice | Series and point | Line and marker (line and scatter charts) or area |
| 13 | Drop lines | | Line |
| 14 | Hi-lo lines | | Line |
| 15 | A data point, in a line or scatter chart with one series (`53.03`) | Series and point | Marker |

Offsets 514–587 are run-time state, worked out as the chart is drawn and clicked. A preferred format holds
whatever its block had there, left over from memory: in `Resume Excel 2` offsets 514–545 hold four doubles
(−40, 40, 10, 10) rather than rectangles and a selection.

Each chart structure is the 392-byte form of the chart record (§16.5, the table's "mem" column). In `Resume
Excel 2` the table at 652 holds 666 and 1058, the main (scatter) and overlay (pie) charts of `CLOCK.CH`, and
the count is 2. The rest of the block is a heap: Excel allocates the chart structures, point formats, series
and strings in it as it needs them (`30.11`), so their places vary.
