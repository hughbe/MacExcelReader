# Excel's code

Where the [format specification](Spec.md) comes from in Excel's own code: how the code is organised and
numbered, which releases were read, and, section by section, the procedures each part of the spec was read from.
The behaviour documents ([Display](Display.md), [Calculation](Calculation.md), [SYLK](Sylk.md), [bugs and
quirks](ExcelBugs.md)) cite procedures in place; the spec does not, and this index keeps its sources.

## Segments and procedures

**Segments and procedures.** The p-code is divided into numbered **segments** (118 in Excel 1.00), each of
numbered **procedures** (from one to over a hundred). In Excel 1.00 the application's `PCOD` 2 resource holds
every segment, and `PCOD` 0 holds the table of where each starts together with Excel's global data: the function
table, built-in names, operator and pattern tables and so on. The interpreter loads a segment from disk when one
of its procedures is called and unloads the least recently used ones when memory runs short. The documents cite code
by these numbers: `47.0b` is procedure `0x0B` of segment `0x47`, Excel 1.00's Save, in hexadecimal. Procedures
change number between releases when one is inserted, so a citation says which release a number is from where it
matters (below). A name like `g0812` is a global variable, named by its offset in Excel's global data.

## Releases read

**Code.** The p-code and 68000 code in the applications' `PCOD` resources, disassembled (`tools/pcode.py`):

| Release | Disk | Used for |
|---------|------|----------|
| Excel 1.00 (1985) | "Excel Program" disk (400K MFS), in two copies whose code differs in 4 bytes (the copy protection) | Save (`47.0b`), Open (`46.06`), the formula compiler (`11.00`), recalculation (`25.01`), the TABLE command (`4b`), chart writer, reader and drawing code (`5d`, `5e`, `37`), workspace files (`64.00`, `64.01`), function table and names. Each was compared with its 1.01 and 1.03 counterparts instruction by instruction; the differences are noted where they matter. |
| Excel 1.01 (French) | `Microsoft Excel 1.01 [French].dsk` (800K HFS), the only 1.01 available | Every procedure, matched to its 1.00 counterpart. 1.01 keeps 1.00's segment and procedure numbers apart from a few procedures it inserts; where a number differs this index gives both (`46.06` in 1.00, `46.07` in 1.01). |
| Excel 1.03 | `Microsoft Excel 1.03.dsk` (800K HFS) | Save (`45.0c`), Open (`44.08`), the external-link reader (`72.01`), the formula compiler (`0f.00`), chart reader and writer (`5d`, `5c`), the counterparts of 1.00's `47.0b`, `46.06`, `73.01`, `11.00`, `5e` and `5d`. This index cites 1.00's numbers first. |
| Excel 1.04 | `Microsoft Excel 1.04.dsk` (800K HFS, from WinWorld) | Every procedure, matched to its 1.03 counterpart. 1.04 inserts two segments (`0x47`, `0x48`), so 1.03's `0x47`–`0x73` are `0x49`–`0x75` in 1.04 and lower numbers are unchanged apart from a few inserted procedures; this index gives 1.04's numbers where a section describes a 1.04 change. |
| Excel 1.06 | `Microsoft Excel 1.06.dsk` (800K HFS) | Every procedure, matched to its 1.04 counterpart. 1.06 keeps 1.04's segment and procedure numbers apart from a few procedures it inserts or drops. |
| Excel 1.5 (May 1988) | Excel 1.5 folder | Save (`47.0d`), format and token translation (`47.18`), function table and file-number table, version-18 rules. Each 1.5 change the spec names was checked against 1.06's code, the release before it. |

Code is cited as `segment.procedure` in hexadecimal, for example `0f.00`.

## Tables in the application

The tables the spec reads from Excel's global data, as offsets into `PCOD` 0:

| Table | 1.00 | 1.01 | 1.03 | 1.5 | Spec |
|-------|------|------|------|-----|------|
| Operator characters (pairs tokens with `+ - = > ^ / * & : < , < < >`) | `0xCFD` | `0xCD5` | `0xCE9` | | §15.1 |
| Function table (10-byte entries) | `0x1B8C` | | `0x1512` | `0x1C92` | §15.3 |
| Function names (Pascal strings; COUNT at `0x00B`, TRUE at 0) | `0x14FA` | `STR#` 257 | `STR#` 257, read into `g0812` | | §15.3 |
| Fill and line patterns (8 bytes each) | `0x254C` | | | | §16.4 |

## Procedures behind the spec

Each spec section's sources, with the words the citation stood beside. Numbers are Excel 1.00's unless a
release is given.

**Sources**

* `47.0b`: …differs in 4 bytes (the copy protection) Save
* `46.06`: …4 bytes (the copy protection) Save (47.0b), Open
* `11.00`: …protection) Save (47.0b), Open (46.06), the formula compiler
* `25.01`: …47.0b), Open (46.06), the formula compiler (11.00), recalculation
* `64.00`, `64.01`: …and drawing code (5d, 5e, 37), workspace files
* `46.06` in 1.00, `46.07` in 1.01: …where a number differs the spec gives both
* `45.0c`: …Excel 1.03 Microsoft Excel 1.03.dsk (800K HFS) Save
* `44.08`: …Microsoft Excel 1.03.dsk (800K HFS) Save (45.0c), Open
* `72.01`: …HFS) Save (45.0c), Open (44.08), the external-link reader
* `0f.00`: …44.08), the external-link reader (72.01), the formula compiler
* `47.0d`: …Excel 1.5 (May 1988) Excel 1.5 folder Save
* `47.18`: …1.5 folder Save (47.0d), format and token translation

**2.1 Releases**

* `0d.0d` in 1.01: …Its code asks for a string by number
* `44.0c`, `44.0d` in 1.04: …between saves and flushes the volume after saving
* `01.1a` in 1.06: …application: a SIZE resource for MultiFinder, one routine
* `15.2c`: …no longer reserving master pointers around Open (1.04's 15.2b
* `46.06`: …§4). Excel 1.00's Open accepts only version 17
* `46.07`: …1.01's
* `44.0e` in 1.04 and 1.06: …1.01's (46.07), 1.03's, 1.04's and 1.06's
* `47.0d`: …1.5 or greater"
* `47.10` in 1.5: …number-format check
* `0d.19`: …

**4.1 Embedded p-code segments**

* `73.01` in 1.00: …opened, the link reader

**5. Data area preamble**

* `46.07` in 1.01: …alert 0x106 if it differs from the reader's
* `46.06`: …worksheet" (alert 0x0159
* `45.0d` in 1.04 and 1.06: …names; 1.03 does the same. Excel 1.04's Save

**6. Sheet state block (630 bytes)**

* `25.0e`: …it as it is and the recalculation driver
* `73.0d` in 1.00, `72.0d` in 1.03: …something they depend on changes. The external-link reader
* `10.00`: …sets it: a new sheet starts with 0
* `23.02`: …flag bits make the routine rescan each row
* `46.06`: …0x80 for a table waiting to be calculated
* `14.08`: …0x02E word Pointer, used by the name code
* `14.08`: …0x030 word Used by the name code
* `1d.06`, `3f.00`: …a list of tables waiting to be calculated
* `27.0c`, `27.0f`, `3f.00`: …a second list of tables, used by recalculation
* `3f.04`: …FFFF Head of the list of all tables
* `1d.06`: …0x03C word Used when a table is unlinked
* `17.16`, `23.01`: …Pointer used by the print-title and print code
* `10.0b`: …the file's values until then (only New clears them
* `46.06`: …into a scratch buffer and drops them
* `09.1c`: …It also recomputes the font metrics at 0x248–0x254

**6.1 Grid dimensions**

* `22.08`: …routine

**6.2 Password hash**

* `24.01` in 1.00, `22.01` in 1.03: …algorithm is from the password dialog's handler

**7. Section 0 — number formats**

* `47.11`: …1.5 writes the 3-slot form to version-17 files
* `07.16` in 1.00: …Excel 1.00–1.06
* `0d.19`: …Excel 1.5
* `47.11`: …the text section alone (unless Save is recompiling it
* `0d.1a`: …Excel 1.5 also marks literals
* `07.17` in 1.00, `05.17` in 1.03: …
* `07.09` in 1.00, `05.09` in 1.03: …08, 1B Reserved. The format compiler
* `07.09`: …Excel 1.00 compiles a section at a time
* `07.09` in 1.00: …and a count of 0
* `07.10`: …the digits only turns thousands on. The assembler
* `0d.0b`, `0d.12`: …Excel 1.5's compiler
* `47.11`: …each format and keeps the first three slots

**8. Section 1 — external documents and row index**

* `16.00` in 1.00: …Excel clears them when it makes an entry
* `16.09`: …in 1.00), Open reads them into the entry
* `47.0b`: …the entry (16.09) and Save writes them back

**9. Section 2 — cell table**

* `46.06`: …every column, but not in the margin row (§6.1
* `46.06`: …on the cell type handle 0–6 only; Open
* `47.0b`: …reads cell data for types 2–6 and Save

**11. Section 4 — defined names**

* `47.0b` in 1.00, `45.0c` in 1.03: …so Save numbers the names in this order

**12.3 Formula record (types 5 and 6)**

* `27.06`: …and FF FE that it was taken off

**12.6 Array formulas and tables**

* `46.06` in 1.00, `46.07` in 1.01: …reads as it is
* `4b.06` in 1.00: …46.07 in 1.01) and the TABLE command builds
* `3f.00` in 1.00: …goes on a list that the table calculation
* `1d.07` in 1.01: …3f.00 in 1.00) takes and empties, and 1d.06
* `46.06`, `46.07` in 1.01: …Open links the table into a list again
* `3f.04`, `46.06` in 1.00, `46.07` in 1.01: …all tables; Open links every table it reads
* `4f.0e`: …together and keeps the old record for Undo
* `0e.01`: …0 of an input with bit 1 set
* `4f.0b`: …0e.01), and the next command clears bit 1
* `1d.01` in 1.00: …and the formula shows #REF! in its place
* `4b.05` in 1.00: …from memory it never set ([Excel bugs and quirks](ExcelBugs.md)

**13. Section 6 — windows**

* `18.01`: …and scrolling shifts the bytes with the rows
* `22.02`: …every window of the document that shows it
* `0f.40`: …lower panes, counting from 0. FREEZE.PANES stores it
* `18.01`: …while the panes are frozen scrolling stops there

**14. Section 7 — page breaks and print settings**

* `47.0b`: …table
* `46.06`: …file ("File error", alert 0x2038) on any difference
* `47.0b`, `46.07` in 1.01: …the list in both Save and Open

**15.2 Tokens with operands**

* `25.01`: …Excel rewrites a memory token by its result
* `11.09` in 1.00, `11.0a` in 1.01: …one area when a single value is needed
* `25.01`: …it cannot get memory for the area list
* `25.01` in 1.00: …rewrites the token's low five bits to 0F
* `46.06`, `46.07` in 1.01: …28, and Open turns 2F back into 2E

**15.3 Function numbers**

* `22.07`, `4a.05` in 1.03: …numbering (PROTECT.DOCUMENT is 24.07 and TABLE 4b.05 in 1.00
* `11.00` in 1.00: …accepts NAME?(…) only for functions with this flag

**15.4 How Excel compiles formula text**

* `11.00` in 1.00, `0f.00` in 1.03: …
* `11.00`: …Excel 1.01's compiler
* `1d.01`: …in its operator-character table) and in TABLE's text
* `04.0e` in 1.01: …settings' country code is 0, the United States
* `02.11` in 1.03: …from a string resource instead
* `47.19`: …samples lacks the exception (§2.3). Excel 1.5's Save

**16.2 Header (128 bytes)**

* `64.00` in 1.5: …102 categories (the word at chart header offset 4

**16.4 Common sub-records**

* `37.03` in 1.00: …twips, worked out when the chart is drawn
* `31.07`, `38.02`: …only value Excel 1.00 draws or lays out

**16.5 Chart record (378 bytes)**

* `0f.04` in 1.00: …hides the overlay's axes at the same time
* `0f.06`: …No effect: a new chart's record clears them
* `37.0a`: …the overlay plots against the main chart's axis
* `37.34`: …out again each time it draws the chart
* `0f.05` in 1.00: …both with kind 4, data label
* `33.10`: …label (0f.05 in 1.00), and the label code
* `37.2b`: …the categories between tick labels to place them
* `37.13` in 1.00: …series, worked out when the chart is drawn
* `0f.05` in 1.00: …whatever the file had (0 in a new chart
* `37.03` in 1.00: …the width before layout less the width after
* `37.03` in 1.00: …drawing size (§16.6) when it draws the chart
* `36.02`: …only then does a click find the legend
* `30.1b` in 1.00: …LEGEND and Chart > Add Legend set it
* `37.03` in 1.00: …the area being laid out divided by 70
* `37.3e` in 1.00: …0 when the axis has no tick labels
* `37.02`, `37.22`: …in twips, worked out from the crossing settings

**16.6 Chart-set header and trailer**

* `30.0f` in 1.00: …scroll bar, at 72 pixels to the inch
* `31.00`: …pixels. Excel scales the chart's fonts from them

**16.7 Text objects**

* `33.19` in 1.00: …box's twips position at mem 30–33 in step
* `33.19`: …44, not saved), twips at 46 and 48

**16.8 Data-point formats and the series word list**

* `37.13` in 1.00: …which the drawing code collects into a list
* `3c.01` in 1.00: …record by its byte 6
* `33.10`: …one series, the point (category) number, from 0
* `35.18` in 1.00: …series format). GALLERY.PIE stores it as a word
* `35.17`: …format a chart has tests the whole word
* `0f.12`: …13 Set when a format is made
* `33.10` in 1.00: …the series, made from built-in text object −2
* `33.00`: …for a pie the label code totals the series
* `0f.12`: …clears a point-format record when it makes one

**Appendix B. Workspace files**

* `63.00`: …64.00 and read by 64.01 in Excel 1.00
* `0a.01`: …the folder Excel started from. Excel 1.03's Quit
* `17.0c`: …else 0. Open zooms the front window again
* `5a.13` in 1.00: …formatting into a data block of its own (below
* `48.22` in 1.00: …which Excel asks the File Manager for (HFS Dispatch
* `0f.00`: …block 1,078 bytes long for a new chart
* `0f.02`: …fields
* `30.11`: …
* `5a.14`: …active chart into it, but not its series
* `0a.03`: …manager
* `3c.02`, `35.17`: …code fills for one series at a time
* `0f.02`: …in its 60-byte in-memory form (§16.7), kind 1
* `31.03`, `57.00`: …area, plot area and legend as last drawn
* `36.03`: …is tested against them for selection kinds 0–2
* `0f.02`, `36.06`: …chart object selected (below); −1 when nothing is
* `36.17`: …number of selection handles drawn round the selection
* `36.17`, `36.36`: …between its corners, depending on the handle style
* `36.17`: …handle, up to eight: where it is drawn
* `5d.00`: …is why a file has the overlay first (§16.1
* `34.00`, `0f.02`: …the worksheet selection when the chart is made
* `5d.00`: …Save adds the two for the series count
* `59.00` in 1.00: …with what Format > Patterns shows for each
* `36.02`, `36.04`, `36.05`, `31.0a`: …code
* `3a.08`, `3a.0f`: …11 Gridlines Which axis's, major or minor
* `53.03`: …a line or scatter chart with one series
* `30.11`: …and strings in it as it needs them
