# Excel 1.x and BIFF2

How the Excel 1.x for Macintosh document format ([the spec](Spec.md)) differs from BIFF2, the format Excel 2.x
wrote and the oldest BIFF version. BIFF2 is described in OpenOffice.org's [Microsoft Excel File Format
documentation](https://www.openoffice.org/sc/excelfileformat.pdf) (“OOo” below, cited by its section
numbers); Excel 1.x by this repository's spec (“§” below).

In short: the **formulas** are almost the same encoding, the **cell formatting bits** are close relatives, and
many sheet settings carry the same values, but the **container** is entirely different. BIFF2 is a stream of
self-describing little-endian records; an Excel 1.x file is a big-endian dump of Excel's memory, in fixed
sections that have to be read in order.

## Contents

1. [At a glance](#1-at-a-glance)
2. [Container](#2-container)
3. [Sheet settings](#3-sheet-settings)
4. [Cells](#4-cells)
5. [Cell formatting](#5-cell-formatting)
6. [Formulas](#6-formulas)
7. [Array formulas and tables](#7-array-formulas-and-tables)
8. [Defined names](#8-defined-names)
9. [External references](#9-external-references)
10. [Number formats](#10-number-formats)
11. [Windows](#11-windows)
12. [Page setup](#12-page-setup)
13. [Only in one format](#13-only-in-one-format)
14. [Converting Excel 1.x to BIFF2](#14-converting-excel-1x-to-biff2)

## 1. At a glance

| | Excel 1.x | BIFF2 |
|---|---|---|
| Byte order | Big-endian (§1) | Little-endian (OOo §1.3) |
| Structure | 24-byte header, two embedded code segments, a 630-byte state block, then eight sections in a fixed order (§3) | Records, each a 2-byte identifier and a 2-byte length, between BOF and EOF (OOo §2.4, §4.1.1) |
| Skipping unknown data | Not possible: sections have no type, and most have no length | Any record can be skipped by its length |
| Strings | Pascal strings in Mac Roman | Byte strings with an 8- or 16-bit length, in the CODEPAGE record's encoding (OOo §2.5.2, §5.17) |
| Cells | A dense grid of 6-byte entries for every row and column in use, values elsewhere (§9, §12) | One record per non-empty cell, holding its row and column (OOo §4.7) |
| Cell formatting | Bits in each cell entry; one font per sheet (§9) | A 3-byte attribute field in each cell record and an XF table (OOo §2.5.13, §5.115) |
| Formulas | Reverse-Polish tokens, with pointers patched by trailing fix-ups (§12.5, §15) | The same tokens, with indexes stored in place (OOo §3) |
| Number formats | Compiled token bodies (§7) | FORMAT records holding the format text (OOo §5.49) |
| Dates | Always the 1904 system (§9) | 1900 or 1904, by the DATEMODE record (OOo §5.28) |
| Charts | Separate files, in Microsoft Chart format (§16) | BIFF streams of their own (BOF type `0x20`, OOo §7) |

## 2. Container

**Header.** An Excel 1.x file starts with the magic word `0x5327`, the format version (17, or 18 from Excel 1.5
when needed) and the writing build's code signature (§4). Two segments of Excel's own interpreted code follow,
for reading values from closed files (§4.1); BIFF has nothing like them. A BIFF2 file starts with a BOF record
(identifier `0x0009`) whose version field is not used (OOo §5.8).

**Document kind.** The two formats use the same values: the kind byte after the Excel 1.x state block (§5) and
the BIFF2 BOF type are `0x10` for a worksheet, `0x20` for a chart and `0x40` for a macro sheet. On the Mac
the Finder type also says it (`XLBN`, `XLPG`; §1.1).

**Sections, not records.** After the state block come a section table and eight sections: number formats,
external documents and the row index, the cell table, external names, defined names, cell data, windows, and
page setup (§3). The section table gives each section's position as a 512-byte block number and an offset
(§1.2), but Excel's own Open ignores it and reads straight through, so the sections must be contiguous and in
order. Most have no length of their own; a reader has to understand each to find the next. The only
forward-compatibility hook is the extra header after the section table, whose length Excel 1.00 already reads
and skips (§5).

**Leftovers from memory.** Excel 1.x writes many structures by copying them out of memory, so files are full
of pointers and padding that mean nothing on disk (§1). BIFF2 records were designed as an interchange format,
and the OOo document marks only a few fields “not used”.

## 3. Sheet settings

Most of BIFF2's settings records match a field of Excel 1.x's state block (§6), often with the same values:

| Excel 1.x state block (§6) | BIFF2 record | Difference |
|---|---|---|
| `0x000` maximum iterations | CALCCOUNT (`0x0C`) | — |
| `0x002` calculation mode: 1, −1, 0 | CALCMODE (`0x0D`) | Same values: 1 automatic, −1 automatic except tables, 0 manual |
| `0x004` precision: 1 full, 0 as displayed | PRECISION (`0x0E`) | Same values |
| `0x006` reference style: 1 A1, 0 R1C1 | REFMODE (`0x0F`) | Same values |
| `0x00A` maximum change (a double) | DELTA (`0x10`) | — |
| `0x012` iteration enabled | ITERATION (`0x11`) | — |
| `0x01E` document protected | PROTECT (`0x12`) | — |
| `0x242` password hash | PASSWORD (`0x13`) | The same hash (§6.2, OOo §4.18.4) |
| `0x014`, `0x274`, `0x020` grid size | DIMENSION (`0x00`) | Excel 1.x sizes the whole cell table with it (§6.1); BIFF2 gives the first and last used row and column |
| `0x042` column widths | COLWIDTH (`0x24`), DEFCOLWIDTH (`0x55`) | Excel 1.x: 256 widths in pixels, 75 by default; BIFF2: ranges of columns in 1/256 of the width of `0` |
| `0x244`, `0x246` standard font and size | FONT (`0x31`) | Excel 1.x has one font for the whole sheet: a Toolbox font number and size. BIFF2 has a font list, with names and heights in twips |
| `0x24C` row height (cached) | DEFAULTROWHEIGHT (`0x25`) | Excel 1.x derives it from the font; rows cannot have their own height |

Excel 1.x also keeps settings with no BIFF2 record: the recalculation state (`0x018`, which BIFF2 has per formula
instead), the flag for functions newer than the base set (`0x01C`), and the count of defined names (`0x022`).
BIFF2's DATEMODE has no counterpart, since Excel 1.x always uses 1904 dates.

## 4. Cells

**Layout.** Excel 1.x stores a dense, row-major grid of 6-byte entries, `rows × columns` of them, so a cell's
position is implicit (§9). Numbers, text and formulas live in a separate section, in the same order as their
cells, found by a 20-bit offset packed into the entry (§12). Before the grid comes a row index, one word per
row, which is a size hint for the recalculation dependency tables and can be ignored (§8). BIFF2 writes a record
only for each cell in use, with its row and column, grouped into blocks of 32 rows with ROW records and an INDEX
record to find them (OOo §4.7).

**Cell types.** Each Excel 1.x type corresponds to a BIFF2 cell record:

| Excel 1.x type (§9) | Value | BIFF2 record |
|---|---|---|
| 0 Empty | Formatting only | BLANK (`0x01`) |
| 1 Integer | Signed 16-bit, in the entry itself | INTEGER (`0x02`), which is **unsigned** 16-bit |
| 2 Number | A double, in section 5 | NUMBER (`0x03`) |
| 3 Text | A Pascal string, in section 5 | LABEL (`0x04`) |
| 5, 6 Formula | A formula record, in section 5 | FORMULA (`0x06`), with STRING (`0x07`) for a text result |
| — | — | BOOLERR (`0x05`): a constant TRUE, FALSE or error. Excel 1.x's type 4 occurs in no sample. |

**Formula results.** Both formats keep a formula's last value in 8 bytes, and both mark a value that is not a
number by setting the high-order word of the double to `0xFFFF`, which makes it a NaN. In Excel 1.x that word
comes first, being big-endian, and byte 2 gives the kind, with the value in byte 4 (§12.4). In BIFF the marker
is in bytes 6–7, the kind in byte 0 and the value in byte 2 (OOo §5.50). The kinds are numbered alike: 0 text,
1 Boolean, 2 error. Excel 1.x keeps a text result after the expression; BIFF2 writes it in a STRING record
after the FORMULA. The error codes are identical (§12.4, OOo §2.5.6).

## 5. Cell formatting

BIFF2's 3-byte cell attribute field (OOo §2.5.13) and Excel 1.x's words 1 and 2 (§9) hold much the same
things:

| Attribute | Excel 1.x | BIFF2 |
|---|---|---|
| Number format | Word 1 bits 5–0: built-in formats, then the user formats | Byte 1 bits 5–0: an index into the FORMAT records |
| Alignment | Word 1 bits 8–6 | Byte 2 bits 2–0. Same values: 0 General, 1 left, 2 centre, 3 right, 4 fill |
| Borders | Word 2 bits 2–5: left, right, top, bottom | Byte 2 bits 3–6, in the same order |
| Locked, hidden | Word 2 bits 14, 15 | Byte 0 bits 6, 7 |
| Font | Bold and italic, word 2 bits 0 and 1, on the sheet's one font | Byte 1 bits 7–6: one of four FONT records |
| Shading | — | Byte 2 bit 7 |
| XF index | — | Byte 0 bits 5–0, with IXFE for an index over 62 |

Excel 1.x has no XF table, no row or column default formats (BIFF2's ROW attributes and COLUMNDEFAULT) and no
font colours. Word 2 bit 13, which marks a formula that uses an array-valued function, has no BIFF2 equivalent.

## 6. Formulas

The formula encoding is where the formats agree most: the spec calls Excel 1.x's the direct ancestor of BIFF's.
Both are reverse-Polish token arrays with a one-byte length (two bytes from BIFF3), and both use bits 6–5 of
tokens `0x20` and up for the reference, value and array classes (§15, OOo §3.2). Token numbers and sizes agree
wherever both formats have the token:

| Token | Excel 1.x (§15.2) | BIFF2 (OOo §3.4) | Operand bytes |
|---|---|---|---|
| `03`–`16` | Operators, parentheses, missing argument | The same | 0 |
| `17` | String | tStr | 1 + n |
| `19` | Attribute | tAttr | 2 (+ n + 1 for CHOOSE) |
| `1A`, `1B` | External sheet start, end | tSheet, tEndSheet | 7, 3 |
| `1C`–`1F` | Error, Boolean, integer, number | tErr, tBool, tInt, tNum | 1, 1, 2, 8 |
| `20` | Array constant | tArray | 6 |
| `21`, `22` | Function, fixed and variable argument count | tFunc, tFuncVar | 1, 2 |
| `23` | Name | tName | 7 |
| `24`, `25` | Cell, area | tRef, tArea | 3, 6 |
| `26`–`29` | Memory area, error, no memory, function | tMemArea, tMemErr, tMemNoMem, tMemFunc | 4, 4, 4, 1 |
| `2A`, `2B` | Deleted cell, area | tRefErr, tAreaErr | 3, 6 |
| `2C`, `2D` | Cell, area in a name definition | tRefN, tAreaN | 3, 6 |
| `2E`, `2F` | Memory function in a name definition, and its no-memory form | tMemAreaN, tMemNoMemN | 1, 1 |

The OOo overview table (OOo §3.4.7) gives tFunc and tFuncVar one more byte each in BIFF2 than its own token
descriptions do (OOo §3.7.1, §3.7.2); Excel 1.x uses the sizes in the descriptions.

Cell addresses are encoded the same way: a row word whose bits 15 and 14 mark a relative row and column, and a
column byte (§15.2, OOo §3.3.3). The attribute token has the same kinds with the same one-byte data: volatile
`01`, IF `02`, CHOOSE `04` with its jump table of one-byte offsets, skip `08` and SUM `10` (OOo §3.10.5). BIFF3
widened these to two bytes.

**Pointers and fix-ups.** The main difference. In memory, Excel 1.x's name and external-sheet tokens hold
pointers, and Save writes them as they are, appending the real values after the expression in token order: the
name's index, the external document's index, the elements of each array constant and the areas of each memory
token (§12.5). BIFF2 puts the name index in the token's first two bytes (OOo §3.9.1), which is exactly what
Excel 1.x does inside name definitions (§11). BIFF also appends data after the token array, for the same two
tokens and in the same formats (OOo §3.1.6):

* an array constant's elements: a column count byte (0 meaning 256), a row count word, then typed values, 1
  for a number and 2 for a string (§12.5, OOo §2.5.8). BIFF also has types 0, 4 and 16 (empty, Boolean,
  error), which no Excel 1.x file contains;
* a memory area's areas: a count word, then 6-byte ranges of first row, last row, first column and last
  column (§12.5, OOo §2.5.15).

**Tokens only Excel 1.x has.** `30`–`37` are working copies of reference tokens that should not reach a file.
`38` is Excel 1.5's in-memory form of a variable-count call, translated on Save. `39` and `3A` are Excel 1.5's
calls to functions numbered 256 and up (§15.2): a one-byte function number could not reach them, and BIFF
solved the same problem in BIFF4 by widening the number to two bytes. In BIFF2 `38` is tFuncCE, a macro-command
call, and BIFF5 reuses `39`–`3D` for tNameX and the 3D references; Excel 1.x calls macro commands with the
ordinary function tokens.

**Function numbers.** Numbers 0–55 mean the same in both. From 56 on they differ, because Excel 1.x put
COLUMN.WIDTH and FORMULA at 56 and 57 (§15.3). Excel 1.5 already numbers its functions in BIFF's order
internally and translates on Save and Open (§2.2), so its internal number is the BIFF one: PV is 56 internally
and in BIFF, and 58 in an Excel 1.x file. Appendix A of the spec lists every Excel 1.x number.

**Dialog-box form.** Bit 7 of a variable-count call's argument count marks a command's dialog-box form, as in
`RUN?()`, in Excel 1.x (§15.2) and in BIFF4 and later (OOo §3.7.2).

**References in names.** BIFF2 stores relative references in names as signed offsets in tRefN and tAreaN
(“method [B]”, OOo §3.3.1). Excel 1.x uses the same token numbers in definitions; every name in the samples
refers to absolute cells, so whether their relative parts are offsets is not known.

## 7. Array formulas and tables

**Array formulas.** In both formats every cell of the range has a formula of one token, `01`, and the shared
expression is kept once, with the top-left cell. In BIFF2 the token holds the top-left cell's row word and
column byte (OOo §3.10.1), and an ARRAY record after that cell's FORMULA holds the range, a recalculation flag
and the expression (OOo §5.4). In Excel 1.x the token's operand is a pointer; each other cell's trailer holds
the same row word and column byte that BIFF2 puts in the token, and the top-left cell's trailer holds the
marker word `0x4000`, the range, the expression and its fix-ups (§12.6).

**Tables.** Table cells use token `02` in both. Excel 1.x keeps an 18-byte record in the top-left cell's
trailer: the range, up to two input cells with deleted marks, and flags for whether there are two inputs and
whether the first is the row input (§12.6). BIFF2 has DATATABLE for a table with one input, with a flag for row
or column input, and DATATABLE2 for two (OOo §5.24, §5.25).

## 8. Defined names

The entries hold the same parts in a different order, and both repeat the definition's length after it:

| Excel 1.x (§11) | BIFF2 DEFINEDNAME (OOo §5.33) |
|---|---|
| name length, definition length, flags | flags, macro kind, shortcut key, name length, definition length |
| shortcut key and macro kind, only for a macro | |
| definition, copy of its length, name | name, definition, copy of its length |

The flags differ: Excel 1.x has `0x04` for a reference, `0x20` for an array-valued definition and `0x40` for a
macro, and a macro kind of `0x40` for a command or `0x80` for a function. BIFF2 has `0x02` for a macro and
`0x04` for a complex function, and kinds 1 (function) and 2 (command).

Excel 1.x stores the names in case-insensitive alphabetical order and ends the list with a zero byte; a name's
place in that order is the index name tokens use. The array constants of every definition follow the list, name
by name (§11). A new Excel 1.x sheet always has five names with empty definitions (`Criteria`, `Database`,
`Print_Area`, `Print_Titles`, `Recorder`), translated in localised releases.

## 9. External references

Excel 1.x keeps one table of external documents, each a Pascal string (§8), and a list of the names used from
each (§10). Token `1A` refers to a document by its 1-based place in the table, and a name token between `1A` and
`1B` to a name by its place in that document's list. BIFF2 has EXTERNCOUNT, then an EXTERNSHEET record for each
document, each followed by its EXTERNNAME records (OOo §5.39–§5.41).

Document names differ. Excel 1.x stores the name or Mac path as it is (`::sampler files - current 1/13:Files:DATA`).
BIFF encodes it with control characters: `01` before an encoded path, `02` for the document itself, and
`03` between folders (OOo §2.5.9). Excel 1.x's equivalent of `02` is an empty name (Excel 1.03 and later), and it
also has the one-character name `00` for whichever sheet is active, as in `=!A1`.

## 10. Number formats

**Storage.** BIFF2 keeps each format as its text (`0.00%`) in a FORMAT record, built-in formats included, with
BUILTINFMTCOUNT saying how many are built in (OOo §5.12, §5.49). Excel 1.x keeps only the user-defined
formats, each already compiled: section offsets and lengths, and token bodies (§7). It writes the number of
built-in formats (19) before them and renumbers a file's format indexes on Open if its own count differs, which
is what BUILTINFMTCOUNT is for.

So a converter compiles and decompiles: `ExcelNumberFormat.ToText` turns an Excel 1.x format into its text, and
`ExcelNumberFormat.Parse` compiles text the way Excel 1.00 or 1.5 does.

**Indexes.** Both store a cell's format as a 6-bit index, built-in formats first, then the user formats in
order.

**Features.** Excel 1.00 formats have three sections, negative, zero and positive. Excel 1.5's version 18 adds a
text section, scaling by trailing commas, `@` and colours (§2.1).

## 11. Windows

Excel 1.x saves every open window of the document (§13); BIFF2 one window, as WINDOW1, WINDOW2, PANE and
SELECTION records (OOo §5.75, §5.93, §5.109, §5.110).

| Excel 1.x window (§13) | BIFF2 |
|---|---|
| Window rectangle, in screen pixels | WINDOW1: position and size in twips, and a hidden flag |
| State bytes 26–28: display formulas, gridlines, headings | WINDOW2 bytes 0–2, in the same order |
| State byte 121 (version 18): panes frozen | WINDOW2 byte 3 |
| — | WINDOW2 byte 4: show zero values |
| Top row and left column of the upper-left pane (state offsets 3, 11) | WINDOW2: first visible row and column |
| State bytes 127–128 (version 18): automatic colour, or one of the eight Mac colours | WINDOW2: automatic flag, or an RGB colour |
| Split positions in pixels (state offsets 37, 39); frozen split row and column (offsets 123, 125) | PANE: split positions in twips, or rows and columns when frozen, and the active pane |
| Active cell, the index of the active area, and the selected areas | SELECTION: active pane, active cell, index of the active area and the areas, in that order |

## 12. Page setup

Excel 1.x keeps page breaks, page header and footer, margins and print options together in section 7 (§14);
BIFF2 has a record for each.

* **Page breaks.** Excel 1.x stores the last row or column **before** each break (§14); BIFF2's
  HORIZONTALPAGEBREAKS and VERTICALPAGEBREAKS store the first one **after** it (OOo §5.57, §5.108), one more.
  Excel 1.x writes column breaks as bytes and row breaks as words, in chunks of up to five.
* **Header and footer.** Strings in both, with Excel's `&` codes; Excel 1.x's defaults are `&f` and `Page &p`.
* **Margins.** Doubles in inches in both: the Excel 1.x page-setup record's offset 62, and BIFF2's LEFTMARGIN,
  RIGHTMARGIN, TOPMARGIN and BOTTOMMARGIN.
* **Print headings and gridlines.** Bytes 97 and 101 of the Excel 1.x record; PRINTHEADERS and PRINTGRIDLINES.
* **Printer.** Excel 1.x keeps the Mac's own print record (`TPrint`) and settings for a serial “TTY” printer,
  which BIFF2 lacks.

## 13. Only in one format

Only Excel 1.x has:

* the embedded code for reading closed files (§4.1) and the code signature that goes with it;
* the function count and, in Excel 1.00, the names of extra functions (§5, §14);
* the recalculation state and dependency hints (§6, §8);
* more than one saved window;
* the TTY printer settings;
* the workspace file, `XLIN` (Appendix B), which BIFF only gained in BIFF3 (BIFF3W).

Only BIFF2 has:

* a font list (FONT, FONT2, FONTCOLOR), shading and row heights;
* row and column default formats (ROW, COLUMNDEFAULT) and an XF table;
* cell notes (NOTE);
* the 1900 date system (DATEMODE) and a code page (CODEPAGE);
* file encryption (FILEPASS), and BACKUP, WINDOWPROTECT and the show-zero-values switch;
* charts in the same file format as sheets.

## 14. Converting Excel 1.x to BIFF2

A reader of either format that wants to write the other needs, beyond moving fields between records:

* byte order: swap every word, long and double;
* strings: Mac Roman, which BIFF2 supports with CODEPAGE `0x8000` (Apple Roman);
* dates: DATEMODE 1;
* cells: walk the dense grid, write a record for each cell that is not an unformatted blank, and turn negative
  Excel 1.x integers into NUMBER records;
* formatting: one FONT for the sheet's font and its bold and italic variants, and an XF for each combination
  used;
* number formats: decompile each user format to text and write all the built-in ones first;
* formulas: copy the token array, put each name's and external document's index into its token, move array
  elements and memory areas to the end, renumber functions 56 and up, and rewrite Excel 1.5's `39` and `3A`
  calls, which BIFF2 cannot express;
* array formulas and tables: write the shared expression as ARRAY, or the table as DATATABLE or DATATABLE2;
* page breaks: add one.

The library does the Excel 1.x half: `ExcelDocument` gives cells, formulas as text and tokens, formats, names,
windows and page setup as objects ([usage guide](Usage.md)).
