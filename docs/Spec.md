# Microsoft Excel 1.x for Macintosh — document file format

Reverse-engineered from `Microsoft Excel 1.03.dsk` (Excel 1.03, 800K HFS disk).

Sources of truth, in order of weight:

1. The application's **Save** routine (p-code segment `0x45`, procedure `0x0C`) and
   **Open** routine (segment `0x44`, procedure `0x08`), plus the external-link
   reader (segment `0x72`, procedure `0x01`), disassembled from the `PCOD`
   resources of the application.
2. The five documents on the disk: `Worksheet1`, `Expenses` (worksheets) and
   `COMMANDS`, `DB.FORM`, `REPLACE` (macro sheets).

Every structural claim below was checked against all five files. A walker built
from this description consumes every section of every sample byte-exactly
(1,285 + 147 + 189 + 648 cell slots, 472 formulas). Where something is inferred
rather than read from code or confirmed in data, it is marked **(inferred)** or
**(unknown)**.

Charts use a **different format** (Microsoft Chart, file type `MCBN`),
described in §16 from Excel's chart writer, chart reader and chart dialogs.

---

## 1. Conventions

* All multi-byte integers are **big-endian**.
* `byte` = 8 bits, `word` = 16 bits, `long` = 32 bits, `double` = IEEE 754
  64-bit.
* *Pascal string* = one length byte followed by that many characters
  (Mac Roman).
* Rows and columns are **zero-based** on disk. Row limit 16,384 (14 bits),
  column limit 256.
* The whole document lives in the **data fork**. The resource fork is empty.

### 1.1 Finder types

| Type   | Creator | Meaning                         |
|--------|---------|---------------------------------|
| `XLBN` | `XCEL`  | Worksheet                       |
| `XLPG` | `XCEL`  | Macro sheet                     |
| `XLIN` | `XCEL`  | "Resume Excel" workspace file (not this format; 35 bytes on the sample disk) |
| `MCBN` |         | Chart, saved in Microsoft Chart format (§16) |

Save dispatches on the document kind: charts go to the chart writer; otherwise
the file type is `XLPG` when the document kind byte (§4) is `0x40`, else
`XLBN`.

### 1.2 Stream positions

The writer is a 512-byte buffered stream. A stream position is stored as a
long made of two words, **(block number, offset within block)**:

```
absolute offset = hi_word * 512 + lo_word
```

e.g. `00 02 01 F0` → 2 × 512 + 0x1F0 = 0x5F0.

### 1.3 Packed pointers ("seg:off")

Excel 1.x's p-code uses 8086-style 20-bit pointers: `address = seg * 16 + off`
with `off` in 0…15 once normalised. Several structures were dumped from memory
with such pointers still in them. Those are **junk on disk** except in the cell
table, where the writer replaces them with section-relative offsets (§7).

---

## 2. Overall layout

```
+0x00  File header                     24 bytes
+0x18  Embedded p-code segment 1       header.len1 bytes
       Embedded p-code segment 2       header.len2 bytes
       Capacity word                   2 bytes          <- header.dataOffset
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

Sections are contiguous and in this order; the section table gives the start of
each.

---

## 3. File header (24 bytes)

| Offset | Size | Value in samples | Meaning |
|-------:|-----:|------------------|---------|
| 0  | word | `0x5327` | Magic. Open rejects anything else. |
| 2  | word | `0x0011` (17) | Format version. Open rejects anything else. |
| 4  | word | `0xAB27` | Code signature of the writing build (see §3.1). |
| 6  | 3 words | 0, 0, 0 | Written as zero; never read. |
| 12 | word | 2 | Number of embedded code segments. |
| 14 | word | `0x0326` / `0x038A` | Length of segment 1. |
| 16 | word | `0x0018` | Offset of segment 1 (always 24). |
| 18 | word | `0x0016` | Length of segment 2. |
| 20 | word | 24 + len1 | Offset of segment 2. |
| 22 | word | 24 + len1 + len2 | **Offset of the data area.** Open seeks here directly. |

### 3.1 Embedded p-code segments

Each Excel 1.x document carries two segments of the application's own
interpreted code (segments `0x6F` and `0x70` of the running program), copied
verbatim out of memory.

They exist for **external references to closed documents**. When Excel needs a
value from a file it has not opened, the link reader looks at header word 4:

* `0xAB27` (same build) → use the application's own copy of segment `0x6F`.
* `0x13E8`, or any `0xABxx` → load segment 1 *from the file*, patch its single
  import to point back at the host (segment `0x72`, procedure `0x0A`), and call
  its procedure 0 with message 25 to fetch the value.

So a document is readable for linking by Excel builds that do not otherwise
know its exact layout. A reader that only wants the data can **skip both
segments**.

Segment image layout (for completeness):

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 | word | Offset of the constant/import table end (`2 × nProcs + 8` — exactly one import word precedes it) |
| 2 | word | LRU stamp from the running interpreter — garbage |
| 4 | word | Number of procedures (3 in segment 1, 1 in segment 2) |
| 6 | word × n | Offset of each procedure |
| … | word | Import: `0x720A` = segment `0x72`, procedure `0x0A` |
| … | | p-code |
| last 10 bytes | | Header of whatever heap block followed in memory — garbage |

The stored length is the heap block size, which includes a 10-byte block header
that is *not* written, so each image over-runs 10 bytes into the next block.

Four samples carry a 0x326-byte segment 1 that is byte-identical (apart from
the stamp) to segment `0x6F` of the 1.03 application. `Expenses` carries a
0x38A-byte one. Disassembled side by side they are the same three procedures
with the same logic; the only difference is that the 1.03 copy uses the
p-code compiler's shared-sequence opcodes to fold repeated instruction runs,
and the `Expenses` copy has those runs written out in full. So `Expenses` was
saved by an earlier build compiled without that size optimisation, which
nevertheless used the same code signature `0xAB27`.

---

## 4. Data area preamble

At `header.dataOffset`:

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 | word | **Number of functions known to the writing build**: 208 (`0x00D0`) in all samples, which is the length of the application's function table (208 ten-byte entries) and of its name list `STR#` 257. Open refuses the file if this is larger than its own count, since formulas could then use function numbers it does not have. |
| 2 | 630 bytes | Sheet state block (§5) |
| 632 | byte | **Document kind**: `0x10` worksheet, `0x40` macro sheet. (`0x20` is the in-memory kind for charts, which never reach this writer.) |
| 633 | 8 longs | Section table: stream positions (§1.2) of sections 0…7. Written last, by seeking back. Open *skips* this table and reads sequentially; it is redundant but reliable. |
| 665 | word | `n` = number of extra header bytes that follow. 0 in the four older samples, 1 in `Worksheet1` (saved by 1.03). |
| 667 | n bytes | Extra header. Byte 0, when present, is a system/locale code (high byte of an application global); Open warns with alert 0x106 if it differs from the reader's. Open skips any further bytes, so this is the format's forward-compatibility hook. |

Section 0 begins immediately after the extra header.

---

## 5. Sheet state block (630 bytes)

A raw dump of a contiguous run of application globals. Offsets are relative to
the start of the block.

| Offset | Size | Samples | Meaning |
|-------:|-----:|---------|---------|
| 0x000 | word | 100 | Maximum iterations (Calculation dialog; 1…32767) |
| 0x002 | word | 1 | Calculation mode. From the dialog's radio order: 1 = Automatic, 0 = Automatic except tables, −1 = Manual **(mapping inferred from dialog item order)** |
| 0x004 | word | 1 | Precision: 1 = full precision, 0 = precision as displayed (results are rounded through the cell format when 0) |
| 0x006 | word | 1 | Reference-style flag, saved and restored around formula parsing; presumably 1 = A1, 0 = R1C1 **(inferred)** |
| 0x008 | word | 0 | Unused — no code in the application reads or writes it |
| 0x00A | double | 0.001 | Maximum change (iteration) |
| 0x012 | word | 0 | Iteration enabled (0 / 1) |
| 0x014 | word | | **Last used row** (zero-based) |
| 0x016 | word | 255 | Index of the last column whose width is not zero. Starts at 255; the column-width routine lowers it when the last such column is set to width 0 and raises it when a higher column is given a width. (A column width of 0 is how a column is hidden.) |
| 0x018 | word | 0 / 1 | Recalculation state: 0 clean, 1 needs recalculation, 2 a distinct pending state the code never downgrades to 1 |
| 0x01A | word | 0 | Pending-work flag: when set, the next pass re-applies an operation to the whole sheet (rows 0–16383, columns 0–255) and clears it |
| 0x01C | word | 0 | Flag that selects which constant is written as the capacity word (§4) |
| 0x01E | word | 0 / 1 | **Document protected.** Selects the "Protect Document…" / "Unprotect Document…" menu text and is checked by ~50 editing routines. 1 only in `DB.FORM`. |
| 0x020 | word | 0 / 1 | **Non-empty flag**: 0 for a sheet with no cell table, else 1 |
| 0x022 | word | | Number of defined names that have a definition (0 for a new sheet, 6 in `REPLACE`) |
| 0x024 | word | | Number of those whose definition passes a "simple reference" test (5 in `REPLACE`) |
| 0x026 | long | | memory pointer — garbage |
| 0x02A | long | | head of the in-memory name list — garbage |
| 0x02E | 4 longs | | memory pointers / `FFFF` markers — garbage |
| 0x03E | word | 0 / 1 | Index of an external-document entry that link refresh skips **(inferred)**; 1 only in `DB.FORM` |
| 0x040 | word | 1 | Window counter for the document — set to 1 for a new sheet and updated as windows are opened and closed |
| 0x042 | 256 words | 75 | **Column widths**, columns 0…255. Default 75. Unit is presumably pixels **(inferred)**. |
| 0x242 | word | 0 | **Protection password hash**; 0 = none. Unprotect hashes the typed password and refuses on mismatch. See §5.2. |
| 0x244 | word | 3 | Standard font number (passed to `TextFont`; 3 = Geneva) |
| 0x246 | word | 10 | Standard font size (passed to `TextSize`) |
| 0x248 | word | 5 | Cached: half the font ascent |
| 0x24A | word | 12 | Cached: ascent + descent |
| 0x24C | word | 13 | Cached: default row height in pixels |
| 0x24E | word | 10 | Cached: font ascent |
| 0x250 | word | | Cached: packed font metric |
| 0x252 | word | 8 | Cached: width of `0` in bold |
| 0x254 | word | 7 | Cached: width of `0` in plain |
| 0x256 | long | | memory pointer — garbage |
| 0x25A | long | 0 | Head of the row page-break list in memory (rebuilt from §13) |
| 0x25E | long | 0 | Head of the column page-break list in memory (rebuilt from §13) |
| 0x262 | word | 0 | Cached: 1 if a print area is defined |
| 0x264 | word | 0 | Cached: pixel width of the print-title columns |
| 0x266 | long | | memory pointer — garbage |
| 0x26A | word | 0 | Cached: number of print-title rows |
| 0x26C | word | 0 | Cached print-titles value |
| 0x26E | 6 bytes | 0 | Cached print-titles range (first row word, last row word, first column byte, last column byte) |

Offsets 0x262–0x273 are all derived from the `Print_Area` and `Print_Titles`
names by one routine, which is re-run whenever a name or a column width
changes; a writer can leave them zero.
| 0x274 | word | | **Last used column** (zero-based) |

The cached font values at 0x248–0x254 are recomputed from the font number and
size by a single routine (`GetFontInfo` / `CharWidth`); a writer can copy them
from a sample.

### 5.1 Grid dimensions

```
nRows = word[0x014] + word[0x020]
nCols = word[0x274] + word[0x020]
```

These size the row index (§6.2) and the cell table (§7). An empty sheet has
`nRows = nCols = 0`.

### 5.2 Password hash

From the password dialog's handler (segment `0x22`, procedure `0x01`). For a
password of `n` characters `c[1…n]`:

```
h = 0
if n > 0:
    for b in c[n], c[n-1], …, c[1], n:        # characters last to first, then the length
        h = ((h << 1) & 0x7FFF) | ((h >> 14) & 1)   # rotate left within 15 bits
        h = h ^ b
    h = h ^ 0xCE4B
```

An empty password gives 0. This is the same hash later Excel versions use for
sheet protection.

---

## 6. Sections 0 and 1

### 6.1 Section 0 — number formats

| Size | Meaning |
|-----:|---------|
| byte | Number of built-in formats in the writing build (19 in all samples). If this differs from the reader's, Open renumbers cell format indices. |
| byte | `n` = number of user-defined formats (0 in all samples) |
| n × … | User-defined format entries |

Each user-defined entry (layout from code; no sample has one):

| Size | Meaning |
|-----:|---------|
| 3 bytes | Offsets of three sections within the body |
| 3 bytes | Lengths of those three sections |
| byte | Flag: 1 when the format has a single section, 0 when it has more |
| n bytes | Tokenised body, `n = max(offset[i] + length[i])` |

The first offset/length slot is the negative-number section and the other two
are the positive and zero sections (built-in format 5,
`$#,##0 ;($#,##0)`, has offsets `0C 00 00`, lengths `0E 0C 0C`)
**(slot order inferred from the built-ins)**.

Built-in formats are not stored in documents; they live in the application's
`FRMT` 256 resource in the same descriptor-plus-body form. Total formats are
capped at 64. The 19 built-ins, by index:

| # | Format | # | Format |
|--:|--------|--:|--------|
| 0 | General | 10 | m/d/yy |
| 1 | 0 | 11 | d-mmm-yy |
| 2 | 0.00 | 12 | d-mmm |
| 3 | #,##0 | 13 | mmm-yy |
| 4 | #,##0.00 | 14 | h:mm AM/PM |
| 5 | $#,##0 ;($#,##0) | 15 | h:mm:ss AM/PM |
| 6 | $#,##0.00 ;($#,##0.00) | 16 | h:mm |
| 7 | 0% | 17 | h:mm:ss |
| 8 | 0.00% | 18 | m/d/yy h:mm |
| 9 | 0.00E+00 | | |

The format names above are the standard Excel 1.x list (they are not stored as
text anywhere). The token meanings below come from the routine that turns a
tokenised format back into text for the Format Number dialog (segment `0x05`,
procedure `0x17`), and agree with every built-in body. Letters and separators
are taken from an international character table at run time, so the ones shown
are the US defaults.

| Token | Operands | Text produced |
|------:|----------|---------------|
| `01` | | `General` |
| `02` | 1 byte `c` | Literal character `c`. A backslash is written before it unless it is one of the characters that may appear unescaped. |
| `03` | 3 bytes `a b c` | Starts a number section; produces no text. `a` = number of integer digit positions (drives thousands-separator placement), `b` = decimal places, `(c − b) / 2` = number of percent signs. |
| `1C` | 3 bytes | Same, for a scientific section |
| `04` | 1 byte `n` | `n` × `#` |
| `06` | 1 byte `n` | `n` × `0` |
| `05` | | Decimal point; digit runs after it are written without separators |
| `07` | | Turns on the thousands separator, inserted every three integer positions |
| `08`, `1B` | | Reserved. The format compiler (segment `0x05`, procedure `0x09`) never emits them and the routines that read formats skip them. |
| `09` | 1 byte `n`, `n` chars | Quoted string `"…"` |
| `0A` | 1 byte `c` | A prefix character from the international table followed by `c` — the repeat-to-fill form **(inferred)** |
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

So `#,##0` is `03 04 00 00 | 07 | 04 03 | 06 01`: header, thousands on,
`###`, `0`.

The compiler confirms the header bytes: it writes the integer-position count,
the decimal count, and decimals plus twice the number of percent signs. It
also shows how the hour tokens are chosen — `h`/`hh` compile to `13`/`14`,
and are bumped by 12 to `1F`/`20` when an AM/PM designator follows — and that
`m` is compiled as minutes rather than month when it directly follows an hour
token.

### 6.2 Section 1 — external documents, then row index

**External document table**

| Size | Meaning |
|-----:|---------|
| word | 0 if the sheet has no external-document table. Otherwise `n × 12 + 8`. |
| word | `n` = number of external documents (present only if the previous word is non-zero) |
| n × … | One entry per external document |

Entry:

| Size | Meaning |
|-----:|---------|
| 6 bytes | First 6 bytes of the in-memory entry **(unknown; zero in the sample)** |
| Pascal string | Document name. A zero-length string means "this document". `DB.FORM` has one entry whose name is the single character `0x00` — apparently the "active sheet" target of `!NAME` references on a macro sheet **(inferred)**. |

**Row index** — immediately after, `nRows` words, one per row. Zero in every
sample.

In memory each row has a 4-byte slot that may point at a per-row auxiliary
table. Save writes that table's entry count (low 15 bits of its first word), or
0 if the row has none. Open, for a value greater than 2, allocates an *empty*
table with that capacity (`2n + 2` bytes) and does not read any contents. So
the word is only a size hint for a structure Excel rebuilds at run time; a
writer can emit 0 and a reader can ignore it.

The structure is the recalculation dependency index. A row's slot records
which rows hold formulas that refer to cells in this row; its top two bits tag
the form — 1 for a single row held inline, 2 for an inline pair, 3 for a
pointer to a table (count word, capacity word, then row numbers whose own top
two bits are flags). Open rebuilds it by calling the dependency routine once
per formula as the cell data is read. The transient bits 12–6 of cell word 2
(§7) belong to the same machinery: the routine that walks dependents tests
bits 12, 11, 10, 9 and 8 of each cell as it goes.

---

## 7. Section 2 — cell table

A dense, row-major grid: `nRows × nCols` entries of 6 bytes. The entry for
(row, col) is at index `row × nCols + col`. Row and column are therefore
implicit.

| Bits | Meaning |
|------|---------|
| word 0 | Value or offset, depending on type |
| word 1, bits 15–12 | Low 4 bits of the data offset (types 2, 3, 5, 6) |
| word 1, bits 11–9 | **Cell type** |
| word 1, bits 8–6 | Alignment: 0 General, 1 Left, 2 Center, 3 Right, 4 Fill. Only 0 occurs in the samples; the codes come from the importer's mapping of alignment letters C, G, L, R, X to 2, 0, 1, 3, 4. |
| word 1, bits 5–0 | Number format index (0 = General; 5 seen on currency cells, 10 on date cells) |
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

The style and border bits are confirmed by the importer, which maps the
letters D, I, L, R, T, B to bits 0, 1, 2, 3, 4, 5 respectively.

Cell types:

| Type | Meaning | word 0 / data |
|-----:|---------|---------------|
| 0 | Empty (formatting only) | word 0 is a stale memory pointer — ignore |
| 1 | Integer | word 0 = signed 16-bit value, stored inline. Bits 15–12 of word 1 are not meaningful (always `0xD` in the sample). |
| 2 | Number | 8-byte double in cell data |
| 3 | Text | Pascal string in cell data |
| 4 | **(not observed)** | No data in cell data |
| 5 | Formula, numeric result | Formula record in cell data (§8) |
| 6 | Formula, non-numeric result | Formula record in cell data (§8) |

For types 2, 3, 5 and 6 the cell's data is at

```
section5_start + (word0 << 4) + (word1 >> 12)
```

Data records appear in section 5 in the same order as their cells appear in the
table, with no gaps, so the offsets are strictly increasing.

Dates are serial numbers in the **1904 system** (29220 = 1 Jan 1984 in
`Expenses`) **(inferred)**.

---

## 8. Section 5 — cell data

### 8.1 Number (type 2)

8 bytes, IEEE double.

### 8.2 Text (type 3)

Pascal string (up to 255 characters).

### 8.3 Formula record (types 5 and 6)

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0  | 8 bytes | Current value (below) |
| 8  | 2 bytes | **(unknown)**, zero in samples |
| 10 | byte | Recalculation scratch — varies, ignore |
| 11 | byte | Column |
| 12 | word | Row in bits 13–0. Bit 14 mirrors "alignment is Fill" for the cell (set from bit 2 of the alignment code whenever alignment is applied). Bit 15 is a transient "needs fix-up" mark set while importing and cleared by a later pass. Both are zero in the samples. |
| 14 | byte | `cce` = length of the parsed expression |
| 15 | cce bytes | Parsed expression (§9) |
| … | Pascal string | Present only if the value is a string |
| … | | Trailing fix-ups (§8.5) |

### 8.4 Current value

* If bytes 0–1 are **not** `FF FF`: an IEEE double (cell type 5).
* If bytes 0–1 are `FF FF` (cell type 6), byte 2 selects:

| Byte 2 | Meaning | Where the value is |
|-------:|---------|--------------------|
| 0 | String | Bytes 4–7 are a stale pointer. The string follows the expression as a Pascal string. |
| 1 | Boolean | Byte 4: 0 = FALSE, 1 = TRUE |
| 2 | Error | Byte 4 = error code. `0x0F` observed; consistent with the later BIFF codes (0 #NULL!, 7 #DIV/0!, 15 #VALUE!, 23 #REF!, 29 #NAME?, 36 #NUM!, 42 #N/A) **(inferred)** |

### 8.5 Trailing fix-ups

In memory, some tokens hold pointers. Save writes the expression unchanged
(pointer bytes and all) and then appends replacements, in token order:

| Token | Appended | Meaning |
|-------|----------|---------|
| `0x23` name | word | 1-based index of the name in section 4 |
| `0x1A` external sheet | word | Index of the external document (from §6.2) |
| `0x20` array constant | array block (below) | **(not observed in samples)** |
| `0x26` memory area | word `n`, then `n` × 6-byte areas | The cached area list the token points at **(not observed in samples)** |

Array block: byte column count (0 means 256), word row count, then
`rows × cols` elements, each a type byte followed by either a Pascal string
(type 2) or 8 bytes (any other type).

### 8.6 Array formulas and tables

If the first token of the expression is `0x01` (array formula) or `0x02`
(TABLE), the cell does not own its expression: the token's three operand bytes
point at a shared record. On disk the trailer replaces the fix-ups of §8.5.
From code only — no sample contains one.

*Cell that is not the anchor* (its row/column differ from the shared record's
first row/column):

| Size | Meaning |
|-----:|---------|
| word | Anchor row |
| byte | Anchor column |

*Anchor cell*:

| Size | Meaning |
|-----:|---------|
| word | `0x4000` — marker; cannot be a row number |
| byte | 0 |
| 10 bytes | Shared record header: first row word, last row word, first column byte, last column byte, then 4 bytes of which the last (offset 9) is the expression length for an array formula |
| n bytes | Array formula: the shared expression, `n` = header byte 9. TABLE: 8 bytes. |
| … | Array formula only: fix-ups for the shared expression, as in §8.5 |

---

## 9. Parsed expressions

Reverse-Polish token stream. This is the direct ancestor of the BIFF2 `ptg`
encoding: token numbers and operand sizes coincide wherever they could be
checked.

A token byte has a base value in bits 5–0 and a class in bits 6–5. The
application normalises with: *if bit 6 is set, set bit 5; keep the low 6 bits*.
So `0x44` and `0x24` are the same token (reference), differing only in class
(value vs. reference).

### 9.1 Tokens with no operands (`0x01`–`0x16`)

Observed in samples: `03 04 06 08 09 0B 0C 0D 0E 11 13 16`. Meanings follow
BIFF **(inferred; consistent with every sample formula)**:

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

### 9.2 Tokens with operands

Operand sizes for `0x19`–`0x37` come from the application's own size table and
were verified by walking every formula in the samples.

| Token | Operand bytes | Meaning |
|------:|--------------:|---------|
| 17 | 1 + n | String constant: length byte, characters |
| 19 | 2 (+ n + 1) | Attribute: flags byte, data byte. If flag `0x04` (CHOOSE) is set, `data + 1` further bytes follow. Flag `0x02` = IF jump, `0x08` = skip. |
| 1A | 7 | External sheet reference start (last 3 bytes are a stale pointer; see §8.5) |
| 1B | 3 | External sheet reference end |
| 1C | 1 | Error constant |
| 1D | 1 | Boolean constant |
| 1E | 2 | Integer constant (word) |
| 1F | 8 | Number constant (double) |
| 20 | 6 | Array constant (data appended, §8.5) |
| 21 | 1 | Function, fixed argument count: function number |
| 22 | 2 | Function, variable argument count: argument count, function number |
| 23 | 7 | Name (first 3 bytes are a stale pointer; index appended, §8.5) |
| 24 | 3 | Cell reference: row word, column byte |
| 25 | 6 | Area reference: first row word, last row word, first column byte, last column byte |
| 26 | 4 | Memory area: 3-byte stale pointer to a cached area list, 1 more byte. The list is appended (§8.5). |
| 27, 28 | 4 | Memory error / no-memory (as BIFF) **(inferred)** |
| 29 | 1 | Memory function |
| 2A | 3 | Deleted cell reference |
| 2B | 6 | Deleted area reference |
| 2C | 3 | Cell reference inside a name definition |
| 2D | 6 | Area reference inside a name definition |
| 2E, 2F | 1 | Name-definition forms of the memory tokens (as BIFF) **(inferred from size only)** |
| 30 | 7 | Marked copy of `23` (name) |
| 31 | 3 | Marked copy of `24` (cell reference) |
| 32 | 6 | Marked copy of `25` (area reference) |
| 33, 34, 35 | 4 | Marked copies of `26`, `27`, `28` |
| 36 | 1 | Marked copy of `29` |
| 37 | 3 | Second marked copy of `24` |

Tokens `30`–`37` are working forms. While references are being adjusted, a
routine adds 13 to tokens `23`–`29` that involve a relative reference (and 19
to a cell reference whose column falls outside a given range, giving `37`);
the matching pass subtracts the same amount to restore them. They have the
same operands as the token they stand for and are not expected in a saved
file.

Row words in references: bits 13–0 = row, bit 15 = row is relative, bit 14 =
column is relative (`C0 12` = relative row 18, relative column).

### 9.3 Function numbers

One byte. The names are in the application's `STR#` 257 resource (208 entries).
With `list` as that resource, zero-based:

```
number 0…33   -> list[number + 2]     (COUNT, IF, ISNA, … VALUE)
number 34, 35 -> TRUE, FALSE          (list[0], list[1])
number 36…    -> list[number]         (AND, OR, NOT, … )
```

So 1 = IF, 31 = MID, 32 = LEN, 38 = NOT, 53 = GOTO, 55 = RETURN, and macro
commands simply continue the same numbering past 127: 169 = ACTIVE.CELL,
185 = INPUT, 188 = GET.FORMULA, 194 = SET.VALUE. Numbers 0–55 agree with later
BIFF; from 56 on they do not (Excel 1.x inserts COLUMN.WIDTH and FORMULA at
56–57).

---

## 10. Section 3 — external names

Names in other documents that this sheet refers to. For each external document
in §6.2, in order, for each such name:

| Size | Meaning |
|-----:|---------|
| Pascal string | Name |

The list ends with a single `0x00` byte. If the sheet has no external-document
table the section is **empty** (no terminator). `DB.FORM`:
`08 "DATABASE" 04 "FORM" 00`.

---

## 11. Section 4 — defined names

A sequence of entries, in alphabetical order (case-insensitive), terminated by
a `0x00` byte. A name's 1-based position here is the index used by `0x23`
tokens.

| Size | Meaning |
|-----:|---------|
| byte | `cch` = name length (never 0; 0 is the terminator) |
| byte | `cce` = definition length |
| byte | Flags. `0x04` = defined on a macro sheet as a function/command name; `0x20` = the definition is array-valued, so formulas using the name get cell word 2 bit 13 (§7); `0x40` = command macro, two extra bytes follow. Open keeps only bits `0x64`. |
| 2 bytes | Only if flag `0x40`: shortcut key character, then one more byte **(unknown)** |
| cce bytes | Definition — a parsed expression (§9). Unlike cell formulas, fix-ups are applied **in place**: in a `0x23` token the first two operand bytes hold the name index; in a `0x1A` token the last three operand bytes are `00 00 <document index>`. |
| byte | Copy of `cce` |
| cch bytes | Name characters |

A new worksheet always has five names with empty definitions: `Criteria`,
`Database`, `Print_Area`, `Print_Titles`, `Recorder`.

Example (`Expenses`): `08 07 04 | 2D 00 09 00 FF 00 03 | 07 | "Database"` —
area rows 9–255, columns 0–3.

---

## 12. Section 6 — windows

One record per open window on the document, front to back, terminated by a
`0x00` byte.

| Size | Meaning |
|-----:|---------|
| byte | `0xFF` marker |
| 4 words | Window rectangle in global (screen) coordinates: top, left, bottom, right |
| 41 bytes | Window state, copied from offsets 183–223 of the window record |
| word | `n` = number of selected areas |
| n × 6 bytes | Each: first row word, last row word, first column byte, last column byte |

On Open the rectangle is first adjusted to fit the current screen.

Window state block, by offset within the 41 bytes. Taken from the routine that
initialises a new window, the one that stores the cursor, and the DISPLAY
command:

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
| 15 | word | cols − 1 | Word companion of offset 13; compared against the last-non-hidden-column bound (§5, 0x016) when drawing |
| 17 | word | cols − 1 | Word companion of offset 14 |
| 19 | 4 bytes | 0 | One byte per pane: the pattern phase (0–7). Excel rotates its 8-byte QuickDraw pattern by this many bytes before drawing in the pane, so that patterned drawing stays aligned as the pane scrolls. |
| 23 | 3 bytes | | Stale pointer to the selection list |
| 26 | byte | 1 on a macro sheet | **Display formulas** |
| 27 | byte | 1 | **Display gridlines** |
| 28 | byte | 1 | **Display row and column headings** |
| 29 | 3 bytes | | Stale pointer to the document |
| 32 | byte | | Padding — never initialised or read; whatever was in memory |
| 33 | word | 0 | **Index of the active area** within the selection list that follows (the area containing the active cell) |
| 35 | word | 0 | **Active cell row** |
| 37 | word | 0 | Vertical pixel position of the horizontal split (between upper and lower panes); 0 = not split |
| 39 | word | 0 | Horizontal pixel position of the vertical split (between left and right panes); 0 = not split |

"rows" and "cols" are the counts that fit the window when it is created. A
window has up to four panes: the first value of each pair describes the upper
(or left) panes, the second the lower (or right) panes. This comes from the
print code, which builds a temporary window and stores the title rows in
offsets 3 and 7, the body rows in 5 and 9, and the boundary in 37, and does
the same for columns with 11/13, 12/14 and 39.

None of the samples is split, so their second values are leftovers (for
example bottom rows 17 and 20 in `Expenses`). When formulas are displayed,
columns are drawn at double width, which is why the macro sheets show a right
column of 2.

Check against `Worksheet1`: offset 0 is 3 and offset 35 is 4, and its single
selection is row 4, column 3.

---

## 13. Section 7 — page breaks and print settings

Three parts, each self-terminating. All three are a single `0x00` in every
sample, so the interiors are from code only.

1. **Column page breaks** — a sequence of chunks, each a count byte `n`
   (1…5) followed by `n` bytes, one column number per byte. Terminated by
   `0x00`.
2. **Row page breaks** — a sequence of chunks, each a count byte `n` (1…5)
   followed by `n` words, one row number per word. Terminated by `0x00`.

   In memory both are linked lists of 14-byte nodes holding five word slots;
   every node but the last is full, which is why a chunk never exceeds five.
   That the byte-sized list is columns and the word-sized list is rows follows
   from the value ranges **(inferred)**.
3. **Print record** — `0x00` if the document has none. Otherwise `0xFF`, then
   a 122-byte page-setup structure, then three strings, each either `0x00`
   (absent) or a Pascal string.

   The three strings are, in order, the **page header**, the **page footer**
   and the **printer message** (the set-up string sent to a serial "TTY"
   printer). Default header and footer are `&f` and `Page &p`.

   The 122-byte structure is Excel's Page Setup record:

   | Offset | Size | Meaning |
   |-------:|-----:|---------|
   | 0 | 62 bytes | The first 62 bytes of the Toolbox print record (`TPrint`: version, `TPrInfo`, paper rectangle, `TPrStl`, `prInfoPT`, `TPrXInfo`), copied in and out around the system print dialogs |
   | 62 | 4 doubles | Left, right, top and bottom margins, as shown in the graphics-printer Page Setup dialog **(order from the dialog's item order; unit presumably inches)** |
   | 94 | 3 bytes | Stale pointer — page header |
   | 97 | byte | Print row and column headings |
   | 98 | 3 bytes | Stale pointer — page footer |
   | 101 | byte | Print gridlines |
   | 102 | 4 words | Printable-area rectangle (top, left, bottom, right) in pixels; Excel derives print width and rows per page from it |
   | 110 | word | TTY left margin, in characters (0…256) |
   | 112 | word | TTY print width, in characters |
   | 114 | word | TTY top margin, in lines |
   | 116 | word | TTY print length, in lines |
   | 118 | 3 bytes | Stale pointer — printer message |
   | 121 | byte | Padding after the 3-byte pointer, bringing the record to an even 122 bytes |

   Offsets 97, 101, 110–118 and the three pointers are read directly from the
   dialog code (dialog items "Print Row & Column Headings", "Page Header",
   "Left Margin", "Print Width", "Top Margin", "Print Length", "Printer
   Message"). No sample contains a print record.

The file ends after this section.

---

## 14. Reading recipe

1. Check words at 0 and 2 are `0x5327` and `17`.
2. Seek to the word at offset 22. Skip the capacity word.
3. Read the 630-byte state block; compute `nRows`, `nCols` (§5.1).
4. Read the flags byte and the 8-entry section table; convert each entry with
   `hi × 512 + lo`.
5. Section 2: read `nRows × nCols` six-byte entries.
6. For each entry of type 2, 3, 5 or 6, read its record at
   `section5 + (word0 << 4) + (word1 >> 12)`.
7. Section 4: read names until the `0x00` terminator.
8. For formulas, walk the token stream to find how many fix-up words follow.

---

## 15. Open questions

* A few chart fields (§16): four option flags in the chart record, the role
  of the two built-in text objects, the 12-byte record trailer words, and the
  per-series word list.
* Number-format tokens beyond those the built-ins and the format compiler use
  (§6.1).

Everything that is described from code alone — array formulas, tables, array
constants, user formats, page breaks, the print record and all of §16 — has
no sample on the disk to confirm it against.

---

## 16. Chart files

Charts are not written in the worksheet format. Save hands them to a separate
writer (segment `0x5C`) that produces a Microsoft Chart document, Finder type
`MCBN`; Open reads them with segment `0x5D`. There is no chart on the disk, so
this section comes from those two routines and from the chart dialogs
(segments `0x57`–`0x5A`), which show what each field means.

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
          Page setup                   2 or 666 bytes
          Series lists                 <- header.seriesOffset
```

### 16.2 Header (128 bytes)

| Offset | Size | Value | Meaning |
|-------:|-----:|-------|---------|
| 0 | word | `0x018C` | Magic. Open rejects anything else. |
| 2 | word | 12 | Format version. Open rejects 1, treats 2 and 3 (older Microsoft Chart files) as 12, and uses bit 2 and bit 3 to decide how to read fonts and series headers. |
| 4 | word | `0x13E8` | Code signature (§3.1) |
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

The embedded segment is the application's `XCPC` 256 resource (598 bytes of
p-code). Open skips it.

### 16.3 Numbers

Chart files do not store IEEE doubles. Every number is converted to
Microsoft's 8-byte **decimal** format:

| Byte | Meaning |
|-----:|---------|
| 0 | Bit 7 = sign (1 = negative). Bits 6–0 = decimal exponent + 64. |
| 1–7 | 14 BCD digits, most significant first, two per byte |

```
value = (−1)^sign × 0.d1 d2 … d14 × 10^(byte0 & 0x7F − 64)
```

Zero is eight zero bytes. A value whose exponent exceeds 63 is written as
exponent 63 with all digits 9; one below −63 as exponent −63 with digits
`10 00 …`. So 1.0 is `41 10 00 00 00 00 00 00`.

### 16.4 Common sub-records

**Font descriptor** (4 bytes):

| Offset | Meaning |
|-------:|---------|
| 0 | Not used |
| 1 | Font |
| 2 | Size |
| 3 | Bit 5 = bold, bit 4 = italic |

In memory bytes 1 and 2 are a Macintosh font number and point size. On disk
they are Microsoft Chart codes: the font byte is 16 for New York (font 2), 1
for Geneva (font 3) and 32 for anything else, plus `0x40` for bold and `0x80`
for italic; the size byte is an index 0–2, `(size − 9) / 3` for Geneva and New
York and `(size − 12) / 6` otherwise. Excel also hides the real Macintosh font
number and size in spare bits nearby (listed with each record) and restores
them on load when the version has bit 2 set; for other versions it converts
the codes back.

**Line format** (4 bytes) — axis lines, gridlines, drop lines, hi-lo lines:

| Offset | Meaning |
|-------:|---------|
| 0 | Not used |
| 1 | Pattern; **0 = line not shown** |
| 2 | Word: bits 15–12 = weight (0–2, hairline to thick, in dialog order) |

**Area format** (16 bytes) — chart area, plot area, legend, text boxes:

| Offset | Meaning |
|-------:|---------|
| 1 | Border pattern (0 = none) |
| 2 | Word: bits 15–12 = border weight |
| 5 | Background pattern (0 = none) |
| 6 | Word: bits 11–8 non-zero = shadow; bit 14 = automatic size (text boxes); bits 15 and 14 are set when a chart is loaded |

Pattern bytes index Excel's 15-entry pattern palette, 1-based, with 0 meaning
"invisible". "Automatic" is a separate flag where the object supports it.

### 16.5 Chart record (378 bytes)

Offsets are within the on-disk record; "mem" is the offset in Excel's
392-byte chart structure that the bytes come from and go back to.

| Offset | mem | Size | Meaning |
|-------:|----:|-----:|---------|
| 0 | 0 | 18 | Chart header (below) |
| 18 | — | 6 | Built for the file: two words — (0, 0) for a lone chart, (1, 0) for the overlay, (0, 2) for a main chart that has an overlay (ignored on load) — then a word whose bits 11–8 are four chart option flags (mem 18, bits 15–12; **not identified**) |
| 24 | 20 | 16 | **Chart area** format |
| 40 | 36 | 24 | **Category axis** (axis record, below) |
| 64 | 62 | 24 | **Value axis** (axis record) |
| 88 | 88 | 44 | **Value-axis scale** (scale record, below) |
| 132 | 134 | 44 | Category-axis options, plot area and legend (below) |
| 176 | — | 78 | Zero; skipped on load |
| 254 | 178 | 4 | **Drop lines** line format (line and area charts) |
| 258 | 182 | 4 | **Hi-lo lines** line format (line charts) |
| 262 | 186 | 2 | Memory pointer — junk |
| 264 | 188 | 2 | **Number of text objects** that follow (§16.7), including the titles |
| 266 | 190 | 46 | Built-in text object −2 (§16.7) |
| 312 | 250 | 46 | Built-in text object −1 |
| 358 | 310 | 2 | Memory pointer — skipped on load |
| 360 | 312 | 2 | Number of words in the series word list |
| 362 | 314 | 2 | Memory pointer — junk |
| 364 | 316 | 2 | **Number of data-point format records** |
| 366 | 318 | 12 | **Default series format** (point-format record, §16.8) |

The two built-in text objects are addressed as indices −2 and −1 of the
chart's text list and are formatted together with the axis labels by "apply
to all"; what they display was not established.

**Chart header** (18 bytes):

| Offset | Meaning |
|-------:|---------|
| 0 | Word of flags: bit 14 = 100%; bit 13 = stacked; bit 12 = bars (horizontal) rather than columns; bit 11 = vary colours by category; overlay chart only: bit 10 = automatic series distribution, bits 9–2 = first series in the overlay chart |
| 2 | Word: **chart type** — 1 column or bar, 2 line, 3 pie, 4 area, 5 scatter |
| 4 | 4 bytes, not identified |
| 8 | Word: % overlap (column/bar). Negative means the "Overlapped" option is on; the magnitude is the percentage. |
| 10 | Word: % cluster spacing (column/bar) |
| 12 | Word: angle of the first pie slice, degrees from vertical (0–359) |
| 14 | 4 bytes, not identified |

The Gallery's "Main Chart" choices map as Area → 4, Bar → 1 with bit 12 set,
Column → 1, Line → 2, Pie → 3, Scatter → 5.

**Axis record** (24 bytes on disk):

| Offset | Meaning |
|-------:|---------|
| 0 | Word: bits 15–12 = major tick marks, bits 11–8 = minor tick marks (0 invisible, 1 inside, 2 outside, 3 cross); bits 7–4 = tick label position (0 none, 1 next to axis, 2 low, 3 high); bit 3 is set on load |
| 2 | Font descriptor for the tick labels |
| 6 | 6 bytes, not identified |
| 12 | **Axis line** format; 0 pattern = axis hidden. Bits 9–2 of its word at +2 hold the saved Macintosh font number. |
| 16 | **Major gridlines** format. Bits 9–2 of its word at +2 hold the saved font size. |
| 20 | **Minor gridlines** format |

**Scale record** (44 bytes) — the value axis, and for scatter charts also the
X axis:

| Offset | Meaning |
|-------:|---------|
| 0 | Minimum (decimal number) |
| 8 | Maximum |
| 16 | Major unit |
| 24 | Minor unit |
| 32 | Category axis crosses at |
| 40 | Word: bit 15 auto minimum, 14 auto maximum, 13 auto major unit, 12 auto minor unit, 11 auto crossing point, 10 logarithmic scale, 9 values in reverse order |
| 42 | Memory pointer — zeroed on load |

**Category-axis options, plot area and legend** (44 bytes, mem 134):

| Offset | mem | Meaning |
|-------:|----:|---------|
| 0 | 134 | Byte: value axis crosses at category number |
| 1 | 135 | Byte: bit 7 = value axis crosses between categories; bit 6 = crosses at maximum category; bit 5 = categories in reverse order |
| 2 | 136 | Byte: number of categories between tick labels |
| 3 | 137 | Byte, not identified |
| 4 | 138 | Word: non-zero (a memory pointer) when an X-axis scale record follows the chart record |
| 6 | 140 | **Plot area** format (16 bytes) |
| 22 | 156 | Word: bits 15–13 = **legend position** (0 bottom, 1 corner, 2 top, 3 vertical); bits 8–1 = saved legend font number |
| 24 | 158 | **Legend** font descriptor |
| 28 | 162 | **Legend** area format. Its byte at +7 holds the saved legend font size. |

For a scatter chart the category axis is a value axis: its scale is the extra
44-byte record after the chart record, and "reverse order" lives in that
record's flags instead of bit 5 above.

### 16.6 Chart-set header and trailer

The 8-byte chart-set header is three words from memory (two pointers and the
number of charts) and the word `0x0011`. Open skips it.

After the last chart:

| Size | Meaning |
|-----:|---------|
| 8 bytes | Window position and size: two 4-byte points, the first stored negated |
| 2 words | Two chart-window words **(not identified)** |
| word | 1 if a page setup follows, else 0 |
| 62 bytes | First 62 bytes of the Page Setup record (§13). Byte 32 carries the "print row and column headings" flag. |
| 58 bytes | Zero |
| 4 × 8 bytes | Margins as decimal numbers, in the order top, bottom, left, right |
| 256 bytes | Header: a Pascal string in a fixed 256-byte field |
| 256 bytes | Footer, likewise |

Without a page setup the word 0 is followed by 664 zero bytes instead, so the
section is always 666 bytes.

### 16.7 Text objects

A chart's text objects are the chart title, the axis titles and any attached
or free text. The chart record gives the count (mem 188); the 46-byte records
follow the chart record (or the X-axis scale), titles first. After all the
records come their strings, one Pascal string per object **except** objects
with automatic text, which have none.

Each 46-byte record is four slices of the 60-byte memory structure: bytes
0–5, 8–37, 38–43 and 46–49. In memory offsets:

| mem | Meaning |
|----:|---------|
| 2 | Word: bits 3–0 = **kind** — 1 chart title, 2 category-axis title, 3 value-axis title, 4 data label; bits 15–14 are cleared on load |
| 4 | Byte: series number (data labels) |
| 5 | Byte: point number, 255 = the whole series (data labels) |
| 8 | Memory pointer to the string — junk |
| 10 | Word: bit 13 = **automatic text**; bit 12 = vertical orientation; bits 11–9 = horizontal alignment (0 centre, 1 left, 2 right); bits 8–6 = vertical alignment (0 bottom, 1 centre, 2 top); bit 5 = show key; bit 4 = show value |
| 12 | Font descriptor |
| 20 | Bytes 20 and 21 hold the saved Macintosh font number and size |
| 22 | Area format of the text box; its word at +6 bit 14 = automatic size |
| 38, 46 | Not identified (likely position) |

On load the titles (kinds 1–3) of the main chart are moved out of the list
into Excel's three title slots.

### 16.8 Data-point formats and the series word list

The **series word list** is one word per series of the chart (count at mem
312). Open skips it.

**Data-point format records** (12 bytes each, count at mem 316) hold
per-point overrides; the same layout is used for the chart's default series
format at mem 318:

| Offset | Meaning |
|-------:|---------|
| 1 | Line or border pattern |
| 2 | Word: bits 15–12 = weight; bit 10 = automatic line/border |
| 4 | Area pattern |
| 6 | Series number |
| 7 | Marker |
| 8 | Word: bit 15 = automatic marker; bit 12 = automatic area |
| 0, 5, 9–11 | Not identified (the point number is among them) |

### 16.9 Series records

The series lists start at the absolute file offset in header word 26.
The first list holds the series flagged with header bit 10, the second the
rest; the header (word 24) gives the total. A chart file carries a **copy of
the plotted data**, not just references to it. Each series is:

| Size | Meaning |
|-----:|---------|
| word | Total length of this record, or 0 for the last record of a list (the next record then follows its strings) |
| word | Offset from the start of the record to its strings (4 + 98 + the arrays) |
| 98 bytes | Series header (below) |
| n × 8 bytes, or n strings | First array, `n` = point count: decimal numbers, or `n` Pascal strings for a text series (type 3). A numeric array shorter than `n` is padded with 1, 2, 3 … |
| n × 8 bytes | Second array, decimal numbers |
| Pascal string | Series name |
| 2 strings | Two label strings, each one length byte, the first part, `.`, and as much of the second part as fits in 255 |
| | Source document: `0x00` if none; otherwise the three bytes `02 1B 06` |
| Pascal string | Name of the source document (empty if none) |
| Pascal string | Second reference, **as formula text** (R1C1) |
| Pascal string | First reference, as formula text |
| 4 bytes | Four empty strings |

A zero word ends the series section.

**Series header** (98 bytes). Open uses only the fields marked †; the rest
are constants Excel writes for Microsoft Chart.

| Offset | Size | Meaning |
|-------:|-----:|---------|
| 0 † | word | Bits 15–13 = series type (0 dates, 1 or 2 numbers, 3 text categories); bit 10 = belongs to the first list; bits 11 and 7 = 1 |
| 2 † | word | Bits 14–13 = 2; bits 12–10 and 9–6 = format class and sub-code of the first array; bits 5–4 = 1 |
| 4 † | word | Bits 14–13 = 2; bits 12–10 and 9–6 = format class and sub-code of the second array |
| 6 | 2 bytes | 10, 10 |
| 8 | 2 words | 60, 60 |
| 14 † | word | Bits 15–14 = how the first reference is held (2 literal text, 3 formula); bits 13–8 = Excel number-format index of the first array |
| 16 † | word | The same for the second reference and array |
| 22 | 2 words | Copied from the series structure |
| 26 | 4 × 8 bytes | Four decimal numbers, all 1.0 |
| 60 † | 2 words | Point count, twice |
| 64 | 2 bytes | 1, 1 (byte 66 † is used for date series) |
| 68 | 11 words | All 2 |

Format class: 2 for General, 1 for formats 1–4, 4 for the currency formats
5–6, 5 for the percent formats 7–8; other formats go through a separate
mapping. For files from Excel (version bit 3 set) Open takes the format
indices from words 14 and 16 directly; for older Microsoft Chart files it
derives them from the class and sub-code.
