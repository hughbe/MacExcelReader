# How Excel 1.x displays a sheet

The [format specification](Spec.md) describes what a file holds. This describes how Excel shows it on screen: how a
value is turned into text in its number format, how that text is fitted into its cell and row, and how the window
round the cells is drawn. It uses the spec's conventions and cites code the same way (`segment.procedure`, spec
Background), and `ExcelValueFormatter` and `ExcelRenderer` follow it.

Drawing also follows the Macintosh's own code: the Mac Plus ROM (checksum `4D1F8172`) that the screens
were captured under, cited by ROM offset, with Apple's 1985 QuickDraw source (`DrawText.a`, `Stretch.a`), and
the scroll bar definition (`CDEF` 1) in the System file on Excel 1.00's Data disk.

## Formatting values

A number format's tokens are described in spec §7. Excel's TEXT function, given each sample format and a range of
values, shows how a format shows a number (`TEXT Formats 1`, `TEXT Formats 2`, `TEXT Date Limits` in Excel 1.00;
`TEXT Formats 1.5`):

* The section follows the sign, but a value that rounds to zero in its section's decimals uses the zero
  section: 0.001234 in `0.0" N";0.0" S";0` shows `0`.
* A single-section format shows a negative number as `-` and the format of its magnitude, dates included:
  -1234.567 in `m/d/yy` is `-5/19/07`.
* Dates and times round to the second (13:36:28.8 shows `13:36:29`). Excel converts the serial to seconds since
  1904 with the Macintosh's date routines, so the last date it shows is 6 February 2040 (serial 49710.27); TEXT
  gives #VALUE! after that.
* Strings and literals show without their quotes, and digits fill the placeholders around literals:
  1234.5 in `00-00-000` is `00-01-235`.

Excel 1.00's display routine (`19.0d`, native code) works through the section's tokens one by one, which
gives the rest:

* Number header `03`: the value goes through Excel's decimal conversion (at most 14 significant digits, a tie at
  the 14th rounding down), `c − b` is added to its decimal exponent (two per `%`, minus three per trailing
  comma), and those decimal digits are rounded half up to the positions there are: 1.005 in `0.00` is `1.01` and
  12345678901234.5 in `0.0` is `12345678901234.0`, since the conversion has already rounded it (`TEXT Digits`).
* Integer digit runs: digits that do not fit go before the first placeholder; a position with no digit shows
  `0` for `06` and nothing for `04`; with `07`, a comma follows every third position from the right that shows
  something.
* Decimal digit runs: leading zeros always show, even for `04`; after them the significant digits, then `0`s
  for `06` only, so `04` drops trailing zeros.
* Exponent `1D`/`1E`: `+` (for `1D`) or `-`, then the exponent's digits in the next run's positions.
* `22` writes `$`; `0A` repeats its character to fill what the column has left; `01` inside a section gets
  General with the width the rest of the section leaves.
* Date and time tokens go to `09.1a`, through the date conversion of `09.1b`, which fails past 2^32 − 1 seconds.
  `mmm` and `ddd` are the first three letters of the name, and 12-hour times show 0 as 12.

A cell whose formatted text is wider than its column shows as many `#` as fit instead (`19.0e`). Drawing cells,
below, gives the widths in pixels.

Excel tests whether a format shows a date in two places. `07.18` looks for a day, month or year token (`0B`–`12`,
`19`, `1A`) in the positive section; times alone do not count. Its caller `34.04` counts the built-in formats
10–18, the time-only ones included, and any other format `07.18` accepts.

TEXT will not take `General` or a fill (`*-0`), and gives #VALUE! for text longer than 255 characters (1E+300 in `0`). General (`19.0c`) fits the number into the characters the column
has room for, W, not counting a minus sign (Drawing cells, below, gives W in pixels, and what happens to a negative
number that rounds to nothing):

* It is written with an exponent when it has more integer digits than W, or when it is below 1E-3 and its digits
  do not fit, but not in a column narrower than 5 characters for numbers from 1E-3 to below 1E+4.
* Written out, it is rounded half up to W digits, less one for a point and one more for a leading `0.`.
* With an exponent, the mantissa has 1 digit up to W = 6 and W − 6 digits beyond that (one more when the
  exponent has two digits), and the exponent at least two digits: `1E+06`, `1.23E-05`.

Turning a number into text (`=""&A1`, in 1.00 and 1.5 alike, `General As Text`) is General with W = 20: 14
significant digits, written out in full from 1E-18 to below 1E+20 (`0.00000000015`, `99900000000000000000`)
and with an exponent outside that (`1E+20`, `1.2345678901234E+20`).

## Drawing cells

Nothing in the file says how a cell looks on screen; Excel works it out from the value, the cell's format,
alignment and style, the column widths and the sheet's standard font each time it draws a row. This section
gives those rules, so that a reader can show a sheet as Excel 1.00 did. They come from Excel 1.00's drawing
code, segment `19` (native 68000 code), and were checked against what Excel 1.00 showed for test sheets of
General numbers in 13 column widths in plain, bold and italic, and of text, fills, alignments and neighbours
([Testing](Testing.md)).

**Coordinates.** Columns sit side by side from the left edge of column A, each as wide as its width in the
state block (spec §6, `0x042`); a hidden column (width 0) takes no room and nothing in it is drawn. A column of width
*w* whose left edge is *L* covers pixels *L* to *L* + *w* − 1, the last being its right gridline. Text is
clipped so that it never covers that gridline.

**Font.** Everything is measured with the sheet's standard font and size (`0x244`, `0x246`), Geneva 10 in every
sample, with QuickDraw's `TextWidth`. `09.1c` caches its metrics (`0x248`–`0x254`). Bold, which QuickDraw draws by
smearing each character a pixel to the right, makes every character a pixel wider: the bold `0` is 8 pixels in
Geneva 10, the plain one 7. Italic text is given half the font's ascent (`0x248`, 5 in Geneva 10) of extra room
for its slant. The widths are the advance widths in the font's `FONT` resource; the System file on Excel 1.00's
Data disk has bitmaps of Chicago 12, New York 10, 12, 20 and 24, Geneva 9, 10, 12, 14, 18, 20 and 24, and
Monaco 9.

Excel also allows any other size. The Font Manager draws one the System has no bitmap for (`RealFont` false) from
the next larger bitmap of the family, or from the largest if none is larger: Geneva 11 from Geneva 12, Geneva 16
from Geneva 18, Monaco 12 from Monaco 9. It returns the scale as numer ÷ denom with denom 256 and numer
`FixRatio`(size, bitmap size) shifted right 8 bits, so not exactly size ÷ bitmap size: Geneva 11 is 234/256 of
Geneva 12, not 11/12 (`FMSwapFont` in the Mac Plus ROM, at ROM offset `0xED7C`; `FixRatio`, at `0x1593A`, rounds
down). The widths stay the bitmap's whole pixels, a pixel more each for bold, and `TextWidth` adds them up before
it scales the total and rounds down (ROM `0x8CE8`): `1E+04`, 38 pixels in Geneva 12, is 34 in Geneva 11, and a `0`
on its own is 7 (8 × 234/256 = 7.3). The ascent, descent and leading scale to the nearest pixel (`GetFontInfo`;
Geneva 11: 11, 3 and 1). Excel 1.00's `09.1c` adds a pixel to both cached digit widths when `RealFont` is false,
and on screen a `#` counts as 9 pixels too. So a Geneva 11 digit counts as 8 pixels and a `#` as 9, while text is
checked against the room at its plain width: `TRUE` (28 pixels) fits a 29-pixel room. Excel 1.03 instead keeps a
flag for a scaled size, and uses it both to widen those divisors and to add a pixel per character to the width it
checks (below).

**Releases.** In a size the System has a bitmap for, Excel 1.03 and 1.5 draw cells exactly as 1.00 does:
every test sheet in Geneva 10 came out the same, pixel for pixel. In a size with no bitmap, Excel 1.5 still
matches 1.00 cell for cell, while Excel 1.03 does not: when it checks whether text fits, it adds a pixel per
character on top of the scaled width. In Geneva 11, `TRUE` (28 pixels) needs 32 pixels of room in 1.03 and 28 in
1.00 and 1.5, and `1E+04` (34) needs 39. The bold column and row headings of a scaled size depend on the System
rather than on Excel: Excel 1.00 draws them with heavier letters under the System on its own Data disk than under
a later System, where it matches 1.03 and 1.5 pixel for pixel. The cells look the same under both. `FMSwapFont`
takes the extra width a style adds from the family's `FOND` resource when there is one and the Font Manager's
flags allow it, scaled to the size (ROM `0xE7A8`–`0xE7CE` and `0xEA36`–`0xEA82`), and otherwise adds a whole pixel
for bold; the Data disk's System has no `FOND` resources, and a later one has, which accounts for the difference
**(inferred)**: the later System was not examined.

**Rows.** Every row has the same height, the font's ascent + descent + 1 (`09.1c`; 13 in Geneva 10, 15 in Geneva
11), and its bottom pixel is its gridline. `09.1c` also sets a floor from the screen's size that no font reaches.
Text sits on a baseline one ascent below the row's top: in Geneva 10 capitals cover pixel rows 3–10 of 13 and
descenders reach row 12.

**The room.** A cell's value has its column's width less 2 pixels each side, and less the italic allowance for
an italic cell, to fit in (`19.0a`): 71 pixels in a default 75-pixel column.

**Numbers** (`19.0d`, `19.0e`):

* General gets *W* = (room − the width of `-` for a negative number) ÷ the width of `0` characters, rounded
  down, and is written as spec §7 describes. A negative number that rounds to nothing there loses its sign: Excel
  writes the `-`, finds the digits gone and writes `0` over it. So −123 shows as `0` in a 12- or 15-pixel column
  (*W* = 0), where 123 shows as nothing (see Excel bugs).
* Any other format fits the room as spec §7 describes, in pixels: General inside a section gets (room − the width of
  the rest of the text) ÷ the width of `0` characters, and a fill (`*`) repeats its character
  (room − 1 − the width of the rest) ÷ its width times, rounded down, to at most 255 characters in all.
* If the text is then wider than the room, the cell shows as many `#` as fit, room ÷ the width of `#` (9 in
  Geneva 10, 10 bold), centred whatever the cell's alignment. A 71-pixel room holds 7.

TRUE, FALSE and errors are fitted the same way: `TRUE` (26 pixels in Geneva 10) shows in a 30-pixel column, and
`FALSE` and `#DIV/0!` become `##`. Text is never turned into `#`.

**Alignment** (`19.0b`). General aligns text left, numbers right, and TRUE, FALSE and errors centre (`19.0e`).
With *L* and *R* = *L* + *w* the cell's edges and *t* the text's width (with the italic allowance):

| Alignment | Text starts at |
|---|---|
| Left | *L* + 2 |
| Right | *R* − 2 − *t* |
| Centre | (*L* + *R* − *t*) ÷ 2, rounded down |

**Text running into the next cells** (`19.07`, with `19.04` and `19.05` finding where to stop). Text wider than
its cell carries on over its neighbours, to the right for left-aligned and General text, to the left for
right-aligned text and both ways for centred text, as long as each neighbour is blank. A cell with any value
stops it, even empty text (`=""` or an empty constant) or a number; a blank cell that is only formatted, in any
alignment including Fill, does not. Hidden columns let it through. The text is cut off at the first cell with a
value, and at the left edge of column A. Numbers, TRUE, FALSE and errors never run on.

**Fills** (`19.0b`). A Fill-aligned cell repeats its text, formatted as for display when it is a number, across
a run: the cell itself and the blank Fill-aligned cells after it (hidden columns included). The text starts at
*L* + 2 and is repeated (run width − 3 − the italic allowance) ÷ *t* times, rounded down: `ab` (11 pixels) 20
times over three 75-pixel columns, `1.50` (24 pixels) 3 times in one. Text longer than that is drawn once and
cut off at the run's end. A fill never runs into a blank cell that is not Fill aligned.

**Clearing under text.** Before drawing a cell's text Excel clears the stretch the text is clipped to, the row's
height less its gridline, so text running into its neighbours covers the gridlines between them.

**Italic and bold.** QuickDraw slants italic text by moving each row of a glyph right by half its distance from
the bottom row of the font's rectangle (ascent + descent), rounded down: in Geneva 10 the row just above the
baseline moves 1 pixel and the top of a capital 4; in Geneva 12, with a descent of 3, 2 and 5. Bold ORs each
glyph with itself a pixel to the right.

**A size with no bitmap** is drawn at the bitmap's size and stretched, by QuickDraw's `DrText` and `StretchBits`
(the Mac Plus ROM's, which match `DrawText.a` and `Stretch.a` in Apple's 1985 QuickDraw source). With *x* and *y*
the pen, *u* the text's width at the bitmap's size (bold included) and *n*/256 the scale:

1. `DrText` draws the text at the bitmap's size, from the pen, into a buffer: the source rectangle runs from *x* to
   *x* + *u* + 32 (slop for bold and italic, in the srcOr to srcBic modes) and from *y* − the bitmap's ascent down
   the height of its font rectangle. Bold and italic are applied in the buffer as for a real size.
2. `MapRect` maps the rectangle to the scaled size: each edge's distance from the pen times *n*/256, rounded to the
   nearest pixel away from the pen.
3. `StretchBits` copies the buffer into the mapped rectangle. Down the rows, an error term starts at minus half the
   source height; each source row read adds the destination height, rows are ORed together while it stays at or
   below 0, and a destination row is written, and repeated while subtracting the source height leaves it at or
   above 0. Across, the fraction *f* = `FixRatio`(smaller width, larger width) steps a 16-bit error term that
   starts at *f* ÷ 2: shrinking, source column *i* goes to destination column ⌊(*f* ÷ 2 + *i* × *f*) ÷ 65536⌋,
   columns falling together ORed; enlarging, destination column *i* takes that source column. For exact ratios
   `StretchBits` has faster routines: ½, ¼ and ¾ give the same columns, 2, 3, 4, 6, 8, 16 and multiples of 8 repeat
   each source column, and 1.5 gives each 4 source columns 1, 2, 1 and 2 destination columns.
4. The pen moves *u* × *n* ÷ 256, rounded down: the scaled `TextWidth`.

Pixels kerned left of the pen are not copied. These rules reproduce Excel 1.00's Geneva 11 pixel for pixel. They
are the Mac Plus ROM's, which the screens were captured under; Excel 1.00 also ran on the 64K ROM of the 512K Mac,
whose text drawing is **(unknown)**.

## Drawing the window

Excel draws the rest of a worksheet or macro sheet's window round the cells, all in window coordinates from the
top-left of its content area (`Render Basics` and the other `Render` samples, captured from Excel 1.00's screen
and matched pixel for pixel by `ExcelRenderer`; docs/Testing.md).

**Headings.** With headings shown (spec §13, offset 28), a strip of the width at window offset 1 runs down the left and
a row of one row height across the top. Excel sets the strip's width from the standard font and the widest row
number showing, and saves it: 29 pixels in Geneva 10, 32 in Geneva 11 and 12, 26 in New York 10 and Monaco 9, and
45 in Geneva 10 when row 16384 shows. Each column letter and row number is in the standard font, bold, centred in
its heading less the line at its right or bottom edge. The headings' lines are solid: one along the strip's right
edge, one under the heading row, and one at each column's right edge and each row's bottom edge.

**Gridlines** (offset 27) are dotted in QuickDraw's `gray` pattern along each column's right edge and each row's
bottom edge, from the headings to the scroll bars. The pattern lines up with the window's content area, rotated
down by the pane's pattern phase (offset 19): in a window with phase 0 a dot is black where x + y is even. Excel
saves phase 0 for Geneva 10 and 5 for Geneva 11 and 12, and 3 for New York 10 and Monaco 9, so their dots fall on
odd x + y.

**Borders** are solid lines on a cell's edges: the left border on the gridline before the cell, the top border on
the gridline above it, the right and bottom borders on its own gridlines.

**Formulas shown** (offset 26): columns are twice their width, and every cell shows what was typed into it, a
formula's text or a constant (numbers in General), from *L* + 2, left aligned and clipped to its own cell.

**The selection.** Each selected area is inverted from the pixel left of the gridline before its first column to
the pixel right of the gridline after its last column, and from the pixel above the gridline over its first row to
the pixel below the gridline under its last row, clipped to the cell area. The active cell is then inverted back
inside a one-pixel frame, from *L* + 1 to *R* − 3 and from its top + 1 to its bottom − 3, so it shows as it is
with an inverted outline.

**The sheet's end.** Below row 16384 and right of column IV the window is grey (`gray`, the headings' strip
included), and the last row's and column's gridlines are solid lines.

**Scroll bars.** The System's scroll bars sit on the content area's right and bottom edges, 16 pixels wide, with
the grow box in the corner. Excel draws a black split box, 4 pixels long, before each bar's first arrow; dragging it
splits the window (spec §13, offsets 37 and 39). The thumb is a 14 × 16 white square inside the track, which is filled
with `ltGray`. The bars are controls with a minimum of 0: the horizontal bar's value is the left column and its
maximum 255, the vertical bar's maximum is 17982 and its value comes from the top row (`18.0c` sets them,
`18.01` and `18.11` set the value with `SetCtlValue`, and dragging the thumb maps the value back with `18.13`;
`16.0c`, `16.01`, `16.11` and `16.13` in 1.03). The value of top row *r* (`18.12`; `16.12` in 1.03) gives the first and last 20 rows 819 each,
so that the thumb moves for each of the first rows, and the 16,343 rows between them the rest:

| Top row *r* (from 0) | Value |
|---|---|
| 0–19 | (*r* × 819 + 10) ÷ 20 |
| 20–16362 | ((*r* − 20) × 16344 + 8171) ÷ 16343 + 819 |
| 16363–16383 | ((*r* − 16363) × 819 + 10) ÷ 20 + 17163 |

Each division rounds down; row 16383 gives 17982. The System's scroll bar (`CDEF` 1) puts the thumb value × travel ÷
maximum pixels along its travel, the track's length less the thumb's, rounding up only when the remainder is more
than half the maximum: (column × travel + 127) ÷ 255 for the horizontal thumb.

## Open questions

* How the 64K ROM that Excel 1.00 also ran on draws a font size with no bitmap, and what a later System's `FOND`
  resources change in bold headings of that size.
