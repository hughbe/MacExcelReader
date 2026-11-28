# How Excel 1.x displays a sheet

The [format specification](Spec.md) describes what a file holds. This describes how Excel shows it on screen: how a
value is turned into text in its number format, how that text is fitted into its cell and row, and how the window
round the cells is drawn, and how a chart's automatic axis scales are worked out. It uses the spec's conventions and cites code the same way (`segment.procedure`, spec
Background), and `ExcelValueFormatter`, `ExcelRenderer` and `ChartScale.SetAutomaticValues` follow it.

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

Excel 1.00's display routine (`19.0d`, native code) works through the section's tokens one by one, in the order
they are stored, which gives the rest. It keeps the number as a string of decimal digits with a count of how
many come before the decimal point (negative when zeros follow the point first), and a count of the integer
placeholders still to come, which starts at the header's `a`.

* A negative number shown by the positive section starts with `-`, before any token, so even an empty format
  shows `-` (`TEXT(-1,"")`).
* Number header `03`: the value goes through Excel's decimal conversion (at most 14 significant digits, a tie at
  the 14th rounding down), `c − b` is added to its decimal exponent (two per `%`, minus three per trailing
  comma), and those decimal digits are rounded half up to the positions there are: 1.005 in `0.00` is `1.01` and
  12345678901234.5 in `0.0` is `12345678901234.0`, since the conversion has already rounded it (`TEXT Digits`).
* Scientific header `1C`: the digits are rounded to `a + b` and their trailing zeros dropped; the exponent is the
  decimal exponent minus `a`, and `a` digits come before the point. Percent signs do not move the point here:
  29220.75 in `0E-0000.0%` is `3E00004.0%`.
* Integer digit runs (`04`, `06`) before any `.`: while there are more placeholders to come than digits before
  the point, each placeholder shows `0` for `06` and nothing for `04`. The run that reaches the digits shows all
  of them that the placeholders after it cannot hold, so digits that do not fit go there, and a placeholder with
  no digit left shows `0`. Once `07` has been met, a comma follows every digit or `0` that leaves a multiple of
  three placeholders to come.
* `05`: integer digits that no run has shown come out first (`.000` shows 10000000000 as `10000000000.000`),
  then `.`. A second `05` only prints another `.`: the decimal places carry on across it.
* Decimal digit runs (after a `05`): leading zeros first, even for `04`, then the digits that are left, then
  `0`s for the rest of the run for `06` only. Each run pads on its own, so `#.##0.000,0` shows -0.5 as
  `-.50.000,0` and `0.0.0` shows 0 as `0.0.0`.
* Exponent `1D`/`1E`: `+` (for `1D`) or `-`, then the exponent's digits take the place of the number's, as an
  integer with the token's width (the placeholders after the `E`) as its placeholders, and the comma flag off.
  So they spread over the runs after the `E` like integer digits, and a `.` after them prints and starts decimal
  runs that have no digits left: 1234.5678 in `0E+00 0` is `1E+00 3`, in `0E+00.0` `1E+003.0`, and 1.5E+123 in
  `0E+00000` is `2E+00123`.
* `22` writes `$`; `0A` repeats its character to fill what the column has left; `01` inside a section gets
  General with the width the rest of the section leaves.
* Date and time tokens go to `09.1a`, through the date conversion of `09.1b`, which fails past 2^32 − 1 seconds.
  `mmm` and `ddd` are the first three letters of the name. The 12-hour hours `1F` and `20` show 0 as 12; `13`
  and `14` are 24-hour even in a section with AM/PM, since the compiler makes only the hour just before an AM/PM
  12-hour (spec §7): `h:mm h:mm A/P` shows 1 as `0:00 12:00 A`.

Checked against TEXT in Excel 1.00 for 1,042 random formats: all agree except five whose exponent count Excel's
compiler misplaced (spec §7), where the count byte holds a value left over from memory.

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
for bold. The Data disk's System has no `FOND` resources, while System 3.2 (on Excel 1.04's disk) has five, and
its Geneva `FOND` gives bold an extra width of 0 (the style table at offset 28 is all zeros, flags 0), so under it a
scaled bold Geneva gets no extra pixel. Times and Helvetica's `FOND`s do have extra widths (Times bold 341/4096
of the size). `excel-test files` lists and copies out a disk's files to look at them.

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

## Automatic axis scales

When Excel 1.00 draws a chart it works out every scale field marked automatic (spec §16.5) and writes the result
into the scale record, which Save then writes out: a chart Excel has drawn holds the values it drew with
(`37.0a` in 1.00). A pie chart has no axes and its scale is left alone. Every step below uses Excel's own
floating-point operations in this order; `ChartScale.SetAutomaticValues` follows them and gives the same doubles
for every chart sample, Excel's own from 1.00, 1.03 and 1.5 among them.

**The data** (`37.16`): the smallest and largest value on the axis, over the chart's series. With an overlay, the
main chart plots the series before the overlay's first series and the overlay the rest.
A chart with no series keeps its scale as it is.

* A stacked chart (header bit 13, not scatter) runs from 0 to the largest total of a category's values, each
  counted positive: two series of −1, −2, −3 and −4, 5, −6 give 0 to 9 (`Scale Stacked Negative`).
* A 100% stacked chart runs from 0 to 100.
* A scatter chart's X axis takes the series' category numbers, or 1, 2, … for a series that has none.
* A logarithmic axis leaves out values that are not positive, and has 1 wherever a linear one has 0.
* With no values both ends are 0 (1 on a logarithmic axis).

**A single value.** When the data's ends are equal, the minimum moves a tenth of its size away from zero, and
for a scatter chart the maximum moves a tenth the other way: 5 gives 4.5 to 5, or 4.5 to 5.5 in a scatter chart.
All zeros give 0 to 1.

**Fixed fields.** A fixed minimum or maximum replaces the data's. On a logarithmic axis a fixed maximum *m*
becomes 10 to the power of ⌈log₁₀ *m*⌉, and a fixed minimum, which the record holds as an exponent *e*, becomes
10 to the power of ⌊*e*⌋.

**Starting at zero.** Unless the minimum is fixed or the chart is a scatter chart, a scale whose data does not
straddle zero starts at zero when the range is at least 8% of the end furthest from zero (|100 × range ÷ end| −
10 + 2 ≥ 0): 92 to 100 starts at 0, 93 to 100 and 92.5 to 100 do not. Negative data ends at zero instead. When
the end it divides by is 0 (data starting at 0 under a fixed maximum below it), the scale never starts at zero.

**The major unit.** With *p* the power of ten below the range (10 to the power of the exponent `Num2Str` gives
the range, less 1) and *n* intervals, Excel takes the smallest of 10, 20, 50, 100, 200, 500 and 1000 (a table in
1.00's initial data) at least ⌈100 × |range| ÷ *n* ÷ *p*⌉, and the unit is that × *p* ÷ 100. A fixed end beyond
the data's other end makes the range negative; the unit is the same as for its size (fixed maximum −20 over data
0 to 10 gives a unit of 2 and a minimum of −22). 0 to 101 over 10
intervals needs 10.1, so the unit is 20 × 100 ÷ 100 = 20 and the axis runs to 120.

*n* is the number of tick labels that fit along the axis, from 1 to 10. `37.08` works it out before the labels
are laid out: it sets the plot area to the space left after the titles and legend (`37.10`), works out the scales,
and only then measures the labels (`37.09`), takes their room off the plot area (`37.0e`) and sets it again. So
the count is against that provisional plot area:

* Along a vertical axis: its height, less twice the other axis's label height, divided by the axis's label height.
* Along a horizontal axis (a bar chart's value axis, a scatter chart's X axis): its width divided by three widths
  of the digit 0 in the axis's label font.

Each division rounds down. A label height here is the font's ascent and descent and a width that of 0 (in bold
for bold labels), from `37.0b`, in twips: 20 a pixel (`32.01` maps the window's pixels onto the drawing's twips).
On screen Excel 1.00 draws a chart's labels at their own size in every window; `30.26`, which picks smaller sizes
to fit a window's share of each font, belongs to another path, perhaps printing **(inferred)**.

**Wrapping category labels** (`37.2b`, `33.1e`, `33.1f`). A category label's slot is the plot area's width divided
among the labels (as many as the categories, one fewer when the value axis does not cross between them, divided by
the categories between labels), but never narrower than a 0; in a bar chart it is a third of the chart area's width.
Excel converts it to pixels and fills each line a word at a time (`33.21`), a word running to and including the
next space or hyphen, while the words' widths, trailing spaces included, fit. A first word too wide for its line
breaks between characters, and a line takes at least one character. A return or line feed ends a line. Excel's own
column charts with labels from one to nine lines of Geneva 9 (the `Labels` samples) lose exactly 240 twips of plot
height for each line more.

**The room the labels take** (`37.0e`), on each side of the plot area in turn:

* An axis's labels measure as the largest label's box: its lines of text (each the font's ascent, descent and
  leading) and its widest line, with a pixel of margin all round (`37.09`, `33.1d`). Category labels are the first
  series' categories, every so many as the category axis says, each wrapped to its slot. A hidden axis, or one
  without labels, takes no room.
* Number labels (`37.2a`, `37.2c`) are not all measured. Excel formats the scale's minimum and maximum, and when the
  major unit is not whole the minimum plus a unit and the maximum less one, as General shows them in a cell as wide
  as the label's slot in pixels (at least four 0s wide: `33.01` calls the cell routine `19.0d`), and lays out the one
  with the most characters, wrapped to the slot. A horizontal axis shares the plot area's width among its intervals;
  a vertical one has a quarter of the chart area's width. A logarithmic scale's labels are the powers of ten at its
  ends; one too large for a number has none. Categories that are numbers are labelled 1 to the number of categories,
  in slots of at least 210 twips.
* Along the top and bottom of a column or line chart the category axis's labels lie (a bar chart's value axis);
  along the left and right the value axis's, with 100 twips more.
* On the side the labels are on (the low side, or the high side when the scale runs in reverse): labels next to
  the axis take their size less how far the axis is from that side, never less than 0; "low" labels take their
  size. On the other side only "high" labels do.
* The other axis's labels overhang the ends: each side takes at least half the other axis's label size, except
  the value axis's labels on a chart whose value axis crosses between categories (and is not a scatter chart).
* Each side also takes twice the axis's tick length.

`37.08` takes the room off the left and right first, measuring the labels in the provisional plot area's width;
then it measures again the labels along the top and bottom (a bar chart's value labels, otherwise the category
labels) in the narrower width, and takes the room off the bottom and top.

The saved plot area is the provisional one less this room, so a chart Excel has drawn gives back the provisional
plot area and with it *n* (`ChartDocument.GetLabelSpaces`; `GetTickLabelSize` gives the label box). Label counts worked out this way reproduce Excel's
scales for all 728 chart saves of the experiments, among them 495 library charts in windows from 100 to 340
points, with labels in the System font (Chicago 12), Geneva 9 and New York 9.

**The other fields.** An automatic minor unit is a fifth of the major unit. Then each unit, fixed or not, is
limited: one too small to change either end of the scale becomes the whole range, and one that would divide the
range into more than 500 parts becomes the range ÷ 500. An automatic maximum moves up and an automatic minimum
down to a whole number of major units, unless it is one already; the multiple is the unit times the quotient
rounded towards zero. An automatic crossing point is 0, or the end of the scale nearer to zero when zero is
outside it. Last, an automatic minimum not below the maximum becomes the maximum less a major unit, and an
automatic maximum not above the minimum the minimum plus one.

**Overlays.** Each time it draws a chart with an overlay, and both charts have series, Excel works out again
whether the overlay shares each axis with the main chart (`37.34`), and saves the result in the overlay option
flags (spec §16.5):

* The value axis is shared when either chart's value axis is hidden (its line pattern is 0, as Excel makes an
  overlay), neither chart is a pie, both are bar charts or neither, and both are 100% charts or neither. The chart
  whose axis is hidden takes the other's logarithmic and reverse-order settings.
* The X axis is shared when both are scatter charts. It is also shared when either chart's category axis is hidden,
  neither is a pie or scatter chart, and both are bar charts or neither. The chart whose axis is hidden takes the
  other's reverse order (of its X scale for a scatter chart) and, for scatter charts, its logarithmic setting; the
  overlay takes the main chart's crossing between categories.

On a shared axis both charts get the range of their data together (`37.2f`): the smaller minimum and the larger
maximum of the two, each chart's own taken as above. A scatter overlay with 2 points on a scatter chart with 3 gets
the main chart's X axis, 1 to 3. The scale counts as a scatter chart's only when both are scatter charts.

**Logarithmic scales** (`37.2d`) hold exponents of ten. The maximum is ⌈log₁₀ max⌉, allowing 1E−10 for rounding
in the logarithm, and the minimum ⌊log₁₀ |min|⌋: 0.5 to 30 gives −1 to 2. The minor unit is 1, and the major unit
is (⌊range⌋ − 1) ÷ (*n* − 1) + 1 in whole numbers, where *n* is not limited to 10. An automatic maximum is then
the minimum plus the range moved up to a whole number of major units.

## Laying out a chart

When Excel 1.00 draws a chart it lays it out from the outside in (`37.03`), in twips, with vertical coordinates
running upwards from the chart area's foot (so a box's top field is its lower edge, spec §16.5):

1. **The chart area** fills the drawing less its margins. A chart Excel has not drawn gets 15/16 of the drawing's
   width and 23/25 of its height, centred: 9075 by 4784 twips in the 9680 by 5200 drawing of a new chart, with
   margins of 302 and 208. After that Excel scales the saved chart area with the window, so a chart's layout
   follows its history; the saved one is the one to start from.
2. **The border and shadow** (`37.04`): a border takes its weight plus two pixels off each side, and a shadow a
   fortieth of the chart area's smaller side more off the right and bottom **(inferred: no sample has either)**.
3. **The title** (`37.07`), laid out in four-fifths of the chart area's width, takes its box and 5% of the chart
   area's height off the top. A text object's box (`33.1d`) is its lines and widest line with a pixel of margin all
   round, a pixel more each side for a border; a shadow adds a fortieth of the chart area's smaller side to the
   room it takes (`38.02`). Its font is the Macintosh number and size at mem 20 and 21 (spec §16.7): Geneva 10 in
   Excel's own charts, 0 and 0, the System font at 12, in the library's. A chart without a title of its own but with a single named series is titled with the
   series' name, in Geneva 9 **(inferred)**: library charts lose 280 + 5% at every window size. Excel's own charts
   made from unnamed series have no title.
4. **The legend** (`57.00`–`57.04`): its entries are the series' names, or a single series' categories (not in an
   area chart, or with an overlay's series). Each entry is a key, as wide as the legend font's ascent and widest
   character, and its text, wrapped beside it, with a pixel of margin all round; entries are 80 twips apart (the
   spacing is (bits 12–10 of the legend word + 2) × 40). A legend on the right or in the corner stacks them, the text
   wrapped in a third of the chart area's width; at the top or bottom they go in as many columns as fit the chart
   area's width (the text in a fifth of it, or its share), rows as even as they can be, no more than half the
   chart area high. A legend of fixed size keeps its saved box. The box and a pixel (20 twips) come off the bottom,
   the top or the right; a corner legend takes them off the right and the top, unless it is more than half the chart
   area high, when it goes on the right.
5. **The tick length** of both axes is the smaller side of what is left, divided by 70.
6. **The axis titles** (`37.05`, `37.06`), the category axis's first: a title beside the plot area (the value
   axis's, or a bar chart's category axis's) is laid out in a fifth of the chart area's width and takes its box's
   width and 20 twips off the left; one below it, in half the width, takes its box's height and 20 twips off the
   bottom. Vertical text swaps the widths, and stacks its characters, each a line high, in columns an M wide with an
   n between them.
7. **The provisional plot area** is what is left: Excel counts the tick labels that fit against it (`37.0a`, below),
   works out the scales, and takes the labels' room off it (`37.0e`) to leave the plot area.

Each step takes its room off one side (`37.0f`), but never leaves less than a quarter of the chart area's width or
height; below the plot area the top first moves up as far as the chart area allows. Laid out this way, every chart
Excel 1.00 saved without an overlay comes out with the plot area and tick length it saved: 745 charts, with titles,
legends in every position and axis titles among them, and library charts Excel had not drawn before
(`ChartDocument.LayOut`). Charts Excel 1.03 and 1.5 saved differ by a pixel or so in places: 1.03's value labels of
fractions and its scatter X labels, 1.5's automatic titles and automatic axis titles.

## Drawing a chart

Excel 1.00 draws a chart in twips and maps each point to the window's pixels as it goes (`32.04`): across, a twip
coordinate becomes x × (window width − 16) / drawing width, and down, (drawing height − y) × (window height − 16) /
drawing height, both truncated, where the window's size is words 3 and 2 of the chart's window block. A length
mapped on its own (`32.02`) is truncated the same way, so where Excel draws a line by its length rather than to its
end, the end can fall a pixel short of the mapped point. The chart's drawing coordinates are the chart area's plus
the margins (spec §16.5), its foot standing the top margin up from the drawing's even where the drawing leaves an
odd twip below it **(inferred: windows sized so show it)**.

**Lines.** Every line is QuickDraw's: a one-pixel pen run from the top end down a scanline at a time, each row a
slab whose edges start half a pixel in and half the slope along and step by the slope, 16.16 fixed point with the
slope truncated (the Mac Plus ROM's DrawLine). Excel moves to a point and draws a line of a given length
(`32.05`, `32.07`), with the start half a pen (10 twips) left of and below the point:

* **Axes** (`3a.03`) run from the plot area's foot or left edge for the plot area's height or width. The value axis
  stands at its saved position in a column or line chart; a bar chart swaps the two axes.
* **Ticks** (`3a.04`–`3a.06`) are drawn a 70th of the plot area's smaller side long, and at least a pixel and a
  half (`3a.00`; the saved tick length, a 70th of the chart area's, only lays the chart out). A tick outside or
  across starts its length out from the axis and runs back to it.
* **Gridlines** (`3a.08`, `3a.0f`) cross the plot area in their pattern, before the axes are drawn: the value
  axis's minor then major, then the category axis's. Minor gridlines on a logarithmic scale fall at 2 to 9 times
  each power of ten.
* **The order**: the chart and plot areas, gridlines, axes and ticks, then the series, then the tick labels, the legend and the text objects.
  A tick label erases its box first, as a data label does (below), so a marker or column under a label is cut
  away (`Draw Log Point Below Axis`, `Draw Labels Column Negative`) **(inferred: the samples show it)**. The main chart is drawn first and an overlay over it.
* **A line chart's series** (`53.03`) joins its points with segments that each start half a pen from one point and
  run for the distance to the next, its length truncated to pixels. A point is the category's middle (or its edge,
  when the value axis crosses at a category) and the plot area's foot plus the value's share of its height
  (`37.23`), mapped without the pen's offset.

**Markers** (`53.08`) are 8-pixel bitmaps from Excel's resources, from 4 pixels left of and above the point to 3
right of and below it: Excel clears a mask a pixel larger than the marker, then draws the marker (CopyBits, srcBic
then srcOr), so lines stop a pixel short of it. The star has no mask. Automatic markers and patterns go through the
markers and patterns in turn, series by series (`3c.03`), the overlay's series first: the main chart's count on
from the overlay's.

**Columns and bars** (`39.03`, `39.04`, `39.00`): each category's share of the plot area holds its series' columns.
A column is 100 units wide, the next series' starts a column and the overlap further on, and half the gap between
clusters lies either side, so a column is the share × 100 / (gap + (series − 1) × (100 + overlap) + 100). A column
runs from where the category axis crosses to its value. Its corners are mapped to pixels; a column no pixel wide
is drawn one pixel wide, and one no pixel high is not drawn. It is filled with its pattern and framed with a
one-pixel pen. A bar chart's categories run up from the plot area's foot (`37.24`). A stacked column runs from the
running total of the values' sizes below it to the total including its own, as a percentage of the category's
total in a 100% chart; it still stands a column and the overlap on from the series before it (`39.04`).

**Areas** (`55.00`, `55.01`): each series is a polygon along its tops and back along its base: the series below's
tops when stacked, which adds the values' sizes; otherwise the category axis's line moved half a pen (10 twips)
toward the series, so the area leaves the axis showing (`55.02`). QuickDraw fills it as
the region its edges make: each edge that is not upright adds inversion points where its left edge, worked out as
DrawLine works it out, steps to another pixel (PutLine). The border runs along the tops, up the left side where the
polygon closes and down the right side a pixel in, but not along the base, clipped to the area's own region, so it
stops short of the polygon's right edge (`55.01`; the right side's pixel **(inferred)**).

Columns and areas are clipped to the plot area; a line chart's lines and markers, and data labels, are not.

**Pies** (`56.00`–`56.05`): the pie is a circle as wide as the plot area's smaller side, in its middle (`56.01`).
Each slice runs from the running total of the values' sizes so far, as a share of 360°, rounded, turned by the
first slice's angle (`56.02`). The pie shrinks to 100/(100 + the largest explosion) of that (`56.0e`, `56.0f`), and
a slice pulled out moves along its middle angle by the diameter × its explosion / 200, in Excel's sines, mapped as a
length (`56.0d`). A slice is filled with FillArc, which only a chart that varies by category does for
automatic slices, and its edge drawn with FrameArc and its two radii. QuickDraw's oval takes, on each scanline,
half-pixel steps out from the centre while (2x/w)² + (2y/h)² stays within 1, y measured to the row's middle; a
frame takes away the oval a pixel inside; an arc keeps what lies between the rays at its start and stop angles,
clockwise from 12 o'clock. A radius runs from the centre to the circle's edge, at half the sum of the rect's sides
plus the side times Excel's sine of the angle (a table of round(32767 × sine) for each whole degree, `56.06`), and
both ends are moved a pixel by the pen according to the radius's quadrant (`56.05`).

**Drop lines and hi-lo lines** are clipped to the plot area widened by 4 pixels each side (`3c.06`). A drop line
(`3c.04`) runs from half a pen off a point to the category axis, its length truncated to pixels, and none is drawn
on the category axis or where the value axis stands (`3c.05`). A line chart draws each series' drop lines as it
draws its lines, an area chart once every area is drawn, from the top series' points. Hi-lo lines (`53.01`) run, at
each category, from half a pen off the lowest point up to the highest, after the lines and before the markers. A
column or area chart with drop lines starts its automatic patterns 2 on, or 4 on when the drop lines are 50% grey
(`3c.03`).

**The chart and plot areas** (`31.03`, `31.07`–`31.09`) are drawn first, the chart area only for the main chart.
Each is its box mapped as a rectangle (`38.04`): its lower left corner mapped as a point, its width and height plus
19 twips mapped as lengths. A shadow (spec §16.5) is the rectangle moved a fortieth of the chart area's smaller
side down and right, in black; then the box is filled with its background, white under a shadow, and framed a
pixel inside its right and bottom edges with its border's pattern. An area with no border, background or shadow
draws nothing.

**Text objects** (`33.1b`): an unbordered object erases its box as a data label does (below); a bordered one moves
its text in by the border's pen and narrows its room by twice the pen. Each line is aligned a pixel inside: left,
right, or centred at (room + 1 − width) / 2. Free text is aligned up and down in its room, the box's height less a
pixel and twice the pen (`33.03`): at the top, at the foot, or centred at (room − text height + 1) / 2. Vertical
text stacks its characters a line apart in a column an M wide (`33.09`), the first baseline two pixels and an
ascent below the top. The renderer draws a text object bold when bit 4 (`0x10`) of its font descriptor's byte 3 is set (spec §16.4),
as `CLOCK.CH`'s hour labels are, though the spec's codes make them italic **(inferred: the capture shows bold)**.

**Arrows** (`52.00`–`52.04`) are text objects (spec §16.7) drawn from the tail, the point of the box nearest the
tip, to the tip, both mapped as points. A head's barbs stand (length + 1) × 100 twips back along the arrow and
(width + 1) × 50 twips either side, mapped as lengths and turned onto the arrow by its direction over its length
in pixels, an integer square root found by Newton's method from the larger side (`52.02`). The shaft runs to the
middle of the barbs under a filled head and to the tip otherwise; a filled head is the triangle of the tip and the
barbs filled with the arrow's pattern, an open one two lines from the barbs to the tip. Every line is moved half
the shaft's weight in pixels left and down.

**The legend** (`57.00`) lays its entries out as the layout does (above): stacked a spacing apart down from the
top of its box on the right or in the corner, or across in columns as wide as the widest entry and the spacing.
An entry's box starts a spacing in from the legend's; its key fills the first half of the key's width (the
legend font's ascent and widest character) a pixel in: a box of the series' pattern as high as the ascent, framed,
for columns, bars, areas and pie slices; a line through the series' marker, half an ascent above the baseline, for
lines; the marker alone for scatter charts. The text starts a pixel short of the key's width in, its first baseline
a pixel and an ascent below the box's top.

**Data labels** (`33.0d`, `33.10`) copy the chart's built-in text objects: −2 for a series' label (point format bit
11), −1 for each point's (bit 10). Each is laid out in a fifth of the chart area's width and placed at a point by
its chart type (`33.05`), then aligned on it (`33.08`): centred, starting there or ending there across, and with
its foot there, centred or its top there up and down. A label erases its box first, from its left edge to a
pixel short of its right and from a pixel below its top **(inferred)**.

* A pie's labels (`56.0b`, `56.0c`) stand 9/8 of the radius out along each slice's middle, centred on that point
  up and down, and across centred when it is within an eighth of the diameter of the middle, otherwise starting
  there on the right and ending there on the left. A series label shows each slice's category; a point label its
  value, or with bit 9 its percentage of the total in 0.00% **(inferred)**.
* An area's series label (`55.03`) is centred halfway up the series between its two middle categories, the
  middles of the tops and bases averaged (`55.02`'s mode 1), and shows the series' name.
* A column's point label (`39.01`) stands 2 pixels above the column's end, centred on it, or below a column that
  runs down; a bar's starts 2 pixels beyond its end. A line or scatter point's label (`53.04`) starts 5 pixels
  right of the point, centred on it up and down. They show the value in General **(inferred)**.

**Labels and titles** sit as the layout put them (above): value labels right-aligned a pixel inside the widest
label's box, which ends two tick lengths left of the plot area, each box a line and a pixel's margin high and
centred on its tick; labels along the foot centred under their place, their box's top two tick lengths below the
axis. A scatter chart's X labels stand at its X scale's major units, and a logarithmic scale's labels are the
powers of ten. Titles are drawn in their saved boxes, each line centred a pixel inside; the automatic title is
centred over the plot area at the top of the chart area.

Drawn this way, all 64 chart windows captured from Excel 1.00's screen match it pixel for pixel, title bar and
frame included: line, column, bar, area, scatter, logarithmic and pie charts, stacked and 100%, with overlays,
exploded slices, legends in every position, data labels, titles, axis titles, Geneva 9 labels and windows of a
dozen sizes, among them charts Excel made itself from a sheet (`ChartRenderer`, `tests/Data/Screens/charts.txt`).
So do charts with drop lines, hi-lo lines, free text and a bordered, shadowed plot area, and `CLOCK.CH`'s
arrows; five charts of other arrows still differ by 16 to 230 pixels, mostly a pixel along parts of a shaft.

## Open questions

* How the 64K ROM that Excel 1.00 also ran on draws a font size with no bitmap. Running it needs the 64K ROM and
  MAME's `mac512k`, neither of which this project has, and a 400K boot disk, which the 64K ROM needs.
* Where Excel 1.03 and 1.5 lay a chart out differently from 1.00: a pixel more for 1.03's value labels of fractions,
  its scatter X labels, and 1.5's automatic titles and axis titles (`Area Drop Lines`, `Scatter`, `BREAK CHART`,
  `Chart 102 Categories`).
* How Excel draws what no capture has yet: stacked columns' and lines' data labels, minor ticks, thicker lines and
  borders, the legend's border and background, and charts in Excel 1.03 and 1.5.
* Where an arrow's shaft steps a pixel away from the line the renderer draws, and how the wider and longer heads
  are rounded.
