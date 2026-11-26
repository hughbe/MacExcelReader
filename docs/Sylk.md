# How Excel 1.x exports SYLK

Excel's Save As with the SYLK type writes a sheet as SYLK, Microsoft's text interchange format, instead of the
document format the [format specification](Spec.md) describes. This describes what Excel 1.00 to 1.5 write, byte
for byte; `ExcelDocument.ExportSylk` follows it. It uses the spec's conventions and cites code the same way.

Save As and `SAVE.AS(name, type)` take type 1 Normal, 2 SYLK, 3 Text (tab-delimited) and 4 WKS (Lotus 1-2-3).
Excel 1.00 numbers them 0–3 internally and dispatches in `47.05`: Normal goes to the document writer (`47.0b`),
the others through `72.01` to the SYLK writer (`69.00`), the text writer (`6e.00`) and the WKS writer
(`74.00`). The SYLK writer is described here from its code (`69.00`–`69.10`, with `72.20` for format codes) and
checked against 84 samples Excel 1.00 exported and 94 Excel 1.5 exported: a writer that follows this
description produces the same bytes for all of them. The writers of Excel 1.01–1.06 (`6a` in 1.04 and 1.06)
differ from 1.00's in one way: they take TRUE, FALSE and the error names from a string list that a translation
can change. Excel 1.5's (`70`) differs in
three ways, given under **Excel 1.5** below.

SYLK is a text form of the same model, so it also shows how Excel reads each field. The 84 exports
(7,936 cells, 2,853 formulas) were compared with what this specification says the files hold. They disagreed
in two places only, the format-text and whole-row rules, which spec §7 and spec §15.4 now give; every constant, number
format, alignment, style, border, protection flag and name agreed.

**Lines.** Each record is a line of Mac OS Roman text ending in a carriage return. Fields are separated by `;`,
so a `;` inside a format, a text value or a formula is doubled. In `P`, `NN`, `NE` and `C` records a byte below
32 or above 126 is escaped with ESC (`1B`) and:

* `N` and a letter for the 11 characters of Excel's first table (PCOD 0 at `0x2CB6`): `a` for Æ, `q` for æ,
  `i` for Ø, `y` for ø, `j` for Œ, `z` for œ, `{` for ß, `x` for †, `c` for ª, `k` for º and `` ` `` for Ω;
* `N`, an accent letter and the base letter for an accented character (PCOD 0 at `0x2CA5`): `A` grave, `B`
  acute, `C` circumflex, `D` tilde, `H` umlaut, `J` ring, `K` cedilla, so é is ESC `N B e`;
* otherwise the byte's high nibble plus `0x20` and its low nibble plus `0x30`: £ (`A3`) is ESC `* 3`, and the
  active-sheet document name `00` is ESC space `0`.

**Records**, in the order Excel writes them:

| Record | Written |
|--------|---------|
| `ID;PXL;N;E` | Always |
| `P;P<format>` | One per number format: the 19 built-in, then the document's own, as format text (spec §7) |
| `F;P<n>;D<code>10[;S…][;E][;G][;H];N<font> <size>` | The defaults: the format, alignment and style of the cell at the grid's margin corner (spec §6.1), or General with none when the grid has no margin. `E` if the front window shows formulas, `G` if it hides gridlines, `H` if it hides headings (spec §13); then the standard font number and size (spec §6). |
| `B;Y<rows>;X<columns>` | The grid size (state words `0x014` and `0x274`), unless it has no rows |
| `O[;P][;L][;M][;R][;E][;A<n> <change>]` | `P` protected, `L` A1 references, `M` manual calculation, `R` precision as displayed, `E` a macro sheet, `A` iteration with its count and maximum change (spec §6) |
| `F;W<first> <last> <width>` | Each run of columns of one width in characters, except runs of 10: `(pixels + digit / 2 − 5) / digit`, rounded down, where `digit` is the width of 0 in the standard font, and 0 for a hidden column |
| `F;…;C<column>` | A whole-column format: the margin row's cell in that column, for each column, when it differs from the defaults |
| `F;…;R<row>` | A whole-row format: the margin column's cell in that row |
| `NN;N<name>[;F][;G<key>];E<definition>` | Each name with a definition: `F` for a function macro, `G` and its shortcut key for a command macro |
| `F;…` and `C;…` | For each cell, row by row: a format record if it differs from the defaults, then a cell record if it has a value, is unlocked or is hidden |
| `NE;F<document>;S<source>;E<cells>` | After a formula that is only an external reference to absolute cells or a name (`=Budget!$A$1`, `=Budget!Rate`): the document, the cells (`R1C1:R2C2`) or name, and the cell or array range that uses it |
| `E` | Always |

**Formats.** An `F` record compares a cell with the defaults, not with the record before it. `P<n>` and the
format code are written when the format or the alignment differs, `S` when the style differs (even to give an
empty `S`). The code is `F` for a cell, `D` for the defaults, then a letter for the format and its decimals,
then the alignment: `G` General, `L` left, `C` centre, `R` right, `X` fill. The format letter comes from its
positive section (`72.20`): `%` for a number whose scaling is at least two more than its decimals, `$` for one
whose first token after the header is the currency symbol, `I` for another with no decimals, `F` for another
with decimals, `E` for a scientific section, and `G` for anything else, dates included. So `$#,##0.00` is
`F$2`, `0.0%` `F%1` and `m/d/yy` `FG0`. The style letters are written in the order I, D, T, L, B, R: italic,
bold, top, left, bottom and right borders.

**Positions.** `Y` and `X` (1-based) are left out when they equal the last ones written. The `B` record is
written the same way, so it sets them: in a one-column sheet the first cell has no `X`.

**Cells.** `C`, the position, `N` if unlocked, `H` if hidden, then for a value `K` and:

* an integer in decimal, a number as General shows it with room for 21 characters (spec §7; 1E+20 is
  `100000000000000000000`);
* text in quotes, with any quotes in it as they are;
* TRUE, FALSE or an error name.

A formula has `K` with its current value, then `E` and its R1C1 text. A cell of an array formula or table has
instead (spec §12.6): at the anchor, `R` and `C` for the range's last cell, then `M` and the shared expression, or
`T` and TABLE's inputs, such as `TR7C25,R5C24` (`#REF!` for a deleted input, nothing for a missing one). Every
other cell of the range has `I` and the anchor's `R` and `C`. Like `Y` and `X`, `R` and `C` are left out when
they repeat the last ones written.

**R1C1 text.** The formula decompiler's R1C1 form:

* An absolute row or column is `R5`, `C2`. A relative one is an offset with its sign, `R[+2]`, `C[-1]`, or `R`,
  `C` alone for no offset.
* An area that spans every column is rows, `R1:R16384` or `R[-20]:R[+16363]`, and one row alone is written once,
  `R[-20]`; an area that spans every row is columns in the same way.
* In a name definition a relative row or column is already an offset (spec §11). Excel 1.00 reads the row offset with
  the relative-column flag as its sign bit, inverted: with an absolute column, `FF FF` read as −1 gives
  `R[-1]C2`; with a relative column, −3 gives `R[+16381]C[-2]` (`Relative Names` sample; see Excel bugs).
* References inside an external reference are written the same way: `CASHBUDGET!R33C2`.
* Numbers, strings, names and functions are as in A1 style (spec §15.4).

Cell type 7 (spec §9) has no case in the writer, as in every other cell switch.

**Excel 1.5.** Its writer (`70.00`–`70.10`) is 1.00's with these changes:

* Formats are written in 1.5's export form (spec §7): `P;P"$"#,##0\ ;;\("$"#,##0\)`, `P;Pm\/d\/yy`, `P;P@`.
* When the grid has a margin (state word `0x020`, spec §6.1), every cell of the grid gets an `F` record, and each
  has `P`, the format code and `S` whether or not they differ from the defaults: `F;P0;FG0G;S;Y1;X1`. The
  `C` record that follows then needs no position. A grid in the count form gets `F` records as in 1.00.
* The defaults record adds `L` and the gridline and heading colour, counting from 1, when the front window's
  colour is not automatic (window state bytes 127 and 128, spec §13): `F;P0;DG0G10;L3;N3 10` for red.
