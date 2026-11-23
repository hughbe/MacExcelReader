# Excel 1.x bugs and quirks

Behaviour of Excel itself found while writing and testing files: crashes, silent data loss and output that
looks wrong but isn't. Each entry says which Excel showed it and how. The format is in the
[specification](Spec.md); [Testing in a real Excel](Testing.md) describes how files are checked in a real
Excel under Mini vMac.

## Files that break the cell-table margin (Excel 1.03)

Seen in Excel 1.03 under Mini vMac, November 2026.

Excel uses only rows `0 … word[0x014] − 1` and columns `0 … word[0x274] − 1` of the cell table (§6.1). In the
last-index form, which every file Excel 1.03 writes uses, that leaves the table's last row and column as a
margin. Excel's own files never have values there, but Open does not check. A file with values in the margin
makes Excel 1.03:

| Symptom | Example |
|---|---|
| **Hangs** while drawing the window, with the watch cursor showing until the emulator is killed | Any table with one row (word `0x014` = 0) |
| Leaves values in the last row and column undrawn. They are still there: selecting the cell shows them in the formula bar. | Text in A13 of a 13-row table |
| Loads the cells of the last column of a row over one another. One cell ends up with the last one's record and the others are lost when the file is saved. | D1:F1 of a 6-column table all become F1's formula |
| Drops rows from the end of the sheet. The rows are gone when the file is saved. | A 115-row macro sheet re-saved with 89 rows |
| Leaves defined command macros out of Macro > Run | The same 115-row sheet |

The library always leaves the margin blank when it writes the last-index form.

Excel 3.0 also hung while converting a 578-formula sheet written before this fix, which broke the margin. That
the margin caused it is **(inferred)**: the sheet was not retested after the fix.

## Opening a large macro sheet runs out of memory (Excel 1.03)

Seen in Excel 1.03 under Mini vMac. A macro sheet with all 578 formulas of `tests/Data/formulas-1.0.txt`
stopped at 92% of Open with "Out of memory". Sheets of about 100 formulas (around 13 KB) open fine. The limit
depends on the memory given to Excel, so this is a practical limit rather than a format rule. `sheet --split`
exists for it.

## OPEN stops a macro to ask about links (Excel 1.03)

Seen in Excel 1.03. A macro that opens a file linked to other documents, such as `CLOCK.CH`, stops at "Update
references to non-resident sheets?" until someone answers. While the dialog is up, the macro's next commands
can act on the wrong window. Pass 0 as OPEN's second argument (`=OPEN("Disk:File",0)`) to skip the question.

## Macro FORMULA takes only R1C1 text (Excel 1.03)

Seen in Excel 1.03. The macro command `FORMULA("=A1+1")` returns `#VALUE!` and enters nothing;
`FORMULA("=R1C1+1")` works. In R1C1 mode `A1` is not a reference at all, so `=Other!A1` became a reference to
an external *name* `A1`. Text from GET.FORMULA, which is R1C1, can be passed straight back.

## Calculation mode and reference style are global

Seen in Excel 1.03. CALCULATION and A1.R1C1 change a setting for the whole program, not for the active sheet:
every sheet saved afterwards carries the new value in its state block (§6, 0x002 and 0x006), and formulas
typed afterwards are read in the new reference style.

## Out of memory makes OPEN fail quietly (Excel 1.03)

Seen in Excel 1.03. With a large macro sheet open (a 345-row Driver after recompiling 100 formulas), the macro
command OPEN returned `#VALUE!` for a small file, and opening it by hand reported "Out of memory". A macro
that does not check OPEN's result carries on with the wrong window active.

## SYLK export leaves a position out after the grid size (Excel 1.00)

Excel 1.00's SYLK writer writes the grid size, `B;Y<rows>;X<columns>`, through the same code that places cells,
so the first cell record leaves out a `Y` or `X` that equals the grid size: in a sheet one column wide, the first
cell has no `X` (`Format Compiler`). A reader that does not treat `B` as setting the position puts that cell in
the wrong column.

## SYLK export misreads a name's relative row (Excel 1.00)

In a name definition a relative row is a 14-bit offset (spec §11). Excel 1.00's R1C1 writer takes the sign from
the relative-column flag instead of from the offset, so a name defined with a relative row and column three rows
up is written `R[+16381]C[-2]` rather than `R[-3]C[-2]`, while one with an absolute column is right (`R[-1]C2`;
`Relative Names` sample).

## SIGN(-0) is −1 (Excel 1.00)

Excel folds a minus sign into a number constant, so `=SIGN(-0)` stores negative zero (`1F 80 00 00 00 00 00 00 00`),
and SIGN of negative zero is −1 rather than 0. Excel then shows the formula as `=SIGN(0)` (`Calc 1` sample;
docs/Calculation.md).

## A negative number in a narrow column can show as 0 (Excel 1.00)

General writes the minus sign before it knows whether any digits will fit. When the number rounds to nothing in
the room the column has, Excel writes `0` over the sign: −123 shows as `0` in a 12- or 15-pixel column (Geneva
10), where 123 shows as nothing and −7 as `#`, and −0.4 shows as `0` in a 19-pixel column. A positive number
never shows as a 0 it is not (spec Appendix D; `display-1.00.txt`).

## Some DATE arguments give nonsense (Excel 1.00)

`DATE(1904,1,0)` gives 15824.73037037 and `DATE(2040,12,31)` 328.73037037037, neither a date at all; the day
before 1 January 1904 should be #NUM! and the last day of 2040 is 49978. `DATE(1984,0,1)` gives 29701 (26 April
1985) rather than 1 December 1983, and `DATE(1984,1,-1)` gives 1 January 1984 (`Calc 2` sample).

## ACTIVATE needs the window title (Excel 1.00)

Excel 1.00 titles a window with its volume when the document is not on the startup disk (`Ws:F 01 saved`), so
`ACTIVATE("F 01 saved")` fails. The same title is written as the document name in workspace files (Appendix
B).

## GET.FORMULA's whole-row text does not compile back the same (Excel 1.03)

Seen in Excel 1.03. `=A1:IV16384` comes back from GET.FORMULA as `=R[-20]:R[16363]`, and entering that text
gives a different formula: a range of two whole-row areas wrapped in a memory area (`46 … 25 … 25 … 11`).

## Display quirks of GET.FORMULA (Excel 1.03, 1.5)

Seen in Excel 1.03; the whole-row and array-formula ones also in Excel 1.5. These look like mismatches but the formula is right:

* GET.FORMULA returns R1C1 text even when the sheet shows A1 references: `=A1%` in B4 is `=R[-3]C[-1]%`.
* A range of whole rows is shortened: `=A1:IV16384` becomes `=R[-20]:R[16363]`, and `=$A$1:$IV$16384` becomes
  `=R1:R16384`.
* External documents get a leading colon inside the quotes: `=Budget!Rate*2` becomes `=':Budget'!Rate*2`.
* Array formulas come back without braces, and every cell of the range gets the anchor's text, its offsets
  counted from the anchor: `{=A1:A2*2}` in B60:B61 is `=R[-59]C[-1]:R[-58]C[-1]*2` in both cells.
* Small numbers are written out in full: `1.5E-10` becomes `0.00000000015`.

## Loaded formulas keep their stored values (Excel 1.03)

Seen in Excel 1.03 under Mini vMac. A sheet written with every formula's result set to 0 and state word
`0x018` at 0 still showed 0 for `=1/0` and `=SQRT(-1)` after CALCULATE.NOW, in automatic calculation. Excel
recalculates a loaded formula only when something it depends on changes, when it opens a file whose state
word `0x018` is 1 or 2 (§6), or when the formula's own record has a non-zero link at offset 8 other than `FF FE`,
which puts it on the list of formulas waiting to be recalculated (§12.3; `Dirty` experiment in Excel 1.00). A
writer that cannot work out the results should set the state word, or `FF FF` at offset 8 of each formula.

## TABLE with one argument takes a second input it was not given (Excel 1.00, 1.03)

Seen in Excel 1.03, and read from Excel 1.00's code (`4b.05`). The macro command TABLE copies its arguments
into a buffer on the stack without clearing it, and treats each slot whose row is not the "missing" marker
`0x4000` as an input. With one argument, `=TABLE(!Y7)`, the second slot keeps whatever was there before, so
the table gets a column input it was not given: P1 every time in these tests. Write `=TABLE(!Y7,)` for a table
with only a row input.

## Old fixed-argument tokens in Excel 1.5 (deliberate)

Not a bug, but it looks like one. It was read from Excel 1.5's Save code, and the samples bear it out. Excel
1.5 gave FORMULA, ACTIVATE, CLOSE, VSCROLL, HSCROLL, FORMULA.ARRAY and FORMULA.FILL an optional argument. Its
Save still writes a call with only the old minimum number of arguments in the fixed-count form (`0x21`), so
older releases can read it. A pre-release 1.5 build in the samples lacks this and writes `FORMULA(A1)` as `42
01 39`. See §2.3 and §15.4 of the specification.
