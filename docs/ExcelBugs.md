# Excel 1.x bugs and quirks

Behaviour of Excel itself found while writing and testing files: crashes, silent data loss and output that
looks wrong but isn't. Each entry says how it was found and how sure it is. The format is in the
[specification](Spec.md); [Testing in a real Excel](Testing.md) describes how files are
checked in a real Excel under Mini vMac.

Labels: **Confirmed** means reproduced in a real Excel. **Inferred** means read from the code or from one
observation, not reproduced.

## Files that break the cell-table margin (Excel 1.03)

**Confirmed** in Excel 1.03 under Mini vMac, November 2026.

Excel uses only rows `0 … word[0x014] − 1` and columns `0 … word[0x274] − 1` of the cell table (§6.1). In
the last-index form, which every file Excel 1.03 writes uses, that leaves the table's last row and column as
a margin. Excel's own files never have values there, but Open does not check. A file with values in the
margin makes Excel 1.03:

| Symptom | Example |
|---|---|
| **Hangs** while drawing the window, with the watch cursor showing until the emulator is killed | Any table with one row (word `0x014` = 0) |
| Leaves values in the last row and column undrawn. They are still there: selecting the cell shows them in the formula bar. | Text in A13 of a 13-row table |
| Loads the cells of the last column of a row over one another. One cell ends up with the last one's record and the others are lost when the file is saved. | D1:F1 of a 6-column table all become F1's formula |
| Drops rows from the end of the sheet. The rows are gone when the file is saved. | A 115-row macro sheet re-saved with 89 rows |
| Leaves defined command macros out of Macro > Run | The same 115-row sheet |

The library always leaves the margin blank when it writes the last-index form.

**Inferred:** Excel 3.0 also hung while converting a 578-formula sheet written before this fix. That sheet
broke the margin, but it was not retested after the fix.

## Opening a large macro sheet runs out of memory (Excel 1.03)

**Confirmed** in Excel 1.03 under Mini vMac. A macro sheet with all 578 formulas of `tests/Data/formulas-1.0.txt` stopped at 92% of
Open with "Out of memory". Sheets of about 100 formulas (around 13 KB) open fine. The limit depends on the
memory given to Excel, so this is a practical limit rather than a format rule. `sheet --split` exists for it.

## OPEN stops a macro to ask about links (Excel 1.03)

**Confirmed.** A macro that opens a file linked to other documents, such as `CLOCK.CH`, stops at "Update
references to non-resident sheets?" until someone answers. While the dialog is up, the macro's next commands
can act on the wrong window. Pass 0 as OPEN's second argument (`=OPEN("Disk:File",0)`) to skip the question.

## Display quirks of GET.FORMULA (Excel 1.03)

**Confirmed.** These look like mismatches but the formula is right:

* GET.FORMULA returns R1C1 text even when the sheet shows A1 references: `=A1%` in B4 is `=R[-3]C[-1]%`.
* A range of whole rows is shortened: `=A1:IV16384` becomes `=R[-20]:R[16363]`, and `=$A$1:$IV$16384`
  becomes `=R1:R16384`.
* External documents get a leading colon inside the quotes: `=Budget!Rate*2` becomes `=':Budget'!Rate*2`.
* Array formulas come back without braces: `{=A1:A2*2}` becomes `=R[-59]C[-1]:R[-58]C[-1]*2`.
* Small numbers are written out in full: `1.5E-10` becomes `0.00000000015`.

## Old fixed-argument tokens in Excel 1.5 (deliberate)

Not a bug, but it looks like one. It was read from Excel 1.5's Save code, and the samples bear it out. Excel
1.5 gave FORMULA, ACTIVATE, CLOSE, VSCROLL, HSCROLL, FORMULA.ARRAY and FORMULA.FILL an optional argument.
Its Save still writes a call with only the old minimum number of arguments in the fixed-count form (`0x21`),
so older releases can read it. A pre-release 1.5 build in the samples lacks this and writes `FORMULA(A1)` as
`42 01 39`. See §2.3 and §15.4 of the specification.
