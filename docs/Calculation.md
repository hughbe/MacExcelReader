# How Excel 1.x calculates

The format records each formula's last result, but not how Excel worked it out. This records what Excel 1.00
calculates, as measured: 181 formulas in `tests/Data/calc-1.00.txt`, each with the exact double Excel stored
(samples `Calc 1` and `Calc 2`). They are a reference for anyone writing a calculator for these files; the library
itself does not calculate.

## Method

The formulas were built into two worksheets with `dumper sheet --worksheet`, each formula given `FF FF` at offset 8
of its record so that Excel 1.00 puts it on the list of formulas waiting to be recalculated when it opens the file
(spec §12.3), then opened and saved by Driver (docs/Testing.md). Two things that do not work: a macro sheet, whose
formulas Excel calculates only when they run as a macro, and CALCULATE.NOW on formulas that are not waiting.

## Numbers

* The four operators give exactly the IEEE double result: `0.1+0.2` is `3FD3333333333334` (0.30000000000000004,
  shown as 0.3), `100*1.1` is 110.00000000000001 and `1-0.9` 0.09999999999999998, bit for bit what any IEEE double
  arithmetic gives. Excel's floating point runs in the p-code interpreter's own operations (`fadd`, `fmul`, …); it
  does not call the Macintosh's SANE package.
* `^` is computed as exp(y·ln x): `2^0.5*2^0.5` is 1.9999999999999996, as exp(0.5·ln 2) squared gives, where a
  correctly rounded power gives 2.0000000000000004. `0^0` and a negative number to a fractional power are #NUM!.
  `-2^2` is 4 (the minus binds first) and `2^3^2` is 64 (left to right).
* No result is outside the normal double range: overflow (`2^1023*2`, `1E308*10`, `EXP(710)`) is #NUM!, and
  underflow (`2^-1074`, `1E-308/1E10`) is 0, with no denormals. `1/0` and `0/0` are #DIV/0!.
* The built-in functions are Excel's own code and are not correctly rounded: `SQRT(2)` is one unit in the last place
  low (`3FF6A09E667F3BCC`), `EXP(1)` one high, `ATAN(1)*4` and `ATAN2(1,1)` two high. `SIN(PI())` is
  1.2246402351403E-16 where the double π gives 1.2246467991473532E-16, so the argument is reduced with a π more
  precise than a double; `SIN(1E10)` is #NUM!.
* `ROUND` rounds half away from zero on the decimal value: `ROUND(2.5,0)` 3, `ROUND(-2.5,0)` −3, `ROUND(1.005,2)`
  1.01 and `ROUND(2.675,2)` 2.68, though neither is exactly representable. A fractional count of digits is
  truncated: `ROUND(15,-1.5)` is 20. `INT` rounds down (`INT(-2.7)` is −3).
* `MOD` takes the sign of the divisor: `MOD(-7,3)` is 2 and `MOD(7,-3)` −2. `MOD(1E20,7)` is #NUM!.

## Dates and times

* Serial numbers are days from 1 January 1904 (0), up to 31 December 2040: `DATE(2040,1,1)` is 49674 and
  `DATE(2041,1,1)` and `DATE(1903,12,31)` are #NUM!. A two-digit year adds 1900.
* A month or day past the end rolls over (`DATE(1984,13,1)` is 1 January 1985, `DATE(1984,2,30)` 1 March 1984).
* `TIME(25,0,0)` wraps to 1:00; a negative time is #NUM!. `SECOND` rounds to the nearest second.
* `WEEKDAY` counts from Sunday = 1; day 0 is a Friday (6).
* See [Excel bugs and quirks](ExcelBugs.md) for the dates Excel gets wrong.

## Text and conversion

* Text is at most 255 characters: `REPT("x",300)` gives 255.
* `VALUE` and arithmetic accept number text in any form Excel can enter: `"1,234"`, `"$5"`, `"12%"`, `"1E3"`,
  `" 12 "`, dates and times (`VALUE("1/2/84")` is 29221). Other text is #VALUE!, as is text in `AVERAGE` or `MAX`.
* Joining a number to text formats it as General (`"a"&1/3` is `a0.33333333333333`); TRUE joins as `TRUE`.
* `MID` from position 0 or with a negative length is #VALUE!; past the end it is empty.
* `FIXED(0.5,0)` is `1`; `DOLLAR` puts a negative in parentheses.

## Comparisons and logic

* Text compares without regard to case (`"a"="A"`), and as text even when it looks like a number (`"10"<"9"`).
  Numbers sort before text, and text before TRUE and FALSE (`1<"a"`, `"a"<TRUE`, `TRUE>1`).
* TRUE and FALSE are 1 and 0 in arithmetic; `"TRUE"` is TRUE in `AND` and `IF`.
* `CHOOSE` truncates its index; `MATCH` and `SEARCH` ignore case.

## Errors

* An operator returns the error of its first operand that has one: `1/0+NA()` is #DIV/0! and `NA()+1/0` is #N/A.
* `COUNT(1,"a",TRUE,1/0)` is 2: it counts the number and TRUE.
* `INDEX` outside its array is #REF!; `IRR` that does not converge is #NUM!; `STDEV` of one value is #DIV/0!.

## Recalculation and iteration

Measured with sheets whose formulas were all marked waiting (`Order`, `Circ`), or entered by the test harness with
`FORMULA` after `CALCULATION(1,TRUE,max,change)` (`It` sheets), then saved by Excel 1.00 and read back.

* **Order.** One recalculation follows the dependencies, whatever the layout: the chain `A1 = A2+1` … `A9 = A10+1`
  with `A10` = 1, written top to bottom, comes out 10, 9, … 1 at once, as do chains across columns and a `SUM` of
  formulas. The calculate loop (in segment `27`) runs a pass (native code) over each open document with formulas
  waiting, and repeats while a pass reports more to do. A pass checks for keystrokes and mouse clicks, so a long
  recalculation can be interrupted.
* **Circular references, iteration off.** Excel says "Can't resolve circular references" (no alert when a macro
  does it) and leaves the cells in the loop with the values they had; formulas that depend on them are not
  calculated either (`Circ`: `A1 = B1+1`, `B1 = A1+1`, `A3 = A1*10` all stay 0).
* **Iteration on.** Excel repeats the pass until, in one iteration, every value changes by strictly less than the
  maximum change, or until it has run the maximum number of iterations:
  * `A1 = A1+1` gives exactly 100 with the default settings, and 5 with a maximum of 5;
  * `A1 = A1/2+1` gives 1.9990234375 (2 − 2⁻¹⁰): the 11th iteration was the first to change it by less than
    0.001;
  * with a maximum change of 0.5 it gives 1.75 after 3 iterations, because the second iteration's change of
    exactly 0.5 does not stop it.
* **Within an iteration** the cell that started the recalculation goes first, then the cells that depend on it:
  with `B1 = A1+1` in place and `A1 = B1+1` entered, 100 iterations give `A1` 199 and `B1` 200.
* A formula just entered starts from 0, whatever its cell held.
* The settings themselves belong to Excel, not to the sheet: a file's maximum iterations, maximum change and
  iteration flag take effect only when it is the first document opened (spec §6).
