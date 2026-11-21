"""Check the results of the coverage sheets' Check macro, as saved by Excel.

usage: python3 check_coverage.py EXCEL.csv OURS.csv [EXCEL.csv OURS.csv ...]

Each pair is `mac-excel-dumper export` of a sheet Excel saved after running Check, and of the sheet as the
library wrote it. Check copies GET.FORMULA's text into column C for every row whose formula text differs from
column A. GET.FORMULA returns R1C1 text, so this converts it back to A1 (relative to the formula in column B)
and reports only the rows that still differ. It also reports rows Excel lost.
"""
import csv
import re
import sys


def column_name(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def to_a1(text, row, column=2):
    def convert(m):
        rr, cc = m.group(1), m.group(2)
        relative_row = rr == '' or rr.startswith('[')
        relative_column = cc == '' or cc.startswith('[')
        r = row + (int(rr[1:-1]) if rr.startswith('[') else 0) if relative_row else int(rr)
        c = column + (int(cc[1:-1]) if cc.startswith('[') else 0) if relative_column else int(cc)
        return f"{'' if relative_column else '$'}{column_name(c)}{'' if relative_row else '$'}{r}"

    def whole_rows(m):
        # Excel writes a range of whole rows, such as A1:IV16384, as R1:R16384.
        rows = []
        for rr in (m.group(1), m.group(2)):
            relative = rr.startswith('[')
            rows.append(f"{'' if relative else '$'}{row + int(rr[1:-1]) if relative else int(rr)}")
        dollar = '' if m.group(1).startswith('[') else '$'
        return f'{dollar}A{rows[0]}:{dollar}IV{rows[1]}'

    text = re.sub(r"(?<![A-Za-z.])R(\[-?\d+\]|\d+):R(\[-?\d+\]|\d+)(?![C\d\[])", whole_rows, text)
    a1 = re.sub(r"(?<![A-Za-z.])R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)(?![A-Za-z(])", convert, text)
    # Excel shows external documents as ':Name'; the coverage sheets write Name or 'Two Words'.
    return re.sub(r"':([^']+)'!", lambda m: m.group(1) + '!' if re.fullmatch(r'\w+', m.group(1)) else f"'{m.group(1)}'!", a1)


def main(paths):
    total = unexplained = 0
    for excel_path, ours_path in zip(paths[::2], paths[1::2]):
        excel = list(csv.reader(open(excel_path, encoding='utf-8')))
        ours = list(csv.reader(open(ours_path, encoding='utf-8')))
        flagged = bad = 0
        for index, cells in enumerate(excel):
            if len(cells) > 2 and cells[2] and cells[0][:1] in '={':
                flagged += 1
                intended = cells[0][1:-1] if cells[0].startswith('{') else cells[0]
                back = to_a1(cells[2], index + 1)
                if back != intended and back.replace('0.00000000015', '1.5E-10') != intended:
                    bad += 1
                    print(f'{excel_path} row {index + 1}: {intended} | Excel: {cells[2]} | as A1: {back}')
        lost = len(ours) - len(excel)
        print(f'{excel_path}: {flagged} flagged, {bad} unexplained' + (f', {lost} rows lost' if lost else ''))
        total += flagged
        unexplained += bad
    print(f'total: {total} flagged, {unexplained} unexplained')


main(sys.argv[1:])
