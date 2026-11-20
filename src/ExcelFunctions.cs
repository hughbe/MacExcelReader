namespace MacExcelReader;

/// <summary>
/// Describes a worksheet function or macro command (spec §9.3).
/// </summary>
/// <param name="Name">The function name.</param>
/// <param name="MinArguments">The minimum number of arguments.</param>
/// <param name="MaxArguments">The maximum number of arguments.</param>
public sealed record ExcelFunction(string Name, int MinArguments, int MaxArguments);

/// <summary>
/// The functions and macro commands known to Excel 1.03, indexed by function number. Names come
/// from the application's STR# 257 resource and argument counts from its function table.
/// </summary>
public static class ExcelFunctions
{
    private static readonly ExcelFunction[] s_functions =
    [
        new("COUNT", 1, 14), // 0
        new("IF", 2, 3), // 1
        new("ISNA", 1, 1), // 2
        new("ISERROR", 1, 1), // 3
        new("SUM", 1, 14), // 4
        new("AVERAGE", 1, 14), // 5
        new("MIN", 1, 14), // 6
        new("MAX", 1, 14), // 7
        new("ROW", 0, 1), // 8
        new("COLUMN", 0, 1), // 9
        new("NA", 0, 0), // 10
        new("NPV", 2, 14), // 11
        new("STDEV", 1, 14), // 12
        new("DOLLAR", 1, 2), // 13
        new("FIXED", 1, 2), // 14
        new("SIN", 1, 1), // 15
        new("COS", 1, 1), // 16
        new("TAN", 1, 1), // 17
        new("ATAN", 1, 1), // 18
        new("PI", 0, 0), // 19
        new("SQRT", 1, 1), // 20
        new("EXP", 1, 1), // 21
        new("LN", 1, 1), // 22
        new("LOG10", 1, 1), // 23
        new("ABS", 1, 1), // 24
        new("INT", 1, 1), // 25
        new("SIGN", 1, 1), // 26
        new("ROUND", 2, 2), // 27
        new("LOOKUP", 2, 3), // 28
        new("INDEX", 2, 4), // 29
        new("REPT", 2, 2), // 30
        new("MID", 3, 3), // 31
        new("LEN", 1, 1), // 32
        new("VALUE", 1, 1), // 33
        new("TRUE", 0, 0), // 34
        new("FALSE", 0, 0), // 35
        new("AND", 1, 14), // 36
        new("OR", 1, 14), // 37
        new("NOT", 1, 1), // 38
        new("MOD", 2, 2), // 39
        new("DCOUNT", 3, 3), // 40
        new("DSUM", 3, 3), // 41
        new("DAVERAGE", 3, 3), // 42
        new("DMIN", 3, 3), // 43
        new("DMAX", 3, 3), // 44
        new("DSTDEV", 3, 3), // 45
        new("VAR", 1, 14), // 46
        new("DVAR", 3, 3), // 47
        new("TEXT", 2, 2), // 48
        new("LINEST", 1, 2), // 49
        new("TREND", 1, 3), // 50
        new("LOGEST", 1, 2), // 51
        new("GROWTH", 1, 3), // 52
        new("GOTO", 1, 1), // 53
        new("HALT", 0, 0), // 54
        new("RETURN", 0, 1), // 55
        new("COLUMN.WIDTH", 0, 2), // 56
        new("FORMULA", 1, 1), // 57
        new("PV", 3, 5), // 58
        new("FV", 3, 5), // 59
        new("NPER", 3, 5), // 60
        new("PMT", 3, 5), // 61
        new("RATE", 3, 6), // 62
        new("MIRR", 3, 3), // 63
        new("IRR", 1, 2), // 64
        new("RAND", 0, 0), // 65
        new("MATCH", 2, 3), // 66
        new("DATE", 3, 3), // 67
        new("TIME", 3, 3), // 68
        new("DAY", 1, 1), // 69
        new("MONTH", 1, 1), // 70
        new("YEAR", 1, 1), // 71
        new("WEEKDAY", 1, 1), // 72
        new("HOUR", 1, 1), // 73
        new("MINUTE", 1, 1), // 74
        new("SECOND", 1, 1), // 75
        new("NOW", 0, 0), // 76
        new("AREAS", 1, 1), // 77
        new("ROWS", 1, 1), // 78
        new("COLUMNS", 1, 1), // 79
        new("OFFSET", 3, 3), // 80
        new("ABSREF", 2, 2), // 81
        new("RELREF", 2, 2), // 82
        new("ARGUMENT", 0, 3), // 83
        new("SEARCH", 2, 3), // 84
        new("FORMAT.NUMBER", 0, 1), // 85
        new("ALIGNMENT", 0, 1), // 86
        new("STYLE", 0, 2), // 87
        new("BORDER", 0, 5), // 88
        new("CELL.PROTECTION", 0, 2), // 89
        new("CLEAR", 0, 1), // 90
        new("EDIT.DELETE", 0, 1), // 91
        new("INSERT", 0, 1), // 92
        new("CUT", 0, 0), // 93
        new("COPY", 0, 0), // 94
        new("PASTE", 0, 0), // 95
        new("FILL.RIGHT", 0, 0), // 96
        new("FILL.DOWN", 0, 0), // 97
        new("TRANSPOSE", 1, 1), // 98
        new("ERROR", 0, 2), // 99
        new("SELECT", 1, 2), // 100
        new("STEP", 0, 0), // 101
        new("DEFINE.NAME", 0, 4), // 102
        new("TYPE", 1, 1), // 103
        new("ECHO", 0, 1), // 104
        new("NEW", 0, 1), // 105
        new("OPEN", 0, 2), // 106
        new("DELETE.NAME", 0, 1), // 107
        new("SAVE", 0, 0), // 108
        new("A1.R1C1", 0, 1), // 109
        new("FILE.DELETE", 0, 1), // 110
        new("PRINTER.SETUP", 0, 3), // 111
        new("NEW.WINDOW", 0, 0), // 112
        new("SELECT.LAST.CELL", 0, 0), // 113
        new("SHOW.ACTIVE.CELL", 0, 0), // 114
        new("DISPLAY", 0, 3), // 115
        new("FONT", 0, 2), // 116
        new("ACTIVATE", 1, 1), // 117
        new("CLOSE", 0, 0), // 118
        new("SET.NAME", 1, 2), // 119
        new("CALLER", 0, 0), // 120
        new("PASTE.SPECIAL", 0, 3), // 121
        new("OPEN.LINKS", 0, 14), // 122
        new("PAGE.SETUP", 0, 8), // 123
        new("PRINT", 0, 5), // 124
        new("UNDO", 0, 0), // 125
        new("CREATE.NAMES", 0, 2), // 126
        new("FORMULA.FIND", 0, 5), // 127
        new("DATA.FIND", 0, 1), // 128
        new("EXTRACT", 0, 1), // 129
        new("DATA.DELETE", 0, 0), // 130
        new("SET.DATABASE", 0, 0), // 131
        new("SET.CRITERIA", 0, 0), // 132
        new("SORT", 0, 7), // 133
        new("DATA.SERIES", 0, 5), // 134
        new("TABLE", 0, 2), // 135
        new("CALCULATE.NOW", 0, 0), // 136
        new("PRECISION", 0, 1), // 137
        new("CALCULATION", 0, 4), // 138
        new("PROTECT.DOCUMENT", 0, 1), // 139
        new("FORMULA.GOTO", 0, 1), // 140
        new("SET.PRINT.AREA", 0, 0), // 141
        new("SET.PRINT.TITLES", 0, 0), // 142
        new("SET.PAGE.BREAK", 0, 0), // 143
        new("DELETE.FORMAT", 0, 1), // 144
        new("RUN", 0, 1), // 145
        new("DEREF", 1, 1), // 146
        new("SHOW.CLIPBOARD", 0, 0), // 147
        new("WINDOWS", 0, 0), // 148
        new("MESSAGE", 1, 2), // 149
        new("HLINE", 1, 1), // 150
        new("SERIES", 4, 4), // 151
        new("REMOVE.PAGE.BREAK", 0, 0), // 152
        new("SAVE.AS", 0, 2), // 153
        new("DOCUMENTS", 0, 0), // 154
        new("VSCROLL", 1, 1), // 155
        new("ACTIVATE.NEXT", 0, 0), // 156
        new("ACTIVATE.PREV", 0, 0), // 157
        new("DATA.FIND.NEXT", 0, 0), // 158
        new("DATA.FIND.PREV", 0, 0), // 159
        new("FORMULA.FIND.NEXT", 0, 0), // 160
        new("FORMULA.FIND.PREV", 0, 0), // 161
        new("GALLERY.AREA", 0, 1), // 162
        new("GALLERY.LINE", 0, 1), // 163
        new("GALLERY.SCATTER", 0, 1), // 164
        new("GALLERY.COLUMN", 0, 1), // 165
        new("GALLERY.BAR", 0, 1), // 166
        new("GALLERY.PIE", 0, 1), // 167
        new("COMBINATION", 0, 1), // 168
        new("ACTIVE.CELL", 0, 0), // 169
        new("SELECTION", 0, 0), // 170
        new("RESULT", 1, 1), // 171
        new("MAIN.CHART.TYPE", 0, 1), // 172
        new("OVERLAY.CHART.TYPE", 0, 1), // 173
        new("COPY.CHART", 0, 1), // 174
        new("COPY.PICTURE", 0, 0), // 175
        new("7", 1, 4), // 176
        new("ATAN2", 2, 2), // 177
        new("ASIN", 1, 1), // 178
        new("ACOS", 1, 1), // 179
        new("CHOOSE", 2, 14), // 180
        new("HLOOKUP", 3, 3), // 181
        new("VLOOKUP", 3, 3), // 182
        new("LINKS", 0, 1), // 183
        new("ALERT", 2, 2), // 184
        new("INPUT", 2, 3), // 185
        new("BEEP", 0, 0), // 186
        new("ISREF", 1, 1), // 187
        new("GET.FORMULA", 1, 1), // 188
        new("GET.NAME", 1, 1), // 189
        new("LEGEND", 0, 1), // 190
        new("PREFERRED", 0, 0), // 191
        new("HSCROLL", 1, 1), // 192
        new("VPAGE", 1, 1), // 193
        new("SET.VALUE", 2, 2), // 194
        new("FORMULA.ARRAY", 1, 1), // 195
        new("FORMULA.FILL", 1, 1), // 196
        new("SELECT.CHART", 0, 0), // 197
        new("HPAGE", 1, 1), // 198
        new("VLINE", 1, 1), // 199
        new("MOVE", 2, 3), // 200
        new("SIZE", 2, 3), // 201
        new("FULL", 0, 1), // 202
        new("8", 1, 1), // 203
        new("9", 0, 1), // 204
        new("CANCEL.COPY", 0, 0), // 205
        new("UNLOCKED.NEXT", 0, 0), // 206
        new("UNLOCKED.PREV", 0, 0), // 207
    ];

    /// <summary>
    /// Gets the number of functions (208), which Excel also writes as the first word of the data area.
    /// </summary>
    public static int Count => s_functions.Length;

    /// <summary>
    /// Gets the function with the given number, or null if the number is unknown.
    /// </summary>
    public static ExcelFunction? Get(int number)
        => number >= 0 && number < s_functions.Length ? s_functions[number] : null;

    /// <summary>
    /// Finds a function number by name (case-insensitive), or returns -1.
    /// </summary>
    public static int FindNumber(string name)
    {
        for (int i = 0; i < s_functions.Length; i++)
        {
            if (string.Equals(s_functions[i].Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }
}
