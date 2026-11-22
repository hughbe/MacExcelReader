namespace MacExcelReader;

/// <summary>
/// Value types, as used for a function's return type and argument types (a bit mask).
/// </summary>
[Flags]
public enum ExcelValueTypes : byte
{
    /// <summary>No value.</summary>
    None = 0,

    /// <summary>A number.</summary>
    Number = 0x01,

    /// <summary>Text.</summary>
    Text = 0x02,

    /// <summary>A logical value.</summary>
    Logical = 0x04,

    /// <summary>A reference.</summary>
    Reference = 0x08,

    /// <summary>An error value.</summary>
    Error = 0x10,

    /// <summary>A macro command's result (no value).</summary>
    Command = 0x20,

    /// <summary>An array.</summary>
    Array = 0x40,

    /// <summary>An empty argument of this type means the argument was not given (spec §15.3).</summary>
    EmptyMeansOmitted = 0x80,
}

/// <summary>
/// Describes a worksheet function or macro command (spec §15.3), from Excel 1.03's function table.
/// </summary>
/// <param name="Name">The function name.</param>
/// <param name="MinArguments">The minimum number of arguments.</param>
/// <param name="MaxArguments">The maximum number of arguments.</param>
/// <param name="ReturnType">What the function can return.</param>
/// <param name="FirstArgumentType">What the first argument accepts.</param>
/// <param name="OtherArgumentType">What the second and later arguments accept.</param>
/// <remarks>The table entry's flags byte also gives <see cref="IsMacroOnly"/>, <see cref="HasDialogForm"/> and <see cref="IsVolatile"/>.</remarks>
public sealed record ExcelFunction(
    string Name,
    int MinArguments,
    int MaxArguments,
    ExcelValueTypes ReturnType,
    ExcelValueTypes FirstArgumentType,
    ExcelValueTypes OtherArgumentType)
{
    private readonly int _flags;

    internal ExcelFunction(string name, int min, int max, int returnType, int firstArgument, int otherArguments, int flags)
        : this(name, min, max, (ExcelValueTypes)returnType, (ExcelValueTypes)firstArgument, (ExcelValueTypes)otherArguments)
    {
        _flags = flags;
    }

    /// <summary>
    /// Gets the low nibble of the function-table entry's second byte: 0x01 macro-only, 0x02 not for cell formulas
    /// (inferred; set on SERIES and the three placeholder functions), 0x04 dialog-box form, 0x08 volatile.
    /// </summary>
    public int TableFlags => _flags;

    /// <summary>Gets whether the function is only available on macro sheets.</summary>
    public bool IsMacroOnly => (_flags & 0x01) != 0;

    /// <summary>Gets whether the command has a dialog-box form, written with a '?' (for example <c>RUN?()</c>).</summary>
    public bool HasDialogForm => (_flags & 0x04) != 0;

    /// <summary>
    /// Gets whether the function is volatile (recalculated every time, like RAND and NOW). Formulas that call
    /// one start with an attribute token <c>19 01</c>.
    /// </summary>
    public bool IsVolatile => (_flags & 0x08) != 0;

    /// <summary>Gets what argument <paramref name="index"/> (zero-based) accepts.</summary>
    public ExcelValueTypes GetArgumentType(int index) => index == 0 ? FirstArgumentType : OtherArgumentType;

    /// <summary>Gets whether the function takes a fixed number of arguments (encoded with token 0x21).</summary>
    public bool HasFixedArguments => MinArguments == MaxArguments;

    /// <summary>Gets the first Excel release that has the function.</summary>
    public ExcelVersion IntroducedIn { get; init; } = ExcelVersion.Excel100;
}

/// <summary>
/// The functions and macro commands of Excel 1.x, indexed by the function number stored in files. Numbers
/// 0–207 come from Excel 1.00–1.03 (names from STR# 257, rules from the function table); 208 was added by
/// Excel 1.04 and 209–291 by Excel 1.5, from Excel 1.5's tables and its file-number translation table.
/// </summary>
public static class ExcelFunctions
{
    private static readonly ExcelFunction[] s_functions =
    [
        new("COUNT", 1, 14, 0x01, 0x59, 0x59, 0x0), // 0
        new("IF", 2, 3, 0x6F, 0x44, 0x7F, 0x0), // 1
        new("ISNA", 1, 1, 0x04, 0x17, 0x00, 0x0), // 2
        new("ISERROR", 1, 1, 0x04, 0x17, 0x00, 0x0), // 3
        new("SUM", 1, 14, 0x01, 0x49, 0x49, 0x0), // 4
        new("AVERAGE", 1, 14, 0x01, 0x49, 0x49, 0x0), // 5
        new("MIN", 1, 14, 0x01, 0x49, 0x49, 0x0), // 6
        new("MAX", 1, 14, 0x01, 0x49, 0x49, 0x0), // 7
        new("ROW", 0, 1, 0x41, 0x08, 0x00, 0x0), // 8
        new("COLUMN", 0, 1, 0x41, 0x08, 0x00, 0x0), // 9
        new("NA", 0, 0, 0x03, 0x00, 0x00, 0x0), // 10
        new("NPV", 2, 14, 0x01, 0x01, 0x49, 0x0), // 11
        new("STDEV", 1, 14, 0x01, 0x49, 0x49, 0x0), // 12
        new("DOLLAR", 1, 2, 0x02, 0x01, 0x01, 0x0), // 13
        new("FIXED", 1, 2, 0x02, 0x01, 0x01, 0x0), // 14
        new("SIN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 15
        new("COS", 1, 1, 0x01, 0x01, 0x00, 0x0), // 16
        new("TAN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 17
        new("ATAN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 18
        new("PI", 0, 0, 0x01, 0x00, 0x00, 0x0), // 19
        new("SQRT", 1, 1, 0x01, 0x01, 0x00, 0x0), // 20
        new("EXP", 1, 1, 0x01, 0x01, 0x00, 0x0), // 21
        new("LN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 22
        new("LOG10", 1, 1, 0x01, 0x01, 0x00, 0x0), // 23
        new("ABS", 1, 1, 0x01, 0x01, 0x00, 0x0), // 24
        new("INT", 1, 1, 0x01, 0x01, 0x00, 0x0), // 25
        new("SIGN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 26
        new("ROUND", 2, 2, 0x01, 0x01, 0x01, 0x0), // 27
        new("LOOKUP", 2, 3, 0x07, 0x07, 0x48, 0x0), // 28
        new("INDEX", 2, 4, 0x4F, 0x48, 0x01, 0x8), // 29
        new("REPT", 2, 2, 0x02, 0x02, 0x01, 0x0), // 30
        new("MID", 3, 3, 0x02, 0x02, 0x01, 0x0), // 31
        new("LEN", 1, 1, 0x01, 0x02, 0x00, 0x0), // 32
        new("VALUE", 1, 1, 0x01, 0x02, 0x00, 0x0), // 33
        new("TRUE", 0, 0, 0x04, 0x00, 0x00, 0x0), // 34
        new("FALSE", 0, 0, 0x04, 0x00, 0x00, 0x0), // 35
        new("AND", 1, 14, 0x04, 0x4C, 0x4C, 0x0), // 36
        new("OR", 1, 14, 0x04, 0x4C, 0x4C, 0x0), // 37
        new("NOT", 1, 1, 0x04, 0x04, 0x00, 0x0), // 38
        new("MOD", 2, 2, 0x01, 0x01, 0x01, 0x0), // 39
        new("DCOUNT", 3, 3, 0x01, 0x08, 0x8B, 0x0), // 40
        new("DSUM", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 41
        new("DAVERAGE", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 42
        new("DMIN", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 43
        new("DMAX", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 44
        new("DSTDEV", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 45
        new("VAR", 1, 14, 0x01, 0x49, 0x49, 0x0), // 46
        new("DVAR", 3, 3, 0x01, 0x08, 0x0B, 0x0), // 47
        new("TEXT", 2, 2, 0x02, 0x01, 0x02, 0x0), // 48
        new("LINEST", 1, 2, 0x40, 0x48, 0x48, 0x0), // 49
        new("TREND", 1, 3, 0x40, 0x48, 0x48, 0x0), // 50
        new("LOGEST", 1, 2, 0x40, 0x48, 0x48, 0x0), // 51
        new("GROWTH", 1, 3, 0x40, 0x48, 0x48, 0x0), // 52
        new("GOTO", 1, 1, 0x20, 0x08, 0x00, 0x1), // 53
        new("HALT", 0, 0, 0x20, 0x00, 0x00, 0x1), // 54
        new("RETURN", 0, 1, 0x20, 0x7F, 0x00, 0x1), // 55
        new("COLUMN.WIDTH", 0, 2, 0x04, 0x81, 0x8A, 0x5), // 56
        new("FORMULA", 1, 1, 0x04, 0x47, 0x00, 0x1), // 57
        new("PV", 3, 5, 0x01, 0x01, 0x01, 0x0), // 58
        new("FV", 3, 5, 0x01, 0x01, 0x01, 0x0), // 59
        new("NPER", 3, 5, 0x01, 0x01, 0x01, 0x0), // 60
        new("PMT", 3, 5, 0x01, 0x01, 0x01, 0x0), // 61
        new("RATE", 3, 6, 0x01, 0x01, 0x01, 0x0), // 62
        new("MIRR", 3, 3, 0x01, 0x48, 0x01, 0x0), // 63
        new("IRR", 1, 2, 0x01, 0x48, 0x01, 0x0), // 64
        new("RAND", 0, 0, 0x01, 0x00, 0x00, 0x8), // 65
        new("MATCH", 2, 3, 0x01, 0x07, 0x49, 0x0), // 66
        new("DATE", 3, 3, 0x01, 0x01, 0x01, 0x0), // 67
        new("TIME", 3, 3, 0x01, 0x01, 0x01, 0x0), // 68
        new("DAY", 1, 1, 0x01, 0x01, 0x00, 0x0), // 69
        new("MONTH", 1, 1, 0x01, 0x01, 0x00, 0x0), // 70
        new("YEAR", 1, 1, 0x01, 0x01, 0x00, 0x0), // 71
        new("WEEKDAY", 1, 1, 0x01, 0x01, 0x00, 0x0), // 72
        new("HOUR", 1, 1, 0x01, 0x01, 0x00, 0x0), // 73
        new("MINUTE", 1, 1, 0x01, 0x01, 0x00, 0x0), // 74
        new("SECOND", 1, 1, 0x01, 0x01, 0x00, 0x0), // 75
        new("NOW", 0, 0, 0x01, 0x00, 0x00, 0x8), // 76
        new("AREAS", 1, 1, 0x01, 0x08, 0x00, 0x8), // 77
        new("ROWS", 1, 1, 0x01, 0x48, 0x00, 0x8), // 78
        new("COLUMNS", 1, 1, 0x01, 0x48, 0x00, 0x8), // 79
        new("OFFSET", 3, 3, 0x08, 0x08, 0x01, 0x1), // 80
        new("ABSREF", 2, 2, 0x08, 0x02, 0x08, 0x1), // 81
        new("RELREF", 2, 2, 0x02, 0x08, 0x08, 0x1), // 82
        new("ARGUMENT", 0, 3, 0x04, 0x82, 0x89, 0x1), // 83
        new("SEARCH", 2, 3, 0x01, 0x02, 0x02, 0x0), // 84
        new("FORMAT.NUMBER", 0, 1, 0x04, 0x02, 0x00, 0x5), // 85
        new("ALIGNMENT", 0, 1, 0x04, 0x01, 0x00, 0x5), // 86
        new("STYLE", 0, 2, 0x04, 0x84, 0x84, 0x5), // 87
        new("BORDER", 0, 5, 0x04, 0x84, 0x84, 0x5), // 88
        new("CELL.PROTECTION", 0, 2, 0x04, 0x84, 0x84, 0x5), // 89
        new("CLEAR", 0, 1, 0x04, 0x01, 0x00, 0x5), // 90
        new("EDIT.DELETE", 0, 1, 0x04, 0x01, 0x00, 0x5), // 91
        new("INSERT", 0, 1, 0x04, 0x01, 0x00, 0x5), // 92
        new("CUT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 93
        new("COPY", 0, 0, 0x04, 0x00, 0x00, 0x1), // 94
        new("PASTE", 0, 0, 0x04, 0x00, 0x00, 0x1), // 95
        new("FILL.RIGHT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 96
        new("FILL.DOWN", 0, 0, 0x04, 0x00, 0x00, 0x1), // 97
        new("TRANSPOSE", 1, 1, 0x40, 0x47, 0x00, 0x0), // 98
        new("ERROR", 0, 2, 0x04, 0x04, 0x08, 0x1), // 99
        new("SELECT", 1, 2, 0x04, 0x8A, 0x8A, 0x1), // 100
        new("STEP", 0, 0, 0x04, 0x00, 0x00, 0x1), // 101
        new("DEFINE.NAME", 0, 4, 0x04, 0x82, 0xDF, 0x5), // 102
        new("TYPE", 1, 1, 0x01, 0x57, 0x00, 0x0), // 103
        new("ECHO", 0, 1, 0x04, 0x04, 0x00, 0x1), // 104
        new("NEW", 0, 1, 0x04, 0x01, 0x00, 0x5), // 105
        new("OPEN", 0, 2, 0x04, 0x02, 0x84, 0x5), // 106
        new("DELETE.NAME", 0, 1, 0x04, 0x02, 0x00, 0x1), // 107
        new("SAVE", 0, 0, 0x04, 0x00, 0x00, 0x1), // 108
        new("A1.R1C1", 0, 1, 0x04, 0x04, 0x00, 0x1), // 109
        new("FILE.DELETE", 0, 1, 0x04, 0x02, 0x00, 0x5), // 110
        new("PRINTER.SETUP", 0, 3, 0x04, 0x81, 0x81, 0x5), // 111
        new("NEW.WINDOW", 0, 0, 0x04, 0x00, 0x00, 0x1), // 112
        new("SELECT.LAST.CELL", 0, 0, 0x04, 0x00, 0x00, 0x1), // 113
        new("SHOW.ACTIVE.CELL", 0, 0, 0x04, 0x00, 0x00, 0x1), // 114
        new("DISPLAY", 0, 3, 0x04, 0x84, 0x84, 0x5), // 115
        new("FONT", 0, 2, 0x04, 0x82, 0x81, 0x5), // 116
        new("ACTIVATE", 1, 1, 0x04, 0x02, 0x00, 0x1), // 117
        new("CLOSE", 0, 0, 0x04, 0x00, 0x00, 0x1), // 118
        new("SET.NAME", 1, 2, 0x04, 0x02, 0x5F, 0x1), // 119
        new("CALLER", 0, 0, 0x08, 0x00, 0x00, 0x1), // 120
        new("PASTE.SPECIAL", 0, 3, 0x04, 0x81, 0x85, 0x5), // 121
        new("OPEN.LINKS", 0, 14, 0x04, 0x52, 0x52, 0x5), // 122
        new("PAGE.SETUP", 0, 8, 0x04, 0x82, 0x87, 0x5), // 123
        new("PRINT", 0, 5, 0x04, 0x81, 0x81, 0x5), // 124
        new("UNDO", 0, 0, 0x04, 0x00, 0x00, 0x1), // 125
        new("CREATE.NAMES", 0, 2, 0x04, 0x84, 0x84, 0x5), // 126
        new("FORMULA.FIND", 0, 5, 0x04, 0x82, 0x81, 0x5), // 127
        new("DATA.FIND", 0, 1, 0x04, 0x04, 0x00, 0x1), // 128
        new("EXTRACT", 0, 1, 0x04, 0x04, 0x00, 0x5), // 129
        new("DATA.DELETE", 0, 0, 0x04, 0x00, 0x00, 0x5), // 130
        new("SET.DATABASE", 0, 0, 0x04, 0x00, 0x00, 0x1), // 131
        new("SET.CRITERIA", 0, 0, 0x04, 0x00, 0x00, 0x1), // 132
        new("SORT", 0, 7, 0x04, 0x81, 0x8B, 0x5), // 133
        new("DATA.SERIES", 0, 5, 0x04, 0x81, 0x81, 0x5), // 134
        new("TABLE", 0, 2, 0x04, 0x8A, 0x8A, 0x5), // 135
        new("CALCULATE.NOW", 0, 0, 0x04, 0x00, 0x00, 0x1), // 136
        new("PRECISION", 0, 1, 0x04, 0x04, 0x00, 0x1), // 137
        new("CALCULATION", 0, 4, 0x04, 0x81, 0x85, 0x5), // 138
        new("PROTECT.DOCUMENT", 0, 1, 0x04, 0x04, 0x00, 0x5), // 139
        new("FORMULA.GOTO", 0, 1, 0x04, 0x0A, 0x00, 0x5), // 140
        new("SET.PRINT.AREA", 0, 0, 0x04, 0x00, 0x00, 0x1), // 141
        new("SET.PRINT.TITLES", 0, 0, 0x04, 0x00, 0x00, 0x1), // 142
        new("SET.PAGE.BREAK", 0, 0, 0x04, 0x00, 0x00, 0x1), // 143
        new("DELETE.FORMAT", 0, 1, 0x04, 0x02, 0x00, 0x5), // 144
        new("RUN", 0, 1, 0x0C, 0x0A, 0x00, 0x5), // 145
        new("DEREF", 1, 1, 0x57, 0x08, 0x00, 0x1), // 146
        new("SHOW.CLIPBOARD", 0, 0, 0x04, 0x00, 0x00, 0x1), // 147
        new("WINDOWS", 0, 0, 0x40, 0x00, 0x00, 0x1), // 148
        new("MESSAGE", 1, 2, 0x04, 0x84, 0x02, 0x1), // 149
        new("HLINE", 1, 1, 0x04, 0x01, 0x00, 0x1), // 150
        new("SERIES", 4, 4, 0x04, 0x82, 0xC9, 0x2), // 151
        new("REMOVE.PAGE.BREAK", 0, 0, 0x04, 0x00, 0x00, 0x1), // 152
        new("SAVE.AS", 0, 2, 0x04, 0x82, 0x81, 0x5), // 153
        new("DOCUMENTS", 0, 0, 0x40, 0x00, 0x00, 0x1), // 154
        new("VSCROLL", 1, 1, 0x04, 0x01, 0x00, 0x1), // 155
        new("ACTIVATE.NEXT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 156
        new("ACTIVATE.PREV", 0, 0, 0x04, 0x00, 0x00, 0x1), // 157
        new("DATA.FIND.NEXT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 158
        new("DATA.FIND.PREV", 0, 0, 0x04, 0x00, 0x00, 0x1), // 159
        new("FORMULA.FIND.NEXT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 160
        new("FORMULA.FIND.PREV", 0, 0, 0x04, 0x00, 0x00, 0x1), // 161
        new("GALLERY.AREA", 0, 1, 0x04, 0x01, 0x00, 0x5), // 162
        new("GALLERY.LINE", 0, 1, 0x04, 0x01, 0x00, 0x5), // 163
        new("GALLERY.SCATTER", 0, 1, 0x04, 0x01, 0x00, 0x5), // 164
        new("GALLERY.COLUMN", 0, 1, 0x04, 0x01, 0x00, 0x5), // 165
        new("GALLERY.BAR", 0, 1, 0x04, 0x01, 0x00, 0x5), // 166
        new("GALLERY.PIE", 0, 1, 0x04, 0x01, 0x00, 0x5), // 167
        new("COMBINATION", 0, 1, 0x04, 0x01, 0x00, 0x5), // 168
        new("ACTIVE.CELL", 0, 0, 0x08, 0x00, 0x00, 0x1), // 169
        new("SELECTION", 0, 0, 0x08, 0x00, 0x00, 0x1), // 170
        new("RESULT", 1, 1, 0x04, 0x01, 0x00, 0x1), // 171
        new("MAIN.CHART.TYPE", 0, 1, 0x04, 0x01, 0x00, 0x5), // 172
        new("OVERLAY.CHART.TYPE", 0, 1, 0x04, 0x01, 0x00, 0x5), // 173
        new("COPY.CHART", 0, 1, 0x04, 0x01, 0x00, 0x5), // 174
        new("COPY.PICTURE", 0, 0, 0x04, 0x00, 0x00, 0x1), // 175
        new("7", 1, 4, 0x00, 0x00, 0x00, 0x2), // 176
        new("ATAN2", 2, 2, 0x01, 0x01, 0x01, 0x0), // 177
        new("ASIN", 1, 1, 0x01, 0x01, 0x00, 0x0), // 178
        new("ACOS", 1, 1, 0x01, 0x01, 0x00, 0x0), // 179
        new("CHOOSE", 2, 14, 0x6F, 0x41, 0x7F, 0x0), // 180
        new("HLOOKUP", 3, 3, 0x07, 0x07, 0x49, 0x0), // 181
        new("VLOOKUP", 3, 3, 0x07, 0x07, 0x49, 0x0), // 182
        new("LINKS", 0, 1, 0x40, 0x02, 0x00, 0x1), // 183
        new("ALERT", 2, 2, 0x04, 0x02, 0x01, 0x1), // 184
        new("INPUT", 2, 3, 0x7F, 0x02, 0x83, 0x1), // 185
        new("BEEP", 0, 0, 0x04, 0x00, 0x00, 0x1), // 186
        new("ISREF", 1, 1, 0x04, 0x5F, 0x00, 0x0), // 187
        new("GET.FORMULA", 1, 1, 0x02, 0x08, 0x00, 0x1), // 188
        new("GET.NAME", 1, 1, 0x02, 0x02, 0x00, 0x1), // 189
        new("LEGEND", 0, 1, 0x04, 0x04, 0x00, 0x1), // 190
        new("PREFERRED", 0, 0, 0x04, 0x00, 0x00, 0x1), // 191
        new("HSCROLL", 1, 1, 0x04, 0x01, 0x00, 0x1), // 192
        new("VPAGE", 1, 1, 0x04, 0x01, 0x00, 0x1), // 193
        new("SET.VALUE", 2, 2, 0x04, 0x08, 0x57, 0x1), // 194
        new("FORMULA.ARRAY", 1, 1, 0x04, 0x47, 0x00, 0x1), // 195
        new("FORMULA.FILL", 1, 1, 0x04, 0x47, 0x00, 0x1), // 196
        new("SELECT.CHART", 0, 0, 0x04, 0x00, 0x00, 0x1), // 197
        new("HPAGE", 1, 1, 0x04, 0x01, 0x00, 0x1), // 198
        new("VLINE", 1, 1, 0x04, 0x01, 0x00, 0x1), // 199
        new("MOVE", 2, 3, 0x04, 0x81, 0x83, 0x1), // 200
        new("SIZE", 2, 3, 0x04, 0x81, 0x83, 0x1), // 201
        new("FULL", 0, 1, 0x04, 0x04, 0x00, 0x1), // 202
        new("8", 1, 1, 0x00, 0x00, 0x00, 0x2), // 203
        new("9", 0, 1, 0x00, 0x00, 0x00, 0x2), // 204
        new("CANCEL.COPY", 0, 0, 0x04, 0x00, 0x00, 0x1), // 205
        new("UNLOCKED.NEXT", 0, 0, 0x04, 0x00, 0x00, 0x1), // 206
        new("UNLOCKED.PREV", 0, 0, 0x04, 0x00, 0x00, 0x1), // 207
    ];

    // Functions added after Excel 1.03, by file number from 208. Gaps are numbers Excel 1.5 never assigns.
    private static readonly ExcelFunction?[] s_later =
    [
        new("CHANGE.LINK", 0, 2, 0x00, 0x02, 0x02, 0x5) { IntroducedIn = ExcelVersion.Excel104 }, // 208
        new("DIALOG.BOX", 1, 1, 0x05, 0x0A, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 209
        new("ADD.MENU", 2, 2, 0x01, 0x01, 0x08, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 210
        new("ADD.COMMAND", 3, 3, 0x01, 0x01, 0x0B, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 211
        new("ADD.BAR", 0, 0, 0x01, 0x00, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 212
        new("SHOW.BAR", 0, 1, 0x04, 0x01, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 213
        new("ENABLE.COMMAND", 4, 4, 0x04, 0x01, 0x07, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 214
        new("CHECK.COMMAND", 4, 4, 0x04, 0x01, 0x07, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 215
        new("RENAME.COMMAND", 4, 4, 0x04, 0x01, 0x03, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 216
        new("DELETE.MENU", 2, 2, 0x04, 0x01, 0x03, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 217
        new("DELETE.COMMAND", 3, 3, 0x04, 0x01, 0x03, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 218
        new("GET.BAR", 0, 0, 0x01, 0x00, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 219
        new("SLN", 3, 3, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 220
        new("SYD", 4, 4, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 221
        new("DDB", 4, 4, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 222
        new("IPMT", 4, 6, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 223
        new("PPMT", 4, 6, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 224
        new("REPLACE", 4, 4, 0x02, 0x02, 0x02, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 225
        new("SUBSTITUTE", 3, 4, 0x02, 0x02, 0x02, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 226
        new("CODE", 1, 1, 0x01, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 227
        new("FIND", 2, 3, 0x01, 0x02, 0x02, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 228
        new("MDETERM", 1, 1, 0x01, 0x40, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 229
        new("MINVERSE", 1, 1, 0x40, 0x40, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 230
        new("MMULT", 2, 2, 0x40, 0x40, 0x40, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 231
        new("LOG", 1, 2, 0x01, 0x01, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 232
        null, // 233
        new("WAIT", 1, 1, 0x04, 0x01, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 234
        new("SPLIT", 0, 2, 0x04, 0x81, 0x81, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 235
        new("FREEZE.PANES", 0, 1, 0x04, 0x04, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 236
        new("GET.CELL", 1, 2, 0x47, 0x01, 0x08, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 237
        new("GET.DOCUMENT", 1, 2, 0x47, 0x01, 0x02, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 238
        new("GET.WINDOW", 1, 2, 0x47, 0x01, 0x82, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 239
        new("GET.WORKSPACE", 1, 1, 0x07, 0x01, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 240
        new("GET.DEF", 1, 2, 0x02, 0x02, 0x02, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 241
        new("FILES", 0, 2, 0x40, 0x02, 0x02, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 242
        new("DIRECTORY", 0, 1, 0x02, 0x02, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 243
        new("QUIT", 0, 0, 0x04, 0x00, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 244
        null, // 245
        null, // 246
        null, // 247
        null, // 248
        null, // 249
        null, // 250
        new("CHAR", 1, 1, 0x02, 0x01, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 251
        new("LOWER", 1, 1, 0x02, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 252
        new("UPPER", 1, 1, 0x02, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 253
        new("PROPER", 1, 1, 0x02, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 254
        null, // 255
        new("LEFT", 1, 2, 0x02, 0x02, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 256
        new("RIGHT", 1, 2, 0x02, 0x02, 0x01, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 257
        new("EXACT", 2, 2, 0x04, 0x02, 0x02, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 258
        new("TRIM", 1, 1, 0x02, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 259
        new("ISERR", 1, 1, 0x04, 0x17, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 260
        new("ISTEXT", 1, 1, 0x04, 0x1F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 261
        new("ISNUMBER", 1, 1, 0x04, 0x1F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 262
        new("ISBLANK", 1, 1, 0x04, 0x1F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 263
        new("T", 1, 1, 0x02, 0x0F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 264
        new("N", 1, 1, 0x01, 0x0F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 265
        new("DATEVALUE", 1, 1, 0x01, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 266
        new("TIMEVALUE", 1, 1, 0x01, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 267
        new("CLEAN", 1, 1, 0x02, 0x02, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 268
        new("COUNTA", 1, 14, 0x01, 0x59, 0x59, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 269
        new("PRODUCT", 1, 14, 0x01, 0x49, 0x49, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 270
        new("FACT", 1, 1, 0x01, 0x01, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 271
        new("DPRODUCT", 3, 3, 0x01, 0x08, 0x0B, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 272
        new("ISNONTEXT", 1, 1, 0x04, 0x1F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 273
        new("TRUNC", 1, 1, 0x01, 0x01, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 274
        new("ISLOGICAL", 1, 1, 0x04, 0x1F, 0x00, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 275
        new("DCOUNTA", 3, 3, 0x01, 0x08, 0x8B, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 276
        new("STDEVP", 1, 14, 0x01, 0x49, 0x49, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 277
        new("DSTDEVP", 3, 3, 0x01, 0x08, 0x0B, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 278
        new("VARP", 1, 14, 0x01, 0x49, 0x49, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 279
        new("DVARP", 3, 3, 0x01, 0x08, 0x0B, 0x0) { IntroducedIn = ExcelVersion.Excel150 }, // 280
        null, // 281
        null, // 282
        new("DELETE.BAR", 1, 1, 0x04, 0x01, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 283
        new("FOR", 3, 4, 0x20, 0x02, 0x01, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 284
        new("WHILE", 1, 1, 0x20, 0x04, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 285
        new("BREAK", 0, 0, 0x20, 0x00, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 286
        new("NEXT", 0, 0, 0x20, 0x00, 0x00, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 287
        new("REFTEXT", 1, 2, 0x02, 0x08, 0x04, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 288
        new("TEXTREF", 1, 2, 0x08, 0x02, 0x04, 0x1) { IntroducedIn = ExcelVersion.Excel150 }, // 289
        new("INDIRECT", 1, 2, 0x57, 0x02, 0x04, 0x8) { IntroducedIn = ExcelVersion.Excel150 }, // 290
        new("CELL", 1, 2, 0x07, 0x02, 0x08, 0x8) { IntroducedIn = ExcelVersion.Excel150 }, // 291
    ];

    // Functions whose argument rules Excel 1.5 changed (from Excel 1.5's function table).
    private static readonly Dictionary<int, ExcelFunction> s_excel15Rules = new()
    {
        [48] = new("TEXT", 2, 2, 0x02, 0x07, 0x02, 0x0),
        [57] = new("FORMULA", 1, 2, 0x04, 0xC7, 0x8A, 0x1),
        [106] = new("OPEN", 0, 3, 0x04, 0x02, 0x84, 0x5),
        [115] = new("DISPLAY", 0, 5, 0x04, 0x84, 0x85, 0x5),
        [117] = new("ACTIVATE", 1, 2, 0x04, 0x82, 0x01, 0x1),
        [118] = new("CLOSE", 0, 1, 0x04, 0x04, 0x00, 0x1),
        [124] = new("PRINT", 0, 9, 0x04, 0x81, 0x85, 0x5),
        [155] = new("VSCROLL", 1, 2, 0x04, 0x01, 0x04, 0x1),
        [188] = new("GET.FORMULA", 1, 1, 0x02, 0x0A, 0x00, 0x1),
        [192] = new("HSCROLL", 1, 2, 0x04, 0x01, 0x04, 0x1),
        [195] = new("FORMULA.ARRAY", 1, 2, 0x04, 0xC7, 0x8A, 0x1),
        [196] = new("FORMULA.FILL", 1, 2, 0x04, 0xC7, 0x8A, 0x1),
    };

    /// <summary>
    /// Gets the number of functions in Excel 1.00–1.03 (208), the function count those versions write at the
    /// start of the data area.
    /// </summary>
    public static int Count => s_functions.Length;

    /// <summary>Gets the highest function number any Excel 1.x release uses (291, in Excel 1.5).</summary>
    public static int MaxNumber => s_functions.Length + s_later.Length - 1;

    /// <summary>
    /// Gets the function with the given number, as Excel 1.00–1.03 define it, or (for numbers above 207) as the
    /// release that added it defines it. Returns null for numbers no release uses.
    /// </summary>
    public static ExcelFunction? Get(int number)
    {
        if (number >= 0 && number < s_functions.Length)
        {
            return s_functions[number];
        }

        int later = number - s_functions.Length;
        return later >= 0 && later < s_later.Length ? s_later[later] : null;
    }

    /// <summary>
    /// Gets the function with the given number as a particular release defines it, or null if that release
    /// does not have it.
    /// </summary>
    public static ExcelFunction? Get(int number, ExcelVersion version)
    {
        var function = Get(number);
        if (function is null || function.IntroducedIn > version)
        {
            return null;
        }

        if (version >= ExcelVersion.Excel150 && s_excel15Rules.TryGetValue(number, out var changed))
        {
            return changed;
        }

        // Excel 1.04 already gave OPEN its third argument (its function table, spec §2.1).
        return version >= ExcelVersion.Excel104 && number == 106 ? s_excel15Rules[106] : function;
    }

    /// <summary>
    /// Finds a function number by name (case-insensitive) among the functions of <paramref name="version"/>,
    /// or returns -1.
    /// </summary>
    public static int FindNumber(string name, ExcelVersion version = ExcelVersion.Excel150)
    {
        ArgumentNullException.ThrowIfNull(name);
        return s_byName.TryGetValue(name, out int number) && Get(number)!.IntroducedIn <= version ? number : -1;
    }

    private static readonly Dictionary<string, int> s_byName = BuildNameIndex();

    private static Dictionary<string, int> BuildNameIndex()
    {
        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i <= MaxNumber; i++)
        {
            if (Get(i) is { } function)
            {
                index.TryAdd(function.Name, i);
            }
        }

        return index;
    }
}
