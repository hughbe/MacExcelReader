namespace MacExcelReader;

/// <summary>
/// The type of a cell (bits 11–9 of the second cell-table word, spec §9).
/// </summary>
public enum ExcelCellType : byte
{
    /// <summary>The cell has no value; it may still carry formatting.</summary>
    Empty = 0,

    /// <summary>A 16-bit integer stored in the cell table.</summary>
    Integer = 1,

    /// <summary>A double-precision number.</summary>
    Number = 2,

    /// <summary>A text constant.</summary>
    Text = 3,

    /// <summary>A type not seen in any known file; it has no cell data.</summary>
    Unknown4 = 4,

    /// <summary>A formula whose current value is a number.</summary>
    NumericFormula = 5,

    /// <summary>A formula whose current value is text, a logical value or an error.</summary>
    Formula = 6,

    /// <summary>A type not seen in any known file.</summary>
    Unknown7 = 7,
}

/// <summary>
/// Cell alignment (bits 8–6 of the second cell-table word).
/// </summary>
public enum ExcelAlignment : byte
{
    /// <summary>General alignment.</summary>
    General = 0,

    /// <summary>Left aligned.</summary>
    Left = 1,

    /// <summary>Centered.</summary>
    Center = 2,

    /// <summary>Right aligned.</summary>
    Right = 3,

    /// <summary>Fill (repeat to fill the cell).</summary>
    Fill = 4,
}

/// <summary>
/// The kind of document.
/// </summary>
public enum ExcelDocumentKind : byte
{
    /// <summary>A worksheet (Finder type XLBN).</summary>
    Worksheet = 0x10,

    /// <summary>A chart; charts use a separate format, see <see cref="Charts.ChartDocument"/>.</summary>
    Chart = 0x20,

    /// <summary>A macro sheet (Finder type XLPG).</summary>
    MacroSheet = 0x40,
}
