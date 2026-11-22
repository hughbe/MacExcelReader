namespace MacExcelReader.Internal;

/// <summary>
/// The row word of a cell or area reference (spec §15.2): bit 15 marks a relative row, bit 14 a relative column,
/// bits 13–0 hold the row, or in a name definition the offset of a relative row (spec §11).
/// </summary>
internal static class RowWord
{
    public const ushort RowRelative = 0x8000;
    public const ushort ColumnRelative = 0x4000;
    public const ushort Flags = RowRelative | ColumnRelative;
    public const ushort RowMask = 0x3FFF;

    public static int Row(ushort rowWord) => rowWord & RowMask;

    public static bool IsRowRelative(ushort rowWord) => (rowWord & RowRelative) != 0;

    public static bool IsColumnRelative(ushort rowWord) => (rowWord & ColumnRelative) != 0;

    public static ushort Create(int row, bool rowRelative, bool columnRelative)
        => (ushort)((row & RowMask) | (rowRelative ? RowRelative : 0) | (columnRelative ? ColumnRelative : 0));

    /// <summary>
    /// Moves the relative parts of a reference by <paramref name="rows"/> and <paramref name="columns"/>, wrapping
    /// round the sheet's 16,384 rows and 256 columns as Excel does; the absolute parts stay.
    /// </summary>
    public static (ushort RowWord, byte Column) Move(ushort rowWord, byte column, int rows, int columns)
    {
        if (IsRowRelative(rowWord))
        {
            rowWord = (ushort)((rowWord & Flags) | ((rowWord + rows) & RowMask));
        }

        if (IsColumnRelative(rowWord))
        {
            column = (byte)(column + columns);
        }

        return (rowWord, column);
    }
}
