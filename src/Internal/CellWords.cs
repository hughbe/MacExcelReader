namespace MacExcelReader.Internal;

/// <summary>
/// The fields of a cell-table entry's second word (spec §9): data-offset nibble, type, alignment, number format.
/// </summary>
internal static class CellWords
{
    public const ushort TypeMask = 0x0E00;
    public const ushort AlignmentMask = 0x01C0;
    public const ushort NumberFormatMask = 0x003F;

    public static ExcelCellType Type(ushort word1) => (ExcelCellType)((word1 & TypeMask) >> 9);

    public static ushort WithType(ushort word1, ExcelCellType type) => (ushort)((word1 & ~TypeMask) | (((int)type & 7) << 9));

    public static ExcelAlignment Alignment(ushort word1) => (ExcelAlignment)((word1 & AlignmentMask) >> 6);

    public static ushort WithAlignment(ushort word1, ExcelAlignment alignment) => (ushort)((word1 & ~AlignmentMask) | (((int)alignment & 7) << 6));

    public static int NumberFormat(ushort word1) => word1 & NumberFormatMask;

    public static ushort WithNumberFormat(ushort word1, int format) => (ushort)((word1 & ~NumberFormatMask) | format);

    // Bits of the third word (spec §9).
    public const int BoldBit = 0;
    public const int ItalicBit = 1;
    public const int LeftBorderBit = 2;
    public const int RightBorderBit = 3;
    public const int TopBorderBit = 4;
    public const int BottomBorderBit = 5;
    public const int ArrayFunctionBit = 13;
    public const int LockedBit = 14;
    public const int HiddenBit = 15;

    public static bool Has(ushort word2, int bit) => (word2 & (1 << bit)) != 0;
}
