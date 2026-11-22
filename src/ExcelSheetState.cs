using System.Buffers.Binary;

namespace MacExcelReader;

/// <summary>
/// The 630-byte sheet state block (spec §5): calculation settings, grid extent, column widths,
/// protection and the standard font. The raw bytes are kept so unknown fields round-trip; the
/// properties read and write the fields whose meaning is known.
/// </summary>
public sealed class ExcelSheetState
{
    /// <summary>
    /// The size of the state block in bytes.
    /// </summary>
    public const int Size = 630;

    /// <summary>
    /// The number of columns whose widths are stored.
    /// </summary>
    public const int ColumnCount = 256;

    private const int MaxIterationsOffset = 0x000;
    private const int CalculationModeOffset = 0x002;
    private const int PrecisionOffset = 0x004;
    private const int ReferenceStyleOffset = 0x006;
    private const int MaxChangeOffset = 0x00A;
    private const int IterationOffset = 0x012;
    private const int LastRowOffset = 0x014;
    private const int LastVisibleColumnOffset = 0x016;
    private const int RecalculationStateOffset = 0x018;
    private const int ProtectedOffset = 0x01E;
    private const int NonEmptyOffset = 0x020;
    private const int ColumnWidthsOffset = 0x042;
    private const int PasswordHashOffset = 0x242;
    private const int StandardFontOffset = 0x244;
    private const int StandardFontSizeOffset = 0x246;
    private const int LastColumnOffset = 0x274;

    private readonly byte[] _data;

    /// <summary>
    /// Initializes a state block from raw bytes.
    /// </summary>
    public ExcelSheetState(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        if (data.Length != Size)
        {
            throw new ArgumentException($"The state block must be {Size} bytes.", nameof(data));
        }

        _data = data;
    }

    /// <summary>
    /// Gets the raw bytes of the block. Changes to the array are reflected in the document.
    /// </summary>
    public byte[] RawData => _data;

    private ushort GetWord(int offset) => BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(offset));

    private void SetWord(int offset, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(_data.AsSpan(offset), value);

    /// <summary>
    /// Gets or sets the maximum number of iterations.
    /// </summary>
    public ushort MaxIterations
    {
        get => GetWord(MaxIterationsOffset);
        set => SetWord(MaxIterationsOffset, value);
    }

    /// <summary>Gets or sets the calculation mode (spec §6, 0x002).</summary>
    public ExcelCalculationMode CalculationMode
    {
        get => (ExcelCalculationMode)unchecked((short)GetWord(CalculationModeOffset));
        set => SetWord(CalculationModeOffset, unchecked((ushort)(short)value));
    }

    /// <summary>
    /// Gets or sets whether calculation uses full precision (true) or precision as displayed (false).
    /// </summary>
    public bool FullPrecision
    {
        get => GetWord(PrecisionOffset) != 0;
        set => SetWord(PrecisionOffset, (ushort)(value ? 1 : 0));
    }

    /// <summary>Gets or sets the reference style formulas are shown in (spec §6, 0x006).</summary>
    public ExcelReferenceStyle ReferenceStyle
    {
        get => (ExcelReferenceStyle)GetWord(ReferenceStyleOffset);
        set => SetWord(ReferenceStyleOffset, (ushort)value);
    }

    /// <summary>
    /// Gets or sets the maximum change for iteration.
    /// </summary>
    public double MaxChange
    {
        get => BinaryPrimitives.ReadDoubleBigEndian(_data.AsSpan(MaxChangeOffset));
        set => BinaryPrimitives.WriteDoubleBigEndian(_data.AsSpan(MaxChangeOffset), value);
    }

    /// <summary>
    /// Gets or sets whether iteration is enabled.
    /// </summary>
    public bool Iteration
    {
        get => GetWord(IterationOffset) != 0;
        set => SetWord(IterationOffset, (ushort)(value ? 1 : 0));
    }

    /// <summary>
    /// Gets or sets the index of the last column with a non-zero width.
    /// </summary>
    public ushort LastVisibleColumn
    {
        get => GetWord(LastVisibleColumnOffset);
        set => SetWord(LastVisibleColumnOffset, value);
    }

    /// <summary>
    /// Gets or sets the recalculation state: 0 clean, 1 needs recalculation, 2 a distinct pending state.
    /// </summary>
    public ushort RecalculationState
    {
        get => GetWord(RecalculationStateOffset);
        set => SetWord(RecalculationStateOffset, value);
    }

    /// <summary>
    /// Gets or sets whether the document is protected.
    /// </summary>
    public bool IsProtected
    {
        get => GetWord(ProtectedOffset) != 0;
        set => SetWord(ProtectedOffset, (ushort)(value ? 1 : 0));
    }

    /// <summary>
    /// Gets or sets the protection password hash (see <see cref="ExcelPassword"/>); 0 means no password.
    /// </summary>
    public ushort PasswordHash
    {
        get => GetWord(PasswordHashOffset);
        set => SetWord(PasswordHashOffset, value);
    }

    /// <summary>
    /// Gets or sets the standard font number (3 = Geneva).
    /// </summary>
    public ushort StandardFont
    {
        get => GetWord(StandardFontOffset);
        set => SetWord(StandardFontOffset, value);
    }

    /// <summary>
    /// Gets or sets the standard font size in points.
    /// </summary>
    public ushort StandardFontSize
    {
        get => GetWord(StandardFontSizeOffset);
        set => SetWord(StandardFontSizeOffset, value);
    }

    /// <summary>
    /// Gets the width of a column (0–255). A width of 0 hides the column.
    /// </summary>
    public ushort GetColumnWidth(int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, ColumnCount);
        return GetWord(ColumnWidthsOffset + 2 * column);
    }

    /// <summary>
    /// Sets the width of a column (0–255).
    /// </summary>
    public void SetColumnWidth(int column, ushort width)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, ColumnCount);
        SetWord(ColumnWidthsOffset + 2 * column, width);
    }

    /// <summary>
    /// Gets the number of rows in the cell table.
    /// </summary>
    internal int RowCount => GetWord(LastRowOffset) + GetWord(NonEmptyOffset);

    /// <summary>
    /// Gets the number of columns in the cell table.
    /// </summary>
    internal int ColumnCountInTable => GetWord(LastColumnOffset) + GetWord(NonEmptyOffset);

    /// <summary>
    /// Gets whether the extent words hold a table of exactly this size in the count form (word 0x020 = 0), in which
    /// Excel uses every row and column of the table, so it needs no blank margin.
    /// </summary>
    internal bool HoldsCountForm(int rows, int columns)
        => rows > 0 && GetWord(NonEmptyOffset) == 0 && RowCount == rows && ColumnCountInTable == columns;

    /// <summary>
    /// Sets the words Excel recomputes from the names and external documents when it saves (spec §6): 0x024, the
    /// number of names whose definition is a reference, and 0x03E, the 1-based position of the active-sheet
    /// external entry, or 0. (0x022 is only counted up by Define Name; Open and Save carry it through as it is.)
    /// </summary>
    internal void SetNameCounts(int references, int activeSheetEntry)
    {
        SetWord(0x024, (ushort)references);
        SetWord(0x03E, (ushort)activeSheetEntry);
    }

    /// <summary>
    /// Updates the extent words so they describe a cell table of the given size.
    /// </summary>
    internal void SetExtent(int rows, int columns)
    {
        // The grid size is stored as last row/column plus the word at 0x020. Excel normally stores the last
        // index and 1, but some files (LOAN1 from Excel 1.5) store the count and 0; keep whatever pair gives
        // the right size.
        if (RowCount == rows && ColumnCountInTable == columns)
        {
            return;
        }

        if (rows == 0 || columns == 0)
        {
            SetWord(LastRowOffset, 0);
            SetWord(LastColumnOffset, 0);
            SetWord(NonEmptyOffset, 0);
        }
        else
        {
            SetWord(LastRowOffset, (ushort)(rows - 1));
            SetWord(LastColumnOffset, (ushort)(columns - 1));
            SetWord(NonEmptyOffset, 1);
        }
    }

    /// <summary>
    /// Creates the state block of a new, empty sheet, using the values Excel 1.03 writes for one.
    /// </summary>
    public static ExcelSheetState CreateDefault()
    {
        var data = new byte[Size];
        var state = new ExcelSheetState(data);
        state.MaxIterations = 100;
        state.CalculationMode = ExcelCalculationMode.Automatic;
        state.FullPrecision = true;
        state.ReferenceStyle = ExcelReferenceStyle.A1;
        state.MaxChange = 0.001;
        state.LastVisibleColumn = 255;

        // Null far-pointer markers Excel writes for an empty sheet.
        foreach (var offset in new[] { 0x26, 0x32, 0x36, 0x3A })
        {
            state.SetWord(offset, 0xFFFF);
        }

        state.SetWord(0x40, 1);
        for (int i = 0; i < ColumnCount; i++)
        {
            state.SetColumnWidth(i, 75);
        }

        state.StandardFont = 3;
        state.StandardFontSize = 10;

        // Cached Geneva 10 metrics (recomputed by Excel on open).
        state.SetWord(0x248, 5);
        state.SetWord(0x24A, 12);
        state.SetWord(0x24C, 13);
        state.SetWord(0x24E, 10);
        state.SetWord(0x250, 0x0CF8);
        state.SetWord(0x252, 8);
        state.SetWord(0x254, 7);
        return state;
    }
}

/// <summary>The calculation mode: state word 0x002 (spec §6).</summary>
public enum ExcelCalculationMode : short
{
    /// <summary>Automatic except tables (CALCULATION(2)).</summary>
    AutomaticExceptTables = -1,

    /// <summary>Manual (CALCULATION(3)).</summary>
    Manual = 0,

    /// <summary>Automatic.</summary>
    Automatic = 1,
}

/// <summary>The reference style: state word 0x006 (spec §6).</summary>
public enum ExcelReferenceStyle : ushort
{
    /// <summary>R1C1 references.</summary>
    R1C1 = 0,

    /// <summary>A1 references.</summary>
    A1 = 1,
}
