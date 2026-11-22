using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A cell range used in a window's selection: first and last row, first and last column.
/// </summary>
/// <param name="FirstRow">The first row.</param>
/// <param name="LastRow">The last row.</param>
/// <param name="FirstColumn">The first column.</param>
/// <param name="LastColumn">The last column.</param>
public readonly record struct ExcelRange(ushort FirstRow, ushort LastRow, byte FirstColumn, byte LastColumn);

/// <summary>
/// A saved window (spec §13).
/// </summary>
public sealed class ExcelWindow
{
    /// <summary>The size of the window state block in version-17 files.</summary>
    public const int StateSize = 41;

    /// <summary>
    /// The size of the window state block in version-18 files (Excel 1.5): the 41 bytes, then 88 more bytes
    /// of the window record (spec §13).
    /// </summary>
    public const int Version18StateSize = 129;

    /// <summary>
    /// Creates a window with the given state block.
    /// </summary>
    public ExcelWindow(short top, short left, short bottom, short right, byte[] state)
    {
        ArgumentNullException.ThrowIfNull(state);
        if (state.Length is not (StateSize or Version18StateSize))
        {
            throw new ArgumentException($"The window state must be {StateSize} (or, for version 18, {Version18StateSize}) bytes.", nameof(state));
        }

        Top = top;
        Left = left;
        Bottom = bottom;
        Right = right;
        State = state;
    }

    /// <summary>Gets or sets the top of the window rectangle (global coordinates).</summary>
    public short Top { get; set; }

    /// <summary>Gets or sets the left of the window rectangle.</summary>
    public short Left { get; set; }

    /// <summary>Gets or sets the bottom of the window rectangle.</summary>
    public short Bottom { get; set; }

    /// <summary>Gets or sets the right of the window rectangle.</summary>
    public short Right { get; set; }

    /// <summary>Gets the raw window state: 41 bytes, or 129 in a version-18 file.</summary>
    public byte[] State { get; }

    /// <summary>Gets the selected areas.</summary>
    public List<ExcelRange> Selection { get; } = [];

    /// <summary>Gets or sets the active cell's column.</summary>
    public byte ActiveColumn { get => State[0]; set => State[0] = value; }

    /// <summary>Gets or sets the active cell's row.</summary>
    public ushort ActiveRow
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(35));
        set => BinaryPrimitives.WriteUInt16BigEndian(State.AsSpan(35), value);
    }

    /// <summary>Gets or sets the index of the active area within <see cref="Selection"/>.</summary>
    public ushort ActiveArea
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(33));
        set => BinaryPrimitives.WriteUInt16BigEndian(State.AsSpan(33), value);
    }

    /// <summary>Gets or sets whether formulas are displayed instead of values.</summary>
    public bool DisplayFormulas { get => State[26] != 0; set => State[26] = (byte)(value ? 1 : 0); }

    /// <summary>Gets or sets whether gridlines are displayed.</summary>
    public bool DisplayGridlines { get => State[27] != 0; set => State[27] = (byte)(value ? 1 : 0); }

    /// <summary>Gets or sets whether row and column headings are displayed.</summary>
    public bool DisplayHeadings { get => State[28] != 0; set => State[28] = (byte)(value ? 1 : 0); }

    /// <summary>Gets or sets the vertical pixel position of the horizontal split (0 = none).</summary>
    public ushort HorizontalSplit
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(37));
        set => BinaryPrimitives.WriteUInt16BigEndian(State.AsSpan(37), value);
    }

    /// <summary>Gets or sets the horizontal pixel position of the vertical split (0 = none).</summary>
    public ushort VerticalSplit
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(39));
        set => BinaryPrimitives.WriteUInt16BigEndian(State.AsSpan(39), value);
    }

    /// <summary>Gets whether the state block is the 129-byte one of a version-18 file (Excel 1.5).</summary>
    public bool HasVersion18State => State.Length == Version18StateSize;

    /// <summary>
    /// Gets or sets whether the panes are frozen (Excel 1.5, version-18 state offset 121; spec §13). Always false
    /// for a 41-byte state.
    /// </summary>
    public bool PanesFrozen
    {
        get => HasVersion18State && State[121] != 0;
        set => Version18State()[121] = (byte)(value ? 1 : 0);
    }

    /// <summary>
    /// Gets or sets the first row below the split where Excel 1.5 froze the panes (version-18 state offset 123;
    /// spec §13), or null for a 41-byte state. Left over from memory until the window's panes are first frozen.
    /// </summary>
    public ushort? FrozenRow
    {
        get => HasVersion18State ? BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(123)) : null;
        set => BinaryPrimitives.WriteUInt16BigEndian(Version18State().AsSpan(123), value ?? 0);
    }

    /// <summary>
    /// Gets or sets the first column right of the split where Excel 1.5 froze the panes (version-18 state offset
    /// 125; spec §13), or null for a 41-byte state. Left over from memory until the window's panes are first frozen.
    /// </summary>
    public ushort? FrozenColumn
    {
        get => HasVersion18State ? BinaryPrimitives.ReadUInt16BigEndian(State.AsSpan(125)) : null;
        set => BinaryPrimitives.WriteUInt16BigEndian(Version18State().AsSpan(125), value ?? 0);
    }

    /// <summary>
    /// Gets or sets the gridline and heading colour (Excel 1.5, version-18 state offsets 127 and 128; spec §13), or
    /// null for automatic. A 41-byte state is always automatic.
    /// </summary>
    public ExcelColor? GridlineColor
    {
        get => HasVersion18State && State[127] == 0 ? (ExcelColor)State[128] : null;
        set
        {
            var state = Version18State();
            state[127] = (byte)(value is null ? 1 : 0);
            state[128] = (byte)(value ?? 0);
        }
    }

    private byte[] Version18State()
        => HasVersion18State ? State : throw new InvalidOperationException("Only a version-18 window state holds this setting.");

    /// <summary>
    /// Creates the window Excel 1.03 saves for a new sheet, with A1 selected.
    /// </summary>
    public static ExcelWindow CreateDefault(bool macroSheet)
    {
        // The state Excel 1.03 saved for the empty sample sheet, with the active cell moved to A1.
        byte[] state =
        [
            0x00, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x12, 0x00, 0x12, 0x00, 0x00, 0x05, 0x06, 0x00,
            0x05, 0x00, 0x06, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x2F, 0x0C, 0x00, 0x01, 0x01, 0xFF, 0x4B, 0x0E,
            0x55, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        var window = new ExcelWindow(0x3D, 0x09, 0x14F, 0x1F6, state) { DisplayFormulas = macroSheet };
        window.Selection.Add(new ExcelRange(0, 0, 0, 0));
        return window;
    }

    internal static ExcelWindow Read(BigEndianReader reader, int stateSize)
    {
        var window = new ExcelWindow(reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), reader.ReadBytes(stateSize));
        int count = reader.ReadUInt16();
        for (int i = 0; i < count; i++)
        {
            window.Selection.Add(new ExcelRange(reader.ReadUInt16(), reader.ReadUInt16(), reader.ReadByte(), reader.ReadByte()));
        }

        return window;
    }

    internal void Write(BigEndianWriter writer)
    {
        writer.WriteByte(0xFF);
        WriteBody(writer);
    }

    // The rectangle, state and selection: everything after the record's marker byte.
    internal void WriteBody(BigEndianWriter writer)
    {
        writer.WriteInt16(Top);
        writer.WriteInt16(Left);
        writer.WriteInt16(Bottom);
        writer.WriteInt16(Right);
        writer.WriteBytes(State);
        writer.WriteUInt16((ushort)Selection.Count);
        foreach (var range in Selection)
        {
            writer.WriteUInt16(range.FirstRow);
            writer.WriteUInt16(range.LastRow);
            writer.WriteByte(range.FirstColumn);
            writer.WriteByte(range.LastColumn);
        }
    }
}

/// <summary>The eight colours of Excel 1.5, in the order of <c>STR#</c> 263 (spec §7, §13).</summary>
public enum ExcelColor
{
    /// <summary>Black.</summary>
    Black = 0,

    /// <summary>White.</summary>
    White = 1,

    /// <summary>Red.</summary>
    Red = 2,

    /// <summary>Green.</summary>
    Green = 3,

    /// <summary>Blue.</summary>
    Blue = 4,

    /// <summary>Yellow.</summary>
    Yellow = 5,

    /// <summary>Magenta.</summary>
    Magenta = 6,

    /// <summary>Cyan.</summary>
    Cyan = 7,
}
