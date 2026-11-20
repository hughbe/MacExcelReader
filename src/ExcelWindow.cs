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
/// A saved window (spec §12).
/// </summary>
public sealed class ExcelWindow
{
    /// <summary>The size of the window state block.</summary>
    public const int StateSize = 41;

    /// <summary>
    /// Creates a window with the given state block.
    /// </summary>
    public ExcelWindow(short top, short left, short bottom, short right, byte[] state)
    {
        ArgumentNullException.ThrowIfNull(state);
        if (state.Length != StateSize)
        {
            throw new ArgumentException($"The window state must be {StateSize} bytes.", nameof(state));
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

    /// <summary>Gets the raw 41-byte window state.</summary>
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

    internal static ExcelWindow Read(BigEndianReader reader)
    {
        var window = new ExcelWindow(reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), reader.ReadBytes(StateSize));
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
