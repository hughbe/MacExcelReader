using System.Buffers.Binary;

namespace MacExcelReader.Charts;

/// <summary>
/// What a chart text object is attached to.
/// </summary>
public enum ChartTextKind
{
    /// <summary>Not attached.</summary>
    Unattached = 0,

    /// <summary>The chart title.</summary>
    ChartTitle = 1,

    /// <summary>The category-axis title.</summary>
    CategoryAxisTitle = 2,

    /// <summary>The value-axis title.</summary>
    ValueAxisTitle = 3,

    /// <summary>A data label attached to a series or point.</summary>
    DataLabel = 4,

    /// <summary>Free text placed anywhere on the chart, or an arrow (see <see cref="ChartText.IsArrow"/>).</summary>
    Free = 6,
}

/// <summary>
/// A chart text object (spec §16.7): a 46-byte record and, unless the text is automatic, a string.
/// Disk offsets are the memory offsets minus 2 for bytes 8–37, minus 2 for 38–43 and minus 4 for 46–49.
/// </summary>
public sealed class ChartText
{
    /// <summary>The size of the record in bytes.</summary>
    public const int RecordSize = 46;

    /// <summary>
    /// Creates a text object from a raw record.
    /// </summary>
    public ChartText(byte[] record, string? text = null)
    {
        ArgumentNullException.ThrowIfNull(record);
        if (record.Length != RecordSize)
        {
            throw new ArgumentException($"A text record is {RecordSize} bytes.", nameof(record));
        }

        Record = record;
        Text = text;
    }

    /// <summary>
    /// Creates a text object of the given kind.
    /// </summary>
    public static ChartText Create(ChartTextKind kind, string text)
    {
        var item = new ChartText(new byte[RecordSize], text) { Kind = kind };
        return item;
    }

    /// <summary>Gets the raw record.</summary>
    public byte[] Record { get; }

    /// <summary>Gets or sets the text; null when <see cref="IsAutomaticText"/> is set.</summary>
    public string? Text { get; set; }

    private ushort Flags
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(Record.AsSpan(8));
        set => BinaryPrimitives.WriteUInt16BigEndian(Record.AsSpan(8), value);
    }

    private void SetFlag(int bit, bool value) => Flags = (ushort)(value ? Flags | (1 << bit) : Flags & ~(1 << bit));

    /// <summary>Gets or sets what the text is attached to.</summary>
    public ChartTextKind Kind
    {
        get => (ChartTextKind)(Record[3] & 0xF);
        set => Record[3] = (byte)((Record[3] & 0xF0) | ((int)value & 0xF));
    }

    /// <summary>Gets or sets the series number of a data label.</summary>
    public byte SeriesNumber { get => Record[4]; set => Record[4] = value; }

    /// <summary>Gets or sets the point number of a data label (255 = the whole series).</summary>
    public byte PointNumber { get => Record[5]; set => Record[5] = value; }

    /// <summary>
    /// Gets or sets whether the object is an arrow rather than text (memory offset 17). An arrow is a free object
    /// with automatic text and no string; its two points are the arrow's ends (spec §16.7).
    /// </summary>
    public bool IsArrow { get => ArrowLine.Pattern != ChartPattern.None; set => ArrowLine.Pattern = value ? ChartPattern.Solid : ChartPattern.None; }

    /// <summary>
    /// Gets an arrow's line: its pattern (memory offset 17), which is <see cref="ChartPattern.None"/> for text, and
    /// its shaft weight.
    /// </summary>
    public ChartLineFormat ArrowLine => new(Record, 14);

    /// <summary>Gets the text box's area format: border, background and the box Excel last drew it in.</summary>
    public ChartAreaFormat Area => new(Record, 20);

    /// <summary>Gets or sets an arrow's head width.</summary>
    public ArrowHeadWidth HeadWidth { get => (ArrowHeadWidth)((Word(18) >> 12) & 3); set => SetWord(18, (ushort)((Word(18) & ~0x3000) | (((int)value & 3) << 12))); }

    /// <summary>Gets or sets an arrow's head length; <see cref="ArrowHeadLength.None"/> means no head.</summary>
    public ArrowHeadLength HeadLength { get => (ArrowHeadLength)((Word(18) >> 8) & 3); set => SetWord(18, (ushort)((Word(18) & ~0x0300) | (((int)value & 3) << 8))); }

    /// <summary>
    /// Gets or sets an arrow's head style. Setting <see cref="ArrowHeadStyle.None"/> sets the length to
    /// <see cref="ArrowHeadLength.None"/>; setting a head on an arrow without one gives it a medium length.
    /// </summary>
    public ArrowHeadStyle HeadStyle
    {
        get => HeadLength == ArrowHeadLength.None ? ArrowHeadStyle.None : (Word(18) & 0x0080) != 0 ? ArrowHeadStyle.Open : ArrowHeadStyle.Filled;
        set
        {
            if (value == ArrowHeadStyle.None)
            {
                HeadLength = ArrowHeadLength.None;
                return;
            }

            if (HeadLength == ArrowHeadLength.None)
            {
                HeadLength = ArrowHeadLength.Medium;
            }

            SetWord(18, (ushort)(value == ArrowHeadStyle.Open ? Word(18) | 0x0080 : Word(18) & ~0x0080));
        }
    }

    private ushort Word(int offset) => System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(Record.AsSpan(offset));

    private void SetWord(int offset, ushort value) => System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(Record.AsSpan(offset), value);

    /// <summary>Gets or sets whether the text is generated automatically (no string is stored).</summary>
    public bool IsAutomaticText { get => (Flags & 0x2000) != 0; set => SetFlag(13, value); }

    /// <summary>Gets or sets vertical orientation.</summary>
    public bool IsVertical { get => (Flags & 0x1000) != 0; set => SetFlag(12, value); }

    /// <summary>Gets or sets the horizontal alignment: 0 centre, 1 left, 2 right.</summary>
    public int HorizontalAlignment
    {
        get => (Flags >> 9) & 7;
        set => Flags = (ushort)((Flags & ~0x0E00) | ((value & 7) << 9));
    }

    /// <summary>Gets or sets the vertical alignment: 0 bottom, 1 centre, 2 top.</summary>
    public int VerticalAlignment
    {
        get => (Flags >> 6) & 7;
        set => Flags = (ushort)((Flags & ~0x01C0) | ((value & 7) << 6));
    }

    /// <summary>Gets or sets "show key".</summary>
    public bool ShowKey { get => (Flags & 0x20) != 0; set => SetFlag(5, value); }

    /// <summary>Gets or sets "show value".</summary>
    public bool ShowValue { get => (Flags & 0x10) != 0; set => SetFlag(4, value); }

    /// <inheritdoc/>
    public override string ToString() => $"{Kind}: {Text}";
}
