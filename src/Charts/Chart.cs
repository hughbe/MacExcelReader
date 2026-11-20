using System.Buffers.Binary;

namespace MacExcelReader.Charts;

/// <summary>
/// The chart type (chart header word 2).
/// </summary>
public enum ChartType : ushort
{
    /// <summary>Column chart, or bar chart when <see cref="Chart.IsBar"/> is set.</summary>
    ColumnOrBar = 1,

    /// <summary>Line chart.</summary>
    Line = 2,

    /// <summary>Pie chart.</summary>
    Pie = 3,

    /// <summary>Area chart.</summary>
    Area = 4,

    /// <summary>Scatter chart.</summary>
    Scatter = 5,
}

/// <summary>
/// The legend position (bits 15–13 of chart-structure word 156).
/// </summary>
public enum LegendPosition
{
    /// <summary>Bottom.</summary>
    Bottom = 0,

    /// <summary>Corner.</summary>
    Corner = 1,

    /// <summary>Top.</summary>
    Top = 2,

    /// <summary>Vertical (right).</summary>
    Vertical = 3,
}

/// <summary>
/// Tick mark style.
/// </summary>
public enum TickMarkType
{
    /// <summary>No tick marks.</summary>
    Invisible = 0,

    /// <summary>Inside the axis.</summary>
    Inside = 1,

    /// <summary>Outside the axis.</summary>
    Outside = 2,

    /// <summary>Crossing the axis.</summary>
    Cross = 3,
}

/// <summary>
/// Tick label position.
/// </summary>
public enum TickLabelPosition
{
    /// <summary>No labels.</summary>
    None = 0,

    /// <summary>Next to the axis.</summary>
    NextToAxis = 1,

    /// <summary>Low.</summary>
    Low = 2,

    /// <summary>High.</summary>
    High = 3,
}

/// <summary>
/// One chart of a chart document (spec §16.5): the main chart or the overlay chart.
/// </summary>
public sealed class Chart
{
    /// <summary>The size of the chart record in bytes.</summary>
    public const int RecordSize = 378;

    /// <summary>The size of a scale record in bytes.</summary>
    public const int ScaleSize = 44;

    /// <summary>On-disk offset of the category axis record.</summary>
    public const int CategoryAxisOffset = 40;

    /// <summary>On-disk offset of the value axis record.</summary>
    public const int ValueAxisOffset = 64;

    /// <summary>On-disk offset of the value-axis scale record.</summary>
    public const int ValueScaleOffset = 88;

    /// <summary>
    /// Creates a chart from a raw record.
    /// </summary>
    public Chart(byte[] record)
    {
        ArgumentNullException.ThrowIfNull(record);
        if (record.Length != RecordSize)
        {
            throw new ArgumentException($"A chart record is {RecordSize} bytes.", nameof(record));
        }

        Record = record;
    }

    /// <summary>Gets the raw 378-byte chart record.</summary>
    public byte[] Record { get; }

    /// <summary>Gets or sets the X-axis scale record of a scatter chart (44 bytes), or null.</summary>
    public byte[]? XAxisScale { get; set; }

    /// <summary>Gets the text objects: titles first, then attached and free text.</summary>
    public List<ChartText> Texts { get; } = [];

    /// <summary>Gets the per-series word list (ignored by Excel on load).</summary>
    public List<ushort> SeriesWords { get; } = [];

    /// <summary>Gets the 12-byte data-point format records.</summary>
    public List<byte[]> PointFormats { get; } = [];

    private ushort GetWord(int offset) => BinaryPrimitives.ReadUInt16BigEndian(Record.AsSpan(offset));

    private void SetWord(int offset, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(Record.AsSpan(offset), value);

    private bool GetBit(int offset, int bit) => (GetWord(offset) & (1 << bit)) != 0;

    private void SetBit(int offset, int bit, bool value)
    {
        ushort word = GetWord(offset);
        SetWord(offset, (ushort)(value ? word | (1 << bit) : word & ~(1 << bit)));
    }

    private int GetField(int offset, int shift, int bits) => (GetWord(offset) >> shift) & ((1 << bits) - 1);

    private void SetField(int offset, int shift, int bits, int value)
    {
        int mask = ((1 << bits) - 1) << shift;
        SetWord(offset, (ushort)((GetWord(offset) & ~mask) | ((value << shift) & mask)));
    }

    /// <summary>Gets or sets the chart type.</summary>
    public ChartType Type { get => (ChartType)GetWord(2); set => SetWord(2, (ushort)value); }

    /// <summary>Gets or sets whether a column chart is drawn as horizontal bars.</summary>
    public bool IsBar { get => GetBit(0, 12); set => SetBit(0, 12, value); }

    /// <summary>Gets or sets whether series are stacked.</summary>
    public bool IsStacked { get => GetBit(0, 13); set => SetBit(0, 13, value); }

    /// <summary>Gets or sets whether a stacked chart is drawn as 100%.</summary>
    public bool IsPercent { get => GetBit(0, 14); set => SetBit(0, 14, value); }

    /// <summary>Gets or sets whether colours vary by category.</summary>
    public bool VaryByCategory { get => GetBit(0, 11); set => SetBit(0, 11, value); }

    /// <summary>Gets or sets automatic series distribution (overlay chart).</summary>
    public bool AutomaticSeriesDistribution { get => GetBit(0, 10); set => SetBit(0, 10, value); }

    /// <summary>Gets or sets the first series in the overlay chart (overlay chart).</summary>
    public int FirstOverlaySeries { get => GetField(0, 2, 8); set => SetField(0, 2, 8, value); }

    /// <summary>Gets or sets the column/bar overlap percentage; negative when "Overlapped" is on.</summary>
    public short Overlap { get => unchecked((short)GetWord(8)); set => SetWord(8, unchecked((ushort)value)); }

    /// <summary>Gets or sets the column/bar cluster spacing percentage.</summary>
    public short ClusterSpacing { get => unchecked((short)GetWord(10)); set => SetWord(10, unchecked((ushort)value)); }

    /// <summary>Gets or sets the angle of the first pie slice in degrees.</summary>
    public short FirstSliceAngle { get => unchecked((short)GetWord(12)); set => SetWord(12, unchecked((ushort)value)); }

    /// <summary>Gets a view of the category axis record.</summary>
    public ChartAxis CategoryAxis => new(Record, CategoryAxisOffset);

    /// <summary>Gets a view of the value axis record.</summary>
    public ChartAxis ValueAxis => new(Record, ValueAxisOffset);

    /// <summary>Gets a view of the value-axis scale.</summary>
    public ChartScale ValueScale => new(Record, ValueScaleOffset);

    /// <summary>Gets a view of the X-axis scale of a scatter chart, or null.</summary>
    public ChartScale? XScale => XAxisScale is null ? null : new ChartScale(XAxisScale, 0);

    /// <summary>Gets or sets the category number the value axis crosses at.</summary>
    public byte ValueAxisCrossesAtCategory { get => Record[132]; set => Record[132] = value; }

    /// <summary>Gets or sets whether the value axis crosses between categories.</summary>
    public bool ValueAxisCrossesBetweenCategories
    {
        get => (Record[133] & 0x80) != 0;
        set => Record[133] = (byte)(value ? Record[133] | 0x80 : Record[133] & ~0x80);
    }

    /// <summary>Gets or sets whether the value axis crosses at the maximum category.</summary>
    public bool ValueAxisCrossesAtMaximumCategory
    {
        get => (Record[133] & 0x40) != 0;
        set => Record[133] = (byte)(value ? Record[133] | 0x40 : Record[133] & ~0x40);
    }

    /// <summary>Gets or sets whether categories are in reverse order.</summary>
    public bool CategoriesInReverseOrder
    {
        get => (Record[133] & 0x20) != 0;
        set => Record[133] = (byte)(value ? Record[133] | 0x20 : Record[133] & ~0x20);
    }

    /// <summary>Gets or sets the number of categories between tick labels.</summary>
    public byte CategoriesBetweenTickLabels { get => Record[134]; set => Record[134] = value; }

    /// <summary>Gets or sets the legend position.</summary>
    public LegendPosition LegendPosition { get => (LegendPosition)GetField(154, 13, 3); set => SetField(154, 13, 3, (int)value); }

    /// <summary>Gets or sets the number of text objects as stored (updated on write).</summary>
    internal ushort StoredTextCount { get => GetWord(264); set => SetWord(264, value); }

    internal ushort StoredSeriesWordCount { get => GetWord(360); set => SetWord(360, value); }

    internal ushort StoredPointFormatCount { get => GetWord(364); set => SetWord(364, value); }

    internal bool HasXScaleMarker
    {
        get => GetWord(136) != 0;
        set => SetWord(136, (ushort)(value ? Math.Max(GetWord(136), (ushort)1) : 0));
    }

    /// <summary>
    /// Creates a chart record with the given type and Excel's usual option values.
    /// </summary>
    public static Chart Create(ChartType type)
    {
        var chart = new Chart(new byte[RecordSize]) { Type = type, ClusterSpacing = 50, CategoriesBetweenTickLabels = 1, ValueAxisCrossesAtCategory = 1 };
        chart.ValueAxisCrossesBetweenCategories = true;
        chart.CategoryAxis.TickLabelPosition = TickLabelPosition.NextToAxis;
        chart.ValueAxis.TickLabelPosition = TickLabelPosition.NextToAxis;
        chart.CategoryAxis.MajorTickMarks = TickMarkType.Outside;
        chart.ValueAxis.MajorTickMarks = TickMarkType.Outside;
        chart.CategoryAxis.AxisLinePattern = 1;
        chart.ValueAxis.AxisLinePattern = 1;
        var scale = chart.ValueScale;
        scale.AutoMinimum = scale.AutoMaximum = scale.AutoMajorUnit = scale.AutoMinorUnit = scale.AutoCrossesAt = true;
        if (type == ChartType.Scatter)
        {
            chart.XAxisScale = new byte[ScaleSize];
            var x = chart.XScale!;
            x.AutoMinimum = x.AutoMaximum = x.AutoMajorUnit = x.AutoMinorUnit = x.AutoCrossesAt = true;
        }

        return chart;
    }
}

/// <summary>
/// A view of an axis record inside a chart record (spec §16.5).
/// </summary>
public sealed class ChartAxis
{
    private readonly byte[] _data;
    private readonly int _offset;

    internal ChartAxis(byte[] data, int offset)
    {
        _data = data;
        _offset = offset;
    }

    private ushort Word => BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(_offset));

    private void SetField(int shift, int value)
    {
        int mask = 0xF << shift;
        BinaryPrimitives.WriteUInt16BigEndian(_data.AsSpan(_offset), (ushort)((Word & ~mask) | ((value << shift) & mask)));
    }

    /// <summary>Gets or sets the major tick marks.</summary>
    public TickMarkType MajorTickMarks { get => (TickMarkType)((Word >> 12) & 0xF); set => SetField(12, (int)value); }

    /// <summary>Gets or sets the minor tick marks.</summary>
    public TickMarkType MinorTickMarks { get => (TickMarkType)((Word >> 8) & 0xF); set => SetField(8, (int)value); }

    /// <summary>Gets or sets the tick label position.</summary>
    public TickLabelPosition TickLabelPosition { get => (TickLabelPosition)((Word >> 4) & 0xF); set => SetField(4, (int)value); }

    /// <summary>Gets or sets the axis line pattern; 0 hides the axis.</summary>
    public byte AxisLinePattern { get => _data[_offset + 13]; set => _data[_offset + 13] = value; }

    /// <summary>Gets or sets the major gridline pattern; 0 hides them.</summary>
    public byte MajorGridlinePattern { get => _data[_offset + 17]; set => _data[_offset + 17] = value; }

    /// <summary>Gets or sets the minor gridline pattern; 0 hides them.</summary>
    public byte MinorGridlinePattern { get => _data[_offset + 21]; set => _data[_offset + 21] = value; }

    /// <summary>Gets the tick-label font descriptor bytes (on disk, Microsoft Chart codes).</summary>
    public ReadOnlySpan<byte> Font => _data.AsSpan(_offset + 2, 4);
}

/// <summary>
/// A view of a scale record (spec §16.5).
/// </summary>
public sealed class ChartScale
{
    private readonly byte[] _data;
    private readonly int _offset;

    internal ChartScale(byte[] data, int offset)
    {
        _data = data;
        _offset = offset;
    }

    private double Get(int index) => MicrosoftDecimal.FromBytes(_data.AsSpan(_offset + 8 * index)).ToDouble();

    private void Set(int index, double value) => MicrosoftDecimal.FromDouble(value).WriteTo(_data.AsSpan(_offset + 8 * index));

    private bool GetFlag(int bit) => (BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(_offset + 40)) & (1 << bit)) != 0;

    private void SetFlag(int bit, bool value)
    {
        var span = _data.AsSpan(_offset + 40);
        ushort word = BinaryPrimitives.ReadUInt16BigEndian(span);
        BinaryPrimitives.WriteUInt16BigEndian(span, (ushort)(value ? word | (1 << bit) : word & ~(1 << bit)));
    }

    /// <summary>Gets or sets the minimum.</summary>
    public double Minimum { get => Get(0); set => Set(0, value); }

    /// <summary>Gets or sets the maximum.</summary>
    public double Maximum { get => Get(1); set => Set(1, value); }

    /// <summary>Gets or sets the major unit.</summary>
    public double MajorUnit { get => Get(2); set => Set(2, value); }

    /// <summary>Gets or sets the minor unit.</summary>
    public double MinorUnit { get => Get(3); set => Set(3, value); }

    /// <summary>Gets or sets where the other axis crosses.</summary>
    public double CrossesAt { get => Get(4); set => Set(4, value); }

    /// <summary>Gets or sets automatic minimum.</summary>
    public bool AutoMinimum { get => GetFlag(15); set => SetFlag(15, value); }

    /// <summary>Gets or sets automatic maximum.</summary>
    public bool AutoMaximum { get => GetFlag(14); set => SetFlag(14, value); }

    /// <summary>Gets or sets automatic major unit.</summary>
    public bool AutoMajorUnit { get => GetFlag(13); set => SetFlag(13, value); }

    /// <summary>Gets or sets automatic minor unit.</summary>
    public bool AutoMinorUnit { get => GetFlag(12); set => SetFlag(12, value); }

    /// <summary>Gets or sets automatic crossing point.</summary>
    public bool AutoCrossesAt { get => GetFlag(11); set => SetFlag(11, value); }

    /// <summary>Gets or sets a logarithmic scale.</summary>
    public bool Logarithmic { get => GetFlag(10); set => SetFlag(10, value); }

    /// <summary>Gets or sets values in reverse order.</summary>
    public bool ReverseOrder { get => GetFlag(9); set => SetFlag(9, value); }
}
