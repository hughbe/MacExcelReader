using MacExcelReader.Internal;

namespace MacExcelReader.Charts;

/// <summary>
/// An Excel 1.x chart document, saved in Microsoft Chart format (Finder type MCBN, spec §16).
/// </summary>
public sealed class ChartDocument
{
    /// <summary>The magic number of chart files.</summary>
    public const ushort Magic = 0x018C;

    /// <summary>The format version Excel writes.</summary>
    public const ushort CurrentVersion = 12;

    /// <summary>The code signature Excel writes in chart files.</summary>
    public const ushort ChartSignature = 0x13E8;

    private const int HeaderSize = 128;
    private const int PageSetupBodySize = 664;

    /// <summary>
    /// Gets or sets the format version: 12 from Excel, 1 from Excel 1.5 for a chart with more than 102 categories
    /// (same layout; earlier releases refuse it), and 2 or 3 for older Microsoft Chart files (spec §16.2).
    /// </summary>
    public ushort Version { get; set; } = CurrentVersion;

    /// <summary>Gets or sets the code signature.</summary>
    public ushort Signature { get; set; } = ChartSignature;

    /// <summary>Gets or sets the embedded p-code segment (empty for documents created here).</summary>
    public byte[] CodeSegment { get; set; } = [];

    /// <summary>Gets or sets the 8-byte chart-set header (skipped by Excel on load).</summary>
    public byte[] ChartSetHeader { get; set; } = [0, 0, 0, 0, 0, 1, 0, 0x11];

    /// <summary>
    /// Gets the charts in file order: when there is an overlay chart it comes first, then the main chart.
    /// </summary>
    public List<Chart> Charts { get; } = [];

    /// <summary>Gets the main chart (the last in file order).</summary>
    public Chart MainChart => Charts[^1];

    /// <summary>Gets the overlay chart, or null.</summary>
    public Chart? OverlayChart => Charts.Count > 1 ? Charts[0] : null;

    /// <summary>
    /// Gets or sets the 12-byte window block: the window's position (stored negated) and size, then the chart's
    /// drawing size (see <see cref="ChartWidth"/>, <see cref="ChartHeight"/>).
    /// </summary>
    public byte[] Window { get; set; } = new byte[12];

    /// <summary>
    /// Gets or sets the chart's drawing width in twips (1/1440 inch): the window's content width less 16 pixels of
    /// scroll bar, at 72 pixels to the inch (spec §16.1). Excel scales the chart's fonts from it.
    /// </summary>
    public short ChartWidth
    {
        get => System.Buffers.Binary.BinaryPrimitives.ReadInt16BigEndian(Window.AsSpan(8));
        set => System.Buffers.Binary.BinaryPrimitives.WriteInt16BigEndian(Window.AsSpan(8), value);
    }

    /// <summary>Gets or sets the chart's drawing height in twips, likewise (spec §16.1).</summary>
    public short ChartHeight
    {
        get => System.Buffers.Binary.BinaryPrimitives.ReadInt16BigEndian(Window.AsSpan(10));
        set => System.Buffers.Binary.BinaryPrimitives.WriteInt16BigEndian(Window.AsSpan(10), value);
    }

    /// <summary>
    /// Gets or sets the 664-byte page-setup body (62 bytes of the print record, 58 zeros, four decimal
    /// margins and 256-byte header and footer fields), or null if the chart has no page setup.
    /// </summary>
    public byte[]? PageSetup { get; set; }

    /// <summary>Gets the series, in file order.</summary>
    public List<ChartSeries> Series { get; } = [];

    /// <summary>
    /// Creates a chart document with one chart of the given type.
    /// </summary>
    public static ChartDocument Create(ChartType type)
    {
        // The window Excel 1.5 saved for PROFIT CHART: at (58, 4), 276 by 500 pixels, so a chart drawing of
        // 484 by 260 pixels, in twips.
        var document = new ChartDocument { Window = Convert.FromHexString("003A00040114 01F4 25D0 1450".Replace(" ", "")) };
        document.Charts.Add(Chart.Create(type));
        return document;
    }

    /// <summary>Reads a chart document from a file.</summary>
    public static ChartDocument Load(string path) => Read(File.ReadAllBytes(path));

    /// <summary>
    /// Loads a chart file asynchronously.
    /// </summary>
    public static async Task<ChartDocument> LoadAsync(string path, CancellationToken cancellationToken = default)
        => Read(await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false));

    /// <summary>
    /// Reads a chart document from its bytes.
    /// </summary>
    public static ChartDocument Read(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        var reader = new BigEndianReader(data);
        if (data.Length < HeaderSize || reader.ReadUInt16() != Magic)
        {
            throw new ExcelFormatException("Not an Excel chart (bad magic number).");
        }

        var document = new ChartDocument { Version = reader.ReadUInt16(), Signature = reader.ReadUInt16() };
        reader.Skip(6);
        int segmentCount = reader.ReadUInt16();
        int segmentLength = reader.ReadUInt16();
        int segmentOffset = reader.ReadUInt16();
        int chartCount = reader.ReadUInt16();
        int dataOffset = reader.ReadUInt16();
        reader.ReadUInt16();
        int seriesCount = reader.ReadUInt16();
        int seriesOffset = reader.ReadUInt16();
        if (segmentCount > 0)
        {
            document.CodeSegment = new BigEndianReader(data, segmentOffset).ReadBytes(segmentLength);
        }

        reader.Position = dataOffset;
        document.ChartSetHeader = reader.ReadBytes(8);
        for (int i = 0; i < chartCount; i++)
        {
            document.Charts.Add(ReadChart(reader));
        }

        document.Window = reader.ReadBytes(12);
        bool hasPageSetup = reader.ReadUInt16() != 0;
        var pageSetup = reader.ReadBytes(PageSetupBodySize);
        document.PageSetup = hasPageSetup ? pageSetup : null;

        if (seriesCount > 0)
        {
            int position = seriesOffset;
            for (int i = 0; i < seriesCount; i++)
            {
                document.Series.Add(ReadSeries(data, ref position));
            }
        }

        return document;
    }

    private static Chart ReadChart(BigEndianReader reader)
    {
        var chart = new Chart(reader.ReadBytes(Chart.RecordSize));
        if (chart.Type == ChartType.Scatter || chart.HasXScaleMarker)
        {
            chart.XAxisScale = reader.ReadBytes(Chart.ScaleSize);
        }

        int textCount = chart.StoredTextCount;
        for (int i = 0; i < textCount; i++)
        {
            chart.Texts.Add(new ChartText(reader.ReadBytes(ChartText.RecordSize)));
        }

        foreach (var text in chart.Texts)
        {
            text.Text = text.IsAutomaticText ? null : reader.ReadPascalString();
        }

        for (int i = 0; i < chart.StoredSeriesWordCount; i++)
        {
            chart.SeriesWords.Add(reader.ReadUInt16());
        }

        for (int i = 0; i < chart.StoredPointFormatCount; i++)
        {
            chart.PointFormats.Add(reader.ReadBytes(12));
        }

        return chart;
    }

    private static ChartSeries ReadSeries(byte[] data, ref int position)
    {
        var reader = new BigEndianReader(data, position);
        int start = position;
        int totalLength = reader.ReadUInt16();

        // The offset of the second array from the start of the record (the first array may hold strings).
        int secondArrayOffset = reader.ReadUInt16();
        var series = new ChartSeries(reader.ReadBytes(ChartSeries.HeaderSize));
        int count = new BigEndianReader(series.Header, 60).ReadUInt16();
        for (int i = 0; i < count; i++)
        {
            if (series.SeriesType == 3)
            {
                series.CategoryTexts.Add(reader.ReadPascalString());
            }
            else
            {
                series.CategoryValues.Add(MicrosoftDecimal.FromBytes(reader.ReadBytes(8)));
            }
        }

        reader.Position = start + secondArrayOffset;
        for (int i = 0; i < count; i++)
        {
            series.Values.Add(MicrosoftDecimal.FromBytes(reader.ReadBytes(8)));
        }

        series.Name = reader.ReadPascalString();
        series.FirstLabel = reader.ReadPascalString();
        series.SecondLabel = reader.ReadPascalString();
        series.SourceMarker = reader.ReadPascalString();
        series.SourceDocument = reader.ReadPascalString();
        series.ValuesReference = reader.ReadPascalString();
        series.CategoriesReference = reader.ReadPascalString();
        for (int i = 0; i < series.TrailingStrings.Length; i++)
        {
            series.TrailingStrings[i] = reader.ReadPascalString();
        }

        position = totalLength != 0 ? start + totalLength : reader.Position;
        return series;
    }

    /// <summary>Writes the chart to a file (Finder type MCBN).</summary>
    public void Save(string path) => File.WriteAllBytes(path, ToArray());

    /// <summary>
    /// Writes the chart to a file asynchronously.
    /// </summary>
    public Task SaveAsync(string path, CancellationToken cancellationToken = default)
        => File.WriteAllBytesAsync(path, ToArray(), cancellationToken);

    /// <summary>
    /// Serialises the chart document.
    /// </summary>
    public byte[] ToArray()
    {
        if (Charts.Count is 0 or > 2)
        {
            throw new InvalidOperationException("A chart document has one or two charts.");
        }

        RequireLength(CodeSegment, null, "Code segment");
        RequireLength(ChartSetHeader, 8, "Chart-set header");
        RequireLength(Window, 12, "Window");
        RequireLength(PageSetup, PageSetupBodySize, "Page setup", allowNull: true);
        foreach (var chart in Charts)
        {
            RequireLength(chart.XAxisScale, Chart.ScaleSize, "Chart X-axis scale", allowNull: true);
            foreach (var format in chart.PointFormats)
            {
                RequireLength(format, 12, "Chart point format");
            }
        }

        var writer = new BigEndianWriter();
        writer.WriteUInt16(Magic);
        writer.WriteUInt16(Version);
        writer.WriteUInt16(Signature);
        writer.WriteZeros(6);
        writer.WriteUInt16(1);
        writer.WriteUInt16((ushort)CodeSegment.Length);
        writer.WriteUInt16(HeaderSize);
        writer.WriteUInt16((ushort)Charts.Count);
        int dataOffset = HeaderSize + CodeSegment.Length;
        writer.WriteUInt16(checked((ushort)dataOffset));
        writer.WriteUInt16((ushort)(Series.Count == 0 ? 0 : Series.Max(s => s.PointCount)));
        writer.WriteUInt16((ushort)Series.Count);
        int seriesOffsetPosition = writer.Position;
        writer.WriteUInt16(0);
        writer.WriteZeros(HeaderSize - writer.Position);
        writer.WriteBytes(CodeSegment);

        writer.WriteBytes(ChartSetHeader);
        foreach (var chart in Charts)
        {
            WriteChart(writer, chart);
        }

        writer.WriteBytes(Window);
        writer.WriteUInt16((ushort)(PageSetup is null ? 0 : 1));
        if (PageSetup is null)
        {
            writer.WriteZeros(PageSetupBodySize);
        }
        else
        {
            writer.WriteBytes(PageSetup);
        }

        if (Series.Count > 0)
        {
            writer.PatchUInt16(seriesOffsetPosition, checked((ushort)writer.Position));
            for (int i = 0; i < Series.Count; i++)
            {
                bool lastInList = i == Series.Count - 1 || Series[i + 1].InFirstList != Series[i].InFirstList;
                WriteSeries(writer, Series[i], lastInList);
            }
        }

        writer.WriteUInt16(0);
        return writer.ToArray();
    }

    private static void RequireLength(byte[]? data, int? length, string name, bool allowNull = false)
    {
        if (data is null)
        {
            if (allowNull)
            {
                return;
            }

            throw new InvalidOperationException($"{name} cannot be null.");
        }

        if (length is { } expected && data.Length != expected)
        {
            throw new InvalidOperationException($"{name} must be {expected} bytes, not {data.Length}.");
        }
    }

    private static void WriteChart(BigEndianWriter writer, Chart chart)
    {
        chart.StoredTextCount = (ushort)chart.Texts.Count;
        chart.StoredSeriesWordCount = (ushort)chart.SeriesWords.Count;
        chart.StoredPointFormatCount = (ushort)chart.PointFormats.Count;
        chart.HasXScaleMarker = chart.XAxisScale is not null;
        writer.WriteBytes(chart.Record);
        if (chart.XAxisScale is not null)
        {
            writer.WriteBytes(chart.XAxisScale);
        }

        foreach (var text in chart.Texts)
        {
            writer.WriteBytes(text.Record);
        }

        foreach (var text in chart.Texts)
        {
            if (!text.IsAutomaticText)
            {
                writer.WritePascalString(text.Text ?? string.Empty);
            }
        }

        foreach (var word in chart.SeriesWords)
        {
            writer.WriteUInt16(word);
        }

        foreach (var format in chart.PointFormats)
        {
            writer.WriteBytes(format);
        }
    }

    private static void WriteSeries(BigEndianWriter writer, ChartSeries series, bool lastInList)
    {
        int count = series.PointCount;
        var body = new BigEndianWriter();
        var header = (byte[])series.Header.Clone();
        header[60] = header[62] = (byte)(count >> 8);
        header[61] = header[63] = (byte)count;
        body.WriteBytes(header);
        for (int i = 0; i < count; i++)
        {
            if (series.SeriesType == 3)
            {
                body.WritePascalString(i < series.CategoryTexts.Count ? series.CategoryTexts[i] : string.Empty);
            }
            else
            {
                body.WriteBytes((i < series.CategoryValues.Count ? series.CategoryValues[i] : MicrosoftDecimal.FromDouble(i + 1)).ToBytes());
            }
        }

        int secondArrayOffset = 4 + body.Position;
        for (int i = 0; i < count; i++)
        {
            body.WriteBytes((i < series.Values.Count ? series.Values[i] : MicrosoftDecimal.Zero).ToBytes());
        }

        body.WritePascalString(series.Name);
        body.WritePascalString(series.FirstLabel);
        body.WritePascalString(series.SecondLabel);
        body.WritePascalString(series.SourceMarker);
        body.WritePascalString(series.SourceDocument);
        body.WritePascalString(series.ValuesReference);
        body.WritePascalString(series.CategoriesReference);
        foreach (var text in series.TrailingStrings)
        {
            body.WritePascalString(text);
        }

        var bytes = body.ToArray();
        writer.WriteUInt16((ushort)(lastInList ? 0 : 4 + bytes.Length));
        writer.WriteUInt16(checked((ushort)secondArrayOffset));
        writer.WriteBytes(bytes);
    }
}
