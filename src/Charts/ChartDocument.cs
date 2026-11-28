using System.Runtime.InteropServices;
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
    public ReadOnlyMemory<byte> CodeSegment { get; set; }

    /// <summary>Gets or sets the 8-byte chart-set header (skipped by Excel on load).</summary>
    public byte[] ChartSetHeader { get; set; } = [0, 0, 0, 0, 0, 1, 0, 0x11];

    /// <summary>
    /// Gets the charts in file order: when there is an overlay chart it comes first, then the main chart.
    /// </summary>
    public List<Chart> Charts { get; } = [];

    /// <summary>Gets the main chart (the last in file order).</summary>
    /// <exception cref="InvalidOperationException"><see cref="Charts"/> is empty.</exception>
    public Chart MainChart => Charts.Count > 0 ? Charts[^1] : throw new InvalidOperationException("The document has no charts.");

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

        // A chart file holds the main chart and, if there is one, the overlay (spec §16.1).
        if (chartCount is 0 or > 2)
        {
            throw new ExcelFormatException($"A chart file has one or two charts, not {chartCount}.");
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


    /// <summary>Reads a chart document from its bytes. The bytes are copied.</summary>
    public static ChartDocument Read(ReadOnlySpan<byte> data) => Read(data.ToArray());

    /// <summary>Reads a chart document from a stream, from its current position to its end.</summary>
    public static ChartDocument Read(Stream stream) => Read(StreamBytes.ReadToEnd(stream));

    /// <summary>Reads a chart document from a stream asynchronously, from its current position to its end.</summary>
    public static async Task<ChartDocument> ReadAsync(Stream stream, CancellationToken cancellationToken = default)
        => Read(await StreamBytes.ReadToEndAsync(stream, cancellationToken).ConfigureAwait(false));

    /// <summary>Writes a chart document to a stream.</summary>
    public void Write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        stream.Write(ToArray());
    }

    /// <summary>Writes a chart document to a stream asynchronously.</summary>
    public async Task WriteAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        await stream.WriteAsync(ToArray(), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Serialises the chart document.
    /// </summary>
    public byte[] ToArray()
    {
        if (Charts.Count is 0 or > 2)
        {
            throw new InvalidOperationException("A chart document has one or two charts.");
        }

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
        writer.WriteBytes(CodeSegment.Span);

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
        if (chart.PointFormats.Count == 0)
        {
            chart.PointFormatListHead = 0;
        }

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

    /// <summary>
    /// Gets the series a chart plots: all of them for a lone chart; with an overlay, the main chart has those before
    /// the overlay's first series (<see cref="Chart.FirstOverlaySeries"/>, counted from 1) and the overlay the rest.
    /// </summary>
    public IReadOnlyList<ChartSeries> GetSeries(Chart chart) => SeriesOf(chart).ToArray();

    internal ReadOnlySpan<ChartSeries> SeriesOf(Chart chart)
    {
        var all = CollectionsMarshal.AsSpan(Series);
        if (OverlayChart is not { } overlay)
        {
            return all;
        }

        int first = Math.Clamp(overlay.FirstOverlaySeries - 1, 0, all.Length);
        return chart == overlay ? all[first..] : all[..first];
    }

    /// <summary>
    /// Gets a chart text's font as Excel 1.00 opens it: the Macintosh font number and size Excel hides in the axis
    /// record (<see cref="ChartAxis.SavedFontNumber"/>, <see cref="ChartAxis.SavedFontSize"/>), or for a Microsoft Chart
    /// file, whose version lacks bit 2, the font converted back from its codes (spec §16.4).
    /// </summary>
    public ChartFont GetLabelFont(ChartAxis axis) => axis.GetFont(Version);

    /// <summary>
    /// Gets the range of the data an axis of a chart shows, as Excel 1.00 takes it for an automatic scale
    /// (<c>37.16</c>, <c>37.2f</c>): the smallest and largest values; for a stacked chart 0 and the largest total of
    /// a category's values, each counted positive; 0 and 100 for a 100% chart. A scatter chart's category axis takes
    /// the series' category numbers (or 1, 2, … for a series without them). A logarithmic axis leaves out values that
    /// are not positive, and has 1 where a linear one has 0. With no values both ends are 0 (1 on a logarithmic
    /// axis). When the overlay shares the axis, both charts get the range of their data together.
    /// </summary>
    public (double Minimum, double Maximum) GetDataRange(Chart chart, ChartAxisKind axis = ChartAxisKind.Value)
    {
        if (OverlayChart is { } overlay && Shares(overlay, axis))
        {
            if (SeriesOf(MainChart).IsEmpty)
            {
                return GetOwnDataRange(overlay, axis);
            }

            var (min, max) = GetOwnDataRange(MainChart, axis);
            if (SeriesOf(overlay).IsEmpty)
            {
                return (min, max);
            }

            var (overlayMin, overlayMax) = GetOwnDataRange(overlay, axis);
            return (min - overlayMin < 0 ? min : overlayMin, max - overlayMax > 0 ? max : overlayMax);
        }

        return GetOwnDataRange(chart, axis);
    }

    /// <summary>
    /// Gets how many tick labels Excel 1.00 reckons fit along an axis of a chart it has drawn (<c>37.0a</c>), from the
    /// layout it saved: the plot area it counts against is the saved one plus the room the axis labels and ticks took
    /// (<c>37.0e</c>), with category labels wrapped to their slots, all measured in the axes' fonts. Along a vertical
    /// axis it is that height, less two of the other axis's label heights, over a label height; along a horizontal one
    /// that width over three widths of 0. A linear scale takes 1 to 10 intervals from it.
    /// </summary>
    /// <param name="chart">The main chart or the overlay.</param>
    /// <param name="fonts">The System's fonts, in which Excel measures the labels.</param>
    /// <param name="axis">The value axis, or a scatter chart's category (X) axis.</param>
    public int GetLabelSpaces(Chart chart, Rendering.MacFontSet fonts, ChartAxisKind axis = ChartAxisKind.Value)
        => new ChartLayoutEngine(this, chart, fonts).GetLabelSpacesFromSaved(axis == ChartAxisKind.Category);

    /// <summary>
    /// Lays a chart out as Excel 1.00 does when it draws it, from its chart area (or, for a chart Excel has not
    /// drawn, the default 15/16 of the drawing's width by 23/25 of its height): the border, the title, the tick
    /// length, how many tick labels fit, and the plot area left when the labels have taken their room. The labels are
    /// measured from the scales as they are; <see cref="UpdateAutomaticScales(Rendering.MacFontSet)"/> works them out.
    /// Legends and axis titles are not covered.
    /// </summary>
    public ChartLayout LayOut(Chart chart, Rendering.MacFontSet fonts) => new ChartLayoutEngine(this, chart, fonts).LayOut();

    /// <summary>
    /// Gets the size, in twips, of the box Excel 1.00 lays out for an axis's largest tick label (<c>37.09</c>): its
    /// lines of text and widest line with a pixel of margin all round. Number labels run from the scale's minimum to its
    /// maximum in major units, in General, on one line each. Category labels wrap to a category's share of the plot
    /// area's width (a third of the chart's width in a bar chart): a word at a time, a word running to the next space
    /// or hyphen, and a word too wide for a line breaking between characters. A hidden axis, or one without labels,
    /// has none.
    /// </summary>
    /// <param name="chart">The main chart or the overlay.</param>
    /// <param name="fonts">The System's fonts, in which Excel measures the labels.</param>
    /// <param name="axis">The value axis, or the category axis (a scatter chart's X axis).</param>
    public (int Height, int Width) GetTickLabelSize(Chart chart, Rendering.MacFontSet fonts, ChartAxisKind axis = ChartAxisKind.Value)
        => new ChartLayoutEngine(this, chart, fonts).GetTickLabelSize(axis == ChartAxisKind.Category);

    /// <summary>
    /// Works out every automatic scale field of the document's charts as Excel 1.00 does when it draws them in their
    /// windows (a chart Excel has not drawn, at the default layout of <see cref="LayOut"/>): first whether the overlay shares the main chart's scales (<c>37.34</c>), then each value axis and
    /// scatter X axis from its data (<see cref="GetDataRange"/>) and the tick labels that fit along it
    /// (<see cref="GetLabelSpaces"/>). Fixed fields stay as they are. A pie chart, which has no axes, and a chart with
    /// no series keep their scales.
    /// </summary>
    /// <param name="fonts">The System's fonts, in which Excel measures the labels.</param>
    public void UpdateAutomaticScales(Rendering.MacFontSet fonts)
    {
        ArgumentNullException.ThrowIfNull(fonts);
        UpdateSharedScales();
        foreach (var chart in Charts)
        {
            if (HasScales(chart))
            {
                // Both counts come from the saved layout, before either scale changes; a chart Excel has not drawn is
                // laid out afresh, which the counts do not depend on the scales for.
                int value, category;
                if (chart.PlotArea.Box.Height > 0 && chart.PlotArea.Box.Width > 0)
                {
                    value = GetLabelSpaces(chart, fonts);
                    category = chart.XScale is null ? 0 : GetLabelSpaces(chart, fonts, ChartAxisKind.Category);
                }
                else
                {
                    var layout = LayOut(chart, fonts);
                    (value, category) = (layout.ValueLabelSpaces, layout.CategoryLabelSpaces);
                }

                UpdateScale(chart, ChartAxisKind.Value, value);
                if (chart.XScale is not null)
                {
                    UpdateScale(chart, ChartAxisKind.Category, category);
                }
            }
        }
    }

    /// <summary>
    /// Works out every automatic scale field of the document's charts as <see cref="UpdateAutomaticScales(Rendering.MacFontSet)"/>
    /// does, with the same number of tick labels fitting along every axis: 10 or more gives a linear scale up to 10
    /// intervals, as Excel draws a chart in a window of its usual size.
    /// </summary>
    public void UpdateAutomaticScales(int labelSpaces = 10)
    {
        UpdateSharedScales();
        foreach (var chart in Charts)
        {
            if (HasScales(chart))
            {
                UpdateScale(chart, ChartAxisKind.Value, labelSpaces);
                if (chart.XScale is not null)
                {
                    UpdateScale(chart, ChartAxisKind.Category, labelSpaces);
                }
            }
        }
    }

    private bool HasScales(Chart chart) => chart.Type != ChartType.Pie && !SeriesOf(chart).IsEmpty;

    private static bool Shares(Chart overlay, ChartAxisKind axis)
        => axis == ChartAxisKind.Category ? overlay.SharesXAxisScale : overlay.SharesValueAxisScale;

    // 37.0a for one axis. On a shared axis, the scale is a scatter chart's only when both charts are scatter charts.
    private void UpdateScale(Chart chart, ChartAxisKind axis, int labelSpaces)
    {
        var scale = axis == ChartAxisKind.Category ? chart.XScale! : chart.ValueScale;
        var (min, max) = GetDataRange(chart, axis);
        bool scatter = chart.Type == ChartType.Scatter;
        if (OverlayChart is { } overlay && Shares(overlay, axis))
        {
            scatter &= MainChart.Type == ChartType.Scatter && overlay.Type == ChartType.Scatter;
        }

        scale.SetAutomaticValues(min, max, labelSpaces, scatter);
    }

    // 37.34: with an overlay and series in both charts, whether the overlay shares each axis. The value axis is shared
    // when either chart's value axis is hidden, neither is a pie, and both are bars or neither, and 100% or neither;
    // the category axis when both are scatter charts, or when either's category axis is hidden, neither is a pie or a
    // scatter chart, and both are bars or neither. The chart whose axis is hidden takes the other's logarithmic and
    // reverse-order settings, and on a shared category axis the overlay takes the main chart's crossing.
    private void UpdateSharedScales()
    {
        if (OverlayChart is not { } overlay || SeriesOf(MainChart).IsEmpty || SeriesOf(overlay).IsEmpty)
        {
            return;
        }

        var main = MainChart;
        bool noPie = main.Type != ChartType.Pie && overlay.Type != ChartType.Pie;
        bool overlayValueHidden = overlay.ValueAxis.Line.Pattern == ChartPattern.None;
        overlay.SharesValueAxisScale = (overlayValueHidden || main.ValueAxis.Line.Pattern == ChartPattern.None)
            && noPie && main.IsBar == overlay.IsBar && main.IsPercent == overlay.IsPercent;
        if (overlay.SharesValueAxisScale)
        {
            var (from, to) = overlayValueHidden ? (main.ValueScale, overlay.ValueScale) : (overlay.ValueScale, main.ValueScale);
            to.IsLogarithmic = from.IsLogarithmic;
            to.ValuesInReverseOrder = from.ValuesInReverseOrder;
        }

        bool overlayCategoryHidden = overlay.CategoryAxis.Line.Pattern == ChartPattern.None;
        bool mainScatter = main.Type == ChartType.Scatter, overlayScatter = overlay.Type == ChartType.Scatter;
        overlay.SharesXAxisScale = mainScatter && overlayScatter
            || (overlayCategoryHidden || main.CategoryAxis.Line.Pattern == ChartPattern.None) && noPie
                && !mainScatter && !overlayScatter && main.IsBar == overlay.IsBar;
        if (overlay.SharesXAxisScale)
        {
            var (from, to) = overlayCategoryHidden ? (main, overlay) : (overlay, main);
            bool reverse = from.XScale is { } fromScale ? fromScale.ValuesInReverseOrder : from.CategoriesInReverseOrder;
            if (to.XScale is { } toScale)
            {
                toScale.ValuesInReverseOrder = reverse;
            }
            else
            {
                to.CategoriesInReverseOrder = reverse;
            }

            if (overlayScatter && from.XScale is { } fromX && to.XScale is { } toX)
            {
                toX.IsLogarithmic = fromX.IsLogarithmic;
            }

            overlay.ValueAxisCrossesBetweenCategories = main.ValueAxisCrossesBetweenCategories;
        }
    }

    // 37.16 for one chart.
    private (double Minimum, double Maximum) GetOwnDataRange(Chart chart, ChartAxisKind axis)
    {
        bool scatter = chart.Type == ChartType.Scatter;
        bool category = axis == ChartAxisKind.Category;
        var scale = category ? chart.XScale ?? chart.ValueScale : chart.ValueScale;
        bool log = scale.IsLogarithmic;
        double empty = log ? 1 : 0;
        bool stacked = !scatter && chart.IsStacked;
        if (stacked && chart.IsPercent)
        {
            return (empty, 100);
        }

        var series = SeriesOf(chart);
        double min = double.PositiveInfinity, max = double.NegativeInfinity;
        if (stacked)
        {
            int points = 0;
            foreach (var s in series)
            {
                points = Math.Max(points, s.Values.Count);
            }

            for (int i = 0; i < points; i++)
            {
                double total = 0;
                foreach (var s in series)
                {
                    if (i < s.Values.Count)
                    {
                        total += Math.Abs(s.Values[i].ToDouble());
                    }
                }

                Add(total, ref min, ref max);
            }
        }
        else
        {
            foreach (var s in series)
            {
                for (int i = 0; i < s.Values.Count; i++)
                {
                    Add(!category ? s.Values[i].ToDouble()
                        : s.SeriesType != 3 && i < s.CategoryValues.Count ? s.CategoryValues[i].ToDouble() : i + 1, ref min, ref max);
                }
            }
        }

        return min > max ? (empty, empty) : (stacked ? empty : min, max);

        void Add(double value, ref double min, ref double max)
        {
            if (log && value <= 0)
            {
                return;
            }

            min = Math.Min(min, value);
            max = Math.Max(max, value);
        }
    }
}
