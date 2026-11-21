using System.Buffers.Binary;

namespace MacExcelReader.Charts;

/// <summary>
/// A series record (spec §16.9): a copy of the plotted data together with its references.
/// </summary>
public sealed class ChartSeries
{
    /// <summary>The size of the series header.</summary>
    public const int HeaderSize = 98;

    /// <summary>
    /// Creates a series from a raw header.
    /// </summary>
    public ChartSeries(byte[] header)
    {
        ArgumentNullException.ThrowIfNull(header);
        if (header.Length != HeaderSize)
        {
            throw new ArgumentException($"A series header is {HeaderSize} bytes.", nameof(header));
        }

        Header = header;
    }

    /// <summary>
    /// Creates a numeric series with the constants Excel writes.
    /// </summary>
    public static ChartSeries Create(IEnumerable<double> values, string name, string reference)
    {
        var header = new byte[HeaderSize];
        // Series type 1 (numbers), bits 11 and 7, and bit 10: in the first list, the main chart's. Without bit 10
        // Excel 1.00 draws nothing of the chart (spec §16.9).
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(0), 0x2C80);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(2), 0x4000 | (2 << 10) | 0x10);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(4), 0x4000 | (2 << 10));
        header[6] = header[7] = 10;
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), 60);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(10), 60);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(14), 0xC000);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(16), 0xC000);
        for (int i = 0; i < 4; i++)
        {
            MicrosoftDecimal.FromDouble(1).WriteTo(header.AsSpan(26 + 8 * i));
        }

        header[65] = header[66] = 1;
        for (int i = 0; i < 11; i++)
        {
            BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(68 + 2 * i), 2);
        }

        // Excel writes each label as the series name and "." (spec §16.9).
        var series = new ChartSeries(header) { Name = name, FirstReference = reference, FirstLabel = name + ".", SecondLabel = name + "." };
        series.SecondValues.AddRange(values.Select(MicrosoftDecimal.FromDouble));
        for (int i = 1; i <= series.SecondValues.Count; i++)
        {
            series.FirstValues.Add(MicrosoftDecimal.FromDouble(i));
        }

        return series;
    }

    /// <summary>Gets the raw 98-byte header.</summary>
    public byte[] Header { get; }

    /// <summary>Gets or sets the series type: 0 dates, 1 or 2 numbers, 3 text categories.</summary>
    public int SeriesType
    {
        get => Header[0] >> 5;
        set => Header[0] = (byte)((Header[0] & 0x1F) | ((value & 7) << 5));
    }

    /// <summary>Gets or sets whether the series belongs to the first list.</summary>
    public bool InFirstList
    {
        get => (Header[0] & 0x04) != 0;
        set => Header[0] = (byte)(value ? Header[0] | 0x04 : Header[0] & ~0x04);
    }

    /// <summary>Gets the first array's numbers (empty for a text series).</summary>
    public List<MicrosoftDecimal> FirstValues { get; } = [];

    /// <summary>Gets the first array's strings (text series only).</summary>
    public List<string> FirstTexts { get; } = [];

    /// <summary>Gets the second array's numbers.</summary>
    public List<MicrosoftDecimal> SecondValues { get; } = [];

    /// <summary>Gets or sets the series name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the first label string.</summary>
    public string FirstLabel { get; set; } = string.Empty;

    /// <summary>Gets or sets the second label string.</summary>
    public string SecondLabel { get; set; } = string.Empty;

    /// <summary>Gets or sets the source-document marker string (empty, or the two bytes 0x1B 0x06).</summary>
    public string SourceMarker { get; set; } = string.Empty;

    /// <summary>Gets or sets the name of the source document.</summary>
    public string SourceDocument { get; set; } = string.Empty;

    /// <summary>Gets or sets the second reference, as formula text.</summary>
    public string SecondReference { get; set; } = string.Empty;

    /// <summary>Gets or sets the first reference, as formula text.</summary>
    public string FirstReference { get; set; } = string.Empty;

    /// <summary>Gets the four trailing strings (empty in files written by Excel).</summary>
    public string[] TrailingStrings { get; } = [string.Empty, string.Empty, string.Empty, string.Empty];

    /// <summary>Gets the number of points.</summary>
    public int PointCount => Math.Max(SeriesType == 3 ? FirstTexts.Count : FirstValues.Count, SecondValues.Count);
}
