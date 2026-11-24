using System.Buffers.Binary;
using MacExcelReader.Internal;

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
    /// Creates a numeric series as Excel writes one: its values, and where they come from. Excel 1.x plots only data
    /// from a document, so a series names its source document and the references to the values and, optionally, the
    /// categories on it, in R1C1 form without a leading "=" (spec §16.9).
    /// </summary>
    /// <param name="name">The series name.</param>
    /// <param name="values">The values, which the chart file keeps a copy of.</param>
    /// <param name="document">The source document, such as "Sales".</param>
    /// <param name="valuesReference">The values' reference on that document, such as "R2C1:R2C3".</param>
    /// <param name="categoriesReference">The categories' reference, or null for none (categories are 1, 2, 3…).</param>
    /// <param name="categories">The category values, or null for 1, 2, 3….</param>
    public static ChartSeries Create(string name, IEnumerable<double> values, string document, string valuesReference, string? categoriesReference = null, IEnumerable<double>? categories = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(document);
        ArgumentException.ThrowIfNullOrEmpty(valuesReference);
        var header = new byte[HeaderSize];
        // Series type 1 (numbers), bits 11 and 7, and bit 10: in the first list, the main chart's. Without bit 10
        // Excel 1.00 draws nothing of the chart (spec §16.9).
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(0), 0x2C80);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(2), 0x4000 | (2 << 10) | 0x10);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(4), 0x4000 | (2 << 10));
        header[6] = header[7] = 10;
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), 60);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(10), 60);
        // Bits 15–14: 3, a reference held as formula text; 2, literal text (an empty categories reference).
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(14), categoriesReference is null ? (ushort)0x8000 : (ushort)0xC000);
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
        var series = new ChartSeries(header)
        {
            Name = name,
            FirstLabel = name + ".",
            SecondLabel = name + ".",
            SourceMarker = "\u001B\u0006",
            SourceDocument = document,
            CategoriesReference = categoriesReference ?? string.Empty,
            ValuesReference = valuesReference,
        };
        series.Values.AddRange(values.Select(MicrosoftDecimal.FromDouble));
        var x = categories?.ToList() ?? Enumerable.Range(1, series.Values.Count).Select(i => (double)i).ToList();
        series.CategoryValues.AddRange(x.Select(MicrosoftDecimal.FromDouble));
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

    /// <summary>Gets the category numbers, the first array (empty for a text series; spec §16.9).</summary>
    public List<MicrosoftDecimal> CategoryValues { get; } = [];

    /// <summary>Gets the category strings, the first array of a text series.</summary>
    public List<string> CategoryTexts { get; } = [];

    /// <summary>Gets the values, the second array.</summary>
    public List<MicrosoftDecimal> Values { get; } = [];

    /// <summary>Gets or sets the series name.</summary>
    public string Name
    {
        get => _name;
        set => _name = PascalText.Check(value, nameof(value));
    }

    private string _name = string.Empty;

    /// <summary>Gets or sets the first label string.</summary>
    public string FirstLabel
    {
        get => _firstLabel;
        set => _firstLabel = PascalText.Check(value, nameof(value));
    }

    private string _firstLabel = string.Empty;

    /// <summary>Gets or sets the second label string.</summary>
    public string SecondLabel
    {
        get => _secondLabel;
        set => _secondLabel = PascalText.Check(value, nameof(value));
    }

    private string _secondLabel = string.Empty;

    /// <summary>Gets or sets the source-document marker string (empty, or the two bytes 0x1B 0x06).</summary>
    public string SourceMarker
    {
        get => _sourceMarker;
        set => _sourceMarker = PascalText.Check(value, nameof(value));
    }

    private string _sourceMarker = string.Empty;

    /// <summary>Gets or sets the name of the source document.</summary>
    public string SourceDocument
    {
        get => _sourceDocument;
        set => _sourceDocument = PascalText.Check(value, nameof(value));
    }

    private string _sourceDocument = string.Empty;

    /// <summary>Gets or sets the reference to the values, as R1C1 formula text without <c>=</c> or the document.</summary>
    public string ValuesReference
    {
        get => _valuesReference;
        set => _valuesReference = PascalText.Check(value, nameof(value));
    }

    private string _valuesReference = string.Empty;

    /// <summary>Gets or sets the reference to the categories, as R1C1 formula text without <c>=</c> or the document.</summary>
    public string CategoriesReference
    {
        get => _categoriesReference;
        set => _categoriesReference = PascalText.Check(value, nameof(value));
    }

    private string _categoriesReference = string.Empty;

    /// <summary>Gets the four trailing strings (empty in files written by Excel).</summary>
    public string[] TrailingStrings { get; } = [string.Empty, string.Empty, string.Empty, string.Empty];

    /// <summary>Gets the number of points.</summary>
    public int PointCount => Math.Max(SeriesType == 3 ? CategoryTexts.Count : CategoryValues.Count, Values.Count);
}
