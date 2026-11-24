using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The Page Setup record and its strings (spec §14).
/// </summary>
public sealed class ExcelPrintSetup
{
    /// <summary>The size of the Page Setup record.</summary>
    public const int RecordSize = 122;

    /// <summary>
    /// Creates a print setup from a raw 122-byte record.
    /// </summary>
    public ExcelPrintSetup(byte[] record)
    {
        ArgumentNullException.ThrowIfNull(record);
        if (record.Length != RecordSize)
        {
            throw new ArgumentException($"The Page Setup record must be {RecordSize} bytes.", nameof(record));
        }

        Record = record;
    }

    /// <summary>Gets the raw record. Bytes 0–61 are the start of the Toolbox print record.</summary>
    public byte[] Record { get; }

    /// <summary>Gets or sets the page header, or null if there is none.</summary>
    public string? Header
    {
        get => _header;
        set => _header = PascalText.CheckOptional(value, nameof(value));
    }

    private string? _header;

    /// <summary>Gets or sets the page footer, or null if there is none.</summary>
    public string? Footer
    {
        get => _footer;
        set => _footer = PascalText.CheckOptional(value, nameof(value));
    }

    private string? _footer;

    /// <summary>Gets or sets the printer set-up message for serial printers, or null if there is none.</summary>
    public string? PrinterMessage
    {
        get => _printerMessage;
        set => _printerMessage = PascalText.CheckOptional(value, nameof(value));
    }

    private string? _printerMessage;

    /// <summary>Gets or sets the left margin in inches (spec §14, offset 62).</summary>
    public double LeftMargin { get => GetDouble(62); set => SetDouble(62, value); }

    /// <summary>Gets or sets the right margin in inches (offset 70).</summary>
    public double RightMargin { get => GetDouble(70); set => SetDouble(70, value); }

    /// <summary>Gets or sets the top margin in inches (offset 78).</summary>
    public double TopMargin { get => GetDouble(78); set => SetDouble(78, value); }

    /// <summary>Gets or sets the bottom margin in inches (offset 86).</summary>
    public double BottomMargin { get => GetDouble(86); set => SetDouble(86, value); }

    private double GetDouble(int offset) => BinaryPrimitives.ReadDoubleBigEndian(Record.AsSpan(offset));

    private void SetDouble(int offset, double value) => BinaryPrimitives.WriteDoubleBigEndian(Record.AsSpan(offset), value);

    /// <summary>Gets or sets whether row and column headings are printed.</summary>
    public bool PrintHeadings { get => Record[97] != 0; set => Record[97] = (byte)(value ? 1 : 0); }

    /// <summary>Gets or sets whether gridlines are printed.</summary>
    public bool PrintGridlines { get => Record[101] != 0; set => Record[101] = (byte)(value ? 1 : 0); }

    private ushort GetWord(int offset) => BinaryPrimitives.ReadUInt16BigEndian(Record.AsSpan(offset));

    private void SetWord(int offset, ushort value) => BinaryPrimitives.WriteUInt16BigEndian(Record.AsSpan(offset), value);

    /// <summary>Gets or sets the serial-printer left margin in characters.</summary>
    public ushort TtyLeftMargin { get => GetWord(110); set => SetWord(110, value); }

    /// <summary>Gets or sets the serial-printer print width in characters.</summary>
    public ushort TtyPrintWidth { get => GetWord(112); set => SetWord(112, value); }

    /// <summary>Gets or sets the serial-printer top margin in lines.</summary>
    public ushort TtyTopMargin { get => GetWord(114); set => SetWord(114, value); }

    /// <summary>Gets or sets the serial-printer print length in lines.</summary>
    public ushort TtyPrintLength { get => GetWord(116); set => SetWord(116, value); }
}
