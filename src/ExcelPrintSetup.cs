using System.Buffers.Binary;

namespace MacExcelReader;

/// <summary>
/// The Page Setup record and its strings (spec §13).
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
    public string? Header { get; set; }

    /// <summary>Gets or sets the page footer, or null if there is none.</summary>
    public string? Footer { get; set; }

    /// <summary>Gets or sets the printer set-up message for serial printers, or null if there is none.</summary>
    public string? PrinterMessage { get; set; }

    /// <summary>
    /// Gets a margin as stored: 0 left, 1 right, 2 top, 3 bottom (order inferred from the dialog).
    /// </summary>
    public double GetMargin(int index)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, 3);
        return BinaryPrimitives.ReadDoubleBigEndian(Record.AsSpan(62 + 8 * index));
    }

    /// <summary>Sets a margin; see <see cref="GetMargin"/>.</summary>
    public void SetMargin(int index, double value)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(index, 3);
        BinaryPrimitives.WriteDoubleBigEndian(Record.AsSpan(62 + 8 * index), value);
    }

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
