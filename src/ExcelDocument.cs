using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A user-defined number format (spec §6.1): a 7-byte descriptor and the tokenised body.
/// </summary>
/// <param name="Descriptor">Three section offsets, three section lengths and the single-section flag.</param>
/// <param name="Body">The tokenised format.</param>
public sealed record ExcelNumberFormat(byte[] Descriptor, byte[] Body);

/// <summary>
/// An Excel 1.x for Macintosh worksheet or macro sheet.
/// </summary>
public sealed class ExcelDocument
{
    /// <summary>The magic number at the start of every worksheet and macro sheet.</summary>
    public const ushort Magic = 0x5327;

    /// <summary>The only format version Excel 1.x reads.</summary>
    public const ushort FormatVersion = 17;

    /// <summary>The code signature of Excel 1.03.</summary>
    public const ushort Excel103Signature = 0xAB27;

    private ExcelDocument(ExcelSheetState state)
    {
        State = state;
    }

    /// <summary>Gets or sets the code signature of the build that wrote the file (header word 4).</summary>
    public ushort Signature { get; set; } = Excel103Signature;

    /// <summary>Gets the six zero bytes that follow the signature.</summary>
    public byte[] HeaderReserved { get; private set; } = new byte[6];

    /// <summary>
    /// Gets the embedded p-code segments (spec §3.1). They are only used when another build of Excel links
    /// to this document while it is closed. Documents created with this library have empty segments.
    /// </summary>
    public List<byte[]> CodeSegments { get; } = [];

    /// <summary>Gets or sets the function count of the writing build (208 for Excel 1.x).</summary>
    public ushort FunctionCount { get; set; } = (ushort)ExcelFunctions.Count;

    /// <summary>Gets the sheet state block.</summary>
    public ExcelSheetState State { get; }

    /// <summary>Gets or sets whether this is a worksheet or a macro sheet.</summary>
    public ExcelDocumentKind Kind { get; set; } = ExcelDocumentKind.Worksheet;

    /// <summary>Gets or sets the extra header bytes (byte 0 is the system code).</summary>
    public byte[] ExtraHeader { get; set; } = [0];

    /// <summary>Gets or sets the number of built-in number formats of the writing build (19).</summary>
    public byte BuiltInFormatCount { get; set; } = 19;

    /// <summary>Gets the user-defined number formats.</summary>
    public List<ExcelNumberFormat> NumberFormats { get; } = [];

    /// <summary>
    /// Gets or sets whether the external-document table is written. It is written whenever
    /// <see cref="ExternalDocuments"/> is non-empty.
    /// </summary>
    public bool HasExternalDocumentTable { get; set; }

    /// <summary>Gets the documents this sheet refers to.</summary>
    public List<ExcelExternalDocument> ExternalDocuments { get; } = [];

    /// <summary>Gets the names this sheet uses from external documents (section 3).</summary>
    public List<string> ExternalNames { get; } = [];

    /// <summary>
    /// Gets the per-row size hints of the recalculation dependency index (spec §6.2). Excel rebuilds the
    /// index on open; zeros are always valid.
    /// </summary>
    public List<ushort> RowIndex { get; } = [];

    /// <summary>Gets the cells.</summary>
    public ExcelCellGrid Cells { get; } = new();

    /// <summary>Gets the defined names, in file order. Name tokens refer to them by 1-based index.</summary>
    public List<ExcelDefinedName> Names { get; } = [];

    /// <summary>Gets the saved windows.</summary>
    public List<ExcelWindow> Windows { get; } = [];

    /// <summary>Gets the column page breaks, as written in byte-sized chunks.</summary>
    public List<byte> ColumnPageBreaks { get; } = [];

    /// <summary>Gets the row page breaks.</summary>
    public List<ushort> RowPageBreaks { get; } = [];

    /// <summary>Gets or sets the print setup, or null if there is none.</summary>
    public ExcelPrintSetup? PrintSetup { get; set; }

    /// <summary>
    /// Creates a new, empty worksheet with the defaults Excel 1.03 writes.
    /// </summary>
    public static ExcelDocument CreateWorksheet() => Create(ExcelDocumentKind.Worksheet);

    /// <summary>
    /// Creates a new, empty macro sheet.
    /// </summary>
    public static ExcelDocument CreateMacroSheet() => Create(ExcelDocumentKind.MacroSheet);

    private static ExcelDocument Create(ExcelDocumentKind kind)
    {
        var document = new ExcelDocument(ExcelSheetState.CreateDefault()) { Kind = kind };
        document.CodeSegments.Add([]);
        document.CodeSegments.Add([]);
        foreach (var name in new[] { "Criteria", "Database", "Print_Area", "Print_Titles", "Recorder" })
        {
            document.Names.Add(new ExcelDefinedName(name));
        }

        document.Windows.Add(ExcelWindow.CreateDefault(kind == ExcelDocumentKind.MacroSheet));
        return document;
    }

    /// <summary>
    /// Reads a document from a file.
    /// </summary>
    public static ExcelDocument Load(string path) => Read(File.ReadAllBytes(path));

    /// <summary>
    /// Reads a document from a stream.
    /// </summary>
    public static ExcelDocument Read(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Read(buffer.ToArray());
    }

    /// <summary>
    /// Reads a document from its bytes (the data fork of the Macintosh file).
    /// </summary>
    public static ExcelDocument Read(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        var reader = new BigEndianReader(data);
        if (data.Length < 24 || reader.ReadUInt16() != Magic)
        {
            throw new ExcelFormatException("Not an Excel 1.x worksheet or macro sheet (bad magic number).");
        }

        ushort version = reader.ReadUInt16();
        if (version != FormatVersion)
        {
            throw new ExcelFormatException($"Unsupported format version {version}; expected {FormatVersion}.");
        }

        ushort signature = reader.ReadUInt16();
        var headerReserved = reader.ReadBytes(6);
        int segmentCount = reader.ReadUInt16();
        var segmentInfo = new (int Length, int Offset)[segmentCount];
        for (int i = 0; i < segmentCount; i++)
        {
            segmentInfo[i] = (reader.ReadUInt16(), reader.ReadUInt16());
        }

        int dataOffset = reader.ReadUInt16();
        reader.Position = dataOffset;
        ushort functionCount = reader.ReadUInt16();
        var document = new ExcelDocument(new ExcelSheetState(reader.ReadBytes(ExcelSheetState.Size)))
        {
            Signature = signature,
            HeaderReserved = headerReserved,
            FunctionCount = functionCount,
            Kind = (ExcelDocumentKind)reader.ReadByte(),
        };

        foreach (var (length, offset) in segmentInfo)
        {
            if (offset + length > data.Length)
            {
                throw new ExcelFormatException("An embedded code segment lies outside the file.");
            }

            document.CodeSegments.Add(data.AsSpan(offset, length).ToArray());
        }

        var sections = new int[8];
        for (int i = 0; i < sections.Length; i++)
        {
            uint position = reader.ReadUInt32();
            sections[i] = (int)((position >> 16) * 512 + (position & 0xFFFF));
            if (sections[i] > data.Length)
            {
                throw new ExcelFormatException($"Section {i} starts beyond the end of the file.");
            }
        }

        document.ExtraHeader = reader.ReadBytes(reader.ReadUInt16());
        int rows = document.State.RowCount;
        int columns = document.State.ColumnCountInTable;

        document.ReadNumberFormats(new BigEndianReader(data, sections[0]));
        document.ReadExternalDocuments(new BigEndianReader(data, sections[1]), rows);
        document.ReadCells(data, sections[2], sections[5], rows, columns);
        if (sections[3] < sections[4])
        {
            var namesReader = new BigEndianReader(data, sections[3]);
            for (byte length; (length = namesReader.ReadByte()) != 0;)
            {
                document.ExternalNames.Add(MacRoman.GetString(namesReader.ReadBytes(length)));
            }
        }

        document.ReadNames(new BigEndianReader(data, sections[4]));
        document.ReadWindows(new BigEndianReader(data, sections[6]));
        document.ReadPrintSettings(new BigEndianReader(data, sections[7]));
        return document;
    }

    private void ReadNumberFormats(BigEndianReader reader)
    {
        BuiltInFormatCount = reader.ReadByte();
        int count = reader.ReadByte();
        for (int i = 0; i < count; i++)
        {
            var descriptor = reader.ReadBytes(7);
            int bodyLength = 0;
            for (int s = 0; s < 3; s++)
            {
                bodyLength = Math.Max(bodyLength, descriptor[s] + descriptor[3 + s]);
            }

            NumberFormats.Add(new ExcelNumberFormat(descriptor, reader.ReadBytes(bodyLength)));
        }
    }

    private void ReadExternalDocuments(BigEndianReader reader, int rows)
    {
        ushort tableSize = reader.ReadUInt16();
        HasExternalDocumentTable = tableSize != 0;
        if (HasExternalDocumentTable)
        {
            int count = reader.ReadUInt16();
            for (int i = 0; i < count; i++)
            {
                var reserved = reader.ReadBytes(6);
                ExternalDocuments.Add(new ExcelExternalDocument(reader.ReadPascalString()) { Reserved = reserved });
            }
        }

        for (int r = 0; r < rows; r++)
        {
            RowIndex.Add(reader.ReadUInt16());
        }
    }

    private void ReadCells(byte[] data, int tableStart, int cellDataStart, int rows, int columns)
    {
        var table = new BigEndianReader(data, tableStart);
        Cells.Load(rows, columns, (row, column) =>
        {
            var cell = new ExcelCell(table.ReadUInt16(), table.ReadUInt16(), table.ReadUInt16());
            int offset = (cell.Word0 << 4) + (cell.Word1 >> 12);
            switch (cell.Type)
            {
                case ExcelCellType.Number:
                    cell.LoadNumber(new BigEndianReader(data, cellDataStart + offset).ReadDouble());
                    break;
                case ExcelCellType.Text:
                    cell.LoadText(new BigEndianReader(data, cellDataStart + offset).ReadPascalString());
                    break;
                case ExcelCellType.NumericFormula:
                case ExcelCellType.Formula:
                    try
                    {
                        cell.LoadFormula(ExcelFormula.Read(new BigEndianReader(data, cellDataStart + offset)));
                    }
                    catch (ExcelFormatException ex)
                    {
                        throw new ExcelFormatException($"Invalid formula in row {row + 1}, column {column + 1}: {ex.Message}", ex);
                    }

                    break;
            }

            return cell;
        });
    }

    private void ReadNames(BigEndianReader reader)
    {
        for (byte length; (length = reader.ReadByte()) != 0;)
        {
            Names.Add(ExcelDefinedName.Read(reader, length));
        }
    }

    private void ReadWindows(BigEndianReader reader)
    {
        for (byte marker; (marker = reader.ReadByte()) != 0;)
        {
            if (marker != 0xFF)
            {
                throw new ExcelFormatException($"Unexpected window marker 0x{marker:X2}.");
            }

            Windows.Add(ExcelWindow.Read(reader));
        }
    }

    private void ReadPrintSettings(BigEndianReader reader)
    {
        for (byte count; (count = reader.ReadByte()) != 0;)
        {
            for (int i = 0; i < count; i++)
            {
                ColumnPageBreaks.Add(reader.ReadByte());
            }
        }

        for (byte count; (count = reader.ReadByte()) != 0;)
        {
            for (int i = 0; i < count; i++)
            {
                RowPageBreaks.Add(reader.ReadUInt16());
            }
        }

        if (reader.ReadByte() != 0)
        {
            PrintSetup = new ExcelPrintSetup(reader.ReadBytes(ExcelPrintSetup.RecordSize))
            {
                Header = ReadOptionalString(reader),
                Footer = ReadOptionalString(reader),
                PrinterMessage = ReadOptionalString(reader),
            };
        }
    }

    private static string? ReadOptionalString(BigEndianReader reader)
    {
        var bytes = reader.ReadPascalBytes();
        return bytes.Length == 0 ? null : MacRoman.GetString(bytes);
    }

    /// <summary>
    /// Writes the document to a file (the data fork). Set the Finder type to XLBN for a worksheet or
    /// XLPG for a macro sheet, and the creator to XCEL.
    /// </summary>
    public void Save(string path) => File.WriteAllBytes(path, ToArray());

    /// <summary>
    /// Writes the document to a stream.
    /// </summary>
    public void Write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        stream.Write(ToArray());
    }

    /// <summary>
    /// Gets the Finder file type for this document.
    /// </summary>
    public string FinderType => Kind == ExcelDocumentKind.MacroSheet ? "XLPG" : "XLBN";

    /// <summary>
    /// Serialises the document.
    /// </summary>
    public byte[] ToArray()
    {
        int rows = Cells.ColumnCount == 0 ? 0 : Cells.RowCount;
        int columns = rows == 0 ? 0 : Cells.ColumnCount;
        State.SetExtent(rows, columns);
        while (RowIndex.Count < rows)
        {
            RowIndex.Add(0);
        }

        var writer = new BigEndianWriter();
        writer.WriteUInt16(Magic);
        writer.WriteUInt16(FormatVersion);
        writer.WriteUInt16(Signature);
        writer.WriteBytes(HeaderReserved);
        writer.WriteUInt16((ushort)CodeSegments.Count);
        int headerSize = 14 + 4 * CodeSegments.Count + 2;
        int segmentOffset = headerSize;
        foreach (var segment in CodeSegments)
        {
            writer.WriteUInt16((ushort)segment.Length);
            writer.WriteUInt16((ushort)segmentOffset);
            segmentOffset += segment.Length;
        }

        writer.WriteUInt16(checked((ushort)segmentOffset));
        foreach (var segment in CodeSegments)
        {
            writer.WriteBytes(segment);
        }

        writer.WriteUInt16(FunctionCount);
        writer.WriteBytes(State.RawData);
        writer.WriteByte((byte)Kind);
        int sectionTable = writer.Position;
        writer.WriteZeros(32);
        writer.WriteUInt16((ushort)ExtraHeader.Length);
        writer.WriteBytes(ExtraHeader);

        // Cell data is built first so the cell table can point into it.
        var cellData = new BigEndianWriter();
        var offsets = new int[rows * columns];
        for (int r = 0, k = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++, k++)
            {
                var cell = Cells.TryGet(r, c)!;
                offsets[k] = cellData.Position;
                switch (cell.Type)
                {
                    case ExcelCellType.Number:
                        cellData.WriteDouble(cell.NumberValue);
                        break;
                    case ExcelCellType.Text:
                        cellData.WritePascalString(cell.TextValue);
                        break;
                    case ExcelCellType.NumericFormula:
                    case ExcelCellType.Formula:
                        cell.Formula!.Write(cellData, r, c);
                        break;
                }
            }
        }

        var sections = new int[8];

        // Section 0: number formats.
        sections[0] = writer.Position;
        writer.WriteByte(BuiltInFormatCount);
        writer.WriteByte(checked((byte)NumberFormats.Count));
        foreach (var format in NumberFormats)
        {
            writer.WriteBytes(format.Descriptor);
            writer.WriteBytes(format.Body);
        }

        // Section 1: external documents, then the row index.
        sections[1] = writer.Position;
        bool writeTable = HasExternalDocumentTable || ExternalDocuments.Count > 0;
        if (writeTable)
        {
            writer.WriteUInt16((ushort)(ExternalDocuments.Count * 12 + 8));
            writer.WriteUInt16((ushort)ExternalDocuments.Count);
            foreach (var external in ExternalDocuments)
            {
                writer.WriteBytes(external.Reserved);
                writer.WritePascalString(external.Name);
            }
        }
        else
        {
            writer.WriteUInt16(0);
        }

        for (int r = 0; r < rows; r++)
        {
            writer.WriteUInt16(RowIndex[r]);
        }

        // Section 2: cell table.
        sections[2] = writer.Position;
        for (int r = 0, k = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++, k++)
            {
                var cell = Cells.TryGet(r, c)!;
                ushort word0 = cell.Word0;
                ushort word1 = cell.Word1;
                if (cell.Type is ExcelCellType.Number or ExcelCellType.Text or ExcelCellType.NumericFormula or ExcelCellType.Formula)
                {
                    word0 = checked((ushort)(offsets[k] >> 4));
                    word1 = (ushort)((word1 & 0x0FFF) | ((offsets[k] & 0xF) << 12));
                }

                writer.WriteUInt16(word0);
                writer.WriteUInt16(word1);
                writer.WriteUInt16(cell.Word2);
            }
        }

        // Section 3: external names.
        // Excel writes the list (with its terminator) only when there are external documents; a table
        // with no entries has an empty section 3.
        sections[3] = writer.Position;
        if (ExternalDocuments.Count > 0 || ExternalNames.Count > 0)
        {
            foreach (var name in ExternalNames)
            {
                writer.WritePascalString(name);
            }

            writer.WriteByte(0);
        }

        // Section 4: defined names.
        sections[4] = writer.Position;
        foreach (var name in Names)
        {
            name.Write(writer);
        }

        writer.WriteByte(0);

        // Section 5: cell data.
        sections[5] = writer.Position;
        writer.WriteBytes(cellData.ToArray());

        // Section 6: windows.
        sections[6] = writer.Position;
        foreach (var window in Windows)
        {
            window.Write(writer);
        }

        writer.WriteByte(0);

        // Section 7: page breaks and print setup.
        sections[7] = writer.Position;
        WriteChunks(writer, ColumnPageBreaks, (w, b) => w.WriteByte(b));
        WriteChunks(writer, RowPageBreaks, (w, v) => w.WriteUInt16(v));
        if (PrintSetup is null)
        {
            writer.WriteByte(0);
        }
        else
        {
            writer.WriteByte(0xFF);
            writer.WriteBytes(PrintSetup.Record);
            writer.WritePascalString(PrintSetup.Header ?? string.Empty);
            writer.WritePascalString(PrintSetup.Footer ?? string.Empty);
            writer.WritePascalString(PrintSetup.PrinterMessage ?? string.Empty);
        }

        for (int i = 0; i < sections.Length; i++)
        {
            writer.PatchUInt32(sectionTable + 4 * i, (uint)(((sections[i] / 512) << 16) | (sections[i] % 512)));
        }

        return writer.ToArray();
    }

    private static void WriteChunks<T>(BigEndianWriter writer, List<T> values, Action<BigEndianWriter, T> write)
    {
        for (int i = 0; i < values.Count; i += 5)
        {
            int count = Math.Min(5, values.Count - i);
            writer.WriteByte((byte)count);
            for (int j = 0; j < count; j++)
            {
                write(writer, values[i + j]);
            }
        }

        writer.WriteByte(0);
    }
}
