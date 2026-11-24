using System.Buffers;
using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// An Excel 1.x for Macintosh worksheet or macro sheet.
/// </summary>
public sealed class ExcelDocument
{
    /// <summary>The magic number at the start of every worksheet and macro sheet.</summary>
    public const ushort Magic = 0x5327;

    /// <summary>
    /// Gets or sets the format version: 17, readable by every Excel 1.x, or 18, written by Excel 1.5 when the
    /// document uses something only Excel 1.5 can read (spec §2.1).
    /// </summary>
    public ExcelFormatVersion Version
    {
        get => _version;
        set => _version = Enum.IsDefined(value) ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The format version is 17 or 18.");
    }

    private ExcelFormatVersion _version = ExcelFormatVersion.Version17;

    /// <summary>The code signature every Excel 1.x release writes in header word 4 (spec §4).</summary>
    public const ushort CodeSignature = 0xAB27;

    private int _sourceLength;

    private ExcelDocument(ExcelSheetState state)
    {
        State = state;
    }

    /// <summary>Gets or sets the code signature of the build that wrote the file (header word 4).</summary>
    public ushort Signature { get; set; } = CodeSignature;

    /// <summary>Gets the six zero bytes that follow the signature.</summary>
    public byte[] HeaderReserved { get; private set; } = new byte[6];

    /// <summary>
    /// Gets the two embedded p-code segments (spec §4.1). They are only used when another build of Excel links
    /// to this document while it is closed. Documents created with this library have empty segments. Excel
    /// reads the data offset from a fixed position in the 24-byte header, so there must be exactly two.
    /// </summary>
    public List<byte[]> CodeSegments { get; } = [];

    /// <summary>Gets or sets the function count of the writing build (208 for Excel 1.x).</summary>
    public ushort FunctionCount { get; set; } = (ushort)ExcelFunctions.Count;

    /// <summary>
    /// Gets the names of the functions numbered from 208 up to <see cref="FunctionCount"/>, which a file without an
    /// extra header carries at the start of section 7 when its count is above 208 (spec §14), so that Open can
    /// check they mean the same functions. Excel 1.00's code supports the list for a build whose function table
    /// is longer than its start-up count, but released 1.00 has 208 of each: it never writes a name and refuses
    /// any file whose count is above 208. Excel 1.01 and later dropped the list, so no known file has one.
    /// </summary>
    public List<string> ExtraFunctionNames { get; } = [];

    /// <summary>Gets the sheet state block.</summary>
    public ExcelSheetState State { get; }

    /// <summary>Gets or sets whether this is a worksheet or a macro sheet. Charts are <see cref="Charts.ChartDocument"/>s.</summary>
    public ExcelDocumentKind Kind
    {
        get => _kind;
        set => _kind = value is ExcelDocumentKind.Worksheet or ExcelDocumentKind.MacroSheet
            ? value
            : throw new ArgumentOutOfRangeException(nameof(value), value, "A document is a worksheet or a macro sheet; charts are ChartDocuments.");
    }

    private ExcelDocumentKind _kind = ExcelDocumentKind.Worksheet;

    /// <summary>Gets or sets the extra header bytes (byte 0 is the system code).</summary>
    public byte[] ExtraHeader { get; set; } = [0];

    /// <summary>Gets or sets the number of built-in number formats of the writing build (19).</summary>
    public byte BuiltInFormatCount { get; set; } = 19;

    /// <summary>Gets the user-defined number formats, numbered from 19 (spec §7).</summary>
    public List<ExcelNumberFormat> NumberFormats { get; } = [];

    /// <summary>
    /// Compiles format text with the given release's rules (see <see cref="ExcelNumberFormat.Parse(string, ExcelVersion)"/>),
    /// adds it to <see cref="NumberFormats"/> in the form this document's <see cref="Version"/> stores, and returns the
    /// index cells use for it. A format only Excel 1.5 can read needs <see cref="Version"/> 18.
    /// </summary>
    /// <exception cref="InvalidOperationException">The format needs version 18 and the document is version 17, or the
    /// document already has the most formats Excel allows.</exception>
    public int AddNumberFormat(string text, ExcelVersion version = ExcelVersion.Excel100)
    {
        var format = ExcelNumberFormat.Parse(text, version, Version >= ExcelFormatVersion.Version18);
        if (Version < ExcelFormatVersion.Version18 && ExcelNumberFormat.Parse(text, version).RequiresVersion18)
        {
            throw new InvalidOperationException($"\"{text}\" can only be read by Excel 1.5: set Version to {ExcelFormatVersion.Version18} first.");
        }

        if (ExcelNumberFormat.BuiltIn.Count + NumberFormats.Count >= 64)
        {
            throw new InvalidOperationException("A document has at most 64 number formats, built-in ones included.");
        }

        NumberFormats.Add(format);
        return ExcelNumberFormat.BuiltIn.Count + NumberFormats.Count - 1;
    }

    /// <summary>
    /// Gets the number format a cell's format index refers to: one of the built-in formats, or one of
    /// <see cref="NumberFormats"/> from 19 on. Returns null for an index past the last format.
    /// </summary>
    public ExcelNumberFormat? GetNumberFormat(int index)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(index);
        if (index < ExcelNumberFormat.BuiltIn.Count)
        {
            return ExcelNumberFormat.BuiltIn[index];
        }

        index -= ExcelNumberFormat.BuiltIn.Count;
        return index < NumberFormats.Count ? NumberFormats[index] : null;
    }

    /// <summary>
    /// Gets or sets whether the external-document table is written. It is written whenever
    /// <see cref="ExternalDocuments"/> is non-empty.
    /// </summary>
    public bool HasExternalDocumentTable { get; set; }

    /// <summary>Gets the documents this sheet refers to.</summary>
    public List<ExcelExternalDocument> ExternalDocuments { get; } = [];


    /// <summary>
    /// Gets the per-row size hints of the recalculation dependency index (spec §8). Excel rebuilds the
    /// index on open; zeros are always valid.
    /// </summary>
    public List<ushort> RowIndex => _rowIndex ??= ReadRowIndex();

    // Read from the file when first used: Excel ignores it, so most callers never do.
    private List<ushort>? _rowIndex;
    private byte[]? _rowIndexSource;
    private int _rowIndexOffset;
    private int _rowIndexCount;

    private List<ushort> ReadRowIndex()
    {
        var index = new List<ushort>(_rowIndexCount);
        for (int r = 0; r < _rowIndexCount; r++)
        {
            index.Add(BinaryPrimitives.ReadUInt16BigEndian(_rowIndexSource.AsSpan(_rowIndexOffset + 2 * r)));
        }

        return index;
    }

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
    public static ExcelDocument Load(string path) => ReadOwned(File.ReadAllBytes(path));

    /// <summary>
    /// Loads a document from a file (its data fork) asynchronously.
    /// </summary>
    public static async Task<ExcelDocument> LoadAsync(string path, CancellationToken cancellationToken = default)
        => ReadOwned(await File.ReadAllBytesAsync(path, cancellationToken).ConfigureAwait(false));

    /// <summary>
    /// Reads a document from a stream asynchronously, from its current position to its end.
    /// </summary>
    public static async Task<ExcelDocument> ReadAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var copy = stream.CanSeek ? new MemoryStream(checked((int)(stream.Length - stream.Position))) : new MemoryStream();
        await stream.CopyToAsync(copy, cancellationToken).ConfigureAwait(false);
        return ReadOwned(copy.Length == copy.Capacity ? copy.GetBuffer() : copy.ToArray());
    }

    /// <summary>
    /// Reads a document from a stream.
    /// </summary>
    public static ExcelDocument Read(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (stream.CanSeek)
        {
            var buffer = new byte[stream.Length - stream.Position];
            stream.ReadExactly(buffer);
            return ReadOwned(buffer);
        }

        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return ReadOwned(copy.ToArray());
    }

    /// <summary>
    /// Reads a document from its bytes (the data fork of the Macintosh file). The bytes are copied, so
    /// the caller may reuse the array.
    /// </summary>
    /// <remarks>
    /// Text and formula cells are decoded on first access, so a malformed record surfaces as an
    /// <see cref="ExcelFormatException"/> from the cell rather than from this method.
    /// </remarks>
    public static ExcelDocument Read(byte[] data) => Read(data, copyData: true);

    /// <summary>
    /// Reads a document from its bytes, optionally without copying them.
    /// </summary>
    /// <param name="data">The file's data fork.</param>
    /// <param name="copyData">
    /// True to copy the bytes first. Pass false to avoid the copy when the caller hands the array over:
    /// the document reads cells from it lazily, so it must not be modified afterwards.
    /// </param>
    public static ExcelDocument Read(byte[] data, bool copyData)
    {
        ArgumentNullException.ThrowIfNull(data);
        return ReadOwned(copyData ? data.AsSpan().ToArray() : data);
    }

    /// <summary>
    /// Reads a document from bytes. The bytes are copied.
    /// </summary>
    public static ExcelDocument Read(ReadOnlySpan<byte> data) => ReadOwned(data.ToArray());

    private static ExcelDocument ReadOwned(byte[] data)
    {
        var document = ReadHeader(data);
        document._sourceLength = data.Length;
        return document;
    }

    private static ExcelDocument ReadHeader(byte[] data)
    {
        var reader = new BigEndianReader(data);
        if (data.Length < 24 || reader.ReadUInt16() != Magic)
        {
            throw new ExcelFormatException("Not an Excel 1.x worksheet or macro sheet (bad magic number).");
        }

        ushort version = reader.ReadUInt16();
        if (version is not ((ushort)ExcelFormatVersion.Version17 or (ushort)ExcelFormatVersion.Version18))
        {
            throw new ExcelFormatException($"Unsupported format version {version}; expected 17 or 18.");
        }

        ushort signature = reader.ReadUInt16();
        var headerReserved = reader.ReadBytes(6);
        // Excel's Open reads neither the segment count nor the segment table (spec §4), and opens a file with junk in
        // them; it uses the segments only to link to a closed document. Keep the two segments the table describes when
        // it describes them, and two empty ones otherwise.
        reader.ReadUInt16();
        var segmentInfo = new (int Length, int Offset)[2];
        for (int i = 0; i < segmentInfo.Length; i++)
        {
            segmentInfo[i] = (reader.ReadUInt16(), reader.ReadUInt16());
        }

        // Like Excel's Open, take the data offset from header word 22.
        int dataOffset = reader.ReadUInt16();
        reader.Position = dataOffset;
        ushort functionCount = reader.ReadUInt16();
        var document = new ExcelDocument(new ExcelSheetState(reader.ReadBytes(ExcelSheetState.Size)))
        {
            Signature = signature,
            Version = (ExcelFormatVersion)version,
            HeaderReserved = headerReserved,
            FunctionCount = functionCount,
            Kind = reader.ReadByte() switch
            {
                (byte)ExcelDocumentKind.Worksheet => ExcelDocumentKind.Worksheet,
                (byte)ExcelDocumentKind.MacroSheet => ExcelDocumentKind.MacroSheet,
                var kind => throw new ExcelFormatException($"Unknown document kind 0x{kind:X2}; expected 0x10 (worksheet) or 0x40 (macro sheet)."),
            },
        };

        bool segmentsInFile = segmentInfo.All(s => s.Offset >= 0 && s.Offset + s.Length <= dataOffset);
        foreach (var (length, offset) in segmentInfo)
        {
            document.CodeSegments.Add(segmentsInFile ? data.AsSpan(offset, length).ToArray() : []);
        }

        var table = new int[8];
        for (int i = 0; i < table.Length; i++)
        {
            uint position = reader.ReadUInt32();
            table[i] = (int)((position >> 16) * 512 + (position & 0xFFFF));
        }

        document.ExtraHeader = reader.ReadBytes(reader.ReadUInt16());
        int rows = document.State.RowCount;
        int columns = document.State.ColumnCountInTable;

        // Like Excel's Open, read the sections one after another rather than trusting the section table (spec §5):
        // a file with junk in the table opens in Excel. Only the end of the cell data needs the table, or failing
        // that a pass over the cell table.
        document.ReadNumberFormats(reader);
        document.ReadExternalDocuments(reader, rows);
        int cellTable = reader.Position;
        reader.Position = checked(cellTable + rows * columns * CellWords.EntrySize);
        if (reader.Position > data.Length)
        {
            throw new ExcelFormatException("The cell table runs past the end of the file.");
        }

        // Section 3: for each external document, the names used from it, ending with a zero byte.
        foreach (var external in document.ExternalDocuments)
        {
            for (byte length; (length = reader.ReadByte()) != 0;)
            {
                external.Names.Add(MacRoman.GetString(reader.ReadSpan(length)));
            }
        }

        document.ReadNames(reader);
        int cellData = reader.Position;
        document.Cells.AttachSource(data, cellTable, cellData, rows, columns);
        int windows = table[6] >= cellData && table[6] <= table[7] && table[7] <= data.Length
            ? table[6]
            : MeasureCellData(data, cellTable, cellData, rows, columns);
        reader.Position = windows;
        document.ReadWindows(reader);
        var printSettings = reader;
        if (document.ExtraHeader.Length == 0)
        {
            for (int i = ExcelFunctions.Count; i < document.FunctionCount; i++)
            {
                document.ExtraFunctionNames.Add(printSettings.ReadPascalString());
            }
        }

        document.ReadPrintSettings(printSettings);
        return document;
    }

    // The end of the cell data: the furthest end of any cell's record (spec §9, §12).
    private static int MeasureCellData(byte[] data, int cellTable, int cellData, int rows, int columns)
    {
        int end = cellData;
        for (int i = 0; i < rows * columns; i++)
        {
            var entry = data.AsSpan(cellTable + i * CellWords.EntrySize);
            ushort word0 = BinaryPrimitives.ReadUInt16BigEndian(entry);
            ushort word1 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(2));
            var type = CellWords.Type(word1);
            if (!ExcelCellGrid.HasCellData(type))
            {
                continue;
            }

            int offset = cellData + (word0 << 4) + (word1 >> 12);
            if (offset >= data.Length)
            {
                throw new ExcelFormatException("A cell's data lies outside the file.");
            }

            end = Math.Max(end, offset + ExcelFormula.MeasureCellData(type, data.AsSpan(offset)));
        }

        return end;
    }

    private void ReadNumberFormats(BigEndianReader reader)
    {
        BuiltInFormatCount = reader.ReadByte();
        int count = reader.ReadByte();
        for (int i = 0; i < count; i++)
        {
            // Version 17: three section offsets, three lengths and a flag. Version 18 (Excel 1.5) adds a fourth
            // section, for text, so four offsets and four lengths (spec §7).
            int sections = Version >= ExcelFormatVersion.Version18 ? 4 : 3;
            var descriptor = reader.ReadBytes(2 * sections + 1);
            int bodyLength = 0;
            for (int s = 0; s < sections; s++)
            {
                bodyLength = Math.Max(bodyLength, descriptor[s] + descriptor[sections + s]);
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

        _rowIndexSource = reader.Data;
        _rowIndexOffset = reader.Position;
        _rowIndexCount = rows;
        reader.Skip(2 * rows);
    }

    private void ReadNames(BigEndianReader reader)
    {
        for (byte length; (length = reader.ReadByte()) != 0;)
        {
            Names.Add(ExcelDefinedName.Read(reader, length));
        }

        // After the terminator: the data of every array constant in the definitions, name by name.
        var data = reader.Data.AsSpan(reader.Position);
        int position = 0;
        foreach (var name in Names)
        {
            var blocks = FormulaTokens.ReadConsecutiveArrayBlocks(name.Definition, data, ref position);
            if (blocks.Count > 0)
            {
                name.ArrayConstants.AddRange(blocks);
            }
        }

        reader.Position += position;
    }

    private void ReadWindows(BigEndianReader reader)
    {
        for (byte marker; (marker = reader.ReadByte()) != 0;)
        {
            if (marker != 0xFF)
            {
                throw new ExcelFormatException($"Unexpected window marker 0x{marker:X2}.");
            }

            Windows.Add(ExcelWindow.Read(reader, Version >= ExcelFormatVersion.Version18 ? ExcelWindow.Version18StateSize : ExcelWindow.StateSize));
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
        var bytes = reader.ReadPascalSpan();
        return bytes.IsEmpty ? null : MacRoman.GetString(bytes);
    }

    /// <summary>
    /// Writes the document to a file (the data fork). Set the Finder type to XLBN for a worksheet or
    /// XLPG for a macro sheet, and the creator to XCEL.
    /// </summary>
    public void Save(string path)
    {
        using var stream = File.Create(path);
        Write(stream);
    }

    /// <summary>
    /// Writes the document to a file (the data fork) asynchronously.
    /// </summary>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = File.Create(path);
        await using (stream.ConfigureAwait(false))
        {
            await WriteAsync(stream, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Writes the document to a stream asynchronously.
    /// </summary>
    public async Task WriteAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        await stream.WriteAsync(Serialize().WrittenMemory, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Protects the document, as Protect Document does: locked cells cannot be changed and hidden formulas are not
    /// shown. A password, if given, is stored as Excel's hash (spec §6.2) and must be given to unprotect it.
    /// </summary>
    public void Protect(string? password = null)
    {
        State.IsProtected = true;
        State.PasswordHash = string.IsNullOrEmpty(password) ? (ushort)0 : ExcelPassword.Hash(password);
    }

    /// <summary>
    /// Unprotects the document, as Unprotect Document does, if <paramref name="password"/> matches the stored hash
    /// (or there is none).
    /// </summary>
    /// <returns>False, leaving the document protected, if the password is wrong.</returns>
    public bool Unprotect(string? password = null)
    {
        ushort hash = string.IsNullOrEmpty(password) ? (ushort)0 : ExcelPassword.Hash(password);
        if (State.PasswordHash != 0 && hash != State.PasswordHash)
        {
            return false;
        }

        State.IsProtected = false;
        State.PasswordHash = 0;
        return true;
    }

    /// <summary>
    /// Writes the document to a stream.
    /// </summary>
    public void Write(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        stream.Write(Serialize().WrittenSpan);
    }

    /// <summary>
    /// Writes the document as SYLK, the text interchange format, byte for byte as Excel 1.00's Save As does with
    /// the SYLK type (spec Appendix C): Mac OS Roman text with carriage returns, formulas in R1C1 style, and
    /// characters outside printable ASCII escaped. Give the file Finder type <c>TEXT</c>.
    /// </summary>
    /// <remarks>Like <see cref="Save(string)"/>, this first sets the grid size from the cells (spec §6.1).</remarks>
    /// <param name="output">Where to write.</param>
    /// <param name="version">
    /// The release whose export to match. Excel 1.00 to 1.06 (the default) write the same; Excel 1.5 writes format
    /// text with every literal escaped, a whole format record for every cell, and the gridline colour.
    /// </param>
    public void ExportSylk(IBufferWriter<byte> output, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(output);
        using var writer = new SylkWriter(this, output, version);
        writer.Write();
    }

    /// <summary>
    /// Writes the document as SYLK to a stream (see <see cref="ExportSylk(IBufferWriter{byte}, ExcelVersion)"/>).
    /// </summary>
    public void ExportSylk(Stream stream, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(stream);
        var buffer = new ArrayBufferWriter<byte>();
        ExportSylk(buffer, version);
        stream.Write(buffer.WrittenSpan);
    }

    /// <summary>
    /// Writes the document as SYLK to a file (see <see cref="ExportSylk(IBufferWriter{byte}, ExcelVersion)"/>).
    /// </summary>
    public void ExportSylk(string path, ExcelVersion version = ExcelVersion.Excel100)
    {
        using var stream = File.Create(path);
        ExportSylk(stream, version);
    }

    /// <summary>
    /// Writes the document as SYLK to a stream asynchronously (see <see cref="ExportSylk(IBufferWriter{byte}, ExcelVersion)"/>).
    /// </summary>
    public async Task ExportSylkAsync(Stream stream, ExcelVersion version = ExcelVersion.Excel100, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        var buffer = new ArrayBufferWriter<byte>();
        ExportSylk(buffer, version);
        await stream.WriteAsync(buffer.WrittenMemory, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Gets the Finder file type for this document.
    /// </summary>
    public string FinderType => Kind == ExcelDocumentKind.MacroSheet ? "XLPG" : "XLBN";

    /// <summary>
    /// Enters an array formula, such as <c>=TREND(B1:B5)</c>, into a range of cells (spec §12.6). The top-left cell
    /// holds the shared formula; the other cells refer to it. Braces around the text are optional.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public void SetArrayFormula(ExcelRange range, string text, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(text);
        range.Validate(nameof(range));

        text = text.Trim();
        if (text.StartsWith('{') && text.EndsWith('}'))
        {
            text = text[1..^1];
        }

        var (expression, fixups) = ExcelFormulaParser.ParseArrayFormula(text, this, version);

        // Anchor trailer: 0x4000 and 0, a 10-byte header (range, three unused bytes, expression length), the
        // shared expression and its fix-ups.
        var anchor = new byte[13 + expression.Length + fixups.Length];
        BinaryPrimitives.WriteUInt16BigEndian(anchor, FormulaTokens.SharedAnchorMarker);
        BinaryPrimitives.WriteUInt16BigEndian(anchor.AsSpan(3), (ushort)range.FirstRow);
        BinaryPrimitives.WriteUInt16BigEndian(anchor.AsSpan(5), (ushort)range.LastRow);
        anchor[7] = (byte)range.FirstColumn;
        anchor[8] = (byte)range.LastColumn;
        anchor[12] = (byte)expression.Length;
        expression.CopyTo(anchor, 13);
        fixups.CopyTo(anchor, 13 + expression.Length);

        for (int row = range.FirstRow; row <= range.LastRow; row++)
        {
            for (int column = range.FirstColumn; column <= range.LastColumn; column++)
            {
                byte[] trailer = row == range.FirstRow && column == range.FirstColumn
                    ? anchor
                    : [(byte)(range.FirstRow >> 8), (byte)range.FirstRow, (byte)range.FirstColumn];
                var formula = new ExcelFormula([0x01, 0, 0, 0], trailer);
                formula.SetResult(0);
                Cells[row, column].SetFormula(formula);
            }
        }
    }

    /// <summary>
    /// Sets the grid-size words (spec §6.1) from the cells, as Save and the SYLK export need them.
    /// </summary>
    internal void UpdateExtent()
    {
        int rows = Cells.ColumnCount == 0 ? 0 : Cells.RowCount;
        int columns = rows == 0 ? 0 : Cells.ColumnCount;

        // Excel uses rows 0 to word 0x014 - 1 and columns 0 to word 0x274 - 1. In the usual last-index form (word
        // 0x020 = 1) that leaves the table's last row and column as a margin: values there are not drawn, cells
        // in the last column load over one another, a one-row table hangs Excel 1.03, and it can drop rows from
        // the end of the sheet. Every sheet Excel saves in that form has no values in its margin, so keep a blank
        // row and column after the last one in use. A table read in the count form is kept as it is.
        if (rows > 0 && !State.HoldsCountForm(rows, columns))
        {
            // Excel keeps the margin even after the sheet's last row or column: a value in row 16384 makes a
            // 16385-row table, and one in column IV a 257-column table.
            bool padRow = rows <= ExcelCellGrid.MaxRows && !IsBlankRow(rows - 1, columns);
            bool padColumn = columns <= ExcelCellGrid.MaxColumns && !IsBlankColumn(columns - 1, rows);
            if (padRow || padColumn)
            {
                rows += padRow ? 1 : 0;
                columns += padColumn ? 1 : 0;
                Cells.EnsureSize(rows, columns);
            }
        }

        State.SetExtent(rows, columns);
    }

    /// <summary>
    /// Serialises the document.
    /// </summary>
    public byte[] ToArray() => Serialize().ToArray();

    private BigEndianWriter Serialize()
    {
        UpdateExtent();
        int rows = State.RowCount;
        int columns = State.ColumnCountInTable;
        State.SetNameCounts(
            Names.Count(n => n.Definition.Length > 0 && (n.Flags & ExcelNameFlags.Reference) != 0),
            ExternalDocuments.FindIndex(d => d.Name == "\0") + 1);

        // A document read from a file usually saves to the same size, so a buffer of that size is returned without
        // a copy.
        var writer = new BigEndianWriter(_sourceLength > 0 ? _sourceLength : 4096);
        writer.WriteUInt16(Magic);
        writer.WriteUInt16((ushort)Version);
        writer.WriteUInt16(Signature);
        writer.WriteBytes(HeaderReserved);
        if (CodeSegments.Count != 2)
        {
            throw new InvalidOperationException("A document must have exactly two code segments (they may be empty); Excel reads the data offset from header offset 22.");
        }

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

        ValidateExtraFunctionNames();
        writer.WriteUInt16(FunctionCount);
        writer.WriteBytes(State.RawData);
        writer.WriteByte((byte)Kind);
        int sectionTable = writer.Position;
        writer.WriteZeros(32);
        writer.WriteUInt16((ushort)ExtraHeader.Length);
        writer.WriteBytes(ExtraHeader);

        var sections = new int[8];

        // Section 0: number formats.
        sections[0] = writer.Position;
        writer.WriteByte(BuiltInFormatCount);
        writer.WriteByte(checked((byte)NumberFormats.Count));
        int formatSections = Version >= ExcelFormatVersion.Version18 ? 4 : 3;
        foreach (var format in NumberFormats)
        {
            if (format.Descriptor is null || format.Descriptor.Length != 2 * formatSections + 1)
            {
                throw new InvalidOperationException($"A number-format descriptor must be {2 * formatSections + 1} bytes in a version-{Version} file.");
            }

            int bodyLength = 0;
            for (int section = 0; section < formatSections; section++)
            {
                bodyLength = Math.Max(bodyLength, format.Descriptor[section] + format.Descriptor[formatSections + section]);
            }

            if (format.Body is null || format.Body.Length != bodyLength)
            {
                throw new InvalidOperationException($"A number-format body must be exactly {bodyLength} bytes for its descriptor, not {format.Body?.Length.ToString() ?? "null"}.");
            }

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

        // Untouched, the file's words are copied; rows beyond them get 0.
        int known = 0;
        if (_rowIndex is null && _rowIndexSource is not null)
        {
            known = Math.Min(rows, _rowIndexCount);
            writer.WriteBytes(_rowIndexSource.AsSpan(_rowIndexOffset, 2 * known));
        }
        else if (_rowIndex is not null)
        {
            known = Math.Min(rows, _rowIndex.Count);
            for (int r = 0; r < known; r++)
            {
                writer.WriteUInt16(_rowIndex[r]);
            }
        }

        writer.WriteZeros(2 * (rows - known));

        // Section 2: cell table. It is reserved here and filled in as the cell data is written.
        sections[2] = writer.Position;
        int tableStart = writer.Position;
        writer.WriteZeros(rows * columns * 6);

        // Section 3: for each external document, the names used from it and a zero byte. With no external
        // documents the section is empty.
        sections[3] = writer.Position;
        foreach (var external in ExternalDocuments)
        {
            foreach (var name in external.Names)
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
        foreach (var name in Names)
        {
            if (name.HasArrayConstants)
            {
                foreach (var block in name.ArrayConstants)
                {
                    writer.WriteBytes(block);
                }
            }
        }

        // Section 5: cell data. Cells that were never accessed are copied byte for byte.
        sections[5] = writer.Position;
        WriteCells(writer, tableStart, rows, columns);

        // Section 6: windows.
        sections[6] = writer.Position;
        foreach (var window in Windows)
        {
            window.Write(writer, Version >= ExcelFormatVersion.Version18);
        }

        writer.WriteByte(0);

        // Section 7: Excel 1.00's list of extra function names, then page breaks and print setup.
        sections[7] = writer.Position;
        foreach (string name in ExtraFunctionNames)
        {
            writer.WritePascalString(name);
        }

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
            // Excel's 512-byte buffer is written out only when the next byte arrives, so a position on a block
            // boundary is the end of the full block before it: (block - 1, 512) (spec §1.2).
            int block = sections[i] / 512, offset = sections[i] % 512;
            if (offset == 0 && block > 0)
            {
                block--;
                offset = 512;
            }

            writer.PatchUInt32(sectionTable + 4 * i, (uint)((block << 16) | offset));
        }

        return writer;
    }

    private bool IsBlankRow(int row, int columns)
    {
        for (int c = 0; c < columns; c++)
        {
            if (!IsBlankEntry(row, c))
            {
                return false;
            }
        }

        return true;
    }

    private bool IsBlankColumn(int column, int rows)
    {
        for (int r = 0; r < rows; r++)
        {
            if (!IsBlankEntry(r, column))
            {
                return false;
            }
        }

        return true;
    }

    // A cell with no value. Excel's own files can have formatting in the margin, but never a value.
    private bool IsBlankEntry(int row, int column) => Cells.GetCellType(row, column) == ExcelCellType.Empty;

    private void WriteCells(BigEndianWriter writer, int tableStart, int rows, int columns)
    {
        int dataStart = writer.Position;
        int entry = tableStart;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++, entry += 6)
            {
                int offset = writer.Position - dataStart;
                if (!Cells.TryGetRawCell(r, c, out ushort word0, out ushort word1, out ushort word2, out var raw))
                {
                    var cell = Cells.Find(r, c)!;
                    word0 = cell.Word0;
                    word1 = cell.Word1;
                    word2 = cell.Word2;
                    switch (cell.Type)
                    {
                        case ExcelCellType.Number:
                            writer.WriteDouble(cell.NumberValue);
                            break;
                        case ExcelCellType.Text:
                            writer.WritePascalString(cell.TextValue);
                            break;
                        case ExcelCellType.NumericFormula:
                        case ExcelCellType.Formula:
                            cell.Formula!.Write(writer, r, c);
                            break;
                    }
                }
                else
                {
                    writer.WriteBytes(raw);
                }

                if (ExcelCellGrid.HasCellData(CellWords.Type(word1)))
                {
                    if (offset > 0xFFFFF)
                    {
                        throw new InvalidOperationException("The cell data is larger than the 1 MB the cell table can address.");
                    }

                    word0 = (ushort)(offset >> 4);
                    word1 = (ushort)((word1 & 0x0FFF) | ((offset & 0xF) << 12));
                }

                writer.PatchUInt16(entry, word0);
                writer.PatchUInt16(entry + 2, word1);
                writer.PatchUInt16(entry + 4, word2);
            }
        }
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

    private void ValidateExtraFunctionNames()
    {
        // Open reads one name per function from 208 up to the count, but only when there is no extra header (spec §14).
        int expected = ExtraHeader.Length == 0 ? Math.Max(0, FunctionCount - ExcelFunctions.Count) : 0;
        if (ExtraFunctionNames.Count != expected)
        {
            throw new InvalidOperationException(ExtraHeader.Length == 0
                ? $"A file without an extra header needs one extra function name per function from {ExcelFunctions.Count} up to the function count ({expected}), not {ExtraFunctionNames.Count}. To save a document using later functions in the form Excel 1.01 and later read, set ExtraHeader to [0]."
                : "Only a file without an extra header can hold extra function names.");
        }
    }
}
