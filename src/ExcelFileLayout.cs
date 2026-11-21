using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The sections of a worksheet or macro sheet file.
/// </summary>
public enum ExcelSection
{
    /// <summary>Section 0: number formats.</summary>
    NumberFormats = 0,

    /// <summary>Section 1: external documents and the row index.</summary>
    ExternalDocuments = 1,

    /// <summary>Section 2: the cell table.</summary>
    CellTable = 2,

    /// <summary>Section 3: names used from external documents.</summary>
    ExternalNames = 3,

    /// <summary>Section 4: defined names.</summary>
    DefinedNames = 4,

    /// <summary>Section 5: cell data (numbers, text and formula records).</summary>
    CellData = 5,

    /// <summary>Section 6: windows.</summary>
    Windows = 6,

    /// <summary>Section 7: page breaks and print setup.</summary>
    PrintSettings = 7,
}

/// <summary>
/// Where each part of a worksheet or macro sheet file lies (spec §2), read without parsing the contents.
/// Useful for inspecting and debugging files.
/// </summary>
public sealed class ExcelFileLayout
{
    private readonly int[] _sections;

    private ExcelFileLayout(int fileLength, IReadOnlyList<(int Offset, int Length)> codeSegments, int dataOffset, int extraHeaderOffset, int extraHeaderLength, int[] sections)
    {
        FileLength = fileLength;
        CodeSegments = codeSegments;
        DataOffset = dataOffset;
        ExtraHeaderOffset = extraHeaderOffset;
        ExtraHeaderLength = extraHeaderLength;
        _sections = sections;
    }

    /// <summary>Gets the file length.</summary>
    public int FileLength { get; }

    /// <summary>Gets the offset and length of each embedded code segment.</summary>
    public IReadOnlyList<(int Offset, int Length)> CodeSegments { get; }

    /// <summary>Gets the offset of the data area (the function-count word).</summary>
    public int DataOffset { get; }

    /// <summary>Gets the offset of the 630-byte sheet state block.</summary>
    public int StateOffset => DataOffset + 2;

    /// <summary>Gets the offset of the document-kind byte.</summary>
    public int KindOffset => StateOffset + ExcelSheetState.Size;

    /// <summary>Gets the offset of the section table.</summary>
    public int SectionTableOffset => KindOffset + 1;

    /// <summary>Gets the offset of the extra header bytes.</summary>
    public int ExtraHeaderOffset { get; }

    /// <summary>Gets the number of extra header bytes.</summary>
    public int ExtraHeaderLength { get; }

    /// <summary>Gets the file offset where a section starts.</summary>
    public int GetSectionOffset(ExcelSection section) => _sections[(int)section];

    /// <summary>Gets a section's length: the distance to the next section, or to the end of the file for the last.</summary>
    public int GetSectionLength(ExcelSection section)
    {
        int index = (int)section;
        int end = index + 1 < _sections.Length ? _sections[index + 1] : FileLength;
        return Math.Max(0, end - _sections[index]);
    }

    /// <summary>
    /// Reads the layout of a file.
    /// </summary>
    /// <exception cref="ExcelFormatException">The file is not a worksheet or macro sheet.</exception>
    public static ExcelFileLayout Read(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        var reader = new BigEndianReader(data);
        if (data.Length < 24 || reader.ReadUInt16() != ExcelDocument.Magic)
        {
            throw new ExcelFormatException("Not an Excel 1.x worksheet or macro sheet (bad magic number).");
        }

        reader.Position = 12;
        int count = reader.ReadUInt16();
        var segments = new List<(int, int)>(count);
        for (int i = 0; i < count; i++)
        {
            int length = reader.ReadUInt16();
            segments.Add((reader.ReadUInt16(), length));
        }

        int dataOffset = reader.ReadUInt16();
        reader.Position = dataOffset + 2 + ExcelSheetState.Size + 1;
        var sections = new int[8];
        for (int i = 0; i < sections.Length; i++)
        {
            uint position = reader.ReadUInt32();
            sections[i] = (int)((position >> 16) * 512 + (position & 0xFFFF));
        }

        int extraLength = reader.ReadUInt16();
        return new ExcelFileLayout(data.Length, segments, dataOffset, reader.Position, extraLength, sections);
    }
}
