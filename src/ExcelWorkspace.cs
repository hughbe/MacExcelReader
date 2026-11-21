using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// One document window in a workspace file: which document, where it lives and how its window looked.
/// </summary>
public sealed class ExcelWorkspaceWindow
{
    /// <summary>Window kind: a worksheet.</summary>
    public const byte WorksheetKind = 0x10;

    /// <summary>Window kind: a chart. A chart window keeps only its rectangle.</summary>
    public const byte ChartKind = 0x20;

    /// <summary>Window kind: a macro sheet.</summary>
    public const byte MacroSheetKind = 0x40;

    /// <summary>Gets or sets the kind of document: <see cref="WorksheetKind"/>, <see cref="ChartKind"/> or <see cref="MacroSheetKind"/>.</summary>
    public byte Kind { get; set; }

    /// <summary>Gets or sets the document's file name.</summary>
    public string DocumentName { get; set; } = string.Empty;

    /// <summary>Gets or sets the folder holding it: a volume name, or a path of folder names, each ending in ':'.</summary>
    public string Folder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the low word of the directory ID of the document's folder, which Excel writes only on a system
    /// with HFS, and 0 otherwise: 2 (a volume's root directory) for a document at the top of a floppy (spec
    /// Appendix B).
    /// </summary>
    public ushort Directory { get; set; }

    /// <summary>
    /// Gets or sets the window: rectangle, 41-byte state and selection, as in a document's section 6. For a
    /// chart only the rectangle is stored, and the state is all zero.
    /// </summary>
    public ExcelWindow Window { get; set; } = new(0, 0, 0, 0, new byte[ExcelWindow.StateSize]);

    /// <inheritdoc/>
    public override string ToString() => Folder + DocumentName;
}

/// <summary>
/// A workspace file (Finder type <c>XLIN</c>, the "Resume Excel" document): the calculation settings and the
/// open document windows, front to back, that Excel restores when it is opened (spec Appendix B).
/// </summary>
public sealed class ExcelWorkspace
{
    private const int SettingsSize = 20;

    /// <summary>Gets the 20 bytes of calculation settings, laid out as offsets 0x000–0x013 of a sheet's state block (spec §6).</summary>
    public byte[] Settings { get; } = new byte[SettingsSize];

    /// <summary>Gets or sets the maximum number of iterations.</summary>
    public ushort MaxIterations
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(Settings);
        set => BinaryPrimitives.WriteUInt16BigEndian(Settings, value);
    }

    /// <summary>Gets or sets the calculation mode word (1 = automatic; spec §6).</summary>
    public short CalculationMode
    {
        get => BinaryPrimitives.ReadInt16BigEndian(Settings.AsSpan(2));
        set => BinaryPrimitives.WriteInt16BigEndian(Settings.AsSpan(2), value);
    }

    /// <summary>Gets or sets whether calculation uses full precision.</summary>
    public bool FullPrecision
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(Settings.AsSpan(4)) != 0;
        set => BinaryPrimitives.WriteUInt16BigEndian(Settings.AsSpan(4), (ushort)(value ? 1 : 0));
    }

    /// <summary>Gets or sets the maximum change for iteration.</summary>
    public double MaxChange
    {
        get => BinaryPrimitives.ReadDoubleBigEndian(Settings.AsSpan(10));
        set => BinaryPrimitives.WriteDoubleBigEndian(Settings.AsSpan(10), value);
    }

    /// <summary>Gets or sets whether iteration is on.</summary>
    public bool Iteration
    {
        get => BinaryPrimitives.ReadUInt16BigEndian(Settings.AsSpan(18)) != 0;
        set => BinaryPrimitives.WriteUInt16BigEndian(Settings.AsSpan(18), (ushort)(value ? 1 : 0));
    }

    /// <summary>
    /// Gets or sets whether a window was zoomed to fill the screen (the FULL command, or a click in the zoom box).
    /// Resume zooms the front window again. The rectangles are always the windows' normal size.
    /// </summary>
    public bool Zoomed { get; set; }

    /// <summary>Gets the document windows, front to back.</summary>
    public List<ExcelWorkspaceWindow> Windows { get; } = [];

    /// <summary>
    /// Gets or sets the preferred chart format (Chart > Set Preferred Format, applied by Gallery > Preferred): Excel's
    /// in-memory image of a chart, whose first word is its length. Empty when none has been set.
    /// </summary>
    public byte[] PreferredChartFormat { get; set; } = [];

    /// <summary>Gets or sets the printer type from Printer Setup: 0 = Macintosh (an ImageWriter), 1 = TTY.</summary>
    public ushort PrinterType { get; set; }

    /// <summary>Gets or sets the TTY printer's baud rate from Printer Setup: 0–5 for 300, 600, 1200, 2400, 4800, 9600.</summary>
    public ushort PrinterBaud { get; set; }

    /// <summary>Gets or sets the TTY printer's port from Printer Setup: 0 = printer port, 1 = modem port.</summary>
    public ushort PrinterPort { get; set; }

    /// <summary>Reads a workspace file (its data fork).</summary>
    /// <exception cref="ExcelFormatException">The data is not a workspace file.</exception>
    public static ExcelWorkspace Read(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        try
        {
            var reader = new BigEndianReader(data);
            if (reader.ReadUInt16() != 0)
            {
                throw new ExcelFormatException("Not an Excel workspace file: the first word must be 0.");
            }

            var workspace = new ExcelWorkspace();
            reader.ReadBytes(SettingsSize).CopyTo(workspace.Settings, 0);
            workspace.Zoomed = reader.ReadUInt16() != 0;
            for (byte kind; (kind = reader.ReadByte()) != 0;)
            {
                var entry = new ExcelWorkspaceWindow
                {
                    Kind = kind,
                    DocumentName = reader.ReadPascalString(),
                    Folder = reader.ReadPascalString(),
                    Directory = reader.ReadUInt16(),
                };
                if (kind == ExcelWorkspaceWindow.ChartKind)
                {
                    entry.Window = new ExcelWindow(reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), reader.ReadInt16(), new byte[ExcelWindow.StateSize]);
                }
                else
                {
                    entry.Window = ExcelWindow.Read(reader, ExcelWindow.StateSize);
                }

                workspace.Windows.Add(entry);
            }

            workspace.PreferredChartFormat = reader.ReadBytes(reader.ReadUInt16());
            reader.ReadByte();
            workspace.PrinterType = reader.ReadUInt16();
            workspace.PrinterBaud = reader.ReadUInt16();
            workspace.PrinterPort = reader.ReadUInt16();
            reader.ReadByte();
            return workspace;
        }
        catch (Exception exception) when (exception is ArgumentOutOfRangeException or IndexOutOfRangeException or EndOfStreamException)
        {
            throw new ExcelFormatException("The workspace file is truncated.", exception);
        }
    }

    /// <summary>Loads a workspace file.</summary>
    public static ExcelWorkspace Load(string path) => Read(File.ReadAllBytes(path));

    /// <summary>Serialises the workspace.</summary>
    public byte[] ToArray()
    {
        var writer = new BigEndianWriter(256);
        writer.WriteUInt16(0);
        writer.WriteBytes(Settings);
        writer.WriteUInt16((ushort)(Zoomed ? 1 : 0));
        foreach (var entry in Windows)
        {
            if (entry.Kind == 0)
            {
                throw new InvalidOperationException("A window's kind cannot be 0, which ends the list.");
            }

            writer.WriteByte(entry.Kind);
            writer.WritePascalString(entry.DocumentName);
            writer.WritePascalString(entry.Folder);
            writer.WriteUInt16(entry.Directory);
            if (entry.Kind == ExcelWorkspaceWindow.ChartKind)
            {
                writer.WriteInt16(entry.Window.Top);
                writer.WriteInt16(entry.Window.Left);
                writer.WriteInt16(entry.Window.Bottom);
                writer.WriteInt16(entry.Window.Right);
            }
            else
            {
                entry.Window.WriteBody(writer);
            }
        }

        writer.WriteByte(0);
        writer.WriteUInt16(checked((ushort)PreferredChartFormat.Length));
        writer.WriteBytes(PreferredChartFormat);
        writer.WriteByte(0);
        writer.WriteUInt16(PrinterType);
        writer.WriteUInt16(PrinterBaud);
        writer.WriteUInt16(PrinterPort);
        writer.WriteByte(0);
        return writer.ToArray();
    }
}
