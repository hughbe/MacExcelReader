using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A defined name (spec §11).
/// </summary>
public sealed class ExcelDefinedName
{
    /// <summary>Flag: defined on a macro sheet as a function or command name.</summary>
    public const byte MacroFlag = 0x04;

    /// <summary>Flag: the definition is array-valued.</summary>
    public const byte ArrayFlag = 0x20;

    /// <summary>Flag: a command macro; a shortcut key follows the flags.</summary>
    public const byte CommandFlag = 0x40;

    /// <summary>
    /// Creates a name.
    /// </summary>
    /// <param name="name">The name, 1–255 characters.</param>
    /// <param name="definition">The definition as a parsed expression; empty for an undefined name.</param>
    /// <param name="flags">The flags byte.</param>
    public ExcelDefinedName(string name, byte[]? definition = null, byte flags = 0)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        Name = name;
        Definition = definition ?? [];
        Flags = flags;
    }

    /// <summary>Gets or sets the name.</summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the definition, a parsed expression whose name and external-document indices are
    /// stored inside the tokens.
    /// </summary>
    public byte[] Definition { get; set; }

    /// <summary>Gets or sets the flags byte.</summary>
    public byte Flags { get; set; }

    /// <summary>Gets or sets the command-macro shortcut key (used when <see cref="CommandFlag"/> is set).</summary>
    public byte ShortcutKey { get; set; }

    /// <summary>Gets or sets the unidentified byte that follows the shortcut key.</summary>
    public byte ShortcutExtra { get; set; }

    /// <summary>Gets whether this is a command macro.</summary>
    public bool IsCommand => (Flags & CommandFlag) != 0;

    /// <summary>
    /// Gets the definition as A1-style text, without '='.
    /// </summary>
    public string GetDefinitionText(ExcelDocument? document = null)
        => FormulaDecompiler.Decompile(Definition, [], document, fixupsInPlace: true);

    internal static ExcelDefinedName Read(BigEndianReader reader, byte nameLength)
    {
        int definitionLength = reader.ReadByte();
        byte flags = reader.ReadByte();
        byte key = 0, extra = 0;
        if ((flags & CommandFlag) != 0)
        {
            key = reader.ReadByte();
            extra = reader.ReadByte();
        }

        var definition = reader.ReadBytes(definitionLength);
        if (reader.ReadByte() != definitionLength)
        {
            throw new ExcelFormatException("A defined name's two definition lengths disagree.");
        }

        var name = MacRoman.GetString(reader.ReadBytes(nameLength));
        return new ExcelDefinedName(name, definition, flags) { ShortcutKey = key, ShortcutExtra = extra };
    }

    internal void Write(BigEndianWriter writer)
    {
        var nameBytes = MacRoman.GetBytes(Name);
        if (nameBytes.Length is 0 or > 255 || Definition.Length > 255)
        {
            throw new InvalidOperationException($"Name '{Name}' or its definition is too long.");
        }

        writer.WriteByte((byte)nameBytes.Length);
        writer.WriteByte((byte)Definition.Length);
        writer.WriteByte(Flags);
        if ((Flags & CommandFlag) != 0)
        {
            writer.WriteByte(ShortcutKey);
            writer.WriteByte(ShortcutExtra);
        }

        writer.WriteBytes(Definition);
        writer.WriteByte((byte)Definition.Length);
        writer.WriteBytes(nameBytes);
    }

    /// <inheritdoc/>
    public override string ToString() => Name;
}
