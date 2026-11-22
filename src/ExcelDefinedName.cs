using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A defined name (spec §11).
/// </summary>
public sealed class ExcelDefinedName
{
    /// <summary>
    /// Flag: the definition is a reference: a cell, an area, or a union of them. Excel sets it on every such name,
    /// on worksheets and macro sheets alike, so a name used as a macro has it too (spec §11).
    /// <see cref="SetDefinition(string, ExcelDocument?, ExcelVersion)"/> sets it.
    /// </summary>
    public const byte ReferenceFlag = 0x04;

    /// <summary>Flag: the definition is array-valued.</summary>
    public const byte ArrayFlag = 0x20;

    /// <summary>
    /// Flag: a macro, function or command. Two bytes follow the flags: the shortcut key and the macro kind
    /// (<see cref="MacroKind"/>). Excel 1.03's DEFINE.NAME sets it for both kinds (spec §11).
    /// </summary>
    public const byte MacroFlag = 0x40;

    /// <summary>Macro kind bit: a command macro, listed in Macro > Run.</summary>
    public const byte CommandKind = 0x40;

    /// <summary>Macro kind bit: a function macro, called from formulas.</summary>
    public const byte FunctionKind = 0x80;

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

    /// <summary>Gets or sets the command macro's shortcut key, typed with Option-Command; 0 for none (used when <see cref="MacroFlag"/> is set).</summary>
    public byte ShortcutKey { get; set; }

    /// <summary>
    /// Gets or sets the macro-kind byte that follows the shortcut key: <see cref="CommandKind"/> or
    /// <see cref="FunctionKind"/>. Excel's own files have bytes left over from memory in the low nibble, which Open discards.
    /// </summary>
    public byte MacroKind { get; set; }

    /// <summary>Gets whether this is a command macro.</summary>
    public bool IsCommand => (Flags & MacroFlag) != 0 && (MacroKind & FunctionKind) == 0;

    /// <summary>Gets whether this is a function macro defined with its kind (Excel's DEFINE.NAME with type 1).</summary>
    public bool IsFunctionMacro => (Flags & MacroFlag) != 0 && (MacroKind & FunctionKind) != 0;

    /// <summary>
    /// Gets the definition as A1-style text, without '=', with its relative references as Excel shows them while
    /// A1 is the active cell.
    /// </summary>
    public string GetDefinitionText(ExcelDocument? document = null)
        => GetDefinitionText(document, default);

    /// <summary>
    /// Gets the definition as A1-style text, without '=', with its relative references as Excel shows them while
    /// <paramref name="relativeTo"/> is the active cell. A definition stores a relative row or column as an offset
    /// from the cell the name is used in (spec §11): a name defined as <c>=B1</c> with C2 active refers to the cell
    /// one row up and one column left of each cell that uses it.
    /// </summary>
    public string GetDefinitionText(ExcelDocument? document, ExcelCellAddress relativeTo)
        => FormulaDecompiler.Decompile(Definition, [], document, fixupsInPlace: true, _arrayConstants, relativeTo);

    /// <summary>
    /// Gets the data of the array constants (<c>0x20</c> tokens) in the definition, one block per token in order
    /// (spec §12.5). They are stored after the list of names (spec §11).
    /// </summary>
    public List<byte[]> ArrayConstants => _arrayConstants ??= [];

    internal bool HasArrayConstants => _arrayConstants is { Count: > 0 };

    // Created on first use: most names have no array constants.
    private List<byte[]>? _arrayConstants;

    /// <summary>
    /// Sets the definition from formula text, such as <c>=$A$1:$B$5</c> or <c>={1,2,3}</c>, and sets or clears
    /// <see cref="ReferenceFlag"/> to match it.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public void SetDefinition(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
        => SetDefinition(text, default, document, version);

    /// <summary>
    /// Sets the definition from formula text whose relative references are relative to
    /// <paramref name="relativeTo"/>, as when Excel defines a name with that cell active: <c>=B1</c> with C2 as
    /// <paramref name="relativeTo"/> refers to the cell one row up and one column left of each cell that uses the
    /// name (spec §11). Sets or clears <see cref="ReferenceFlag"/> to match the definition.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public void SetDefinition(string text, ExcelCellAddress relativeTo, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
    {
        var (definition, arrays) = ExcelFormulaParser.ParseNameDefinition(text, document, version);
        Rebase(definition, relativeTo);
        Definition = definition;
        ArrayConstants.Clear();
        ArrayConstants.AddRange(arrays);
        Flags = IsReference(definition) ? (byte)(Flags | ReferenceFlag) : (byte)(Flags & ~ReferenceFlag);
    }

    // Turns the relative parts of the definition's references, parsed as cell indices, into offsets from relativeTo.
    private static void Rebase(byte[] definition, ExcelCellAddress relativeTo)
    {
        if (relativeTo == default)
        {
            return;
        }

        foreach (var token in FormulaTokens.Enumerate(definition))
        {
            int p = token.Offset + 1;
            switch (token.Code)
            {
                case 0x2C:
                    RebaseCell(definition, p, p + 2, relativeTo);
                    break;
                case 0x2D:
                    RebaseCell(definition, p, p + 4, relativeTo);
                    RebaseCell(definition, p + 2, p + 5, relativeTo);
                    break;
            }
        }
    }

    private static void RebaseCell(byte[] definition, int rowOffset, int columnOffset, ExcelCellAddress relativeTo)
    {
        var (rowWord, column) = RowWord.Move(
            BinaryPrimitives.ReadUInt16BigEndian(definition.AsSpan(rowOffset)), definition[columnOffset], -relativeTo.Row, -relativeTo.Column);
        BinaryPrimitives.WriteUInt16BigEndian(definition.AsSpan(rowOffset), rowWord);
        definition[columnOffset] = column;
    }

    // A single cell or area token (2C, 2D in any class), or a union in its definition form (2E n ...).
    private static bool IsReference(ReadOnlySpan<byte> definition)
    {
        if (definition.IsEmpty)
        {
            return false;
        }

        int code = definition[0] < 0x20 ? definition[0] : (definition[0] & 0x1F) | 0x20;
        return code switch
        {
            0x2C => definition.Length == 4,
            0x2D => definition.Length == 7,
            0x2E => true,
            _ => false,
        };
    }

    internal static ExcelDefinedName Read(BigEndianReader reader, byte nameLength)
    {
        int definitionLength = reader.ReadByte();
        byte flags = reader.ReadByte();
        byte key = 0, extra = 0;
        if ((flags & MacroFlag) != 0)
        {
            key = reader.ReadByte();
            extra = reader.ReadByte();
        }

        var definition = reader.ReadBytes(definitionLength);
        if (reader.ReadByte() != definitionLength)
        {
            throw new ExcelFormatException("A defined name's two definition lengths disagree.");
        }

        var name = MacRoman.GetString(reader.ReadSpan(nameLength));
        return new ExcelDefinedName(name, definition, flags) { ShortcutKey = key, MacroKind = extra };
    }

    internal void Write(BigEndianWriter writer)
    {
        // Mac OS Roman has one byte per character.
        if (Name.Length is 0 or > 255 || Definition.Length > 255)
        {
            throw new InvalidOperationException($"Name '{Name}' or its definition is too long.");
        }

        writer.WriteByte((byte)Name.Length);
        writer.WriteByte((byte)Definition.Length);
        writer.WriteByte(Flags);
        if ((Flags & MacroFlag) != 0)
        {
            writer.WriteByte(ShortcutKey);
            writer.WriteByte(MacroKind);
        }

        writer.WriteBytes(Definition);
        writer.WriteByte((byte)Definition.Length);
        writer.WriteMacRoman(Name);
    }

    /// <inheritdoc/>
    public override string ToString() => Name;
}
