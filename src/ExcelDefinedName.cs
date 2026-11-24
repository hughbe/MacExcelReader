using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The flags of a defined name (spec §11). Excel's files also carry bits 0x01, 0x02 and 0x80 left over from memory,
/// which Open discards; they are kept as they are.
/// </summary>
[Flags]
public enum ExcelNameFlags : byte
{
    /// <summary>No flags.</summary>
    None = 0,

    /// <summary>
    /// The definition is a reference: a cell, an area, or a union of them. Excel sets it on every such name, on
    /// worksheets and macro sheets alike, so a name used as a macro has it too. <see cref="ExcelDefinedName.SetDefinition(string, ExcelDocument?, ExcelVersion)"/>
    /// sets it.
    /// </summary>
    Reference = 0x04,

    /// <summary>The definition is array-valued, so formulas using the name are marked as using an array function.</summary>
    Array = 0x20,

    /// <summary>The name is a macro: a command or a function. A shortcut key and a macro kind follow in the file.</summary>
    Macro = 0x40,
}

/// <summary>
/// What kind of macro a defined name is.
/// </summary>
public enum ExcelMacroKind
{
    /// <summary>Not a macro.</summary>
    None,

    /// <summary>A command macro, listed in Macro &gt; Run and run with its shortcut key.</summary>
    Command,

    /// <summary>A function macro, called from formulas.</summary>
    Function,
}

/// <summary>
/// A defined name (spec §11).
/// </summary>
public sealed class ExcelDefinedName
{
    // Bits of the macro-kind byte (spec §11); its low nibble is left over from memory.
    private const byte CommandBit = 0x40;
    private const byte FunctionBit = 0x80;

    /// <summary>
    /// Creates a name.
    /// </summary>
    /// <param name="name">The name, 1–255 characters.</param>
    /// <param name="definition">The definition as a parsed expression; empty for an undefined name.</param>
    /// <param name="flags">The flags.</param>
    public ExcelDefinedName(string name, ReadOnlySpan<byte> definition = default, ExcelNameFlags flags = ExcelNameFlags.None)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        Name = name;
        Definition = definition.IsEmpty ? default : definition.ToArray();
        Flags = flags;
    }

    /// <summary>Gets or sets the name.</summary>
    public string Name
    {
        get => _name;
        set
        {
            _name = string.IsNullOrEmpty(value) ? throw new ArgumentException("A name cannot be empty.", nameof(value)) : PascalText.Check(value, nameof(value));
            _unchanged = false;
        }
    }

    private string _name = string.Empty;

    // Whether the name and definition are still those read from the file (ExcelDocument sorts changed names on save).
    private bool _unchanged;

    internal void MarkUnchanged() => _unchanged = true;

    internal bool IsUnchanged => _unchanged;

    /// <summary>
    /// Gets or sets the definition, a parsed expression whose name and external-document indices are
    /// stored inside the tokens. A name read from a file refers to the file's bytes rather than copying them.
    /// </summary>
    /// <exception cref="ArgumentException">The definition is longer than 255 bytes.</exception>
    public ReadOnlyMemory<byte> Definition
    {
        get => _definition;
        set
        {
            _definition = value.Length <= 255 ? value : throw new ArgumentException("A definition cannot be longer than 255 bytes.", nameof(value));
            _unchanged = false;
        }
    }

    private ReadOnlyMemory<byte> _definition;

    /// <summary>Gets or sets the flags.</summary>
    public ExcelNameFlags Flags { get; set; }

    /// <summary>
    /// Gets or sets what kind of macro the name is. Setting <see cref="ExcelMacroKind.Command"/> or
    /// <see cref="ExcelMacroKind.Function"/> sets <see cref="ExcelNameFlags.Macro"/>, and <see cref="ExcelMacroKind.None"/>
    /// clears it.
    /// </summary>
    public ExcelMacroKind MacroKind
    {
        get => (Flags & ExcelNameFlags.Macro) == 0 ? ExcelMacroKind.None
            : (RawMacroKind & FunctionBit) != 0 ? ExcelMacroKind.Function
            : ExcelMacroKind.Command;
        set
        {
            if (value == ExcelMacroKind.None)
            {
                Flags &= ~ExcelNameFlags.Macro;
                return;
            }

            Flags |= ExcelNameFlags.Macro;
            RawMacroKind = (byte)((RawMacroKind & 0x0F) | (value == ExcelMacroKind.Function ? FunctionBit : CommandBit));
        }
    }

    /// <summary>
    /// Gets or sets a command macro's shortcut key, typed with Option-Command, or null for none.
    /// </summary>
    public char? ShortcutKey
    {
        get => RawShortcutKey == 0 ? null : MacRoman.GetChar(RawShortcutKey);
        set => RawShortcutKey = value is { } key ? MacRoman.GetByte(key) : (byte)0;
    }

    /// <summary>Gets or sets the shortcut-key byte as stored, a Mac OS Roman character or 0.</summary>
    public byte RawShortcutKey { get; set; }

    /// <summary>
    /// Gets or sets the macro-kind byte as stored: bit 0x40 for a command, 0x80 for a function, and a low nibble left
    /// over from memory in Excel's own files, which Open discards. <see cref="MacroKind"/> reads it.
    /// </summary>
    public byte RawMacroKind { get; set; }

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
        => FormulaDecompiler.Decompile(Definition.Span, [], document, fixupsInPlace: true, _arrayConstants, relativeTo);

    /// <summary>
    /// Writes the definition as <see cref="GetDefinitionText(ExcelDocument?, ExcelCellAddress)"/> gives it into
    /// <paramref name="destination"/> without allocating.
    /// </summary>
    /// <returns>False if the text does not fit in <paramref name="destination"/>.</returns>
    public bool TryFormatDefinitionText(Span<char> destination, out int charsWritten, ExcelDocument? document = null, ExcelCellAddress relativeTo = default)
    {
        var text = new TextBuffer(stackalloc char[256]);
        try
        {
            FormulaDecompiler.AppendDefinition(ref text, this, document, relativeTo);
            return text.TryCopyTo(destination, out charsWritten);
        }
        finally
        {
            text.Dispose();
        }
    }

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
    /// <see cref="ExcelNameFlags.Reference"/> to match it.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public void SetDefinition(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
        => SetDefinition(text, default, document, version);

    /// <summary>
    /// Sets the definition from formula text whose relative references are relative to
    /// <paramref name="relativeTo"/>, as when Excel defines a name with that cell active: <c>=B1</c> with C2 as
    /// <paramref name="relativeTo"/> refers to the cell one row up and one column left of each cell that uses the
    /// name (spec §11). Sets or clears <see cref="ExcelNameFlags.Reference"/> to match the definition.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public void SetDefinition(string text, ExcelCellAddress relativeTo, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
    {
        var (definition, arrays) = ExcelFormulaParser.ParseNameDefinition(text, document, version);
        Rebase(definition, relativeTo);
        Definition = definition;
        ArrayConstants.Clear();
        ArrayConstants.AddRange(arrays);
        Flags = IsReference(definition) ? Flags | ExcelNameFlags.Reference : Flags & ~ExcelNameFlags.Reference;
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
        if ((flags & (byte)ExcelNameFlags.Macro) != 0)
        {
            key = reader.ReadByte();
            extra = reader.ReadByte();
        }

        var definition = reader.ReadMemory(definitionLength);
        if (reader.ReadByte() != definitionLength)
        {
            throw new ExcelFormatException("A defined name's two definition lengths disagree.");
        }

        var name = MacRoman.GetString(reader.ReadSpan(nameLength));
        return new ExcelDefinedName(name, default, (ExcelNameFlags)flags) { Definition = definition, RawShortcutKey = key, RawMacroKind = extra };
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
        writer.WriteByte((byte)Flags);
        if ((Flags & ExcelNameFlags.Macro) != 0)
        {
            writer.WriteByte(RawShortcutKey);
            writer.WriteByte(RawMacroKind);
        }

        writer.WriteBytes(Definition.Span);
        writer.WriteByte((byte)Definition.Length);
        writer.WriteMacRoman(Name);
    }

    /// <inheritdoc/>
    public override string ToString() => Name;
}
