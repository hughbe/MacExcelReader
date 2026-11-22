using System.Buffers.Binary;

namespace MacExcelReader;

/// <summary>
/// A token of a parsed expression (spec §15).
/// </summary>
/// <param name="Offset">The offset of the token within the expression.</param>
/// <param name="RawCode">The token byte as stored, including its class bits.</param>
/// <param name="Code">The token with its class normalised (bit 6 folded into bit 5, then the low 6 bits).</param>
/// <param name="Length">The total length of the token including operands.</param>
public readonly record struct FormulaToken(int Offset, byte RawCode, byte Code, int Length);

/// <summary>
/// Walks parsed expressions and measures the fix-ups that follow them in a file. All methods work on
/// spans and do not allocate.
/// </summary>
public static class FormulaTokens
{
    // Operand sizes of tokens 0x19–0x37 (from the application's own size table).
    private static ReadOnlySpan<byte> OperandSizes =>
    [
        2, 7, 3, 1, 1, 2, 8, 6, 1, 2, 7, 3, 6, 4, 4, 4, 1, 3, 6, 3, 6, 1, 1, 7, 3, 6, 4, 4, 4, 1, 3,

        // Excel 1.5: 0x38, its in-memory form of a call with an argument count; 0x39 and 0x3A, calls to functions
        // numbered 256 and up, with a fixed argument count and with a count (spec §15.2).
        2, 1, 2,
    ];

    /// <summary>Token: name reference.</summary>
    public const byte Name = 0x23;

    /// <summary>Token: start of an external-sheet reference.</summary>
    public const byte ExternalSheet = 0x1A;

    /// <summary>Token: array constant.</summary>
    public const byte Array = 0x20;

    /// <summary>Token: memory area.</summary>
    public const byte MemoryArea = 0x26;

    /// <summary>
    /// Normalises a token's class: bit 6 is folded into bit 5 and the low 6 bits are kept.
    /// </summary>
    public static byte Normalize(byte raw)
    {
        int b = raw;
        if ((b & 0x40) != 0)
        {
            b |= 0x20;
        }

        return (byte)(b & 0x3F);
    }

    /// <summary>
    /// Gets a short descriptive name for a normalised token code, such as "Add" or "Reference".
    /// </summary>
    public static string GetName(byte code) => code switch
    {
        0x01 => "ArrayFormula",
        0x02 => "Table",
        0x03 => "Add",
        0x04 => "Subtract",
        0x05 => "Multiply",
        0x06 => "Divide",
        0x07 => "Power",
        0x08 => "Concatenate",
        0x09 => "LessThan",
        0x0A => "LessEqual",
        0x0B => "Equal",
        0x0C => "GreaterEqual",
        0x0D => "GreaterThan",
        0x0E => "NotEqual",
        0x0F => "Intersection",
        0x10 => "Union",
        0x11 => "Range",
        0x12 => "UnaryPlus",
        0x13 => "UnaryMinus",
        0x14 => "Percent",
        0x15 => "Parentheses",
        0x16 => "MissingArgument",
        0x17 => "String",
        0x19 => "Attribute",
        0x1A => "ExternalSheet",
        0x1B => "ExternalSheetEnd",
        0x1C => "Error",
        0x1D => "Boolean",
        0x1E => "Integer",
        0x1F => "Number",
        0x20 => "Array",
        0x21 => "Function",
        0x22 => "FunctionVarArgs",
        0x23 => "Name",
        0x24 => "Reference",
        0x25 => "Area",
        0x26 => "MemoryArea",
        0x27 => "MemoryError",
        0x28 => "MemoryNoMemory",
        0x29 => "MemoryFunction",
        0x2A => "DeletedReference",
        0x2B => "DeletedArea",
        0x2C => "NameReference",
        0x2D => "NameArea",
        0x2E => "DefinitionMemoryFunction",
        0x2F => "DefinitionMemoryNoMemory",
        >= 0x30 and <= 0x37 => "Marked" + GetName((byte)(code - (code == 0x37 ? 0x13 : 0x0D))),
        0x38 => "FunctionVarArgsInternal",
        0x39 => "FunctionHigh",
        0x3A => "FunctionVarArgsHigh",
        _ => $"Token{code:X2}",
    };

    /// <summary>
    /// Enumerates the tokens of an expression without allocating.
    /// </summary>
    public static TokenEnumerator Enumerate(ReadOnlySpan<byte> expression) => new(expression);

    /// <summary>
    /// Returns a copy of a parsed expression with the bytes Excel leaves over from memory set to zero (spec §15.2),
    /// so that two encodings of the same formula compare equal: the pointer bytes of name (<c>23</c>),
    /// external-sheet (<c>1A</c>, <c>1B</c>), array-constant (<c>20</c>) and deleted-reference (<c>2A</c>) tokens, and
    /// the skip byte of an IF or CHOOSE attribute, and the pointer bytes of a memory area (<c>26</c>) or memory error (<c>27</c>), and the shared-record pointer of an
    /// array formula or table cell (<c>01</c>, <c>02</c>).
    /// </summary>
    /// <param name="expression">The expression.</param>
    /// <param name="definition">Whether it is a name definition, whose fix-ups are applied in place (spec §11).</param>
    public static byte[] ClearStaleBytes(ReadOnlySpan<byte> expression, bool definition = false)
    {
        var copy = expression.ToArray();
        foreach (var token in Enumerate(expression))
        {
            var operands = copy.AsSpan(token.Offset + 1, token.Length - 1);
            switch (token.Code)
            {
                case 0x23:
                    operands[(definition ? 2 : 0)..].Clear();
                    break;
                case 0x1A:
                    operands[..(definition ? 6 : 7)].Clear();
                    break;
                case Array:
                    operands[3..].Clear();
                    break;
                case 0x1B:
                case 0x2A:
                    operands.Clear();
                    break;
                case MemoryArea:
                    operands[..3].Clear();
                    break;
                case 0x27:
                    operands[2] = 0;
                    break;
                case 0x01 or 0x02 when token.Offset == 0:
                    operands.Clear();
                    break;
                case 0x19 when operands[0] is 0x01 or 0x10:
                    operands[1] = 0;
                    break;
            }
        }

        return copy;
    }

    /// <summary>
    /// Validates an expression's token stream and returns the number of tokens.
    /// </summary>
    public static int CountTokens(ReadOnlySpan<byte> expression)
    {
        int count = 0;
        foreach (var _ in Enumerate(expression))
        {
            count++;
        }

        return count;
    }

    /// <summary>
    /// An allocation-free enumerator over the tokens of an expression.
    /// </summary>
    public ref struct TokenEnumerator
    {
        private readonly ReadOnlySpan<byte> _expression;
        private int _position;

        internal TokenEnumerator(ReadOnlySpan<byte> expression)
        {
            _expression = expression;
            _position = 0;
            Current = default;
        }

        /// <summary>Gets the current token.</summary>
        public FormulaToken Current { get; private set; }

        /// <summary>Returns this enumerator.</summary>
        public readonly TokenEnumerator GetEnumerator() => this;

        /// <summary>Advances to the next token.</summary>
        public bool MoveNext()
        {
            int p = _position;
            if (p >= _expression.Length)
            {
                return false;
            }

            byte raw = _expression[p];
            byte code = Normalize(raw);
            int length = 1;
            if (code == 0x17)
            {
                length += 1 + At(p + 1);
            }
            else if (code == 0x19)
            {
                length += 2;
                if ((At(p + 1) & 0x04) != 0)
                {
                    length += At(p + 2) + 1;
                }
            }
            else if (code is 0x01 or 0x02)
            {
                // Array-formula / table reference: a 3-byte pointer to the shared record.
                length += 3;
            }
            else if (code >= 0x19)
            {
                if (code - 0x19 >= OperandSizes.Length)
                {
                    throw new ExcelFormatException($"Unknown token 0x{raw:X2} at offset {p}.");
                }

                length += OperandSizes[code - 0x19];
            }
            else if (code is 0x00 or 0x18)
            {
                throw new ExcelFormatException($"Unknown token 0x{raw:X2} at offset {p}.");
            }

            if (p + length > _expression.Length)
            {
                throw new ExcelFormatException($"Token 0x{raw:X2} at offset {p} runs past the end of the expression.");
            }

            Current = new FormulaToken(p, raw, code, length);
            _position = p + length;
            return true;
        }

        private readonly int At(int index)
        {
            if ((uint)index >= (uint)_expression.Length)
            {
                throw new ExcelFormatException("A token runs past the end of the expression.");
            }

            return _expression[index];
        }
    }

    /// <summary>
    /// Returns true if the expression is an array-formula or table reference (first token 0x01 or 0x02).
    /// </summary>
    public static bool IsSharedReference(ReadOnlySpan<byte> expression)
        => expression.Length > 0 && Normalize(expression[0]) is 0x01 or 0x02;

    // The text of the binary operators, tokens 0x03–0x11 in order (spec §15.1). 0x0F is the space, Excel's
    // intersection operator.
    private static readonly string[] s_binaryOperators = ["+", "-", "*", "/", "^", "&", "<", "<=", "=", ">=", ">", "<>", " ", ",", ":"];

    /// <summary>Gets the text of a binary operator token (0x03–0x11), or "?".</summary>
    internal static string BinaryOperatorText(byte code)
        => code is >= 0x03 and <= 0x11 ? s_binaryOperators[code - 0x03] : "?";

    /// <summary>Gets the token of a binary operator's text.</summary>
    internal static byte BinaryOperatorCode(string text)
    {
        int index = System.Array.IndexOf(s_binaryOperators, text);
        return index >= 0 ? (byte)(index + 0x03) : throw new ArgumentException($"'{text}' is not a binary operator.", nameof(text));
    }

    /// <summary>
    /// The word that starts the trailer of an array formula's or table's top-left cell, where other cells hold the
    /// anchor's row; it cannot be a row (spec §12.6).
    /// </summary>
    internal const ushort SharedAnchorMarker = 0x4000;

    /// <summary>
    /// Measures the fix-ups that follow an expression in a cell formula record (spec §12.5, §12.6).
    /// </summary>
    /// <param name="expression">The expression.</param>
    /// <param name="data">The bytes that follow the expression (and its text result, if any).</param>
    /// <returns>The length of the fix-ups in bytes.</returns>
    internal static int MeasureFixups(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> data)
    {
        int position = 0;
        if (IsSharedReference(expression))
        {
            ushort row = ReadWord(data, ref position);
            Skip(data, ref position, 1);
            if (row == SharedAnchorMarker)
            {
                Require(data, position, 10);
                byte sharedLength = data[position + 9];
                Skip(data, ref position, 10);
                if (Normalize(expression[0]) == 0x01)
                {
                    Require(data, position, sharedLength);
                    var shared = data.Slice(position, sharedLength);
                    position += sharedLength;
                    SkipFixups(shared, data, ref position);
                }
                else
                {
                    Skip(data, ref position, 8);
                }
            }

            return position;
        }

        SkipFixups(expression, data, ref position);
        return position;
    }

    private static void SkipFixups(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> data, ref int position)
    {
        foreach (var token in Enumerate(expression))
        {
            switch (token.Code)
            {
                case Name:
                case ExternalSheet:
                    Skip(data, ref position, 2);
                    break;
                case Array:
                    SkipArrayBlock(data, ref position);
                    break;
                case MemoryArea:
                    int count = ReadWord(data, ref position);
                    Skip(data, ref position, count * 6);
                    break;
            }
        }
    }

    private static void SkipArrayBlock(ReadOnlySpan<byte> data, ref int position)
    {
        Require(data, position, 1);
        int columns = data[position++];
        if (columns == 0)
        {
            columns = 256;
        }

        int rows = ReadWord(data, ref position);
        for (int i = 0; i < rows * columns; i++)
        {
            Require(data, position, 1);
            byte type = data[position++];
            if (type == 2)
            {
                Require(data, position, 1);
                Skip(data, ref position, 1 + data[position]);
            }
            else
            {
                Skip(data, ref position, 8);
            }
        }
    }

    private static void Require(ReadOnlySpan<byte> data, int position, int count)
    {
        if (position + count > data.Length)
        {
            throw new ExcelFormatException("A formula's fix-ups run past the end of the cell data.");
        }
    }

    private static void Skip(ReadOnlySpan<byte> data, ref int position, int count)
    {
        Require(data, position, count);
        position += count;
    }

    private static ushort ReadWord(ReadOnlySpan<byte> data, ref int position)
    {
        Require(data, position, 2);
        ushort value = BinaryPrimitives.ReadUInt16BigEndian(data.Slice(position));
        position += 2;
        return value;
    }

    /// <summary>
    /// Reads the fix-up words of a cell formula in token order: for each name token its 1-based name
    /// index, and for each external-sheet token its external-document index. Array and memory-area
    /// blocks are skipped.
    /// </summary>
    /// <summary>
    /// Returns the array-constant blocks (spec §12.5) for the <c>0x20</c> tokens of an expression, in token order,
    /// from fix-up data laid out as in a cell formula.
    /// </summary>
    internal static IReadOnlyList<byte[]> ReadArrayBlocks(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups)
    {
        if (!ContainsArray(expression))
        {
            return [];
        }

        var blocks = new List<byte[]>();
        int position = 0;
        foreach (var token in Enumerate(expression))
        {
            switch (token.Code)
            {
                case Name:
                case ExternalSheet:
                    Skip(fixups, ref position, 2);
                    break;
                case Array:
                    int start = position;
                    SkipArrayBlock(fixups, ref position);
                    blocks.Add(fixups[start..position].ToArray());
                    break;
                case MemoryArea:
                    Skip(fixups, ref position, ReadWord(fixups, ref position) * 6);
                    break;
            }
        }

        return blocks;
    }

    /// <summary>
    /// Reads one array-constant block per <c>0x20</c> token of <paramref name="expression"/> from consecutive data
    /// (as in the array data after the defined names, spec §11).
    /// </summary>
    internal static IReadOnlyList<byte[]> ReadConsecutiveArrayBlocks(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> data, ref int position)
    {
        if (!ContainsArray(expression))
        {
            return [];
        }

        var blocks = new List<byte[]>();
        foreach (var token in Enumerate(expression))
        {
            if (token.Code == Array)
            {
                int start = position;
                SkipArrayBlock(data, ref position);
                blocks.Add(data[start..position].ToArray());
            }
        }

        return blocks;
    }

    private static bool ContainsArray(ReadOnlySpan<byte> expression)
    {
        foreach (var token in Enumerate(expression))
        {
            if (token.Code == Array)
            {
                return true;
            }
        }

        return false;
    }

    internal static List<ushort> ReadFixupWords(ReadOnlySpan<byte> expression, ReadOnlySpan<byte> fixups)
    {
        var words = new List<ushort>();
        if (IsSharedReference(expression) || fixups.IsEmpty)
        {
            return words;
        }

        int position = 0;
        foreach (var token in Enumerate(expression))
        {
            switch (token.Code)
            {
                case Name:
                case ExternalSheet:
                    words.Add(ReadWord(fixups, ref position));
                    break;
                case Array:
                    SkipArrayBlock(fixups, ref position);
                    break;
                case MemoryArea:
                    Skip(fixups, ref position, ReadWord(fixups, ref position) * 6);
                    break;
            }
        }

        return words;
    }
}
