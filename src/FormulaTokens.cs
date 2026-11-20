using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A token of a parsed expression (spec §9).
/// </summary>
/// <param name="Offset">The offset of the token within the expression.</param>
/// <param name="RawCode">The token byte as stored, including its class bits.</param>
/// <param name="Code">The token with its class normalised (bit 6 folded into bit 5, then the low 6 bits).</param>
/// <param name="Length">The total length of the token including operands.</param>
public readonly record struct FormulaToken(int Offset, byte RawCode, byte Code, int Length);

/// <summary>
/// Walks parsed expressions and measures the fix-ups that follow them in a file.
/// </summary>
public static class FormulaTokens
{
    // Operand sizes of tokens 0x19–0x37 (from the application's own size table).
    private static readonly byte[] s_operandSizes =
    [
        2, 7, 3, 1, 1, 2, 8, 6, 1, 2, 7, 3, 6, 4, 4, 4, 1, 3, 6, 3, 6, 1, 1, 7, 3, 6, 4, 4, 4, 1, 3,
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
    /// Enumerates the tokens of an expression.
    /// </summary>
    public static IEnumerable<FormulaToken> Enumerate(byte[] expression)
    {
        ArgumentNullException.ThrowIfNull(expression);
        int p = 0;
        while (p < expression.Length)
        {
            byte raw = expression[p];
            byte code = Normalize(raw);
            int length = 1;
            if (code == 0x17)
            {
                length += 1 + At(expression, p + 1);
            }
            else if (code == 0x19)
            {
                length += 2;
                if ((At(expression, p + 1) & 0x04) != 0)
                {
                    length += At(expression, p + 2) + 1;
                }
            }
            else if (code is 0x01 or 0x02)
            {
                // Array-formula / table reference: a 3-byte pointer to the shared record.
                length += 3;
            }
            else if (code >= 0x19)
            {
                length += s_operandSizes[code - 0x19];
            }

            if (p + length > expression.Length)
            {
                throw new ExcelFormatException($"Token 0x{raw:X2} at offset {p} runs past the end of the expression.");
            }

            yield return new FormulaToken(p, raw, code, length);
            p += length;
        }
    }

    private static int At(byte[] data, int index)
    {
        if (index >= data.Length)
        {
            throw new ExcelFormatException("A token runs past the end of the expression.");
        }

        return data[index];
    }

    /// <summary>
    /// Returns true if the expression is an array-formula or table reference (first token 0x01 or 0x02).
    /// </summary>
    public static bool IsSharedReference(byte[] expression)
        => expression.Length > 0 && Normalize(expression[0]) is 0x01 or 0x02;

    /// <summary>
    /// Reads the fix-ups that follow an expression in a cell formula record (spec §8.5, §8.6) and returns
    /// their raw bytes.
    /// </summary>
    internal static byte[] ReadFixups(byte[] expression, BigEndianReader reader)
    {
        int start = reader.Position;
        if (IsSharedReference(expression))
        {
            ReadSharedTrailer(expression, reader);
        }
        else
        {
            SkipFixups(expression, reader);
        }

        int end = reader.Position;
        reader.Position = start;
        return reader.ReadBytes(end - start);
    }

    private static void SkipFixups(byte[] expression, BigEndianReader reader)
    {
        foreach (var token in Enumerate(expression))
        {
            switch (token.Code)
            {
                case Name:
                case ExternalSheet:
                    reader.Skip(2);
                    break;
                case Array:
                    SkipArrayBlock(reader);
                    break;
                case MemoryArea:
                    int count = reader.ReadUInt16();
                    reader.Skip(count * 6);
                    break;
            }
        }
    }

    private static void SkipArrayBlock(BigEndianReader reader)
    {
        int columns = reader.ReadByte();
        if (columns == 0)
        {
            columns = 256;
        }

        int rows = reader.ReadUInt16();
        for (int i = 0; i < rows * columns; i++)
        {
            byte type = reader.ReadByte();
            if (type == 2)
            {
                reader.ReadPascalBytes();
            }
            else
            {
                reader.Skip(8);
            }
        }
    }

    private static void ReadSharedTrailer(byte[] expression, BigEndianReader reader)
    {
        ushort row = reader.ReadUInt16();
        reader.ReadByte();
        if (row != 0x4000)
        {
            return;
        }

        var header = reader.ReadBytes(10);
        if (Normalize(expression[0]) == 0x01)
        {
            var shared = reader.ReadBytes(header[9]);
            SkipFixups(shared, reader);
        }
        else
        {
            reader.Skip(8);
        }
    }

    /// <summary>
    /// Reads the fix-up words of a cell formula in token order: for each name token its 1-based name
    /// index, and for each external-sheet token its external-document index. Array and memory-area
    /// blocks are skipped.
    /// </summary>
    internal static List<ushort> ReadFixupWords(byte[] expression, byte[] fixups)
    {
        var words = new List<ushort>();
        if (IsSharedReference(expression) || fixups.Length == 0)
        {
            return words;
        }

        var reader = new BigEndianReader(fixups);
        foreach (var token in Enumerate(expression))
        {
            switch (token.Code)
            {
                case Name:
                case ExternalSheet:
                    words.Add(reader.ReadUInt16());
                    break;
                case Array:
                    SkipArrayBlock(reader);
                    break;
                case MemoryArea:
                    reader.Skip(reader.ReadUInt16() * 6);
                    break;
            }
        }

        return words;
    }
}
