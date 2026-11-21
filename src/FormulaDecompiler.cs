using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace MacExcelReader;

/// <summary>
/// Converts parsed expressions (spec §9) to A1-style formula text.
/// </summary>
public static class FormulaDecompiler
{
    /// <summary>
    /// Converts a parsed expression to text, without the leading '='.
    /// </summary>
    /// <param name="expression">The expression bytes.</param>
    /// <param name="fixupWords">Name and external-document indices in token order (cell formulas).</param>
    /// <param name="document">The document used to resolve names and external documents; may be null.</param>
    /// <param name="fixupsInPlace">True for name definitions, whose indices are stored inside the tokens.</param>
    public static string Decompile(ReadOnlySpan<byte> expression, IReadOnlyList<ushort> fixupWords, ExcelDocument? document, bool fixupsInPlace)
    {
        var stack = new Stack<string>();
        int fixupIndex = 0;
        string? externalPrefix = null;
        int external = 0;

        foreach (var token in FormulaTokens.Enumerate(expression))
        {
            int p = token.Offset + 1;
            switch (token.Code)
            {
                case 0x01:
                case 0x02:
                    stack.Push(token.Code == 0x01 ? "{array formula}" : "TABLE()");
                    break;
                case >= 0x03 and <= 0x11:
                {
                    string right = Pop(stack);
                    string left = Pop(stack);
                    stack.Push(string.Concat(left, BinaryOperator(token.Code), right));
                    break;
                }

                case 0x12:
                    stack.Push("+" + Pop(stack));
                    break;
                case 0x13:
                    stack.Push("-" + Pop(stack));
                    break;
                case 0x14:
                    stack.Push(Pop(stack) + "%");
                    break;
                case 0x15:
                    stack.Push("(" + Pop(stack) + ")");
                    break;
                case 0x16:
                    stack.Push(string.Empty);
                    break;
                case 0x17:
                {
                    var text = MacRoman.GetString(expression.Slice(p + 1, expression[p]));
                    stack.Push("\"" + text.Replace("\"", "\"\"") + "\"");
                    break;
                }

                case 0x19:
                    if ((expression[p] & 0x10) != 0)
                    {
                        stack.Push("SUM(" + Pop(stack) + ")");
                    }

                    break;
                case 0x1A:
                {
                    ushort index = NextFixup(expression, token, fixupWords, fixupsInPlace, ref fixupIndex);
                    externalPrefix = ExternalDocumentName(document, index) + "!";
                    external = index;
                    break;
                }

                case 0x1B:
                    external = 0;
                    break;
                case 0x1C:
                    stack.Push(ExcelFormulaResult.ErrorText(expression[p]));
                    break;
                case 0x1D:
                    stack.Push(expression[p] != 0 ? "TRUE" : "FALSE");
                    break;
                case 0x1E:
                    stack.Push(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p)).ToString(CultureInfo.InvariantCulture));
                    break;
                case 0x1F:
                    stack.Push(BinaryPrimitives.ReadDoubleBigEndian(expression.Slice(p)).ToString("R", CultureInfo.InvariantCulture));
                    break;
                case 0x20:
                    stack.Push("{array}");
                    break;
                case 0x21:
                {
                    var function = ExcelFunctions.Get(expression[p]);
                    stack.Push(CallText(stack, function, expression[p], function?.MinArguments ?? 0));
                    break;
                }

                case 0x22:
                {
                    // Bit 7 of the argument count marks a command's dialog-box form, such as RUN?().
                    var function = ExcelFunctions.Get(expression[p + 1]);
                    bool dialog = (expression[p] & 0x80) != 0;
                    var call = CallText(stack, function, expression[p + 1], expression[p] & 0x7F);
                    stack.Push(dialog ? call.Insert(call.IndexOf('('), "?") : call);
                    break;
                }

                case 0x23:
                case 0x30:
                {
                    ushort index = NextFixup(expression, token, fixupWords, fixupsInPlace, ref fixupIndex);
                    // Inside an external reference a name token indexes that document's names (spec §10).
                    stack.Push(WithPrefix(ref externalPrefix, external != 0 ? ExternalNameText(document, external, index) : NameText(document, index)));
                    break;
                }

                case 0x24:
                case 0x2C:
                case 0x31:
                case 0x37:
                    stack.Push(WithPrefix(ref externalPrefix, CellText(BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p)), expression[p + 2])));
                    break;
                case 0x25:
                case 0x2D:
                case 0x32:
                {
                    ushort row1 = BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p));
                    ushort row2 = BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(p + 2));
                    stack.Push(WithPrefix(ref externalPrefix, CellText(row1, expression[p + 4]) + ":" + CellText(row2, expression[p + 5])));
                    break;
                }

                case 0x2A:
                case 0x2B:
                    stack.Push(WithPrefix(ref externalPrefix, "#REF!"));
                    break;
                default:
                    // Memory tokens (0x26–0x29, 0x2E, 0x2F, 0x33–0x36) carry no text of their own.
                    break;
            }
        }

        return stack.Count switch
        {
            0 => string.Empty,
            1 => stack.Pop(),
            _ => string.Join(" ", stack.Reverse()),
        };
    }

    private static string Pop(Stack<string> stack) => stack.Count > 0 ? stack.Pop() : "?";

    private static string WithPrefix(ref string? externalPrefix, string operand)
    {
        if (externalPrefix is null)
        {
            return operand;
        }

        var result = externalPrefix + operand;
        externalPrefix = null;
        return result;
    }

    private static ushort NextFixup(ReadOnlySpan<byte> expression, FormulaToken token, IReadOnlyList<ushort> fixupWords, bool fixupsInPlace, ref int fixupIndex)
    {
        if (fixupsInPlace)
        {
            return token.Code == FormulaTokens.ExternalSheet
                ? expression[token.Offset + 7]
                : BinaryPrimitives.ReadUInt16BigEndian(expression.Slice(token.Offset + 1));
        }

        return fixupIndex < fixupWords.Count ? fixupWords[fixupIndex++] : (ushort)0;
    }

    private static string CallText(Stack<string> stack, ExcelFunction? function, int number, int argumentCount)
    {
        var arguments = new string[argumentCount];
        for (int i = argumentCount - 1; i >= 0; i--)
        {
            arguments[i] = stack.Count > 0 ? stack.Pop() : "?";
        }

        // Function 255 calls a user-defined function macro; its first argument is the macro.
        if (number == 255 && argumentCount > 0)
        {
            return arguments[0] + "(" + string.Join(",", arguments.Skip(1)) + ")";
        }

        string name = function?.Name ?? $"FUNCTION{number}";
        return name + "(" + string.Join(",", arguments) + ")";
    }

    private static string BinaryOperator(byte code) => code switch
    {
        0x03 => "+",
        0x04 => "-",
        0x05 => "*",
        0x06 => "/",
        0x07 => "^",
        0x08 => "&",
        0x09 => "<",
        0x0A => "<=",
        0x0B => "=",
        0x0C => ">=",
        0x0D => ">",
        0x0E => "<>",
        0x0F => " ",
        0x10 => ",",
        0x11 => ":",
        _ => "?",
    };

    /// <summary>
    /// Formats a cell reference from a row word (bit 15 = relative row, bit 14 = relative column) and a column.
    /// </summary>
    public static string CellText(ushort rowWord, byte column)
    {
        bool rowRelative = (rowWord & 0x8000) != 0;
        bool columnRelative = (rowWord & 0x4000) != 0;
        int row = rowWord & 0x3FFF;
        var builder = new StringBuilder();
        if (!columnRelative)
        {
            builder.Append('$');
        }

        builder.Append(ColumnName(column));
        if (!rowRelative)
        {
            builder.Append('$');
        }

        builder.Append(row + 1);
        return builder.ToString();
    }

    /// <summary>
    /// Gets the letters of a zero-based column index (0 = A, 255 = IV).
    /// </summary>
    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        var name = string.Empty;
        column++;
        while (column > 0)
        {
            int remainder = (column - 1) % 26;
            name = (char)('A' + remainder) + name;
            column = (column - 1) / 26;
        }

        return name;
    }

    private static string QuoteIfNeeded(string name)
    {
        foreach (char c in name)
        {
            if (!(char.IsLetterOrDigit(c) || c is '.' or '_'))
            {
                return "'" + name.Replace("'", "''") + "'";
            }
        }

        return name;
    }

    private static string NameText(ExcelDocument? document, ushort index)
    {
        if (document is not null && index >= 1 && index <= document.Names.Count)
        {
            return document.Names[index - 1].Name;
        }

        return $"NAME{index}";
    }

    private static string ExternalNameText(ExcelDocument? document, int external, ushort index)
    {
        if (document is not null && external <= document.ExternalDocuments.Count)
        {
            var names = document.ExternalDocuments[external - 1].Names;
            if (index >= 1 && index <= names.Count)
            {
                return names[index - 1];
            }
        }

        return $"NAME{index}";
    }

    private static string ExternalDocumentName(ExcelDocument? document, ushort index)
    {
        // External-document indices are 1-based. A document named "\0" (or "") means the active sheet,
        // written as !A1.
        if (document is not null && index >= 1 && index <= document.ExternalDocuments.Count)
        {
            var name = document.ExternalDocuments[index - 1].Name;
            return name.Length == 0 || name == "\0" ? string.Empty : QuoteIfNeeded(name);
        }

        return $"DOC{index}";
    }
}
