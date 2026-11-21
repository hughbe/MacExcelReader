using System.Buffers.Binary;
using System.Globalization;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The exception thrown when formula text cannot be parsed.
/// </summary>
public sealed class ExcelFormulaSyntaxException : FormatException
{
    /// <summary>
    /// Initializes a new instance.
    /// </summary>
    public ExcelFormulaSyntaxException(string message, int position)
        : base($"{message} (at position {position + 1})")
    {
        Position = position;
    }

    /// <summary>Gets the zero-based position in the text where the problem was found.</summary>
    public int Position { get; }
}

/// <summary>
/// Compiles A1-style formula text, such as <c>=SUM($A$1:$A$5)*2</c>, into parsed expressions (spec §9),
/// choosing token classes and <c>IF</c> jump tokens the way Excel 1.x does.
/// </summary>
/// <remarks>
/// <para>
/// Supported: numbers, strings, TRUE and FALSE, error values, cell references and areas in A1 notation
/// (with <c>$</c> for absolute parts), external references (<c>Budget!A1</c>, <c>'My Doc'!A1</c>),
/// defined names, every Excel 1.x function and macro command, calls to user-defined function macros,
/// the operators <c>: , % ^ * / + - &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;=</c> and parentheses,
/// references to the active sheet (<c>!A1</c>) and the dialog-box forms of commands (<c>RUN?()</c>).
/// </para>
/// <para>
/// Like Excel, the parser folds a minus sign into a number constant, starts formulas that call volatile
/// functions (RAND, NOW, INDEX, …) with a volatile attribute, and wraps ranges built from function results
/// (<c>INDEX(A1:C3,1,1):C3</c>) in a memory-function token.
/// </para>
/// <para>
/// Excel's own parser was not disassembled in full. The token classes and <c>IF</c> encoding are derived from
/// Excel's function table and from every formula in the sample documents, all of which this parser
/// reproduces byte for byte (apart from stale pointer bytes inside name tokens, which Excel rebuilds when
/// it opens a file).
/// </para>
/// </remarks>
public static class ExcelFormulaParser
{
    /// <summary>
    /// Parses formula text into a formula for a cell. A leading '=' is optional.
    /// </summary>
    /// <param name="text">The formula text.</param>
    /// <param name="document">
    /// The document, used to resolve defined names and external documents. May be null if the formula
    /// uses neither.
    /// </param>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public static ExcelFormula Parse(string text, ExcelDocument? document = null)
    {
        var (expression, fixups) = Compile(text, document, definition: false);
        return new ExcelFormula(expression, fixups);
    }

    /// <summary>
    /// Parses formula text into a name definition (see <see cref="ExcelDefinedName.Definition"/>). Name and
    /// external-document indices are stored inside the tokens, as name definitions require.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    public static byte[] ParseNameDefinition(string text, ExcelDocument? document = null)
        => Compile(text, document, definition: true).Expression;

    private static (byte[] Expression, byte[] Fixups) Compile(string text, ExcelDocument? document, bool definition)
    {
        ArgumentNullException.ThrowIfNull(text);
        var lexer = new Lexer(text);
        var parser = new Parser(lexer, document);
        var root = parser.ParseFormula();
        var emitter = new Emitter(definition);
        if (parser.IsVolatile)
        {
            // Formulas that call a volatile function start with a volatile attribute.
            emitter.Expression.AddRange([0x19, 0x01, 0x00]);
        }

        emitter.Emit(root, definition ? Context.DefinitionTop : Context.Top);
        var expression = emitter.Expression.ToArray();
        if (expression.Length > 255)
        {
            throw new ExcelFormulaSyntaxException($"The formula compiles to {expression.Length} bytes; the limit is 255.", text.Length - 1);
        }

        return (expression, emitter.Fixups.ToArray());
    }

    // ----- Syntax tree -----

    private abstract record Node(int Position);

    private sealed record NumberNode(int Position, double Value, bool ForceDouble = false) : Node(Position);

    private sealed record StringNode(int Position, byte[] Bytes) : Node(Position);

    private sealed record BoolNode(int Position, bool Value) : Node(Position);

    private sealed record ErrorNode(int Position, byte Code) : Node(Position);

    private sealed record MissingNode(int Position) : Node(Position);

    private sealed record RefNode(int Position, ushort RowWord, byte Column, int? External) : Node(Position);

    private sealed record AreaNode(int Position, ushort FirstRowWord, ushort LastRowWord, byte FirstColumn, byte LastColumn, int? External) : Node(Position);

    private sealed record NameNode(int Position, int Index, int? External) : Node(Position);

    private sealed record FunctionNode(int Position, int Number, ExcelFunction? Function, NameNode? Macro, List<Node> Arguments, bool Dialog = false) : Node(Position);

    private sealed record UnaryNode(int Position, byte Operator, Node Operand) : Node(Position);

    private sealed record BinaryNode(int Position, byte Operator, Node Left, Node Right) : Node(Position);

    private sealed record ParenNode(int Position, Node Inner) : Node(Position);

    // ----- Lexer -----

    private enum TokenKind
    {
        End,
        Number,
        String,
        QuotedName,
        Identifier,
        Error,
        Operator,
        OpenParen,
        CloseParen,
        Comma,
        Bang,
    }

    private readonly record struct Token(TokenKind Kind, string Text, int Position, double Number = 0);

    private sealed class Lexer
    {
        private static readonly string[] s_errors = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];
        private readonly string _text;
        private int _position;

        public Lexer(string text)
        {
            _text = text;
            _position = 0;
            SkipSpace();
            if (_position < _text.Length && _text[_position] == '=')
            {
                _position++;
            }
        }

        private void SkipSpace()
        {
            while (_position < _text.Length && char.IsWhiteSpace(_text[_position]))
            {
                _position++;
            }
        }

        public Token Next()
        {
            SkipSpace();
            int start = _position;
            if (_position >= _text.Length)
            {
                return new Token(TokenKind.End, string.Empty, start);
            }

            char c = _text[_position];
            switch (c)
            {
                case '(':
                    _position++;
                    return new Token(TokenKind.OpenParen, "(", start);
                case ')':
                    _position++;
                    return new Token(TokenKind.CloseParen, ")", start);
                case ',':
                    _position++;
                    return new Token(TokenKind.Comma, ",", start);
                case '!':
                    _position++;
                    return new Token(TokenKind.Bang, "!", start);
                case '"':
                    return ReadString(start);
                case '\'':
                    return ReadQuotedName(start);
                case '#':
                    foreach (var error in s_errors)
                    {
                        if (string.Compare(_text, _position, error, 0, error.Length, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            _position += error.Length;
                            return new Token(TokenKind.Error, error, start);
                        }
                    }

                    throw new ExcelFormulaSyntaxException("Unknown error value", start);
                case '<':
                    _position++;
                    if (_position < _text.Length && (_text[_position] == '=' || _text[_position] == '>'))
                    {
                        _position++;
                    }

                    return new Token(TokenKind.Operator, _text[start.._position], start);
                case '>':
                    _position++;
                    if (_position < _text.Length && _text[_position] == '=')
                    {
                        _position++;
                    }

                    return new Token(TokenKind.Operator, _text[start.._position], start);
                case '+' or '-' or '*' or '/' or '^' or '&' or '=' or '%' or ':':
                    _position++;
                    return new Token(TokenKind.Operator, c.ToString(), start);
            }

            if (char.IsAsciiDigit(c) || (c == '.' && _position + 1 < _text.Length && char.IsAsciiDigit(_text[_position + 1])))
            {
                return ReadNumber(start);
            }

            if (IsNameChar(c))
            {
                while (_position < _text.Length && IsNameChar(_text[_position]))
                {
                    _position++;
                }

                return new Token(TokenKind.Identifier, _text[start.._position], start);
            }

            throw new ExcelFormulaSyntaxException($"Unexpected character '{c}'", start);
        }

        private static bool IsNameChar(char c) => char.IsLetterOrDigit(c) || c is '.' or '_' or '$' or '\\' or '?';

        private Token ReadNumber(int start)
        {
            while (_position < _text.Length && (char.IsAsciiDigit(_text[_position]) || _text[_position] == '.'))
            {
                _position++;
            }

            if (_position < _text.Length && (_text[_position] is 'e' or 'E'))
            {
                int save = _position;
                _position++;
                if (_position < _text.Length && _text[_position] is '+' or '-')
                {
                    _position++;
                }

                if (_position < _text.Length && char.IsAsciiDigit(_text[_position]))
                {
                    while (_position < _text.Length && char.IsAsciiDigit(_text[_position]))
                    {
                        _position++;
                    }
                }
                else
                {
                    _position = save;
                }
            }

            var span = _text.AsSpan(start, _position - start);
            if (!double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                throw new ExcelFormulaSyntaxException($"Invalid number '{span}'", start);
            }

            return new Token(TokenKind.Number, span.ToString(), start, value);
        }

        private Token ReadString(int start)
        {
            _position++;
            var builder = new System.Text.StringBuilder();
            while (true)
            {
                if (_position >= _text.Length)
                {
                    throw new ExcelFormulaSyntaxException("Unterminated string", start);
                }

                char c = _text[_position++];
                if (c == '"')
                {
                    if (_position < _text.Length && _text[_position] == '"')
                    {
                        builder.Append('"');
                        _position++;
                        continue;
                    }

                    return new Token(TokenKind.String, builder.ToString(), start);
                }

                builder.Append(c);
            }
        }

        private Token ReadQuotedName(int start)
        {
            _position++;
            var builder = new System.Text.StringBuilder();
            while (true)
            {
                if (_position >= _text.Length)
                {
                    throw new ExcelFormulaSyntaxException("Unterminated quoted name", start);
                }

                char c = _text[_position++];
                if (c == '\'')
                {
                    if (_position < _text.Length && _text[_position] == '\'')
                    {
                        builder.Append('\'');
                        _position++;
                        continue;
                    }

                    return new Token(TokenKind.QuotedName, builder.ToString(), start);
                }

                builder.Append(c);
            }
        }
    }

    // ----- Parser -----

    private sealed class Parser
    {
        private readonly Lexer _lexer;
        private readonly ExcelDocument? _document;
        private Token _current;
        private Token _next;

        public Parser(Lexer lexer, ExcelDocument? document)
        {
            _lexer = lexer;
            _document = document;
            _current = lexer.Next();
            _next = lexer.Next();
        }

        private void Advance()
        {
            _current = _next;
            _next = _lexer.Next();
        }

        /// <summary>Gets whether the formula calls a volatile function.</summary>
        public bool IsVolatile { get; private set; }

        private bool IsOperator(string op) => _current.Kind == TokenKind.Operator && _current.Text == op;

        public Node ParseFormula()
        {
            if (_current.Kind == TokenKind.End)
            {
                throw new ExcelFormulaSyntaxException("The formula is empty", 0);
            }

            var node = ParseComparison();
            if (_current.Kind != TokenKind.End)
            {
                throw new ExcelFormulaSyntaxException($"Unexpected '{_current.Text}'", _current.Position);
            }

            return node;
        }

        private Node ParseComparison()
        {
            var left = ParseConcatenation();
            while (_current.Kind == TokenKind.Operator && _current.Text is "=" or "<>" or "<" or ">" or "<=" or ">=")
            {
                var op = _current;
                Advance();
                byte code = op.Text switch
                {
                    "<" => 0x09,
                    "<=" => 0x0A,
                    "=" => 0x0B,
                    ">=" => 0x0C,
                    ">" => 0x0D,
                    _ => 0x0E,
                };
                left = new BinaryNode(op.Position, code, left, ParseConcatenation());
            }

            return left;
        }

        private Node ParseConcatenation()
        {
            var left = ParseAdditive();
            while (IsOperator("&"))
            {
                int position = _current.Position;
                Advance();
                left = new BinaryNode(position, 0x08, left, ParseAdditive());
            }

            return left;
        }

        private Node ParseAdditive()
        {
            var left = ParseMultiplicative();
            while (IsOperator("+") || IsOperator("-"))
            {
                var op = _current;
                Advance();
                left = new BinaryNode(op.Position, op.Text == "+" ? (byte)0x03 : (byte)0x04, left, ParseMultiplicative());
            }

            return left;
        }

        private Node ParseMultiplicative()
        {
            var left = ParsePower();
            while (IsOperator("*") || IsOperator("/"))
            {
                var op = _current;
                Advance();
                left = new BinaryNode(op.Position, op.Text == "*" ? (byte)0x05 : (byte)0x06, left, ParsePower());
            }

            return left;
        }

        private Node ParsePower()
        {
            var left = ParsePercent();
            while (IsOperator("^"))
            {
                int position = _current.Position;
                Advance();
                left = new BinaryNode(position, 0x07, left, ParsePercent());
            }

            return left;
        }

        private Node ParsePercent()
        {
            var node = ParseUnary();
            while (IsOperator("%"))
            {
                node = new UnaryNode(_current.Position, 0x14, node);
                Advance();
            }

            return node;
        }

        private Node ParseUnary()
        {
            if (IsOperator("-") || IsOperator("+"))
            {
                var op = _current;
                Advance();
                var operand = ParseUnary();

                // Excel folds a minus sign on a number constant into the constant.
                if (op.Text == "-" && operand is NumberNode number)
                {
                    return new NumberNode(op.Position, -number.Value, ForceDouble: true);
                }

                return new UnaryNode(op.Position, op.Text == "-" ? (byte)0x13 : (byte)0x12, operand);
            }

            return ParseRange();
        }

        private Node ParseRange()
        {
            var left = ParsePrimary();
            while (IsOperator(":"))
            {
                int position = _current.Position;
                Advance();
                var right = ParsePrimary();
                left = left is RefNode a && right is RefNode b && a.External == b.External
                    ? new AreaNode(a.Position, a.RowWord, b.RowWord, a.Column, b.Column, a.External)
                    : new BinaryNode(position, 0x11, left, right);
            }

            return left;
        }

        private Node ParsePrimary()
        {
            var token = _current;
            switch (token.Kind)
            {
                case TokenKind.Number:
                    Advance();
                    return new NumberNode(token.Position, token.Number);
                case TokenKind.String:
                {
                    var bytes = MacRoman.GetBytes(token.Text);
                    if (bytes.Length > 255)
                    {
                        throw new ExcelFormulaSyntaxException("A string constant cannot be longer than 255 characters", token.Position);
                    }

                    Advance();
                    return new StringNode(token.Position, bytes);
                }

                case TokenKind.Error:
                    Advance();
                    return new ErrorNode(token.Position, token.Text.ToUpperInvariant() switch
                    {
                        "#NULL!" => 0,
                        "#DIV/0!" => 7,
                        "#VALUE!" => 15,
                        "#REF!" => 23,
                        "#NAME?" => 29,
                        "#NUM!" => 36,
                        _ => 42,
                    });
                case TokenKind.OpenParen:
                {
                    Advance();
                    var inner = ParseComparison();
                    while (_current.Kind == TokenKind.Comma)
                    {
                        int position = _current.Position;
                        Advance();
                        inner = new BinaryNode(position, 0x10, inner, ParseComparison());
                    }

                    Expect(TokenKind.CloseParen, ")");
                    return new ParenNode(token.Position, inner);
                }

                case TokenKind.QuotedName:
                case TokenKind.Identifier when _next.Kind == TokenKind.Bang:
                {
                    Advance();
                    Expect(TokenKind.Bang, "!");
                    int external = ResolveExternal(token.Text, token.Position);
                    var operand = ParsePrimaryReference(external);
                    return operand;
                }

                case TokenKind.Bang:
                {
                    Advance();
                    return ParsePrimaryReference(ResolveExternal(string.Empty, token.Position));
                }

                case TokenKind.Identifier:
                    return ParseIdentifier(external: null);
            }

            throw new ExcelFormulaSyntaxException(token.Kind == TokenKind.End ? "Unexpected end of formula" : $"Unexpected '{token.Text}'", token.Position);
        }

        private Node ParsePrimaryReference(int external)
        {
            if (_current.Kind != TokenKind.Identifier)
            {
                throw new ExcelFormulaSyntaxException("Expected a reference or name after '!'", _current.Position);
            }

            var node = ParseIdentifier(external);
            if (IsOperator(":") && node is RefNode first)
            {
                int position = _current.Position;
                Advance();
                if (_current.Kind == TokenKind.Identifier && ParseIdentifier(external) is RefNode last)
                {
                    return new AreaNode(first.Position, first.RowWord, last.RowWord, first.Column, last.Column, external);
                }

                throw new ExcelFormulaSyntaxException("Expected a cell reference after ':'", position);
            }

            return node;
        }

        private Node ParseIdentifier(int? external)
        {
            var token = _current;
            Advance();
            if (_current.Kind == TokenKind.OpenParen && external is null)
            {
                return ParseCall(token);
            }

            if (external is null && string.Equals(token.Text, "TRUE", StringComparison.OrdinalIgnoreCase))
            {
                return new BoolNode(token.Position, true);
            }

            if (external is null && string.Equals(token.Text, "FALSE", StringComparison.OrdinalIgnoreCase))
            {
                return new BoolNode(token.Position, false);
            }

            if (TryParseReference(token.Text, out ushort rowWord, out byte column))
            {
                return new RefNode(token.Position, rowWord, column, external);
            }

            return new NameNode(token.Position, ResolveName(token.Text, token.Position), external);
        }

        private Node ParseCall(Token name)
        {
            Expect(TokenKind.OpenParen, "(");
            var arguments = new List<Node>();
            if (_current.Kind != TokenKind.CloseParen)
            {
                while (true)
                {
                    arguments.Add(_current.Kind is TokenKind.Comma or TokenKind.CloseParen
                        ? new MissingNode(_current.Position)
                        : ParseComparison());
                    if (_current.Kind == TokenKind.Comma)
                    {
                        Advance();
                        continue;
                    }

                    break;
                }
            }

            Expect(TokenKind.CloseParen, ")");

            bool dialog = name.Text.EndsWith('?') && name.Text.Length > 1;
            int number = ExcelFunctions.FindNumber(dialog ? name.Text[..^1] : name.Text);
            if (number >= 0)
            {
                var function = ExcelFunctions.Get(number)!;
                if (dialog && !function.HasDialogForm)
                {
                    throw new ExcelFormulaSyntaxException($"{function.Name} has no dialog-box form", name.Position);
                }

                if (function.IsVolatile)
                {
                    IsVolatile = true;
                }
                if (arguments.Count < function.MinArguments || arguments.Count > function.MaxArguments)
                {
                    string expected = function.MinArguments == function.MaxArguments
                        ? function.MinArguments.ToString(CultureInfo.InvariantCulture)
                        : $"{function.MinArguments}–{function.MaxArguments}";
                    throw new ExcelFormulaSyntaxException($"{function.Name} takes {expected} argument(s), not {arguments.Count}", name.Position);
                }

                return new FunctionNode(name.Position, number, function, null, arguments, dialog);
            }

            if (TryFindName(name.Text, out int index))
            {
                if (arguments.Count > 254)
                {
                    throw new ExcelFormulaSyntaxException("Too many arguments", name.Position);
                }

                return new FunctionNode(name.Position, 255, null, new NameNode(name.Position, index, null), arguments);
            }

            throw new ExcelFormulaSyntaxException($"Unknown function '{name.Text}'", name.Position);
        }

        private void Expect(TokenKind kind, string text)
        {
            if (_current.Kind != kind)
            {
                throw new ExcelFormulaSyntaxException($"Expected '{text}'", _current.Position);
            }

            Advance();
        }

        private bool TryFindName(string name, out int index)
        {
            index = 0;
            if (_document is null)
            {
                return false;
            }

            for (int i = 0; i < _document.Names.Count; i++)
            {
                if (string.Equals(_document.Names[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    index = i + 1;
                    return true;
                }
            }

            return false;
        }

        private int ResolveName(string name, int position)
        {
            if (TryFindName(name, out int index))
            {
                return index;
            }

            throw new ExcelFormulaSyntaxException(
                _document is null ? $"'{name}' is not a cell reference, and names cannot be resolved without a document" : $"Unknown name '{name}'",
                position);
        }

        private int ResolveExternal(string name, int position)
        {
            // Indices are 1-based. The active sheet (!A1) is the entry named "\0" or "".
            if (_document is not null)
            {
                for (int i = 0; i < _document.ExternalDocuments.Count; i++)
                {
                    var candidate = _document.ExternalDocuments[i].Name;
                    if (name.Length == 0 ? candidate is "" or "\0" : string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return i + 1;
                    }
                }
            }

            throw new ExcelFormulaSyntaxException(
                name.Length == 0
                    ? "A reference to the active sheet (!A1) needs an ExternalDocuments entry named \"\\0\""
                    : $"Unknown external document '{name}'; add it to ExternalDocuments first",
                position);
        }

        private static bool TryParseReference(string text, out ushort rowWord, out byte column)
        {
            rowWord = 0;
            column = 0;
            int i = 0;
            bool columnRelative = true;
            if (i < text.Length && text[i] == '$')
            {
                columnRelative = false;
                i++;
            }

            int start = i;
            while (i < text.Length && char.IsAsciiLetter(text[i]))
            {
                i++;
            }

            int letters = i - start;
            if (letters is 0 or > 2)
            {
                return false;
            }

            bool rowRelative = true;
            if (i < text.Length && text[i] == '$')
            {
                rowRelative = false;
                i++;
            }

            if (i >= text.Length || !int.TryParse(text.AsSpan(i), NumberStyles.None, CultureInfo.InvariantCulture, out int row))
            {
                return false;
            }

            int col = 0;
            for (int k = start; k < start + letters; k++)
            {
                col = col * 26 + (char.ToUpperInvariant(text[k]) - 'A' + 1);
            }

            if (row < 1 || row > ExcelCellGrid.MaxRows || col > ExcelCellGrid.MaxColumns)
            {
                return false;
            }

            rowWord = (ushort)((row - 1) | (rowRelative ? 0x8000 : 0) | (columnRelative ? 0x4000 : 0));
            column = (byte)(col - 1);
            return true;
        }
    }

    // ----- Emitter -----

    private enum ContextKind
    {
        Top,
        DefinitionTop,
        Operator,
        Range,
        Argument,
        IfArgument,
        MacroTarget,
    }

    private readonly record struct Context(ContextKind Kind, ExcelValueTypes Accepts = ExcelValueTypes.None)
    {
        public static Context Top => new(ContextKind.Top);

        public static Context DefinitionTop => new(ContextKind.DefinitionTop);

        public static Context Operator => new(ContextKind.Operator);

        public static Context Range => new(ContextKind.Range);
    }

    private const int ReferenceClass = 1;
    private const int ValueClass = 2;
    private const int ArrayClass = 3;

    private sealed class Emitter
    {
        private readonly bool _definition;

        public Emitter(bool definition) => _definition = definition;

        public List<byte> Expression { get; } = [];

        public List<byte> Fixups { get; } = [];

        private static byte WithClass(byte baseCode, int tokenClass) => (byte)((baseCode & 0x1F) | (tokenClass << 5));

        private static int OperandClass(Context context) => context.Kind switch
        {
            ContextKind.DefinitionTop or ContextKind.Range or ContextKind.MacroTarget => ReferenceClass,
            ContextKind.Argument or ContextKind.IfArgument => (context.Accepts & ExcelValueTypes.Reference) != 0 ? ReferenceClass : ValueClass,
            _ => ValueClass,
        };

        private static int FunctionClass(Context context, ExcelValueTypes returns) => context.Kind switch
        {
            ContextKind.Range or ContextKind.MacroTarget or ContextKind.IfArgument or ContextKind.DefinitionTop => ReferenceClass,
            ContextKind.Argument when returns == ExcelValueTypes.Array && (context.Accepts & ExcelValueTypes.Array) != 0 => ArrayClass,
            ContextKind.Argument when (context.Accepts & ExcelValueTypes.Reference) != 0 && (returns & ExcelValueTypes.Reference) != 0 => ReferenceClass,
            _ => ValueClass,
        };

        private void Add(byte value) => Expression.Add(value);

        private void AddWord(ushort value)
        {
            Expression.Add((byte)(value >> 8));
            Expression.Add((byte)value);
        }

        private void AddFixup(int value)
        {
            Fixups.Add((byte)(value >> 8));
            Fixups.Add((byte)value);
        }

        public void Emit(Node node, Context context)
        {
            switch (node)
            {
                case NumberNode number:
                    if (!number.ForceDouble && number.Value >= 0 && number.Value <= ushort.MaxValue && number.Value == Math.Floor(number.Value))
                    {
                        Add(0x1E);
                        AddWord((ushort)number.Value);
                    }
                    else
                    {
                        Add(0x1F);
                        Span<byte> bytes = stackalloc byte[8];
                        BinaryPrimitives.WriteDoubleBigEndian(bytes, number.Value);
                        foreach (var b in bytes)
                        {
                            Add(b);
                        }
                    }

                    break;
                case StringNode text:
                    Add(0x17);
                    Add((byte)text.Bytes.Length);
                    Expression.AddRange(text.Bytes);
                    break;
                case BoolNode logical:
                    Add(0x1D);
                    Add((byte)(logical.Value ? 1 : 0));
                    break;
                case ErrorNode error when error.Code == 0x17 && (context == Context.DefinitionTop || context == Context.Range):
                    // A name or range that is #REF! refers to deleted cells: a deleted-reference token, whose
                    // row and column bytes are no longer meaningful.
                    Add(WithClass(0x2A, OperandClass(context)));
                    Add(0);
                    Add(0);
                    Add(0);
                    break;
                case ErrorNode error:
                    Add(0x1C);
                    Add(error.Code);
                    break;
                case MissingNode:
                    Add(0x16);
                    break;
                case RefNode reference:
                    EmitExternal(reference.External, () =>
                    {
                        Add(WithClass(_definition ? (byte)0x2C : (byte)0x24, OperandClass(context)));
                        AddWord(reference.RowWord);
                        Add(reference.Column);
                    });
                    break;
                case AreaNode area:
                    EmitExternal(area.External, () =>
                    {
                        Add(WithClass(_definition ? (byte)0x2D : (byte)0x25, OperandClass(context)));
                        AddWord(area.FirstRowWord);
                        AddWord(area.LastRowWord);
                        Add(area.FirstColumn);
                        Add(area.LastColumn);
                    });
                    break;
                case NameNode name:
                    EmitExternal(name.External, () => EmitName(name.Index, OperandClass(context)));
                    break;
                case UnaryNode unary:
                    Emit(unary.Operand, Context.Operator);
                    Add(unary.Operator);
                    break;
                case BinaryNode binary:
                {
                    bool reference = binary.Operator is 0x10 or 0x11;
                    var operandContext = reference ? Context.Range : Context.Operator;

                    // A range or union whose operands call functions is wrapped in a memory-function
                    // token giving the length of the subexpression.
                    int start = -1;
                    if (reference && ContainsCall(binary))
                    {
                        Add(0x29);
                        Add(0);
                        start = Expression.Count;
                    }

                    Emit(binary.Left, operandContext);
                    Emit(binary.Right, operandContext);
                    Add(binary.Operator);
                    if (start >= 0)
                    {
                        int length = Expression.Count - start;
                        if (length > byte.MaxValue)
                        {
                            throw new ExcelFormulaSyntaxException("The range expression is too long", binary.Position);
                        }

                        Expression[start - 1] = (byte)length;
                    }

                    break;
                }
                case ParenNode paren:
                    Emit(paren.Inner, context);
                    Add(0x15);
                    break;
                case FunctionNode function when function.Number == 1:
                    EmitIf(function, context);
                    break;
                case FunctionNode function:
                    EmitCall(function, context);
                    break;
            }
        }

        private static bool ContainsCall(Node node) => node switch
        {
            FunctionNode => true,
            BinaryNode binary => ContainsCall(binary.Left) || ContainsCall(binary.Right),
            UnaryNode unary => ContainsCall(unary.Operand),
            ParenNode paren => ContainsCall(paren.Inner),
            _ => false,
        };

        private void EmitName(int index, int tokenClass)
        {
            Add(WithClass(0x23, tokenClass));
            if (_definition)
            {
                AddWord((ushort)index);
                for (int i = 0; i < 5; i++)
                {
                    Add(0);
                }
            }
            else
            {
                for (int i = 0; i < 7; i++)
                {
                    Add(0);
                }

                AddFixup(index);
            }
        }

        private void EmitExternal(int? external, Action emitOperand)
        {
            if (external is not { } index)
            {
                emitOperand();
                return;
            }

            // 1A: four bytes Excel leaves stale, 00 00, then the 1-based document index (which cell formulas
            // also carry in the fix-ups).
            Add(0x1A);
            for (int i = 0; i < 6; i++)
            {
                Add(0);
            }

            Add((byte)index);
            if (!_definition)
            {
                AddFixup(index);
            }

            emitOperand();
            Add(0x1B);
            Add(0);
            Add(0);
            Add(0);
        }

        private void EmitCall(FunctionNode function, Context context)
        {
            if (function.Macro is { } macro)
            {
                EmitName(macro.Index, ReferenceClass);
                foreach (var argument in function.Arguments)
                {
                    Emit(argument, new Context(ContextKind.Argument, (ExcelValueTypes)0x7F));
                }

                Add(WithClass(0x22, FunctionClass(context, (ExcelValueTypes)0x7F)));
                Add((byte)(function.Arguments.Count + 1));
                Add(255);
                return;
            }

            var definition = function.Function!;
            for (int i = 0; i < function.Arguments.Count; i++)
            {
                Emit(function.Arguments[i], new Context(ContextKind.Argument, definition.GetArgumentType(i)));
            }

            int tokenClass = FunctionClass(context, definition.ReturnType);
            if (function.Dialog)
            {
                Add(WithClass(0x22, tokenClass));
                Add((byte)(function.Arguments.Count | 0x80));
            }
            else if (definition.HasFixedArguments)
            {
                Add(WithClass(0x21, tokenClass));
            }
            else
            {
                Add(WithClass(0x22, tokenClass));
                Add((byte)function.Arguments.Count);
            }

            Add((byte)function.Number);
        }

        // IF(condition, true[, false]): condition, a jump over the true branch, the true branch and a skip
        // past the IF token, then the false branch and its skip, then the IF token itself.
        private void EmitIf(FunctionNode function, Context context)
        {
            var definition = function.Function!;
            var arguments = function.Arguments;
            Emit(arguments[0], new Context(ContextKind.IfArgument, definition.GetArgumentType(0)));
            Add(0x19);
            Add(0x02);
            int jump = Expression.Count;
            Add(0);
            int afterJump = Expression.Count;

            var skips = new List<int>();
            for (int i = 1; i < arguments.Count; i++)
            {
                if (i == 2)
                {
                    SetOffset(jump, Expression.Count - afterJump, function.Position);
                }

                Emit(arguments[i], new Context(ContextKind.IfArgument, definition.GetArgumentType(i)));
                Add(0x19);
                Add(0x08);
                skips.Add(Expression.Count);
                Add(0);
            }

            if (arguments.Count == 2)
            {
                SetOffset(jump, Expression.Count - afterJump, function.Position);
            }

            Add(WithClass(0x22, FunctionClass(context, definition.ReturnType)));
            Add((byte)arguments.Count);
            Add(1);
            int end = Expression.Count;
            foreach (int skip in skips)
            {
                SetOffset(skip, end - (skip + 1) - 1, function.Position);
            }
        }

        private void SetOffset(int index, int value, int position)
        {
            if (value is < 0 or > 255)
            {
                throw new ExcelFormulaSyntaxException("An IF branch is too long to encode (over 255 bytes)", position);
            }

            Expression[index] = (byte)value;
        }
    }
}
