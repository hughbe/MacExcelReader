using System.Buffers;
using System.Globalization;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// Compiles A1-style formula text, such as <c>=SUM($A$1:$A$5)*2</c>, into parsed expressions (spec §15),
/// choosing token classes and <c>IF</c> jump tokens the way Excel 1.x does.
/// </summary>
/// <remarks>
/// <para>
/// The compiler reads the text once, building a syntax tree of struct nodes, and then writes the tokens from the tree.
/// Its arrays and buffers are kept per thread and reused. Tokens are spans of the text, so names, references and strings are
/// never copied out of it: compiling a cell formula allocates only the <see cref="ExcelFormula"/> it returns.
/// </para>
/// <para>
/// The tree is needed because a token's class depends on where its operand ends up (<c>=A1</c> and <c>=A1+1</c>
/// differ), and IF and CHOOSE jumps and memory tokens need the length of what follows them (spec §15.4).
/// </para>
/// </remarks>
internal static class ExcelFormulaParser
{
    /// <summary>
    /// Parses formula text into a formula for a cell. A leading '=' is optional.
    /// </summary>
    /// <param name="text">The formula text.</param>
    /// <param name="document">
    /// The document, used to resolve defined names and external documents. May be null if the formula
    /// uses neither.
    /// </param>
    /// <param name="version">
    /// The Excel release whose functions and encoding to use. Excel 1.00–1.03 (the default) produce formulas
    /// every Excel 1.x can read; Excel 1.5 adds functions and encodes a few things differently (spec §15.4).
    /// </param>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    internal static ExcelFormula Parse(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(text);
        var compiler = new Compiler(text, document, version, definition: false, arrayFormula: false);
        try
        {
            compiler.Compile();
            return ExcelFormula.FromCompiled(compiler.Expression.AsSpan(), compiler.Fixups.AsSpan());
        }
        finally
        {
            compiler.Dispose();
        }
    }

    /// <summary>
    /// Parses formula text into a name definition (see <see cref="ExcelDefinedName.Definition"/>). Name and
    /// external-document indices are stored inside the tokens, as name definitions require.
    /// </summary>
    /// <exception cref="ExcelFormulaSyntaxException">The text is not a valid formula.</exception>
    internal static (byte[] Definition, List<byte[]> ArrayConstants) ParseNameDefinition(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(text);
        var compiler = new Compiler(text, document, version, definition: true, arrayFormula: false);
        try
        {
            compiler.Compile();
            return (compiler.Expression.ToArray(), compiler.ArrayConstants ?? []);
        }
        finally
        {
            compiler.Dispose();
        }
    }

    /// <summary>
    /// Parses the formula of an array formula (without the braces) into the shared expression and its fix-ups.
    /// Use <see cref="ExcelDocument.SetArrayFormula"/> to put an array formula on a sheet.
    /// </summary>
    internal static (byte[] Expression, byte[] Fixups) ParseArrayFormula(string text, ExcelDocument? document = null, ExcelVersion version = ExcelVersion.Excel100)
    {
        ArgumentNullException.ThrowIfNull(text);
        var compiler = new Compiler(text, document, version, definition: false, arrayFormula: true);
        try
        {
            compiler.Compile();
            return (compiler.Expression.ToArray(), compiler.Fixups.ToArray());
        }
        finally
        {
            compiler.Dispose();
        }
    }

    // ----- Tokens -----

    private enum TokenKind : byte
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
        OpenBrace,
        CloseBrace,
        Semicolon,
    }

    // Operators, numbered so that the binary ones (Plus to NotEqual) are their tokens, 0x03–0x0E (spec §15.1).
    private enum Op : byte
    {
        None = 0,
        Plus = 0x03,
        Minus = 0x04,
        Times = 0x05,
        Divide = 0x06,
        Power = 0x07,
        Concatenate = 0x08,
        Less = 0x09,
        LessEqual = 0x0A,
        Equal = 0x0B,
        GreaterEqual = 0x0C,
        Greater = 0x0D,
        NotEqual = 0x0E,
        Percent = 0x14,
        Colon = 0x11,
    }

    private const byte IntersectionOperator = 0x0F;
    private const byte UnionOperator = 0x10;
    private const byte RangeOperator = 0x11;

    // A token: where it is in the text and what it holds. Strings and quoted names are the text between the quotes,
    // with doubled quotes still doubled (Escaped says whether there are any).
    private readonly struct Token
    {
        public TokenKind Kind { get; init; }

        public Op Op { get; init; }

        /// <summary>Gets where the token starts, quote included: the position errors report.</summary>
        public int Position { get; init; }

        public int Start { get; init; }

        public int Length { get; init; }

        public double Number { get; init; }

        /// <summary>Gets whether a number token was written as plain digits, as a row number is.</summary>
        public bool IsDigits { get; init; }

        public bool Escaped { get; init; }

        public ExcelError Error { get; init; }

        /// <summary>Gets whether white space came before the token: between two references, the intersection operator.</summary>
        public bool SpaceBefore { get; init; }
    }

    // ----- Lexer -----

    private ref struct Lexer
    {
        private readonly ReadOnlySpan<char> _text;
        private int _position;

        public Lexer(ReadOnlySpan<char> text)
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
            int before = _position;
            SkipSpace();
            bool spaceBefore = _position > before;
            return NextToken() with { SpaceBefore = spaceBefore };
        }

        private static Token Simple(TokenKind kind, int start, int length, Op op = Op.None)
            => new() { Kind = kind, Op = op, Position = start, Start = start, Length = length };

        private Token NextToken()
        {
            int start = _position;
            if (_position >= _text.Length)
            {
                return Simple(TokenKind.End, start, 0);
            }

            char c = _text[_position];
            switch (c)
            {
                case '(':
                    _position++;
                    return Simple(TokenKind.OpenParen, start, 1);
                case ')':
                    _position++;
                    return Simple(TokenKind.CloseParen, start, 1);
                case ',':
                    _position++;
                    return Simple(TokenKind.Comma, start, 1);
                case '!':
                    _position++;
                    return Simple(TokenKind.Bang, start, 1);
                case '{':
                    _position++;
                    return Simple(TokenKind.OpenBrace, start, 1);
                case '}':
                    _position++;
                    return Simple(TokenKind.CloseBrace, start, 1);
                case ';':
                    _position++;
                    return Simple(TokenKind.Semicolon, start, 1);
                case '"':
                    return ReadQuoted(start, '"', TokenKind.String, "Unterminated string");
                case '\'':
                    return ReadQuoted(start, '\'', TokenKind.QuotedName, "Unterminated quoted name");
                case '#':
                    if (ExcelErrorExtensions.TryMatch(_text[_position..], out var error, out int length))
                    {
                        _position += length;
                        return Simple(TokenKind.Error, start, length) with { Error = error };
                    }

                    throw new ExcelFormulaSyntaxException("Unknown error value", start);
                case '<':
                    _position++;
                    if (_position < _text.Length && _text[_position] is '=' or '>')
                    {
                        _position++;
                        return Simple(TokenKind.Operator, start, 2, _text[start + 1] == '=' ? Op.LessEqual : Op.NotEqual);
                    }

                    return Simple(TokenKind.Operator, start, 1, Op.Less);
                case '>':
                    _position++;
                    if (_position < _text.Length && _text[_position] == '=')
                    {
                        _position++;
                        return Simple(TokenKind.Operator, start, 2, Op.GreaterEqual);
                    }

                    return Simple(TokenKind.Operator, start, 1, Op.Greater);
                case '+' or '-' or '*' or '/' or '^' or '&' or '=' or '%' or ':':
                    _position++;
                    return Simple(TokenKind.Operator, start, 1, c switch
                    {
                        '+' => Op.Plus,
                        '-' => Op.Minus,
                        '*' => Op.Times,
                        '/' => Op.Divide,
                        '^' => Op.Power,
                        '&' => Op.Concatenate,
                        '=' => Op.Equal,
                        '%' => Op.Percent,
                        _ => Op.Colon,
                    });
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

                return Simple(TokenKind.Identifier, start, _position - start);
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

            if (_position < _text.Length && _text[_position] is 'e' or 'E')
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

            var span = _text[start.._position];
            bool digits = !span.ContainsAnyExcept("0123456789");
            if (digits && span.Length <= 15)
            {
                // Up to 15 digits is a whole number a double holds exactly, the value double.TryParse would give.
                long whole = 0;
                foreach (char d in span)
                {
                    whole = whole * 10 + (d - '0');
                }

                return Simple(TokenKind.Number, start, span.Length) with { Number = whole, IsDigits = true };
            }

            // Excel 1.00 takes a constant other than 0 only from 1E-307 up to below 1E+306, however it is written:
            // =9.99999E305 and =100E-309 compile, =0.1E307 and =9.9E-308 are not formulas (`Constant Edges`).
            if (!double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                || (value == 0 ? !IsZero(span) : value < MinConstant || value >= MaxConstant))
            {
                throw new ExcelFormulaSyntaxException($"Invalid number '{span}'", start);
            }

            return Simple(TokenKind.Number, start, span.Length) with { Number = value, IsDigits = digits };
        }

        // Whether number text is zero, as opposed to a number too small for a double (1E-400 reads as 0).
        private static bool IsZero(ReadOnlySpan<char> number)
        {
            int exponent = number.IndexOfAny('E', 'e');
            return !(exponent < 0 ? number : number[..exponent]).ContainsAnyExcept('0', '.');
        }

        private const double MinConstant = 1E-307;
        private const double MaxConstant = 1E306;

        // A string or quoted name: the text up to the closing quote, a doubled quote standing for one.
        private Token ReadQuoted(int start, char quote, TokenKind kind, string unterminated)
        {
            int contentStart = ++_position;
            bool escaped = false;
            while (true)
            {
                int close = _position < _text.Length ? _text[_position..].IndexOf(quote) : -1;
                if (close < 0)
                {
                    throw new ExcelFormulaSyntaxException(unterminated, start);
                }

                _position += close + 1;
                if (_position < _text.Length && _text[_position] == quote)
                {
                    escaped = true;
                    _position++;
                    continue;
                }

                return new Token
                {
                    Kind = kind,
                    Position = start,
                    Start = contentStart,
                    Length = _position - 1 - contentStart,
                    Escaped = escaped,
                };
            }
        }
    }

    // ----- Syntax tree -----

    private enum NodeKind : byte
    {
        Number,
        String,
        Bool,
        Error,
        Missing,
        Ref,
        Area,
        Name,
        Function,
        Unary,
        Binary,
        Paren,
        Array,
    }

    // A node of the syntax tree, kept by index in a pooled array. Which fields are used depends on the kind:
    // Number: Number, Flag (folded sign, so always a double). String: A, B (the text between the quotes), Flag
    // (doubled quotes). Bool, Error: Code. Ref: Row1, Column1, External. Area: Row1, Row2, Column1, Column2, External.
    // Name: A (index), External. Function: A (number), B, C (first argument in the argument list, count), D (the
    // name node of a macro, or -1), Flag (dialog form). Unary: Code, A. Binary: Code, A, B. Paren: A. Array: A, B
    // (the block in the array buffer).
    private struct Node
    {
        public NodeKind Kind;
        public byte Code;
        public bool Flag;
        public byte Column1;
        public byte Column2;
        public ushort Row1;
        public ushort Row2;
        public int Position;
        public int A;
        public int B;
        public int C;
        public int D;
        public int External;
        public double Number;
    }

    // ----- Buffers -----

    // The compiler's buffers, kept per thread between compilations so that compiling rents and returns nothing.
    // Buffers that grew past what any formula that compiles needs are dropped rather than kept.
    private sealed class Workspace
    {
        public Node[] Nodes = new Node[64];
        public int[] Arguments = new int[32];
        public int[] Pending = new int[32];
        public byte[] Expression = new byte[256];
        public byte[] Fixups = new byte[64];
        public byte[] Arrays = new byte[64];
    }

    [ThreadStatic]
    private static Workspace? t_workspace;

    private const int KeptNodes = 4096;
    private const int KeptBytes = 16384;

    // A growable byte buffer over a workspace array. A mutable struct: keep it in a field and never copy it.
    private struct Bytes(byte[] buffer)
    {
        public byte[] Buffer = buffer;

        public int Count;

        public byte this[int index]
        {
            readonly get => Buffer[index];
            set => Buffer[index] = value;
        }

        public void Add(byte value)
        {
            if (Count == Buffer.Length)
            {
                Array.Resize(ref Buffer, Buffer.Length * 2);
            }

            Buffer[Count++] = value;
        }

        public void AddWord(int value)
        {
            Add((byte)(value >> 8));
            Add((byte)value);
        }

        public void AddDouble(double value)
        {
            Reserve(8);
            System.Buffers.Binary.BinaryPrimitives.WriteDoubleBigEndian(Buffer.AsSpan(Count), value);
            Count += 8;
        }

        public void AddRange(ReadOnlySpan<byte> values)
        {
            Reserve(values.Length);
            values.CopyTo(Buffer.AsSpan(Count));
            Count += values.Length;
        }

        private void Reserve(int count)
        {
            if (Count + count > Buffer.Length)
            {
                Array.Resize(ref Buffer, Math.Max(Buffer.Length * 2, Count + count));
            }
        }

        public readonly ReadOnlySpan<byte> AsSpan() => Buffer.AsSpan(0, Count);

        public readonly ReadOnlySpan<byte> Slice(int start, int length) => Buffer.AsSpan(start, length);

        public readonly byte[] ToArray() => AsSpan().ToArray();
    }

    // ----- Compiler -----

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
        // For an operator's operands: whether the operator itself sits where an array is accepted.
        public bool InArrayArgument { get; init; }

        public static Context Top => new(ContextKind.Top);

        public static Context DefinitionTop => new(ContextKind.DefinitionTop);

        public static Context Operator => new(ContextKind.Operator);

        public static Context Range => new(ContextKind.Range);
    }

    private const int ReferenceClass = 1;
    private const int ValueClass = 2;
    private const int ArrayClass = 3;

    // An area a memory-area token caches: rows and columns of the sheet.
    private readonly record struct CachedArea(ushort FirstRow, ushort LastRow, byte FirstColumn, byte LastColumn);

    private ref struct Compiler
    {
        private readonly ReadOnlySpan<char> _text;
        private readonly ExcelDocument? _document;
        private readonly ExcelVersion _version;
        private readonly bool _definition;
        private readonly bool _arrayFormula;

        private Lexer _lexer;
        private Token _current;
        private Token _next;
        private Token _third;

        // The tree, the arguments of its calls (node indices, each call's in a run), the arguments of the calls being
        // parsed (innermost last), and the blocks of its array constants.
        private Node[] _nodes;
        private int _nodeCount;
        private int[] _arguments;
        private int _argumentCount;
        private int[] _pending;
        private int _pendingCount;
        private Bytes _arrays;
        private readonly Workspace _workspace;

        // Every level of nesting compiles to at least one byte, so a formula nested deeper than an expression can be
        // long cannot compile; stopping there also keeps the recursion off the end of the stack.
        private const int MaxDepth = 256;
        private int _depth;

        private int _highestFunction;
        private bool _isVolatile;
        private bool _insideMemoryFunction;

        // Whether a comma is the union operator: everywhere but directly in a function's arguments, where it
        // separates them.
        private bool _commaIsUnion = true;

        // The output. Fields, not properties: the builders are mutable structs.
        public Bytes Expression;
        public Bytes Fixups;

        // The data blocks of a name definition's array constants, created for the first one.
        public List<byte[]>? ArrayConstants;

        public Compiler(string text, ExcelDocument? document, ExcelVersion version, bool definition, bool arrayFormula)
        {
            _text = text;
            _document = document;
            _version = version;
            _definition = definition;
            _arrayFormula = arrayFormula;
            // Take this thread's workspace (or a new one, if a compilation on this thread already has it).
            _workspace = t_workspace ?? new Workspace();
            t_workspace = null;
            _nodes = _workspace.Nodes;
            _arguments = _workspace.Arguments;
            _pending = _workspace.Pending;
            _arrays = new Bytes(_workspace.Arrays);
            Expression = new Bytes(_workspace.Expression);
            Fixups = new Bytes(_workspace.Fixups);
            _highestFunction = -1;
            _lexer = default;
            _current = default;
            _next = default;
            _third = default;
        }

        // Gives the workspace back to the thread, with the buffers as they grew unless they grew too large to keep.
        public void Dispose()
        {
            var workspace = _workspace;
            workspace.Nodes = _nodes.Length <= KeptNodes ? _nodes : new Node[64];
            workspace.Arguments = _arguments.Length <= KeptNodes ? _arguments : new int[32];
            workspace.Pending = _pending.Length <= KeptNodes ? _pending : new int[32];
            workspace.Arrays = _arrays.Buffer.Length <= KeptBytes ? _arrays.Buffer : new byte[64];
            workspace.Expression = Expression.Buffer.Length <= KeptBytes ? Expression.Buffer : new byte[256];
            workspace.Fixups = Fixups.Buffer.Length <= KeptBytes ? Fixups.Buffer : new byte[64];
            t_workspace = workspace;
        }

        // Parses the text and writes its tokens.
        public void Compile()
        {
            _lexer = new Lexer(_text);
            _current = _lexer.Next();
            _next = _lexer.Next();
            _third = _lexer.Next();
            int root = ParseFormula();

            // A document that uses a function newer than Excel 1.03 records the highest number it needs (spec §5).
            if (_document is not null && _highestFunction >= _document.FunctionCount)
            {
                _document.FunctionCount = (ushort)(_highestFunction + 1);
            }

            if (_isVolatile)
            {
                // Formulas that call a volatile function start with a volatile attribute.
                Expression.Add(0x19);
                Expression.Add(0x01);
                Expression.Add(0x00);
            }

            Emit(root, _definition ? Context.DefinitionTop : Context.Top);
            if (Expression.Count > 255)
            {
                throw new ExcelFormulaSyntaxException($"The formula compiles to {Expression.Count} bytes; the limit is 255.", _text.Length - 1);
            }
        }

        // ----- Parser -----

        private void Advance()
        {
            _current = _next;
            _next = _third;
            _third = _lexer.Next();
        }

        private readonly ReadOnlySpan<char> TextOf(Token token) => _text.Slice(token.Start, token.Length);

        private readonly bool IsOperator(Op op) => _current.Kind == TokenKind.Operator && _current.Op == op;

        // The token as written, for error messages: a string or quoted name without its quotes.
        private readonly string Display(Token token) => token.Kind switch
        {
            TokenKind.Number => token.Number.ToString("R", CultureInfo.InvariantCulture),
            TokenKind.String => Unescape(TextOf(token), '"'),
            TokenKind.QuotedName => Unescape(TextOf(token), '\''),
            TokenKind.Error => token.Error.ToText(),
            _ => TextOf(token).ToString(),
        };

        private static string Unescape(ReadOnlySpan<char> text, char quote)
            => text.Contains(quote) ? text.ToString().Replace(new string(quote, 2), quote.ToString(), StringComparison.Ordinal) : text.ToString();

        private int Add(in Node node)
        {
            if (_nodeCount == _nodes.Length)
            {
                Array.Resize(ref _nodes, _nodes.Length * 2);
            }

            _nodes[_nodeCount] = node;
            return _nodeCount++;
        }

        private int Binary(int position, byte code, int left, int right)
            => Add(new Node { Kind = NodeKind.Binary, Position = position, Code = code, A = left, B = right });

        private int ParseFormula()
        {
            if (_current.Kind == TokenKind.End)
            {
                throw new ExcelFormulaSyntaxException("The formula is empty", 0);
            }

            int node = ParseComparison();

            if (_current.Kind != TokenKind.End)
            {
                throw new ExcelFormulaSyntaxException($"Unexpected '{Display(_current)}'", _current.Position);
            }

            return node;
        }

        private void Enter()
        {
            if (++_depth > MaxDepth)
            {
                throw new ExcelFormulaSyntaxException($"The formula is nested more than {MaxDepth} levels deep", _current.Position);
            }
        }

        private int ParseComparison()
        {
            Enter();
            int left = ParseConcatenation();
            while (_current.Kind == TokenKind.Operator && _current.Op is Op.Equal or Op.NotEqual or Op.Less or Op.Greater or Op.LessEqual or Op.GreaterEqual)
            {
                var op = _current;
                Advance();
                int right = ParseConcatenation();
                left = Binary(op.Position, (byte)op.Op, left, right);
            }

            _depth--;
            return left;
        }

        private int ParseConcatenation()
        {
            int left = ParseAdditive();
            while (IsOperator(Op.Concatenate))
            {
                int position = _current.Position;
                Advance();
                int right = ParseAdditive();
                left = Binary(position, (byte)Op.Concatenate, left, right);
            }

            return left;
        }

        private int ParseAdditive()
        {
            int left = ParseMultiplicative();
            while (IsOperator(Op.Plus) || IsOperator(Op.Minus))
            {
                var op = _current;
                Advance();
                int right = ParseMultiplicative();
                left = Binary(op.Position, (byte)op.Op, left, right);
            }

            return left;
        }

        private int ParseMultiplicative()
        {
            int left = ParsePower();
            while (IsOperator(Op.Times) || IsOperator(Op.Divide))
            {
                var op = _current;
                Advance();
                int right = ParsePower();
                left = Binary(op.Position, (byte)op.Op, left, right);
            }

            return left;
        }

        private int ParsePower()
        {
            int left = ParsePercent();
            while (IsOperator(Op.Power))
            {
                int position = _current.Position;
                Advance();
                int right = ParsePercent();
                left = Binary(position, (byte)Op.Power, left, right);
            }

            return left;
        }

        private int ParsePercent()
        {
            int node = ParseUnary();
            while (IsOperator(Op.Percent))
            {
                node = Add(new Node { Kind = NodeKind.Unary, Position = _current.Position, Code = 0x14, A = node });
                Advance();
            }

            return node;
        }

        private int ParseUnary()
        {
            if (IsOperator(Op.Minus) || IsOperator(Op.Plus))
            {
                var op = _current;
                Advance();

                // A sign written against a number makes a signed constant before ':' is looked at, so =-30:31 and
                // =+30:31 are refused where =- 30:31 and =-$30:$31 are whole rows (Excel 1.00, `Typed Signs`).
                if (_current.Kind == TokenKind.Number && !_current.SpaceBefore && _next.Kind == TokenKind.Operator && _next.Op == Op.Colon)
                {
                    throw new ExcelFormulaSyntaxException("':' needs a reference", _current.Position);
                }

                Enter();
                int operand = ParseUnary();
                _depth--;

                // Excel folds a minus sign on a number constant into the constant, and drops a plus sign on one
                // (Excel compiles =+1 to 1E 00 01).
                if (_nodes[operand].Kind == NodeKind.Number)
                {
                    if (op.Op == Op.Minus)
                    {
                        return Add(new Node { Kind = NodeKind.Number, Position = op.Position, Number = -_nodes[operand].Number, Flag = true });
                    }

                    return operand;
                }

                return Add(new Node { Kind = NodeKind.Unary, Position = op.Position, Code = op.Op == Op.Minus ? (byte)0x13 : (byte)0x12, A = operand });
            }

            return ParseUnion();
        }

        // Union: a comma outside a function's arguments (=D30:E31,F32, or (A1,B2) in parentheses). It binds more
        // tightly than any operator but range and intersection: Excel 1.00 compiles =1+D30,E31 as 1+(D30,E31),
        // =D30,E31=1 as (D30,E31)=1 and =-D30,E31 as -(D30,E31) (`Typed Unions`).
        private int ParseUnion()
        {
            int left = ParseIntersection();
            while (_commaIsUnion && _current.Kind == TokenKind.Comma)
            {
                int position = _current.Position;
                Advance();
                int right = RequireReference(ParseIntersection(), "A union");
                left = Binary(position, UnionOperator, RequireReference(left, "A union"), right);
            }

            return left;
        }

        // Intersection: white space between two references (=A1:B2 B1:C3). It binds less tightly than a range and
        // more tightly than a union.
        private int ParseIntersection()
        {
            int left = ParseRange();
            while (_current.SpaceBefore && IsReferenceNode(left)
                && (_current.Kind is TokenKind.Identifier or TokenKind.QuotedName or TokenKind.Bang or TokenKind.OpenParen || StartsWholeRows()))
            {
                int position = _current.Position;
                int right = RequireReference(ParseRange(), "An intersection");
                left = Binary(position, IntersectionOperator, left, right);
            }

            return left;
        }

        // Whether the current token starts whole rows written with a relative first row, as in =30:32 31:33.
        private readonly bool StartsWholeRows()
            => _current.Kind == TokenKind.Number && _next.Kind == TokenKind.Operator && _next.Op == Op.Colon && !_next.SpaceBefore && !_third.SpaceBefore
                && TryRow(_current, out _, out _) && TryRow(_third, out _, out _);

        private readonly bool IsReferenceNode(int index)
        {
            ref readonly var node = ref _nodes[index];
            return node.Kind switch
            {
                NodeKind.Ref or NodeKind.Area or NodeKind.Name => true,
                NodeKind.Paren => IsReferenceNode(node.A),
                NodeKind.Binary => node.Code is IntersectionOperator or UnionOperator or RangeOperator,
                // A function that can return a reference (IF, CHOOSE, INDEX), or a macro's, whose result is unknown.
                NodeKind.Function => node.A == 255 || (ExcelFunctions.Get(node.A, _version) is { } function
                    && (function.ReturnType & ExcelValueTypes.Reference) != 0),
                _ => false,
            };
        }

        // Excel's compiler takes only references as the operands of ':', intersection and union, and as an argument
        // that accepts nothing else (ROW, COLUMN, AREAS): =A1:"", =SUM(A1):B2 and =ROW(SUM(A1)) are syntax errors
        // (Excel 1.00, docs/Testing.md).
        private readonly int RequireReference(int index, string what)
        {
            if (!IsReferenceNode(index))
            {
                throw new ExcelFormulaSyntaxException($"{what} needs a reference", _nodes[index].Position);
            }

            return index;
        }

        private int ParseRange()
        {
            int left = TryParseWholeRowsOrColumns(0) is int whole ? whole : ParsePrimary();
            while (IsOperator(Op.Colon))
            {
                int position = _current.Position;
                Advance();
                int right = TryParseWholeRowsOrColumns(0) is int rightWhole ? rightWhole : ParsePrimary();
                ref readonly var a = ref _nodes[left];
                if (a.Kind == NodeKind.Ref && _nodes[right].Kind == NodeKind.Ref && a.External == _nodes[right].External)
                {
                    left = Area(a.Position, a.Row1, _nodes[right].Row1, a.Column1, _nodes[right].Column1, a.External);
                    continue;
                }

                // Excel pairs cell:cell into an area before it applies ':' to anything else, so B2:C3 in
                // =B$2:$C3:B2:C3 or =(A1,C3):B2:C3 is one area token (Excel 1.00, docs/Testing.md).
                if (_nodes[right].Kind == NodeKind.Ref && IsOperator(Op.Colon) && _next.Kind == TokenKind.Identifier)
                {
                    int pairPosition = _current.Position;
                    Advance();
                    int last = TryParseWholeRowsOrColumns(0) is int lastWhole ? lastWhole : ParsePrimary();
                    ref readonly var b = ref _nodes[right];
                    if (_nodes[last].Kind == NodeKind.Ref && b.External == _nodes[last].External)
                    {
                        left = Binary(position, RangeOperator, RequireReference(left, "':'"), Area(b.Position, b.Row1, _nodes[last].Row1, b.Column1, _nodes[last].Column1, b.External));
                    }
                    else
                    {
                        left = Binary(pairPosition, RangeOperator, Binary(position, RangeOperator, RequireReference(left, "':'"), right), RequireReference(last, "':'"));
                    }

                    continue;
                }

                left = Binary(position, RangeOperator, RequireReference(left, "':'"), RequireReference(right, "':'"));
            }

            return left;
        }

        private int Area(int position, ushort firstRow, ushort lastRow, byte firstColumn, byte lastColumn, int external)
            => Add(new Node { Kind = NodeKind.Area, Position = position, Row1 = firstRow, Row2 = lastRow, Column1 = firstColumn, Column2 = lastColumn, External = external });

        // Whole rows, 1:3 or $18:$19, and whole columns, A:B or $A:$C, as Excel 1.00 reads them. Whole rows take every
        // column, A to IV, as absolute columns, and whole columns every row as absolute rows.
        private int? TryParseWholeRowsOrColumns(int external)
        {
            if (!(_next.Kind == TokenKind.Operator && _next.Op == Op.Colon && !_next.SpaceBefore && !_third.SpaceBefore))
            {
                return null;
            }

            if (TryRow(_current, out int firstRow, out bool firstRelative) && TryRow(_third, out int lastRow, out bool lastRelative))
            {
                int position = _current.Position;
                Advance();
                Advance();
                Advance();
                return Area(position, RowWord.Create(firstRow, firstRelative, false), RowWord.Create(lastRow, lastRelative, false),
                    0, ExcelCellGrid.MaxColumns - 1, external);
            }

            if (TryColumn(_current, out int firstColumn, out firstRelative) && TryColumn(_third, out int lastColumn, out lastRelative))
            {
                int position = _current.Position;
                Advance();
                Advance();
                Advance();
                return Area(position, RowWord.Create(0, false, firstRelative), RowWord.Create(ExcelCellGrid.MaxRows - 1, false, lastRelative),
                    (byte)firstColumn, (byte)lastColumn, external);
            }

            return null;
        }

        private readonly bool TryRow(Token token, out int row, out bool relative)
        {
            row = 0;
            relative = token.Kind == TokenKind.Number;
            if (token.Kind == TokenKind.Number && token.IsDigits && token.Number is >= 1 and <= ExcelCellGrid.MaxRows)
            {
                row = (int)token.Number - 1;
                return true;
            }

            var text = TextOf(token);
            if (token.Kind == TokenKind.Identifier && text.Length > 1 && text[0] == '$'
                && int.TryParse(text[1..], NumberStyles.None, CultureInfo.InvariantCulture, out int number)
                && number is >= 1 and <= ExcelCellGrid.MaxRows)
            {
                row = number - 1;
                return true;
            }

            return false;
        }

        private readonly bool TryColumn(Token token, out int column, out bool relative)
        {
            column = 0;
            relative = true;
            if (token.Kind != TokenKind.Identifier)
            {
                return false;
            }

            var letters = TextOf(token);
            if (letters.Length > 0 && letters[0] == '$')
            {
                relative = false;
                letters = letters[1..];
            }

            if (letters.Length is < 1 or > 2 || !IsAsciiLetters(letters))
            {
                return false;
            }

            foreach (char c in letters)
            {
                column = column * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
            }

            column--;
            return column < ExcelCellGrid.MaxColumns;
        }

        private static bool IsAsciiLetters(ReadOnlySpan<char> text)
        {
            foreach (char c in text)
            {
                if (!char.IsAsciiLetter(c))
                {
                    return false;
                }
            }

            return true;
        }

        private int ParsePrimary()
        {
            var token = _current;
            switch (token.Kind)
            {
                case TokenKind.Number:
                    Advance();
                    return Add(new Node { Kind = NodeKind.Number, Position = token.Position, Number = token.Number });
                case TokenKind.String:
                    if (StringLength(token) > 255)
                    {
                        throw new ExcelFormulaSyntaxException("A string constant cannot be longer than 255 characters", token.Position);
                    }

                    Advance();
                    return Add(new Node { Kind = NodeKind.String, Position = token.Position, A = token.Start, B = token.Length, Flag = token.Escaped });
                case TokenKind.Error:
                    Advance();
                    return Add(new Node { Kind = NodeKind.Error, Position = token.Position, Code = (byte)token.Error });
                case TokenKind.OpenParen:
                {
                    Advance();
                    bool union = _commaIsUnion;
                    _commaIsUnion = true;
                    int inner = ParseComparison();
                    _commaIsUnion = union;
                    Expect(TokenKind.CloseParen, ")");
                    return Add(new Node { Kind = NodeKind.Paren, Position = token.Position, A = inner });
                }

                case TokenKind.QuotedName:
                case TokenKind.Identifier when _next.Kind == TokenKind.Bang:
                {
                    Advance();
                    Expect(TokenKind.Bang, "!");
                    int external = ResolveExternal(token);
                    return ParsePrimaryReference(external);
                }

                case TokenKind.Bang:
                {
                    Advance();
                    return ParsePrimaryReference(ResolveExternal(token));
                }

                case TokenKind.Identifier:
                    return ParseIdentifier(0);

                case TokenKind.OpenBrace:
                    return ParseArrayConstant();
            }

            throw new ExcelFormulaSyntaxException(token.Kind == TokenKind.End ? "Unexpected end of formula" : $"Unexpected '{Display(token)}'", token.Position);
        }

        // The number of characters, and so of Mac OS Roman bytes, a string token stands for.
        private readonly int StringLength(Token token)
            => token.Escaped ? token.Length - TextOf(token).Count("\"\"") : token.Length;

        // {1,2;3,4}: rows separated by ';', columns by ','. Elements are numbers (optionally negative), strings,
        // TRUE, FALSE and errors (spec §12.5). The block is written as the elements are read: column count (0 =
        // 256), row count, then each element: type 1 and a double, type 2 and a Pascal string, or type 4 (TRUE,
        // FALSE) or 16 (an error) and 8 bytes holding the value in the second.
        private int ParseArrayConstant()
        {
            var start = _current;
            Advance();
            int block = _arrays.Count;
            _arrays.Add(0);
            _arrays.Add(0);
            _arrays.Add(0);
            int rows = 0;
            int columns = -1;
            int inRow = 0;
            bool ragged = false;
            while (true)
            {
                bool negative = false;
                if (_current.Kind == TokenKind.Operator && _current.Op is Op.Minus or Op.Plus)
                {
                    negative = _current.Op == Op.Minus;
                    Advance();
                }

                if (_current.Kind == TokenKind.Number)
                {
                    _arrays.Add(1);
                    _arrays.AddDouble(negative ? -_current.Number : _current.Number);
                }
                else if (_current.Kind == TokenKind.String && !negative)
                {
                    int length = StringLength(_current);
                    if (length > 255)
                    {
                        throw new ExcelFormulaSyntaxException("A string constant cannot be longer than 255 characters", _current.Position);
                    }

                    _arrays.Add(2);
                    _arrays.Add((byte)length);
                    AddMacRoman(ref _arrays, TextOf(_current), _current.Escaped);
                }
                else if (_current.Kind == TokenKind.Identifier && !negative && BooleanText.TryParse(TextOf(_current), out bool logical))
                {
                    AddElement(4, (byte)(logical ? 1 : 0));
                }
                else if (_current.Kind == TokenKind.Error && !negative)
                {
                    AddElement(16, (byte)_current.Error);
                }
                else
                {
                    throw new ExcelFormulaSyntaxException("An array constant can hold only numbers, strings, TRUE, FALSE and errors", _current.Position);
                }

                inRow++;
                Advance();
                if (_current.Kind == TokenKind.Comma)
                {
                    Advance();
                    continue;
                }

                bool last = _current.Kind != TokenKind.Semicolon;
                if (last)
                {
                    Expect(TokenKind.CloseBrace, "}");
                }
                else
                {
                    Advance();
                }

                // The row ends: every row needs the first row's number of columns, checked once the constant is read.
                rows++;
                ragged |= columns >= 0 && inRow != columns;
                if (columns < 0)
                {
                    columns = inRow;
                }

                inRow = 0;
                if (last)
                {
                    break;
                }
            }

            if (ragged)
            {
                throw new ExcelFormulaSyntaxException("Every row of an array constant needs the same number of columns", start.Position);
            }

            if (columns > 256 || rows > ushort.MaxValue)
            {
                throw new ExcelFormulaSyntaxException("The array constant is too large", start.Position);
            }

            _arrays[block] = (byte)columns;
            _arrays[block + 1] = (byte)(rows >> 8);
            _arrays[block + 2] = (byte)rows;
            return Add(new Node { Kind = NodeKind.Array, Position = start.Position, A = block, B = _arrays.Count - block });
        }

        private void AddElement(byte type, byte value)
        {
            _arrays.Add(type);
            _arrays.Add(0);
            _arrays.Add(value);
            for (int i = 0; i < 6; i++)
            {
                _arrays.Add(0);
            }
        }

        private int ParsePrimaryReference(int external)
        {
            if (TryParseWholeRowsOrColumns(external) is int whole)
            {
                return whole;
            }

            if (_current.Kind != TokenKind.Identifier)
            {
                throw new ExcelFormulaSyntaxException("Expected a reference or name after '!'", _current.Position);
            }

            int node = ParseIdentifier(external);
            if (IsOperator(Op.Colon) && _nodes[node].Kind == NodeKind.Ref)
            {
                int position = _current.Position;
                Advance();
                if (_current.Kind == TokenKind.Identifier)
                {
                    int last = ParseIdentifier(external);
                    if (_nodes[last].Kind == NodeKind.Ref)
                    {
                        ref readonly var first = ref _nodes[node];
                        return Area(first.Position, first.Row1, _nodes[last].Row1, first.Column1, _nodes[last].Column1, external);
                    }
                }

                throw new ExcelFormulaSyntaxException("Expected a cell reference after ':'", position);
            }

            return node;
        }

        private int ParseIdentifier(int external)
        {
            var token = _current;
            Advance();
            if (_current.Kind == TokenKind.OpenParen)
            {
                return external != 0 ? ParseExternalMacroCall(token, external) : ParseCall(token);
            }

            var text = TextOf(token);
            if (external == 0 && BooleanText.TryParse(text, out bool logical))
            {
                return Add(new Node { Kind = NodeKind.Bool, Position = token.Position, Code = (byte)(logical ? 1 : 0) });
            }

            if (TryParseReference(text, out ushort rowWord, out byte column))
            {
                return Add(new Node { Kind = NodeKind.Ref, Position = token.Position, Row1 = rowWord, Column1 = column, External = external });
            }

            int index = external != 0 ? ResolveExternalName(external, text, token.Position) : ResolveName(text, token.Position);
            return Add(new Node { Kind = NodeKind.Name, Position = token.Position, A = index, External = external });
        }

        private int ParseCall(Token name)
        {
            var (first, count) = ParseArguments();
            var text = TextOf(name);
            bool dialog = text.Length > 1 && text[^1] == '?';
            if (ExcelFunctions.TryFindNumber(dialog ? text[..^1] : text, out int number, _version))
            {
                var function = ExcelFunctions.Get(number, _version)!;
                _highestFunction = Math.Max(_highestFunction, number);
                if (dialog && !function.HasDialogForm)
                {
                    throw new ExcelFormulaSyntaxException($"{function.Name} has no dialog-box form", name.Position);
                }

                if (function.IsVolatile)
                {
                    _isVolatile = true;
                }

                if (count < function.MinArguments || count > function.MaxArguments)
                {
                    string expected = function.MinArguments == function.MaxArguments
                        ? function.MinArguments.ToString(CultureInfo.InvariantCulture)
                        : $"{function.MinArguments}–{function.MaxArguments}";
                    throw new ExcelFormulaSyntaxException($"{function.Name} takes {expected} argument(s), not {count}", name.Position);
                }

                for (int i = 0; i < count; i++)
                {
                    int argument = _arguments[first + i];
                    if (function.GetArgumentType(i) == ExcelValueTypes.Reference && _nodes[argument].Kind != NodeKind.Missing)
                    {
                        RequireReference(argument, $"{function.Name}'s argument {i + 1}");
                    }
                }

                return Add(new Node { Kind = NodeKind.Function, Position = name.Position, A = number, B = first, C = count, D = -1, Flag = dialog });
            }

            if (TryFindName(text, out int index))
            {
                if (count > 254)
                {
                    throw new ExcelFormulaSyntaxException("Too many arguments", name.Position);
                }

                int macro = Add(new Node { Kind = NodeKind.Name, Position = name.Position, A = index });
                return Add(new Node { Kind = NodeKind.Function, Position = name.Position, A = 255, B = first, C = count, D = macro });
            }

            throw new ExcelFormulaSyntaxException($"Unknown function '{text}'", name.Position);
        }

        // 'Other Sheet'!Macro(arguments): a user-defined function on another sheet.
        private int ParseExternalMacroCall(Token name, int external)
        {
            int index = ResolveExternalName(external, TextOf(name), name.Position);
            int macro = Add(new Node { Kind = NodeKind.Name, Position = name.Position, A = index, External = external });
            var (first, count) = ParseArguments();
            if (count > 254)
            {
                throw new ExcelFormulaSyntaxException("Too many arguments", name.Position);
            }

            return Add(new Node { Kind = NodeKind.Function, Position = name.Position, A = 255, B = first, C = count, D = macro });
        }

        // Reads a call's arguments onto the pending stack, then moves them to the argument list as one run.
        private (int First, int Count) ParseArguments()
        {
            Expect(TokenKind.OpenParen, "(");
            int first = _pendingCount;
            bool union = _commaIsUnion;
            _commaIsUnion = false;
            if (_current.Kind != TokenKind.CloseParen)
            {
                while (true)
                {
                    int argument = _current.Kind is TokenKind.Comma or TokenKind.CloseParen
                        ? Add(new Node { Kind = NodeKind.Missing, Position = _current.Position })
                        : ParseComparison();
                    Push(ref _pending, ref _pendingCount, argument);
                    if (_current.Kind == TokenKind.Comma)
                    {
                        Advance();
                        continue;
                    }

                    break;
                }
            }

            _commaIsUnion = union;
            Expect(TokenKind.CloseParen, ")");
            int count = _pendingCount - first;
            int start = _argumentCount;
            for (int i = first; i < _pendingCount; i++)
            {
                Push(ref _arguments, ref _argumentCount, _pending[i]);
            }

            _pendingCount = first;
            return (start, count);
        }

        private static void Push(ref int[] array, ref int count, int value)
        {
            if (count == array.Length)
            {
                Array.Resize(ref array, array.Length * 2);
            }

            array[count++] = value;
        }

        private void Expect(TokenKind kind, string text)
        {
            if (_current.Kind != kind)
            {
                throw new ExcelFormulaSyntaxException($"Expected '{text}'", _current.Position);
            }

            Advance();
        }

        private readonly bool TryFindName(ReadOnlySpan<char> name, out int index)
        {
            index = 0;
            if (_document is null)
            {
                return false;
            }

            // Excel matches a name ignoring case and accents, as it orders them (spec §11).
            var names = _document.Names;
            for (int i = 0; i < names.Count; i++)
            {
                if (NameKey.Equal(name, names[i].Name))
                {
                    index = i + 1;
                    return true;
                }
            }

            return false;
        }

        private readonly int ResolveName(ReadOnlySpan<char> name, int position)
        {
            if (TryFindName(name, out int index))
            {
                return index;
            }

            throw new ExcelFormulaSyntaxException(
                _document is null ? $"'{name}' is not a cell reference, and names cannot be resolved without a document" : $"Unknown name '{name}'",
                position);
        }

        // Names inside an external reference index the document's own list of names used (spec §10).
        private readonly int ResolveExternalName(int external, ReadOnlySpan<char> name, int position)
        {
            var entry = _document!.ExternalDocuments[external - 1];
            for (int i = 0; i < entry.Names.Count; i++)
            {
                if (name.Equals(entry.Names[i], StringComparison.OrdinalIgnoreCase))
                {
                    return i + 1;
                }
            }

            throw new ExcelFormulaSyntaxException(
                $"Unknown name '{name}' in external document '{entry.Name.Replace("\0", "(active sheet)")}'; add it to that entry's Names first",
                position);
        }

        // The external document a 'Name'!, Name! or ! prefix refers to. Indices are 1-based; the active sheet (!A1)
        // is the entry named "\0" or "".
        private readonly int ResolveExternal(Token token)
        {
            ReadOnlySpan<char> raw = token.Kind == TokenKind.Bang ? default : TextOf(token);
            bool escaped = token.Kind == TokenKind.QuotedName && token.Escaped;
            if (_document is not null)
            {
                for (int i = 0; i < _document.ExternalDocuments.Count; i++)
                {
                    var candidate = _document.ExternalDocuments[i].Name;
                    if (raw.IsEmpty ? candidate is "" or "\0" : EqualsName(raw, escaped, candidate))
                    {
                        return i + 1;
                    }
                }
            }

            throw new ExcelFormulaSyntaxException(
                raw.IsEmpty
                    ? "A reference to the active sheet (!A1) needs an ExternalDocuments entry named \"\\0\""
                    : $"Unknown external document '{(escaped ? Unescape(raw, '\'') : raw.ToString())}'; add it to ExternalDocuments first",
                token.Position);
        }

        // Compares a document name as written, quotes doubled if escaped, with a name, ignoring case.
        private static bool EqualsName(ReadOnlySpan<char> raw, bool escaped, string name)
        {
            if (!escaped)
            {
                return raw.Equals(name, StringComparison.OrdinalIgnoreCase);
            }

            int j = 0;
            for (int i = 0; i < raw.Length; i++, j++)
            {
                if (raw[i] == '\'')
                {
                    i++;
                }

                if (j >= name.Length || char.ToUpperInvariant(raw[i]) != char.ToUpperInvariant(name[j]))
                {
                    return false;
                }
            }

            return j == name.Length;
        }

        private static bool TryParseReference(ReadOnlySpan<char> text, out ushort rowWord, out byte column)
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

            if (i >= text.Length || !int.TryParse(text[i..], NumberStyles.None, CultureInfo.InvariantCulture, out int row))
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

            rowWord = RowWord.Create(row - 1, rowRelative, columnRelative);
            column = (byte)(col - 1);
            return true;
        }

        // Encodes string text as Mac OS Roman, a doubled quote standing for one.
        private static void AddMacRoman(ref Bytes output, ReadOnlySpan<char> text, bool escaped)
        {
            for (int i = 0; i < text.Length; i++)
            {
                output.Add(MacRoman.GetByte(text[i]));
                if (escaped && text[i] == '"')
                {
                    i++;
                }
            }
        }

        // ----- Emitter -----

        private static byte WithClass(byte baseCode, int tokenClass) => (byte)((baseCode & 0x1F) | (tokenClass << 5));

        // In a name definition, operator operands and arguments that accept arrays take array class where they would
        // otherwise take value class. In an array formula, arguments that accept arrays do, and operator operands only
        // where the operator is itself in such an argument: {=A1:A2*2} keeps value class in Excel 1.00 and 1.5, while
        // {=SUM(A1:A2*B1:B2)} takes array class (spec §15.4). The formula as a whole keeps the class the other rules give.
        private readonly int InArrayFormula(Context context, int tokenClass)
            => tokenClass == ValueClass
                && (context.Kind == ContextKind.Operator ? _definition || (_arrayFormula && context.InArrayArgument) : (_arrayFormula || _definition) && AcceptsArrays(context))
                ? ArrayClass
                : tokenClass;

        private static bool AcceptsArrays(Context context)
            => context.Kind is ContextKind.Argument or ContextKind.IfArgument && (context.Accepts & ExcelValueTypes.Array) != 0;

        private static Context OperatorContext(Context parent)
            => Context.Operator with { InArrayArgument = parent.Kind == ContextKind.Operator ? parent.InArrayArgument : AcceptsArrays(parent) };

        private readonly int OperandClass(Context context) => InArrayFormula(context, context.Kind switch
        {
            ContextKind.DefinitionTop or ContextKind.Range or ContextKind.MacroTarget => ReferenceClass,
            ContextKind.Argument or ContextKind.IfArgument => (context.Accepts & ExcelValueTypes.Reference) != 0 ? ReferenceClass : ValueClass,
            _ => ValueClass,
        });

        // The arguments of IF and CHOOSE follow the same rules as any other in every release; the 1.00 disk's macro
        // sheets carry reference-class calls from an earlier build's compiler (spec §15.4).
        private readonly int FunctionClass(Context context, ExcelValueTypes returns) => InArrayFormula(context, context.Kind switch
        {
            ContextKind.Range or ContextKind.MacroTarget or ContextKind.DefinitionTop => ReferenceClass,
            ContextKind.Argument or ContextKind.IfArgument when returns == ExcelValueTypes.Array && (context.Accepts & ExcelValueTypes.Array) != 0 => ArrayClass,
            ContextKind.Argument or ContextKind.IfArgument when (context.Accepts & ExcelValueTypes.Reference) != 0 && (returns & ExcelValueTypes.Reference) != 0 => ReferenceClass,
            _ => ValueClass,
        });

        private void Emit(int index, Context context)
        {
            // A copy: emitting the children does not move the tree, but reading fields from a copy keeps this simple.
            var node = _nodes[index];
            switch (node.Kind)
            {
                case NodeKind.Number:
                    if (!node.Flag && node.Number >= 0 && node.Number <= ushort.MaxValue && node.Number == Math.Floor(node.Number))
                    {
                        Expression.Add(0x1E);
                        Expression.AddWord((ushort)node.Number);
                    }
                    else
                    {
                        Expression.Add(0x1F);
                        Expression.AddDouble(node.Number);
                    }

                    break;
                case NodeKind.String:
                {
                    var text = _text.Slice(node.A, node.B);
                    Expression.Add(0x17);
                    Expression.Add((byte)(node.Flag ? node.B - text.Count("\"\"") : node.B));
                    AddMacRoman(ref Expression, text, node.Flag);
                    break;
                }

                case NodeKind.Bool:
                    Expression.Add(0x1D);
                    Expression.Add(node.Code);
                    break;
                case NodeKind.Error when node.Code == (byte)ExcelError.Reference:
                    // #REF! refers to deleted cells: a deleted-reference token, whose row and column bytes are no
                    // longer meaningful. Excel 1.03 compiles a typed =#REF! to 4A 00 00 00, not an error constant.
                    Expression.Add(WithClass(0x2A, OperandClass(context)));
                    Expression.Add(0);
                    Expression.Add(0);
                    Expression.Add(0);
                    break;
                case NodeKind.Error:
                    Expression.Add(0x1C);
                    Expression.Add(node.Code);
                    break;
                case NodeKind.Missing:
                    Expression.Add(0x16);
                    break;
                case NodeKind.Array:
                {
                    // 0x20: column count, row count, three bytes left over from memory. The data goes in the fix-ups of
                    // a cell formula, or after the names for a name definition. Array class in a definition and where an
                    // argument accepts arrays; elsewhere the class of a value, so ={1,2} in a cell is 40 (spec §15.4).
                    var block = _arrays.Slice(node.A, node.B);
                    Expression.Add(WithClass(0x20, _definition || (context.Kind is ContextKind.Argument or ContextKind.IfArgument && (context.Accepts & ExcelValueTypes.Array) != 0)
                        ? ArrayClass
                        : InArrayFormula(context, ValueClass)));
                    Expression.Add(block[0]);
                    Expression.Add(block[1]);
                    Expression.Add(block[2]);
                    Expression.Add(0);
                    Expression.Add(0);
                    Expression.Add(0);
                    if (_definition)
                    {
                        (ArrayConstants ??= []).Add(block.ToArray());
                    }
                    else
                    {
                        Fixups.AddRange(block);
                    }

                    break;
                }

                case NodeKind.Ref:
                    BeginExternal(node.External);
                    Expression.Add(WithClass(_definition ? (byte)0x2C : (byte)0x24, OperandClass(context)));
                    Expression.AddWord(node.Row1);
                    Expression.Add(node.Column1);
                    EndExternal(node.External);
                    break;
                case NodeKind.Area:
                    BeginExternal(node.External);
                    Expression.Add(WithClass(_definition ? (byte)0x2D : (byte)0x25, OperandClass(context)));
                    Expression.AddWord(node.Row1);
                    Expression.AddWord(node.Row2);
                    Expression.Add(node.Column1);
                    Expression.Add(node.Column2);
                    EndExternal(node.External);
                    break;
                case NodeKind.Name:
                    BeginExternal(node.External);
                    EmitName(node.A, OperandClass(context));
                    EndExternal(node.External);
                    break;
                case NodeKind.Unary:
                    Emit(node.A, OperatorContext(context));
                    Expression.Add(node.Code);
                    break;
                case NodeKind.Binary:
                    EmitBinary(node, context);
                    break;
                case NodeKind.Paren:
                    Emit(node.A, context);
                    Expression.Add(0x15);
                    break;
                case NodeKind.Function when node.A == 1:
                    EmitIf(node, context);
                    break;
                case NodeKind.Function when node.A == 180 && !node.Flag:
                    EmitChoose(node, context);
                    break;
                case NodeKind.Function:
                    EmitCall(node, context);
                    break;
            }
        }

        private void EmitBinary(Node binary, Context context)
        {
            bool reference = binary.Code is IntersectionOperator or UnionOperator or RangeOperator;
            var operandContext = reference ? Context.Range : OperatorContext(context);

            // The outermost union or intersection, or range whose operands are not both plain references, is
            // wrapped in a memory token giving the length of what follows. Built from cell references alone, it is a
            // memory area (0x26: three bytes left over from memory, then the length), and Excel caches the resulting
            // areas in the fix-ups; otherwise a memory function (0x29; 0x2E in a definition). Both take the class of
            // a reference in their place, as Excel 1.03's own compilation shows.
            int start = -1;
            if (reference && !_insideMemoryFunction
                && (binary.Code != RangeOperator || !(_nodes[binary.A].Kind == NodeKind.Ref && _nodes[binary.B].Kind == NodeKind.Ref)))
            {
                int memoryClass = OperandClass(context);
                var areas = new AreaList(stackalloc CachedArea[16]);
                try
                {
                    bool cached = !_definition && CollectAreas(ref areas, binary);
                    if (cached && (areas.Count == 0 || (areas.Count > 1 && memoryClass == ValueClass)))
                    {
                        // No areas, as for an intersection of cells that do not overlap, or more than one where a
                        // single value is needed: Excel compiles a memory error (0x27) instead, with no cached areas:
                        // #NULL! (0x00) for no areas, in any class (=SUM((A1 B2):C3), `Empty Intersection Ranges`),
                        // and #VALUE! (0x0F) for several (spec §15.2).
                        Expression.Add(WithClass(0x27, memoryClass));
                        Expression.Add(0);
                        Expression.Add(areas.Count == 0 ? (byte)0x00 : (byte)0x0F);
                        Expression.Add(0);
                        Expression.Add(0);
                    }
                    else if (cached)
                    {
                        Expression.Add(WithClass(0x26, memoryClass));
                        Expression.Add(0);
                        Expression.Add(0);
                        Expression.Add(0);
                        Expression.Add(0);
                        Fixups.AddWord(areas.Count);
                        foreach (var area in areas.AsSpan())
                        {
                            Fixups.AddWord(area.FirstRow);
                            Fixups.AddWord(area.LastRow);
                            Fixups.Add(area.FirstColumn);
                            Fixups.Add(area.LastColumn);
                        }
                    }
                    else
                    {
                        Expression.Add(_definition ? (byte)0x2E : WithClass(0x29, memoryClass));
                        Expression.Add(0);
                    }
                }
                finally
                {
                    areas.Dispose();
                }

                start = Expression.Count;
                _insideMemoryFunction = true;
            }

            Emit(binary.A, operandContext);
            Emit(binary.B, operandContext);
            Expression.Add(binary.Code);
            if (start >= 0)
            {
                _insideMemoryFunction = false;
                int length = Expression.Count - start;
                if (length > byte.MaxValue)
                {
                    throw new ExcelFormulaSyntaxException("The range expression is too long", binary.Position);
                }

                Expression[start - 1] = (byte)length;
            }
        }

        // Adds the areas a reference expression built only from cell references on this sheet stands for, as Excel
        // caches them after a memory-area token, or returns false if it involves anything else (names, functions,
        // other sheets).
        private readonly bool CollectAreas(ref AreaList areas, in Node node)
        {
            switch (node.Kind)
            {
                case NodeKind.Ref when node.External == 0:
                    areas.Add(new CachedArea((ushort)RowWord.Row(node.Row1), (ushort)RowWord.Row(node.Row1), node.Column1, node.Column1));
                    return true;
                case NodeKind.Area when node.External == 0:
                    areas.Add(new CachedArea((ushort)RowWord.Row(node.Row1), (ushort)RowWord.Row(node.Row2), node.Column1, node.Column2));
                    return true;
                case NodeKind.Paren:
                    return CollectAreas(ref areas, _nodes[node.A]);
                case NodeKind.Binary when node.Code is IntersectionOperator or UnionOperator or RangeOperator:
                {
                    int start = areas.Count;
                    if (!CollectAreas(ref areas, _nodes[node.A]))
                    {
                        return false;
                    }

                    int middle = areas.Count;
                    if (!CollectAreas(ref areas, _nodes[node.B]))
                    {
                        return false;
                    }

                    int end = areas.Count;
                    if (node.Code == UnionOperator)
                    {
                        // A side with no areas makes the union empty too: =D30 E30,F32 is a memory error, #NULL!
                        // (`Typed Unions`, Excel 1.00).
                        if (middle == start || end == middle)
                        {
                            areas.Truncate(start);
                        }

                        return true;
                    }

                    if (node.Code == RangeOperator)
                    {
                        // The range of both sides: the box around every area of either. A side with no areas, an
                        // intersection of cells that do not overlap, makes the range empty too (`Empty Intersection
                        // Ranges`, Excel 1.00).
                        var all = areas.AsSpan()[start..end];
                        if (middle == start || end == middle)
                        {
                            areas.Truncate(start);
                            return true;
                        }

                        ushort firstRow = ushort.MaxValue, lastRow = 0;
                        byte firstColumn = byte.MaxValue, lastColumn = 0;
                        foreach (var area in all)
                        {
                            firstRow = Math.Min(firstRow, area.FirstRow);
                            lastRow = Math.Max(lastRow, area.LastRow);
                            firstColumn = Math.Min(firstColumn, area.FirstColumn);
                            lastColumn = Math.Max(lastColumn, area.LastColumn);
                        }

                        areas.Truncate(start);
                        areas.Add(new CachedArea(firstRow, lastRow, firstColumn, lastColumn));
                        return true;
                    }

                    // The intersection: every overlap of an area of the left with one of the right, in that order.
                    for (int i = start; i < middle; i++)
                    {
                        for (int k = middle; k < end; k++)
                        {
                            var x = areas[i];
                            var y = areas[k];
                            ushort firstRow = Math.Max(x.FirstRow, y.FirstRow), lastRow = Math.Min(x.LastRow, y.LastRow);
                            byte firstColumn = Math.Max(x.FirstColumn, y.FirstColumn), lastColumn = Math.Min(x.LastColumn, y.LastColumn);
                            if (firstRow <= lastRow && firstColumn <= lastColumn)
                            {
                                areas.Add(new CachedArea(firstRow, lastRow, firstColumn, lastColumn));
                            }
                        }
                    }

                    areas.RemoveRange(start, end - start);
                    return true;
                }

                default:
                    return false;
            }
        }

        private void EmitName(int index, int tokenClass)
        {
            Expression.Add(WithClass(0x23, tokenClass));
            if (_definition)
            {
                Expression.AddWord(index);
                for (int i = 0; i < 5; i++)
                {
                    Expression.Add(0);
                }
            }
            else
            {
                for (int i = 0; i < 7; i++)
                {
                    Expression.Add(0);
                }

                Fixups.AddWord(index);
            }
        }

        // An operand in another document: 1A, the operand, 1B.
        private void BeginExternal(int external)
        {
            if (external == 0)
            {
                return;
            }

            // 1A: four bytes left over from memory, 00 00, then the 1-based document index (which cell formulas
            // also carry in the fix-ups).
            Expression.Add(0x1A);
            for (int i = 0; i < 6; i++)
            {
                Expression.Add(0);
            }

            Expression.Add(_definition ? (byte)external : (byte)0);
            if (!_definition)
            {
                Fixups.AddWord(external);
            }
        }

        private void EndExternal(int external)
        {
            if (external == 0)
            {
                return;
            }

            Expression.Add(0x1B);
            Expression.Add(0);
            Expression.Add(0);
            Expression.Add(0);
        }

        private void EmitCall(Node function, Context context)
        {
            int count = function.C;
            if (function.D >= 0)
            {
                // A call to a function macro: its name, the arguments, then function 255 with one more argument.
                var macro = _nodes[function.D];
                BeginExternal(macro.External);
                EmitName(macro.A, ReferenceClass);
                EndExternal(macro.External);
                for (int i = 0; i < count; i++)
                {
                    Emit(_arguments[function.B + i], new Context(ContextKind.Argument, (ExcelValueTypes)0x7F));
                }

                Expression.Add(WithClass(0x22, FunctionClass(context, (ExcelValueTypes)0x7F)));
                Expression.Add((byte)(count + 1));
                Expression.Add(255);
                return;
            }

            int number = function.A;
            bool dialog = function.Flag;
            var definition = ExcelFunctions.Get(number, _version)!;
            for (int i = 0; i < count; i++)
            {
                Emit(_arguments[function.B + i], new Context(ContextKind.Argument, definition.GetArgumentType(i)));
            }

            // Excel compiles SUM with exactly one argument to a SUM attribute instead of a function token.
            if (number == 4 && count == 1 && !dialog)
            {
                Expression.Add(0x19);
                Expression.Add(0x10);
                Expression.Add(0);
                return;
            }

            int tokenClass = FunctionClass(context, definition.ReturnType);

            // Functions numbered 256 and up use 0x39 and 0x3A, holding the number minus 256 (Excel 1.5's Save).
            int high = number > 255 ? 0x18 : 0;
            // Excel 1.5's Save writes every call of a fixed-count function in the fixed form, dropping a dialog
            // form's bit (47.19 in 1.5): DATA.DELETE?() is 41 82, where 1.00's compiler keeps 42 80 82.
            if (dialog && !(definition.HasFixedArguments && _version >= ExcelVersion.Excel150))
            {
                Expression.Add(WithClass((byte)(0x22 + high), tokenClass));
                Expression.Add((byte)(count | 0x80));
            }
            else if (definition.HasFixedArguments || WritesOldForm(number, count, definition))
            {
                Expression.Add(WithClass((byte)(0x21 + high), tokenClass));
            }
            else
            {
                Expression.Add(WithClass((byte)(0x22 + high), tokenClass));
                Expression.Add((byte)count);
            }

            Expression.Add((byte)number);
        }

        // Excel 1.5 gave FORMULA, ACTIVATE, CLOSE, VSCROLL, HSCROLL, FORMULA.ARRAY and FORMULA.FILL an optional
        // argument. Its Save still writes a call with the minimum number of arguments in the fixed-count form, so
        // that older releases can read it (spec §15.4).
        private readonly bool WritesOldForm(int number, int count, ExcelFunction function)
            => _version >= ExcelVersion.Excel150
                && number is 57 or 117 or 118 or 155 or 192 or 195 or 196
                && count == function.MinArguments;

        // IF(condition, true[, false]): condition, a jump over the true branch, the true branch and a skip past the
        // IF token, then the false branch and its skip, then the IF token itself.
        private void EmitIf(Node function, Context context)
        {
            var definition = ExcelFunctions.Get(1, _version)!;
            int count = function.C;
            Emit(_arguments[function.B], new Context(ContextKind.IfArgument, definition.GetArgumentType(0)));
            Expression.Add(0x19);
            Expression.Add(0x02);
            int jump = Expression.Count;
            Expression.Add(0);
            int afterJump = Expression.Count;

            Span<int> skips = stackalloc int[3];
            for (int i = 1; i < count; i++)
            {
                if (i == 2)
                {
                    SetOffset(jump, Expression.Count - afterJump, function.Position);
                }

                Emit(_arguments[function.B + i], new Context(ContextKind.IfArgument, definition.GetArgumentType(i)));
                Expression.Add(0x19);
                Expression.Add(0x08);
                skips[i - 1] = Expression.Count;
                Expression.Add(0);
            }

            if (count == 2)
            {
                SetOffset(jump, Expression.Count - afterJump, function.Position);
            }

            Expression.Add(WithClass(0x22, FunctionClass(context, definition.ReturnType)));
            Expression.Add((byte)count);
            Expression.Add(1);
            int end = Expression.Count;
            foreach (int skip in skips[..(count - 1)])
            {
                SetOffset(skip, end - (skip + 1) - 1, function.Position);
            }
        }

        // CHOOSE(index, choice1, …): the index, then 19 04 n and a table of n + 1 byte offsets, then each choice
        // followed by a skip past the CHOOSE token, then the CHOOSE token. Offsets in the table are measured from its
        // first byte: entry k is the start of choice k + 1, and the last entry is the CHOOSE token.
        private void EmitChoose(Node function, Context context)
        {
            var definition = ExcelFunctions.Get(180, _version)!;
            int count = function.C;
            Emit(_arguments[function.B], new Context(ContextKind.IfArgument, definition.GetArgumentType(0)));
            Expression.Add(0x19);
            Expression.Add(0x04);
            Expression.Add((byte)(count - 1));
            int table = Expression.Count;
            for (int i = 0; i < count; i++)
            {
                Expression.Add(0);
            }

            int[]? rented = count > 32 ? ArrayPool<int>.Shared.Rent(count) : null;
            Span<int> skips = rented ?? stackalloc int[32];
            try
            {
                for (int i = 1; i < count; i++)
                {
                    SetOffset(table + i - 1, Expression.Count - table, function.Position);
                    Emit(_arguments[function.B + i], new Context(ContextKind.IfArgument, definition.GetArgumentType(i)));
                    Expression.Add(0x19);
                    Expression.Add(0x08);
                    skips[i - 1] = Expression.Count;
                    Expression.Add(0);
                }

                SetOffset(table + count - 1, Expression.Count - table, function.Position);
                Expression.Add(WithClass(0x22, FunctionClass(context, definition.ReturnType)));
                Expression.Add((byte)count);
                Expression.Add(180);
                int end = Expression.Count;
                foreach (int skip in skips[..(count - 1)])
                {
                    SetOffset(skip, end - (skip + 1) - 1, function.Position);
                }
            }
            finally
            {
                if (rented is not null)
                {
                    ArrayPool<int>.Shared.Return(rented);
                }
            }
        }

        private void SetOffset(int index, int value, int position)
        {
            if (value is < 0 or > 255)
            {
                throw new ExcelFormulaSyntaxException("An IF or CHOOSE branch is too long to encode (over 255 bytes)", position);
            }

            Expression[index] = (byte)value;
        }
    }

    // A list of cached areas that starts on the stack and moves to a pooled array if it grows.
    private ref struct AreaList
    {
        private Span<CachedArea> _items;
        private CachedArea[]? _rented;

        public AreaList(Span<CachedArea> initial)
        {
            _items = initial;
            _rented = null;
            Count = 0;
        }

        public int Count { get; private set; }

        public readonly CachedArea this[int index] => _items[index];

        public readonly ReadOnlySpan<CachedArea> AsSpan() => _items[..Count];

        public void Add(CachedArea area)
        {
            if (Count == _items.Length)
            {
                var larger = ArrayPool<CachedArea>.Shared.Rent(_items.Length * 2);
                _items[..Count].CopyTo(larger);
                if (_rented is not null)
                {
                    ArrayPool<CachedArea>.Shared.Return(_rented);
                }

                _items = _rented = larger;
            }

            _items[Count++] = area;
        }

        public void Truncate(int count) => Count = count;

        public void RemoveRange(int start, int count)
        {
            _items[(start + count)..Count].CopyTo(_items[start..]);
            Count -= count;
        }

        public void Dispose()
        {
            if (_rented is not null)
            {
                ArrayPool<CachedArea>.Shared.Return(_rented);
                _rented = null;
            }
        }
    }
}
