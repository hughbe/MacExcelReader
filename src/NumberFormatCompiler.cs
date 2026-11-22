namespace MacExcelReader;

/// <summary>
/// Compiles number-format text into tokens as Excel does (spec §7): Excel 1.00's section compiler (<c>07.09</c>) and
/// assembler (<c>07.10</c>), or Excel 1.5's (<c>0d.0b</c> and <c>0d.12</c>). The body is built in one stack buffer.
/// </summary>
internal static class NumberFormatCompiler
{
    private const int MaxBody = 255;

    // Room past MaxBody for the bytes a section adds after its last length check (header, pad, colour).
    private const int Slack = 16;


    /// <summary>Compiles a whole format; <paramref name="textSection"/> asks for Excel 1.5's four-slot form.</summary>
    public static ExcelNumberFormat Compile(string text, ExcelVersion version, bool textSection)
    {
        ArgumentNullException.ThrowIfNull(text);
        Span<byte> body = stackalloc byte[MaxBody + Slack];
        return version >= ExcelVersion.Excel150
            ? CompileExcel15(text, body, textSection)
            : CompileExcel100(text, body);
    }

    // 07.10 in Excel 1.00: up to three sections, stored positive, negative, zero. With one section the others share
    // its bytes and the single-section byte is 1; with two, zero shares the positive section's bytes.
    private static ExcelNumberFormat CompileExcel100(string text, Span<byte> body)
    {
        int position = 0;
        int p = new SectionCompiler(text, body, 0, excel15: false, flagsIn: 0).Compile(ref position, out _);
        if (position < 0)
        {
            return Create([0, 0, 0, (byte)p, (byte)p, (byte)p, 1], body[..p]);
        }

        int n = new SectionCompiler(text, body, p, excel15: false, flagsIn: 0).Compile(ref position, out _);
        if (position < 0)
        {
            return Create([(byte)p, 0, 0, (byte)n, (byte)p, (byte)p, 0], body[..(p + n)]);
        }

        int z = new SectionCompiler(text, body, p + n, excel15: false, flagsIn: 0).Compile(ref position, out _);
        if (position >= 0)
        {
            throw Error(text, position - 1, "A number format has at most three sections in Excel 1.0x.");
        }

        return Create([(byte)p, (byte)(p + n), 0, (byte)n, (byte)z, (byte)p, 0], body[..(p + n + z)]);
    }

    // 0d.12 in Excel 1.5: up to four sections, positive, negative, zero, text. A section that uses @ is the text
    // section and ends the format; a positive section that is only General (and a colour) doubles as it; without
    // either, General is appended as the text section. A version-17 file keeps the first three slots (47.11).
    private static ExcelNumberFormat CompileExcel15(string text, Span<byte> body, bool textSection)
    {
        // Slots in the descriptor's order: negative, zero, positive, text.
        Span<int> offset = stackalloc int[4];
        Span<int> length = stackalloc int[4];
        ReadOnlySpan<int> slots = [2, 0, 1, 3];
        bool single = false, anyNumeric = false, textSet = false;
        int used = 0, position = 0;
        for (int i = 0; i < slots.Length && position >= 0; i++)
        {
            int slot = slots[i];
            int start = used;
            int size = new SectionCompiler(text, body, start, excel15: true, flagsIn: slot == 3 ? 2 : 0).Compile(ref position, out int flags);
            used += size;
            if ((flags & (slot == 2 ? 3 : 2)) != 0)
            {
                if (textSection)
                {
                    textSet = true;
                    offset[3] = start;
                    length[3] = size;
                }

                if ((flags & 2) != 0)
                {
                    break;
                }
            }

            anyNumeric = true;
            offset[slot] = start;
            length[slot] = size;
            single = slot == 2;
            if (slot == 2)
            {
                offset[0] = offset[1] = start;
                length[0] = length[1] = size;
            }
        }

        if (position >= 0)
        {
            throw Error(text, position - 1, "A number format has at most four sections, and none after the one with @.");
        }

        if (!(anyNumeric && textSet))
        {
            int general = used;
            body[used++] = 0x01;
            if (textSet)
            {
                // Only a text section: the number sections are General.
                offset[0] = offset[1] = offset[2] = general;
                length[0] = length[1] = length[2] = 1;
                single = true;
            }
            else
            {
                offset[3] = general;
                length[3] = 1;
            }
        }

        byte flag = single ? (byte)1 : (byte)0;
        if (!textSection)
        {
            int end = Math.Max(offset[0] + length[0], Math.Max(offset[1] + length[1], offset[2] + length[2]));
            return Create(
                [(byte)offset[0], (byte)offset[1], (byte)offset[2], (byte)length[0], (byte)length[1], (byte)length[2], flag],
                body[..end]);
        }

        return Create(
            [(byte)offset[0], (byte)offset[1], (byte)offset[2], (byte)offset[3], (byte)length[0], (byte)length[1], (byte)length[2], (byte)length[3], flag],
            body[..used]);
    }

    private static ExcelNumberFormat Create(ReadOnlySpan<byte> descriptor, ReadOnlySpan<byte> body)
        => body.Length > MaxBody
            ? throw new FormatException("A number format compiles to at most 255 bytes.")
            : new ExcelNumberFormat(descriptor.ToArray(), body.ToArray());

    private static FormatException Error(string text, int position, string message)
        => new($"{message} (at position {Math.Max(0, position)} of \"{text}\")");

    // Excel 1.5's keywords, checked before each character: 0 General, 1 AM/PM, 2 A/P, 3 the currency symbol.
    private static int Keyword(ReadOnlySpan<char> rest)
        => rest.StartsWith("General", StringComparison.OrdinalIgnoreCase) ? 0
            : rest.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase) ? 1
            : rest.StartsWith("A/P", StringComparison.OrdinalIgnoreCase) ? 2
            : rest[0] == '$' ? 3
            : -1;

    private static int ColorIndex(ReadOnlySpan<char> name)
    {
        for (int i = 0; i < ExcelNumberFormat.ColorNames.Length; i++)
        {
            if (name.Equals(ExcelNumberFormat.ColorNames[i], StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// One section, from a position to the next ';' (07.09 in 1.00, 0d.0b in 1.5), written into the body at its
    /// start. Insertions (the number header, the thousands token, a colour) shift the section's bytes in place.
    /// </summary>
    private ref struct SectionCompiler
    {
        private readonly string _text;
        private readonly Span<byte> _section;
        private readonly int _start;
        private readonly bool _excel15;
        private int _count;
        private int _i;
        private int _header;
        private int _firstDigit;
        private int _exponent;
        private int _lastHour;
        private int _integers;
        private int _decimals;
        private int _exponentDigits;
        private int _extra;
        private int _color;
        private bool _thousands;
        private bool _afterPoint;
        private bool _general;
        private bool _date;
        private bool _fill;
        private bool _hourPending;
        private bool _at;
        private bool _numeric;
        private bool _percent;
        private bool _scaling;
        private bool _generalLike;

        public SectionCompiler(string text, Span<byte> body, int start, bool excel15, int flagsIn)
        {
            _text = text;
            _section = body[start..];
            _start = start;
            _excel15 = excel15;
            _header = _firstDigit = _exponent = _lastHour = _color = -1;
            _at = (flagsIn & 2) != 0;
            _numeric = (flagsIn & 4) != 0;
            _generalLike = !_numeric && !_at;
        }

        private readonly char Peek(int offset = 0) => _i + offset < _text.Length ? _text[_i + offset] : '\0';

        /// <summary>
        /// Compiles the section and returns its length. <paramref name="position"/> becomes the index after the
        /// ';', or -1 at the end. <paramref name="flags"/> (Excel 1.5): 1 the section is only General and colour,
        /// 2 it uses @, 4 it has numbers, dates or AM/PM.
        /// </summary>
        public int Compile(ref int position, out int flags)
        {
            _i = position;
            while (_i < _text.Length && _text[_i] != ';')
            {
                if (_start + _count > MaxBody - (_excel15 ? 8 : 6))
                {
                    throw Fail("The number format is too long.");
                }

                if (_excel15 && CompileKeyword())
                {
                    continue;
                }

                char ch = _text[_i++];
                _generalLike &= ch == '[';
                CompileCharacter(ch);
            }

            Finish();
            position = _i < _text.Length ? _i + 1 : -1;
            flags = _generalLike && _count > 0 ? 1 : _numeric ? 4 : _at ? 2 : 0;
            return _count;
        }

        private readonly FormatException Fail(string message) => Error(_text, _i - 1, message);

        private void Emit(byte value) => _section[_count++] = value;

        private void Emit(byte first, byte second)
        {
            _section[_count++] = first;
            _section[_count++] = second;
        }

        private void Literal(char c) => Emit(0x02, MacRoman.GetByte(c));

        private void Insert(int at, scoped ReadOnlySpan<byte> bytes)
        {
            _section[at.._count].CopyTo(_section[(at + bytes.Length)..]);
            bytes.CopyTo(_section[at..]);
            _count += bytes.Length;
        }

        private bool CompileKeyword()
        {
            int keyword = Keyword(_text.AsSpan(_i));
            if (keyword < 0)
            {
                return false;
            }

            _generalLike &= keyword == 0;
            switch (keyword)
            {
                case 0:
                    if (_general || _header >= 0 || _date)
                    {
                        throw Fail("General cannot be combined with numbers or dates.");
                    }

                    Emit(0x01);
                    _general = true;
                    _i += "General".Length;
                    break;
                case 1 or 2:
                    if (_at)
                    {
                        throw Fail("AM/PM cannot be used with @.");
                    }

                    _numeric = true;
                    AmPm(keyword == 1 ? 2 : 1);
                    break;
                default:
                    if (_at)
                    {
                        throw Fail("A currency symbol cannot be used with @.");
                    }

                    _numeric = true;
                    Emit(0x22);
                    _i++;
                    break;
            }

            return true;
        }

        private void CompileCharacter(char ch)
        {
            switch (ch)
            {
                case '#' or '0':
                    Digits(ch);
                    return;
                case '.':
                    if (_at)
                    {
                        throw Fail("A decimal point cannot be used with @.");
                    }

                    _numeric = true;
                    Emit(0x05);
                    _afterPoint = true;
                    return;
                case '%':
                    if (_excel15 && (_at || _scaling))
                    {
                        throw Fail("% cannot be combined with @ or with scaling.");
                    }

                    _percent = true;
                    _extra += 2;
                    Literal('%');
                    return;
                case '\\':
                    if (_i >= _text.Length)
                    {
                        throw Fail("A backslash needs a character after it.");
                    }

                    Literal(_text[_i++]);
                    return;
                case '*':
                    if (_fill || _i >= _text.Length)
                    {
                        throw Fail("A section has at most one fill, and * needs a character after it.");
                    }

                    Emit(0x0A, MacRoman.GetByte(_text[_i++]));
                    _fill = true;
                    return;
                case '"':
                    QuotedString();
                    return;
                case '$' when !_excel15:
                    Emit(0x22);
                    return;
                case '[' when _excel15:
                    Color();
                    return;
                case '@' when _excel15:
                    if (_numeric)
                    {
                        throw Fail("@ cannot be combined with numbers or dates.");
                    }

                    _at = true;
                    Emit(0x23);
                    return;
            }

            char lower = char.ToLowerInvariant(ch);
            switch (lower)
            {
                case 'e':
                    Exponent(ch);
                    return;
                case 'd':
                case 'm' when !_hourPending:
                    DayOrMonth(lower);
                    return;
                case 'h' or 'm' or 's':
                    HourMinuteSecond(lower);
                    return;
                case 'y':
                    Year();
                    return;
            }

            // Excel 1.5 keeps any other character as a literal; Excel 1.00 only its separators, General and AM/PM.
            if (_excel15 || ch is ' ' or '(' or ')' or '+' or ',' or '-' or '/' or ':')
            {
                Literal(ch);
                return;
            }

            var rest = _text.AsSpan(_i - 1);
            if (!_general && _header < 0 && !_date && rest.StartsWith("General", StringComparison.OrdinalIgnoreCase))
            {
                Emit(0x01);
                _general = true;
                _i += "General".Length - 1;
            }
            else if (rest.StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
            {
                _i--;
                AmPm(1);
            }
            else if (rest.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
            {
                _i--;
                AmPm(2);
            }
            else
            {
                throw Fail($"'{ch}' is not part of a number format in Excel 1.0x; put it in quotes or after a backslash.");
            }
        }

        private void QuotedString()
        {
            int close = _text.IndexOf('"', _i);
            if (close < 0)
            {
                throw Fail("A quoted string has no closing quote.");
            }

            var quoted = _text.AsSpan(_i, close - _i);
            if (_excel15 && quoted is "$")
            {
                Emit(0x22);
            }
            else
            {
                if (_start + _count + 2 + quoted.Length > MaxBody)
                {
                    throw Fail("The number format is too long.");
                }

                Emit(0x09, (byte)quoted.Length);
                foreach (char c in quoted)
                {
                    Emit(MacRoman.GetByte(c));
                }
            }

            _i = close + 1;
        }

        private void Color()
        {
            int end = _text.IndexOf(']', _i);
            int index = end < 0 ? -1 : ColorIndex(_text.AsSpan(_i, end - _i));
            if (index < 0 || _color >= 0)
            {
                throw Fail("A colour is [black], [white], [red], [green], [blue], [yellow], [magenta] or [cyan], once per section.");
            }

            _color = index;
            _i = end + 1;
        }

        private void Digits(char c)
        {
            if (_date || _general || _at)
            {
                throw Fail("Digits cannot be combined with dates, General or @.");
            }

            _numeric = true;
            int count = 1;
            bool comma = false;
            for (; _i < _text.Length; _i++)
            {
                char next = _text[_i];
                if (next == c)
                {
                    count++;
                }
                else if (next == ',' && !_afterPoint && (_header < 0 || _exponent < 0)
                    && (!_excel15 || Peek(1) is '#' or '0'))
                {
                    comma = true;
                }
                else
                {
                    break;
                }
            }

            // Excel 1.5: commas after the digits scale the number by 1000 each.
            if (_excel15 && Peek() != '.')
            {
                for (; Peek() == ','; _i++)
                {
                    if (_percent)
                    {
                        throw Fail("Scaling cannot be combined with %.");
                    }

                    _scaling = true;
                    _extra -= 3;
                }
            }

            if (_header < 0)
            {
                Insert(0, [0x03, 0, 0, 0]);
                _header = 0;
                if (comma)
                {
                    Emit(0x07);
                }

                _firstDigit = _count;
                _thousands = comma;
            }
            else if (comma && !_thousands)
            {
                Insert(_header + 4, [0x07]);
                _firstDigit++;
                _thousands = true;
            }

            Emit(c == '0' ? (byte)0x06 : (byte)0x04, (byte)count);
            if (_exponent >= 0)
            {
                _exponentDigits += count;
            }
            else if (_afterPoint)
            {
                _decimals += count;
            }
            else
            {
                _integers += count;
            }
        }

        private void Exponent(char c)
        {
            if (_header < 0 || _exponent >= 0 || Peek() is not ('+' or '-') || _at)
            {
                throw Fail("E needs digits before it and + or - after it.");
            }

            _numeric = true;
            _section[_header] = 0x1C;
            Literal(c);
            Emit(_text[_i++] == '+' ? (byte)0x1D : (byte)0x1E);
            _exponent = _count;
            Emit(0);
        }

        private void DayOrMonth(char c)
        {
            if (_general || _header >= 0 || _at)
            {
                throw Fail("Dates cannot be combined with numbers, General or @.");
            }

            _numeric = true;
            int k = 0;
            while (k < 3 && char.ToLowerInvariant(Peek()) == c)
            {
                _i++;
                k++;
            }

            Emit((byte)((c == 'd' ? 0x0B : 0x0F) + k));
            _date = true;
        }

        private void HourMinuteSecond(char c)
        {
            if (_header >= 0 || _general || _at)
            {
                throw Fail("Times cannot be combined with numbers, General or @.");
            }

            _numeric = true;
            int k = 0;
            if (char.ToLowerInvariant(Peek()) == c)
            {
                k = 1;
                _i++;
            }

            switch (c)
            {
                case 'h':
                    _hourPending = true;
                    _lastHour = _count;
                    Emit((byte)(0x13 + k));
                    break;
                case 'm':
                    _hourPending = false;
                    Emit((byte)(0x15 + k));
                    break;
                default:
                    Emit((byte)(0x17 + k));
                    break;
            }

            _date = true;
        }

        private void Year()
        {
            if (_header >= 0 || _general || _at || char.ToLowerInvariant(Peek()) != 'y')
            {
                throw Fail("A year is yy or yyyy.");
            }

            _numeric = true;
            _i++;
            if (char.ToLowerInvariant(Peek()) == 'y' && char.ToLowerInvariant(Peek(1)) == 'y')
            {
                Emit(0x1A);
                _i += 2;
            }
            else
            {
                Emit(0x19);
            }

            _date = true;
        }

        // AM/PM or A/P at the current position, stored as typed; an hour before it becomes a 12-hour one.
        private void AmPm(int half)
        {
            Emit(0x21);
            if (_lastHour >= 0)
            {
                _section[_lastHour] += 12;
                _lastHour = -1;
            }

            _hourPending = false;
            Emit((byte)half, (byte)half);
            for (int k = 0; k < half; k++)
            {
                Emit(MacRoman.GetByte(_text[_i + k]));
            }

            for (int k = 0; k < half; k++)
            {
                Emit(MacRoman.GetByte(_text[_i + half + 1 + k]));
            }

            _i += 2 * half + 1;
        }

        // The end of the section: fill in the number header, pad a thousands format to four integer positions, and
        // put the colour after the header (or first).
        private void Finish()
        {
            if (_header >= 0 && !_at)
            {
                if (_exponent >= 0)
                {
                    _section[_exponent] = (byte)_exponentDigits;
                }

                if (_thousands && _integers < 4)
                {
                    if (_section[_firstDigit] == 0x04)
                    {
                        _section[_firstDigit + 1] += (byte)(4 - _integers);
                    }
                    else
                    {
                        Insert(_firstDigit, [0x04, (byte)(4 - _integers)]);
                    }

                    _integers = 4;
                }

                _section[_header + 1] = (byte)_integers;
                _section[_header + 2] = (byte)_decimals;
                _section[_header + 3] = (byte)(_decimals + _extra);
            }

            if (_color >= 0)
            {
                if (_count == 0)
                {
                    Emit(0x01);
                }

                Insert(_header >= 0 ? _header + 4 : 0, [0x24, (byte)_color]);
            }
        }
    }
}
