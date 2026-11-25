using System.Globalization;
using System.Text;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>A section of a number format (spec §7).</summary>
public enum ExcelNumberFormatSection
{
    /// <summary>Negative numbers.</summary>
    Negative = 0,

    /// <summary>Zero.</summary>
    Zero = 1,

    /// <summary>Positive numbers.</summary>
    Positive = 2,

    /// <summary>Text: Excel 1.5's fourth section, present only in version-18 files.</summary>
    Text = 3,
}

/// <summary>
/// A tokenised number format: a descriptor giving each section's offset and length within the body, then the body
/// (spec §7). User-defined formats are stored in the document; the 19 built-in ones are in <see cref="BuiltIn"/>.
/// </summary>
/// <param name="Descriptor">The section offsets, the section lengths and the single-section byte: 7 bytes with three
/// sections, or 9 with Excel 1.5's text section.</param>
/// <param name="Body">The tokens.</param>
public sealed record ExcelNumberFormat(byte[] Descriptor, byte[] Body)
{
    /// <summary>
    /// Compares two formats by their bytes: formats that compile to the same descriptor and body are equal.
    /// </summary>
    public bool Equals(ExcelNumberFormat? other)
        => other is not null && Descriptor.AsSpan().SequenceEqual(other.Descriptor) && Body.AsSpan().SequenceEqual(other.Body);

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.AddBytes(Descriptor);
        hash.AddBytes(Body);
        return hash.ToHashCode();
    }

    /// <summary>The number of built-in formats, numbered 0–18; a document's own formats are numbered from 19.</summary>
    public const int BuiltInCount = 19;

    // Excel 1.5's colour names (STR# 263), in the order of their numbers in a colour token (spec §7).
    internal static readonly string[] ColorNames = ["black", "white", "red", "green", "blue", "yellow", "magenta", "cyan"];

    // The bodies of Excel 1.00's FRMT 256 resource, in index order.
    private static readonly ExcelNumberFormat[] s_builtIn =
    [
        Create("00000001010101", "01"),
        Create("00000006060601", "030100000601"),
        Create("00000009090901", "030102020601050602"),
        Create("00000009090901", "030400000704030601"),
        Create("0000000C0C0C01", "030402020704030601050602"),
        Create("0C00000E0C0C00", "0304000022070403060102200304000002282207040306010229"),
        Create("0F0000110F0F00", "0304020222070403060105060202200304020202282207040306010506020229"),
        Create("00000008080801", "0301000206010225"),
        Create("0000000B0B0B01", "0301020406010506020225"),
        Create("0000000F0F0F01", "1C010202060105060202451D020602"),
        Create("00000007070701", "0F022F0B022F19"),
        Create("00000007070701", "0B022D11022D19"),
        Create("00000004040401", "0B022D11"),
        Create("00000004040401", "11022D19"),
        Create("0000000D0D0D01", "1F023A160220210202414D504D"),
        Create("00000010101001", "1F023A16023A180220210202414D504D"),
        Create("00000004040401", "13023A16"),
        Create("00000007070701", "13023A16023A18"),
        Create("0000000D0D0D01", "0F022F0B022F19022013023A16"),
    ];

    /// <summary>
    /// Gets Excel's 19 built-in formats, by index, as Excel 1.00's <c>FRMT</c> 256 resource holds them. Files store
    /// only the index (spec §7).
    /// </summary>
    public static IReadOnlyList<ExcelNumberFormat> BuiltIn => s_builtIn;

    /// <summary>
    /// Compiles format text, such as <c>#,##0.00;(#,##0.00)</c>, as the given release's Format Number command does
    /// (spec §7). Releases before 1.5 follow Excel 1.00's compiler and give the three-section form of a version-17
    /// file; <see cref="ExcelVersion.Excel150"/> follows Excel 1.5's, which adds text sections, <c>@</c>, colours
    /// and scaling commas, and gives the four-section form of a version-18 file. Letters, quotes and separators are
    /// the US ones.
    /// </summary>
    /// <exception cref="FormatException">The text is not a format the release accepts.</exception>
    public static ExcelNumberFormat Parse(string text, ExcelVersion version = ExcelVersion.Excel100)
        => NumberFormatCompiler.Compile(text, version, textSection: version >= ExcelVersion.Excel150);

    /// <summary>
    /// Compiles format text for a document of the given format version: as <see cref="Parse(string, ExcelVersion)"/>, but in a version-17
    /// file Excel 1.5's formats keep only their first three sections, as its Save does (spec §7).
    /// </summary>
    internal static ExcelNumberFormat Parse(string text, ExcelVersion version, bool textSection)
        => NumberFormatCompiler.Compile(text, version, textSection && version >= ExcelVersion.Excel150);

    /// <summary>
    /// Gets whether only Excel 1.5 can read the format, so a document holding it must be saved as version 18 (spec
    /// §2.1): it uses <c>@</c> or a colour, scales by trailing commas, or has a text section of its own.
    /// </summary>
    public bool RequiresVersion18
    {
        get
        {
            if (SectionCount == 4 && HasOwnTextSection())
            {
                return true;
            }

            for (int s = 0; s < SectionCount; s++)
            {
                foreach (var token in Tokens((ExcelNumberFormatSection)s))
                {
                    if (token.Code is 0x23 or 0x24 || (token.Code is 0x03 or 0x1C && token.Operands.Length == 3 && HeaderExtra(token.Operands) < 0))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }

    /// <summary>Gets the number of sections: 3, or 4 with Excel 1.5's text section.</summary>
    public int SectionCount => Descriptor.Length == 9 ? 4 : 3;

    /// <summary>
    /// Gets the format in the form a file of the other format version stores (spec §7): with Excel 1.5's text slot
    /// for version 18, where a format without a text section gets General (<c>01</c>) there, or the first three slots
    /// for version 17, as Excel 1.5's Save keeps them.
    /// </summary>
    internal ExcelNumberFormat ForSlotCount(int slots)
    {
        if (slots == SectionCount)
        {
            return this;
        }

        if (slots == 3)
        {
            byte[] three = [Descriptor[0], Descriptor[1], Descriptor[2], Descriptor[4], Descriptor[5], Descriptor[6], Descriptor[8]];
            int length = 0;
            for (int i = 0; i < 3; i++)
            {
                length = Math.Max(length, three[i] + three[3 + i]);
            }

            return new ExcelNumberFormat(three, Body.AsSpan(0, length).ToArray());
        }

        // A positive section that is only General doubles as the text section; otherwise General is appended.
        var positive = GetSection(ExcelNumberFormatSection.Positive);
        bool shared = positive.Length == 1 && positive[0] == 0x01;
        byte textOffset = shared ? Descriptor[2] : checked((byte)Body.Length);
        byte[] four = [Descriptor[0], Descriptor[1], Descriptor[2], textOffset, Descriptor[3], Descriptor[4], Descriptor[5], 1, Descriptor[6]];
        return new ExcelNumberFormat(four, shared ? Body : [.. Body, 0x01]);
    }

    /// <summary>Gets whether the format has a single section: every slot is the same bytes at offset 0.</summary>
    public bool IsSingleSection => Descriptor[^1] != 0;

    /// <summary>
    /// Gets whether the format shows a date: its positive section has a day, month or year token. Hours, minutes,
    /// seconds and AM/PM do not count, as in Excel's own test (<c>07.18</c> in 1.00; spec §7).
    /// </summary>
    public bool IsDate
    {
        get
        {
            foreach (var token in Tokens(ExcelNumberFormatSection.Positive))
            {
                if (token.Code is (>= 0x0B and <= 0x12) or 0x19 or 0x1A)
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>Gets the tokens of one section, or an empty span for a text section the format does not have.</summary>
    public ReadOnlySpan<byte> GetSection(ExcelNumberFormatSection section)
    {
        int count = SectionCount;
        int index = (int)section;
        if (index >= count)
        {
            return [];
        }

        return Body.AsSpan(Descriptor[index], Descriptor[count + index]);
    }

    /// <summary>
    /// Gets the format as text, the way Excel's Format Number dialog shows it (spec §7): the positive section, then
    /// the negative, zero and text sections, separated by <c>;</c>, leaving out those that repeat the positive one.
    /// </summary>
    /// <param name="version">
    /// The release whose rule to follow. Excel 1.00 to 1.06 (<c>07.16</c> in 1.00) write the negative section unless
    /// the format is single-section, and leave out a zero section with the positive section's offset and length,
    /// even an empty one: <c>;;</c> comes back as <c>;</c>. Excel 1.5 (<c>0d.19</c>) leaves out a negative or zero
    /// section that repeats a non-empty positive one, and writes <c>;;</c> back as it is. A format with Excel 1.5's
    /// text section always follows 1.5, and one with only a text section, such as <c>@</c>, is written as that alone.
    /// </param>
    public string ToText(ExcelVersion version = ExcelVersion.Excel100) => ToText(version, export: false);

    // Excel 1.5's export form (the SYLK writer sets g3100 through 79.14) escapes every literal character and quotes
    // the currency symbol: "$"#,##0\ ;\("$"#,##0\).
    internal string ToText(ExcelVersion version, bool export)
    {
        var text = new StringBuilder();
        int count = SectionCount;
        int Offset(ExcelNumberFormatSection s) => Descriptor[(int)s];
        int Length(ExcelNumberFormatSection s) => Descriptor[count + (int)s];
        bool SameBytes(ExcelNumberFormatSection s)
            => Offset(s) == Offset(ExcelNumberFormatSection.Positive) && Length(s) == Length(ExcelNumberFormatSection.Positive);

        // Excel 1.5 writes a format whose text section comes before the General its compiler appended (@ alone,
        // or "pre"@"post") as the text section alone (0d.19).
        if (count == 4 && HasOwnTextSection() && Offset(ExcelNumberFormatSection.Text) < Offset(ExcelNumberFormatSection.Positive))
        {
            AppendSectionText(text, GetSection(ExcelNumberFormatSection.Text), export);
            return text.ToString();
        }

        AppendSectionText(text, GetSection(ExcelNumberFormatSection.Positive), export);
        if (version < ExcelVersion.Excel150 && count == 3)
        {
            if (!IsSingleSection)
            {
                text.Append(';');
                AppendSectionText(text, GetSection(ExcelNumberFormatSection.Negative), export);
                if (!SameBytes(ExcelNumberFormatSection.Zero))
                {
                    text.Append(';');
                    AppendSectionText(text, GetSection(ExcelNumberFormatSection.Zero), export);
                }
            }

            return text.ToString();
        }

        bool SameAsPositive(ExcelNumberFormatSection s) => SameBytes(s) && Length(s) != 0;
        if (!SameAsPositive(ExcelNumberFormatSection.Negative))
        {
            text.Append(';');
            AppendSectionText(text, GetSection(ExcelNumberFormatSection.Negative), export);
            if (!SameAsPositive(ExcelNumberFormatSection.Zero))
            {
                text.Append(';');
                AppendSectionText(text, GetSection(ExcelNumberFormatSection.Zero), export);
            }
        }

        if (count == 4 && HasOwnTextSection())
        {
            text.Append(';');
            AppendSectionText(text, GetSection(ExcelNumberFormatSection.Text), export);
        }

        return text.ToString();
    }

    /// <summary>Formats a number with this format.</summary>
    public string Format(double value) => Format(value, null, TextMeasure.Characters);

    /// <summary>
    /// Formats a number as Excel shows it in a column with room for <paramref name="width"/> characters (spec §7,
    /// <c>19.0d</c> and <c>19.0e</c> in Excel 1.00): General gets the characters the rest of the text leaves, a fill
    /// repeats into the room left, and text that does not fit becomes a row of <c>#</c>. Excel measures in pixels;
    /// this counts every character as one digit width, an approximation. For the pixel-exact result use
    /// <see cref="ExcelValueFormatter.FormatCell(ExcelCell, ExcelDocument, int, ExcelVersion)"/> (spec Appendix D).
    /// </summary>
    public string Format(double value, int width) => Format(value, width, TextMeasure.Characters);

    /// <summary>
    /// Formats a number with this format into <paramref name="destination"/>, as <see cref="Format(double)"/> does,
    /// without allocating.
    /// </summary>
    /// <returns>False if the text does not fit in <paramref name="destination"/>.</returns>
    public bool TryFormat(double value, Span<char> destination, out int charsWritten)
    {
        var text = new TextBuffer(stackalloc char[128]);
        try
        {
            Format(value, null, TextMeasure.Characters, ref text);
            return text.TryCopyTo(destination, out charsWritten);
        }
        finally
        {
            text.Dispose();
        }
    }

    // Formats into a room of `width` units of `measure`: characters, or pixels of the sheet's font (19.0d, 19.0e).
    internal string Format(double value, int? width, TextMeasure measure)
    {
        var text = new TextBuffer(stackalloc char[128]);
        try
        {
            Format(value, width, measure, ref text);
            return text.ToString();
        }
        finally
        {
            text.Dispose();
        }
    }

    internal void Format(double value, int? width, TextMeasure measure, ref TextBuffer text)
    {
        if (width <= 0)
        {
            return;
        }

        int start = text.Length;

        // Excel never makes an infinity or a NaN, but one in a file shows in General as the number its decimal
        // conversion makes of it (spec §12.1), and as a row of # in any other format (`NaN Wide`, Excel 1.00).
        if (!double.IsFinite(value))
        {
            if (GetSection(ExcelNumberFormatSection.Positive) is [0x01])
            {
                text.Append(GeneralMark);
                Resolve(ref text, start, value, width, measure);
            }
            else if (width is { } room)
            {
                AppendHashes(ref text, room, measure);
            }
            else
            {
                AppendGeneral(ref text, value, 20);
            }

            return;
        }

        // Excel picks the section by sign, and uses the zero section for a value that rounds to zero in the
        // section it would otherwise use: 0.001234 in 0.0" N";0.0" S";0 shows 0.
        var section = value > 0 ? ExcelNumberFormatSection.Positive
            : value < 0 && !IsSingleSection ? ExcelNumberFormatSection.Negative
            : value < 0 ? ExcelNumberFormatSection.Positive
            : ExcelNumberFormatSection.Zero;
        double magnitude = Math.Abs(value);
        if (section != ExcelNumberFormatSection.Zero && RoundsToZero(Tokens(section), magnitude))
        {
            section = ExcelNumberFormatSection.Zero;
            magnitude = 0;
        }

        if (!FormatSection(Tokens(section), magnitude, ref text))
        {
            // A date Excel cannot show: TEXT gives #VALUE!, and a cell a row of #.
            text.Remove(start, text.Length - start);
            if (width is { } columns)
            {
                AppendHashes(ref text, columns, measure);
            }
            else
            {
                AppendGeneral(ref text, value, 20);
            }

            return;
        }

        // A single-section format shows a negative number as a minus sign and the format of its magnitude, dates
        // included: -1234.567 in m/d/yy is -5/19/07.
        if (value < 0 && section == ExcelNumberFormatSection.Positive && text.Length > start)
        {
            text.Insert(start, "-");
        }

        Resolve(ref text, start, magnitude, width, measure);
    }

    // Stand-ins for a General token and a fill, resolved once the rest of the text is known.
    private const char GeneralMark = '\uE000';
    private const char FillMark = '\uE001';

    // Puts General's digits and the fill into the text from `start`, and turns it into #s if it does not fit.
    private static void Resolve(ref TextBuffer text, int start, double magnitude, int? width, TextMeasure measure)
    {
        var shown = text.AsSpan()[start..];
        int fill = shown.IndexOf(FillMark);
        char fillChar = fill >= 0 && fill + 1 < shown.Length ? shown[fill + 1] : ' ';
        if (fill >= 0)
        {
            text.Remove(start + fill, Math.Min(2, shown.Length - fill));
        }

        shown = text.AsSpan()[start..];
        int general = shown.IndexOf(GeneralMark);
        if (general >= 0)
        {
            // General gets the room the rest of the text leaves, in digit widths (19.0d).
            int left = 20;
            if (width is { } columns)
            {
                int rest = measure.Width(shown[..general]) + measure.Width(shown[(general + 1)..]);
                left = Math.Max(0, (columns - rest) / measure.DigitWidth);
            }

            text.Remove(start + general, 1);
            int at = start + general;
            var tail = new TextBuffer(stackalloc char[64]);
            try
            {
                AppendGeneral(ref tail, magnitude, left);
                text.Insert(at, tail.AsSpan());
                if (fill > general)
                {
                    fill += tail.Length - 1;
                }
            }
            finally
            {
                tail.Dispose();
            }
        }

        if (width is not { } room)
        {
            return;
        }

        if (fill >= 0)
        {
            // The fill repeats into what is left, less a pixel, up to 255 characters in all (19.0d).
            shown = text.AsSpan()[start..];
            int count = (room - 1 - measure.Width(shown)) / Math.Max(1, measure.Width(fillChar));
            count = Math.Clamp(count, 0, Math.Max(0, 255 - shown.Length));
            Span<char> repeated = stackalloc char[255];
            repeated = repeated[..count];
            repeated.Fill(fillChar);
            text.Insert(start + Math.Min(fill, shown.Length), repeated);
        }

        // Text wider than the room shows as #s (19.0e).
        if (measure.FitWidth(text.AsSpan()[start..]) > room)
        {
            text.Remove(start, text.Length - start);
            AppendHashes(ref text, room, measure);
        }
    }

    // General with room for `width` characters, written into the text.
    private static void AppendGeneral(ref TextBuffer text, double value, int width)
    {
        int size = ExcelValueFormatter.GeneralSize(width);
        char[]? rented = size > 256 ? System.Buffers.ArrayPool<char>.Shared.Rent(size) : null;
        Span<char> buffer = rented ?? stackalloc char[256];
        try
        {
            text.Append(buffer[..ExcelValueFormatter.FormatGeneral(value, width, buffer)]);
        }
        finally
        {
            if (rented is not null)
            {
                System.Buffers.ArrayPool<char>.Shared.Return(rented);
            }
        }
    }

    // As many #s as fit in the room (19.0e).
    internal static string Hashes(int room, TextMeasure measure)
        => new('#', Math.Max(0, room / measure.HashWidth));

    private static void AppendHashes(ref TextBuffer text, int room, TextMeasure measure)
        => text.Append('#', Math.Max(0, room / measure.HashWidth));

    private static bool RoundsToZero(List<Token> tokens, double magnitude)
    {
        int header = tokens.FindIndex(t => t.Code == 0x03);
        if (header < 0 || tokens.Exists(IsDateToken))
        {
            return false;
        }

        byte[] h = tokens[header].Operands;
        int extra = HeaderExtra(h);
        Span<char> converted = stackalloc char[15];
        int length = ExcelValueFormatter.Num2Str(magnitude, converted, out int point);
        point += extra;
        Span<char> rounded = stackalloc char[15];
        return RoundDigits(converted[..length], point + h[1], rounded, ref point) == 0;
    }

    // A number header's c − b: two per % sign, or (Excel 1.5) minus three per trailing comma. Both are bytes and
    // Excel works the difference out in 8 bits, so a format with 128 decimal places or more (b = c = 0x80) has
    // none of either: Excel 1.00 compiles and shows it as plain decimals (`Many Decimals`).
    private static int HeaderExtra(ReadOnlySpan<byte> header) => (sbyte)(byte)(header[2] - header[1]);

    private static bool IsDateToken(Token token) => token.Code is (>= 0x0B and <= 0x1A) or 0x1F or 0x20 or 0x21;

    /// <summary>
    /// Formats text with this format: through Excel 1.5's text section when the format has one of its own (where
    /// <c>@</c> stands for the text), and otherwise as it is.
    /// </summary>
    public string Format(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (SectionCount < 4 || !HasOwnTextSection())
        {
            return value;
        }

        var text = new TextBuffer(stackalloc char[128]);
        try
        {
            foreach (var token in Tokens(ExcelNumberFormatSection.Text))
            {
                if (token.Code == 0x23)
                {
                    text.Append(value);
                }
                else
                {
                    AppendDisplayLiteral(ref text, token);
                }
            }

            return text.ToString();
        }
        finally
        {
            text.Dispose();
        }
    }

    // Excel 1.5's decompiler writes the text section when it is not General and is not the positive section's
    // bytes (0d.19 in 1.5).
    private bool HasOwnTextSection()
    {
        var textSection = GetSection(ExcelNumberFormatSection.Text);
        if (textSection.Length == 1 && textSection[0] == 0x01)
        {
            return false;
        }

        return Descriptor[(int)ExcelNumberFormatSection.Text] != Descriptor[(int)ExcelNumberFormatSection.Positive]
            || Descriptor[4 + (int)ExcelNumberFormatSection.Text] != Descriptor[4 + (int)ExcelNumberFormatSection.Positive]
            || textSection.Length == 0;
    }

    private static ExcelNumberFormat Create(string descriptor, string body)
        => new(Convert.FromHexString(descriptor), Convert.FromHexString(body));

    private readonly record struct Token(byte Code, byte[] Operands);

    // The tokens of each section, worked out on first use: formatting a value would otherwise tokenise its format
    // every time. A format's bytes are treated as fixed once it is made (a `with` gets a cache of its own).
    private List<Token>?[]? _tokens;

    private ExcelNumberFormat(ExcelNumberFormat original)
    {
        Descriptor = original.Descriptor;
        Body = original.Body;
    }

    private List<Token> Tokens(ExcelNumberFormatSection section)
    {
        var cache = _tokens ??= new List<Token>?[4];
        return cache[(int)section] ??= Tokenize(GetSection(section));
    }

    private static List<Token> Tokenize(ReadOnlySpan<byte> section)
    {
        // A section has at most one token per byte, so one array of that size never grows.
        var tokens = new List<Token>(section.Length);
        int i = 0;
        while (i < section.Length)
        {
            byte code = section[i++];
            int length = code switch
            {
                0x02 or 0x04 or 0x06 or 0x0A or 0x1D or 0x1E or 0x24 => 1,
                0x03 or 0x1C => 3,
                0x09 => i < section.Length ? 1 + section[i] : 0,
                0x21 => i + 1 < section.Length ? 2 + section[i] + section[i + 1] : 0,
                _ => 0,
            };
            if (length > section.Length - i)
            {
                throw new ExcelFormatException($"Number-format token 0x{code:X2} runs past the end of its section.");
            }

            tokens.Add(new Token(code, section.Slice(i, length).ToArray()));
            i += length;
        }

        return tokens;
    }

    // Characters a literal can show without a backslash: those at positions 6–14 of Excel 1.5's format-character
    // string (spec §7).
    private static bool IsUnescaped(char c) => c is ',' or ':' or '/' or ' ' or '-' or '+' or '(' or ')' or '.';

    private static void AppendSectionText(StringBuilder text, ReadOnlySpan<byte> section, bool export)
    {
        var tokens = Tokenize(section);
        int integerLeft = 0;
        int digitsLeft = 0;
        int scale = 0;
        int percents = 0;
        bool thousands = false;
        bool afterPoint = false;
        for (int t = 0; t < tokens.Count; t++)
        {
            var token = tokens[t];
            byte[] op = token.Operands;
            switch (token.Code)
            {
                case 0x01:
                    text.Append("General");
                    break;
                case 0x02 when op.Length == 1:
                    char c = (char)op[0];
                    bool beforeExponent = t + 1 < tokens.Count && tokens[t + 1].Code is 0x1D or 0x1E;
                    if (c == '%' && percents > 0)
                    {
                        percents--;
                    }
                    else if (!beforeExponent && (export || !IsUnescaped(c)))
                    {
                        text.Append('\\');
                    }

                    text.Append(MacRoman.GetString([op[0]]));
                    break;
                case 0x03 or 0x1C when op.Length == 3:
                    integerLeft = op[0];
                    digitsLeft = op[0] + op[1];
                    int extra = HeaderExtra(op);
                    percents = extra > 0 ? extra / 2 : 0;
                    scale = extra < 0 ? -extra / 3 : 0;
                    thousands = false;
                    afterPoint = false;
                    break;
                case 0x04 or 0x06 when op.Length == 1:
                    char digit = token.Code == 0x04 ? '#' : '0';
                    for (int n = 0; n < op[0]; n++)
                    {
                        text.Append(digit);
                        integerLeft--;
                        digitsLeft--;
                        if (!afterPoint && thousands && integerLeft > 0 && integerLeft % 3 == 0)
                        {
                            text.Append(',');
                        }
                    }

                    if (digitsLeft == 0)
                    {
                        for (; scale > 0; scale--)
                        {
                            text.Append(',');
                        }
                    }

                    break;
                case 0x05:
                    text.Append('.');
                    afterPoint = true;
                    break;
                case 0x07:
                    thousands = true;
                    break;
                case 0x1D or 0x1E when op.Length == 1:
                    text.Append(token.Code == 0x1D ? '+' : '-');
                    integerLeft = op[0];
                    digitsLeft = op[0];
                    afterPoint = false;
                    thousands = false;
                    break;
                case 0x09 or 0x0A or 0x22 or 0x23 or 0x24 or 0x21:
                case >= 0x0B and <= 0x1A:
                case 0x1F or 0x20:
                    if (export && token.Code == 0x22)
                    {
                        text.Append("\"$\"");
                    }
                    else
                    {
                        AppendLiteral(text, token);
                    }

                    break;
            }
        }
    }

    private static void AppendLiteral(StringBuilder text, Token token)
    {
        byte[] op = token.Operands;
        switch (token.Code)
        {
            case 0x02 when op.Length == 1:
                text.Append(MacRoman.GetString([op[0]]));
                break;
            case 0x09 when op.Length >= 1:
                text.Append('"').Append(MacRoman.GetString(op.AsSpan(1))).Append('"');
                break;
            case 0x0A when op.Length == 1:
                text.Append('*').Append(MacRoman.GetString([op[0]]));
                break;
            case 0x22:
                text.Append('$');
                break;
            case 0x23:
                text.Append('@');
                break;
            case 0x24 when op.Length == 1:
                text.Append('[').Append(op[0] < ColorNames.Length ? ColorNames[op[0]] : op[0].ToString(CultureInfo.InvariantCulture)).Append(']');
                break;
            case 0x21 when op.Length >= 2 && op.Length >= 2 + op[0] + op[1]:
                text.Append(MacRoman.GetString(op.AsSpan(2, op[0]))).Append('/').Append(MacRoman.GetString(op.AsSpan(2 + op[0], op[1])));
                break;
            default:
                text.Append(token.Code switch
                {
                    0x0B => "d", 0x0C => "dd", 0x0D => "ddd", 0x0E => "dddd",
                    0x0F => "m", 0x10 => "mm", 0x11 => "mmm", 0x12 => "mmmm",
                    0x13 or 0x1F => "h", 0x14 or 0x20 => "hh",
                    0x15 => "m", 0x16 => "mm", 0x17 => "s", 0x18 => "ss",
                    0x19 => "yy", 0x1A => "yyyy",
                    _ => string.Empty,
                });
                break;
        }
    }

    // Formats a non-negative magnitude with one section into the text, or returns false for a date Excel cannot show.
    private static bool FormatSection(List<Token> tokens, double magnitude, ref TextBuffer text)
    {
        if (tokens.Count == 0)
        {
            return true;
        }

        if (tokens.Exists(IsDateToken))
        {
            return FormatDate(tokens, magnitude, ref text);
        }

        int header = tokens.FindIndex(static t => t.Code is 0x03 or 0x1C);
        if (header < 0)
        {
            foreach (var token in tokens)
            {
                if (token.Code == 0x01)
                {
                    text.Append(GeneralMark);
                }
                else
                {
                    AppendDisplayLiteral(ref text, token);
                }
            }

            return true;
        }

        FormatNumber(tokens, header, magnitude, ref text);
        return true;
    }

    private static void FormatNumber(List<Token> tokens, int header, double magnitude, ref TextBuffer text)
    {
        byte[] h = tokens[header].Operands;
        bool scientific = tokens[header].Code == 0x1C;
        int decimals = h[1];
        int extra = HeaderExtra(h);
        bool thousands = tokens.Exists(static t => t.Code == 0x07);

        // The digit placeholders in order, '#' or '0': the mantissa's integer places, then its decimal places, then
        // the exponent's. A point (05) moves from integer to decimal places and an exponent sign (1D, 1E) to the
        // exponent's, so each kind is one run.
        int total = 0;
        foreach (var token in tokens)
        {
            if (token.Code is 0x04 or 0x06 && token.Operands.Length == 1)
            {
                total += token.Operands[0];
            }
        }

        char[]? rented = total > 256 ? System.Buffers.ArrayPool<char>.Shared.Rent(total) : null;
        Span<char> places = rented ?? stackalloc char[256];
        try
        {
            int integerCount = 0, fractionCount = 0, count = 0;
            int part = 0; // 0 integer, 1 fraction, 2 exponent
            foreach (var token in tokens)
            {
                if (token.Code == 0x05 && part == 0)
                {
                    part = 1;
                }
                else if (token.Code is 0x1D or 0x1E)
                {
                    part = 2;
                }
                else if (token.Code is 0x04 or 0x06 && token.Operands.Length == 1)
                {
                    char place = token.Code == 0x04 ? '#' : '0';
                    for (int n = 0; n < token.Operands[0]; n++)
                    {
                        places[count++] = place;
                    }

                    integerCount += part == 0 ? token.Operands[0] : 0;
                    fractionCount += part == 1 ? token.Operands[0] : 0;
                }
            }

            ReadOnlySpan<char> integerPlaces = places[..integerCount];
            ReadOnlySpan<char> fractionPlaces = places.Slice(integerCount, fractionCount);
            ReadOnlySpan<char> exponentPlaces = places[(integerCount + fractionCount)..count];

            // Excel's decimal conversion: at most 14 significant digits, the value being 0.digits × 10^point. The
            // percent signs and scaling commas move the point (spec §7), so they cannot overflow the double.
            Span<char> converted = stackalloc char[15];
            int convertedLength = ExcelValueFormatter.Num2Str(magnitude, converted, out int point);
            point += extra;
            int exponent = 0;
            int mantissaDigits = Math.Max(1, integerPlaces.Length);
            if (scientific && convertedLength > 0)
            {
                exponent = point - mantissaDigits;
                point = mantissaDigits;
            }

            // The digits rounded half up to the positions there are.
            Span<char> significant = stackalloc char[15];
            int significantLength = RoundDigits(converted[..convertedLength], point + decimals, significant, ref point);
            if (scientific && significantLength > 0 && point > mantissaDigits)
            {
                // Rounding carried into a new digit (9.99 to 10.0): one more for the exponent instead.
                point--;
                exponent++;
            }

            significant = significant[..significantLength];
            int integerLength = significantLength == 0 ? 0 : Math.Max(0, point);
            // At most 309 + 128 integer digits (the scaling is a signed byte) and 255 decimals.
            Span<char> digitText = stackalloc char[1024];
            for (int i = 0; i < integerLength + decimals; i++)
            {
                int at = i < integerLength ? i : point + (i - integerLength);
                digitText[i] = at >= 0 && at < significant.Length ? significant[at] : '0';
            }

            ReadOnlySpan<char> integerDigits = digitText[..integerLength];
            ReadOnlySpan<char> fractionDigits = digitText.Slice(integerLength, decimals);
            Span<char> exponentText = stackalloc char[11];
            Math.Abs(exponent).TryFormat(exponentText, out int exponentLength, default, CultureInfo.InvariantCulture);
            var exponentDigits = exponentText[..exponentLength].TrimStart('0');

            // Decimal places show their digit, except trailing zeros where the placeholder is #.
            int keep = fractionPlaces.Length;
            while (keep > 0 && fractionPlaces[keep - 1] == '#' && keep - 1 < fractionDigits.Length && fractionDigits[keep - 1] == '0')
            {
                keep--;
            }

            // Walk the tokens again, putting each digit where its placeholder is.
            int integerNext = 0, fractionNext = 0, exponentNext = 0;
            part = 0;
            foreach (var token in tokens)
            {
                switch (token.Code)
                {
                    case 0x04 or 0x06 when token.Operands.Length == 1:
                        for (int n = 0; n < token.Operands[0]; n++)
                        {
                            if (part == 0 && integerNext < integerPlaces.Length)
                            {
                                AppendIntegerPlace(ref text, integerDigits, integerPlaces, integerNext++, thousands);
                            }
                            else if (part == 1 && fractionNext < fractionPlaces.Length)
                            {
                                if (fractionNext < keep && fractionNext < fractionDigits.Length)
                                {
                                    text.Append(fractionDigits[fractionNext]);
                                }

                                fractionNext++;
                            }
                            else if (part == 2 && exponentNext < exponentPlaces.Length)
                            {
                                AppendIntegerPlace(ref text, exponentDigits, exponentPlaces, exponentNext++, thousands: false);
                            }
                        }

                        break;
                    case 0x05 when part == 0:
                        text.Append('.');
                        part = 1;
                        break;
                    case 0x1D or 0x1E:
                        if (exponent < 0)
                        {
                            text.Append('-');
                        }
                        else if (token.Code == 0x1D)
                        {
                            text.Append('+');
                        }

                        part = 2;
                        break;
                    default:
                        AppendDisplayLiteral(ref text, token);
                        break;
                }
            }
        }
        finally
        {
            if (rented is not null)
            {
                System.Buffers.ArrayPool<char>.Shared.Return(rented);
            }
        }
    }

    // Rounds the digits of 0.digits × 10^point half up to the first `keep` of them, moving the point when a carry
    // adds a digit (0.999 to 0.1 × 10^(point + 1)). Returns how many digits are left: 0 when the value rounds to 0.
    private static int RoundDigits(ReadOnlySpan<char> digits, int keep, Span<char> result, ref int point)
    {
        if (keep >= digits.Length)
        {
            digits.CopyTo(result);
            return digits.Length;
        }

        if (keep < 0 || (keep == 0 && digits[0] < '5'))
        {
            return 0;
        }

        digits[..keep].CopyTo(result);
        if (digits[keep] < '5')
        {
            return keep;
        }

        for (int i = keep - 1; i >= 0; i--)
        {
            if (result[i] != '9')
            {
                result[i]++;
                return i + 1;
            }
        }

        result[0] = '1';
        point++;
        return 1;
    }

    // What one integer placeholder shows (`place` counts them from the left): a 0 shows a digit or 0, a # a digit or
    // nothing, digits that do not fit go with the first placeholder, and a separator follows every third position
    // from the right.
    private static void AppendIntegerPlace(ref TextBuffer text, scoped ReadOnlySpan<char> digits, scoped ReadOnlySpan<char> places, int place, bool thousands)
    {
        int width = Math.Max(places.Length, digits.Length);
        int spare = width - places.Length;
        for (int i = place == 0 ? 0 : spare + place; i <= spare + place; i++)
        {
            int at = i - spare;
            int digit = i - (width - digits.Length);
            char shown = digit >= 0 ? digits[digit] : at >= 0 && places[at] == '0' ? '0' : '\0';
            if (shown == '\0')
            {
                continue;
            }

            text.Append(shown);
            int remaining = width - i - 1;
            if (thousands && remaining > 0 && remaining % 3 == 0)
            {
                text.Append(',');
            }
        }
    }

    // A token as it shows in a formatted value: strings without their quotes, and fills (which need a column
    // width) and colours as nothing.
    private static void AppendDisplayLiteral(ref TextBuffer text, Token token)
    {
        byte[] op = token.Operands;
        switch (token.Code)
        {
            case 0x02 when op.Length == 1:
                text.Append(MacRoman.GetChar(op[0]));
                break;
            case 0x09 when op.Length >= 1:
                text.AppendMacRoman(op.AsSpan(1), escapeQuotes: false);
                break;
            case 0x22:
                text.Append('$');
                break;
            case 0x0A when op.Length == 1:
                text.Append(FillMark);
                text.Append(MacRoman.GetChar(op[0]));
                break;
        }
    }

    // Excel converts the serial to seconds since 1904, rounded, and gives up past the Macintosh clock's last second,
    // 2^32 - 1 (6 February 2040, serial 49710.27): TEXT gives #VALUE! there (Excel 1.00, "Dates" run).
    private static bool FormatDate(List<Token> tokens, double serial, ref TextBuffer text)
    {
        double seconds = Math.Round(serial * 86400, MidpointRounding.AwayFromZero);
        if (seconds > uint.MaxValue)
        {
            return false;
        }

        var date = new DateTime(1904, 1, 1).AddSeconds(seconds);
        bool twelveHour = tokens.Exists(static t => t.Code == 0x21);
        var names = CultureInfo.InvariantCulture.DateTimeFormat;
        int hour12 = date.Hour % 12 == 0 ? 12 : date.Hour % 12;
        foreach (var token in tokens)
        {
            switch (token.Code)
            {
                case 0x0B: text.AppendFormatted(date.Day, default); break;
                case 0x0C: text.AppendFormatted(date.Day, "00"); break;
                case 0x0D: text.Append(names.GetAbbreviatedDayName(date.DayOfWeek)); break;
                case 0x0E: text.Append(names.GetDayName(date.DayOfWeek)); break;
                case 0x0F: text.AppendFormatted(date.Month, default); break;
                case 0x10: text.AppendFormatted(date.Month, "00"); break;
                case 0x11: text.Append(names.GetAbbreviatedMonthName(date.Month)); break;
                case 0x12: text.Append(names.GetMonthName(date.Month)); break;
                case 0x13: text.AppendFormatted(twelveHour ? hour12 : date.Hour, default); break;
                case 0x14: text.AppendFormatted(twelveHour ? hour12 : date.Hour, "00"); break;
                case 0x1F: text.AppendFormatted(hour12, default); break;
                case 0x20: text.AppendFormatted(hour12, "00"); break;
                case 0x15: text.AppendFormatted(date.Minute, default); break;
                case 0x16: text.AppendFormatted(date.Minute, "00"); break;
                case 0x17: text.AppendFormatted(date.Second, default); break;
                case 0x18: text.AppendFormatted(date.Second, "00"); break;
                case 0x19: text.AppendFormatted(date.Year % 100, "00"); break;
                case 0x1A: text.AppendFormatted(date.Year, default); break;
                case 0x21 when token.Operands.Length >= 2:
                    byte[] op = token.Operands;
                    text.AppendMacRoman(date.Hour < 12 ? op.AsSpan(2, op[0]) : op.AsSpan(2 + op[0], op[1]), escapeQuotes: false);
                    break;
                default:
                    AppendDisplayLiteral(ref text, token);
                    break;
            }
        }

        return true;
    }
}
