using System.Globalization;
using System.Text;

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
                foreach (var token in Tokenize(GetSection((ExcelNumberFormatSection)s)))
                {
                    if (token.Code is 0x23 or 0x24 || (token.Code is 0x03 or 0x1C && token.Operands.Length == 3 && (sbyte)token.Operands[2] < token.Operands[1]))
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
            foreach (var token in Tokenize(GetSection(ExcelNumberFormatSection.Positive)))
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
    public string Format(double value) => Format(value, null, MacExcelReader.Internal.TextMeasure.Characters);

    /// <summary>
    /// Formats a number as Excel shows it in a column with room for <paramref name="width"/> characters (spec §7,
    /// <c>19.0d</c> and <c>19.0e</c> in Excel 1.00): General gets the characters the rest of the text leaves, a fill
    /// repeats into the room left, and text that does not fit becomes a row of <c>#</c>. Excel measures in pixels;
    /// this counts every character as one digit width, an approximation. For the pixel-exact result use
    /// <see cref="ExcelValueFormatter.FormatCell(ExcelCell, ExcelDocument, int)"/> (spec Appendix D).
    /// </summary>
    public string Format(double value, int width) => Format(value, width, MacExcelReader.Internal.TextMeasure.Characters);

    // Formats into a room of `width` units of `measure`: characters, or pixels of the sheet's font (19.0d, 19.0e).
    internal string Format(double value, int? width, MacExcelReader.Internal.TextMeasure measure)
    {
        if (width <= 0)
        {
            return string.Empty;
        }

        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return ExcelValueFormatter.FormatGeneral(value);
        }

        // Excel picks the section by sign, and uses the zero section for a value that rounds to zero in the
        // section it would otherwise use: 0.001234 in 0.0" N";0.0" S";0 shows 0.
        var section = value > 0 ? ExcelNumberFormatSection.Positive
            : value < 0 && !IsSingleSection ? ExcelNumberFormatSection.Negative
            : value < 0 ? ExcelNumberFormatSection.Positive
            : ExcelNumberFormatSection.Zero;
        double magnitude = Math.Abs(value);
        if (section != ExcelNumberFormatSection.Zero && RoundsToZero(GetSection(section), magnitude))
        {
            section = ExcelNumberFormatSection.Zero;
            magnitude = 0;
        }

        string? shown = FormatSection(GetSection(section), magnitude);
        if (shown is null)
        {
            // A date Excel cannot show: TEXT gives #VALUE!, and a cell a row of #.
            return width is { } columns ? Hashes(columns, measure) : ExcelValueFormatter.FormatGeneral(value);
        }

        // A single-section format shows a negative number as a minus sign and the format of its magnitude, dates
        // included: -1234.567 in m/d/yy is -5/19/07.
        if (value < 0 && section == ExcelNumberFormatSection.Positive && shown.Length > 0)
        {
            shown = "-" + shown;
        }

        return Resolve(shown, magnitude, width, measure);
    }

    // Stand-ins for a General token and a fill, resolved once the rest of the text is known.
    private const char GeneralMark = '\uE000';
    private const char FillMark = '\uE001';

    private static string Resolve(string text, double magnitude, int? width, MacExcelReader.Internal.TextMeasure measure)
    {
        int fill = text.IndexOf(FillMark);
        char fillChar = fill >= 0 && fill + 1 < text.Length ? text[fill + 1] : ' ';
        if (fill >= 0)
        {
            text = text.Remove(fill, Math.Min(2, text.Length - fill));
        }

        int general = text.IndexOf(GeneralMark);
        if (general >= 0)
        {
            // General gets the room the rest of the text leaves, in digit widths (19.0d).
            int left = 20;
            if (width is { } columns)
            {
                int rest = measure.Width(text.AsSpan(0, general)) + measure.Width(text.AsSpan(general + 1));
                left = Math.Max(0, (columns - rest) / measure.Width('0'));
            }

            string shown = ExcelValueFormatter.FormatGeneral(magnitude, left);
            text = text.Remove(general, 1).Insert(general, shown);
            if (fill > general)
            {
                fill += shown.Length - 1;
            }
        }

        if (width is not { } room)
        {
            return text;
        }

        if (fill >= 0)
        {
            // The fill repeats into what is left, less a pixel, up to 255 characters in all (19.0d).
            int count = (room - 1 - measure.Width(text)) / Math.Max(1, measure.Width(fillChar));
            count = Math.Clamp(count, 0, Math.Max(0, 255 - text.Length));
            text = text.Insert(Math.Min(fill, text.Length), new string(fillChar, count));
        }

        // Text wider than the room shows as #s (19.0e).
        return measure.Width(text) > room ? Hashes(room, measure) : text;
    }

    // As many #s as fit in the room (19.0e).
    internal static string Hashes(int room, MacExcelReader.Internal.TextMeasure measure)
        => new('#', Math.Max(0, room / measure.Width('#')));

    private static bool RoundsToZero(ReadOnlySpan<byte> section, double magnitude)
    {
        var tokens = Tokenize(section);
        int header = tokens.FindIndex(t => t.Code == 0x03);
        if (header < 0 || tokens.Exists(IsDateToken))
        {
            return false;
        }

        byte[] h = tokens[header].Operands;
        int extra = (sbyte)h[2] - h[1];
        double scaled = magnitude * Math.Pow(100, extra > 0 ? extra / 2 : 0) / Math.Pow(1000, extra < 0 ? -extra / 3 : 0);
        return Math.Round(scaled, h[1], MidpointRounding.AwayFromZero) == 0;
    }

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

        var text = new StringBuilder();
        foreach (var token in Tokenize(GetSection(ExcelNumberFormatSection.Text)))
        {
            if (token.Code == 0x23)
            {
                text.Append(value);
            }
            else
            {
                AppendDisplayLiteral(text, token);
            }
        }

        return text.ToString();
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
            length = Math.Min(length, section.Length - i);
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
                    int extra = (sbyte)op[2] - op[1];
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

    // Formats a non-negative magnitude with one section, or returns null for a date Excel cannot show.
    private static string? FormatSection(ReadOnlySpan<byte> section, double magnitude)
    {
        var tokens = Tokenize(section);
        if (tokens.Count == 0)
        {
            return string.Empty;
        }

        if (tokens.Exists(IsDateToken))
        {
            return FormatDate(tokens, magnitude);
        }

        int header = tokens.FindIndex(t => t.Code is 0x03 or 0x1C);
        if (header < 0)
        {
            var text = new StringBuilder();
            foreach (var token in tokens)
            {
                if (token.Code == 0x01)
                {
                    text.Append(GeneralMark);
                }
                else
                {
                    AppendDisplayLiteral(text, token);
                }
            }

            return text.ToString();
        }

        return FormatNumber(tokens, header, magnitude);
    }

    private static string FormatNumber(List<Token> tokens, int header, double magnitude)
    {
        byte[] h = tokens[header].Operands;
        bool scientific = tokens[header].Code == 0x1C;
        int decimals = h[1];
        int extra = (sbyte)h[2] - h[1];
        double value = magnitude * Math.Pow(100, extra > 0 ? extra / 2 : 0) / Math.Pow(1000, extra < 0 ? -extra / 3 : 0);
        bool thousands = tokens.Exists(t => t.Code == 0x07);

        // Split the tokens into mantissa digit placeholders (before and after the point) and exponent ones.
        var integerPlaces = new List<char>();
        var fractionPlaces = new List<char>();
        var exponentPlaces = new List<char>();
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
                var list = part == 0 ? integerPlaces : part == 1 ? fractionPlaces : exponentPlaces;
                list.AddRange(Enumerable.Repeat(token.Code == 0x04 ? '#' : '0', token.Operands[0]));
            }
        }

        int exponent = 0;
        if (scientific && value != 0)
        {
            int mantissaDigits = Math.Max(1, integerPlaces.Count);
            exponent = (int)Math.Floor(Math.Log10(value)) - (mantissaDigits - 1);
            double mantissa = Math.Round(value / Math.Pow(10, exponent), decimals, MidpointRounding.AwayFromZero);
            if (mantissa >= Math.Pow(10, mantissaDigits))
            {
                exponent++;
            }

            value /= Math.Pow(10, exponent);
        }

        string fixedText = Math.Round(value, decimals, MidpointRounding.AwayFromZero).ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        int point = fixedText.IndexOf('.');
        string integerDigits = (point < 0 ? fixedText : fixedText[..point]).TrimStart('0');
        string fractionDigits = point < 0 ? string.Empty : fixedText[(point + 1)..];
        string[][] pieces =
        [
            FillInteger(integerDigits, integerPlaces, thousands),
            FillFraction(fractionDigits, fractionPlaces),
            FillInteger(Math.Abs(exponent).ToString(CultureInfo.InvariantCulture).TrimStart('0'), exponentPlaces, false),
        ];

        // Walk the tokens again, putting each digit where its placeholder is.
        var text = new StringBuilder();
        int[] next = new int[3];
        part = 0;
        foreach (var token in tokens)
        {
            switch (token.Code)
            {
                case 0x04 or 0x06 when token.Operands.Length == 1:
                    for (int n = 0; n < token.Operands[0] && next[part] < pieces[part].Length; n++)
                    {
                        text.Append(pieces[part][next[part]++]);
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
                    AppendDisplayLiteral(text, token);
                    break;
            }
        }

        return text.ToString();
    }

    // What each integer placeholder shows, left to right: a 0 shows a digit or 0, a # a digit or nothing, digits
    // that do not fit go with the first placeholder, and a separator follows every third position from the right.
    private static string[] FillInteger(string digits, List<char> places, bool thousands)
    {
        var pieces = new string[places.Count];
        int width = Math.Max(places.Count, digits.Length);
        var piece = new StringBuilder();
        int placeIndex = 0;
        for (int i = 0; i < width; i++)
        {
            int place = i - (width - places.Count);
            int digitIndex = i - (width - digits.Length);
            char? shown = digitIndex >= 0 ? digits[digitIndex] : place >= 0 && places[place] == '0' ? '0' : null;
            if (shown is not null)
            {
                piece.Append(shown.Value);
                int remaining = width - i - 1;
                if (thousands && remaining > 0 && remaining % 3 == 0)
                {
                    piece.Append(',');
                }
            }

            if (place >= 0)
            {
                pieces[placeIndex++] = piece.ToString();
                piece.Clear();
            }
        }

        return pieces;
    }

    // What each decimal placeholder shows: its digit, except trailing zeros where the placeholder is #.
    private static string[] FillFraction(string digits, List<char> places)
    {
        var pieces = new string[places.Count];
        int keep = places.Count;
        while (keep > 0 && places[keep - 1] == '#' && keep - 1 < digits.Length && digits[keep - 1] == '0')
        {
            keep--;
        }

        for (int i = 0; i < places.Count; i++)
        {
            pieces[i] = i < keep && i < digits.Length ? digits[i].ToString() : string.Empty;
        }

        return pieces;
    }

    // A token as it shows in a formatted value: strings without their quotes, and fills (which need a column
    // width) and colours as nothing.
    private static void AppendDisplayLiteral(StringBuilder text, Token token)
    {
        byte[] op = token.Operands;
        switch (token.Code)
        {
            case 0x02 when op.Length == 1:
                text.Append(MacRoman.GetString([op[0]]));
                break;
            case 0x09 when op.Length >= 1:
                text.Append(MacRoman.GetString(op.AsSpan(1)));
                break;
            case 0x22:
                text.Append('$');
                break;
            case 0x0A when op.Length == 1:
                text.Append(FillMark).Append(MacRoman.GetString([op[0]]));
                break;
        }
    }

    // Excel converts the serial to seconds since 1904, rounded, and gives up past the Macintosh clock's last second,
    // 2^32 - 1 (6 February 2040, serial 49710.27): TEXT gives #VALUE! there (Excel 1.00, "Dates" run).
    private static string? FormatDate(List<Token> tokens, double serial)
    {
        double seconds = Math.Round(serial * 86400, MidpointRounding.AwayFromZero);
        if (seconds > uint.MaxValue)
        {
            return null;
        }

        var date = new DateTime(1904, 1, 1).AddSeconds(seconds);
        bool twelveHour = tokens.Exists(t => t.Code == 0x21);
        var invariant = CultureInfo.InvariantCulture;
        var text = new StringBuilder();
        foreach (var token in tokens)
        {
            int hour12 = date.Hour % 12 == 0 ? 12 : date.Hour % 12;
            switch (token.Code)
            {
                case 0x0B: text.Append(date.Day.ToString(invariant)); break;
                case 0x0C: text.Append(date.ToString("dd", invariant)); break;
                case 0x0D: text.Append(date.ToString("ddd", invariant)); break;
                case 0x0E: text.Append(date.ToString("dddd", invariant)); break;
                case 0x0F: text.Append(date.Month.ToString(invariant)); break;
                case 0x10: text.Append(date.ToString("MM", invariant)); break;
                case 0x11: text.Append(date.ToString("MMM", invariant)); break;
                case 0x12: text.Append(date.ToString("MMMM", invariant)); break;
                case 0x13: text.Append((twelveHour ? hour12 : date.Hour).ToString(invariant)); break;
                case 0x14: text.Append((twelveHour ? hour12 : date.Hour).ToString("00", invariant)); break;
                case 0x1F: text.Append(hour12.ToString(invariant)); break;
                case 0x20: text.Append(hour12.ToString("00", invariant)); break;
                case 0x15: text.Append(date.Minute.ToString(invariant)); break;
                case 0x16: text.Append(date.Minute.ToString("00", invariant)); break;
                case 0x17: text.Append(date.Second.ToString(invariant)); break;
                case 0x18: text.Append(date.Second.ToString("00", invariant)); break;
                case 0x19: text.Append(date.ToString("yy", invariant)); break;
                case 0x1A: text.Append(date.Year.ToString(invariant)); break;
                case 0x21 when token.Operands.Length >= 2:
                    byte[] op = token.Operands;
                    text.Append(date.Hour < 12
                        ? MacRoman.GetString(op.AsSpan(2, op[0]))
                        : MacRoman.GetString(op.AsSpan(2 + op[0], op[1])));
                    break;
                default:
                    AppendDisplayLiteral(text, token);
                    break;
            }
        }

        return text.ToString();
    }
}
