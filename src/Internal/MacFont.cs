namespace MacExcelReader.Internal;

/// <summary>
/// The metrics of a bitmap font in the System file Excel 1.00 shipped with (on its Excel Data disk): what Excel's
/// drawing code measures cells with (spec Appendix D). Excel picks the sheet's standard font and size (spec §6, 0x244 and
/// 0x246); QuickDraw draws bold by adding a pixel to every character, and Excel allows half the ascent for the slant
/// of italic text.
/// </summary>
internal sealed class MacFont
{
    // Font number, size, ascent, descent, leading, then the advance widths of Mac Roman characters 0–255 in hex;
    // characters the font lacks have the width of its missing-character symbol.
    private static readonly MacFont[] s_fonts =
    [
        new(0, 12, 12, 3, 1, // Chicago 12
            "000A0A0A0A0A0A0A0A080A0A0A000A0A0A0B0B090B0A0A0A0A0A0A0A0A0A0A0A0406070A070B0A030505070704070407080808080808080808080404060806080B0808080807070808060709070C0908080808070608080C0808080507050808060808070808060808040608040C0808080806070608080C0808080505050808" +
            "080808070908080808080808080808080808060606060808080808080808080805060709070709090A0A0B0506090B090E0706060A0A08090A0B0607070A0C080806070C060A0909090E080808080B0C060A070704040709080A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A"),
        new(2, 10, 10, 2, 0, // New York 10
            "000900090009090909060909090009090909090909090909090909090909090903030508060A0803050507060305030606060606060606060606030304050406080807060706060708050508070A0806070608060608080A0608060406040506030607060606050708040407040A080607070605040707090607050503050500" +
            "080806060806080606060606060606060606040404040806060606060707070705060607060608070707090305090A070909090907090909090909060509090706030909090909060607060808060A09040706060404090707060909090909090909090909090909090909090909090909090909090909090909090909090909"),
        new(2, 12, 12, 3, 1, // New York 12
            "000C0C0C0C0C0C0C0C080C0C0C000C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0404050807090A030505060603060307080808080808080808080303050705070A0A08070807070809060508080C09080808080708090A0D0808060407040407040708060807050708040408040C0807080806060508080B0708060503050600" +
            "0A0A070709080907070707070706070707070404040408070707070708080808040507080708090709090904040C0C080C0C0C0C080C0C0C0C0C0C07060C0A0707040C0C0C0C0C070709080A0A080C0B0509050503030C09080A0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C0C"),
        new(2, 20, 18, 4, 2, // New York 20
            "0013001300131313130D1313130013131313131313131313131313131313131306060A100C1410060A0A0E0C060A060C0C0C0C0D0C0C0C0C0C0C0606080B080C10100E0C0E0C0C0E100A0A100E14100D0E0D0E0C0D1010140D100C080C080A0C060C0D0C0C0C0A0E1008080E0814100C0E0E0C0A080E0E120C0E0A0A060A0A00" +
            "10100C0C100E100C0C0C0C0C0C0B0C0C0C0C08080808100C0C0C0C0C0E0E0E0E0A0C0A0E0C0C100E0E0E12060A13140E131313130E1313131313130C0A13120C0C0613131313130C0C0E0C10100E1412080E0C0C0808130E0E0A1313131313131313131313131313131313131313131313131313131313131313131313131313"),
        new(2, 24, 21, 5, 3, // New York 24
            "001800180018181818101818180018181818181818181818181818181818181808080A100E1214060A0A0C0C060C060E1010101010101010101006060A0E0A0E1414100E100E0E10120C0A10101812101010100E1012141A10100C080E08080E060E100C100E0A0E100808100818100E10100C0C0A1010160E100C0A060A0C00" +
            "14140E0E1210120E0E0E0E0E0E0C0E0E0E0E08080808100E0E0E0E0E10101010080A0E100E10120E121212080818181018181818101818181818180E0C18140E0E0818181818180E0E121014141018160A120A0A0606181210141818181818181818181818181818181818181818181818181818181818181818181818181818"),
        new(3, 9, 10, 2, 0, // Geneva 9
            "0000000000000707070607070700070707070707070707070707070707070707030305070609080304040706040503060606060606060606060604040506050608070606060505060603060605080606060606060606060806060504060404060305050505050405050304050308050505050505040506080606050402040600" +
            "0707060506060605050505050505050505050303040405050505050505050505040506060606060609090904040609070A06050506060506070805050508090606030508050607080808060707060908050707070404060806090707070707070707070707070707070707070707070707070707070707070707070707070707"),
        new(3, 10, 10, 2, 1, // Geneva 10
            "000800000000080808060808080008080808080808080808080808080808080803040509070A080306060708040703060707070707070707070703040507050709060706070606060704060706090706070607060607060A06060605050505070405060505050405060405060409060506050605040606080606050504050700" +
            "060606060706070505050506050505050505040405040605050505060606060605050607060606070A0A0903040709080B070707070705070708050505080807070407090507070807090706070709080609070704040809060C0808080808080808080808080808080808080808080808080808080808080808080808080808"),
        new(3, 12, 12, 3, 1, // Geneva 12
            "000A000000000A0A0A080A0A0A000A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A04030509070909030505080704070307080808080808080808080304060706080A090808080707080805070807090808080808080808090B0707070407040408040707070707060707050607050B0707070706070607070B0707070503050800" +
            "090908070808080707070707070707070707050505050707070707070707070705060708080808090A0A0B0405080C080B07060607090708090A0506060A0B070804070A05080908080A080909080C0B0609070704040709080F0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A"),
        new(3, 14, 14, 4, 1, // Geneva 14
            "000B000000000B0B0B090B0B0B000B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B04040609070B0B030606090904080307090909090909090909090304070B07080C0B09090908080A09050809080B0A0A090A0A09090A0B0D0909090507050709050809080909060908050509050D0809090907080708090B0909080503050900" +
            "0B0B09080A0A0A08080808080808090909090505050608090909090908080808050608090909090A0B0B0C0406090D0A0F090707090A080A0A0B0507070A0E080904080C05090B09090A090B0B0A0E0E070A08080404090B090F0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B0B"),
        new(3, 18, 18, 4, 1, // Geneva 18
            "0001000000000E0E0E0C0E0E0E000E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0606080E0C0E100307070E0C060A060A0C0C0C0C0C0C0C0C0C0C06060A0A0A0C100E0C0C0C0A0A0C0C060A0C0A0E0C0C0C0C0C0C0C0C0E100C0C0A080A08080C060A0A0A0A0A080A0A06080A06100A0A0A0A0A0A080A0A100A0C0A0804080C00" +
            "0E0E0C0A0C0C0C0A0A0A0A0A0A0A0A0A0A0A060606070A0A0A0A0A0A0A0A0A0A08080A0C0C0C0C0C10101208080A120C100C0A0A0C0C0A0C0E0E0A0A0A10120A0C060A100A0A0E0E0E100C0E0E0C12100A0E0E0E08080A100C110E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E0E"),
        new(3, 20, 20, 4, 0, // Geneva 20
            "0010000000001010100C1010100010101010101010101010101010101010101006080A120E1210060C0A0E10080E060C0E0E0E0E0E0E0E0E0E0E06080A0E0A0E120C0E0C0E0C0C0C0E080C0E0C120E0C0E0C0E0C0C0E0C140C0C0A0A0A0A0A0E080A0C0A0A0A080A0C08080C08120C0A0C0A0C0A080C0C100C0C0A0A080A0E00" +
            "0C0C0C0C0E0C0E0A0A0A0A0A0A0A0A0A0A0A08080A080C0A0A0A0A0C0C0C0C0C0A0A0C0E0C0C0C0E14141206080E1210160E0A0E0E0E0A0E0E100A0A0A10100E0E080E120A0E0E100E120E0C0E0E12100B120E0E080810120C181010101010101010101010101010101010101010101010101010101010101010101010101010"),
        new(3, 24, 22, 6, 1, // Geneva 24
            "001400000000141414101414140014141414141414141414141414141414141408060A120E1212060A0A100E080E060E1010101010101010101006080C0E0C1014121010100E0E10100A0E100E12101010101010101012160E0E0E080E080810080E0E0E0E0E0C0E0E0A0C0E0A160E0E0E0E0C0E0C0E0E160E0E0E0A060A1000" +
            "1212100E1010100E0E0E0E0E0E0E0E0E0E0E0A0A0A0A0E0E0E0E0E0E0E0E0E0E0A0C0E1010101012141416080A101810160E0C0C0E120E1012140A0C0C14160E10080E140A10121010141012121018160C120E0E08080E120E1E1414141414141414141414141414141414141414141414141414141414141414141414141414"),
        new(4, 9, 9, 2, 0, // Monaco 9
            "0000000000000606060606060600060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060600" +
            "0606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606060606"),
    ];

    private readonly string _widths;

    // A size the System has no bitmap for is drawn by QuickDraw from another size of the family, scaled by the Font
    // Manager's numer ÷ 256 (MacFontScale).
    private readonly MacFont? _bitmap;
    private readonly int _numer = MacFontScale.Denom;

    private MacFont(int number, int size, int ascent, int descent, int leading, string widths)
    {
        Number = number;
        Size = size;
        Ascent = ascent;
        Descent = descent;
        Leading = leading;
        _widths = widths;
    }

    private MacFont(MacFont bitmap, int size)
    {
        _bitmap = bitmap;
        _widths = bitmap._widths;
        Number = bitmap.Number;
        Size = size;
        _numer = MacFontScale.Numer(size, bitmap.Size);
        // GetFontInfo scales the font's metrics to the nearest pixel.
        Ascent = MacFontScale.Round(bitmap.Ascent, _numer);
        Descent = MacFontScale.Round(bitmap.Descent, _numer);
        Leading = MacFontScale.Round(bitmap.Leading, _numer);
    }

    /// <summary>The font number: 0 Chicago, 2 New York, 3 Geneva, 4 Monaco.</summary>
    public int Number { get; }

    /// <summary>The size in points.</summary>
    public int Size { get; }

    public int Ascent { get; }

    public int Descent { get; }

    public int Leading { get; }

    /// <summary>Whether the System has a bitmap of this size (QuickDraw's <c>RealFont</c>).</summary>
    public bool IsReal => _bitmap is null;

    /// <summary>The room Excel allows for the slant of italic text: half the ascent (spec §6, 0x248).</summary>
    public int ItalicExtra => Ascent / 2;

    /// <summary>The height of every row: ascent + descent + 1 (spec Appendix D).</summary>
    public int RowHeight => Ascent + Descent + 1;

    /// <summary>
    /// Finds the font Excel 1.00 draws with, or null for a family its System lacks. A size with no bitmap is
    /// scaled from the next larger size, or the largest there is (spec Appendix D).
    /// </summary>
    public static MacFont? Find(int number, int size)
    {
        MacFont? larger = null;
        MacFont? largest = null;
        foreach (var font in s_fonts)
        {
            if (font.Number != number)
            {
                continue;
            }

            if (font.Size == size)
            {
                return font;
            }

            if (font.Size > size && (larger is null || font.Size < larger.Size))
            {
                larger = font;
            }

            if (largest is null || font.Size > largest.Size)
            {
                largest = font;
            }
        }

        return size <= 0 || (larger ?? largest) is not { } bitmap ? null : new MacFont(bitmap, size);
    }

    /// <summary>Gets the advance width of a Mac Roman character, as <c>CharWidth</c> gives it.</summary>
    public int CharWidth(byte c, bool bold) => MacFontScale.Width(BitmapWidth(c) + (bold ? 1 : 0), _numer);

    /// <summary>Gets the width of text, as QuickDraw's <c>TextWidth</c> gives it.</summary>
    public int Width(ReadOnlySpan<char> text, bool bold)
    {
        Span<byte> bytes = text.Length <= 256 ? stackalloc byte[text.Length] : new byte[text.Length];
        MacRoman.GetBytes(text, bytes);
        int width = 0;
        foreach (byte b in bytes)
        {
            width += BitmapWidth(b);
        }

        // The bitmap's widths, a pixel more each for bold, add up before the total is scaled and rounded down.
        return MacFontScale.Width(width + (bold ? text.Length : 0), _numer);
    }

    private int BitmapWidth(byte c) => HexValue(_widths[2 * c]) * 16 + HexValue(_widths[2 * c + 1]);

    private static int HexValue(char c) => c <= '9' ? c - '0' : c - 'A' + 10;
}

/// <summary>
/// Measures text either in pixels with a font, or in units of a fixed width per character (1 for counting
/// characters, the cached digit width for a font the library has no metrics for).
/// </summary>
internal readonly struct TextMeasure
{
    private readonly MacFont? _font;
    private readonly int _unit;
    private readonly bool _bold;

    private readonly bool _fitPerCharacter;

    private TextMeasure(MacFont? font, int unit, bool bold, bool fitPerCharacter = false)
    {
        _font = font;
        _unit = unit;
        _bold = bold;
        _fitPerCharacter = fitPerCharacter;
    }

    /// <summary>Counts characters: every character is one unit wide.</summary>
    public static TextMeasure Characters => new(null, 1, false);

    /// <summary>
    /// Measures in pixels of <paramref name="font"/>. Releases measure alike, except that for a size with no bitmap
    /// Excel 1.03 adds a pixel per character when it checks whether text fits (spec Appendix D).
    /// </summary>
    public static TextMeasure Pixels(MacFont font, bool bold, ExcelVersion version = ExcelVersion.Excel100)
        => new(font, 1, bold, version == ExcelVersion.Excel103 && !font.IsReal);

    public static TextMeasure Fixed(int unit) => new(null, Math.Max(1, unit), false);

    public int Width(ReadOnlySpan<char> text) => _font?.Width(text, _bold) ?? text.Length * _unit;

    public int Width(char c) => Width([c]);

    /// <summary>The width Excel compares with the room to decide whether text fits.</summary>
    public int FitWidth(ReadOnlySpan<char> text) => Width(text) + (_fitPerCharacter ? text.Length : 0);

    // Excel adds a pixel to the widths of 0 and # it divides by when the font is scaled (09.1c, 19.0e).
    private int Adjust => _font is { IsReal: false } ? 1 : 0;

    /// <summary>The width of a digit, as Excel divides by it to count the characters that fit.</summary>
    public int DigitWidth => Width('0') + Adjust;

    /// <summary>The width of <c>#</c>, as Excel divides by it to fill a cell with them.</summary>
    public int HashWidth => Width('#') + Adjust;
}
