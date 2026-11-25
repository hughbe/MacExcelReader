using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader.Rendering;

/// <summary>
/// A Macintosh bitmap font of one family and size, read from a <c>FONT</c> resource (Inside Macintosh, "The Font
/// Manager"), drawn as QuickDraw draws text in its plain, bold and italic styles.
/// </summary>
public sealed class MacBitmapFont
{
    // Each glyph's rows as bits from the top of a ulong (glyphs are at most 64 pixels wide), Height rows per glyph,
    // in Mac Roman order; characters the font lacks use its missing-character glyph.
    private readonly ulong[] _rows;
    private readonly byte[] _imageWidths = new byte[256];
    private readonly sbyte[] _origins = new sbyte[256];

    // Advance widths in 16.16 fixed point: whole pixels for a real size, fractions for a scaled one.
    private readonly int[] _advances = new int[256];

    private MacBitmapFont(int family, int size, int ascent, int descent, int leading, int height, ulong[] rows, bool isReal)
    {
        Family = family;
        Size = size;
        Ascent = ascent;
        Descent = descent;
        Leading = leading;
        Height = height;
        _rows = rows;
        IsReal = isReal;
    }

    /// <summary>Gets the font family number: 0 Chicago, 2 New York, 3 Geneva, 4 Monaco and so on.</summary>
    public int Family { get; }

    /// <summary>Gets the size in points.</summary>
    public int Size { get; }

    /// <summary>Gets the ascent: how many pixels the font reaches above its baseline.</summary>
    public int Ascent { get; }

    /// <summary>Gets the descent: how many pixels the font reaches below its baseline.</summary>
    public int Descent { get; }

    /// <summary>Gets the leading: the pixels the font asks for between lines.</summary>
    public int Leading { get; }

    /// <summary>Gets whether the size came from a bitmap of its own, rather than being scaled from another size.</summary>
    public bool IsReal { get; }

    // Rows in each glyph's image: the font rectangle's height.
    private int Height { get; }

    /// <summary>
    /// Reads a <c>FONT</c> or <c>NFNT</c> resource.
    /// </summary>
    /// <param name="family">The family number (the resource ID ÷ 128 for a <c>FONT</c>).</param>
    /// <param name="size">The size (the resource ID mod 128 for a <c>FONT</c>).</param>
    /// <param name="resource">The resource's data.</param>
    public static MacBitmapFont Read(int family, int size, ReadOnlySpan<byte> resource)
    {
        if (resource.Length < 26)
        {
            throw new ExcelFormatException("A FONT resource has a 26-byte header.");
        }

        int firstChar = Word(resource, 2), lastChar = Word(resource, 4);
        int kernMax = Word(resource, 8);
        int rectHeight = Word(resource, 14);
        int owTableOffset = 16 + 2 * (Word(resource, 16) & 0xFFFF);
        int ascent = Word(resource, 18), descent = Word(resource, 20), leading = Word(resource, 22), rowWords = Word(resource, 24);
        if (firstChar < 0 || lastChar > 255 || lastChar < firstChar || rectHeight is <= 0 or > 255 || rowWords <= 0)
        {
            throw new ExcelFormatException("A FONT resource's header is not valid.");
        }

        int strikeOffset = 26;
        int rowBytes = 2 * rowWords;
        int locOffset = strikeOffset + rowBytes * rectHeight;
        int glyphs = lastChar - firstChar + 2;   // the characters and the missing-character glyph
        if (locOffset + 2 * (glyphs + 1) > resource.Length || owTableOffset + 2 * glyphs > resource.Length)
        {
            throw new ExcelFormatException("A FONT resource's tables run past its end.");
        }

        var font = new MacBitmapFont(family, size, ascent, descent, leading, rectHeight, new ulong[256 * rectHeight], isReal: true);
        var strike = resource.Slice(strikeOffset, rowBytes * rectHeight);
        int missing = glyphs - 1;
        for (int c = 0; c < 256; c++)
        {
            int glyph = c >= firstChar && c <= lastChar && Word(resource, owTableOffset + 2 * (c - firstChar)) != -1 ? c - firstChar : missing;
            int offsetWidth = Word(resource, owTableOffset + 2 * glyph);
            if (offsetWidth == -1)
            {
                continue;   // no missing-character glyph either: nothing drawn, no advance
            }

            int location = Word(resource, locOffset + 2 * glyph) & 0xFFFF;
            int imageWidth = (Word(resource, locOffset + 2 * glyph + 2) & 0xFFFF) - location;
            if (imageWidth is < 0 or > 64 || location + imageWidth > rowBytes * 8)
            {
                throw new ExcelFormatException("A FONT resource's glyph lies outside its strike.");
            }

            font._imageWidths[c] = (byte)imageWidth;
            font._origins[c] = (sbyte)(kernMax + ((offsetWidth >> 8) & 0xFF));
            font._advances[c] = (offsetWidth & 0xFF) << 16;
            for (int row = 0; row < rectHeight; row++)
            {
                font._rows[c * rectHeight + row] = ReadStrikeBits(strike.Slice(row * rowBytes, rowBytes), location, imageWidth);
            }
        }

        return font;
    }

    private static int Word(ReadOnlySpan<byte> resource, int offset) => (uint)offset + 2 <= (uint)resource.Length
        ? BinaryPrimitives.ReadInt16BigEndian(resource[offset..])
        : throw new ExcelFormatException("A FONT resource runs past its end.");

    private static long FloorDiv(long a, long b) => a >= 0 ? a / b : -((-a + b - 1) / b);

    // Reads `count` bits of a strike row from bit `location`, into the top of a ulong.
    private static ulong ReadStrikeBits(ReadOnlySpan<byte> row, int location, int count)
    {
        ulong bits = 0;
        for (int i = 0; i < count; i++)
        {
            int bit = location + i;
            if ((row[bit >> 3] & (0x80 >> (bit & 7))) != 0)
            {
                bits |= 1UL << (63 - i);
            }
        }

        return bits;
    }

    /// <summary>
    /// Makes a size the System has no bitmap for from this one, as QuickDraw does (spec Appendix D): the advance widths
    /// scaled by size ÷ this size and kept in fractions of a pixel, the ascent and descent rounded to the nearest pixel,
    /// and text drawn at this size and then stretched to the new one.
    /// </summary>
    internal MacBitmapFont Scale(int size)
    {
        int Round(int value) => (2 * value * size + Size) / (2 * Size);
        return new MacBitmapFont(Family, size, Round(Ascent), Round(Descent), Round(Leading), Height, _rows, isReal: false)
        {
            _source = this,
        };
    }

    // The bitmap font a scaled size is drawn from.
    private MacBitmapFont? _source;

    // Draws text in a scaled size: at the bitmap's size into a buffer, then stretched to this size, each pixel taking
    // the bitmap pixel under its centre. This comes close to QuickDraw but is not exact: Excel 1.00's Geneva 11 differs
    // from it in about one pixel in fifty where glyphs fall between pixels.
    private int DrawScaled(MacBitmapFont source, MonoBitmap target, int x, int baseline, ReadOnlySpan<char> text, bool bold, bool italic,
        int clipLeft, int clipTop, int clipRight, int clipBottom)
    {
        const int margin = 16;
        int sourceWidth = source.MeasureText(text, bold) + 2 * margin + source.Height;
        var buffer = new MonoBitmap(Math.Max(1, sourceWidth), source.Height);
        source.DrawText(buffer, margin, source.Ascent, text, bold, italic);

        int S = source.Size, D = Size;
        int top = baseline - Ascent;
        int height = (int)Math.Min(int.MaxValue, ((long)source.Height * D + S - 1) / S);
        int width = (int)Math.Min(int.MaxValue, ((long)sourceWidth * D + S - 1) / S);
        int dMargin = margin * D / S + 1;
        int originX = x - dMargin;

        // Only the pixels inside the clip rectangle and the target are worked out, however large the size.
        int firstRow = Math.Max(0, Math.Max(clipTop, 0) - top), lastRow = Math.Min(height, Math.Min(clipBottom, target.Height) - top);
        int firstColumn = Math.Max(0, Math.Max(clipLeft, 0) - originX), lastColumn = Math.Min(width, Math.Min(clipRight, target.Width) - originX);
        for (int dy = firstRow; dy < lastRow; dy++)
        {
            int sy = (int)Math.Min(source.Height - 1, (2L * dy + 1) * S / (2 * D));
            for (int dx = firstColumn; dx < lastColumn; dx++)
            {
                int sxp = margin + (int)FloorDiv((2L * (dx - dMargin) + 1) * S, 2L * D);
                if (sxp >= 0 && sxp < buffer.Width && buffer[sxp, sy])
                {
                    target[originX + dx, top + dy] = true;
                }
            }
        }

        return MeasureText(text, bold);
    }

    /// <summary>
    /// Gets the width of text in pixels, as QuickDraw's <c>TextWidth</c> gives it: bold adds a pixel per character,
    /// and the fractions of a scaled size add up before the total is rounded down.
    /// </summary>
    public int MeasureText(ReadOnlySpan<char> text, bool bold = false)
    {
        if (_source is { } source)
        {
            // The bitmap's widths add up in fractions of a pixel at this size, and the total is rounded down.
            return (int)((long)(source.MeasureText(text) * Size) / source.Size) + (bold ? text.Length : 0);
        }

        int width = 0;
        foreach (char c in text)
        {
            width += _advances[MacRoman.GetByte(c)];
        }

        return (width >> 16) + (bold ? text.Length : 0);
    }

    /// <summary>
    /// Draws text with its baseline at <paramref name="baseline"/>, starting at <paramref name="x"/>, and returns the
    /// width it advanced. Pixels outside the clip rectangle (left and top inclusive, right and bottom exclusive) stay
    /// as they were. Bold smears each glyph a pixel to the right and widens it by one; italic slants each row right
    /// by half its height above the baseline.
    /// </summary>
    public int DrawText(MonoBitmap target, int x, int baseline, ReadOnlySpan<char> text, bool bold = false, bool italic = false)
        => DrawText(target, x, baseline, text, bold, italic, 0, 0, target.Width, target.Height);

    /// <inheritdoc cref="DrawText(MonoBitmap, int, int, ReadOnlySpan{char}, bool, bool)"/>
    public int DrawText(MonoBitmap target, int x, int baseline, ReadOnlySpan<char> text, bool bold, bool italic,
        int clipLeft, int clipTop, int clipRight, int clipBottom)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (_source is { } source)
        {
            return DrawScaled(source, target, x, baseline, text, bold, italic, clipLeft, clipTop, clipRight, clipBottom);
        }

        int top = baseline - Ascent;
        int firstRow = Math.Max(0, clipTop - top);
        int lastRow = Math.Min(Height, clipBottom - top);
        int pen = x << 16;
        foreach (char ch in text)
        {
            byte c = MacRoman.GetByte(ch);
            int glyphX = (pen >> 16) + _origins[c];
            int width = _imageWidths[c];
            for (int row = firstRow; row < lastRow; row++)
            {
                ulong bits = _rows[c * Height + row];
                if (bits == 0)
                {
                    continue;
                }

                int count = width;
                if (bold)
                {
                    bits |= bits >> 1;
                    count++;
                }

                // QuickDraw's italic: a row moves right by half its distance from the bottom row of the font's
                // rectangle, rounded down.
                int slant = italic ? (Height - 1 - row) >> 1 : 0;
                target.OrBits(glyphX + slant, top + row, bits, Math.Min(64, count), clipLeft, clipRight);
            }

            pen += _advances[c] + (bold ? 1 << 16 : 0);
        }

        return (pen >> 16) - x;
    }
}
