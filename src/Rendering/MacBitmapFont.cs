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
    /// Makes a size the System has no bitmap for from this one, as the Mac Plus ROM does (docs/Display.md): the Font
    /// Manager's scale is numer ÷ 256 (<see cref="MacFontScale"/>), the ascent, descent and leading scale to the nearest
    /// pixel, widths add up at this size before they are scaled, and text is drawn at this size and stretched.
    /// </summary>
    internal MacBitmapFont Scale(int size)
    {
        int numer = MacFontScale.Numer(size, Size);
        return new MacBitmapFont(Family, size, MacFontScale.Round(Ascent, numer), MacFontScale.Round(Descent, numer),
            MacFontScale.Round(Leading, numer), Height, _rows, isReal: false)
        {
            _source = this,
            _numer = numer,
        };
    }

    // The bitmap font a scaled size is drawn from, and the Font Manager's numer for it (over 256).
    private MacBitmapFont? _source;
    private int _numer;

    // Draws text in a scaled size as QuickDraw's DrText does when the Font Manager asks for stretching (Mac Plus ROM;
    // DrawText.a in Apple's QuickDraw source). DrText draws the text at the bitmap's size, from the pen, into a buffer
    // as tall as the font's rectangle and as wide as the text plus 32 pixels of slop for bold and italic; maps that
    // rectangle to the new size with MapRect; and copies the buffer into the mapped rectangle with StretchBits, which
    // ORs together the source rows, and the source columns, that fall into one destination row or column.
    private int DrawScaled(MacBitmapFont source, MonoBitmap target, int x, int baseline, ReadOnlySpan<char> text, bool bold, bool italic,
        int clipLeft, int clipTop, int clipRight, int clipBottom)
    {
        // StdTxMeas: the bitmap's widths, a pixel more each for bold, added up unscaled.
        int unscaled = source.MeasureText(text, bold);
        int advance = MacFontScale.Width(unscaled, _numer);
        int srcWidth = unscaled + 32, srcHeight = source.Height;
        int srcTop = baseline - source.Ascent;

        // MapRect moves each edge's distance from the pen by numer ÷ 256, rounded to the nearest pixel.
        int dstLeft = x, dstRight = x + MacFontScale.Round(srcWidth, _numer);
        int dstTop = baseline + MacFontScale.Round(srcTop - baseline, _numer);
        int dstBottom = baseline + MacFontScale.Round(srcTop + srcHeight - baseline, _numer);
        int dstWidth = dstRight - dstLeft, dstHeight = dstBottom - dstTop;

        // StretchBits draws only inside the destination rectangle, the clip rectangle and the target.
        int minLeft = Math.Max(Math.Max(clipLeft, 0), dstLeft), minRight = Math.Min(Math.Min(clipRight, target.Width), dstRight);
        int minTop = Math.Max(Math.Max(clipTop, 0), dstTop), minBottom = Math.Min(Math.Min(clipBottom, target.Height), dstBottom);
        if (minLeft >= minRight || minTop >= minBottom)
        {
            return advance;
        }

        // The buffer, with column 0 at the pen: pixels kerned left of the pen are not copied.
        const int margin = 64;
        var buffer = new MonoBitmap(srcWidth + margin, srcHeight);
        source.DrawText(buffer, margin, source.Ascent, text, bold, italic);

        // Horizontally StretchBits picks a routine from FixRatio's fraction of the smaller width over the larger
        // (SetupStretch). The general ones step a 16-bit error term, starting at half the fraction, by the fraction:
        // when shrinking (SHRINK) each source column goes to the destination column the carries have reached, and when
        // stretching (STRETCH) each destination column takes the source column they have reached. The fast routines for
        // exact ratios repeat each source column 2, 3, 4, 6, 8, 16 or any multiple of 8 times, and stretch to 1.5 times
        // by giving each 4 source columns 1, 2, 1 and 2 destination columns (ONE.5). Shrinking to ½, ¼ and ¾ has fast
        // routines too, which give what SHRINK gives.
        bool shrinking = dstWidth < srcWidth;
        int ratio = (int)(shrinking ? MacFontScale.FixRatio(dstWidth, srcWidth) : MacFontScale.FixRatio(srcWidth, dstWidth)) & 0xFFFF;
        int repeat = dstWidth == srcWidth ? 1 : shrinking ? 0 : ratio switch
        {
            0x8000 => 2,
            0x5555 => 3,
            0x4000 => 4,
            0x2AAA => 6,
            0x2000 => 8,
            0x1000 => 16,
            _ => dstWidth % srcWidth == 0 && dstWidth / srcWidth % 8 == 0 ? dstWidth / srcWidth : 0,
        };
        int Step(int i) => repeat > 0 ? i / repeat
            : !shrinking && ratio == 0xAAAA ? i / 6 * 4 + (0b11_11_10_01_01_00 >> (i % 6 * 2) & 3)
            : (int)(((ratio >> 1) + (long)i * ratio) >> 16);

        // Vertically a DDA with an error term starting at minus half the source height: each destination row is the
        // OR of the source rows read since the last, and is repeated while the error stays at or above zero.
        var row = new bool[srcWidth];
        var dstRow = new bool[dstWidth];
        int error = -(srcHeight >> 1), sourceRow = 0, vertical = dstTop;
        while (sourceRow < srcHeight)
        {
            Array.Clear(row);
            do
            {
                for (int i = 0; i < srcWidth; i++)
                {
                    row[i] |= buffer[margin + i, sourceRow];
                }

                sourceRow++;
                error += dstHeight;
            }
            while (error <= 0 && sourceRow < srcHeight);

            Array.Clear(dstRow);
            if (shrinking)
            {
                for (int i = 0; i < srcWidth; i++)
                {
                    if (row[i])
                    {
                        dstRow[Step(i)] = true;
                    }
                }
            }
            else
            {
                for (int i = minLeft - dstLeft; i < minRight - dstLeft; i++)
                {
                    int column = Step(i);
                    dstRow[i] = column < srcWidth && row[column];
                }
            }

            do
            {
                if (vertical >= minTop)
                {
                    for (int px = minLeft; px < minRight; px++)
                    {
                        if (dstRow[px - dstLeft])
                        {
                            target[px, vertical] = true;
                        }
                    }
                }

                if (++vertical >= minBottom)
                {
                    return advance;
                }

                error -= srcHeight;
            }
            while (error >= 0);
        }

        return advance;
    }

    /// <summary>
    /// Gets the width of text in pixels, as QuickDraw's <c>TextWidth</c> gives it: bold adds a pixel per character, and
    /// in a scaled size the bitmap's widths add up before the total is scaled and rounded down.
    /// </summary>
    public int MeasureText(ReadOnlySpan<char> text, bool bold = false)
    {
        if (_source is { } source)
        {
            // TextWidth adds up the bitmap's widths (a pixel more each for bold) and scales the total, rounding down.
            return MacFontScale.Width(source.MeasureText(text, bold), _numer);
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
