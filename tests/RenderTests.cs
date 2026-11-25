using System.Buffers.Binary;
using System.IO.Compression;
using MacExcelReader.Rendering;

namespace MacExcelReader.Tests;

/// <summary>
/// Checks the renderer with a font made here, so they run without a System file: font parsing, text drawing, the PNG
/// writer and the window's geometry. <see cref="ExperimentTests"/> compares whole windows with Excel 1.00's screen
/// when <c>MACEXCEL_SYSTEM</c> names Excel's fonts.
/// </summary>
public class RenderTests
{
    // A font where every printable character is a 5 × 8 box resting on the baseline, one pixel in from its origin, with
    // an advance of 7; Geneva 10's ascent, descent and family number, so a default sheet draws with it.
    internal static byte[] BoxFontResource()
    {
        const int first = 32, last = 126, glyphs = last - first + 2, width = 5, height = 12, ascent = 10;
        int rowWords = (glyphs * width + 15) / 16;
        int strike = 26, loc = strike + rowWords * 2 * height, ow = loc + 2 * (glyphs + 1);
        var font = new byte[ow + 2 * (glyphs + 1)];
        void Word(int offset, int value) => BinaryPrimitives.WriteInt16BigEndian(font.AsSpan(offset), (short)value);
        Word(0, unchecked((short)0x9000));
        Word(2, first);
        Word(4, last);
        Word(6, 7);
        Word(8, 0);
        Word(10, -2);
        Word(12, 7);
        Word(14, height);
        Word(16, (ow - 16) / 2);
        Word(18, ascent);
        Word(20, 2);
        Word(22, 0);
        Word(24, rowWords);
        for (int g = 0; g < glyphs; g++)
        {
            for (int row = 2; row < ascent; row++)
            {
                for (int x = 0; x < width; x++)
                {
                    int bit = g * width + x;
                    font[strike + row * rowWords * 2 + bit / 8] |= (byte)(0x80 >> (bit % 8));
                }
            }

            Word(loc + 2 * g, g * width);
            Word(ow + 2 * g, (1 << 8) | 7);
        }

        Word(loc + 2 * glyphs, glyphs * width);
        Word(ow + 2 * glyphs, -1);
        return font;
    }

    // A resource fork holding one FONT resource.
    internal static byte[] ResourceFork(short id, byte[] resource)
    {
        const int dataOffset = 256;
        int mapOffset = dataOffset + 4 + resource.Length;
        var fork = new byte[mapOffset + 28 + 2 + 8 + 12];
        BinaryPrimitives.WriteInt32BigEndian(fork, dataOffset);
        BinaryPrimitives.WriteInt32BigEndian(fork.AsSpan(4), mapOffset);
        BinaryPrimitives.WriteInt32BigEndian(fork.AsSpan(8), 4 + resource.Length);
        BinaryPrimitives.WriteInt32BigEndian(fork.AsSpan(12), fork.Length - mapOffset);
        BinaryPrimitives.WriteInt32BigEndian(fork.AsSpan(dataOffset), resource.Length);
        resource.CopyTo(fork, dataOffset + 4);
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(mapOffset + 24), 28);
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(mapOffset + 26), (ushort)(fork.Length - mapOffset));
        int types = mapOffset + 28;
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(types), 0);
        "FONT"u8.CopyTo(fork.AsSpan(types + 2));
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(types + 6), 0);
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(types + 8), 10);
        BinaryPrimitives.WriteInt16BigEndian(fork.AsSpan(types + 10), id);
        BinaryPrimitives.WriteInt16BigEndian(fork.AsSpan(types + 12), -1);
        return fork;
    }

    internal static MacFontSet BoxFonts() => MacFontSet.FromResourceFork(ResourceFork(3 * 128 + 10, BoxFontResource()));

    [Fact]
    public void FontSet_ReadsTheFontResource()
    {
        var font = BoxFonts().GetFont(3, 10)!;
        Assert.True(font.IsReal);
        Assert.Equal((3, 10, 10, 2), (font.Family, font.Size, font.Ascent, font.Descent));
        Assert.Equal(14, font.MeasureText("ab"));
        Assert.Equal(16, font.MeasureText("ab", bold: true));
        Assert.Null(BoxFonts().GetFont(4, 9));
    }

    [Fact]
    public void FontSet_ScalesASizeWithNoBitmap()
    {
        // From the largest size there is, the widths kept in fractions: 7 × 15/10 = 10.5 a character.
        var font = BoxFonts().GetFont(3, 15)!;
        Assert.False(font.IsReal);
        Assert.Equal(15, font.Ascent);
        Assert.Equal(21, font.MeasureText("ab"));
        Assert.Equal(31, font.MeasureText("abc"));
    }

    [Fact]
    public void FontSet_RejectsAForkWithoutFonts()
        => Assert.Throws<ExcelFormatException>(() => MacFontSet.FromResourceFork(new byte[64]));

    [Fact]
    public void DrawText_PlainBoldAndItalic()
    {
        var font = BoxFonts().GetFont(3, 10)!;
        var bitmap = new MonoBitmap(64, 20);

        // Plain: the box covers x + 1 to x + 5 and the 8 rows above the baseline.
        Assert.Equal(7, font.DrawText(bitmap, 10, 15, "a"));
        Assert.True(bitmap[11, 7] && bitmap[15, 14]);
        Assert.False(bitmap[10, 10] || bitmap[16, 10] || bitmap[12, 6] || bitmap[12, 15]);

        // Bold smears a pixel right and advances one more.
        bitmap = new MonoBitmap(64, 20);
        Assert.Equal(8, font.DrawText(bitmap, 10, 15, "a", bold: true));
        Assert.True(bitmap[16, 10]);
        Assert.False(bitmap[17, 10]);

        // Italic: a row moves right by half its distance from the font rectangle's bottom (row 11 of 12).
        bitmap = new MonoBitmap(64, 20);
        font.DrawText(bitmap, 10, 15, "a", italic: true);
        Assert.True(bitmap[11 + 4, 7] && !bitmap[11 + 3, 7]);   // row 2 of the glyph: (11 - 2) / 2 = 4
        Assert.True(bitmap[11 + 1, 14] && !bitmap[11, 14]);     // row 9: (11 - 9) / 2 = 1
    }

    [Fact]
    public void DrawText_ClipsToTheRectangle()
    {
        var font = BoxFonts().GetFont(3, 10)!;
        var bitmap = new MonoBitmap(64, 20);
        font.DrawText(bitmap, 10, 15, "aa", false, false, 13, 0, 19, 20);
        Assert.False(bitmap[12, 10]);
        Assert.True(bitmap[13, 10] && bitmap[18, 10]);
        Assert.False(bitmap[19, 10]);
    }

    [Fact]
    public void Png_RoundTrips()
    {
        var bitmap = new MonoBitmap(13, 5);
        bitmap[0, 0] = bitmap[12, 4] = bitmap[7, 2] = true;
        var read = ReadPng(bitmap.ToPng());
        Assert.Equal((13, 5), (read.Width, read.Height));
        Assert.Equal(bitmap.Bits, read.Bits);
    }

    [Fact]
    public void Render_DefaultWorksheetGeometry()
    {
        // The default window: content 493 × 274, so 496 × 295 with the frame and shadow.
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["B2"].SetText("x");
        var bitmap = ExcelRenderer.Render(document, BoxFonts());
        Assert.Equal((496, 295), (bitmap.Width, bitmap.Height));

        // Content starts at (1, 19). The row-heading strip is 29 wide, its line at x 28; the heading row 13 high,
        // its line at y 12; column A's gridline at 28 + 75 = 103, dotted where x + y is even.
        int X(int x) => 1 + x;
        int Y(int y) => 19 + y;
        Assert.True(bitmap[X(28), Y(100)]);
        Assert.True(bitmap[X(50), Y(12)]);
        Assert.True(bitmap[X(103), Y(41)] && !bitmap[X(103), Y(40)]);
        Assert.True(bitmap[X(111), Y(25)] && !bitmap[X(110), Y(25)]);

        // B2's text: left aligned two pixels into the cell (x 104 + 2, plus the glyph's origin), baseline at 26 + 10.
        Assert.True(bitmap[X(104 + 2 + 1), Y(35)]);

        // The active cell A1 is outlined by its inverted frame, clipped at the heading strip: its first column of
        // pixels is black and the next white.
        Assert.True(bitmap[X(29), Y(20)] && !bitmap[X(30), Y(20)]);
    }

    [Fact]
    public void Render_PastTheLastRowIsGrey()
    {
        var document = ExcelDocument.CreateWorksheet();
        document.Windows[0].TopRow = 16383;
        var bitmap = ExcelRenderer.Render(document, BoxFonts(), new ExcelRenderOptions { Frame = false, Selection = false });

        // Row 16384 covers y 13 to 25 with a solid edge at 25; below it everything is grey, the headings included.
        Assert.True(bitmap[200, 25] && bitmap[201, 25]);
        Assert.True(bitmap[10, 30] != bitmap[11, 30]);
        Assert.True(bitmap[200, 30] != bitmap[200, 31]);
    }

    [Fact]
    public void Render_RejectsAMissingWindowOrFont()
    {
        var document = ExcelDocument.CreateWorksheet();
        Assert.Throws<InvalidOperationException>(() => ExcelRenderer.Render(document, BoxFonts(), new ExcelRenderOptions { Window = 1 }));
        document.State.StandardFont = 2;
        Assert.Throws<InvalidOperationException>(() => ExcelRenderer.Render(document, BoxFonts()));
    }

    [Fact]
    public void Render_DamagedFontSizeAndRectangleStayCheap()
    {
        // A damaged file can ask for a 32,778-point font in a window of 30,000 pixels. Excel fits the window to the
        // screen, and only the pixels that show are drawn, so this takes milliseconds.
        var document = ExcelDocument.CreateWorksheet();
        document.Cells["A1"].SetText("text");
        document.State.StandardFontSize = 32778;
        document.Windows[0].Right = 30000;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var bitmap = ExcelRenderer.Render(document, BoxFonts());
        Assert.True(bitmap.Width <= 2048 + 3);
        Assert.True(watch.ElapsedMilliseconds < 2000, $"{watch.ElapsedMilliseconds} ms");
    }

    // Where Excel 1.00's vertical thumb sat for each top row, in a window of the default height (208 pixels of travel).
    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 0)]
    [InlineData(2, 1)]
    [InlineData(3, 1)]
    [InlineData(4, 2)]
    [InlineData(8, 4)]
    [InlineData(16, 8)]
    [InlineData(32, 10)]
    [InlineData(64, 10)]
    [InlineData(100, 10)]
    [InlineData(150, 11)]
    [InlineData(200, 12)]
    [InlineData(500, 15)]
    [InlineData(1000, 21)]
    [InlineData(2000, 32)]
    [InlineData(8000, 102)]
    [InlineData(16000, 194)]
    [InlineData(16383, 208)]
    public void VerticalThumb_MatchesExcel(int topRow, int offset) => Assert.Equal(offset, ExcelRenderer.VerticalThumbOffset(topRow, 208));

    // And the horizontal thumb, 427 pixels of travel.
    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 2)]
    [InlineData(10, 17)]
    [InlineData(100, 167)]
    [InlineData(240, 402)]
    public void HorizontalThumb_MatchesExcel(int leftColumn, int offset) => Assert.Equal(offset, ExcelRenderer.HorizontalThumbOffset(leftColumn, 427));

    /// <summary>
    /// Reads a one-bit greyscale PNG without interlacing, as <see cref="MonoBitmap.WritePng"/> writes them.
    /// </summary>
    internal static MonoBitmap ReadPng(byte[] png)
    {
        int position = 8, width = 0, height = 0;
        var data = new MemoryStream();
        while (position < png.Length)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(position));
            var type = png.AsSpan(position + 4, 4);
            var chunk = png.AsSpan(position + 8, length);
            if (type.SequenceEqual("IHDR"u8))
            {
                width = BinaryPrimitives.ReadInt32BigEndian(chunk);
                height = BinaryPrimitives.ReadInt32BigEndian(chunk[4..]);
                Assert.Equal(1, chunk[8]);
                Assert.Equal(0, chunk[9]);
            }
            else if (type.SequenceEqual("IDAT"u8))
            {
                data.Write(chunk);
            }

            position += 12 + length;
        }

        data.Position = 0;
        using var zlib = new ZLibStream(data, CompressionMode.Decompress);
        var raw = new MemoryStream();
        zlib.CopyTo(raw);
        var bitmap = new MonoBitmap(width, height);
        var bytes = raw.ToArray();
        for (int y = 0; y < height; y++)
        {
            Assert.Equal(0, bytes[y * (bitmap.Stride + 1)]);
            for (int i = 0; i < bitmap.Stride; i++)
            {
                bitmap.Bits[y * bitmap.Stride + i] = (byte)~bytes[y * (bitmap.Stride + 1) + 1 + i];
            }
        }

        // Bits past the right edge are padding.
        if (width % 8 != 0)
        {
            for (int y = 0; y < height; y++)
            {
                bitmap.Bits[y * bitmap.Stride + bitmap.Stride - 1] &= (byte)(0xFF << (8 - width % 8));
            }
        }

        return bitmap;
    }
}
