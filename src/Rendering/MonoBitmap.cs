using System.Buffers.Binary;
using System.IO.Compression;

namespace MacExcelReader.Rendering;

/// <summary>
/// A black-and-white picture, as a Macintosh screen of the time held one: one bit per pixel, rows of whole bytes,
/// the leftmost pixel in a byte's top bit, and a set bit black.
/// </summary>
public sealed class MonoBitmap
{
    /// <summary>
    /// Creates a white picture.
    /// </summary>
    public MonoBitmap(int width, int height)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        Width = width;
        Height = height;
        Stride = (width + 7) / 8;
        Bits = new byte[checked(Stride * height)];
    }

    /// <summary>Gets the width in pixels.</summary>
    public int Width { get; }

    /// <summary>Gets the height in pixels.</summary>
    public int Height { get; }

    /// <summary>Gets the number of bytes in each row.</summary>
    public int Stride { get; }

    /// <summary>Gets the pixels, row by row, <see cref="Stride"/> bytes each; a set bit is black.</summary>
    public byte[] Bits { get; }

    /// <summary>Gets or sets whether a pixel is black. Pixels outside the picture read as white and ignore writes.</summary>
    public bool this[int x, int y]
    {
        get => (uint)x < (uint)Width && (uint)y < (uint)Height && (Bits[y * Stride + (x >> 3)] & (0x80 >> (x & 7))) != 0;
        set
        {
            if ((uint)x < (uint)Width && (uint)y < (uint)Height)
            {
                ref byte b = ref Bits[y * Stride + (x >> 3)];
                b = value ? (byte)(b | (0x80 >> (x & 7))) : (byte)(b & ~(0x80 >> (x & 7)));
            }
        }
    }

    /// <summary>
    /// Fills a rectangle (left and top inclusive, right and bottom exclusive) with an 8 × 8 QuickDraw pattern: byte
    /// <c>y % 8</c> of the pattern, its top bit at <c>x % 8 == 0</c>, gives the row. Coordinates are in this picture;
    /// pass the offset of the picture's origin in the pattern's frame to align the pattern elsewhere.
    /// </summary>
    internal void FillPattern(int left, int top, int right, int bottom, ulong pattern, int patternX = 0, int patternY = 0)
    {
        ClipRect(ref left, ref top, ref right, ref bottom);
        for (int y = top; y < bottom; y++)
        {
            byte row = (byte)(pattern >> (8 * (7 - ((y + patternY) & 7))));
            // Rotate the pattern row so that bit 7 is the pattern's column for x = 0.
            int shift = (patternX & 7);
            byte aligned = (byte)((row << shift) | (row >> (8 - shift)));
            FillRow(y, left, right, aligned);
        }
    }

    /// <summary>Fills a rectangle black or white.</summary>
    internal void FillRect(int left, int top, int right, int bottom, bool black)
    {
        ClipRect(ref left, ref top, ref right, ref bottom);
        for (int y = top; y < bottom; y++)
        {
            FillRow(y, left, right, black ? (byte)0xFF : (byte)0);
        }
    }

    /// <summary>Inverts a rectangle, as QuickDraw's <c>InvertRect</c> does.</summary>
    internal void InvertRect(int left, int top, int right, int bottom)
    {
        ClipRect(ref left, ref top, ref right, ref bottom);
        for (int y = top; y < bottom; y++)
        {
            var row = Bits.AsSpan(y * Stride, Stride);
            for (int x = left; x < right;)
            {
                int bit = x & 7;
                if (bit == 0 && right - x >= 8)
                {
                    row[x >> 3] ^= 0xFF;
                    x += 8;
                }
                else
                {
                    row[x >> 3] ^= (byte)(0x80 >> bit);
                    x++;
                }
            }
        }
    }

    // Sets pixels left to right - 1 of a row to the pattern byte, whose top bit is the column x % 8 == 0.
    private void FillRow(int y, int left, int right, byte pattern)
    {
        if (left >= right)
        {
            return;
        }

        var row = Bits.AsSpan(y * Stride, Stride);
        int first = left >> 3, last = (right - 1) >> 3;
        byte firstMask = (byte)(0xFF >> (left & 7));
        byte lastMask = (byte)(0xFF << (7 - ((right - 1) & 7)));
        if (first == last)
        {
            byte mask = (byte)(firstMask & lastMask);
            row[first] = (byte)((row[first] & ~mask) | (pattern & mask));
            return;
        }

        row[first] = (byte)((row[first] & ~firstMask) | (pattern & firstMask));
        row.Slice(first + 1, last - first - 1).Fill(pattern);
        row[last] = (byte)((row[last] & ~lastMask) | (pattern & lastMask));
    }

    private void ClipRect(ref int left, ref int top, ref int right, ref int bottom)
    {
        left = Math.Max(left, 0);
        top = Math.Max(top, 0);
        right = Math.Min(right, Width);
        bottom = Math.Min(bottom, Height);
    }

    /// <summary>
    /// ORs a run of bits into a row: <paramref name="bits"/> holds them from its top bit, <paramref name="count"/> of
    /// them, the first going to <paramref name="x"/>. Only pixels from <paramref name="clipLeft"/> to
    /// <paramref name="clipRight"/> − 1 change.
    /// </summary>
    internal void OrBits(int x, int y, ulong bits, int count, int clipLeft, int clipRight)
    {
        if ((uint)y >= (uint)Height || bits == 0)
        {
            return;
        }

        clipLeft = Math.Max(clipLeft, 0);
        clipRight = Math.Min(clipRight, Width);
        if (x < clipLeft)
        {
            int skip = clipLeft - x;
            if (skip >= count)
            {
                return;
            }

            bits <<= skip;
            count -= skip;
            x = clipLeft;
        }

        count = Math.Min(count, clipRight - x);
        if (count <= 0)
        {
            return;
        }

        bits &= ulong.MaxValue << (64 - count);
        var row = Bits.AsSpan(y * Stride, Stride);
        int index = x >> 3;
        int shift = x & 7;

        // Up to 64 bits spread over at most 9 bytes.
        ulong high = bits >> shift;
        ulong low = shift == 0 ? 0 : bits << (64 - shift);
        for (int i = 0; i < 8 && index + i < row.Length; i++)
        {
            row[index + i] |= (byte)(high >> (56 - 8 * i));
        }

        if (low != 0 && index + 8 < row.Length)
        {
            row[index + 8] |= (byte)(low >> 56);
        }
    }

    /// <summary>
    /// Copies the part of another picture that falls inside this one, with its top-left corner at (x, y).
    /// </summary>
    public void Draw(MonoBitmap source, int x, int y)
    {
        ArgumentNullException.ThrowIfNull(source);
        for (int row = 0; row < source.Height; row++)
        {
            int ty = y + row;
            if ((uint)ty >= (uint)Height)
            {
                continue;
            }

            FillRect(x, ty, x + source.Width, ty + 1, false);
            for (int column = 0; column < source.Width; column += 64)
            {
                int count = Math.Min(64, source.Width - column);
                OrBits(x + column, ty, source.ReadBits(column, row, count), count, 0, Width);
            }
        }
    }

    // Reads up to 64 pixels of a row from x, into the top bits of the result.
    private ulong ReadBits(int x, int y, int count)
    {
        ulong bits = 0;
        var row = Bits.AsSpan(y * Stride, Stride);
        int index = x >> 3, shift = x & 7;
        for (int i = 0; i < 9 && index + i < row.Length; i++)
        {
            int at = 56 - 8 * i + shift;
            bits |= at >= 0 ? (ulong)row[index + i] << at : (ulong)row[index + i] >> -at;
        }

        return count >= 64 ? bits : bits & (ulong.MaxValue << (64 - count));
    }

    /// <summary>
    /// Writes the picture as a one-bit greyscale PNG.
    /// </summary>
    public void WritePng(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        Span<byte> header = stackalloc byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, Width);
        BinaryPrimitives.WriteInt32BigEndian(header[4..], Height);
        header[8] = 1;   // bit depth
        header[9] = 0;   // greyscale
        header[10] = 0;  // deflate
        header[11] = 0;  // adaptive filtering
        header[12] = 0;  // no interlace

        // In a greyscale PNG 0 is black, so the rows go in inverted, each after a filter byte of 0 (none).
        var pixels = new MemoryStream();
        using (var zlib = new ZLibStream(pixels, CompressionLevel.Optimal, leaveOpen: true))
        {
            var line = new byte[Stride + 1];
            for (int y = 0; y < Height; y++)
            {
                line[0] = 0;
                for (int i = 0; i < Stride; i++)
                {
                    line[i + 1] = (byte)~Bits[y * Stride + i];
                }

                zlib.Write(line);
            }
        }

        stream.Write([0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A]);
        WriteChunk(stream, "IHDR"u8, header);
        WriteChunk(stream, "IDAT"u8, pixels.GetBuffer().AsSpan(0, (int)pixels.Length));
        WriteChunk(stream, "IEND"u8, []);
    }

    /// <summary>
    /// Gets the picture as a one-bit greyscale PNG.
    /// </summary>
    public byte[] ToPng()
    {
        var stream = new MemoryStream();
        WritePng(stream);
        return stream.ToArray();
    }

    private static void WriteChunk(Stream stream, ReadOnlySpan<byte> type, ReadOnlySpan<byte> data)
    {
        Span<byte> word = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(word, data.Length);
        stream.Write(word);
        stream.Write(type);
        stream.Write(data);
        uint crc = Crc32.Update(Crc32.Update(0xFFFFFFFF, type), data) ^ 0xFFFFFFFF;
        BinaryPrimitives.WriteUInt32BigEndian(word, crc);
        stream.Write(word);
    }

    private static class Crc32
    {
        private static readonly uint[] s_table = CreateTable();

        public static uint Update(uint crc, ReadOnlySpan<byte> data)
        {
            foreach (byte b in data)
            {
                crc = s_table[(crc ^ b) & 0xFF] ^ (crc >> 8);
            }

            return crc;
        }

        private static uint[] CreateTable()
        {
            var table = new uint[256];
            for (uint n = 0; n < 256; n++)
            {
                uint c = n;
                for (int k = 0; k < 8; k++)
                {
                    c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
                }

                table[n] = c;
            }

            return table;
        }
    }
}
