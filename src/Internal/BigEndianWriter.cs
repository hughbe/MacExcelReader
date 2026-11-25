using System.Buffers.Binary;

namespace MacExcelReader.Internal;

/// <summary>
/// Growable big-endian writer.
/// </summary>
internal sealed class BigEndianWriter : IDisposable
{
    private byte[] _buffer;

    // Whether the buffer is rented from the shared pool, to be returned by Dispose.
    private readonly bool _pooled;

    public BigEndianWriter(int capacity = 4096, bool pooled = false)
    {
        _pooled = pooled;
        capacity = Math.Max(16, capacity);
        _buffer = pooled ? System.Buffers.ArrayPool<byte>.Shared.Rent(capacity) : new byte[capacity];
    }

    public void Dispose()
    {
        if (_pooled && _buffer.Length > 0)
        {
            System.Buffers.ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = [];
        }
    }

    public ReadOnlySpan<byte> WrittenSpan => _buffer.AsSpan(0, Position);

    public ReadOnlyMemory<byte> WrittenMemory => _buffer.AsMemory(0, Position);

    public int Position { get; private set; }

    private Span<byte> Reserve(int count)
    {
        if (Position + count > _buffer.Length)
        {
            int size = Math.Max(_buffer.Length + _buffer.Length / 2, Position + count);
            if (_pooled)
            {
                var larger = System.Buffers.ArrayPool<byte>.Shared.Rent(size);
                _buffer.AsSpan(0, Position).CopyTo(larger);
                System.Buffers.ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = larger;
            }
            else
            {
                Array.Resize(ref _buffer, size);
            }
        }

        var span = _buffer.AsSpan(Position, count);
        Position += count;
        return span;
    }

    public void WriteByte(byte value) => Reserve(1)[0] = value;

    public void WriteUInt16(ushort value) => BinaryPrimitives.WriteUInt16BigEndian(Reserve(2), value);

    public void WriteInt16(short value) => BinaryPrimitives.WriteInt16BigEndian(Reserve(2), value);

    public void WriteUInt32(uint value) => BinaryPrimitives.WriteUInt32BigEndian(Reserve(4), value);

    public void WriteDouble(double value) => BinaryPrimitives.WriteDoubleBigEndian(Reserve(8), value);

    public void WriteBytes(ReadOnlySpan<byte> bytes) => bytes.CopyTo(Reserve(bytes.Length));

    public void WriteZeros(int count) => Reserve(count).Clear();

    /// <summary>
    /// Writes a Pascal string: a length byte followed by the bytes.
    /// </summary>
    public void WritePascalBytes(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length > 255)
        {
            throw new ArgumentException("A Pascal string cannot be longer than 255 bytes.", nameof(bytes));
        }

        WriteByte((byte)bytes.Length);
        WriteBytes(bytes);
    }

    /// <summary>
    /// Writes a Pascal string, encoding the text as Mac OS Roman (one byte per character) straight into the buffer.
    /// </summary>
    public void WritePascalString(string value)
    {
        if (value.Length > 255)
        {
            throw new ArgumentException("A Pascal string cannot be longer than 255 bytes.", nameof(value));
        }

        WriteByte((byte)value.Length);
        WriteMacRoman(value);
    }

    /// <summary>
    /// Writes text as Mac OS Roman, one byte per character, with no length.
    /// </summary>
    public void WriteMacRoman(string value) => MacRoman.GetBytes(value, Reserve(value.Length));

    /// <summary>
    /// Overwrites a big-endian word at an earlier position.
    /// </summary>
    public void PatchUInt16(int position, ushort value)
        => BinaryPrimitives.WriteUInt16BigEndian(_buffer.AsSpan(position, 2), value);

    public void PatchBytes(int position, ReadOnlySpan<byte> bytes) => bytes.CopyTo(_buffer.AsSpan(position, bytes.Length));

    public void PatchUInt32(int position, uint value)
        => BinaryPrimitives.WriteUInt32BigEndian(_buffer.AsSpan(position, 4), value);

    /// <summary>
    /// Gets the bytes written. When they fill the buffer exactly, the buffer itself is returned and the writer must
    /// not be used again.
    /// </summary>
    public byte[] ToArray() => !_pooled && Position == _buffer.Length ? _buffer : _buffer.AsSpan(0, Position).ToArray();
}
