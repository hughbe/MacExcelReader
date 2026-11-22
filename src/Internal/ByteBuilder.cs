using System.Buffers;

namespace MacExcelReader.Internal;

/// <summary>
/// A growable byte buffer on pooled arrays, for building one expression or record and copying it out once. A mutable
/// struct: keep it in a field or local and never copy it.
/// </summary>
internal struct ByteBuilder : IDisposable
{
    private byte[] _buffer;

    public ByteBuilder(int capacity)
    {
        _buffer = ArrayPool<byte>.Shared.Rent(capacity);
    }

    public int Count { get; private set; }

    public byte this[int index]
    {
        readonly get => _buffer.AsSpan(0, Count)[index];
        set => _buffer.AsSpan(0, Count)[index] = value;
    }

    public void Add(byte value)
    {
        if (Count == _buffer.Length)
        {
            Grow(1);
        }

        _buffer[Count++] = value;
    }

    public void AddRange(ReadOnlySpan<byte> values)
    {
        if (Count + values.Length > _buffer.Length)
        {
            Grow(values.Length);
        }

        values.CopyTo(_buffer.AsSpan(Count));
        Count += values.Length;
    }

    public readonly ReadOnlySpan<byte> AsSpan() => _buffer.AsSpan(0, Count);

    public readonly byte[] ToArray() => AsSpan().ToArray();

    public void Dispose()
    {
        var buffer = _buffer;
        _buffer = [];
        Count = 0;
        if (buffer.Length > 0)
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    private void Grow(int needed)
    {
        var larger = ArrayPool<byte>.Shared.Rent(Math.Max(_buffer.Length * 2, Count + needed));
        _buffer.AsSpan(0, Count).CopyTo(larger);
        ArrayPool<byte>.Shared.Return(_buffer);
        _buffer = larger;
    }
}
