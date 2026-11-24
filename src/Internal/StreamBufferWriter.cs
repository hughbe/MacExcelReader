using System.Buffers;

namespace MacExcelReader.Internal;

/// <summary>
/// An <see cref="IBufferWriter{T}"/> that writes to a stream through one pooled buffer, flushing it as it fills, so
/// writing a whole file costs one rented array however long the file is.
/// </summary>
internal sealed class StreamBufferWriter : IBufferWriter<byte>, IDisposable
{
    private readonly Stream _stream;
    private byte[] _buffer;
    private int _count;

    public StreamBufferWriter(Stream stream, int size = 16384)
    {
        _stream = stream;
        _buffer = ArrayPool<byte>.Shared.Rent(size);
    }

    public void Advance(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(count, _buffer.Length - _count);
        _count += count;
    }

    public Memory<byte> GetMemory(int sizeHint = 0)
    {
        int start = Reserve(sizeHint);   // may replace the buffer, so read it after
        return _buffer.AsMemory(start);
    }

    public Span<byte> GetSpan(int sizeHint = 0)
    {
        int start = Reserve(sizeHint);
        return _buffer.AsSpan(start);
    }

    // Makes room for at least sizeHint bytes (one if 0) and returns where they start.
    private int Reserve(int sizeHint)
    {
        int needed = Math.Max(1, sizeHint);
        if (_buffer.Length - _count < needed)
        {
            Flush();
            if (_buffer.Length < needed)
            {
                ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = ArrayPool<byte>.Shared.Rent(needed);
            }
        }

        return _count;
    }

    public void Flush()
    {
        if (_count > 0)
        {
            _stream.Write(_buffer, 0, _count);
            _count = 0;
        }
    }

    public void Dispose()
    {
        var buffer = _buffer;
        _buffer = [];
        _count = 0;
        if (buffer.Length > 0)
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
}
