using System.Globalization;

namespace MacExcelReader.Internal;

// A text buffer that starts on the stack and grows into pooled arrays.
internal ref struct TextBuffer
{
    private Span<char> _chars;
    private char[]? _rented;

    public TextBuffer(Span<char> initial)
    {
        _chars = initial;
        _rented = null;
        Length = 0;
    }

    public int Length { get; private set; }

    public void Append(char c)
    {
        if (Length == _chars.Length)
        {
            Grow(1);
        }

        _chars[Length++] = c;
    }

    public void Append(scoped ReadOnlySpan<char> value)
    {
        if (Length + value.Length > _chars.Length)
        {
            Grow(value.Length);
        }

        value.CopyTo(_chars[Length..]);
        Length += value.Length;
    }

    public void AppendFormatted<T>(T value, scoped ReadOnlySpan<char> format)
        where T : ISpanFormattable
    {
        int written;
        while (!value.TryFormat(_chars[Length..], out written, format, CultureInfo.InvariantCulture))
        {
            Grow(32);
        }

        Length += written;
    }

    // Mac OS Roman text, optionally doubling quotes as a formula's string constant does.
    public void AppendMacRoman(scoped ReadOnlySpan<byte> bytes, bool escapeQuotes)
    {
        foreach (byte b in bytes)
        {
            char c = MacRoman.GetChar(b);
            if (escapeQuotes && c == '"')
            {
                Append('"');
            }

            Append(c);
        }
    }

    public void Insert(int index, scoped ReadOnlySpan<char> value)
    {
        if (Length + value.Length > _chars.Length)
        {
            Grow(value.Length);
        }

        _chars[index..Length].CopyTo(_chars[(index + value.Length)..]);
        value.CopyTo(_chars[index..]);
        Length += value.Length;
    }

    public void Remove(int index, int count)
    {
        _chars[(index + count)..Length].CopyTo(_chars[index..]);
        Length -= count;
    }

    public override readonly string ToString() => new(_chars[..Length]);

    public readonly ReadOnlySpan<char> AsSpan() => _chars[..Length];

    public void Dispose()
    {
        if (_rented is not null)
        {
            System.Buffers.ArrayPool<char>.Shared.Return(_rented);
            _rented = null;
        }
    }

    private void Grow(int needed)
    {
        var larger = System.Buffers.ArrayPool<char>.Shared.Rent(Math.Max(_chars.Length * 2, Length + needed));
        _chars[..Length].CopyTo(larger);
        if (_rented is not null)
        {
            System.Buffers.ArrayPool<char>.Shared.Return(_rented);
        }

        _chars = _rented = larger;
    }
}
