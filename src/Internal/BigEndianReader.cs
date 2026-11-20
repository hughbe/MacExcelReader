using System.Buffers.Binary;

namespace MacExcelReader.Internal;

/// <summary>
/// Sequential big-endian reader over a byte buffer.
/// </summary>
internal sealed class BigEndianReader
{
    private readonly byte[] _data;

    public BigEndianReader(byte[] data, int position = 0)
    {
        _data = data;
        Position = position;
    }

    public int Position { get; set; }

    public int Length => _data.Length;

    public int Remaining => _data.Length - Position;

    public byte[] Data => _data;

    private void Require(int count)
    {
        if (count < 0 || Position < 0 || Position + count > _data.Length)
        {
            throw new ExcelFormatException(
                $"Unexpected end of data: needed {count} byte(s) at offset {Position}, file is {_data.Length} bytes.");
        }
    }

    public byte ReadByte()
    {
        Require(1);
        return _data[Position++];
    }

    public ushort ReadUInt16()
    {
        Require(2);
        var value = BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(Position));
        Position += 2;
        return value;
    }

    public short ReadInt16() => unchecked((short)ReadUInt16());

    public uint ReadUInt32()
    {
        Require(4);
        var value = BinaryPrimitives.ReadUInt32BigEndian(_data.AsSpan(Position));
        Position += 4;
        return value;
    }

    public double ReadDouble()
    {
        Require(8);
        var value = BinaryPrimitives.ReadDoubleBigEndian(_data.AsSpan(Position));
        Position += 8;
        return value;
    }

    public byte[] ReadBytes(int count)
    {
        Require(count);
        var result = _data.AsSpan(Position, count).ToArray();
        Position += count;
        return result;
    }

    /// <summary>
    /// Reads a Pascal string (a length byte followed by that many bytes) and returns the raw bytes.
    /// </summary>
    public byte[] ReadPascalBytes()
    {
        var length = ReadByte();
        return ReadBytes(length);
    }

    public string ReadPascalString() => MacRoman.GetString(ReadPascalBytes());

    public void Skip(int count)
    {
        Require(count);
        Position += count;
    }
}
