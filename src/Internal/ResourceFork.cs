using System.Buffers.Binary;

namespace MacExcelReader.Internal;

/// <summary>
/// Reads the resources of a Macintosh resource fork (Inside Macintosh, "The Resource Manager"): a 16-byte header
/// giving the offsets of the data and the map, a map listing each type with its resources, and the data, each
/// resource a 4-byte length and its bytes.
/// </summary>
internal static class ResourceFork
{
    /// <summary>
    /// Calls <paramref name="found"/> with the ID and data of every resource of the given type.
    /// </summary>
    public static void ForEach(ReadOnlySpan<byte> fork, uint type, ResourceAction found)
    {
        if (fork.Length < 16)
        {
            throw new ExcelFormatException("A resource fork is at least 16 bytes long.");
        }

        int dataOffset = ReadInt(fork, 0);
        int mapOffset = ReadInt(fork, 4);
        int typeList = mapOffset + ReadWord(fork, mapOffset + 24);
        int types = ReadWord(fork, typeList) + 1 & 0xFFFF;
        for (int t = 0; t < types; t++)
        {
            int entry = typeList + 2 + 8 * t;
            if (ReadUInt(fork, entry) != type)
            {
                continue;
            }

            int count = ReadWord(fork, entry + 4) + 1;
            int references = typeList + ReadWord(fork, entry + 6);
            for (int i = 0; i < count; i++)
            {
                int reference = references + 12 * i;
                short id = (short)ReadWord(fork, reference);
                int offset = dataOffset + (int)(ReadUInt(fork, reference + 4) & 0xFFFFFF);
                int length = ReadInt(fork, offset);
                if (length < 0 || offset + 4 + length > fork.Length)
                {
                    throw new ExcelFormatException($"Resource {id} runs past the end of the resource fork.");
                }

                found(id, fork.Slice(offset + 4, length));
            }
        }
    }

    public delegate void ResourceAction(short id, ReadOnlySpan<byte> data);

    public static uint Type(ReadOnlySpan<byte> fourCharacters) => BinaryPrimitives.ReadUInt32BigEndian(fourCharacters);

    private static int ReadWord(ReadOnlySpan<byte> fork, int offset)
        => (uint)offset + 2 <= (uint)fork.Length
            ? BinaryPrimitives.ReadUInt16BigEndian(fork[offset..])
            : throw new ExcelFormatException("The resource map runs past the end of the resource fork.");

    private static uint ReadUInt(ReadOnlySpan<byte> fork, int offset)
        => (uint)offset + 4 <= (uint)fork.Length
            ? BinaryPrimitives.ReadUInt32BigEndian(fork[offset..])
            : throw new ExcelFormatException("The resource map runs past the end of the resource fork.");

    private static int ReadInt(ReadOnlySpan<byte> fork, int offset) => (int)ReadUInt(fork, offset);
}
