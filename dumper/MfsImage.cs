using System.Buffers.Binary;
using System.Text;

/// <summary>
/// A file on a Macintosh File System (MFS) floppy: its name, Finder type and creator, and both forks.
/// </summary>
sealed record MfsFile(string Name, string Type, string Creator, byte[] Data, byte[] Resource);

/// <summary>
/// Reads and writes 400K and 800K MFS floppy images, which every Mac emulator can mount and Excel 1.x can open
/// files from. MFS is flat: one directory, no folders.
/// </summary>
static class MfsImage
{
    private const ushort Signature = 0xD2D7;
    private const int DirectoryStart = 4;
    private const int DirectoryLength = 12;
    private static readonly Encoding s_macRoman = CodePagesEncoding();

    public static byte[] Write(string volume, IReadOnlyList<MfsFile> files, int sizeK)
    {
        int sectors = sizeK * 2;
        int blockSize = sizeK <= 400 ? 1024 : 2048;   // keeps the block map inside the 1K after the MDB
        int first = DirectoryStart + DirectoryLength;
        int blockCount = (sectors - first) * 512 / blockSize;
        var image = new byte[sectors * 512];
        var map = new int[blockCount];
        int next = 0;
        uint now = (uint)(DateTime.UtcNow - new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;

        (int Start, int Physical) Allocate(byte[] data)
        {
            if (data.Length == 0)
            {
                return (0, 0);
            }

            int count = (data.Length + blockSize - 1) / blockSize;
            if (next + count > blockCount)
            {
                throw new InvalidOperationException($"The files do not fit on a {sizeK}K disk.");
            }

            int start = next;
            for (int i = 0; i < count; i++)
            {
                map[start + i] = i < count - 1 ? start + i + 3 : 1;
                int length = Math.Min(blockSize, data.Length - i * blockSize);
                data.AsSpan(i * blockSize, length).CopyTo(image.AsSpan(first * 512 + (start + i) * blockSize));
            }

            next += count;
            return (start + 2, count * blockSize);
        }

        var directory = new byte[DirectoryLength * 512];
        int position = 0;
        for (int i = 0; i < files.Count; i++)
        {
            var file = files[i];
            var (dataStart, dataPhysical) = Allocate(file.Data);
            var (rsrcStart, rsrcPhysical) = Allocate(file.Resource);
            byte[] name = s_macRoman.GetBytes(file.Name);
            if (name.Length is 0 or > 31)
            {
                throw new ArgumentException($"File names must be 1–31 characters: {file.Name}");
            }

            var entry = new byte[50 + 1 + name.Length + ((51 + name.Length) & 1)];
            entry[0] = 0x80;
            s_macRoman.GetBytes(file.Type).CopyTo(entry, 2);
            s_macRoman.GetBytes(file.Creator).CopyTo(entry, 6);
            // Icon position in the disk's Finder window (vertical, horizontal): four to a row.
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(12), (uint)(((40 + 60 * (i / 4)) << 16) | (30 + 90 * (i % 4))));
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(18), (uint)(i + 1));
            BinaryPrimitives.WriteUInt16BigEndian(entry.AsSpan(22), (ushort)dataStart);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(24), (uint)file.Data.Length);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(28), (uint)dataPhysical);
            BinaryPrimitives.WriteUInt16BigEndian(entry.AsSpan(32), (ushort)rsrcStart);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(34), (uint)file.Resource.Length);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(38), (uint)rsrcPhysical);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(42), now);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(46), now);
            entry[50] = (byte)name.Length;
            name.CopyTo(entry, 51);

            // Entries never cross a sector boundary.
            if (position / 512 != (position + entry.Length - 1) / 512)
            {
                position = (position / 512 + 1) * 512;
            }

            if (position + entry.Length > directory.Length)
            {
                throw new InvalidOperationException("Too many files for the directory.");
            }

            entry.CopyTo(directory, position);
            position += entry.Length;
        }

        directory.CopyTo(image, DirectoryStart * 512);

        var mdb = image.AsSpan(1024);
        BinaryPrimitives.WriteUInt16BigEndian(mdb, Signature);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(2), now);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(6), 0);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(10), 0);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(12), (ushort)files.Count);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(14), DirectoryStart);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(16), DirectoryLength);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(18), (ushort)blockCount);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(20), (uint)blockSize);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(24), (uint)(8 * blockSize));
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(28), (ushort)first);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(30), (uint)(files.Count + 1));
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(34), (ushort)(blockCount - next));
        byte[] volumeName = s_macRoman.GetBytes(volume);
        if (volumeName.Length is 0 or > 27)
        {
            throw new ArgumentException($"Volume names must be 1–27 characters: {volume}");
        }

        mdb[36] = (byte)volumeName.Length;
        volumeName.CopyTo(mdb.Slice(37));

        // The block map: 12 bits per allocation block, after the 64-byte MDB.
        var mapBytes = mdb.Slice(64);
        for (int i = 0; i < blockCount; i += 2)
        {
            int a = map[i];
            int b = i + 1 < blockCount ? map[i + 1] : 0;
            int o = i / 2 * 3;
            mapBytes[o] = (byte)(a >> 4);
            mapBytes[o + 1] = (byte)(((a & 15) << 4) | (b >> 8));
            if (o + 2 < (blockCount * 3 + 1) / 2)
            {
                mapBytes[o + 2] = (byte)b;
            }
        }

        return image;
    }

    /// <summary>
    /// Adds files to a copy of a disk, first growing it to <paramref name="sizeK"/>, without moving anything already on
    /// it: copy protection can look for a file's data at a fixed sector. The disk keeps its allocation block size, so
    /// it grows only as far as its block map, which ends where the directory starts, can reach.
    /// </summary>
    public static byte[] AddFiles(byte[] original, IReadOnlyList<MfsFile> files, int sizeK)
    {
        var image = new byte[Math.Max(sizeK * 1024, original.Length)];
        original.CopyTo(image, 0);
        var mdb = image.AsSpan(1024);
        if (BinaryPrimitives.ReadUInt16BigEndian(mdb) != Signature)
        {
            throw new InvalidDataException("Not an MFS disk image.");
        }

        int fileCount = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(12));
        int directoryStart = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(14));
        int directoryLength = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(16));
        int blockCount = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(18));
        int blockSize = (int)BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(20));
        int allocationStart = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(28));
        uint fileNumber = BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(30));
        var mapBytes = mdb.Slice(64, (directoryStart - 2) * 512 - 64);

        var map = new int[Math.Max(blockCount, Math.Min((image.Length / 512 - allocationStart) * 512 / blockSize, mapBytes.Length * 2 / 3))];
        for (int i = 0; i < blockCount; i++)
        {
            int v = (mapBytes[i * 3 / 2] << 8) | mapBytes[i * 3 / 2 + 1];
            map[i] = i % 2 == 0 ? v >> 4 : v & 0xFFF;
        }

        (int Start, int Physical) Allocate(byte[] data)
        {
            if (data.Length == 0)
            {
                return (0, 0);
            }

            var blocks = Enumerable.Range(0, map.Length).Where(i => map[i] == 0).Take((data.Length + blockSize - 1) / blockSize).ToArray();
            if (blocks.Length * blockSize < data.Length)
            {
                throw new InvalidOperationException($"The files do not fit on a {sizeK}K disk.");
            }

            for (int i = 0; i < blocks.Length; i++)
            {
                map[blocks[i]] = i < blocks.Length - 1 ? blocks[i + 1] + 2 : 1;
                int length = Math.Min(blockSize, data.Length - i * blockSize);
                data.AsSpan(i * blockSize, length).CopyTo(image.AsSpan(allocationStart * 512 + blocks[i] * blockSize));
            }

            return (blocks[0] + 2, blocks.Length * blockSize);
        }

        // The end of the directory's entries.
        var directory = image.AsSpan(directoryStart * 512, directoryLength * 512);
        int position = 0;
        for (int found = 0; found < fileCount;)
        {
            if ((directory[position] & 0x80) == 0)
            {
                position = (position / 512 + 1) * 512;
                continue;
            }

            position += 51 + directory[position + 50];
            position += position & 1;
            found++;
        }

        uint now = (uint)(DateTime.UtcNow - new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
        foreach (var file in files)
        {
            var (dataStart, dataPhysical) = Allocate(file.Data);
            var (rsrcStart, rsrcPhysical) = Allocate(file.Resource);
            byte[] name = s_macRoman.GetBytes(file.Name);
            var entry = new byte[50 + 1 + name.Length + ((51 + name.Length) & 1)];
            entry[0] = 0x80;
            s_macRoman.GetBytes(file.Type).CopyTo(entry, 2);
            s_macRoman.GetBytes(file.Creator).CopyTo(entry, 6);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(18), fileNumber++);
            BinaryPrimitives.WriteUInt16BigEndian(entry.AsSpan(22), (ushort)dataStart);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(24), (uint)file.Data.Length);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(28), (uint)dataPhysical);
            BinaryPrimitives.WriteUInt16BigEndian(entry.AsSpan(32), (ushort)rsrcStart);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(34), (uint)file.Resource.Length);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(38), (uint)rsrcPhysical);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(42), now);
            BinaryPrimitives.WriteUInt32BigEndian(entry.AsSpan(46), now);
            entry[50] = (byte)name.Length;
            name.CopyTo(entry, 51);
            if (position / 512 != (position + entry.Length - 1) / 512)
            {
                position = (position / 512 + 1) * 512;
            }

            if (position + entry.Length > directory.Length)
            {
                throw new InvalidOperationException("Too many files for the directory.");
            }

            entry.CopyTo(directory.Slice(position));
            position += entry.Length;
            fileCount++;
        }

        for (int i = 0; i < map.Length; i++)
        {
            int o = i * 3 / 2;
            if (i % 2 == 0)
            {
                mapBytes[o] = (byte)(map[i] >> 4);
                mapBytes[o + 1] = (byte)((mapBytes[o + 1] & 0x0F) | ((map[i] & 15) << 4));
            }
            else
            {
                mapBytes[o] = (byte)((mapBytes[o] & 0xF0) | (map[i] >> 8));
                mapBytes[o + 1] = (byte)map[i];
            }
        }

        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(12), (ushort)fileCount);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(18), (ushort)map.Length);
        BinaryPrimitives.WriteUInt32BigEndian(mdb.Slice(30), fileNumber);
        BinaryPrimitives.WriteUInt16BigEndian(mdb.Slice(34), (ushort)map.Count(m => m == 0));
        return image;
    }

    public static (string Volume, List<MfsFile> Files) Read(byte[] image)
    {
        var mdb = image.AsSpan(1024);
        if (BinaryPrimitives.ReadUInt16BigEndian(mdb) != Signature)
        {
            throw new InvalidDataException("Not an MFS disk image.");
        }

        int directoryStart = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(14));
        int directoryLength = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(16));
        int blockCount = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(18));
        int blockSize = (int)BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(20));
        int allocationStart = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(28));
        string volume = s_macRoman.GetString(mdb.Slice(37, mdb[36]));
        byte[] map = mdb.Slice(64, (blockCount * 3 + 1) / 2 + 1).ToArray();

        int Next(int block)
        {
            int i = block - 2;
            int o = i * 3 / 2;
            int v = (map[o] << 8) | map[o + 1];
            return i % 2 == 0 ? v >> 4 : v & 0xFFF;
        }

        byte[] Fork(int start, int length)
        {
            var data = new byte[length];
            int written = 0;
            int block = start;
            for (int guard = 0; block is not (0 or 1) && written < length && guard <= blockCount; guard++)
            {
                int count = Math.Min(blockSize, length - written);
                image.AsSpan(allocationStart * 512 + (block - 2) * blockSize, count).CopyTo(data.AsSpan(written));
                written += count;
                block = Next(block);
            }

            return data;
        }

        var files = new List<MfsFile>();
        var directory = image.AsSpan(directoryStart * 512, directoryLength * 512);
        for (int sector = 0; sector < directoryLength; sector++)
        {
            int p = sector * 512;
            while (p < (sector + 1) * 512 - 51 && (directory[p] & 0x80) != 0)
            {
                string type = s_macRoman.GetString(directory.Slice(p + 2, 4));
                string creator = s_macRoman.GetString(directory.Slice(p + 6, 4));
                int dataStart = BinaryPrimitives.ReadUInt16BigEndian(directory.Slice(p + 22));
                int dataLength = (int)BinaryPrimitives.ReadUInt32BigEndian(directory.Slice(p + 24));
                int rsrcStart = BinaryPrimitives.ReadUInt16BigEndian(directory.Slice(p + 32));
                int rsrcLength = (int)BinaryPrimitives.ReadUInt32BigEndian(directory.Slice(p + 34));
                int nameLength = directory[p + 50];
                string name = s_macRoman.GetString(directory.Slice(p + 51, nameLength));
                files.Add(new MfsFile(name, type, creator, Fork(dataStart, dataLength), Fork(rsrcStart, rsrcLength)));
                p += 51 + nameLength;
                p += p & 1;
            }
        }

        return (volume, files);
    }

    private static Encoding CodePagesEncoding()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(10000);
    }
}
