using System.Buffers.Binary;
using System.Text;

/// <summary>
/// Reads the files off an HFS floppy image, as Excel 1.01 to 1.06 shipped on: each file's path from the volume's root
/// (folders joined with ':'), Finder type and creator, and both forks. Enough for `excel-test run` to lift Excel and
/// a System Folder off a release's own disk.
/// </summary>
static class HfsImage
{
    private const ushort Signature = 0x4244;

    public static bool IsHfs(byte[] image) =>
        image.Length >= 1026 && BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(1024)) == Signature;

    public static List<(string Path, MfsFile File)> Read(byte[] image)
    {
        var mdb = image.AsSpan(1024, 512);
        if (BinaryPrimitives.ReadUInt16BigEndian(mdb) != Signature)
        {
            throw new InvalidDataException("Not an HFS disk image.");
        }

        int blockSize = (int)BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(0x14));
        int firstBlock = BinaryPrimitives.ReadUInt16BigEndian(mdb.Slice(0x1C)) * 512;
        int catalogLength = (int)BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(0x92));
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = Encoding.GetEncoding(10000);

        // Copies a fork from its extents: the first three, in the catalog or the MDB, then any more from the extents
        // overflow file, whose leaf records are keyed by fork (0 data, 0xFF resource), file ID and first block.
        var overflow = new Dictionary<(int Fork, uint File), List<(int Start, byte[] Extents)>>();
        byte[] Fork(byte[] extents, int length, int fork, uint file)
        {
            var data = new byte[length];
            int done = 0, blocks = 0;
            var more = overflow.TryGetValue((fork, file), out var records) ? records.OrderBy(r => r.Start).ToList() : [];
            for (int record = -1; done < length && record < more.Count; record++)
            {
                byte[] three = record < 0 ? extents : more[record].Extents;
                if (record >= 0 && more[record].Start != blocks)
                {
                    break;
                }

                for (int i = 0; i < 3 && done < length; i++)
                {
                    int start = BinaryPrimitives.ReadUInt16BigEndian(three.AsSpan(4 * i));
                    int count = BinaryPrimitives.ReadUInt16BigEndian(three.AsSpan(4 * i + 2));
                    int n = Math.Min(count * blockSize, length - done);
                    image.AsSpan(firstBlock + start * blockSize, n).CopyTo(data.AsSpan(done));
                    done += n;
                    blocks += count;
                }
            }

            if (done < length)
            {
                throw new InvalidDataException($"File {file}'s fork is not all in its extents.");
            }

            return data;
        }

        // The leaf records of a B-tree, from the first leaf along the forward links.
        static IEnumerable<(int Offset, int Node)> Leaves(byte[] tree)
        {
            int nodeSize = BinaryPrimitives.ReadUInt16BigEndian(tree.AsSpan(14 + 18));
            int node = (int)BinaryPrimitives.ReadUInt32BigEndian(tree.AsSpan(14 + 10));
            for (int guard = 0; node != 0 && guard < tree.Length / nodeSize; guard++)
            {
                int start = node * nodeSize;
                int records = BinaryPrimitives.ReadUInt16BigEndian(tree.AsSpan(start + 10));
                for (int r = 0; r < records; r++)
                {
                    yield return (start + BinaryPrimitives.ReadUInt16BigEndian(tree.AsSpan(start + nodeSize - 2 * (r + 1))), start);
                }

                node = (int)BinaryPrimitives.ReadUInt32BigEndian(tree.AsSpan(start));
            }
        }

        byte[] extentsTree = Fork(mdb.Slice(0x86, 12).ToArray(), (int)BinaryPrimitives.ReadUInt32BigEndian(mdb.Slice(0x82)), 0, 3);
        foreach (var (offset, _) in Leaves(extentsTree))
        {
            if (extentsTree[offset] < 7)
            {
                continue;
            }

            var key = (extentsTree[offset + 1], BinaryPrimitives.ReadUInt32BigEndian(extentsTree.AsSpan(offset + 2)));
            if (!overflow.TryGetValue(key, out var list))
            {
                overflow[key] = list = [];
            }

            list.Add((BinaryPrimitives.ReadUInt16BigEndian(extentsTree.AsSpan(offset + 6)), extentsTree.AsSpan(offset + 8, 12).ToArray()));
        }

        byte[] catalog = Fork(mdb.Slice(0x96, 12).ToArray(), catalogLength, 0, 4);

        // Folders by ID (the root's parent is 1), then the files in each, from the leaf records in key order.
        var folders = new Dictionary<uint, (uint Parent, string Name)>();
        var files = new List<(uint Parent, MfsFile File)>();
        foreach (var (offset, _) in Leaves(catalog))
        {
            int keyLength = catalog[offset];
            if (keyLength == 0)
            {
                continue;   // a deleted record
            }

            uint parent = BinaryPrimitives.ReadUInt32BigEndian(catalog.AsSpan(offset + 2));
            string name = encoding.GetString(catalog, offset + 7, catalog[offset + 6]);
            var data = catalog.AsSpan(offset + 1 + keyLength + ((1 + keyLength) & 1));
            switch (data[0])
            {
                case 1:
                    folders[BinaryPrimitives.ReadUInt32BigEndian(data.Slice(6))] = (parent, name);
                    break;
                case 2:
                    uint id = BinaryPrimitives.ReadUInt32BigEndian(data.Slice(20));
                    files.Add((parent, new MfsFile(
                        name,
                        encoding.GetString(data.Slice(4, 4)),
                        encoding.GetString(data.Slice(8, 4)),
                        Fork(data.Slice(74, 12).ToArray(), (int)BinaryPrimitives.ReadUInt32BigEndian(data.Slice(26)), 0, id),
                        Fork(data.Slice(86, 12).ToArray(), (int)BinaryPrimitives.ReadUInt32BigEndian(data.Slice(36)), 0xFF, id))));
                    break;
            }
        }
        string PathOf(uint folder)
        {
            var parts = new List<string>();
            for (int guard = 0; guard < 64 && folders.TryGetValue(folder, out var f) && f.Parent != 1; guard++)
            {
                parts.Insert(0, f.Name);
                folder = f.Parent;
            }

            return string.Join(":", parts);
        }

        return [.. files.Select(f => (PathOf(f.Parent) is { Length: > 0 } p ? p + ":" + f.File.Name : f.File.Name, f.File))];
    }
}
