using MacExcelReader.Internal;

namespace MacExcelReader.Rendering;

/// <summary>
/// The bitmap fonts of a Macintosh System file, which Excel draws with. The library carries no fonts of its own:
/// load them from the resource fork of the System file Excel ran under, such as the one on Excel 1.00's Data disk,
/// which has Chicago 12, New York 10, 12, 20 and 24, Geneva 9, 10, 12, 14, 18, 20 and 24, and Monaco 9.
/// </summary>
public sealed class MacFontSet
{
    private readonly Dictionary<(int Family, int Size), MacBitmapFont> _fonts = [];

    // Sizes made by scaling, kept so that each is made once.
    private readonly Dictionary<(int Family, int Size), MacBitmapFont?> _scaled = [];
    private readonly Lock _lock = new();

    private MacFontSet()
    {
    }

    /// <summary>Gets the fonts read from the resource fork.</summary>
    public IEnumerable<MacBitmapFont> Fonts => _fonts.Values;

    /// <summary>
    /// Reads every <c>FONT</c> resource of a resource fork: resource ID ÷ 128 is the family and ID mod 128 the size,
    /// size 0 being the family's name.
    /// </summary>
    /// <exception cref="ExcelFormatException">The resource fork, or a font in it, is not valid.</exception>
    public static MacFontSet FromResourceFork(ReadOnlySpan<byte> resourceFork)
    {
        var set = new MacFontSet();
        ResourceFork.ForEach(resourceFork, ResourceFork.Type("FONT"u8), (id, data) =>
        {
            int family = id / 128, size = id % 128;
            if (id >= 0 && size > 0)
            {
                set._fonts[(family, size)] = MacBitmapFont.Read(family, size, data);
            }
        });

        if (set._fonts.Count == 0)
        {
            throw new ExcelFormatException("The resource fork has no FONT resources.");
        }

        return set;
    }

    /// <summary>
    /// Reads the fonts of a System file's resource fork saved as a file of its own (on macOS a file's resource fork
    /// can be read as <c>FILE/..namedfork/rsrc</c>).
    /// </summary>
    public static MacFontSet Load(string resourceForkPath) => FromResourceFork(File.ReadAllBytes(resourceForkPath));

    /// <summary>
    /// Gets a font as QuickDraw draws it: the bitmap of that size, or for a size with no bitmap one scaled from the
    /// next larger size of the family, or from the largest if none is larger (docs/Display.md). Returns null for a
    /// family the System lacks.
    /// </summary>
    public MacBitmapFont? GetFont(int family, int size)
    {
        if (size <= 0)
        {
            return null;
        }

        if (_fonts.TryGetValue((family, size), out var font))
        {
            return font;
        }

        lock (_lock)
        {
            if (!_scaled.TryGetValue((family, size), out var scaled))
            {
                MacBitmapFont? larger = null, largest = null;
                foreach (var candidate in _fonts.Values)
                {
                    if (candidate.Family != family)
                    {
                        continue;
                    }

                    if (candidate.Size > size && (larger is null || candidate.Size < larger.Size))
                    {
                        larger = candidate;
                    }

                    if (largest is null || candidate.Size > largest.Size)
                    {
                        largest = candidate;
                    }
                }

                scaled = (larger ?? largest)?.Scale(size);
                _scaled[(family, size)] = scaled;
            }

            return scaled;
        }
    }
}
