using System.Runtime.InteropServices;
using System.Text;

/// <summary>
/// Reads Mac files kept in a macOS folder, as copied off a disk with their resource forks and Finder info: the
/// resource fork is the file's `..namedfork/rsrc` and the type and creator are the first 8 bytes of its
/// com.apple.FinderInfo attribute. Only on macOS, whose file systems keep both.
/// </summary>
static class MacFolder
{
    public static List<(string Path, MfsFile File)> Read(string folder)
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException($"{folder} is a folder: only macOS keeps Mac files' resource forks and Finder info in a folder. Use a disk image.");
        }

        var encoding = Encoding.GetEncoding(10000);
        var files = new List<(string, MfsFile)>();
        foreach (string path in Directory.GetFiles(folder, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
        {
            string relative = Path.GetRelativePath(folder, path);
            if (relative.Split(Path.DirectorySeparatorChar).Any(p => p.StartsWith('.')))
            {
                continue;
            }

            var info = new byte[32];
            bool hasInfo = GetXattr(path, "com.apple.FinderInfo", info, info.Length, 0, 0) >= 8;
            string resourcePath = Path.Combine(path, "..namedfork", "rsrc");
            byte[] resource = File.Exists(resourcePath) ? File.ReadAllBytes(resourcePath) : [];
            files.Add((relative.Replace(Path.DirectorySeparatorChar, ':'), new MfsFile(
                Path.GetFileName(path),
                hasInfo ? encoding.GetString(info, 0, 4) : "????",
                hasInfo ? encoding.GetString(info, 4, 4) : "????",
                File.ReadAllBytes(path),
                resource)));
        }

        return files;
    }

    [DllImport("libc", EntryPoint = "getxattr", SetLastError = true)]
    private static extern nint GetXattr(string path, string name, byte[] value, nint size, uint position, int options);
}
