using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using MacExcelReader.Rendering;
using Spectre.Console;
using Spectre.Console.Cli;

// Runs a disk written by `excel-test pack` in a real Excel, 1.00 to 1.5, headless: MAME's Mac Plus boots a disk built
// from the release's Program disk and a System, with Excel as its startup application, and excel-test.lua types the
// keys that open Driver and run it, captures windows from the screen buffer and ejects the test disk. See
// docs/Testing.md.

sealed class RunSettings : CommandSettings
{
    [CommandArgument(0, "<disk>")]
    [Description("The floppy image written by excel-test pack. Excel's saves are written back to it.")]
    public required string Disk { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("The Excel to run: 1.0 (the default), 1.01, 1.03, 1.04, 1.06 or 1.5. Pack the disk with the same --excel.")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--system <IMAGE>")]
    [Description("A startup disk whose System and Finder the Mac boots, such as Excel 1.00's Data disk. Not needed when the Program disk has a System Folder (1.03, 1.04). " +
        "Defaults to $MACEXCEL_SYSTEM for 1.0 and $MACEXCEL_SYSTEM_1_03 and so on for the others.")]
    public string? System { get; init; }

    [CommandOption("--program <IMAGE>")]
    [Description("The disk with the Excel application (MFS or HFS), or a folder holding it with its resource fork (macOS). " +
        "Defaults to $MACEXCEL_PROGRAM for 1.0 and $MACEXCEL_PROGRAM_1_03 and so on for the others.")]
    public string? Program { get; init; }

    [CommandOption("--rom <FILE>")]
    [Description("The Mac Plus ROM (128K; checksum 4D1F8172). Defaults to $MACEXCEL_ROM.")]
    public string? Rom { get; init; }

    [CommandOption("--capture <TITLE>")]
    [Description("After Driver has run, bring the window with this title to the front and save it as a PNG, like tests/Data/Screens. " +
        "TITLE=FILE names the PNG; otherwise it is the title with ':' as ' - '. Repeatable. Pack with --keep-open.")]
    public string[] Capture { get; init; } = [];

    [CommandOption("-o|--output <DIR>")]
    [Description("Where to write captured windows (default: the current folder).")]
    public string Output { get; init; } = ".";

    [CommandOption("--screen <FILE>")]
    [Description("Also save the whole screen as a PNG at the end.")]
    public string? Screen { get; init; }

    [CommandOption("--timeout <SECONDS>")]
    [Description("Give up after this many seconds of emulated time (default 3600). MAME runs the Mac Plus about 30 times faster than real time.")]
    public int Timeout { get; init; } = 3600;

    [CommandOption("--mame <PATH>")]
    [Description("The MAME executable (default: mame on the PATH).")]
    public string Mame { get; init; } = "mame";

    [CommandOption("--show")]
    [Description("Show MAME's window, at real speed, to watch the run.")]
    public bool Show { get; init; }

    [CommandOption("--work <DIR>")]
    [Description("Keep MAME's working files (boot disk, ROMs, script, log) in this folder instead of a temporary one.")]
    public string? Work { get; init; }
}

sealed class RunCommand : Command<RunSettings>
{
    // Excel's name on the boot disk: the boot blocks name the startup application in 15 characters.
    private const int MaxStartupName = 15;

    public static readonly string[] Releases = ["1.0", "1.01", "1.03", "1.04", "1.06", "1.5"];

    // The environment variables' suffix for a release: none for 1.0, "_1_03" for 1.03 and so on.
    public static string Suffix(string release) => release == "1.0" ? "" : "_" + release.Replace('.', '_');

    private static string? Variable(string name) => Environment.GetEnvironmentVariable(name) is { Length: > 0 } value ? value : null;

    public override int Execute(CommandContext context, RunSettings settings, CancellationToken cancellationToken)
    {
        if (!Releases.Contains(settings.Excel))
        {
            AnsiConsole.MarkupLine($"[red]--excel must be one of {string.Join(", ", Releases)}.[/]");
            return 2;
        }

        string suffix = Suffix(settings.Excel);
        string? system = settings.System ?? Variable("MACEXCEL_SYSTEM" + suffix);
        string? program = settings.Program ?? Variable("MACEXCEL_PROGRAM" + suffix);
        string? rom = settings.Rom ?? Variable("MACEXCEL_ROM");
        if (program is null || rom is null)
        {
            AnsiConsole.MarkupLine($"[red]Pass --program and --rom, or set MACEXCEL_PROGRAM{suffix} and MACEXCEL_ROM (and --system or MACEXCEL_SYSTEM{suffix} unless the Program disk has a System Folder).[/]");
            return 2;
        }

        var captures = settings.Capture.Select(ParseCapture).ToList();
        var (testVolume, testFiles) = MfsImage.Read(File.ReadAllBytes(settings.Disk));
        if (!testFiles.Any(f => f.Name == ExcelTest.DriverName))
        {
            AnsiConsole.MarkupLine("[red]No Driver on the disk: write it with excel-test pack.[/]");
            return 2;
        }

        string work = settings.Work ?? Directory.CreateTempSubdirectory("excel-run-").FullName;
        Directory.CreateDirectory(work);
        bool keep = settings.Work is not null;
        try
        {
            string application = WriteBootDisk(system, program, Path.Combine(work, "boot.img"));
            WriteRoms(rom, Path.Combine(work, "roms", "macplus"));
            string disk = Path.Combine(work, "test.img");
            File.Copy(settings.Disk, disk, overwrite: true);
            File.WriteAllText(Path.Combine(work, "excel-test.lua"), Script(application, testVolume, captures, settings));

            int exit = RunMame(settings, work, disk);
            string log = Path.Combine(work, "result.txt");
            var results = File.Exists(log) ? File.ReadAllLines(log) : [];
            foreach (string line in results)
            {
                AnsiConsole.MarkupLine(Markup.Escape(line));
            }

            if (settings.Screen is { } screen && File.Exists(Path.Combine(work, "screen.bin")))
            {
                File.WriteAllBytes(screen, ReadScreen(Path.Combine(work, "screen.bin"), null).ToPng());
            }

            if (!results.Contains("done"))
            {
                AnsiConsole.MarkupLine($"[red]The run did not finish (MAME exit code {exit}).[/] The disk is unchanged; MAME's files are in {Markup.Escape(work)}.");
                if (settings.Screen is null && File.Exists(Path.Combine(work, "screen.bin")))
                {
                    File.WriteAllBytes(Path.Combine(work, "screen.png"), ReadScreen(Path.Combine(work, "screen.bin"), null).ToPng());
                    AnsiConsole.MarkupLine($"The last screen is {Markup.Escape(Path.Combine(work, "screen.png"))}.");
                }

                keep = true;
                return 1;
            }

            File.Copy(disk, settings.Disk, overwrite: true);
            Directory.CreateDirectory(settings.Output);
            int missing = 0;
            for (int i = 0; i < captures.Count; i++)
            {
                string raw = Path.Combine(work, $"capture{i + 1}.bin");
                if (!File.Exists(raw))
                {
                    AnsiConsole.MarkupLine($"[yellow]No window titled \"{Markup.Escape(captures[i].Title)}\".[/]");
                    missing++;
                    continue;
                }

                string png = Path.Combine(settings.Output, captures[i].File);
                File.WriteAllBytes(png, ReadScreen(raw, File.ReadAllText(raw + ".rect")).ToPng());
                AnsiConsole.MarkupLine($"Captured \"{Markup.Escape(captures[i].Title)}\" as [bold]{Markup.Escape(png)}[/].");
            }

            AnsiConsole.MarkupLine($"Driver ran; Excel's saves are on [bold]{Markup.Escape(settings.Disk)}[/]. Next: [bold]excel-test check[/].");
            return missing == 0 ? 0 : 1;
        }
        finally
        {
            if (!keep)
            {
                Directory.Delete(work, recursive: true);
            }
        }
    }

    private static (string Title, string File) ParseCapture(string text)
    {
        int equals = text.LastIndexOf('=');
        return equals > 0 && text.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
            ? (text[..equals], text[(equals + 1)..])
            : (text, text.Replace(":", " - ") + ".png");
    }

    // A disk that boots straight into Excel, with Excel as the startup application named in the boot blocks
    // (bbHelloName, offset 0x5A), which the System launches instead of the Finder.
    //
    // An MFS Program disk (1.00) is kept as it is, since Excel 1.00's copy protection checks the disk it runs from:
    // it is grown to 800K, the System disk's System and Finder are added with everything already on it left where it
    // was, and the System disk's boot blocks are copied in.
    //
    // Later releases came on HFS disks, where the System looks for the startup application in the System Folder:
    // Excel is lifted off the disk (or out of a folder) with the System and Finder, from the disk's own System Folder if
    // it has one, onto a new 800K MFS disk, which is flat, with the boot blocks of the disk the System came from.
    private static string WriteBootDisk(string? systemImage, string program, string path)
    {
        List<(string Path, MfsFile File)> programFiles;
        byte[]? programBytes = null;
        if (Directory.Exists(program))
        {
            programFiles = MacFolder.Read(program);
        }
        else
        {
            programBytes = File.ReadAllBytes(program);
            programFiles = HfsImage.IsHfs(programBytes)
                ? HfsImage.Read(programBytes)
                : [.. MfsImage.Read(programBytes).Files.Select(f => (f.Name, f))];
        }

        var excel = programFiles.FirstOrDefault(f => f.File.Type == "APPL" && f.File.Creator == "XCEL");
        if (excel.File is null)
        {
            throw new InvalidOperationException($"No Excel application (type APPL, creator XCEL) on {program}.");
        }

        // The System and Finder: the Program disk's own, in its System Folder, or the System disk's.
        static bool IsSystem(MfsFile f) => f.Type == "ZSYS" && f.Name == "System";
        static bool IsFinder(MfsFile f) => f.Type == "FNDR" && f.Name == "Finder";
        static string Folder(string path) => path.Contains(':') ? path[..path.LastIndexOf(':')] : "";
        var system = programFiles.FirstOrDefault(f => IsSystem(f.File));
        var finder = programFiles.FirstOrDefault(f => IsFinder(f.File) && system.File is not null && Folder(f.Path) == Folder(system.Path));
        bool hfs = programBytes is not null && HfsImage.IsHfs(programBytes);
        byte[] bootBlocks;
        if (hfs && system.File is not null && finder.File is not null)
        {
            bootBlocks = programBytes![..1024];
        }
        else
        {
            if (systemImage is null)
            {
                throw new InvalidOperationException($"{program} has no System Folder: pass --system, a startup disk.");
            }

            byte[] systemBytes = File.ReadAllBytes(systemImage);
            var systemFiles = HfsImage.IsHfs(systemBytes) ? HfsImage.Read(systemBytes) : [.. MfsImage.Read(systemBytes).Files.Select(f => (f.Name, f))];
            system = systemFiles.FirstOrDefault(f => IsSystem(f.File));
            finder = systemFiles.FirstOrDefault(f => IsFinder(f.File) && system.File is not null && Folder(f.Path) == Folder(system.Path));
            if (system.File is null || finder.File is null)
            {
                throw new InvalidOperationException($"No System and Finder on {systemImage}.");
            }

            bootBlocks = systemBytes[..1024];
        }

        if (bootBlocks[0] != 0x4C || bootBlocks[1] != 0x4B)
        {
            throw new InvalidOperationException("The disk with the System has no boot blocks: it is not a startup disk.");
        }

        byte[] image;
        if (programBytes is not null && !hfs)
        {
            if (excel.File.Name.Length > MaxStartupName)
            {
                throw new InvalidOperationException($"The startup application's name, {excel.File.Name}, is longer than {MaxStartupName} characters.");
            }

            var add = new[] { system.File, finder.File }.Where(f => !programFiles.Any(p => p.File.Name == f.Name)).ToList();
            image = MfsImage.AddFiles(programBytes, add, 800);
        }
        else
        {
            // Excel, renamed if its name is too long for the boot blocks ("Microsoft Excel 1.01F"), and its help
            // file beside it, which it opens by name.
            if (excel.File.Name.Length > MaxStartupName)
            {
                excel = (excel.Path, excel.File with { Name = "Microsoft Excel" });
            }

            var files = new List<MfsFile> { system.File, finder.File, excel.File };
            files.AddRange(programFiles.Where(f => Folder(f.Path) == Folder(excel.Path) && f.File.Creator == "XCEL" && f.File.Type == "XHLP").Select(f => f.File));
            image = MfsImage.Write("Excel", files, 800);
        }

        bootBlocks.CopyTo(image, 0);
        var hello = image.AsSpan(0x5A, 16);
        hello.Clear();
        hello[0] = (byte)excel.File.Name.Length;
        Encoding.GetEncoding(10000).GetBytes(excel.File.Name).CopyTo(hello[1..]);
        File.WriteAllBytes(path, image);
        return excel.File.Name;
    }

    // MAME's macplus set: the ROM's even and odd bytes, and the keyboard's microcontroller, which the run does
    // not need (the script posts key events itself) but MAME will not start without. Its stand-in runs nothing,
    // and MAME warns that its checksum is wrong.
    private static void WriteRoms(string rom, string folder)
    {
        byte[] bytes = File.ReadAllBytes(rom);
        if (bytes.Length < 0x20000 || bytes[0] != 0x4D || bytes[1] != 0x1F || bytes[2] != 0x81 || bytes[3] != 0x72)
        {
            throw new InvalidOperationException($"{rom} is not the Mac Plus ROM (128K, checksum 4D1F8172).");
        }

        Directory.CreateDirectory(folder);
        var even = new byte[0x10000];
        var odd = new byte[0x10000];
        for (int i = 0; i < 0x10000; i++)
        {
            even[i] = bytes[2 * i];
            odd[i] = bytes[2 * i + 1];
        }

        File.WriteAllBytes(Path.Combine(folder, "342-0341-c.u6d"), even);
        File.WriteAllBytes(Path.Combine(folder, "342-0342-b.u8d"), odd);
        File.WriteAllBytes(Path.Combine(folder, "341-0332-a.bin"), new byte[1024]);
    }

    private static string Script(string application, string volume, List<(string Title, string File)> captures, RunSettings settings)
    {
        static string Quote(string text) => "\"" + string.Concat(text.Select(c => c is '\\' or '"' || c < ' ' || c > '~' ? $"\\{(int)c:D3}" : c.ToString())) + "\"";
        var config = new StringBuilder();
        config.AppendLine("CONFIG = {");
        config.AppendLine($"  application = {Quote(application)},");
        config.AppendLine($"  volume = {Quote(volume)},");
        config.AppendLine($"  driver = {Quote(ExcelTest.DriverName)},");
        config.AppendLine($"  finished = {Quote(ExcelTest.FinishedAlert)},");
        config.AppendLine($"  timeout = {settings.Timeout.ToString(CultureInfo.InvariantCulture)},");
        config.AppendLine($"  auto_open = {(settings.Excel == "1.5" ? "true" : "false")},");
        config.AppendLine($"  captures = {{ {string.Join(", ", captures.Select(c => Quote(MacRoman(c.Title))))} }},");
        config.AppendLine("}");
        using var stream = typeof(RunCommand).Assembly.GetManifestResourceStream("excel-test.lua")
            ?? throw new InvalidOperationException("excel-test.lua is not embedded in the dumper.");
        using var reader = new StreamReader(stream);
        return config + reader.ReadToEnd();
    }

    // Lua strings are bytes: the titles go in as Mac Roman, as the Mac holds them.
    private static string MacRoman(string text)
    {
        var encoding = Encoding.GetEncoding(10000);
        return string.Concat(encoding.GetBytes(text).Select(b => (char)b));
    }

    private static int RunMame(RunSettings settings, string work, string disk)
    {
        var start = new ProcessStartInfo(settings.Mame)
        {
            WorkingDirectory = work,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        string[] arguments =
        [
            "macplus", "-rompath", Path.Combine(work, "roms"),
            "-flop1", Path.Combine(work, "boot.img"), "-flop2", disk,
            "-autoboot_script", Path.Combine(work, "excel-test.lua"),
            "-skip_gameinfo", "-sound", "none", "-nvram_directory", Path.Combine(work, "nvram"),
            "-cfg_directory", Path.Combine(work, "cfg"), "-snapshot_directory", Path.Combine(work, "snap"),
            "-seconds_to_run", (settings.Timeout + 60).ToString(CultureInfo.InvariantCulture),
        ];
        foreach (string argument in arguments)
        {
            start.ArgumentList.Add(argument);
        }

        // Fully in the background: with -video none MAME's SDL layer still opens a blank window and registers as a
        // foreground app, taking focus, unless SDL_VIDEODRIVER and SDL_AUDIODRIVER are "dummy" (the -videodriver and
        // -audiodriver options alone are not enough). The script needs no host input: it posts its keys in memory.
        string[] background =
        [
            "-video", "none", "-nothrottle", "-videodriver", "dummy", "-audiodriver", "dummy",
            "-keyboardprovider", "none", "-mouseprovider", "none", "-joystickprovider", "none", "-lightgunprovider", "none",
        ];
        foreach (string argument in settings.Show ? ["-window"] : background)
        {
            start.ArgumentList.Add(argument);
        }

        if (!settings.Show)
        {
            start.Environment["SDL_VIDEODRIVER"] = "dummy";
            start.Environment["SDL_AUDIODRIVER"] = "dummy";
        }

        Process process;
        try
        {
            process = Process.Start(start) ?? throw new InvalidOperationException($"Could not start {settings.Mame}.");
        }
        catch (System.ComponentModel.Win32Exception e)
        {
            throw new InvalidOperationException($"Could not start {settings.Mame} ({e.Message}). Install MAME (brew install mame) or pass --mame.", e);
        }

        using var _ = process;
        var output = new StringBuilder();
        process.OutputDataReceived += (_, e) => { if (e.Data is not null) { lock (output) { output.AppendLine(e.Data); } } };
        process.ErrorDataReceived += (_, e) => { if (e.Data is not null) { lock (output) { output.AppendLine(e.Data); } } };
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        process.WaitForExit();
        File.WriteAllText(Path.Combine(work, "mame.log"), output.ToString());
        return process.ExitCode;
    }

    // The script writes the 512 × 342 screen buffer as it is in memory; a rectangle "left top right bottom" crops it
    // to a window, title bar, frame and shadow included.
    private static MonoBitmap ReadScreen(string path, string? rectangle)
    {
        const int Width = 512, Height = 342, Stride = Width / 8;
        byte[] screen = File.ReadAllBytes(path);
        int left = 0, top = 0, right = Width, bottom = Height;
        if (rectangle is not null)
        {
            var parts = rectangle.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(p => int.Parse(p, CultureInfo.InvariantCulture)).ToArray();
            (left, top, right, bottom) = (Math.Max(parts[0], 0), Math.Max(parts[1], 0), Math.Min(parts[2], Width), Math.Min(parts[3], Height));
        }

        var bitmap = new MonoBitmap(right - left, bottom - top);
        for (int y = top; y < bottom; y++)
        {
            for (int x = left; x < right; x++)
            {
                bitmap[x - left, y - top] = (screen[y * Stride + (x >> 3)] & (0x80 >> (x & 7))) != 0;
            }
        }

        // The shadow is the frame moved 1 pixel right and down, so the top right and bottom left corners are the
        // desktop's: white, as ExcelRenderer draws them.
        if (rectangle is not null)
        {
            bitmap[bitmap.Width - 1, 0] = false;
            bitmap[0, bitmap.Height - 1] = false;
        }

        return bitmap;
    }
}
