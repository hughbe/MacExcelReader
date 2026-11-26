using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using MacExcelReader.Rendering;
using Spectre.Console;
using Spectre.Console.Cli;

// Runs a disk written by `excel-test pack` in a real Excel 1.00, headless: MAME's Mac Plus boots a disk built from
// the System and Program disks, with Excel as its startup application, and excel-test.lua types the keys that open
// Driver and run it, captures windows from the screen buffer and ejects the test disk. See docs/Testing.md.

sealed class RunSettings : CommandSettings
{
    [CommandArgument(0, "<disk>")]
    [Description("The floppy image written by excel-test pack. Excel's saves are written back to it.")]
    public required string Disk { get; init; }

    [CommandOption("--system <IMAGE>")]
    [Description("A bootable MFS disk with a System and Finder, such as Excel 1.00's Data disk. Defaults to $MACEXCEL_SYSTEM.")]
    public string? System { get; init; }

    [CommandOption("--program <IMAGE>")]
    [Description("The MFS disk with the Excel application, such as Excel 1.00's Program disk. Defaults to $MACEXCEL_PROGRAM.")]
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

    public override int Execute(CommandContext context, RunSettings settings, CancellationToken cancellationToken)
    {
        string? system = settings.System ?? Environment.GetEnvironmentVariable("MACEXCEL_SYSTEM");
        string? program = settings.Program ?? Environment.GetEnvironmentVariable("MACEXCEL_PROGRAM");
        string? rom = settings.Rom ?? Environment.GetEnvironmentVariable("MACEXCEL_ROM");
        if (system is null || program is null || rom is null)
        {
            AnsiConsole.MarkupLine("[red]Pass --system, --program and --rom, or set MACEXCEL_SYSTEM, MACEXCEL_PROGRAM and MACEXCEL_ROM.[/]");
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

    // A disk that boots straight into Excel: the Program disk, grown to 800K with the System disk's System and
    // Finder added, everything already on it left where it was, since Excel's copy protection checks the disk it
    // runs from. The boot blocks come from the System disk, with the startup application (bbHelloName, offset
    // 0x5A) set to Excel.
    private static string WriteBootDisk(string systemImage, string programImage, string path)
    {
        byte[] systemBytes = File.ReadAllBytes(systemImage);
        byte[] programBytes = File.ReadAllBytes(programImage);
        var (_, systemFiles) = MfsImage.Read(systemBytes);
        var (_, programFiles) = MfsImage.Read(programBytes);
        var excel = programFiles.FirstOrDefault(f => f.Type == "APPL" && f.Creator == "XCEL")
            ?? throw new InvalidOperationException($"No Excel application (type APPL, creator XCEL) on {programImage}.");
        if (excel.Name.Length > MaxStartupName)
        {
            throw new InvalidOperationException($"The startup application's name, {excel.Name}, is longer than {MaxStartupName} characters.");
        }

        var files = new List<MfsFile>();
        foreach (string name in new[] { "System", "Finder" }.Where(n => !programFiles.Any(f => f.Name == n)))
        {
            files.Add(systemFiles.FirstOrDefault(f => f.Name == name) ?? throw new InvalidOperationException($"No {name} on {systemImage}."));
        }

        byte[] image = MfsImage.AddFiles(programBytes, files, 800);
        systemBytes.AsSpan(0, 1024).CopyTo(image);
        if (image[0] != 0x4C || image[1] != 0x4B)
        {
            throw new InvalidOperationException($"{systemImage} has no boot blocks: it is not a startup disk.");
        }

        var hello = image.AsSpan(0x5A, 16);
        hello.Clear();
        hello[0] = (byte)excel.Name.Length;
        Encoding.GetEncoding(10000).GetBytes(excel.Name).CopyTo(hello[1..]);
        File.WriteAllBytes(path, image);
        return excel.Name;
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

        foreach (string argument in settings.Show ? ["-window"] : new[] { "-video", "none", "-nothrottle" })
        {
            start.ArgumentList.Add(argument);
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
