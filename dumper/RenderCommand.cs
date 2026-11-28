using System.ComponentModel;
using MacExcelReader;
using MacExcelReader.Rendering;
using Spectre.Console;
using Spectre.Console.Cli;

sealed class RenderSettings : FileSettings
{
    [CommandOption("-o|--output <FILE>")]
    [Description("The PNG to write. Defaults to the input's name with .png.")]
    public string? Output { get; init; }

    [CommandOption("-s|--system <FILE>")]
    [Description("Where Excel's fonts come from: an MFS disk image with a System file (such as Excel 1.00's Data disk), " +
        "or a System file's resource fork (on macOS, SYSTEM/..namedfork/rsrc). Defaults to $MACEXCEL_SYSTEM.")]
    public string? System { get; init; }

    [CommandOption("-w|--window <N>")]
    [Description("Which saved window to draw: 0 (the default) is the front one.")]
    public int Window { get; init; }

    [CommandOption("--no-frame")]
    [Description("Draw only the window's contents, without its title bar and border.")]
    public bool NoFrame { get; init; }

    [CommandOption("--no-selection")]
    [Description("Leave out the selection and active cell.")]
    public bool NoSelection { get; init; }

    [CommandOption("--title <TITLE>")]
    [Description("The title bar's text. Defaults to the file's name.")]
    public string? Title { get; init; }

    [CommandOption("--size <WIDTHxHEIGHT>")]
    [Description("The window's content size in pixels, instead of the size it was saved with.")]
    public string? Size { get; init; }

    [CommandOption("--scale <N>")]
    [Description("Enlarge each pixel to N × N in the PNG (default 1).")]
    public int Scale { get; init; } = 1;
}

sealed class RenderCommand : Command<RenderSettings>
{
    public override int Execute(CommandContext context, RenderSettings settings, CancellationToken cancellationToken)
    {
        string? system = settings.System ?? Environment.GetEnvironmentVariable("MACEXCEL_SYSTEM");
        if (system is null)
        {
            AnsiConsole.MarkupLine("[red]Excel's fonts are needed: pass --system with Excel 1.00's Data disk image or a System file's resource fork, or set MACEXCEL_SYSTEM.[/]");
            return 1;
        }

        var fonts = LoadFonts(system);
        MonoBitmap bitmap;
        if (Program.IsChart(settings.Input))
        {
            // A chart's window, at the size it was saved.
            bitmap = MacExcelReader.Rendering.ChartRenderer.Render(MacExcelReader.Charts.ChartDocument.Load(settings.Input), fonts,
                !settings.NoFrame, settings.Title ?? Path.GetFileName(settings.Input));
            return Write(bitmap, settings);
        }

        var document = ExcelDocument.Load(settings.Input);
        var options = new ExcelRenderOptions
        {
            Window = settings.Window,
            Frame = !settings.NoFrame,
            Selection = !settings.NoSelection,
            Title = settings.Title ?? Path.GetFileName(settings.Input),
        };

        if (settings.Size is { } size)
        {
            var parts = size.Split('x', 'X');
            options.Width = int.Parse(parts[0]);
            options.Height = int.Parse(parts[1]);
        }

        bitmap = ExcelRenderer.Render(document, fonts, options);
        return Write(bitmap, settings);
    }

    private static int Write(MonoBitmap bitmap, RenderSettings settings)
    {
        if (settings.Scale > 1)
        {
            bitmap = Enlarge(bitmap, settings.Scale);
        }

        string output = settings.Output ?? settings.Input + ".png";
        using (var stream = File.Create(output))
        {
            bitmap.WritePng(stream);
        }

        AnsiConsole.MarkupLine($"Wrote {Program.Escape(output)} ({bitmap.Width} × {bitmap.Height})");
        return 0;
    }

    // The fonts of a System file: from an MFS disk image's System file, or from a resource fork on its own.
    internal static MacFontSet LoadFonts(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        if (bytes.Length > 1026 && bytes[1024] == 0xD2 && bytes[1025] == 0xD7)
        {
            var (_, files) = MfsImage.Read(bytes);
            var file = files.FirstOrDefault(f => f.Name == "System" && f.Resource.Length > 0)
                ?? throw new InvalidOperationException($"{path} has no System file.");
            return MacFontSet.FromResourceFork(file.Resource);
        }

        return MacFontSet.FromResourceFork(bytes);
    }

    private static MonoBitmap Enlarge(MonoBitmap source, int scale)
    {
        var large = new MonoBitmap(source.Width * scale, source.Height * scale);
        for (int y = 0; y < large.Height; y++)
        {
            for (int x = 0; x < large.Width; x++)
            {
                large[x, y] = source[x / scale, y / scale];
            }
        }

        return large;
    }
}
