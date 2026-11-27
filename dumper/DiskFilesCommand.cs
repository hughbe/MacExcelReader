using System.ComponentModel;
using Spectre.Console;
using Spectre.Console.Cli;

sealed class DiskFilesSettings : CommandSettings
{
    [CommandArgument(0, "<disk>")]
    [Description("A floppy image, MFS (Excel 1.00's disks, the harness's test disks) or HFS (Excel 1.01 to 1.06).")]
    public required string Disk { get; init; }

    [CommandOption("-o|--output <DIR>")]
    [Description("Copy every file out to this folder: NAME with its data fork and NAME.rsrc with its resource fork.")]
    public string? Output { get; init; }
}

/// <summary>
/// Lists, and with -o copies out, the files on a floppy image, as the harness reads them: to look at Excel's saves,
/// or at the System and fonts a release runs with, where macOS no longer mounts these disks.
/// </summary>
sealed class DiskFilesCommand : Command<DiskFilesSettings>
{
    public override int Execute(CommandContext context, DiskFilesSettings settings, CancellationToken cancellationToken)
    {
        byte[] image = File.ReadAllBytes(settings.Disk);
        var files = HfsImage.IsHfs(image)
            ? HfsImage.Read(image)
            : MfsImage.Read(image).Files.Select(f => (Path: f.Name, File: f)).ToList();
        var table = new Table().AddColumn("File").AddColumn("Type").AddColumn("Creator")
            .AddColumn(new TableColumn("Data").RightAligned()).AddColumn(new TableColumn("Resource").RightAligned());
        foreach (var (path, file) in files)
        {
            table.AddRow(Markup.Escape(path), Markup.Escape(file.Type), Markup.Escape(file.Creator), $"{file.Data.Length:N0}", $"{file.Resource.Length:N0}");
            if (settings.Output is { } output)
            {
                string target = Path.Combine(output, path.Replace('/', '-').Replace(':', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.WriteAllBytes(target, file.Data);
                if (file.Resource.Length > 0)
                {
                    File.WriteAllBytes(target + ".rsrc", file.Resource);
                }
            }
        }

        AnsiConsole.Write(table);
        return 0;
    }
}
