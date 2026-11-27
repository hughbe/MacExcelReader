using System.ComponentModel;
using System.Text;
using MacExcelReader;
using Spectre.Console;
using Spectre.Console.Cli;

sealed class ExperimentSettings : CommandSettings
{
    [CommandArgument(0, "<files>")]
    [Description("Worksheets, macro sheets and charts to run, each with its FILE.commands if it has one, or folders of them.")]
    public required string[] Files { get; init; }

    [CommandOption("-o|--output <DIR>")]
    [Description("Where to put everything: the disks, each disk's files as Excel left them, Excel's saves in saved/, and commands.tsv.")]
    public required string Output { get; init; }

    [CommandOption("--volume <NAME>")]
    [Description("The disks' name (default: X). With more than one disk they are X1, X2 and so on; Driver opens files by Volume:Name.")]
    public string Volume { get; init; } = "X";

    [CommandOption("--disk-size <K>")]
    [Description("The floppies' size, 400 or 800. By default each is 400 when its files fit and 800 otherwise.")]
    public int? DiskSize { get; init; }

    [CommandOption("--batch <N>")]
    [Description("At most this many files on a disk (default 30). A batch that does not fit a floppy is halved until it does.")]
    public int Batch { get; init; } = 30;

    [CommandOption("--keep-open")]
    [Description("Leave each file open after saving it.")]
    public bool KeepOpen { get; init; }

    [CommandOption("--no-probe")]
    [Description("Do not record GET.FORMULA for formula cells; only run commands and re-save.")]
    public bool NoProbe { get; init; }

    [CommandOption("--recompile")]
    [Description("Re-enter every formula from Excel's own text, so check compares Excel's tokens with ours.")]
    public bool Recompile { get; init; }

    [CommandOption("--save-as <TYPE>")]
    [Description("Save each file with SAVE.AS's type instead: 2 SYLK, 3 text, 4 WKS.")]
    public int? SaveAsType { get; init; }

    [CommandOption("--each <FORMULA>")]
    [Description("A macro command to run on every file before it is saved. Repeatable.")]
    public string[] Each { get; init; } = [];

    [CommandOption("--verbose")]
    [Description("Show each disk's run and check as they happen, instead of one line per disk (the check is always in DISK/check.txt).")]
    public bool Verbose { get; init; }

    [CommandOption("--all")]
    [Description("Have check show the differences Excel is expected to make too.")]
    public bool All { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("The Excel to run the files in: 1.0 (the default), 1.01, 1.03, 1.04, 1.06 or 1.5.")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--system <IMAGE>")]
    [Description("The startup disk (excel-test run --system). Defaults to $MACEXCEL_SYSTEM, or $MACEXCEL_SYSTEM_1_03 and so on.")]
    public string? System { get; init; }

    [CommandOption("--program <IMAGE>")]
    [Description("Excel's Program disk or folder (excel-test run --program). Defaults to $MACEXCEL_PROGRAM, or $MACEXCEL_PROGRAM_1_03 and so on.")]
    public string? Program { get; init; }

    [CommandOption("--rom <FILE>")]
    [Description("The Mac Plus ROM. Defaults to $MACEXCEL_ROM.")]
    public string? Rom { get; init; }

    [CommandOption("--mame <PATH>")]
    [Description("The MAME executable (default: mame on the PATH).")]
    public string Mame { get; init; } = "mame";

    [CommandOption("--timeout <SECONDS>")]
    [Description("Give up on a disk after this many seconds of emulated time (default 600, about 20 seconds). A disk that does not finish, usually because Excel crashed, is split and run again until the file that stops it is found.")]
    public int Timeout { get; init; } = 600;
}

/// <summary>
/// One command for a whole experiment in Excel under MAME: packs the files onto as many disks as they need, runs
/// Driver on each, checks Excel's saves, and reports what every command in the files' .commands returned.
/// </summary>
sealed class ExperimentCommand : Command<ExperimentSettings>
{
    public override int Execute(CommandContext context, ExperimentSettings settings, CancellationToken cancellationToken)
    {
        // A folder stands for the files in it, in name order, other than .commands files and hidden ones.
        string[] inputs = [.. settings.Files.SelectMany<string, string>(f => Directory.Exists(f)
            ? Directory.GetFiles(f).Where(p => !p.EndsWith(".commands", StringComparison.Ordinal) && !Path.GetFileName(p).StartsWith('.')).Order(StringComparer.Ordinal)
            : [f])];
        settings = new ExperimentSettings
        {
            Files = inputs, Output = settings.Output, Volume = settings.Volume, Batch = settings.Batch, KeepOpen = settings.KeepOpen,
            NoProbe = settings.NoProbe, Recompile = settings.Recompile, SaveAsType = settings.SaveAsType, Each = settings.Each,
            Verbose = settings.Verbose, All = settings.All, Excel = settings.Excel, System = settings.System, Program = settings.Program, Rom = settings.Rom,
            Mame = settings.Mame, Timeout = settings.Timeout, DiskSize = settings.DiskSize,
        };
        var missing = settings.Files.Where(f => !File.Exists(f)).ToList();
        if (missing.Count > 0)
        {
            AnsiConsole.MarkupLine($"[red]Not found: {Markup.Escape(string.Join(", ", missing))}[/]");
            return 2;
        }

        // Every file keeps its name on the disks and among the saves, so names must differ.
        var duplicate = settings.Files.GroupBy(Path.GetFileName).FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
        {
            AnsiConsole.MarkupLine($"[red]Two files are named \"{Markup.Escape(duplicate.Key!)}\": give them different names.[/]");
            return 2;
        }

        Directory.CreateDirectory(settings.Output);
        string disks = Directory.CreateDirectory(Path.Combine(settings.Output, "disks")).FullName;
        string saves = Directory.CreateDirectory(Path.Combine(settings.Output, "saved")).FullName;

        // Pack in batches, halving any batch whose disk overflows (the directory of a 400K or 800K floppy holds
        // about a hundred files, and every file also gets its saved copy).
        var pending = new Queue<string[]>(settings.Files.Chunk(Math.Max(1, settings.Batch)));
        var packed = new List<(string Volume, string Disk)>();
        var batches = new List<string[]>();   // the files on each packed disk
        var stoppers = new List<string>();     // files that stop Excel on a disk of their own
        var refusedFiles = new List<string>(); // files Excel would not open
        int total = (settings.Files.Length + Math.Max(1, settings.Batch) - 1) / Math.Max(1, settings.Batch);
        while (pending.Count > 0)
        {
            var batch = pending.Dequeue();
            string volume = total == 1 && pending.Count == 0 && packed.Count == 0 ? settings.Volume : settings.Volume + (packed.Count + 1);
            string disk = Path.Combine(disks, volume + ".dsk");
            try
            {
                Pack(settings, volume, disk, batch);
                packed.Add((volume, disk));
                batches.Add(batch);
            }
            catch (InvalidOperationException e) when (batch.Length > 1 && (e.Message.Contains("Too many files") || e.Message.Contains("do not fit")))
            {
                // Put both halves back at the front, in order.
                var rest = pending.ToList();
                pending.Clear();
                pending.Enqueue(batch[..(batch.Length / 2)]);
                pending.Enqueue(batch[(batch.Length / 2)..]);
                rest.ForEach(pending.Enqueue);
                total++;
            }
        }

        int failures = 0;
        var report = new StringBuilder("disk\tfile\trow\tcommand\tresult\n");
        var errors = new List<string>();
        var retried = new HashSet<string>();
        for (int d = 0; d < packed.Count; d++)
        {
            var (volume, disk) = packed[d];
            string files = Path.Combine(settings.Output, volume);
            AnsiConsole.Markup($"[bold]{Markup.Escape(volume)}[/]: ");
            Directory.CreateDirectory(files);
            var (run, check, log) = RunAndCheck(context, settings, disk, files, Path.Combine(files, "screen.png"), cancellationToken);
            string logPath = Path.Combine(files, "check.txt");
            File.WriteAllText(logPath, log);
            if (run != 0)
            {
                // Usually Excel crashed or hung on one of the files: run them again in halves until that file is alone.
                var onDisk = batches[d];
                if (onDisk.Length > 1)
                {
                    AnsiConsole.MarkupLine($"[yellow]did not run to the end; running its {onDisk.Length} files again in halves to find the one that stops Excel.[/]");
                    foreach (var half in new[] { onDisk[..(onDisk.Length / 2)], onDisk[(onDisk.Length / 2)..] })
                    {
                        string again = settings.Volume + (packed.Count + 1);
                        Pack(settings, again, Path.Combine(disks, again + ".dsk"), half);
                        packed.Add((again, Path.Combine(disks, again + ".dsk")));
                        batches.Add(half);
                    }
                }
                else
                {
                    AnsiConsole.MarkupLine($"[red]{Markup.Escape(Path.GetFileName(onDisk[0]))} stops Excel[/]: its last screen is {Markup.Escape(Path.Combine(files, "screen.png"))}.");
                    stoppers.Add(onDisk[0]);
                }

                failures++;
                continue;
            }

            AnsiConsole.MarkupLine(check == 0 ? "ran; Excel's saves match." : $"ran; check found differences ({Markup.Escape(logPath)}).");
            failures += check == 0 ? 0 : 1;

            // A save that is missing or did not reach the disk: run those files again, once, on a disk of their own.
            var (unsaved, refusedHere) = Unsaved(disk, settings.Files);
            foreach (string file in refusedHere)
            {
                AnsiConsole.MarkupLine($"  [red]Excel could not open {Markup.Escape(Path.GetFileName(file))}[/] (OPEN gave an error).");
                refusedFiles.Add(file);
            }

            if (unsaved.Count > 0 && unsaved.All(retried.Add))
            {
                AnsiConsole.MarkupLine($"  [yellow]Excel's save of {Markup.Escape(string.Join(", ", unsaved.Select(Path.GetFileName)))} did not reach the disk; running again.[/]");
                string again = settings.Volume + (packed.Count + 1);
                Pack(settings, again, Path.Combine(disks, again + ".dsk"), [.. unsaved]);
                packed.Add((again, Path.Combine(disks, again + ".dsk")));
                batches.Add([.. unsaved]);
            }

            foreach (string saved in Directory.GetFiles(files, "*" + ExcelTest.SavedSuffix))
            {
                if (Path.GetFileName(saved) != ExcelTest.SavedName(ExcelTest.DriverName) && CheckCommand.IsReadable(File.ReadAllBytes(saved)))
                {
                    File.Copy(saved, Path.Combine(saves, Path.GetFileName(saved)), overwrite: true);
                }
            }

            ReportCommands(Path.Combine(files, ExcelTest.DriverName), volume, report, errors);
        }

        File.WriteAllText(Path.Combine(settings.Output, "commands.tsv"), report.ToString());
        int commands = report.ToString().Count(c => c == '\n') - 1;
        AnsiConsole.Write(new Rule("[bold]Experiment[/]").LeftJustified());
        AnsiConsole.MarkupLine($"{settings.Files.Length} file(s) on {packed.Count} disk(s); Excel's saves are in [bold]{Markup.Escape(saves)}[/].");
        if (refusedFiles.Count > 0)
        {
            AnsiConsole.MarkupLine($"[red]Excel could not open {refusedFiles.Count} file(s)[/]: {Markup.Escape(string.Join(", ", refusedFiles.Select(Path.GetFileName)))}.");
        }

        foreach (string stopper in stoppers)
        {
            AnsiConsole.MarkupLine($"[red]{Markup.Escape(Path.GetFileName(stopper))} stops Excel[/] (crash or hang): run it on its own with --show to watch.");
        }
        if (commands > 0)
        {
            AnsiConsole.MarkupLine($"{commands} command(s), each with what it returned, in [bold]{Markup.Escape(Path.Combine(settings.Output, "commands.tsv"))}[/].");
            foreach (string error in errors.Take(20))
            {
                AnsiConsole.MarkupLine($"  [yellow]{Markup.Escape(error)}[/]");
            }

            if (errors.Count > 20)
            {
                AnsiConsole.MarkupLine($"  [yellow]… and {errors.Count - 20} more command(s) that returned an error.[/]");
            }
        }

        return failures == 0 ? 0 : 1;
    }

    // Runs Driver on a disk and checks it, with their output kept for the disk's check.txt unless --verbose shows it.
    private static (int Run, int Check, string Log) RunAndCheck(CommandContext context, ExperimentSettings settings, string disk, string files, string screen, CancellationToken cancellationToken)
    {
        var console = AnsiConsole.Console;
        var log = new StringWriter();
        if (!settings.Verbose)
        {
            AnsiConsole.Console = AnsiConsole.Create(new AnsiConsoleSettings { Out = new AnsiConsoleOutput(log), ColorSystem = ColorSystemSupport.NoColors });
        }

        try
        {
            int run = new RunCommand().Execute(context, new RunSettings
            {
                Disk = disk,
                Excel = settings.Excel,
                System = settings.System,
                Program = settings.Program,
                Rom = settings.Rom,
                Mame = settings.Mame,
                Timeout = settings.Timeout,
                Screen = screen,
            }, cancellationToken);
            int check = -1;
            if (run == 0)
            {
                try
                {
                    check = new CheckCommand().Execute(context, new CheckSettings { Disk = disk, Output = files, All = settings.All }, cancellationToken);
                }
                catch (Exception e) when (e is not OperationCanceledException)
                {
                    // A check that fails on one disk's files must not end the experiment: report it with the rest.
                    AnsiConsole.MarkupLine($"[red]check failed:[/] {Markup.Escape(e.ToString())}");
                    check = 2;
                }
            }

            return (run, check, log.ToString());
        }
        finally
        {
            AnsiConsole.Console = console;
        }
    }

    // The files of a run whose saves are missing or unreadable although Excel opened them (their data was lost on the
    // way to the disk), and the files Excel would not open at all (OPEN gave an error), which running again cannot help.
    private static (List<string> Lost, List<string> Refused) Unsaved(string disk, string[] files)
    {
        var onDisk = MfsImage.Read(File.ReadAllBytes(disk)).Files.ToDictionary(f => f.Name, f => f.Data);
        var refused = new HashSet<string>();
        if (onDisk.TryGetValue(ExcelTest.DriverName, out var driverData) && CheckCommand.IsReadable(driverData))
        {
            var driver = ExcelDocument.Read(driverData);
            for (int row = 0; row < driver.Cells.RowCount; row++)
            {
                if (driver.Cells.GetCellType(row, 1) == ExcelCellType.Text && driver.Cells[row, 1].TextValue is { } label
                    && label.StartsWith("open ", StringComparison.Ordinal) && driver.Cells.Find(row, 0)?.Formula?.Result.Kind == ExcelValueKind.Error)
                {
                    refused.Add(label["open ".Length..]);
                }
            }
        }

        var present = files.Where(f => onDisk.ContainsKey(Path.GetFileName(f))).ToList();
        return (
            [.. present.Where(f => !refused.Contains(Path.GetFileName(f))
                && (!onDisk.TryGetValue(ExcelTest.SavedName(Path.GetFileName(f)), out var saved) || !CheckCommand.IsReadable(saved)))],
            [.. present.Where(f => refused.Contains(Path.GetFileName(f)))]);
    }

    private static void Pack(ExperimentSettings settings, string volume, string disk, string[] files)
    {
        var pack = new PackSettings
        {
            Disk = disk,
            Files = files,
            Volume = volume,
            Excel = settings.Excel,
            KeepOpen = settings.KeepOpen,
            NoProbe = settings.NoProbe,
            Recompile = settings.Recompile,
            SaveAsType = settings.SaveAsType,
            Each = settings.Each,
            DiskSize = settings.DiskSize,
        };
        // pack's own advice (how to run Driver) does not apply here.
        var console = AnsiConsole.Console;
        AnsiConsole.Console = AnsiConsole.Create(new AnsiConsoleSettings { Out = new AnsiConsoleOutput(TextWriter.Null) });
        try
        {
            if (new PackCommand().Execute(null!, pack, default) != 0)
            {
                throw new InvalidOperationException($"Could not pack {disk}.");
            }
        }
        finally
        {
            AnsiConsole.Console = console;
        }
    }

    // Driver labels each command's row "command FILE"; Excel saved the value each one returned (a TRUE, FALSE, number
    // or error) as the row's result.
    private static void ReportCommands(string driverPath, string volume, StringBuilder report, List<string> errors)
    {
        if (!File.Exists(driverPath))
        {
            return;
        }

        var driver = ExcelDocument.Load(driverPath);
        for (int row = 0; row < driver.Cells.RowCount; row++)
        {
            // Commands from .commands files and --each, and --recompile's FORMULA for each formula ("recompile
            // NAME!F5"), whose error means Excel's compiler rejected the text.
            if (driver.Cells.GetCellType(row, 1) != ExcelCellType.Text || driver.Cells[row, 1].TextValue is not { } label
                || driver.Cells.Find(row, 0)?.Formula is not { } formula)
            {
                continue;
            }

            string file;
            if (label.StartsWith("command ", StringComparison.Ordinal))
            {
                file = label["command ".Length..];
            }
            else if (label.StartsWith("recompile ", StringComparison.Ordinal))
            {
                file = label["recompile ".Length..];
            }
            else
            {
                continue;
            }

            string text = formula.ToText(driver);
            var result = formula.Result;
            report.Append(volume).Append('\t').Append(file).Append('\t').Append(row + 1).Append('\t').Append(text).Append('\t').Append(result.ToString()).Append('\n');
            if (result.Kind == ExcelValueKind.Error)
            {
                errors.Add($"{file}: {text} gave {result}");
            }
        }
    }
}
