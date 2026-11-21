using System.ComponentModel;
using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Spectre.Console;
using Spectre.Console.Cli;
using MacExcelReader;
using MacExcelReader.Charts;

// Round trip through a real Excel: `pack` puts files on a floppy image with a generated Driver macro sheet;
// in Excel, run Driver (Macro > Run, or automatically in 1.5); `check` reads Excel's results back. See
// docs/Testing.md.

sealed class PackSettings : CommandSettings
{
    [CommandArgument(0, "<disk>")]
    [Description("The floppy image to write (.dsk).")]
    public required string Disk { get; init; }

    [CommandArgument(1, "<files>")]
    [Description("Worksheets, macro sheets and charts (data forks) to test.")]
    public required string[] Files { get; init; }

    [CommandOption("--volume <NAME>")]
    [Description("The disk's name (default: Test). Driver opens and saves files by Volume:Name.")]
    public string Volume { get; init; } = "Test";

    [CommandOption("--excel <VERSION>")]
    [Description("The Excel that will run Driver: 1.0 (the default) or 1.5, where Driver also runs by itself when opened (Auto_Open).")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--keep-open")]
    [Description("Leave each file open after saving it, so its windows can be looked at.")]
    public bool KeepOpen { get; init; }

    [CommandOption("--no-probe")]
    [Description("Do not record GET.FORMULA for formula cells; only open and re-save.")]
    public bool NoProbe { get; init; }
}

sealed class CheckSettings : CommandSettings
{
    [CommandArgument(0, "<disk>")]
    [Description("The floppy image after Driver has run.")]
    public required string Disk { get; init; }

    [CommandOption("-o|--output <DIR>")]
    [Description("Also extract every file on the disk to this folder.")]
    public string? Output { get; init; }

    [CommandOption("--all")]
    [Description("Show every difference, including the ones Excel is expected to make (recalculated values, windows, table size).")]
    public bool All { get; init; }
}

static class ExcelTest
{
    public const string DriverName = "Driver";
    public const string SavedSuffix = " saved";

    // The suffix keeps names within the 31 characters Excel's dialogs allow.
    public static string SavedName(string name) => (name.Length + SavedSuffix.Length > 31 ? name[..(31 - SavedSuffix.Length)] : name) + SavedSuffix;

    public static bool IsChart(byte[] data) => data.Length >= 2 && data[0] == 0x01 && data[1] == 0x8C;
}

sealed class PackCommand : Command<PackSettings>
{
    public override int Execute(CommandContext context, PackSettings settings, CancellationToken cancellationToken)
    {
        var version = settings.Excel == "1.5" ? ExcelVersion.Excel150 : ExcelVersion.Excel100;
        var files = new List<MfsFile>();
        var targets = new List<(string Name, ExcelDocument? Sheet)>();
        foreach (string path in settings.Files)
        {
            string name = Path.GetFileName(path).Replace(':', '-');
            if (name.Length > 31 || name == ExcelTest.DriverName)
            {
                throw new InvalidOperationException($"Rename {path}: names must be at most 31 characters and not \"{ExcelTest.DriverName}\".");
            }

            byte[] data = File.ReadAllBytes(path);
            if (ExcelTest.IsChart(data))
            {
                files.Add(new MfsFile(name, "MCBN", "XCEL", data, []));
                targets.Add((name, null));
            }
            else
            {
                var sheet = ExcelDocument.Read(data);
                files.Add(new MfsFile(name, sheet.FinderType, "XCEL", data, []));
                targets.Add((name, sheet));
            }
        }

        var (driver, probes) = BuildDriver(settings, version, targets);
        files.Insert(0, new MfsFile(ExcelTest.DriverName, "XLPG", "XCEL", driver.ToArray(), []));

        int size = files.Sum(f => (f.Data.Length + 1023) / 1024 + 1) <= 380 ? 400 : 800;
        File.WriteAllBytes(settings.Disk, MfsImage.Write(settings.Volume, files, size));
        AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(settings.Disk)}[/] ({size}K, volume \"{Markup.Escape(settings.Volume)}\"): Driver and {targets.Count} file(s), {probes} formula cells to probe.");
        if (probes > 300)
        {
            AnsiConsole.MarkupLine("[yellow]Driver is large; Excel 1.03 may run out of memory. Pack fewer or smaller files.[/]");
        }

        AnsiConsole.MarkupLine(version == ExcelVersion.Excel150
            ? "In Excel 1.5: open Driver; it runs by itself. Then run [bold]excel-test check[/]."
            : "In Excel 1.03: open Driver, choose Macro > Run > Driver!Run (Option-Command-R). Then run [bold]excel-test check[/].");
        return 0;
    }

    // Driver: column A holds the commands; column B labels each one, so check can map results back. A probe
    // row is =GET.FORMULA(File!B4); when Driver saves itself, each probe's value is Excel's text for that cell.
    private static (ExcelDocument Driver, int Probes) BuildDriver(PackSettings settings, ExcelVersion version, List<(string Name, ExcelDocument? Sheet)> targets)
    {
        var driver = ExcelDocument.CreateMacroSheet();
        driver.State.SetColumnWidth(0, 200);
        driver.State.SetColumnWidth(1, 120);
        foreach (var (name, sheet) in targets)
        {
            if (sheet is not null)
            {
                driver.ExternalDocuments.Add(new ExcelExternalDocument(name));
            }
        }

        int row = 0, probes = 0;
        void Put(string formula, string label)
        {
            var parsed = ExcelFormula.Parse(formula, driver, version);
            parsed.SetResult(0);
            driver.Cells[row, 0].SetFormula(parsed);
            driver.Cells[row, 1].SetText(label);
            row++;
        }

        static string Quote(string text) => "\"" + text.Replace("\"", "\"\"") + "\"";
        static string Reference(string name) => Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_.]*$") ? name : "'" + name.Replace("'", "''") + "'";

        // Carry on past errors, so one bad file does not stop the rest; its probes come back as error values.
        Put("=ERROR(FALSE)", "start");
        foreach (var (name, sheet) in targets)
        {
            // 0: do not update links, which would otherwise stop Driver with a question.
            Put($"=OPEN({Quote(settings.Volume + ":" + name)},0)", "open " + name);
            if (sheet is not null && !settings.NoProbe)
            {
                foreach (var (r, c, cell) in sheet.Cells.NonEmptyCells)
                {
                    if (cell.Formula is not null)
                    {
                        string address = new ExcelCellAddress(r, c).ToString();
                        Put($"=GET.FORMULA({Reference(name)}!{address})", name + "!" + address);
                        probes++;
                    }
                }
            }

            Put($"=SAVE.AS({Quote(settings.Volume + ":" + ExcelTest.SavedName(name))})", "save " + name);
            if (!settings.KeepOpen)
            {
                Put("=CLOSE()", "close " + name);
            }
        }

        Put($"=ACTIVATE({Quote(ExcelTest.DriverName)})", "activate");
        Put("=SAVE()", "save Driver");
        Put("=RETURN()", "end");

        // Run is the command macro for Macro > Run (Option-Command-R); Excel 1.5 also runs Auto_Open on open.
        var names = new List<string> { "Run" };
        if (version == ExcelVersion.Excel150)
        {
            names.Add("Auto_Open");
        }

        foreach (string name in names)
        {
            driver.Names.Add(new ExcelDefinedName(name, [], ExcelDefinedName.ReferenceFlag | ExcelDefinedName.CommandFlag)
            {
                ShortcutKey = name == "Run" ? (byte)'r' : (byte)0,
                ShortcutExtra = 0x40,
            });
        }

        driver.Names.Sort((a, b) => string.CompareOrdinal(a.Name.ToUpperInvariant(), b.Name.ToUpperInvariant()));
        foreach (string name in names)
        {
            driver.Names.Single(n => n.Name == name).SetDefinition("=$A$1", driver, version);
        }

        return (driver, probes);
    }
}

sealed class CheckCommand : Command<CheckSettings>
{
    public override int Execute(CommandContext context, CheckSettings settings, CancellationToken cancellationToken)
    {
        var (volume, files) = MfsImage.Read(File.ReadAllBytes(settings.Disk));
        var byName = files.ToDictionary(f => f.Name, f => f);
        if (settings.Output is { } output)
        {
            Directory.CreateDirectory(output);
            foreach (var file in files)
            {
                File.WriteAllBytes(Path.Combine(output, file.Name.Replace('/', ':')), file.Data);
            }
        }

        if (!byName.TryGetValue(ExcelTest.DriverName, out var driverFile))
        {
            AnsiConsole.MarkupLine("[red]No Driver on the disk.[/]");
            return 1;
        }

        // Driver's results: label in column B, value of the command in column A.
        var driver = ExcelDocument.Read(driverFile.Data);
        var results = new Dictionary<string, ExcelCell>();
        for (int r = 0; r < driver.Cells.RowCount; r++)
        {
            if (driver.Cells.GetCellType(r, 1) == ExcelCellType.Text)
            {
                results[driver.Cells[r, 1].TextValue] = driver.Cells[r, 0];
            }
        }

        if (!results.TryGetValue("save Driver", out var done) || done.Formula?.Result is null)
        {
            AnsiConsole.MarkupLine("[yellow]Driver does not seem to have run to the end.[/]");
        }

        int problems = 0;
        var originals = files.Where(f => f.Name != ExcelTest.DriverName && !f.Name.EndsWith(ExcelTest.SavedSuffix, StringComparison.Ordinal) && f.Type is "XLBN" or "XLPG" or "MCBN");
        foreach (var original in originals)
        {
            AnsiConsole.Write(new Rule($"[bold]{Markup.Escape(original.Name)}[/]").LeftJustified());
            if (!byName.TryGetValue(ExcelTest.SavedName(original.Name), out var saved))
            {
                AnsiConsole.MarkupLine("[red]Excel did not save it[/] (it could not open it, or Driver did not reach it).");
                problems++;
                continue;
            }

            AnsiConsole.MarkupLine(original.Data.AsSpan().SequenceEqual(saved.Data)
                ? "  Bytes: identical."
                : $"  Bytes: ours {original.Data.Length:N0}, Excel {saved.Data.Length:N0}; {DifferingBytes(original.Data, saved.Data):N0} differ in the common length.");
            problems += original.Type == "MCBN"
                ? CompareJson(JsonCommand.ChartModel(ChartDocument.Read(original.Data)), JsonCommand.ChartModel(ChartDocument.Read(saved.Data)), settings.All)
                : CheckSheet(original, saved, results, settings.All);
        }

        AnsiConsole.MarkupLine(problems == 0 ? "[green]No unexpected differences.[/]" : $"[yellow]{problems} unexpected difference(s).[/]");
        return problems == 0 ? 0 : 1;
    }

    private static int CheckSheet(MfsFile original, MfsFile saved, Dictionary<string, ExcelCell> results, bool all)
    {
        var ours = ExcelDocument.Read(original.Data);
        var excel = ExcelDocument.Read(saved.Data);
        int problems = 0, probed = 0;

        // What Excel itself made of each formula, from GET.FORMULA.
        foreach (var (r, c, cell) in ours.Cells.NonEmptyCells)
        {
            if (cell.Formula is not { } formula || !results.TryGetValue(original.Name + "!" + new ExcelCellAddress(r, c), out var result))
            {
                continue;
            }

            probed++;
            string intended = formula.ToText(ours);
            if (result.Formula?.Result is not { } value || value.Kind != ExcelFormulaResultKind.Text)
            {
                AnsiConsole.MarkupLine($"  [red]{new ExcelCellAddress(r, c)}[/] {Markup.Escape(intended)}: GET.FORMULA gave {Markup.Escape(result.Formula is { } f ? f.Result.Value?.ToString() ?? "nothing" : result.Type.ToString())}");
                problems++;
                continue;
            }

            string text = value.Text!;
            string back = R1C1.ToA1(text, r, c);
            if (!R1C1.Same(intended, back))
            {
                AnsiConsole.MarkupLine($"  [yellow]{new ExcelCellAddress(r, c)}[/] {Markup.Escape(intended)} | Excel: {Markup.Escape(text)} | as A1: {Markup.Escape(back)}");
                problems++;
            }
        }

        if (probed > 0)
        {
            AnsiConsole.MarkupLine($"  GET.FORMULA: {probed} formula(s) probed.");
        }

        problems += CompareJson(JsonCommand.SheetModel(ours, false), JsonCommand.SheetModel(excel, false), all);
        return problems;
    }

    // Compares the JSON models of our file and Excel's re-save. Unless all is set, skips what Excel is expected
    // to change: formula results (it recalculates), windows, the code signature and the table's size.
    private static int CompareJson(object ours, object excel, bool all)
    {
        var options = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
        var a = JsonSerializer.SerializeToNode(ours, options)!;
        var b = JsonSerializer.SerializeToNode(excel, options)!;
        var differences = new List<string>();
        Compare(a, b, "", differences, all);
        foreach (string difference in differences.Take(50))
        {
            AnsiConsole.MarkupLine($"  [yellow]{Markup.Escape(difference)}[/]");
        }

        if (differences.Count > 50)
        {
            AnsiConsole.MarkupLine($"  … and {differences.Count - 50} more.");
        }

        if (differences.Count == 0)
        {
            AnsiConsole.MarkupLine("  Excel's re-save matches.");
        }

        return differences.Count;
    }

    private static int DifferingBytes(byte[] a, byte[] b)
    {
        int count = 0;
        for (int i = 0; i < Math.Min(a.Length, b.Length); i++)
        {
            count += a[i] != b[i] ? 1 : 0;
        }

        return count;
    }

    // Top-level fields Excel is expected to change on re-saving: the code signature and chart version say which
    // Excel saved it, the table size and windows follow from what Excel did while the file was open.
    private static readonly HashSet<string> s_expected = ["signature", "version", "rows", "columns", "windows"];

    private static void Compare(JsonNode? a, JsonNode? b, string path, List<string> differences, bool all)
    {
        if (a is JsonObject oa && b is JsonObject ob)
        {
            foreach (var key in oa.Select(p => p.Key).Union(ob.Select(p => p.Key)))
            {
                if (!all && path.Length == 0 && s_expected.Contains(key))
                {
                    continue;
                }

                // Excel's Save keeps only the external names some formula uses (spec §10).
                if (!all && key == "names" && path.StartsWith("externalDocuments", StringComparison.Ordinal)
                    && oa[key] is JsonArray na && ob[key] is JsonArray nb
                    && nb.All(n => na.Any(m => JsonNode.DeepEquals(m, n))))
                {
                    continue;
                }

                if (key == "cells" && oa[key] is JsonArray ca && ob[key] is JsonArray cb)
                {
                    CompareCells(ca, cb, differences, all);
                    continue;
                }

                Compare(oa[key], ob[key], path.Length == 0 ? key : path + "." + key, differences, all);
            }
        }
        else if (a is JsonArray aa && b is JsonArray ab)
        {
            for (int i = 0; i < Math.Max(aa.Count, ab.Count); i++)
            {
                Compare(i < aa.Count ? aa[i] : null, i < ab.Count ? ab[i] : null, $"{path}[{i}]", differences, all);
            }
        }
        else if (!JsonNode.DeepEquals(a, b))
        {
            differences.Add($"{path}: ours {a?.ToJsonString() ?? "(none)"}, Excel {b?.ToJsonString() ?? "(none)"}");
        }
    }

    private static void CompareCells(JsonArray ours, JsonArray excel, List<string> differences, bool all)
    {
        var a = ours.OfType<JsonObject>().ToDictionary(c => (string)c["address"]!);
        var b = excel.OfType<JsonObject>().ToDictionary(c => (string)c["address"]!);
        foreach (var address in a.Keys.Union(b.Keys))
        {
            a.TryGetValue(address, out var x);
            b.TryGetValue(address, out var y);
            if (x is null || y is null)
            {
                differences.Add(x is null ? $"{address}: only in Excel's copy ({y!["type"]} {y["display"]})" : $"{address}: lost by Excel ({x["type"]} {x["display"]})");
                continue;
            }

            bool formula = x["formula"] is not null && JsonNode.DeepEquals(x["formula"], y["formula"]);
            foreach (var (key, value) in x)
            {
                if (!all && formula && key is "value" or "display" or "type")
                {
                    continue;   // recalculated
                }

                if (!JsonNode.DeepEquals(value, y[key]))
                {
                    differences.Add($"{address}.{key}: ours {value?.ToJsonString() ?? "null"}, Excel {y[key]?.ToJsonString() ?? "null"}");
                }
            }
        }
    }
}

/// <summary>
/// Turns the R1C1 text GET.FORMULA returns into A1 text, and compares it with the decompiler's.
/// </summary>
static class R1C1
{
    private static readonly Regex s_wholeRows = new(@"(?<![A-Za-z.])R(\[-?\d+\]|\d+):R(\[-?\d+\]|\d+)(?![C\d\[])", RegexOptions.Compiled);
    private static readonly Regex s_cell = new(@"(?<![A-Za-z.])R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)(?![A-Za-z(])", RegexOptions.Compiled);
    private static readonly Regex s_external = new(@"':([^']+)'!", RegexOptions.Compiled);
    private static readonly Regex s_number = new(@"(?<![A-Za-z$\d.\[])\d+(\.\d+)?([Ee][+-]?\d+)?(?![\d\]])", RegexOptions.Compiled);

    public static string ToA1(string text, int row, int column)
    {
        string Part(string group, int origin, out bool absolute)
        {
            absolute = !(group.Length == 0 || group[0] == '[');
            return (absolute ? int.Parse(group, CultureInfo.InvariantCulture) : origin + 1 + (group.Length == 0 ? 0 : int.Parse(group[1..^1], CultureInfo.InvariantCulture))).ToString(CultureInfo.InvariantCulture);
        }

        text = s_wholeRows.Replace(text, m =>
        {
            string first = Part(m.Groups[1].Value, row, out bool absolute);
            string last = Part(m.Groups[2].Value, row, out _);
            string d = absolute ? "$" : "";
            return $"{d}A{d}{first}:{d}IV{d}{last}";
        });

        text = s_cell.Replace(text, m =>
        {
            string r = Part(m.Groups[1].Value, row, out bool absoluteRow);
            int c = int.Parse(Part(m.Groups[2].Value, column, out bool absoluteColumn), CultureInfo.InvariantCulture);
            return $"{(absoluteColumn ? "$" : "")}{FormulaDecompiler.ColumnName(c - 1)}{(absoluteRow ? "$" : "")}{r}";
        });

        // Excel names an external document ':Name' (relative to the current folder).
        return s_external.Replace(text, m => Regex.IsMatch(m.Groups[1].Value, @"^\w+$") ? m.Groups[1].Value + "!" : $"'{m.Groups[1].Value}'!");
    }

    // Equal apart from case, array braces and how numbers are written (1.5E-10 / 0.00000000015).
    public static bool Same(string intended, string excel)
    {
        string Normalise(string text)
        {
            if (text.StartsWith("{=", StringComparison.Ordinal) && text.EndsWith('}'))
            {
                text = text[1..^1];
            }

            return s_number.Replace(text, m => double.Parse(m.Value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)).ToUpperInvariant();
        }

        return Normalise(intended) == Normalise(excel);
    }
}
