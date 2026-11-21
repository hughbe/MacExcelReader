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

    [CommandOption("--recompile")]
    [Description("Re-enter every formula from Excel's own text (FORMULA, or FORMULA.ARRAY for an array formula), so the re-saved file holds Excel's compilation of it; check then compares its tokens with ours.")]
    public bool Recompile { get; init; }

    [CommandOption("--each <FORMULA>")]
    [Description("A macro command to run on each file after probing, before it is saved, such as =CALCULATION(3). Repeatable. A file can also have its own commands, one per line, in FILE.commands next to it.")]
    public string[] Each { get; init; } = [];

    [CommandOption("--end <FORMULA>")]
    [Description("A macro command to run after Driver has saved itself, such as =FULL(TRUE). Repeatable.")]
    public string[] End { get; init; } = [];
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

    [CommandOption("--bytes")]
    [Description("Also list raw byte differences the JSON model does not show: the sheet state block, window state, name flag and shortcut bytes, and external-document entries.")]
    public bool Bytes { get; init; }
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
        var commands = new Dictionary<string, string[]>();
        foreach (string path in settings.Files)
        {
            string name = Path.GetFileName(path).Replace(':', '-');
            if (name.Length > 31 || name == ExcelTest.DriverName)
            {
                throw new InvalidOperationException($"Rename {path}: names must be at most 31 characters and not \"{ExcelTest.DriverName}\".");
            }

            byte[] data = File.ReadAllBytes(path);
            if (File.Exists(path + ".commands"))
            {
                commands[name] = File.ReadAllLines(path + ".commands").Where(l => l.TrimStart().StartsWith('=')).ToArray();
            }

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

        var (driver, probes) = BuildDriver(settings, version, targets, commands);
        files.Insert(0, new MfsFile(ExcelTest.DriverName, "XLPG", "XCEL", driver.ToArray(), []));

        int size = files.Sum(f => (f.Data.Length + 1023) / 1024 + 1) <= 380 ? 400 : 800;
        File.WriteAllBytes(settings.Disk, MfsImage.Write(settings.Volume, files, size));
        AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(settings.Disk)}[/] ({size}K, volume \"{Markup.Escape(settings.Volume)}\"): Driver and {targets.Count} file(s), {probes} formula cells to probe.");
        // Excel 1.03 ran out of memory opening a small file after a 345-row Driver had recompiled a 100-formula sheet.
        if (driver.Cells.RowCount > 250)
        {
            AnsiConsole.MarkupLine($"[yellow]Driver has {driver.Cells.RowCount} rows; Excel 1.03 may run out of memory. Pack fewer files, or recompile one sheet per disk.[/]");
        }

        AnsiConsole.MarkupLine(version == ExcelVersion.Excel150
            ? "In Excel 1.5: open Driver; it runs by itself. Then run [bold]excel-test check[/]."
            : "In Excel 1.03: open Driver, choose Macro > Run > Driver!Run (Option-Command-R). Then run [bold]excel-test check[/].");
        return 0;
    }

    // Driver: column A holds the commands; column B labels each one, so check can map results back. A probe
    // row is =GET.FORMULA(File!B4); when Driver saves itself, each probe's value is Excel's text for that cell.
    // Whether a formula refers to another document (an external-sheet token other than the active sheet's).
    private static bool ReferencesOtherDocuments(ExcelFormula formula)
    {
        foreach (var token in FormulaTokens.Enumerate(formula.IsArrayFormula ? formula.SharedExpression : formula.Expression))
        {
            if (token.Code == FormulaTokens.ExternalSheet)
            {
                return true;
            }
        }

        return false;
    }

    private static (ExcelDocument Driver, int Probes) BuildDriver(PackSettings settings, ExcelVersion version, List<(string Name, ExcelDocument? Sheet)> targets, Dictionary<string, string[]> commands)
    {
        var driver = ExcelDocument.CreateMacroSheet();
        driver.State.SetColumnWidth(0, 200);
        driver.State.SetColumnWidth(1, 120);
        // Probes use active-sheet references (!B4), which every release resolves against the document OPEN has
        // just brought to the front. Excel 1.00 does not resolve a reference to a named document opened after
        // Driver was loaded ('Name'!B4 gives #REF!).
        driver.ExternalDocuments.Add(new ExcelExternalDocument("\0"));

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

        // Carry on past errors, so one bad file does not stop the rest; its probes come back as error values.
        Put("=ERROR(FALSE)", "start");
        foreach (var (name, sheet) in targets)
        {
            // The file's block of rows, laid out first so the guards after OPEN know where the next file starts.
            // Each entry is a formula and its label; {P:n} in a formula stands for the row of the block's n-th entry.
            var block = new List<(string Formula, string Label)>();
            var probeRows = new List<(int Row, int Column, ExcelFormula Formula, int Entry)>();
            if (sheet is not null && !settings.NoProbe)
            {
                foreach (var (r, c, cell) in sheet.Cells.NonEmptyCells)
                {
                    if (cell.Formula is { } formula)
                    {
                        string address = new ExcelCellAddress(r, c).ToString();
                        probeRows.Add((r, c, formula, block.Count));
                        block.Add(($"=GET.FORMULA(!{address})", name + "!" + address));
                        probes++;
                    }
                }
            }

            // Excel compiles the text it just gave back, so its re-save holds its own tokens for each formula.
            // Formulas that refer to other documents are left alone: re-entering one makes Excel ask for the file.
            if (settings.Recompile)
            {
                foreach (var (r, c, formula, entry) in probeRows)
                {
                    string address = new ExcelCellAddress(r, c).ToString();
                    if (ReferencesOtherDocuments(formula))
                    {
                        continue;
                    }

                    if (formula.IsArrayFormula)
                    {
                        if (!formula.IsSharedAnchor || formula.SharedRange is not { } range)
                        {
                            continue;
                        }

                        string area = new ExcelCellAddress(range.FirstRow, range.FirstColumn) + ":" + new ExcelCellAddress(range.LastRow, range.LastColumn);
                        block.Add(($"=SELECT(!{area},!{address})", "select " + name + "!" + area));
                        block.Add(($"=FORMULA.ARRAY(A{{P:{entry}}})", "recompile " + name + "!" + address));
                    }
                    else
                    {
                        block.Add(($"=SELECT(!{address})", "select " + name + "!" + address));
                        block.Add(($"=FORMULA(A{{P:{entry}}})", "recompile " + name + "!" + address));
                    }
                }
            }

            foreach (string command in settings.Each.Concat(commands.GetValueOrDefault(name, [])))
            {
                block.Add((command.Trim(), "command " + name));
            }

            block.Add(($"=SAVE.AS({Quote(settings.Volume + ":" + ExcelTest.SavedName(name))})", "save " + name));
            if (!settings.KeepOpen)
            {
                block.Add(("=CLOSE()", "close " + name));
            }

            // 0: do not update links, which would otherwise stop Driver with a question. If OPEN fails (a missing
            // file, or Excel out of memory), skip to the next file rather than run its commands on Driver itself.
            int open = row, first = row + 3, next = first + block.Count;
            Put($"=OPEN({Quote(settings.Volume + ":" + name)},0)", "open " + name);
            Put($"=IF(ISERROR(A{open + 1}),GOTO(A{next + 1}))", "skip if not opened");
            Put($"=IF(NOT(A{open + 1}),GOTO(A{next + 1}))", "skip if not opened");
            foreach (var (formula, label) in block)
            {
                Put(Regex.Replace(formula, @"\{P:(\d+)\}", m => (first + int.Parse(m.Groups[1].Value) + 1).ToString()), label);
            }
        }

        Put($"=ACTIVATE({Quote(ExcelTest.DriverName)})", "activate");
        Put("=SAVE()", "save Driver");
        foreach (string command in settings.End)
        {
            Put(command, "end command");
        }

        Put("=RETURN()", "end");

        // Run is the command macro for Macro > Run (Option-Command-R); Excel 1.5 also runs Auto_Open on open.
        var names = new List<string> { "Run" };
        if (version == ExcelVersion.Excel150)
        {
            names.Add("Auto_Open");
        }

        foreach (string name in names)
        {
            driver.Names.Add(new ExcelDefinedName(name, [], ExcelDefinedName.ReferenceFlag | ExcelDefinedName.MacroFlag)
            {
                ShortcutKey = name == "Run" ? (byte)'r' : (byte)0,
                MacroKind = 0x40,
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
            if (original.Type == "MCBN")
            {
                var ours = ChartDocument.Read(original.Data);
                var excel = ChartDocument.Read(saved.Data);
                problems += CompareJson(JsonCommand.ChartModel(ours), JsonCommand.ChartModel(excel), settings.All);
                if (settings.Bytes)
                {
                    CompareChartBytes(ours, excel);
                }
            }
            else
            {
                problems += CheckSheet(original, saved, results, settings);
            }
        }

        AnsiConsole.MarkupLine(problems == 0 ? "[green]No unexpected differences.[/]" : $"[yellow]{problems} unexpected difference(s).[/]");
        return problems == 0 ? 0 : 1;
    }

    private static int CheckSheet(MfsFile original, MfsFile saved, Dictionary<string, ExcelCell> results, CheckSettings settings)
    {
        bool all = settings.All;
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

        // The tokens in Excel's re-save, with the bytes left over from memory cleared. After --recompile they are
        // Excel's own compilation of each formula.
        int compared = 0;
        foreach (var (r, c, cell) in ours.Cells.NonEmptyCells)
        {
            if (cell.Formula is not { } formula || (formula.IsArrayFormula && !formula.IsSharedAnchor) || excel.Cells.TryGet(r, c)?.Formula is not { } theirs)
            {
                continue;
            }

            compared++;
            var a = FormulaTokens.ClearStaleBytes(formula.IsArrayFormula ? formula.SharedExpression : formula.Expression);
            var b = FormulaTokens.ClearStaleBytes(theirs.IsArrayFormula ? theirs.SharedExpression : theirs.Expression);
            bool sameFixups = formula.IsArrayFormula || formula.GetFixupWords().SequenceEqual(theirs.GetFixupWords());
            if (!a.AsSpan().SequenceEqual(b) || !sameFixups)
            {
                AnsiConsole.MarkupLine($"  [yellow]{new ExcelCellAddress(r, c)} tokens[/] {Markup.Escape(formula.ToText(ours))}: ours {Convert.ToHexString(a)}, Excel {Convert.ToHexString(b)}" + (sameFixups ? "" : " (fix-ups differ)"));
                problems++;
            }
        }

        if (compared > 0)
        {
            AnsiConsole.MarkupLine($"  Tokens: {compared} formula(s) compared.");
        }

        if (settings.Bytes)
        {
            CompareBytes(ours, excel);
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

    // Chart records, text objects and the window block, byte by byte (offsets as in spec §16).
    private static void CompareChartBytes(ChartDocument ours, ChartDocument excel)
    {
        ReportBytes("Window", ours.Window, excel.Window);
        for (int i = 0; i < Math.Min(ours.Charts.Count, excel.Charts.Count); i++)
        {
            ReportBytes($"Chart {i + 1} record", ours.Charts[i].Record, excel.Charts[i].Record);
            for (int t = 0; t < Math.Min(ours.Charts[i].Texts.Count, excel.Charts[i].Texts.Count); t++)
            {
                ReportBytes($"Chart {i + 1} text {t + 1}", ours.Charts[i].Texts[t].Record, excel.Charts[i].Texts[t].Record);
            }

            if (ours.Charts[i].Texts.Count != excel.Charts[i].Texts.Count)
            {
                AnsiConsole.MarkupLine($"  [blue]Chart {i + 1} text objects[/]: ours {ours.Charts[i].Texts.Count}, Excel {excel.Charts[i].Texts.Count}");
            }
        }
    }

    private static void ReportBytes(string what, ReadOnlySpan<byte> a, ReadOnlySpan<byte> b, Func<int, bool>? skip = null)
    {
        var parts = new List<string>();
        for (int i = 0; i < Math.Max(a.Length, b.Length); i++)
        {
            int x = i < a.Length ? a[i] : -1, y = i < b.Length ? b[i] : -1;
            if (x != y && skip?.Invoke(i) != true)
            {
                parts.Add($"{i}: {(x < 0 ? "--" : x.ToString("X2"))}→{(y < 0 ? "--" : y.ToString("X2"))}");
            }
        }

        if (parts.Count > 0)
        {
            AnsiConsole.MarkupLine($"  [blue]{Markup.Escape(what)}[/] {string.Join(", ", parts)}");
        }
    }

    // Raw differences in the parts the JSON model leaves out. Informational: they do not count as problems.
    private static void CompareBytes(ExcelDocument ours, ExcelDocument excel)
    {
        void Report(string what, ReadOnlySpan<byte> a, ReadOnlySpan<byte> b, Func<int, bool>? skip = null)
        {
            var parts = new List<string>();
            for (int i = 0; i < Math.Max(a.Length, b.Length); i++)
            {
                int x = i < a.Length ? a[i] : -1, y = i < b.Length ? b[i] : -1;
                if (x != y && skip?.Invoke(i) != true)
                {
                    parts.Add($"0x{i:X3}: {(x < 0 ? "--" : x.ToString("X2"))}→{(y < 0 ? "--" : y.ToString("X2"))}");
                }
            }

            if (parts.Count > 0)
            {
                AnsiConsole.MarkupLine($"  [blue]{Markup.Escape(what)}[/] {string.Join(", ", parts)}");
            }
        }

        AnsiConsole.MarkupLine($"  Function count: ours {ours.FunctionCount}, Excel {excel.FunctionCount}; extra header: ours {Convert.ToHexString(ours.ExtraHeader)}, Excel {Convert.ToHexString(excel.ExtraHeader)}");

        // State block offsets the spec marks as memory pointers.
        static bool Garbage(int i) => i is >= 0x026 and < 0x03E or >= 0x256 and < 0x25A or >= 0x266 and < 0x26A;
        Report("State block", ours.State.RawData, excel.State.RawData, Garbage);
        for (int i = 0; i < Math.Min(ours.Windows.Count, excel.Windows.Count); i++)
        {
            // Window state bytes 23–25 and 29–32 are stale pointers and padding (spec §13).
            Report($"Window {i + 1} state", ours.Windows[i].State, excel.Windows[i].State, j => j is >= 23 and < 26 or >= 29 and < 33);
        }

        foreach (var name in ours.Names)
        {
            if (excel.Names.FirstOrDefault(n => n.Name == name.Name) is { } other)
            {
                Report($"Name {name.Name} flags, key, key flags", [name.Flags, name.ShortcutKey, name.MacroKind], [other.Flags, other.ShortcutKey, other.MacroKind]);
            }
        }

        for (int i = 0; i < Math.Min(ours.ExternalDocuments.Count, excel.ExternalDocuments.Count); i++)
        {
            Report($"External document {i + 1} entry bytes", ours.ExternalDocuments[i].Reserved, excel.ExternalDocuments[i].Reserved);
        }
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

                // A series with an empty reference gets a few bytes left over from memory as one (spec §16.9).
                if (!all && key is "firstReference" or "secondReference" && oa[key]?.GetValue<string>() is "")
                {
                    continue;
                }

                // An automatic scale limit is whatever Excel worked out the last time it drew the chart (spec §16.5).
                if (!all && key is "Minimum" or "Maximum" && oa[$"Auto{key}"]?.GetValue<bool>() == true)
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
