using System.ComponentModel;
using System.Text;
using Spectre.Console;
using Spectre.Console.Cli;
using MacExcelReader;

sealed class SheetSettings : CommandSettings
{
    [CommandArgument(0, "<list>")]
    [Description("A formula list: one formula per line, with the setup lines described in docs/Testing.md.")]
    public required string List { get; init; }

    [CommandArgument(1, "<output>")]
    [Description("The sheet to write. With --split, \"<output> 01\", \"<output> 02\"… instead.")]
    public required string Output { get; init; }

    [CommandOption("--split <COUNT>")]
    [Description("Write several sheets of at most COUNT formulas each. Excel 1.03 runs out of memory on about 500.")]
    public int Split { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("Compile for Excel 1.0 (1.00–1.04, the default) or 1.5.")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--worksheet")]
    [Description("Write worksheets instead of macro sheets.")]
    public bool Worksheet { get; init; }

    public override ValidationResult Validate()
        => File.Exists(List) ? ValidationResult.Success() : ValidationResult.Error($"File not found: {List}");
}

/// <summary>
/// Builds sheets from a formula list: column A holds each formula as text, column B the formula itself. Every
/// formula is checked to decompile back to its text before the sheet is written.
/// </summary>
sealed class SheetCommand : Command<SheetSettings>
{
    public override int Execute(CommandContext context, SheetSettings settings, CancellationToken cancellationToken)
    {
        var version = settings.Excel == "1.5" ? ExcelVersion.Excel150 : ExcelVersion.Excel100;
        var list = FormulaList.Load(settings.List);
        var parts = new List<(ExcelDocument Document, int Formulas)>();
        var skipped = new List<string>();
        ExcelDocument document = null!;
        int row = 0, partFormulas = 0, total = 0;
        string heading = string.Empty;

        void StartSheet()
        {
            document = list.CreateDocument(settings.Worksheet, version);
            document.Cells[0, 0].SetText("Intended formula (text)");
            document.Cells[0, 1].SetText("Formula");
            document.Cells[0, 0].Bold = document.Cells[0, 1].Bold = true;
            row = 1;
            partFormulas = 0;
            if (heading.Length > 0)
            {
                WriteHeading();
            }
        }

        void WriteHeading()
        {
            row++;
            document.Cells[row, 0].SetText(heading);
            document.Cells[row, 0].Bold = true;
            row++;
        }

        void MakeRoom(int needed)
        {
            if (settings.Split > 0 && partFormulas > 0 && partFormulas + needed > settings.Split)
            {
                parts.Add((document, partFormulas));
                StartSheet();
            }
        }

        StartSheet();
        foreach (var line in list.Lines)
        {
            if (line.Heading is { } text)
            {
                heading = text;
                if (settings.Split > 0 && partFormulas >= settings.Split)
                {
                    MakeRoom(int.MaxValue);
                }
                else
                {
                    WriteHeading();
                }

                continue;
            }

            int rows = line.ArrayRows;
            MakeRoom(Math.Max(rows, 1));
            try
            {
                if (rows > 0)
                {
                    document.SetArrayFormula(new ExcelRange((ushort)row, (ushort)(row + rows - 1), 1, 1), line.Formula!, version);
                }
                else
                {
                    var formula = ExcelFormula.Parse(line.Formula!, document, version);
                    formula.SetResult(0);
                    document.Cells[row, 1].SetFormula(formula);
                }
            }
            catch (ExcelFormulaSyntaxException exception)
            {
                skipped.Add($"{line.Formula}: {exception.Message}");
                continue;
            }

            document.Cells[row, 0].SetText(rows > 0 ? "{" + line.Formula + "}" : line.Formula!);
            row += Math.Max(rows, 1);
            partFormulas += Math.Max(rows, 1);
            total += Math.Max(rows, 1);
        }

        parts.Add((document, partFormulas));
        for (int i = 0; i < parts.Count; i++)
        {
            string path = parts.Count == 1 && settings.Split <= 0 ? settings.Output : $"{settings.Output} {i + 1:D2}";
            WriteAndCheck(path, parts[i].Document);
        }

        AnsiConsole.MarkupLine($"{total} formulas in {parts.Count} sheet(s).");
        foreach (var skip in skipped)
        {
            AnsiConsole.MarkupLine($"[yellow]Skipped[/] {Markup.Escape(skip)}");
        }

        return 0;
    }

    // Writes a sheet and makes sure every formula decompiles back to the text in column A.
    private static void WriteAndCheck(string path, ExcelDocument document)
    {
        var bytes = document.ToArray();
        File.WriteAllBytes(path, bytes);
        var reread = ExcelDocument.Read(bytes);
        int mismatches = 0;
        for (int r = 0; r < reread.Cells.RowCount; r++)
        {
            if (reread.Cells[r, 1].Formula is { } formula && reread.Cells[r, 0].Type == ExcelCellType.Text)
            {
                string expected = reread.Cells[r, 0].TextValue;
                string actual = formula.ToText(reread);
                if (!string.Equals(expected.Replace("E+", "E"), actual.Replace("E+", "E"), StringComparison.OrdinalIgnoreCase)
                    && !(expected.StartsWith('{') && reread.Cells[r - 1, 0].TextValue == expected))
                {
                    mismatches++;
                    AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(Path.GetFileName(path))} row {r + 1}[/]: {Markup.Escape(expected)} decompiles as {Markup.Escape(actual)}");
                }
            }
        }

        AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(path)}[/]: {reread.Cells.RowCount} rows, {bytes.Length:N0} bytes" + (mismatches == 0 ? "." : $", [yellow]{mismatches} formulas decompile differently[/]."));
    }
}

/// <summary>
/// A formula list file. Blank lines and lines starting with # are ignored. Setup lines, which apply to every
/// sheet built from the list:
/// <code>
/// width A 220                     column width
/// external Budget Rate Total      an external document and the names used from it ("!" = the active sheet)
/// name Total =$B$2                a defined name; add "macro", or "command K" for a command macro with key K
/// cell D1 =ARGUMENT("x")          a cell outside columns A and B, such as a macro the names point at
/// recalc 1                        the recalculation state word (0x018): 1 marks the sheet as needing it
/// </code>
/// Then "## Heading" lines and formulas, one per line. "{=formula} N" is an array formula over N rows.
/// </summary>
sealed class FormulaList
{
    public List<(string Column, ushort Width)> Widths { get; } = [];
    public List<(string Name, List<string> Names)> Externals { get; } = [];
    public List<(string Name, string Definition, byte Flags, char Key)> Names { get; } = [];
    public List<(string Address, string Formula)> Cells { get; } = [];
    public ushort? RecalculationState { get; private set; }
    public List<Line> Lines { get; } = [];

    public sealed record Line(string? Heading, string? Formula, int ArrayRows);

    public static FormulaList Load(string path)
    {
        var list = new FormulaList();
        int number = 0;
        foreach (string raw in File.ReadLines(path))
        {
            number++;
            string line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#') && !line.StartsWith("##"))
            {
                continue;
            }

            if (line.StartsWith("##"))
            {
                list.Lines.Add(new Line(line[2..].Trim(), null, 0));
                continue;
            }

            if (line.StartsWith('='))
            {
                list.Lines.Add(new Line(null, line, 0));
                continue;
            }

            if (line.StartsWith("{="))
            {
                int close = line.LastIndexOf('}');
                string rest = line[(close + 1)..].Trim();
                list.Lines.Add(new Line(null, line[1..close], rest.Length == 0 ? 1 : int.Parse(rest)));
                continue;
            }

            var words = Split(line);
            switch (words[0])
            {
                case "width" when words.Count == 3:
                    list.Widths.Add((words[1], ushort.Parse(words[2])));
                    break;
                case "external" when words.Count >= 2:
                    list.Externals.Add((words[1] == "!" ? "\0" : words[1], words.Skip(2).ToList()));
                    break;
                case "name" when words.Count >= 3:
                    byte flags = 0;
                    char key = '\0';
                    if (words.Count >= 4 && words[3] == "macro")
                    {
                        flags = ExcelDefinedName.ReferenceFlag | ExcelDefinedName.MacroFlag;
                        key = '\x80';   // marks a function macro below
                    }
                    else if (words.Count >= 5 && words[3] == "command")
                    {
                        flags = ExcelDefinedName.ReferenceFlag | ExcelDefinedName.MacroFlag;
                        key = words[4][0];
                    }

                    list.Names.Add((words[1], words[2], flags, key));
                    break;
                case "recalc" when words.Count == 2:
                    list.RecalculationState = ushort.Parse(words[1]);
                    break;
                case "cell" when words.Count >= 3:
                    list.Cells.Add((words[1], line[(line.IndexOf(words[1], 4, StringComparison.Ordinal) + words[1].Length)..].Trim()));
                    break;
                default:
                    throw new InvalidDataException($"{path} line {number}: not a formula, heading or setup line: {raw}");
            }
        }

        return list;
    }

    /// <summary>
    /// Creates a sheet with the list's setup: widths, external documents, names (in Excel's order) and cells.
    /// </summary>
    public ExcelDocument CreateDocument(bool worksheet, ExcelVersion version)
    {
        var document = worksheet ? ExcelDocument.CreateWorksheet() : ExcelDocument.CreateMacroSheet();
        foreach (var (column, width) in Widths)
        {
            document.State.SetColumnWidth(ExcelCellAddress.Parse(column + "1").Column, width);
        }

        foreach (var (name, names) in Externals)
        {
            var external = new ExcelExternalDocument(name);
            external.Names.AddRange(names);
            document.ExternalDocuments.Add(external);
        }

        // Tokens refer to names by position, so add them all, sort them as Excel does, then compile definitions.
        foreach (var (name, _, flags, key) in Names)
        {
            document.Names.Add(new ExcelDefinedName(name, [], flags)
            {
                // Excel's DEFINE.NAME writes kind 0x80 for a function macro and 0x40 for a command (spec §11).
                ShortcutKey = key == '\x80' ? (byte)0 : (byte)key,
                MacroKind = (flags & ExcelDefinedName.MacroFlag) == 0 ? (byte)0 : key == '\x80' ? ExcelDefinedName.FunctionKind : ExcelDefinedName.CommandKind,
            });
        }

        document.Names.Sort((a, b) => string.CompareOrdinal(a.Name.ToUpperInvariant(), b.Name.ToUpperInvariant()));
        foreach (var (name, definition, _, _) in Names)
        {
            document.Names.Single(n => n.Name == name).SetDefinition(definition, document, version);
        }

        foreach (var (address, text) in Cells)
        {
            var formula = ExcelFormula.Parse(text, document, version);
            formula.SetResult(0);
            document.Cells[address].SetFormula(formula);
        }

        if (RecalculationState is { } state)
        {
            document.State.RecalculationState = state;
        }

        return document;
    }

    // Splits on spaces, keeping "quoted words" together.
    private static List<string> Split(string line)
    {
        var words = new List<string>();
        var current = new StringBuilder();
        bool quoted = false;
        foreach (char c in line)
        {
            if (c == '"')
            {
                quoted = !quoted;
            }
            else if (c == ' ' && !quoted)
            {
                if (current.Length > 0)
                {
                    words.Add(current.ToString());
                    current.Clear();
                }
            }
            else
            {
                current.Append(c);
            }
        }

        if (current.Length > 0)
        {
            words.Add(current.ToString());
        }

        return words;
    }
}
