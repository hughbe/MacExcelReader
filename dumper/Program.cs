using System.ComponentModel;
using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Spectre.Console;
using Spectre.Console.Cli;
using MacExcelReader;
using MacExcelReader.Charts;

public sealed class Program
{
    public static int Main(string[] args)
    {
        var app = new CommandApp();
        app.Configure(config =>
        {
            config.SetApplicationName("mac-excel-dumper");
            config.AddCommand<InfoCommand>("info")
                .WithDescription("Summarise a worksheet, macro sheet or chart.")
                .WithExample("info", "Expenses");
            config.AddCommand<CellsCommand>("cells")
                .WithDescription("List the non-empty cells, optionally limited to a range.")
                .WithExample("cells", "Expenses", "--range", "A10:D20");
            config.AddCommand<CellCommand>("cell")
                .WithDescription("Show everything about one cell, including its formula tokens.")
                .WithExample("cell", "REPLACE", "B27");
            config.AddCommand<ParseCommand>("parse")
                .WithDescription("Compile formula text to Excel's tokens, optionally resolving names against a file.")
                .WithExample("parse", "=IF(A1>0,SUM(A1:A3),-1)")
                .WithExample("parse", "=COLUMNS(!ADD)", "--file", "DB.FORM");
            config.AddCommand<SheetCommand>("sheet")
                .WithDescription("Build sheets from a formula list, for testing formulas in Excel.")
                .WithExample("sheet", "tests/Data/formulas-1.0.txt", "Formulas", "--split", "100");
            config.AddCommand<NewChartCommand>("newchart")
                .WithDescription("Build a chart file with the library, for testing charts in Excel")
                .WithExample("newchart", "Line", "out/Chart", "--series", "1,2,3", "--series", "3,1,2", "--title", "Sales");
            config.AddBranch("excel-test", branch =>
            {
                branch.SetDescription("Test files in a real Excel under an emulator (see docs/Testing.md).");
                branch.AddCommand<PackCommand>("pack")
                    .WithDescription("Put files on a floppy image with a Driver macro sheet that opens, probes and re-saves each.")
                    .WithExample("excel-test", "pack", "Test.dsk", "Formulas 01", "Chart1");
                branch.AddCommand<JunkCommand>("junk")
                    .WithDescription("Fill every field Excel is said to ignore or rebuild with junk, to test it in Excel.")
                    .WithExample("excel-test", "junk", "Expenses", "out/Expenses");
                branch.AddCommand<RunCommand>("run")
                    .WithDescription("Run Driver in Excel 1.00 under MAME, headless, optionally capturing windows, and write Excel's saves back to the disk.")
                    .WithExample("excel-test", "run", "Test.dsk", "--capture", "Test:Sheet saved");
                branch.AddCommand<ExperimentCommand>("experiment")
                    .WithDescription("Pack, run (Excel 1.00 in MAME) and check files in one go, over as many disks as they need, reporting what each command returned.")
                    .WithExample("excel-test", "experiment", "out/Sheet1", "out/Sheet2", "-o", "results");
                branch.AddCommand<DiskFilesCommand>("files")
                    .WithDescription("List the files on a floppy image (MFS or HFS), or copy them out with -o")
                    .WithExample("excel-test", "files", "Test.dsk", "-o", "out");
                branch.AddCommand<CheckCommand>("check")
                    .WithDescription("After Driver has run in Excel, compare Excel's results and re-saved copies with the originals.")
                    .WithExample("excel-test", "check", "Test.dsk");
            });
            config.AddCommand<ExportCommand>("export")
                .WithDescription("Export the sheet as CSV (or TSV), as displayed values or formulas.")
                .WithExample("export", "Expenses", "-o", "expenses.csv");
            config.AddCommand<NamesCommand>("names")
                .WithDescription("List the defined names and their definitions.");
            config.AddCommand<DumpCommand>("dump")
                .WithDescription("Dump the whole file: header, settings, names, windows, print setup and cells.");
            config.AddCommand<ChartCommand>("chart")
                .WithDescription("Dump a chart (MCBN) file.");
            config.AddCommand<WorkspaceCommand>("workspace")
                .WithDescription("Show a workspace (\"Resume Excel\", type XLIN) file: settings and the windows it reopens.")
                .WithExample("workspace", "Resume Excel");
            config.AddCommand<LayoutCommand>("layout")
                .WithDescription("Show where each part of the file lies: header, code segments and sections.")
                .WithExample("layout", "Expenses");
            config.AddCommand<JsonCommand>("json")
                .WithDescription("Write the whole document as JSON.")
                .WithExample("json", "Expenses", "-o", "expenses.json");
            config.AddCommand<SylkCommand>("sylk")
                .WithDescription("Export the sheet as SYLK, byte for byte as Excel 1.00's Save As does.")
                .WithExample("sylk", "Expenses", "-o", "Expenses.slk");
            config.AddCommand<RenderCommand>("render")
                .WithDescription("Draw a sheet's window as Excel 1.00 shows it, as a PNG, with the fonts of a System file.")
                .WithExample("render", "Expenses", "--system", "Excel-Data.img", "-o", "Expenses.png");
        });

        return app.Run(args);
    }

    internal static bool IsChart(string path)
    {
        using var stream = File.OpenRead(path);
        Span<byte> magic = stackalloc byte[2];
        return stream.Read(magic) == 2 && magic[0] == 0x01 && magic[1] == 0x8C;
    }

    internal static string Escape(string? text) => Markup.Escape(text ?? string.Empty);

    internal static string CalculationModeText(ExcelCalculationMode mode) => mode switch
    {
        ExcelCalculationMode.Automatic => "automatic",
        ExcelCalculationMode.AutomaticExceptTables => "automatic except tables",
        ExcelCalculationMode.Manual => "manual",
        _ => ((short)mode).ToString(CultureInfo.InvariantCulture),
    };

    internal static string Hex(ReadOnlySpan<byte> bytes) => bytes.IsEmpty ? "(none)" : Convert.ToHexString(bytes);

    internal static string YesNo(bool value) => value ? "yes" : "no";

    // A value as stored: numbers in full ("R"), text as it is, TRUE, FALSE and errors as Excel writes them.
    internal static string Raw(ExcelValue value)
        => value.TryGetNumber(out double number) ? number.ToString("R", CultureInfo.InvariantCulture) : value.ToString();

    // A value for the cell table: like Raw, with text quoted and empty shown.
    internal static string Describe(ExcelValue value) => value.Kind switch
    {
        ExcelValueKind.Empty => "(empty)",
        ExcelValueKind.Text => "\"" + value.Text + "\"",
        _ => Raw(value),
    };
}

class FileSettings : CommandSettings
{
    [CommandArgument(0, "<input>")]
    [Description("The file to read (the data fork of the Macintosh file).")]
    public required string Input { get; init; }

    public override ValidationResult Validate()
        => File.Exists(Input) ? ValidationResult.Success() : ValidationResult.Error($"File not found: {Input}");
}

sealed class InfoCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        if (Program.IsChart(settings.Input))
        {
            return ChartCommand.Dump(ChartDocument.Load(settings.Input));
        }

        var document = ExcelDocument.Load(settings.Input);
        int cells = 0, formulas = 0, text = 0, numbers = 0;
        foreach (var cell in document.Cells.NonEmptyCells)
        {
            cells++;
            switch (cell.Type)
            {
                case ExcelCellType.Text:
                    text++;
                    break;
                case ExcelCellType.Integer or ExcelCellType.Number:
                    numbers++;
                    break;
                case ExcelCellType.NumericFormula or ExcelCellType.Formula:
                    formulas++;
                    break;
            }
        }

        var table = new Table().Title("[bold]Summary[/]").AddColumn("Field").AddColumn("Value");
        table.AddRow("Kind", $"{document.Kind} (Finder type {document.FinderType})");
        table.AddRow("Code signature", $"0x{document.Signature:X4}{(document.Signature == ExcelDocument.CodeSignature ? " (Excel 1.03)" : string.Empty)}");
        table.AddRow("Grid", $"{document.Cells.RowCount} rows × {document.Cells.ColumnCount} columns");
        table.AddRow("Cells with values", $"{cells} ({numbers} numbers, {text} text, {formulas} formulas)");
        table.AddRow("Defined names", document.Names.Count.ToString(CultureInfo.InvariantCulture));
        table.AddRow("External documents", document.ExternalDocuments.Count.ToString(CultureInfo.InvariantCulture));
        table.AddRow("Protected", Program.YesNo(document.State.IsProtected));
        table.AddRow("Windows", document.Windows.Count.ToString(CultureInfo.InvariantCulture));
        table.AddRow("Print setup", Program.YesNo(document.PrintSetup is not null));
        AnsiConsole.Write(table);
        return 0;
    }
}

sealed class CellsSettings : FileSettings
{
    [CommandOption("-r|--range <RANGE>")]
    [Description("Only show cells in this range, e.g. A1:D20 or B5.")]
    public string? Range { get; init; }

    [CommandOption("--all")]
    [Description("Include empty cells that carry formatting.")]
    public bool All { get; init; }
}

sealed class CellsCommand : Command<CellsSettings>
{
    public override int Execute(CommandContext context, CellsSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var (first, last) = RangeParser.Parse(settings.Range, document);
        var table = new Table()
            .AddColumn("Cell").AddColumn("Type").AddColumn("Display").AddColumn("Format").AddColumn("Style").AddColumn("Formula");
        int count = 0;
        for (int r = first.Row; r <= last.Row; r++)
        {
            for (int c = first.Column; c <= last.Column; c++)
            {
                var cell = document.Cells.Find(r, c);
                if (cell is null || (!cell.HasValue && !(settings.All && cell.RawAttributes != 0x4000 || settings.All && (cell.NumberFormat != 0 || cell.Alignment != 0))))
                {
                    continue;
                }

                count++;
                table.AddRow(
                    new ExcelCellAddress(r, c).ToString(),
                    cell.Type.ToString(),
                    Program.Escape(ExcelValueFormatter.FormatCell(cell, document)),
                    Program.Escape(document.GetNumberFormat(cell.NumberFormat)?.ToText() ?? $"#{cell.NumberFormat} (missing)"),
                    Program.Escape(CellCommand.Style(cell)),
                    Program.Escape(cell.Formula?.ToText(document)));
            }
        }

        AnsiConsole.Write(table);
        AnsiConsole.MarkupLine($"[grey]{count} cell(s)[/]");
        return 0;
    }
}

sealed class CellSettings : FileSettings
{
    [CommandArgument(1, "<address>")]
    [Description("The cell, e.g. B19.")]
    public required string Address { get; init; }

    public override ValidationResult Validate()
    {
        var baseResult = base.Validate();
        if (!baseResult.Successful)
        {
            return baseResult;
        }

        return ExcelCellAddress.TryParse(Address, out _)
            ? ValidationResult.Success()
            : ValidationResult.Error($"'{Address}' is not a cell address such as B19.");
    }
}

sealed class ParseSettings : CommandSettings
{
    [CommandArgument(0, "<formula>")]
    [Description("The formula text, e.g. =SUM(A1:A5)*2. Quote it for the shell.")]
    public required string Formula { get; init; }

    [CommandOption("-f|--file <FILE>")]
    [Description("A worksheet or macro sheet whose defined names and external documents the formula uses.")]
    public string? File { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("The Excel release to compile for: 1.0 (1.00–1.03, the default), 1.04 or 1.5.")]
    public string Excel { get; init; } = "1.0";

    [CommandOption("--array")]
    [Description("Compile as an array formula (the shared expression).")]
    public bool Array { get; init; }

    [CommandOption("--name")]
    [Description("Compile as a name definition (indices stored inside the tokens).")]
    public bool Name { get; init; }
}

sealed class ParseCommand : Command<ParseSettings>
{
    public override int Execute(CommandContext context, ParseSettings settings, CancellationToken cancellationToken)
    {
        var document = settings.File is null ? null : ExcelDocument.Load(settings.File);
        try
        {
            byte[] expression;
            var version = settings.Excel switch
            {
                "1.0" or "1.00" or "1.03" => ExcelVersion.Excel100,
                "1.04" => ExcelVersion.Excel104,
                "1.5" => ExcelVersion.Excel150,
                _ => throw new ExcelFormulaSyntaxException($"Unknown Excel version '{settings.Excel}' (use 1.0, 1.04 or 1.5)", 0),
            };
            var table = new Table().AddColumn("Field").AddColumn("Value");
            if (settings.Name)
            {
                var name = new ExcelDefinedName("Parsed");
                name.SetDefinition(settings.Formula, document, version);
                expression = name.Definition.ToArray();
                table.AddRow("Decompiled", Program.Escape("=" + name.GetDefinitionText(document)));
                table.AddRow("Array data", Program.Escape(string.Join(" ", name.ArrayConstants.Select(a => Convert.ToHexString(a)))));
            }
            else if (settings.Array)
            {
                // Entered on a scratch copy of the sheet, whose anchor cell then holds the shared expression.
                var scratch = document is null ? ExcelDocument.CreateWorksheet() : ExcelDocument.Read(document.ToArray());
                scratch.SetArrayFormula(new ExcelRange(0, 0, 0, 0), settings.Formula, version);
                var anchor = scratch.Cells[0, 0].Formula!;
                expression = anchor.SharedExpression.ToArray();
                table.AddRow("Fix-ups", Program.Hex(anchor.SharedFixups));
            }
            else
            {
                var formula = ExcelFormula.Parse(settings.Formula, document, version);
                expression = formula.Expression.ToArray();
                table.AddRow("Decompiled", Program.Escape(formula.ToText(document)));
                table.AddRow("Fix-ups", $"{Program.Hex(formula.Fixups)}  words: {string.Join(", ", formula.FixupWords.ToArray())}");
            }

            table.AddRow("Expression", $"{Program.Hex(expression)} ({expression.Length} bytes)");
            AnsiConsole.Write(table);
            AnsiConsole.Write(CellCommand.TokenTable(expression));
            return 0;
        }
        catch (ExcelFormulaSyntaxException exception)
        {
            AnsiConsole.MarkupLine($"[red]{Program.Escape(exception.Message)}[/]");
            AnsiConsole.MarkupLine(Program.Escape(settings.Formula));
            AnsiConsole.MarkupLine(new string(' ', exception.Position) + "[red]^[/]");
            return 1;
        }
    }
}

sealed class CellCommand : Command<CellSettings>
{
    public override int Execute(CommandContext context, CellSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var address = ExcelCellAddress.Parse(settings.Address);
        var cell = document.Cells.Find(address.Row, address.Column);
        if (cell is null)
        {
            AnsiConsole.MarkupLine($"[yellow]{address} is outside the sheet ({document.Cells.RowCount} rows × {document.Cells.ColumnCount} columns) and is empty.[/]");
            return 0;
        }

        Describe(document, address, cell, settings.Input);
        return 0;
    }

    internal static void Describe(ExcelDocument document, ExcelCellAddress address, ExcelCell cell, string? path = null)
    {
        var table = new Table().Title($"[bold]{address}[/]").AddColumn("Field").AddColumn("Value");
        table.AddRow("Type", cell.Type.ToString());
        table.AddRow("Display", Program.Escape(ExcelValueFormatter.FormatCell(cell, document)));
        table.AddRow("In its column", Program.Escape(ExcelValueFormatter.FormatCell(cell, document, address.Column))
            + $" ({document.State.GetColumnWidth(address.Column)} px, about {document.State.GetColumnWidth(address.Column) / Math.Max(1, (int)(cell.Bold ? document.State.BoldDigitWidth : document.State.DigitWidth))} digits)");
        table.AddRow("Value", Program.Escape(Program.Describe(cell.Value)));
        if (cell.TryGetNumber(out double number) && ExcelValueFormatter.IsDateFormat(cell.NumberFormat, document))
        {
            table.AddRow("As date", ExcelValueFormatter.ToDateTime(number).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
        }

        table.AddRow("Number format", $"{cell.NumberFormat}: {Program.Escape(document.GetNumberFormat(cell.NumberFormat)?.ToText() ?? "(missing)")}{(cell.NumberFormat >= ExcelNumberFormat.BuiltInCount ? " (user-defined)" : string.Empty)}");
        table.AddRow("Alignment", cell.Alignment.ToString());
        table.AddRow("Style", Program.Escape(Style(cell)));
        table.AddRow("Protection", $"locked {Program.YesNo(cell.IsLocked)}, hidden {Program.YesNo(cell.IsHidden)}");
        table.AddRow("Attribute word", $"0x{cell.RawAttributes:X4}");

        ExcelFormula? decoded = null;
        try
        {
            decoded = cell.Formula;
        }
        catch (ExcelFormatException e)
        {
            // Show what the file holds, so a formula the library cannot read can still be looked at.
            table.AddRow("Formula", $"[red]cannot be read: {Program.Escape(e.Message)}[/]");
            if (path is not null && ExcelDocument.Load(path).Cells.TryGetRawData(address.Row, address.Column, out var record) && record.Length > 15)
            {
                int length = Math.Min(record.Length, 15 + record[14] + 32);
                table.AddRow("Record", Program.Hex(record[..length]) + "  (value 8, reserved 3, column, row word, length, expression, then fix-ups)");
                table.AddRow("Expression", Program.Hex(record.Slice(15, Math.Min(record[14], record.Length - 15))));
            }
        }

        if (decoded is { } formula)
        {
            var result = formula.Result;
            table.AddRow("Formula", Program.Escape(formula.ToText(document)));
            table.AddRow("Result", $"{result.Kind}: {Program.Escape(Program.Describe(result))}");
            table.AddRow("Uses array function", Program.YesNo(cell.UsesArrayFunction));
            table.AddRow("Stored row/column", $"row word 0x{formula.StoredRowWord:X4}, column {formula.StoredColumn}");
            table.AddRow("Reserved bytes", Program.Hex(formula.Reserved));
            table.AddRow("Expression", Program.Hex(formula.Expression));
            table.AddRow("Fix-ups", $"{Program.Hex(formula.Fixups)}  words: {string.Join(", ", formula.FixupWords.ToArray())}");
        }

        AnsiConsole.Write(table);

        if (decoded is { } f)
        {
            AnsiConsole.Write(TokenTable(f.Expression.ToArray()));
        }
    }

    internal static Table TokenTable(byte[] expression)
    {
        var tokens = new Table().Title("[bold]Tokens[/]").AddColumn("Offset").AddColumn("Byte").AddColumn("Token").AddColumn("Operands").AddColumn("Meaning");
        foreach (var token in FormulaTokens.Enumerate(expression))
        {
            var operands = expression.AsSpan(token.Offset + 1, token.Length - 1);
            tokens.AddRow(
                token.Offset.ToString(CultureInfo.InvariantCulture),
                $"0x{token.RawCode:X2}",
                FormulaTokens.GetName(token.Code),
                Program.Hex(operands),
                Program.Escape(TokenMeaning(token, operands)));
        }

        return tokens;
    }

    private static string TokenMeaning(FormulaToken token, ReadOnlySpan<byte> operands)
    {
        switch (token.Code)
        {
            case 0x21:
                return ExcelFunctions.Get(operands[0])?.Name ?? $"function {operands[0]}";
            case 0x39:
                return ExcelFunctions.Get(operands[0] + 256)?.Name ?? $"function {operands[0] + 256}";
            case 0x3A:
                return $"{ExcelFunctions.Get(operands[1] + 256)?.Name ?? $"function {operands[1] + 256}"}, {operands[0] & 0x7F} argument(s)";
            case 0x22:
            {
                int count = operands[0] & 0x7F;
                string dialog = (operands[0] & 0x80) != 0 ? "?" : string.Empty;
                return operands[1] == 255
                    ? $"call user macro, {count - 1} argument(s)"
                    : $"{ExcelFunctions.Get(operands[1])?.Name ?? $"function {operands[1]}"}{dialog}, {count} argument(s)";
            }

            case 0x1F:
                return System.Buffers.Binary.BinaryPrimitives.ReadDoubleBigEndian(operands).ToString("R", CultureInfo.InvariantCulture);
            case 0x29:
                return $"function result range, {operands[0]} bytes";
            case 0x2A or 0x2B:
                return "#REF! (deleted)";
            case 0x24 or 0x2C or 0x31 or 0x37:
                return FormulaTokens.ReferenceText((ushort)((operands[0] << 8) | operands[1]), operands[2]);
            case 0x25 or 0x2D or 0x32:
                return FormulaTokens.ReferenceText((ushort)((operands[0] << 8) | operands[1]), operands[4]) + ":" +
                       FormulaTokens.ReferenceText((ushort)((operands[2] << 8) | operands[3]), operands[5]);
            case 0x17:
                return "\"" + MacRoman.GetString(operands[1..]) + "\"";
            case 0x1E:
                return ((operands[0] << 8) | operands[1]).ToString(CultureInfo.InvariantCulture);
            case 0x1D:
                return operands[0] != 0 ? "TRUE" : "FALSE";
            case 0x1C:
                return ((ExcelError)operands[0]).ToText();
            case 0x19:
                return operands[0] == 0x01 ? "volatile" : (operands[0] & 0x02) != 0 ? "IF jump" : (operands[0] & 0x08) != 0 ? "skip" : (operands[0] & 0x04) != 0 ? "CHOOSE" : $"flags 0x{operands[0]:X2}";
            default:
                return string.Empty;
        }
    }

    internal static string Style(ExcelCell cell)
    {
        var parts = new List<string>();
        if (cell.Bold) parts.Add("bold");
        if (cell.Italic) parts.Add("italic");
        var borders = new StringBuilder();
        if (cell.LeftBorder) borders.Append('L');
        if (cell.RightBorder) borders.Append('R');
        if (cell.TopBorder) borders.Append('T');
        if (cell.BottomBorder) borders.Append('B');
        if (borders.Length > 0) parts.Add("border " + borders);
        if (cell.Alignment != ExcelAlignment.General) parts.Add(cell.Alignment.ToString().ToLowerInvariant());
        return string.Join(", ", parts);
    }
}

sealed class ExportSettings : FileSettings
{
    [CommandOption("-o|--output <FILE>")]
    [Description("Write to this file instead of the console.")]
    public string? Output { get; init; }

    [CommandOption("--formulas")]
    [Description("Export formulas instead of their displayed values.")]
    public bool Formulas { get; init; }

    [CommandOption("--raw")]
    [Description("Export unformatted values instead of displayed text.")]
    public bool Raw { get; init; }

    [CommandOption("--tsv")]
    [Description("Separate columns with tabs.")]
    public bool Tsv { get; init; }

    [CommandOption("-r|--range <RANGE>")]
    [Description("Only export this range.")]
    public string? Range { get; init; }
}

sealed class ExportCommand : Command<ExportSettings>
{
    public override int Execute(CommandContext context, ExportSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var (first, last) = RangeParser.Parse(settings.Range, document);
        char separator = settings.Tsv ? '\t' : ',';
        var output = new StringBuilder();
        for (int r = first.Row; r <= last.Row; r++)
        {
            for (int c = first.Column; c <= last.Column; c++)
            {
                if (c > first.Column)
                {
                    output.Append(separator);
                }

                var cell = document.Cells.Find(r, c);
                string text = cell is null ? string.Empty
                    : settings.Formulas && cell.Formula is { } formula ? formula.ToText(document)
                    : settings.Raw ? Program.Raw(cell.Value)
                    : ExcelValueFormatter.FormatCell(cell, document);
                output.Append(Quote(text, separator));
            }

            output.Append('\n');
        }

        if (settings.Output is null)
        {
            Console.Write(output);
        }
        else
        {
            File.WriteAllText(settings.Output, output.ToString());
            AnsiConsole.MarkupLine($"Wrote {last.Row - first.Row + 1} row(s) to {Program.Escape(settings.Output)}");
        }

        return 0;
    }

    private static string Quote(string text, char separator)
        => text.IndexOfAny([separator, '"', '\n', '\r']) >= 0 ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
}

sealed class NamesCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        DumpCommand.WriteNames(document);
        return 0;
    }
}

sealed class DumpCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        if (Program.IsChart(settings.Input))
        {
            return ChartCommand.Dump(ChartDocument.Load(settings.Input));
        }

        var document = ExcelDocument.Load(settings.Input);
        var state = document.State;

        var header = new Table().Title("[bold]Header[/]").AddColumn("Field").AddColumn("Value");
        header.AddRow("Kind", $"{document.Kind} (Finder type {document.FinderType})");
        header.AddRow("Code signature", $"0x{document.Signature:X4}");
        header.AddRow("Code segments", $"{document.CodeSegment1.Length} bytes, {document.CodeSegment2.Length} bytes");
        header.AddRow("Function count", document.FunctionCount.ToString(CultureInfo.InvariantCulture));
        header.AddRow("Extra header", Program.Hex(document.ExtraHeader));
        if (document.ExtraFunctionNames.Count > 0)
        {
            header.AddRow("Extra function names", Program.Escape(string.Join(", ", document.ExtraFunctionNames.Select((n, i) => $"{ExcelFunctions.Count + i} {n}"))));
        }

        header.AddRow("Built-in formats", document.BuiltInFormatCount.ToString(CultureInfo.InvariantCulture));
        header.AddRow("User formats", document.NumberFormats.Count == 0 ? "(none)"
            : string.Join("; ", document.NumberFormats.Select((f, i) => $"{ExcelNumberFormat.BuiltInCount + i} {f.ToText()}").Select(Program.Escape)));
        AnsiConsole.Write(header);

        var settingsTable = new Table().Title("[bold]Sheet settings[/]").AddColumn("Setting").AddColumn("Value");
        settingsTable.AddRow("Calculation mode", Program.CalculationModeText(state.CalculationMode));
        settingsTable.AddRow("Precision", state.FullPrecision ? "full" : "as displayed");
        settingsTable.AddRow("Iteration", $"{Program.YesNo(state.Iteration)}, max {state.MaxIterations}, max change {state.MaxChange.ToString(CultureInfo.InvariantCulture)}");
        settingsTable.AddRow("Reference style", state.ReferenceStyle.ToString());
        settingsTable.AddRow("Recalculation state", $"{(ushort)state.RecalculationState} ({state.RecalculationState switch
        {
            ExcelRecalculationState.Clean => "clean",
            ExcelRecalculationState.Needed => "needs recalculation",
            _ => "pending",
        }})");
        settingsTable.AddRow("Uses later functions", Program.YesNo(state.UsesLaterFunctions));
        settingsTable.AddRow("Defined names with a definition", state.DefinedNameCount.ToString(CultureInfo.InvariantCulture));
        settingsTable.AddRow("Protected", $"{Program.YesNo(state.IsProtected)}, password hash 0x{state.PasswordHash:X4}");
        settingsTable.AddRow("Standard font", $"font {state.StandardFont}, {state.StandardFontSize} pt");
        settingsTable.AddRow("Last non-hidden column", state.LastVisibleColumn.ToString(CultureInfo.InvariantCulture));
        var widths = Enumerable.Range(0, ExcelCellGrid.MaxColumns)
            .Where(c => state.GetColumnWidth(c) != 75)
            .Select(c => $"{ExcelCellAddress.ColumnName(c)}={state.GetColumnWidth(c)}");
        settingsTable.AddRow("Column widths", "75 except " + (string.Join(", ", widths) is { Length: > 0 } w ? w : "(none)"));
        AnsiConsole.Write(settingsTable);

        if (document.HasExternalDocumentTable && document.ExternalDocuments.Count == 0)
        {
            AnsiConsole.WriteLine("External documents: an empty table");
        }
        else if (document.HasExternalDocumentTable)
        {
            var external = new Table().Title("[bold]External documents[/]").AddColumn("#").AddColumn("Name").AddColumn("Names used");
            for (int i = 0; i < document.ExternalDocuments.Count; i++)
            {
                var entry = document.ExternalDocuments[i];
                external.AddRow(
                    (i + 1).ToString(CultureInfo.InvariantCulture),
                    Program.Escape(entry.Name.Replace("\0", "(active sheet)")),
                    Program.Escape(string.Join(", ", entry.Names)));
            }

            AnsiConsole.Write(external);
        }

        WriteNames(document);

        foreach (var window in document.Windows)
        {
            var table = new Table().Title("[bold]Window[/]").AddColumn("Field").AddColumn("Value");
            table.AddRow("Rectangle", $"top {window.Top}, left {window.Left}, bottom {window.Bottom}, right {window.Right}");
            table.AddRow("Active cell", new ExcelCellAddress(window.ActiveRow, window.ActiveColumn).ToString());
            table.AddRow("Selection", string.Join(", ", window.Selection.Select(s =>
                new ExcelCellAddress(s.FirstRow, s.FirstColumn) + (s.FirstRow == s.LastRow && s.FirstColumn == s.LastColumn ? string.Empty : ":" + new ExcelCellAddress(s.LastRow, s.LastColumn)))));
            table.AddRow("Display", $"formulas {Program.YesNo(window.DisplayFormulas)}, gridlines {Program.YesNo(window.DisplayGridlines)}, headings {Program.YesNo(window.DisplayHeadings)}");
            table.AddRow("Splits", $"horizontal {window.HorizontalSplit}, vertical {window.VerticalSplit}");
            if (window.HasVersion18State)
            {
                table.AddRow("Frozen panes", window.PanesFrozen ? $"yes, at {new ExcelCellAddress(window.FrozenRow!.Value, window.FrozenColumn!.Value)}" : "no");
                table.AddRow("Gridline colour", window.GridlineColor?.ToString().ToLowerInvariant() ?? "automatic");
            }

            table.AddRow("State bytes", Program.Hex(window.State));
            AnsiConsole.Write(table);
        }

        AnsiConsole.MarkupLine($"[bold]Page breaks[/]: columns {Format(document.ColumnPageBreaks.Select(b => ExcelCellAddress.ColumnName(b)))}; rows {Format(document.RowPageBreaks.Select(r => (r + 1).ToString(CultureInfo.InvariantCulture)))}");
        if (document.PrintSetup is { } print)
        {
            var table = new Table().Title("[bold]Print setup[/]").AddColumn("Field").AddColumn("Value");
            table.AddRow("Header", Program.Escape(print.Header ?? "(none)"));
            table.AddRow("Footer", Program.Escape(print.Footer ?? "(none)"));
            table.AddRow("Printer message", Program.Escape(print.PrinterMessage ?? "(none)"));
            table.AddRow("Margins (inches)", string.Create(CultureInfo.InvariantCulture, $"left {print.LeftMargin}, right {print.RightMargin}, top {print.TopMargin}, bottom {print.BottomMargin}"));
            table.AddRow("Print headings / gridlines", $"{Program.YesNo(print.PrintHeadings)} / {Program.YesNo(print.PrintGridlines)}");
            AnsiConsole.Write(table);
        }
        else
        {
            AnsiConsole.MarkupLine("[bold]Print setup[/]: (none)");
        }

        var cells = new Table().Title("[bold]Cells[/]").AddColumn("Cell").AddColumn("Type").AddColumn("Display").AddColumn("Formula");
        foreach (var entry in document.Cells.NonEmptyCells)
        {
            int row = entry.Row;
            int column = entry.Column;
            var cell = entry.GetCell();
            cells.AddRow(new ExcelCellAddress(row, column).ToString(), cell.Type.ToString(),
                Program.Escape(ExcelValueFormatter.FormatCell(cell, document)), Program.Escape(cell.Formula?.ToText(document)));
        }

        AnsiConsole.Write(cells);
        return 0;
    }

    private static string Format(IEnumerable<string> values) => string.Join(", ", values) is { Length: > 0 } text ? text : "(none)";

    internal static void WriteNames(ExcelDocument document)
    {
        var names = new Table().Title("[bold]Defined names[/]").AddColumn("#").AddColumn("Name").AddColumn("Kind").AddColumn("Definition");
        for (int i = 0; i < document.Names.Count; i++)
        {
            var name = document.Names[i];
            string kind = name.MacroKind switch
            {
                ExcelMacroKind.Command => name.ShortcutKey is { } key ? $"command (Option-⌘-{key})" : "command",
                ExcelMacroKind.Function => "function macro",
                _ => name.Definition.Length == 0 ? "undefined"
                    : (name.Flags & ExcelNameFlags.Reference) != 0 ? "reference" : "formula",
            };
            names.AddRow((i + 1).ToString(CultureInfo.InvariantCulture), Program.Escape(name.Name), Program.Escape(kind),
                Program.Escape(name.Definition.Length == 0 ? string.Empty : "=" + name.GetDefinitionText(document)));
        }

        AnsiConsole.Write(names);
        if (!document.NamesAreSorted)
        {
            AnsiConsole.MarkupLine("[yellow]The names are not in Excel's order (spec §11): Excel gives #NAME? for some of them typed in a formula.[/]");
        }
    }
}

sealed class ChartCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
        => Dump(ChartDocument.Load(settings.Input));

    internal static int Dump(ChartDocument document)
    {
        AnsiConsole.MarkupLine($"[bold]Chart file[/] version {document.Version}, signature 0x{document.Signature:X4}, {document.Charts.Count} chart(s), {document.Series.Count} series, page setup {Program.YesNo(document.PageSetup is not null)}");
        foreach (var chart in document.Charts)
        {
            var table = new Table().Title($"[bold]{(chart == document.MainChart ? "Main" : "Overlay")} chart[/]").AddColumn("Field").AddColumn("Value");
            table.AddRow("Type", chart.Type + (chart.Type == ChartType.ColumnOrBar ? chart.IsBar ? " (bar)" : " (column)" : string.Empty));
            table.AddRow("Options", $"stacked {Program.YesNo(chart.IsStacked)}, 100% {Program.YesNo(chart.IsPercent)}, vary by category {Program.YesNo(chart.VaryByCategory)}");
            table.AddRow("Overlap / cluster spacing", $"{chart.Overlap}% / {chart.ClusterSpacing}%");
            table.AddRow("First pie slice angle", $"{chart.FirstSliceAngle}°");
            table.AddRow("Legend", chart.HasLegend ? chart.LegendPosition.ToString() : $"none (position {chart.LegendPosition} if added)");
            table.AddRow("Category axis", Axis(chart.CategoryAxis));
            table.AddRow("Value axis", Axis(chart.ValueAxis));
            table.AddRow("Value scale", Scale(chart.ValueScale));
            if (chart.XScale is { } x)
            {
                table.AddRow("X scale", Scale(x));
            }

            table.AddRow("Category crossing", $"at category {chart.ValueAxisCrossesAtCategory}, between categories {Program.YesNo(chart.ValueAxisCrossesBetweenCategories)}, at maximum {Program.YesNo(chart.ValueAxisCrossesAtMaximumCategory)}, reverse {Program.YesNo(chart.CategoriesInReverseOrder)}, labels every {chart.CategoriesBetweenTickLabels}");
            table.AddRow("Categories", chart.CategoryCount == 0 && chart.Excel15CategoryCount != 0
                ? $"{chart.Excel15CategoryCount} (Excel 1.5's word)"
                : chart.CategoryCount.ToString(CultureInfo.InvariantCulture));
            if (chart != document.MainChart)
            {
                table.AddRow("Shares main chart's scale", $"value axis {Program.YesNo(chart.SharesValueAxisScale)}, X axis {Program.YesNo(chart.SharesXAxisScale)}");
                table.AddRow("First series", chart.FirstOverlaySeries.ToString(CultureInfo.InvariantCulture));
            }

            var (top, left) = chart.DrawingMargins;
            table.AddRow("Layout (twips)", $"margins top {top}, left {left}; chart area {Box(chart.ChartArea.Box)}; plot area {Box(chart.PlotArea.Box)}; tick length {chart.ValueAxis.TickLength}");
            table.AddRow("Default series format", PointFormat(chart.DefaultSeriesFormat));
            for (int i = 0; i < chart.PointFormats.Count; i++)
            {
                table.AddRow($"Point format {i + 1}", PointFormat(chart.GetPointFormat(i)));
            }
            AnsiConsole.Write(table);

            if (chart.Texts.Count > 0)
            {
                var texts = new Table().Title("[bold]Text objects[/]").AddColumn("Kind").AddColumn("Text").AddColumn("Attached to").AddColumn("Options");
                foreach (var text in chart.Texts)
                {
                    string kind = text.IsArrow ? "Arrow" : text.Kind switch
                    {
                        ChartTextKind.ChartTitle => "Chart title",
                        ChartTextKind.CategoryAxisTitle => "Category-axis title",
                        ChartTextKind.ValueAxisTitle => "Value-axis title",
                        ChartTextKind.DataLabel => "Data label",
                        ChartTextKind.FreeText => "Free text",
                        _ => text.Kind.ToString(),
                    };
                    texts.AddRow(kind, Program.Escape(text.IsArrow ? string.Empty : text.Text ?? "(automatic)"),
                        text.Kind == ChartTextKind.DataLabel ? $"series {text.SeriesNumber}, point {(text.PointNumber == 255 ? "all" : text.PointNumber)}" : string.Empty,
                        text.IsArrow
                            ? $"pattern {text.ArrowLine.Pattern}, shaft {text.ArrowLine.Weight}, "
                              + (text.HeadStyle == ArrowHeadStyle.None ? "no head" : $"{text.HeadStyle.ToString().ToLowerInvariant()} head, {text.HeadWidth.ToString().ToLowerInvariant()} width, {text.HeadLength.ToString().ToLowerInvariant()} length")
                            : $"{(text.IsVertical ? "vertical" : "horizontal")}, aligned {text.HorizontalAlignment.ToString().ToLowerInvariant()} {text.VerticalAlignment.ToString().ToLowerInvariant()}, show key {Program.YesNo(text.ShowKey)}, show value {Program.YesNo(text.ShowValue)}");
                }

                AnsiConsole.Write(texts);
            }
        }

        foreach (var series in document.Series)
        {
            var table = new Table().Title($"[bold]Series {Program.Escape(series.Name)}[/]").AddColumn("#").AddColumn("Category / X").AddColumn("Value");
            for (int i = 0; i < series.PointCount; i++)
            {
                string first = series.SeriesType == 3
                    ? i < series.CategoryTexts.Count ? series.CategoryTexts[i] : string.Empty
                    : i < series.CategoryValues.Count ? series.CategoryValues[i].ToString() : string.Empty;
                table.AddRow((i + 1).ToString(CultureInfo.InvariantCulture), Program.Escape(first),
                    i < series.Values.Count ? series.Values[i].ToString() : string.Empty);
            }

            table.Caption($"{series.SeriesType switch { 0 => "dates", 3 => "text categories", _ => "numbers" }}, references {Program.Escape(series.CategoriesReference)} / {Program.Escape(series.ValuesReference)}{(series.SourceDocument.Length > 0 ? ", from " + Program.Escape(series.SourceDocument) : string.Empty)}");
            AnsiConsole.Write(table);
        }

        return 0;
    }

    private static string Box(ChartBox box) => $"top {box.Top}, left {box.Left}, {box.Height} × {box.Width}";

    private static string PointFormat(ChartPointFormat format)
        => $"series or point {format.SeriesOrPoint}; line {format.Line.Pattern}{(format.AutomaticLine ? " (auto)" : string.Empty)} {format.Line.Weight}; area pattern {format.AreaPattern}{(format.AutomaticArea ? " (auto)" : string.Empty)}; marker {format.Marker}{(format.AutomaticMarker ? " (auto)" : string.Empty)}"
            + (format.SeriesLabel ? "; series label" : string.Empty) + (format.PointLabels ? "; point labels" : string.Empty) + (format.PercentLabels ? " (percentages)" : string.Empty)
            + (format.Explosion != 0 ? $"; pulled out {format.Explosion}%" : string.Empty);

    private static string Axis(ChartAxis axis)
        => $"ticks major {axis.MajorTickMarks}, minor {axis.MinorTickMarks}; labels {axis.TickLabelPosition} (drawn at {axis.TickLabelsAt}); line {(axis.Line.Pattern == ChartPattern.None ? "hidden" : $"{axis.Line.Pattern} {axis.Line.Weight} at {axis.LineAt}")}; gridlines major {axis.MajorGridlines.Pattern}, minor {axis.MinorGridlines.Pattern}";

    // Automatic values show what the record holds too: what Excel last worked out when it drew the chart.
    private static string Scale(ChartScale scale)
        => $"min {Value(scale.AutomaticMinimum, scale.Minimum)}, max {Value(scale.AutomaticMaximum, scale.Maximum)}, major {Value(scale.AutomaticMajorUnit, scale.MajorUnit)}, minor {Value(scale.AutomaticMinorUnit, scale.MinorUnit)}, crosses {Value(scale.AutomaticCrossesAt, scale.CrossesAt)}{(scale.IsLogarithmic ? ", log" : string.Empty)}{(scale.ValuesInReverseOrder ? ", reversed" : string.Empty)}";

    private static string Value(bool automatic, double value) => automatic ? $"auto ({value.ToString("R", CultureInfo.InvariantCulture)})" : value.ToString("R", CultureInfo.InvariantCulture);
}

static class RangeParser
{
    public static (ExcelCellAddress First, ExcelCellAddress Last) Parse(string? range, ExcelDocument document)
    {
        var whole = (new ExcelCellAddress(0, 0), new ExcelCellAddress(Math.Max(0, document.Cells.RowCount - 1), Math.Max(0, document.Cells.ColumnCount - 1)));
        if (string.IsNullOrWhiteSpace(range))
        {
            return document.Cells.RowCount == 0 ? (new ExcelCellAddress(0, 0), new ExcelCellAddress(-1, -1)) : whole;
        }

        var parts = range.Split(':');
        var first = ExcelCellAddress.Parse(parts[0]);
        var last = parts.Length > 1 ? ExcelCellAddress.Parse(parts[1]) : first;
        return (new ExcelCellAddress(Math.Min(first.Row, last.Row), Math.Min(first.Column, last.Column)),
                new ExcelCellAddress(Math.Max(first.Row, last.Row), Math.Max(first.Column, last.Column)));
    }
}

sealed class LayoutCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var bytes = File.ReadAllBytes(settings.Input);
        var layout = ExcelFileLayout.Read(bytes);
        var table = new Table().Title($"[bold]Layout of {Program.Escape(Path.GetFileName(settings.Input))}[/] ({layout.FileLength} bytes)")
            .AddColumn("Part").AddColumn(new TableColumn("Offset").RightAligned()).AddColumn(new TableColumn("Length").RightAligned()).AddColumn("Notes");
        table.AddRow("Header", "0", "24", "magic 0x5327, version 17");
        for (int i = 0; i < layout.CodeSegments.Count; i++)
        {
            var (offset, length) = layout.CodeSegments[i];
            table.AddRow($"Code segment {i + 1}", Hex(offset), length.ToString(CultureInfo.InvariantCulture), "embedded p-code");
        }

        table.AddRow("Function count", Hex(layout.DataOffset), "2", string.Empty);
        table.AddRow("Sheet state", Hex(layout.StateOffset), ExcelSheetState.Size.ToString(CultureInfo.InvariantCulture), string.Empty);
        table.AddRow("Document kind", Hex(layout.KindOffset), "1", $"0x{bytes[layout.KindOffset]:X2}");
        table.AddRow("Section table", Hex(layout.SectionTableOffset), "32", "8 block/offset pairs");
        table.AddRow("Extra header", Hex(layout.ExtraHeaderOffset), (2 + layout.ExtraHeaderLength).ToString(CultureInfo.InvariantCulture), string.Empty);
        foreach (var section in Enum.GetValues<ExcelSection>())
        {
            table.AddRow($"Section {(int)section}: {section}", Hex(layout.GetSectionOffset(section)),
                layout.GetSectionLength(section).ToString(CultureInfo.InvariantCulture), string.Empty);
        }

        AnsiConsole.Write(table);
        return 0;
    }

    private static string Hex(int value) => $"0x{value:X}";
}

sealed class JsonSettings : FileSettings
{
    [CommandOption("-o|--output <FILE>")]
    [Description("Write to this file instead of the console.")]
    public string? Output { get; init; }

    [CommandOption("--empty")]
    [Description("Include empty cells that carry formatting.")]
    public bool Empty { get; init; }
}

sealed class SylkSettings : FileSettings
{
    [CommandOption("-o|--output <FILE>")]
    [Description("Write to this file instead of the console.")]
    public string? Output { get; init; }

    [CommandOption("--excel <VERSION>")]
    [Description("Match this release's export: 1.00 (the default, also 1.01–1.06) or 1.5.")]
    public string? Excel { get; init; }
}

sealed class SylkCommand : Command<SylkSettings>
{
    public override int Execute(CommandContext context, SylkSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var version = settings.Excel == "1.5" ? ExcelVersion.Excel150 : ExcelVersion.Excel100;
        if (settings.Output is not null)
        {
            document.ExportSylk(settings.Output, version);
            AnsiConsole.MarkupLine($"Wrote {Program.Escape(settings.Output)}");
            return 0;
        }

        // Mac OS Roman with carriage returns: shown as lines, with escapes as they are.
        var buffer = new System.Buffers.ArrayBufferWriter<byte>();
        document.ExportSylk(buffer, version);
        Console.Write(MacRoman.GetString(buffer.WrittenSpan).Replace('\r', '\n'));
        return 0;
    }
}

sealed class JsonCommand : Command<JsonSettings>
{
    // A cell value as JSON: a number, a string, true or false, an error's text such as "#N/A", or null.
    private static object? JsonValue(ExcelValue value) => value.Kind switch
    {
        ExcelValueKind.Number => value.Number,
        ExcelValueKind.Text => value.Text,
        ExcelValueKind.Boolean => value.Boolean,
        ExcelValueKind.Error => value.Error.ToText(),
        _ => null,
    };

    public override int Execute(CommandContext context, JsonSettings settings, CancellationToken cancellationToken)
    {
        var options = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowNamedFloatingPointLiterals };
        object model = Program.IsChart(settings.Input) ? ChartModel(ChartDocument.Load(settings.Input)) : SheetModel(ExcelDocument.Load(settings.Input), settings.Empty);
        string json = JsonSerializer.Serialize(model, options);
        if (settings.Output is null)
        {
            Console.WriteLine(json);
        }
        else
        {
            File.WriteAllText(settings.Output, json);
            AnsiConsole.MarkupLine($"Wrote {Program.Escape(settings.Output)}");
        }

        return 0;
    }

    internal static object SheetModel(ExcelDocument document, bool includeEmpty)
    {
        var state = document.State;
        var cells = new List<object>();
        for (int r = 0; r < document.Cells.RowCount; r++)
        {
            for (int c = 0; c < document.Cells.ColumnCount; c++)
            {
                if (document.Cells.GetCellType(r, c) == ExcelCellType.Empty && !includeEmpty)
                {
                    continue;
                }

                var cell = document.Cells.Find(r, c)!;
                if (!cell.HasValue && cell.RawAttributes == 0x4000 && cell.NumberFormat == 0 && cell.Alignment == 0)
                {
                    continue;
                }

                cells.Add(new
                {
                    address = new ExcelCellAddress(r, c).ToString(),
                    type = cell.Type.ToString(),
                    value = JsonValue(cell.Value),
                    display = ExcelValueFormatter.FormatCell(cell, document),
                    formula = cell.Formula?.ToText(document),
                    numberFormat = cell.NumberFormat,
                    alignment = cell.Alignment.ToString(),
                    bold = cell.Bold,
                    italic = cell.Italic,
                    borders = new { left = cell.LeftBorder, right = cell.RightBorder, top = cell.TopBorder, bottom = cell.BottomBorder },
                    locked = cell.IsLocked,
                    hidden = cell.IsHidden,
                });
            }
        }

        return new
        {
            kind = document.Kind.ToString(),
            version = document.Version,
            signature = $"0x{document.Signature:X4}",
            rows = document.Cells.RowCount,
            columns = document.Cells.ColumnCount,
            settings = new
            {
                calculationMode = state.CalculationMode.ToString(),
                fullPrecision = state.FullPrecision,
                iteration = state.Iteration,
                maxIterations = state.MaxIterations,
                maxChange = state.MaxChange,
                isProtected = state.IsProtected,
                passwordHash = state.PasswordHash,
                standardFont = state.StandardFont,
                standardFontSize = state.StandardFontSize,
                columnWidths = Enumerable.Range(0, ExcelCellGrid.MaxColumns).Where(i => state.GetColumnWidth(i) != 75)
                    .ToDictionary(i => ExcelCellAddress.ColumnName(i), i => (int)state.GetColumnWidth(i)),
            },
            numberFormats = document.NumberFormats.Select(f => new { text = f.ToText(), bytes = Convert.ToHexString(f.Descriptor) + " " + Convert.ToHexString(f.Body) }),
            externalDocuments = document.ExternalDocuments.Select(d => new { name = d.Name, names = d.Names }),
            names = document.Names.Select(n => new
            {
                name = n.Name,
                flags = (byte)n.Flags,
                command = n.MacroKind == ExcelMacroKind.Command,
                shortcutKey = n.MacroKind == ExcelMacroKind.Command ? (n.ShortcutKey ?? '\0').ToString() : null,
                definition = n.Definition.Length == 0 ? null : "=" + n.GetDefinitionText(document),
            }),
            windows = document.Windows.Select(w => new
            {
                rectangle = new { w.Top, w.Left, w.Bottom, w.Right },
                activeCell = new ExcelCellAddress(w.ActiveRow, w.ActiveColumn).ToString(),
                selection = w.Selection.Select(s => new ExcelCellAddress(s.FirstRow, s.FirstColumn) + ":" + new ExcelCellAddress(s.LastRow, s.LastColumn)),
                displayFormulas = w.DisplayFormulas,
                displayGridlines = w.DisplayGridlines,
                displayHeadings = w.DisplayHeadings,
            }),
            rowPageBreaks = document.RowPageBreaks,
            columnPageBreaks = document.ColumnPageBreaks,
            printSetup = document.PrintSetup is { } p ? new { p.Header, p.Footer, p.PrinterMessage, p.PrintHeadings, p.PrintGridlines } : null,
            cells,
        };
    }

    internal static object ChartModel(ChartDocument document) => new
    {
        version = document.Version,
        charts = document.Charts.Select(c => new
        {
            type = c.Type.ToString(),
            bar = c.IsBar,
            stacked = c.IsStacked,
            percent = c.IsPercent,
            legend = c.LegendPosition.ToString(),
            valueScale = new { c.ValueScale.AutomaticMinimum, c.ValueScale.Minimum, c.ValueScale.AutomaticMaximum, c.ValueScale.Maximum, c.ValueScale.IsLogarithmic },
            texts = c.Texts.Select(t => new { kind = t.Kind.ToString(), text = t.Text, automatic = t.IsAutomaticText }),
        }),
        series = document.Series.Select(s => new
        {
            name = s.Name,
            type = s.SeriesType,
            categoriesReference = s.CategoriesReference,
            valuesReference = s.ValuesReference,
            categories = s.SeriesType == 3 ? s.CategoryTexts.Cast<object>() : s.CategoryValues.Select(v => (object)v.ToDouble()),
            values = s.Values.Select(v => v.ToDouble()),
        }),
    };
}

sealed class WorkspaceCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var workspace = ExcelWorkspace.Load(settings.Input);
        var summary = new Table().Title("[bold]Workspace[/]").AddColumn("Field").AddColumn("Value");
        summary.AddRow("Calculation mode", Program.CalculationModeText(workspace.CalculationMode));
        summary.AddRow("Precision", workspace.FullPrecision ? "full" : "as displayed");
        summary.AddRow("Iteration", $"{Program.YesNo(workspace.Iteration)}, max {workspace.MaxIterations}, max change {workspace.MaxChange.ToString(CultureInfo.InvariantCulture)}");
        summary.AddRow("A window zoomed", Program.YesNo(workspace.Zoomed));
        summary.AddRow("Preferred chart format", workspace.PreferredChartFormat.Length == 0
            ? "(none)"
            : $"{workspace.PreferredChartFormat.Length:N0} bytes; "
              + string.Join(", ", workspace.PreferredChartOffsets.Zip(workspace.PreferredChartTypes, (o, t) => (o, t))
                  .Select((c, i) => $"{(i == 0 ? "main" : "overlay")} {c.t} chart at offset {c.o}")));
        string[] bauds = ["300", "600", "1200", "2400", "4800", "9600"];
        summary.AddRow("Printer setup", workspace.PrinterType == 0 ? "Macintosh"
            : $"TTY, {(workspace.PrinterPort == 0 ? "printer" : "modem")} port, {(workspace.PrinterBaud < bauds.Length ? bauds[workspace.PrinterBaud] : workspace.PrinterBaud.ToString(CultureInfo.InvariantCulture))} baud");
        AnsiConsole.Write(summary);

        var windows = new Table().Title("[bold]Windows, front to back[/]").AddColumn("Kind").AddColumn("Document").AddColumn("Folder").AddColumn("Directory").AddColumn("Rectangle").AddColumn("Active cell");
        foreach (var entry in workspace.Windows)
        {
            string kind = entry.Kind switch
            {
                ExcelDocumentKind.Worksheet => "worksheet",
                ExcelDocumentKind.MacroSheet => "macro sheet",
                ExcelDocumentKind.Chart => "chart",
                _ => $"0x{entry.Kind:X2}",
            };
            var w = entry.Window;
            windows.AddRow(kind, Program.Escape(entry.DocumentName), Program.Escape(entry.Folder), entry.Directory.ToString(CultureInfo.InvariantCulture),
                $"{w.Top}, {w.Left}, {w.Bottom}, {w.Right}",
                entry.Kind == ExcelDocumentKind.Chart ? "" : new ExcelCellAddress(w.ActiveRow, w.ActiveColumn).ToString());
        }

        AnsiConsole.Write(windows);
        return 0;
    }
}


sealed class NewChartSettings : CommandSettings
{
    [CommandArgument(0, "<type>")]
    [Description("ColumnOrBar, Line, Pie, Area or Scatter.")]
    public required ChartType Type { get; init; }

    [CommandArgument(1, "<output>")]
    [Description("The chart file to write.")]
    public required string Output { get; init; }

    [CommandOption("--series <VALUES>")]
    [Description("A series as comma-separated numbers; repeat for more series. Default: one series 1,3,2.")]
    public string[] Series { get; init; } = [];

    [CommandOption("--title <TEXT>")]
    [Description("A chart title.")]
    public string? Title { get; init; }

    [CommandOption("--bar")]
    [Description("Horizontal bars (with ColumnOrBar).")]
    public bool Bar { get; init; }

    [CommandOption("--scale <FIELD=VALUE>")]
    [Description("Fix a value-axis scale field instead of leaving it automatic: min, max, major, minor or crosses, as min=2. Repeatable.")]
    public string[] Scale { get; init; } = [];

    [CommandOption("--log")]
    [Description("A logarithmic value axis.")]
    public bool Log { get; init; }

    [CommandOption("--overlay <TYPE>")]
    [Description("Add an overlay chart of this type, plotting the series from --overlay-from on.")]
    public ChartType? Overlay { get; init; }

    [CommandOption("--overlay-from <N>")]
    [Description("The overlay's first series, counted from 1 (default: the last series).")]
    public int? OverlayFrom { get; init; }

    [CommandOption("--share")]
    [Description("Scale the overlay's value axis over the data of both charts.")]
    public bool Share { get; init; }

    [CommandOption("--share-x")]
    [Description("Scale a scatter overlay's X axis over the data of both charts.")]
    public bool ShareX { get; init; }

    [CommandOption("--hide-overlay-axes")]
    [Description("Hide the overlay's axes, as Excel does when it makes an overlay.")]
    public bool HideOverlayAxes { get; init; }

    [CommandOption("--hide-value-axis")]
    [Description("Hide the main chart's value axis.")]
    public bool HideValueAxis { get; init; }
}

/// <summary>
/// Builds a chart file with <see cref="ChartDocument.Create"/>, so the harness can check what Excel makes of it.
/// </summary>
sealed class NewChartCommand : Command<NewChartSettings>
{
    public override int Execute(CommandContext context, NewChartSettings settings, CancellationToken cancellationToken)
    {
        var document = ChartDocument.Create(settings.Type);
        document.MainChart.IsBar = settings.Bar;
        if (settings.Title is { } title)
        {
            document.MainChart.Texts.Add(ChartText.Create(ChartTextKind.ChartTitle, title));
        }

        var series = settings.Series.Length == 0 ? ["1,3,2"] : settings.Series;
        for (int i = 0; i < series.Length; i++)
        {
            var values = series[i].Split(',').Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToList();
            document.Series.Add(ChartSeries.Create($"Series {i + 1}", values, "Data", $"R{i + 1}C1:R{i + 1}C{values.Count}"));
        }

        var scale = document.MainChart.ValueScale;
        scale.IsLogarithmic = settings.Log;
        if (settings.HideValueAxis)
        {
            document.MainChart.ValueAxis.Line.Pattern = ChartPattern.None;
        }

        foreach (string field in settings.Scale)
        {
            var parts = field.Split('=', 2);
            double value = double.Parse(parts[1], CultureInfo.InvariantCulture);
            switch (parts[0])
            {
                case "min": scale.Minimum = value; scale.AutomaticMinimum = false; break;
                case "max": scale.Maximum = value; scale.AutomaticMaximum = false; break;
                case "major": scale.MajorUnit = value; scale.AutomaticMajorUnit = false; break;
                case "minor": scale.MinorUnit = value; scale.AutomaticMinorUnit = false; break;
                case "crosses": scale.CrossesAt = value; scale.AutomaticCrossesAt = false; break;
                default: throw new ArgumentException($"Unknown scale field {parts[0]}: use min, max, major, minor or crosses.");
            }
        }

        if (settings.Overlay is { } overlayType)
        {
            var overlay = Chart.Create(overlayType);
            overlay.FirstOverlaySeries = settings.OverlayFrom ?? series.Length;
            overlay.SharesValueAxisScale = settings.Share;
            overlay.SharesXAxisScale = settings.ShareX;
            if (settings.HideOverlayAxes)
            {
                overlay.CategoryAxis.Line.Pattern = overlay.ValueAxis.Line.Pattern = ChartPattern.None;
            }
            document.Charts.Insert(0, overlay);
        }

        File.WriteAllBytes(settings.Output, document.ToArray());
        AnsiConsole.MarkupLine($"Wrote [bold]{Markup.Escape(settings.Output)}[/]: {settings.Type}, {series.Length} series.");
        return 0;
    }
}
