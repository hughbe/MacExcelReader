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
            config.AddCommand<ExportCommand>("export")
                .WithDescription("Export the sheet as CSV (or TSV), as displayed values or formulas.")
                .WithExample("export", "Expenses", "-o", "expenses.csv");
            config.AddCommand<NamesCommand>("names")
                .WithDescription("List the defined names and their definitions.");
            config.AddCommand<DumpCommand>("dump")
                .WithDescription("Dump the whole file: header, settings, names, windows, print setup and cells.");
            config.AddCommand<ChartCommand>("chart")
                .WithDescription("Dump a chart (MCBN) file.");
            config.AddCommand<LayoutCommand>("layout")
                .WithDescription("Show where each part of the file lies: header, code segments and sections.")
                .WithExample("layout", "Expenses");
            config.AddCommand<JsonCommand>("json")
                .WithDescription("Write the whole document as JSON.")
                .WithExample("json", "Expenses", "-o", "expenses.json");
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

    internal static string Hex(ReadOnlySpan<byte> bytes) => bytes.IsEmpty ? "(none)" : Convert.ToHexString(bytes);

    internal static string YesNo(bool value) => value ? "yes" : "no";
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
        foreach (var (_, _, cell) in document.Cells.NonEmptyCells)
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
        table.AddRow("Code signature", $"0x{document.Signature:X4}{(document.Signature == ExcelDocument.Excel103Signature ? " (Excel 1.03)" : string.Empty)}");
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
                var cell = document.Cells.TryGet(r, c);
                if (cell is null || (!cell.HasValue && !(settings.All && cell.RawAttributes != 0x4000 || settings.All && (cell.NumberFormat != 0 || cell.Alignment != 0))))
                {
                    continue;
                }

                count++;
                table.AddRow(
                    new ExcelCellAddress(r, c).ToString(),
                    cell.Type.ToString(),
                    Program.Escape(ExcelValueFormatter.FormatCell(cell)),
                    Program.Escape(ExcelValueFormatter.GetBuiltInFormat(cell.NumberFormat) ?? $"user #{cell.NumberFormat}"),
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
            var table = new Table().AddColumn("Field").AddColumn("Value");
            if (settings.Name)
            {
                expression = ExcelFormulaParser.ParseNameDefinition(settings.Formula, document);
                table.AddRow("Decompiled", Program.Escape("=" + FormulaDecompiler.Decompile(expression, [], document, fixupsInPlace: true)));
            }
            else
            {
                var formula = ExcelFormula.Parse(settings.Formula, document);
                expression = formula.Expression.ToArray();
                table.AddRow("Decompiled", Program.Escape(formula.ToText(document)));
                table.AddRow("Fix-ups", $"{Program.Hex(formula.Fixups)}  words: {string.Join(", ", formula.GetFixupWords())}");
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
        var cell = document.Cells.TryGet(address.Row, address.Column);
        if (cell is null)
        {
            AnsiConsole.MarkupLine($"[yellow]{address} is outside the sheet ({document.Cells.RowCount} rows × {document.Cells.ColumnCount} columns) and is empty.[/]");
            return 0;
        }

        Describe(document, address, cell);
        return 0;
    }

    internal static void Describe(ExcelDocument document, ExcelCellAddress address, ExcelCell cell)
    {
        var table = new Table().Title($"[bold]{address}[/]").AddColumn("Field").AddColumn("Value");
        table.AddRow("Type", cell.Type.ToString());
        table.AddRow("Display", Program.Escape(ExcelValueFormatter.FormatCell(cell)));
        table.AddRow("Value", Program.Escape(cell.Value switch
        {
            null => "(empty)",
            string s => "\"" + s + "\"",
            double d => d.ToString("R", CultureInfo.InvariantCulture),
            var other => Convert.ToString(other, CultureInfo.InvariantCulture),
        }));
        if (cell.TryGetNumber(out double number) && ExcelValueFormatter.IsDateFormat(cell.NumberFormat))
        {
            table.AddRow("As date", ExcelValueFormatter.ToDateTime(number).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
        }

        table.AddRow("Number format", $"{cell.NumberFormat} ({Program.Escape(ExcelValueFormatter.GetBuiltInFormat(cell.NumberFormat) ?? "user-defined")})");
        table.AddRow("Alignment", cell.Alignment.ToString());
        table.AddRow("Style", Program.Escape(Style(cell)));
        table.AddRow("Protection", $"locked {Program.YesNo(cell.IsLocked)}, hidden {Program.YesNo(cell.IsHidden)}");
        table.AddRow("Attribute word", $"0x{cell.RawAttributes:X4}");

        if (cell.Formula is { } formula)
        {
            var result = formula.Result;
            table.AddRow("Formula", Program.Escape(formula.ToText(document)));
            table.AddRow("Result", $"{result.Kind}: {Program.Escape(Convert.ToString(result.Value, CultureInfo.InvariantCulture))}");
            table.AddRow("Uses array function", Program.YesNo(cell.UsesArrayFunction));
            table.AddRow("Stored row/column", $"row word 0x{formula.StoredRowWord:X4}, column {formula.StoredColumn}");
            table.AddRow("Reserved bytes", Program.Hex(formula.Reserved));
            table.AddRow("Expression", Program.Hex(formula.Expression));
            table.AddRow("Fix-ups", $"{Program.Hex(formula.Fixups)}  words: {string.Join(", ", formula.GetFixupWords())}");
        }

        AnsiConsole.Write(table);

        if (cell.Formula is { } f)
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
                return FormulaDecompiler.CellText((ushort)((operands[0] << 8) | operands[1]), operands[2]);
            case 0x25 or 0x2D or 0x32:
                return FormulaDecompiler.CellText((ushort)((operands[0] << 8) | operands[1]), operands[4]) + ":" +
                       FormulaDecompiler.CellText((ushort)((operands[2] << 8) | operands[3]), operands[5]);
            case 0x17:
                return "\"" + MacRoman.GetString(operands[1..]) + "\"";
            case 0x1E:
                return ((operands[0] << 8) | operands[1]).ToString(CultureInfo.InvariantCulture);
            case 0x1D:
                return operands[0] != 0 ? "TRUE" : "FALSE";
            case 0x1C:
                return ExcelFormulaResult.ErrorText(operands[0]);
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

                var cell = document.Cells.TryGet(r, c);
                string text = cell is null ? string.Empty
                    : settings.Formulas && cell.Formula is { } formula ? formula.ToText(document)
                    : settings.Raw ? Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? string.Empty
                    : ExcelValueFormatter.FormatCell(cell);
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
        header.AddRow("Code segments", string.Join(", ", document.CodeSegments.Select(s => $"{s.Length} bytes")));
        header.AddRow("Function count", document.FunctionCount.ToString(CultureInfo.InvariantCulture));
        header.AddRow("Extra header", Program.Hex(document.ExtraHeader));
        header.AddRow("Built-in formats", document.BuiltInFormatCount.ToString(CultureInfo.InvariantCulture));
        header.AddRow("User formats", document.NumberFormats.Count == 0 ? "(none)"
            : string.Join("; ", document.NumberFormats.Select(f => Program.Hex(f.Descriptor) + " " + Program.Hex(f.Body))));
        AnsiConsole.Write(header);

        var settingsTable = new Table().Title("[bold]Sheet settings[/]").AddColumn("Setting").AddColumn("Value");
        settingsTable.AddRow("Calculation mode", state.CalculationMode switch { 1 => "automatic", 0 => "automatic except tables", -1 => "manual", var m => m.ToString(CultureInfo.InvariantCulture) });
        settingsTable.AddRow("Precision", state.FullPrecision ? "full" : "as displayed");
        settingsTable.AddRow("Iteration", $"{Program.YesNo(state.Iteration)}, max {state.MaxIterations}, max change {state.MaxChange.ToString(CultureInfo.InvariantCulture)}");
        settingsTable.AddRow("Reference style word", state.ReferenceStyle.ToString(CultureInfo.InvariantCulture));
        settingsTable.AddRow("Recalculation state", state.RecalculationState.ToString(CultureInfo.InvariantCulture));
        settingsTable.AddRow("Protected", $"{Program.YesNo(state.IsProtected)}, password hash 0x{state.PasswordHash:X4}");
        settingsTable.AddRow("Standard font", $"font {state.StandardFont}, {state.StandardFontSize} pt");
        settingsTable.AddRow("Last non-hidden column", state.LastVisibleColumn.ToString(CultureInfo.InvariantCulture));
        var widths = Enumerable.Range(0, ExcelSheetState.ColumnCount)
            .Where(c => state.GetColumnWidth(c) != 75)
            .Select(c => $"{FormulaDecompiler.ColumnName(c)}={state.GetColumnWidth(c)}");
        settingsTable.AddRow("Column widths", "75 except " + (string.Join(", ", widths) is { Length: > 0 } w ? w : "(none)"));
        AnsiConsole.Write(settingsTable);

        if (document.HasExternalDocumentTable)
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
            table.AddRow("State bytes", Program.Hex(window.State));
            AnsiConsole.Write(table);
        }

        AnsiConsole.MarkupLine($"[bold]Page breaks[/]: columns {Format(document.ColumnPageBreaks.Select(b => FormulaDecompiler.ColumnName(b)))}; rows {Format(document.RowPageBreaks.Select(r => (r + 1).ToString(CultureInfo.InvariantCulture)))}");
        if (document.PrintSetup is { } print)
        {
            var table = new Table().Title("[bold]Print setup[/]").AddColumn("Field").AddColumn("Value");
            table.AddRow("Header", Program.Escape(print.Header ?? "(none)"));
            table.AddRow("Footer", Program.Escape(print.Footer ?? "(none)"));
            table.AddRow("Printer message", Program.Escape(print.PrinterMessage ?? "(none)"));
            table.AddRow("Margins (L, R, T, B)", string.Join(", ", Enumerable.Range(0, 4).Select(i => print.GetMargin(i).ToString(CultureInfo.InvariantCulture))));
            table.AddRow("Print headings / gridlines", $"{Program.YesNo(print.PrintHeadings)} / {Program.YesNo(print.PrintGridlines)}");
            AnsiConsole.Write(table);
        }
        else
        {
            AnsiConsole.MarkupLine("[bold]Print setup[/]: (none)");
        }

        var cells = new Table().Title("[bold]Cells[/]").AddColumn("Cell").AddColumn("Type").AddColumn("Display").AddColumn("Formula");
        foreach (var (row, column, cell) in document.Cells.NonEmptyCells)
        {
            cells.AddRow(new ExcelCellAddress(row, column).ToString(), cell.Type.ToString(),
                Program.Escape(ExcelValueFormatter.FormatCell(cell)), Program.Escape(cell.Formula?.ToText(document)));
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
            string kind = name.IsCommand ? $"command (Option-⌘-{(char)name.ShortcutKey})"
                : (name.Flags & ExcelDefinedName.MacroFlag) != 0 ? "macro"
                : name.Definition.Length == 0 ? "undefined" : "name";
            names.AddRow((i + 1).ToString(CultureInfo.InvariantCulture), Program.Escape(name.Name), Program.Escape(kind),
                Program.Escape(name.Definition.Length == 0 ? string.Empty : "=" + name.GetDefinitionText(document)));
        }

        AnsiConsole.Write(names);
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
            table.AddRow("Legend", chart.LegendPosition.ToString());
            table.AddRow("Category axis", Axis(chart.CategoryAxis));
            table.AddRow("Value axis", Axis(chart.ValueAxis));
            table.AddRow("Value scale", Scale(chart.ValueScale));
            if (chart.XScale is { } x)
            {
                table.AddRow("X scale", Scale(x));
            }

            table.AddRow("Category crossing", $"at category {chart.ValueAxisCrossesAtCategory}, between categories {Program.YesNo(chart.ValueAxisCrossesBetweenCategories)}, at maximum {Program.YesNo(chart.ValueAxisCrossesAtMaximumCategory)}, reverse {Program.YesNo(chart.CategoriesInReverseOrder)}, labels every {chart.CategoriesBetweenTickLabels}");
            table.AddRow("Point formats", chart.PointFormats.Count.ToString(CultureInfo.InvariantCulture));
            AnsiConsole.Write(table);

            if (chart.Texts.Count > 0)
            {
                var texts = new Table().Title("[bold]Text objects[/]").AddColumn("Kind").AddColumn("Text").AddColumn("Attached to").AddColumn("Options");
                foreach (var text in chart.Texts)
                {
                    texts.AddRow(text.Kind.ToString(), Program.Escape(text.Text ?? "(automatic)"),
                        text.Kind == ChartTextKind.DataLabel ? $"series {text.SeriesNumber}, point {(text.PointNumber == 255 ? "all" : text.PointNumber)}" : string.Empty,
                        $"{(text.IsVertical ? "vertical" : "horizontal")}, show key {Program.YesNo(text.ShowKey)}, show value {Program.YesNo(text.ShowValue)}");
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
                    ? i < series.FirstTexts.Count ? series.FirstTexts[i] : string.Empty
                    : i < series.FirstValues.Count ? series.FirstValues[i].ToString() : string.Empty;
                table.AddRow((i + 1).ToString(CultureInfo.InvariantCulture), Program.Escape(first),
                    i < series.SecondValues.Count ? series.SecondValues[i].ToString() : string.Empty);
            }

            table.Caption($"type {series.SeriesType}, references {Program.Escape(series.FirstReference)} / {Program.Escape(series.SecondReference)}{(series.SourceDocument.Length > 0 ? ", from " + Program.Escape(series.SourceDocument) : string.Empty)}");
            AnsiConsole.Write(table);
        }

        return 0;
    }

    private static string Axis(ChartAxis axis)
        => $"ticks major {axis.MajorTickMarks}, minor {axis.MinorTickMarks}; labels {axis.TickLabelPosition}; line {(axis.AxisLinePattern == 0 ? "hidden" : "shown")}; gridlines major {(axis.MajorGridlinePattern == 0 ? "off" : "on")}, minor {(axis.MinorGridlinePattern == 0 ? "off" : "on")}";

    private static string Scale(ChartScale scale)
        => $"min {(scale.AutoMinimum ? "auto" : scale.Minimum)}, max {(scale.AutoMaximum ? "auto" : scale.Maximum)}, major {(scale.AutoMajorUnit ? "auto" : scale.MajorUnit)}, minor {(scale.AutoMinorUnit ? "auto" : scale.MinorUnit)}, crosses {(scale.AutoCrossesAt ? "auto" : scale.CrossesAt)}{(scale.Logarithmic ? ", log" : string.Empty)}{(scale.ReverseOrder ? ", reversed" : string.Empty)}";
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

sealed class JsonCommand : Command<JsonSettings>
{
    public override int Execute(CommandContext context, JsonSettings settings, CancellationToken cancellationToken)
    {
        var options = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
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

    private static object SheetModel(ExcelDocument document, bool includeEmpty)
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

                var cell = document.Cells.TryGet(r, c)!;
                if (!cell.HasValue && cell.RawAttributes == 0x4000 && cell.NumberFormat == 0 && cell.Alignment == 0)
                {
                    continue;
                }

                cells.Add(new
                {
                    address = new ExcelCellAddress(r, c).ToString(),
                    type = cell.Type.ToString(),
                    value = cell.Value,
                    display = ExcelValueFormatter.FormatCell(cell),
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
            signature = $"0x{document.Signature:X4}",
            rows = document.Cells.RowCount,
            columns = document.Cells.ColumnCount,
            settings = new
            {
                calculationMode = state.CalculationMode,
                fullPrecision = state.FullPrecision,
                iteration = state.Iteration,
                maxIterations = state.MaxIterations,
                maxChange = state.MaxChange,
                isProtected = state.IsProtected,
                passwordHash = state.PasswordHash,
                standardFont = state.StandardFont,
                standardFontSize = state.StandardFontSize,
                columnWidths = Enumerable.Range(0, ExcelSheetState.ColumnCount).Where(i => state.GetColumnWidth(i) != 75)
                    .ToDictionary(i => FormulaDecompiler.ColumnName(i), i => (int)state.GetColumnWidth(i)),
            },
            externalDocuments = document.ExternalDocuments.Select(d => new { name = d.Name, names = d.Names }),
            names = document.Names.Select(n => new
            {
                name = n.Name,
                flags = n.Flags,
                command = n.IsCommand,
                shortcutKey = n.IsCommand ? ((char)n.ShortcutKey).ToString() : null,
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

    private static object ChartModel(ChartDocument document) => new
    {
        version = document.Version,
        charts = document.Charts.Select(c => new
        {
            type = c.Type.ToString(),
            bar = c.IsBar,
            stacked = c.IsStacked,
            percent = c.IsPercent,
            legend = c.LegendPosition.ToString(),
            valueScale = new { c.ValueScale.AutoMinimum, c.ValueScale.Minimum, c.ValueScale.AutoMaximum, c.ValueScale.Maximum, c.ValueScale.Logarithmic },
            texts = c.Texts.Select(t => new { kind = t.Kind.ToString(), text = t.Text, automatic = t.IsAutomaticText }),
        }),
        series = document.Series.Select(s => new
        {
            name = s.Name,
            type = s.SeriesType,
            firstReference = s.FirstReference,
            secondReference = s.SecondReference,
            categories = s.SeriesType == 3 ? s.FirstTexts.Cast<object>() : s.FirstValues.Select(v => (object)v.ToDouble()),
            values = s.SecondValues.Select(v => v.ToDouble()),
        }),
    };
}
