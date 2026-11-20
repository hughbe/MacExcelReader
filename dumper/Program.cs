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
                .WithDescription("Show a worksheet or macro sheet's settings, names and windows.");
            config.AddCommand<CellsCommand>("cells")
                .WithDescription("List every non-empty cell, with formulas decompiled.");
            config.AddCommand<ChartCommand>("chart")
                .WithDescription("Show the contents of a chart (MCBN) file.");
        });

        return app.Run(args);
    }
}

sealed class FileSettings : CommandSettings
{
    [CommandArgument(0, "<input>")]
    public required string Input { get; init; }
}

sealed class InfoCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var table = new Table().AddColumn("Field").AddColumn("Value");
        table.AddRow("Kind", document.Kind.ToString());
        table.AddRow("Signature", $"0x{document.Signature:X4}");
        table.AddRow("Size", $"{document.Cells.RowCount} rows x {document.Cells.ColumnCount} columns");
        table.AddRow("Calculation mode", document.State.CalculationMode.ToString());
        table.AddRow("Max iterations", document.State.MaxIterations.ToString());
        table.AddRow("Max change", document.State.MaxChange.ToString());
        table.AddRow("Protected", document.State.IsProtected.ToString());
        table.AddRow("Standard font", $"{document.State.StandardFont} / {document.State.StandardFontSize} pt");
        table.AddRow("External documents", string.Join(", ", document.ExternalDocuments.Select(d => Markup.Escape($"'{d.Name}'"))));
        AnsiConsole.Write(table);

        var names = new Table().AddColumn("Name").AddColumn("Flags").AddColumn("Definition");
        foreach (var name in document.Names)
        {
            names.AddRow(Markup.Escape(name.Name), $"0x{name.Flags:X2}", Markup.Escape(name.GetDefinitionText(document)));
        }

        AnsiConsole.Write(names);
        foreach (var window in document.Windows)
        {
            AnsiConsole.MarkupLine($"Window ({window.Top},{window.Left})-({window.Bottom},{window.Right}), active cell {FormulaDecompiler.ColumnName(window.ActiveColumn)}{window.ActiveRow + 1}");
        }

        return 0;
    }
}

sealed class CellsCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var document = ExcelDocument.Load(settings.Input);
        var table = new Table().AddColumn("Cell").AddColumn("Type").AddColumn("Value").AddColumn("Formula");
        foreach (var (row, column, cell) in document.Cells.NonEmptyCells)
        {
            table.AddRow(
                $"{FormulaDecompiler.ColumnName(column)}{row + 1}",
                cell.Type.ToString(),
                Markup.Escape(cell.Value?.ToString() ?? string.Empty),
                Markup.Escape(cell.Formula?.ToText(document) ?? string.Empty));
        }

        AnsiConsole.Write(table);
        return 0;
    }
}

sealed class ChartCommand : Command<FileSettings>
{
    public override int Execute(CommandContext context, FileSettings settings, CancellationToken cancellationToken)
    {
        var document = ChartDocument.Load(settings.Input);
        foreach (var chart in document.Charts)
        {
            AnsiConsole.MarkupLine($"[bold]{(chart == document.MainChart ? "Main" : "Overlay")} chart[/]: {chart.Type}{(chart.IsBar ? " (bar)" : string.Empty)}, stacked={chart.IsStacked}, legend={chart.LegendPosition}");
            var scale = chart.ValueScale;
            AnsiConsole.MarkupLine($"  Value axis: min={(scale.AutoMinimum ? "auto" : scale.Minimum)} max={(scale.AutoMaximum ? "auto" : scale.Maximum)} log={scale.Logarithmic}");
            foreach (var text in chart.Texts)
            {
                AnsiConsole.MarkupLine($"  Text {text.Kind}: {Markup.Escape(text.Text ?? "(automatic)")}");
            }
        }

        foreach (var series in document.Series)
        {
            var values = string.Join(", ", series.SecondValues.Select(v => v.ToString()));
            AnsiConsole.MarkupLine($"Series {Markup.Escape(series.Name)} [{Markup.Escape(series.FirstReference)}]: {Markup.Escape(values)}");
        }

        return 0;
    }
}
