using MacExcelReader.Charts;
using static MacExcelReader.Tests.TestData;

namespace MacExcelReader.Tests;

public class WorkspaceTests
{
    [Fact]
    public void ResumeExcel_ReadsSettingsAndWindowsAndRoundTrips()
    {
        // The "Resume Excel" file on the Excel 1.00 program disk: Expenses open from the Excel Data disk.
        var bytes = Sample("Resume Excel");
        var workspace = ExcelWorkspace.Read(bytes);

        Assert.Equal(100, workspace.MaxIterations);
        Assert.Equal(ExcelCalculationMode.Automatic, workspace.CalculationMode);
        Assert.True(workspace.FullPrecision);
        Assert.False(workspace.Iteration);
        Assert.Equal(0.001, workspace.MaxChange);
        var window = Assert.Single(workspace.Windows);
        Assert.Equal(ExcelDocumentKind.Worksheet, window.Kind);
        Assert.Equal("Expenses", window.DocumentName);
        Assert.Equal("Excel Data:", window.Folder);
        Assert.Equal((short)73, window.Window.Top);
        Assert.Equal((short)503, window.Window.Right);
        Assert.Equal(new ExcelRange(0, 0, 0, 0), Assert.Single(window.Window.Selection));
        Assert.Empty(workspace.PreferredChartFormat);
        Assert.False(workspace.Zoomed);
        Assert.Equal(0, workspace.PrinterType);
        Assert.Equal(5, workspace.PrinterBaud);   // 9600
        Assert.Equal(0, workspace.PrinterPort);

        Assert.Equal(bytes, workspace.ToArray());
    }

    [Fact]
    public void ResumeExcel2_HasZoomPrinterSetupAndPreferredChart()
    {
        // Saved by Excel 1.00 on Quit after FULL(TRUE), PRINTER.SETUP(2,2,3) and Chart > Set Preferred Format, with a
        // chart and two macro sheets open.
        var bytes = Sample("Resume Excel 2");
        var workspace = ExcelWorkspace.Read(bytes);

        Assert.True(workspace.Zoomed);
        Assert.Equal(1, workspace.PrinterType);   // TTY
        Assert.Equal(2, workspace.PrinterBaud);   // 1200
        Assert.Equal(1, workspace.PrinterPort);   // modem
        Assert.Equal(1640, workspace.PreferredChartFormat.Length);
        Assert.Equal([ExcelDocumentKind.Chart, ExcelDocumentKind.MacroSheet, ExcelDocumentKind.MacroSheet], workspace.Windows.Select(w => w.Kind));
        Assert.Equal("Ws:Clock saved", workspace.Windows[0].DocumentName);
        Assert.Equal(bytes, workspace.ToArray());
    }

    [Fact]
    public void ResumeExcel2_PreferredChartFormatIsTheChartDataBlock()
    {
        // The preferred format is CLOCK.CH's chart data block (spec Appendix B): its length word, the three titles at
        // their fixed offsets, and the chart table at 652 pointing at the main chart and then the overlay.
        var workspace = ExcelWorkspace.Read(Sample("Resume Excel 2"));
        var block = workspace.PreferredChartFormat;

        Assert.Equal(block.Length, (block[0] << 8) | block[1]);
        Assert.Equal(1, block[334 + 3] & 0x0F);
        Assert.Equal(2, block[394 + 3] & 0x0F);
        Assert.Equal(3, block[454 + 3] & 0x0F);
        Assert.Equal([666, 1058], workspace.PreferredChartOffsets);
        Assert.Equal([ChartType.Scatter, ChartType.Pie], workspace.PreferredChartTypes);

        // Set Preferred Format copies the charts and titles but not the series, so the series lists are empty.
        Assert.All(new[] { 658, 660, 662, 664 }, offset => Assert.Equal(0, (block[offset] << 8) | block[offset + 1]));
    }

    [Fact]
    public void ChartWindows_KeepOnlyTheirRectangle()
    {
        var workspace = new ExcelWorkspace();
        workspace.Windows.Add(new ExcelWorkspaceWindow
        {
            Kind = ExcelDocumentKind.Chart,
            DocumentName = "Chart1",
            Folder = "Disk:",
            Window = new ExcelWindow(40, 10, 300, 400, new byte[ExcelWindow.StateSize]),
        });

        var reread = ExcelWorkspace.Read(workspace.ToArray());
        var window = Assert.Single(reread.Windows);
        Assert.Equal("Chart1", window.DocumentName);
        Assert.Equal((short)300, window.Window.Bottom);
        Assert.Empty(window.Window.Selection);
    }

    [Fact]
    public void NonZeroFirstWord_IsRejected()
        => Assert.Throws<ExcelFormatException>(() => ExcelWorkspace.Read([0, 1, 0, 0]));
}
