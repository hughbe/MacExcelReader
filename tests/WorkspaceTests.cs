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
        Assert.Equal(1, workspace.CalculationMode);
        Assert.True(workspace.FullPrecision);
        Assert.False(workspace.Iteration);
        Assert.Equal(0.001, workspace.MaxChange);
        var window = Assert.Single(workspace.Windows);
        Assert.Equal(ExcelWorkspaceWindow.WorksheetKind, window.Kind);
        Assert.Equal("Expenses", window.DocumentName);
        Assert.Equal("Excel Data:", window.Folder);
        Assert.Equal((short)73, window.Window.Top);
        Assert.Equal((short)503, window.Window.Right);
        Assert.Equal(new ExcelRange(0, 0, 0, 0), Assert.Single(window.Window.Selection));
        Assert.Empty(workspace.Block);

        Assert.Equal(bytes, workspace.ToArray());
    }

    [Fact]
    public void ChartWindows_KeepOnlyTheirRectangle()
    {
        var workspace = new ExcelWorkspace();
        workspace.Windows.Add(new ExcelWorkspaceWindow
        {
            Kind = ExcelWorkspaceWindow.ChartKind,
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
