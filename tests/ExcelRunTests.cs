using Xunit;

namespace MacExcelReader.Tests;

/// <summary>
/// Runs the harness of docs/Testing.md end to end in a real Excel 1.00 under MAME: pack, run, capture and check.
/// </summary>
public class ExcelRunTests
{
    // Needs MAME on the PATH and the three files the run command reads: MACEXCEL_SYSTEM (Excel 1.00's Data disk),
    // MACEXCEL_PROGRAM (its Program disk) and MACEXCEL_ROM (the Mac Plus ROM). Passes without checking anything
    // otherwise.
    private static bool CanRun => new[] { "MACEXCEL_SYSTEM", "MACEXCEL_PROGRAM", "MACEXCEL_ROM" }
        .All(v => Environment.GetEnvironmentVariable(v) is { Length: > 0 } path && File.Exists(path));

    [Fact]
    public void Run_ReSavesAndCapturesAWindowAsExcelDrawsIt()
    {
        if (!CanRun)
        {
            return;
        }

        string folder = Directory.CreateTempSubdirectory("excel-run-test-").FullName;
        try
        {
            // Render Basics is Excel's save of R Basics, captured on volume Render.
            string sheet = Path.Combine(folder, "R Basics");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Samples", "Excel 1.00", "Render Basics"), sheet);
            string disk = Path.Combine(folder, "Test.dsk");
            Assert.Equal(0, Program.Main(["excel-test", "pack", disk, sheet, "--volume", "Render", "--keep-open"]));
            Assert.Equal(0, Program.Main(["excel-test", "run", disk, "--capture", "Render:R Basics saved=Captured.png", "-o", folder]));
            Assert.Equal(0, Program.Main(["excel-test", "check", disk]));

            var expected = RenderTests.ReadPng(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "Screens", "Render Basics.png")));
            var actual = RenderTests.ReadPng(File.ReadAllBytes(Path.Combine(folder, "Captured.png")));
            Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
            Assert.Equal(expected.Bits, actual.Bits);
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }
}
