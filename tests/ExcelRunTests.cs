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

    [Fact]
    public void Pack_RefusesCommandsThatAreNotUtf8()
    {
        // A .commands file saved as Latin-1 or Mac OS Roman would turn its accented letters into replacement
        // characters; pack says so instead.
        string folder = Directory.CreateTempSubdirectory("excel-pack-test-").FullName;
        try
        {
            string sheet = Path.Combine(folder, "Sheet");
            ExcelDocument.CreateWorksheet().Save(sheet);
            File.WriteAllBytes(sheet + ".commands", System.Text.Encoding.Latin1.GetBytes("=DEFINE.NAME(\"aéx\",1)\n"));
            Assert.NotEqual(0, Program.Main(["excel-test", "pack", Path.Combine(folder, "Test.dsk"), sheet]));
            Assert.False(File.Exists(Path.Combine(folder, "Test.dsk")));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public void Experiment_RunsFilesOverSeveralDisksAndReportsEveryCommand()
    {
        if (!CanRun)
        {
            return;
        }

        string folder = Directory.CreateTempSubdirectory("excel-experiment-test-").FullName;
        try
        {
            // Two sheets, one to a disk: a name with an accent, and a name Excel refuses, whose error must not stop
            // the file's other commands.
            string first = Path.Combine(folder, "First"), second = Path.Combine(folder, "Second");
            ExcelDocument.CreateWorksheet().Save(first);
            ExcelDocument.CreateWorksheet().Save(second);
            File.WriteAllText(first + ".commands", "=DEFINE.NAME(\"aéx\",1)\n");
            File.WriteAllText(second + ".commands", "=DEFINE.NAME(\"a∞x\",1)\n=DEFINE.NAME(\"ok\",2)\n");
            string output = Path.Combine(folder, "out");
            Program.Main(["excel-test", "experiment", first, second, "-o", output, "--batch", "1", "--no-probe"]);

            Assert.True(Directory.Exists(Path.Combine(output, "X1")) && Directory.Exists(Path.Combine(output, "X2")));
            Assert.Contains(ExcelDocument.Load(Path.Combine(output, "saved", "First saved")).Names, n => n.Name == "aéx");
            var secondNames = ExcelDocument.Load(Path.Combine(output, "saved", "Second saved")).Names;
            Assert.Contains(secondNames, n => n.Name == "ok");
            Assert.DoesNotContain(secondNames, n => n.Name == "a∞x");

            var results = File.ReadLines(Path.Combine(output, "commands.tsv")).Skip(1).Select(l => l.Split('\t')).ToList();
            Assert.Equal(["TRUE", "#VALUE!", "TRUE"], results.Select(r => r[4]));
            Assert.Equal(["First", "Second", "Second"], results.Select(r => r[1]));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }
}
