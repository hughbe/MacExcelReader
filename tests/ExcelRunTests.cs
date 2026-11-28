using Xunit;

namespace MacExcelReader.Tests;

/// <summary>
/// Runs the harness of docs/Testing.md end to end in a real Excel under MAME: pack, run, capture and check.
/// </summary>
public class ExcelRunTests
{
    // Needs MAME on the PATH and the files the run command reads: MACEXCEL_ROM (the Mac Plus ROM), and for each release
    // its Program disk (MACEXCEL_PROGRAM for 1.00, MACEXCEL_PROGRAM_1_03 and so on) and, where that disk has no System
    // Folder, a startup disk (MACEXCEL_SYSTEM, MACEXCEL_SYSTEM_1_5 and so on). Passes without checking anything
    // otherwise.
    private static bool CanRun => CanRunRelease("1.0");

    private static bool CanRunRelease(string release)
    {
        static bool Exists(string variable) => Environment.GetEnvironmentVariable(variable) is { Length: > 0 } path && (File.Exists(path) || Directory.Exists(path));
        string suffix = release == "1.0" ? "" : "_" + release.Replace('.', '_');
        bool needsSystem = release is not ("1.03" or "1.04");
        return Exists("MACEXCEL_ROM") && Exists("MACEXCEL_PROGRAM" + suffix) && (!needsSystem || Exists("MACEXCEL_SYSTEM" + suffix));
    }

    [Theory]
    [InlineData("1.01")]
    [InlineData("1.03")]
    [InlineData("1.04")]
    [InlineData("1.06")]
    [InlineData("1.5")]
    public void Run_RunsDriverInEachRelease(string release)
    {
        if (!CanRunRelease(release))
        {
            return;
        }

        string folder = Directory.CreateTempSubdirectory("excel-run-test-").FullName;
        try
        {
            string sheet = Path.Combine(folder, "R Basics");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Samples", "Excel 1.00", "Render Basics"), sheet);
            string disk = Path.Combine(folder, "Test.dsk");
            Assert.Equal(0, Program.Main(["excel-test", "pack", disk, sheet, "--volume", "Render", "--keep-open", "--excel", release]));
            Assert.Equal(0, Program.Main(["excel-test", "run", disk, "--excel", release, "--capture", "Render:R Basics saved=Captured.png", "-o", folder]));

            // The French 1.01 adds its own built-in names, so only the others' saves match ours.
            var saved = MfsImage.Read(File.ReadAllBytes(disk)).Files.Single(f => f.Name == "R Basics saved");
            var document = ExcelDocument.Read(saved.Data);
            Assert.Equal(ExcelDocument.Load(sheet).Cells.RowCount, document.Cells.RowCount);
            if (release != "1.01")
            {
                Assert.Equal(0, Program.Main(["excel-test", "check", disk]));
            }

            Assert.True(new FileInfo(Path.Combine(folder, "Captured.png")).Length > 0);
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

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
