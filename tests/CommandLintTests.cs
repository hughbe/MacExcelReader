using Xunit;

namespace MacExcelReader.Tests;

/// <summary>
/// The warnings pack gives about Driver commands (dumper/CommandLint.cs), from what experiments ran into.
/// </summary>
public class CommandLintTests
{
    [Theory]
    [InlineData("=FORMULA(\"=A1+1\")", "A1 notation")]
    [InlineData("=FORMULA(\"=SUM($B$2:C3)\")", "A1 notation")]
    [InlineData("=DEFINE.NAME(\"Far\",\"=$Z$90\")", "A1 notation")]
    [InlineData("=FORMULA(\"=R1C1:-R2C2\")", "minus sign")]
    [InlineData("=FORMULA(\"=SUM(R1C1):(R1C1,R[-2]C[-3])\")", "relative reference")]
    [InlineData("=TEXT(1,\";\\###0.0E+00.00E+00$\")", "crashes Excel")]
    [InlineData("=FORMAT.NUMBER(\"0E+00E-0\")", "refuses it")]
    public void Check_WarnsAboutWhatExcelRefusesOrCrashesOn(string command, string expected)
    {
        Assert.Contains(CommandLint.Check(command), w => w.Contains(expected, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("=FORMULA(\"=R1C1+1\")")]
    [InlineData("=FORMULA(\"=R[-1]C:R2C[3]\")")]
    [InlineData("=FORMULA(\"=SUM(C1:C3)+R2:R3\")")]
    [InlineData("=FORMULA(\"=LOG10(R1C1)+ATAN2(1,1)\")")]
    [InlineData("=SELECT(!$F$5)")]
    [InlineData("=DEFINE.NAME(\"Down\",\"=R[20]C[3]\")")]
    [InlineData("=TEXT(1234.5,\"0.00E+00\")")]
    [InlineData("=FORMULA(\"=-R1C1:R2C2\")")]
    [InlineData("=FORMULA(\"=\"\"A1\"\"\")")]
    public void Check_AcceptsGoodCommands(string command)
    {
        Assert.Empty(CommandLint.Check(command));
    }
}
