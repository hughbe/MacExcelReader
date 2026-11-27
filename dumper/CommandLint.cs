using System.Text.RegularExpressions;

/// <summary>
/// Warnings about macro commands for Driver that would not do what they say in Excel, found while running
/// experiments (docs/Testing.md, docs/ExcelBugs.md): formula text in A1 notation, which macro commands read as
/// R1C1, and the shapes of formula and format that crash Excel 1.00 instead of being refused.
/// </summary>
static partial class CommandLint
{
    /// <summary>Gets the warnings for one command, empty when there are none.</summary>
    public static List<string> Check(string command)
    {
        var warnings = new List<string>();
        foreach (Match m in QuotedString().Matches(command))
        {
            string text = m.Groups[1].Value.Replace("\"\"", "\"", StringComparison.Ordinal);
            if (text.StartsWith('='))
            {
                CheckFormulaText(text, warnings);
            }
            else if (TwoExponentsSection(text) is int section and > 0)
            {
                warnings.Add($"\"{text}\" has two exponents in section {section + 1}: TEXT with it crashes Excel 1.00 (FORMAT.NUMBER refuses it).");
            }
            else if (TwoExponentsSection(text) == 0)
            {
                warnings.Add($"\"{text}\" has two exponents: Excel 1.00 refuses it (TEXT gives #VALUE!).");
            }
        }

        return warnings;
    }

    private static void CheckFormulaText(string text, List<string> warnings)
    {
        // R1C1 references and whole rows or columns (R1C1, R[-1]C, R2:R3, C[-1]) taken out, what looks like a cell
        // in A1 notation is one: FORMULA("=A1+1") gives #VALUE!.
        // Text constants inside the formula ("A1") are not references.
        string withoutR1C1 = R1C1Reference().Replace(Regex.Replace(text, "\"[^\"]*\"", " "), " ");
        if (A1Reference().Match(withoutR1C1) is { Success: true } a1)
        {
            warnings.Add($"\"{text}\" refers to {a1.Value} in A1 notation, but macro commands read formula text as R1C1 whatever A1.R1C1 says, so Excel will refuse it (#VALUE!). Write it as R1C1, or type it with a FILE.type.");
        }

        if (MinusAfterColon().IsMatch(text))
        {
            warnings.Add($"\"{text}\" has a minus sign straight after ':', which crashes Excel 1.00 (docs/ExcelBugs.md).");
        }

        if (FunctionThenRelativeUnion().IsMatch(text))
        {
            warnings.Add($"\"{text}\" joins a function call with ':' or a space to a union ending in a relative reference, which can crash Excel 1.00 (docs/ExcelBugs.md).");
        }
    }

    // The first section of a number format with two exponent markers, or -1.
    private static int TwoExponentsSection(string format) => Array.FindIndex(format.Split(';'), section => Exponent().Count(section) > 1);

    [GeneratedRegex("\"((?:[^\"]|\"\")*)\"")]
    private static partial Regex QuotedString();

    [GeneratedRegex(@"R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?|R(\[-?\d+\]|\d+)|(?<![A-Za-z])C(\[-?\d+\]|\d+)")]
    private static partial Regex R1C1Reference();

    [GeneratedRegex(@"(?<![A-Za-z0-9_.!])\$?[A-Za-z]{1,2}\$?\d{1,5}(?![\dA-Za-z(_])")]
    private static partial Regex A1Reference();

    [GeneratedRegex(@":\s*-")]
    private static partial Regex MinusAfterColon();

    // A call's closing parenthesis, ':' or a space, then a parenthesised union whose last item has a relative part.
    [GeneratedRegex(@"\)\s*[: ]\s*\([^()]*,[^(),]*\[[^(),]*\)")]
    private static partial Regex FunctionThenRelativeUnion();

    [GeneratedRegex(@"[Ee][+-]")]
    private static partial Regex Exponent();
}
