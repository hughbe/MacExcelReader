namespace MacExcelReader;

/// <summary>
/// The exception thrown when formula text cannot be parsed.
/// </summary>
public sealed class ExcelFormulaSyntaxException : FormatException
{
    /// <summary>
    /// Initializes a new instance.
    /// </summary>
    public ExcelFormulaSyntaxException(string message, int position)
        : base($"{message} (at position {position + 1})")
    {
        Position = position;
    }

    /// <summary>Gets the zero-based position in the text where the problem was found.</summary>
    public int Position { get; }
}
