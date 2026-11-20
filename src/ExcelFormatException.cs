namespace MacExcelReader;

/// <summary>
/// The exception thrown when a file is not a valid Excel 1.x for Macintosh document.
/// </summary>
public class ExcelFormatException : Exception
{
    /// <summary>
    /// Initializes a new instance with a message.
    /// </summary>
    public ExcelFormatException(string message) : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance with a message and an inner exception.
    /// </summary>
    public ExcelFormatException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
