namespace MacExcelReader;

/// <summary>
/// An Excel error value. The numbers are the codes Excel 1.x stores (spec §12.4), the same as BIFF's.
/// </summary>
public enum ExcelError : byte
{
    /// <summary><c>#NULL!</c>: an intersection with no cells.</summary>
    Null = 0x00,

    /// <summary><c>#DIV/0!</c>: division by zero.</summary>
    DivideByZero = 0x07,

    /// <summary><c>#VALUE!</c>: an argument or operand of the wrong type.</summary>
    Value = 0x0F,

    /// <summary><c>#REF!</c>: a reference to a deleted cell.</summary>
    Reference = 0x17,

    /// <summary><c>#NAME?</c>: a name that is not defined.</summary>
    Name = 0x1D,

    /// <summary><c>#NUM!</c>: a number out of range.</summary>
    Number = 0x24,

    /// <summary><c>#N/A</c>: no value is available.</summary>
    NotAvailable = 0x2A,
}

/// <summary>
/// Methods for <see cref="ExcelError"/>.
/// </summary>
public static class ExcelErrorExtensions
{
    /// <summary>
    /// Gets the text Excel shows for an error, such as <c>#N/A</c>; <c>#ERR</c> and the code for a code Excel does
    /// not have.
    /// </summary>
    public static string ToText(this ExcelError error) => error switch
    {
        ExcelError.Null => "#NULL!",
        ExcelError.DivideByZero => "#DIV/0!",
        ExcelError.Value => "#VALUE!",
        ExcelError.Reference => "#REF!",
        ExcelError.Name => "#NAME?",
        ExcelError.Number => "#NUM!",
        ExcelError.NotAvailable => "#N/A",
        _ => $"#ERR{(byte)error}",
    };
}
