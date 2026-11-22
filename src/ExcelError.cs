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
