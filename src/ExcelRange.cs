using System.Diagnostics.CodeAnalysis;

namespace MacExcelReader;

/// <summary>
/// A rectangular range of cells, such as <c>B2:C5</c>, from its top-left to its bottom-right cell.
/// </summary>
/// <param name="First">The top-left cell.</param>
/// <param name="Last">The bottom-right cell.</param>
public readonly record struct ExcelRange(ExcelCellAddress First, ExcelCellAddress Last) : ISpanFormattable
{
    /// <summary>
    /// Creates a range from its first and last row and column, zero-based.
    /// </summary>
    public ExcelRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        : this(new ExcelCellAddress(firstRow, firstColumn), new ExcelCellAddress(lastRow, lastColumn))
    {
    }

    /// <summary>Gets the first row.</summary>
    public int FirstRow => First.Row;

    /// <summary>Gets the first column.</summary>
    public int FirstColumn => First.Column;

    /// <summary>Gets the last row.</summary>
    public int LastRow => Last.Row;

    /// <summary>Gets the last column.</summary>
    public int LastColumn => Last.Column;

    /// <summary>Gets the number of rows.</summary>
    public int RowCount => LastRow - FirstRow + 1;

    /// <summary>Gets the number of columns.</summary>
    public int ColumnCount => LastColumn - FirstColumn + 1;

    /// <summary>Gets whether the range holds the cell.</summary>
    public bool Contains(ExcelCellAddress cell)
        => cell.Row >= FirstRow && cell.Row <= LastRow && cell.Column >= FirstColumn && cell.Column <= LastColumn;

    /// <summary>
    /// Parses a range in A1 notation, <c>B2:C5</c>, or a single cell, <c>B2</c>. Dollar signs are ignored.
    /// </summary>
    /// <exception cref="FormatException">The text is not a range.</exception>
    public static ExcelRange Parse(string text)
        => TryParse(text, out var range) ? range : throw new FormatException($"'{text}' is not a range such as B2:C5.");

    /// <summary>
    /// Tries to parse a range in A1 notation, <c>B2:C5</c>, or a single cell, <c>B2</c>.
    /// </summary>
    public static bool TryParse([NotNullWhen(true)] string? text, out ExcelRange range)
    {
        range = default;
        if (text is null)
        {
            return false;
        }

        int colon = text.IndexOf(':');
        if (colon < 0)
        {
            if (!ExcelCellAddress.TryParse(text, out var cell))
            {
                return false;
            }

            range = new ExcelRange(cell, cell);
            return true;
        }

        if (!ExcelCellAddress.TryParse(text[..colon], out var first) || !ExcelCellAddress.TryParse(text[(colon + 1)..], out var last)
            || first.Row > last.Row || first.Column > last.Column)
        {
            return false;
        }

        range = new ExcelRange(first, last);
        return true;
    }

    /// <summary>Gets the range in A1 notation: <c>B2:C5</c>, or <c>B2</c> for one cell.</summary>
    public override string ToString()
    {
        Span<char> buffer = stackalloc char[48];
        return TryFormat(buffer, out int written) ? new string(buffer[..written]) : string.Empty;
    }

    /// <summary>Formats the range in A1 notation; the format and provider are not used.</summary>
    public string ToString(string? format, IFormatProvider? formatProvider) => ToString();

    /// <summary>
    /// Writes the range in A1 notation into <paramref name="destination"/> without allocating; the format and
    /// provider are not used.
    /// </summary>
    public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format = default, IFormatProvider? provider = null)
    {
        charsWritten = 0;
        if (!First.TryFormat(destination, out int first))
        {
            return false;
        }

        if (First == Last)
        {
            charsWritten = first;
            return true;
        }

        if (first >= destination.Length || !Last.TryFormat(destination[(first + 1)..], out int last))
        {
            return false;
        }

        destination[first] = ':';
        charsWritten = first + 1 + last;
        return true;
    }

    // Checks that the range fits what a file stores for a range: a row word and a column byte each.
    internal void ValidateStorable(string parameterName)
    {
        if (FirstRow is < 0 or > ushort.MaxValue || LastRow is < 0 or > ushort.MaxValue
            || FirstColumn is < 0 or > byte.MaxValue || LastColumn is < 0 or > byte.MaxValue)
        {
            throw new ArgumentException($"The range {this} cannot be stored: rows are words and columns bytes.", parameterName);
        }
    }

    // Checks that the range is a range of the sheet: rows 0–16383, columns 0–255, first before last.
    internal void Validate(string parameterName)
    {
        if (FirstRow < 0 || FirstColumn < 0 || LastRow >= ExcelCellGrid.MaxRows || LastColumn >= ExcelCellGrid.MaxColumns
            || FirstRow > LastRow || FirstColumn > LastColumn)
        {
            throw new ArgumentException($"The range {this} is not a range of the sheet with its first cell first.", parameterName);
        }
    }
}
