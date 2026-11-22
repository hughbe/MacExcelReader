using System.Diagnostics.CodeAnalysis;

namespace MacExcelReader;

/// <summary>
/// A zero-based cell address that converts to and from A1 notation (for example "B19" or "$IV$16384").
/// </summary>
/// <param name="Row">The zero-based row.</param>
/// <param name="Column">The zero-based column.</param>
public readonly record struct ExcelCellAddress(int Row, int Column)
{
    /// <summary>
    /// Parses an A1-style address. Dollar signs are ignored; columns run A–IV and rows 1–16384.
    /// </summary>
    /// <exception cref="FormatException">The text is not a valid address.</exception>
    public static ExcelCellAddress Parse(string text)
        => TryParse(text, out var address) ? address : throw new FormatException($"'{text}' is not a cell address such as B19.");

    /// <summary>
    /// Tries to parse an A1-style address.
    /// </summary>
    public static bool TryParse([NotNullWhen(true)] string? text, out ExcelCellAddress address)
    {
        address = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var span = text.AsSpan().Trim();
        int i = 0;
        if (i < span.Length && span[i] == '$')
        {
            i++;
        }

        int column = 0;
        int letters = 0;
        while (i < span.Length && char.IsAsciiLetter(span[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(span[i]) - 'A' + 1);
            letters++;
            i++;
            if (letters > 2)
            {
                return false;
            }
        }

        if (letters == 0)
        {
            return false;
        }

        if (i < span.Length && span[i] == '$')
        {
            i++;
        }

        if (i >= span.Length || !int.TryParse(span[i..], System.Globalization.NumberStyles.None, null, out int row))
        {
            return false;
        }

        if (row < 1 || row > ExcelCellGrid.MaxRows || column > ExcelCellGrid.MaxColumns)
        {
            return false;
        }

        address = new ExcelCellAddress(row - 1, column - 1);
        return true;
    }

    /// <summary>
    /// Formats the address in A1 notation.
    /// </summary>
    public override string ToString() => ColumnName(Column) + (Row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>
    /// Gets the letters of a zero-based column index: 0 is A, 25 is Z, 26 is AA and 255, the last column, is IV.
    /// </summary>
    public static string ColumnName(int column) => FormulaDecompiler.ColumnName(column);
}
