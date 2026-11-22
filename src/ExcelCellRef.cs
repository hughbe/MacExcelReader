using System.Buffers.Binary;

namespace MacExcelReader;

/// <summary>
/// A lightweight, allocation-free view of a cell, produced by <see cref="ExcelCellGrid.EnumerateNonEmpty"/>.
/// It reads directly from the file's cell table and only creates an <see cref="ExcelCell"/> when
/// <see cref="GetCell"/> is called.
/// </summary>
/// <example>
/// <code>
/// double total = 0;
/// foreach (var cell in document.Cells.EnumerateNonEmpty())
/// {
///     if (cell.TryGetNumber(out double value))
///     {
///         total += value;
///     }
/// }
/// </code>
/// </example>
public readonly struct ExcelCellRef
{
    private readonly ExcelCellGrid _grid;
    private readonly ushort _word0;
    private readonly ushort _word1;
    private readonly ushort _word2;

    internal ExcelCellRef(ExcelCellGrid grid, int row, int column, ushort word0, ushort word1, ushort word2)
    {
        _grid = grid;
        Row = row;
        Column = column;
        _word0 = word0;
        _word1 = word1;
        _word2 = word2;
    }

    /// <summary>Gets the zero-based row.</summary>
    public int Row { get; }

    /// <summary>Gets the zero-based column.</summary>
    public int Column { get; }

    /// <summary>Gets the cell's address.</summary>
    public ExcelCellAddress Address => new(Row, Column);

    /// <summary>Gets the cell type.</summary>
    public ExcelCellType Type => (ExcelCellType)((_word1 >> 9) & 7);

    /// <summary>Gets the number format index.</summary>
    public int NumberFormat => _word1 & 0x3F;

    /// <summary>Gets the alignment.</summary>
    public ExcelAlignment Alignment => (ExcelAlignment)((_word1 >> 6) & 7);

    /// <summary>Gets the raw attribute word (bold, italic, borders, protection).</summary>
    public ushort RawAttributes => _word2;

    /// <summary>Gets whether the cell is bold.</summary>
    public bool Bold => (_word2 & 1) != 0;

    /// <summary>Gets whether the cell is italic.</summary>
    public bool Italic => (_word2 & 2) != 0;

    /// <summary>
    /// Gets the numeric value of an integer, number or numeric-formula cell without allocating.
    /// </summary>
    public bool TryGetNumber(out double value)
    {
        switch (Type)
        {
            case ExcelCellType.Integer:
                value = unchecked((short)_word0);
                return true;
            case ExcelCellType.Number:
            case ExcelCellType.NumericFormula:
                if (_grid.TryGetRawData(Row, Column, out var data) && data.Length >= 8)
                {
                    value = BinaryPrimitives.ReadDoubleBigEndian(data);
                    return true;
                }

                return GetCell().TryGetNumber(out value);
            default:
                value = 0;
                return false;
        }
    }

    /// <summary>
    /// Gets the text of a text cell, or null for other types. Allocates only the string.
    /// </summary>
    public string? GetText()
    {
        if (Type != ExcelCellType.Text)
        {
            return null;
        }

        if (_grid.TryGetRawData(Row, Column, out var data) && data.Length >= 1 && data.Length >= 1 + data[0])
        {
            return MacRoman.GetString(data.Slice(1, data[0]));
        }

        return GetCell().TextValue;
    }

    /// <summary>
    /// Gets the full <see cref="ExcelCell"/>, creating it if needed.
    /// </summary>
    public ExcelCell GetCell() => _grid.Find(Row, Column)!;

    /// <inheritdoc/>
    public override string ToString() => $"{Address} ({Type})";
}
