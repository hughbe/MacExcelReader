namespace MacExcelReader;

/// <summary>
/// The dense cell table of a sheet (spec §7). Every cell inside the grid exists; empty cells may still
/// carry formatting.
/// </summary>
public sealed class ExcelCellGrid
{
    /// <summary>The maximum number of rows.</summary>
    public const int MaxRows = 16384;

    /// <summary>The maximum number of columns.</summary>
    public const int MaxColumns = 256;

    private readonly List<ExcelCell[]> _rows = [];

    /// <summary>Gets the number of rows.</summary>
    public int RowCount => _rows.Count;

    /// <summary>Gets the number of columns.</summary>
    public int ColumnCount { get; private set; }

    /// <summary>
    /// Gets a cell, creating and growing the grid as needed.
    /// </summary>
    public ExcelCell this[int row, int column]
    {
        get
        {
            EnsureSize(row + 1, column + 1);
            return _rows[row][column];
        }
    }

    /// <summary>
    /// Gets a cell, or null if it is outside the grid.
    /// </summary>
    public ExcelCell? TryGet(int row, int column)
        => row >= 0 && row < RowCount && column >= 0 && column < ColumnCount ? _rows[row][column] : null;

    /// <summary>
    /// Gets the cells that have a value, in row-major order.
    /// </summary>
    public IEnumerable<(int Row, int Column, ExcelCell Cell)> NonEmptyCells
    {
        get
        {
            for (int r = 0; r < _rows.Count; r++)
            {
                for (int c = 0; c < ColumnCount; c++)
                {
                    if (_rows[r][c].HasValue)
                    {
                        yield return (r, c, _rows[r][c]);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Grows the grid to at least the given size. The grid never shrinks.
    /// </summary>
    public void EnsureSize(int rows, int columns)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rows);
        ArgumentOutOfRangeException.ThrowIfNegative(columns);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(rows, MaxRows);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(columns, MaxColumns);
        if (columns > ColumnCount)
        {
            for (int r = 0; r < _rows.Count; r++)
            {
                var old = _rows[r];
                var grown = new ExcelCell[columns];
                old.CopyTo(grown, 0);
                for (int c = old.Length; c < columns; c++)
                {
                    grown[c] = new ExcelCell();
                }

                _rows[r] = grown;
            }

            ColumnCount = columns;
        }

        while (_rows.Count < rows)
        {
            var row = new ExcelCell[ColumnCount];
            for (int c = 0; c < row.Length; c++)
            {
                row[c] = new ExcelCell();
            }

            _rows.Add(row);
        }
    }

    internal void Load(int rows, int columns, Func<int, int, ExcelCell> factory)
    {
        _rows.Clear();
        ColumnCount = columns;
        for (int r = 0; r < rows; r++)
        {
            var row = new ExcelCell[columns];
            for (int c = 0; c < columns; c++)
            {
                row[c] = factory(r, c);
            }

            _rows.Add(row);
        }
    }
}
