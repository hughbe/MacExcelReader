using System.Diagnostics;
using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// The dense cell table of a sheet (spec §9). Every position inside the grid is a cell; empty cells may
/// still carry formatting.
/// </summary>
/// <remarks>
/// <para>
/// A grid read from a file is lazy: rows and cells are created from the file's cell table the first time
/// they are accessed, and a cell's text or formula is decoded the first time its value is read.
/// Reading a document therefore costs almost nothing beyond reading the file, and iterating
/// <see cref="NonEmptyCells"/> only creates objects for cells that have values.
/// </para>
/// <para>
/// Cells that are never accessed are written back by copying their bytes, so loading and saving an
/// unchanged document does not decode its cells at all.
/// </para>
/// </remarks>
public sealed class ExcelCellGrid
{
    /// <summary>The maximum number of rows.</summary>
    public const int MaxRows = 16384;

    /// <summary>The maximum number of columns.</summary>
    public const int MaxColumns = 256;

    private const int EntrySize = 6;

    // The created cells, row by row; null until the first is created, and a row's array null until a cell in it
    // is, so reading a document does not allocate per row.
    private ExcelCell?[]?[]? _rows;
    private int _rowCount;

    // The file the grid was read from, or null for a grid built in memory.
    private byte[]? _source;
    private int _tableStart;
    private int _dataStart;
    private int _sourceRows;
    private int _sourceColumns;

    /// <summary>Gets the number of rows.</summary>
    public int RowCount => _rowCount;

    /// <summary>Gets the number of columns.</summary>
    public int ColumnCount { get; private set; }

    /// <summary>
    /// Gets a cell, creating it and growing the grid as needed.
    /// </summary>
    /// <param name="row">The zero-based row (0–16383).</param>
    /// <param name="column">The zero-based column (0–255).</param>
    /// <exception cref="ArgumentOutOfRangeException">The row or column is outside the sheet.</exception>
    public ExcelCell this[int row, int column]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(row);
            ArgumentOutOfRangeException.ThrowIfNegative(column);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(row, MaxRows);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, MaxColumns);
            Grow(row + 1, column + 1);
            return GetCell(row, column);
        }
    }

    /// <summary>
    /// Gets a cell by A1-style address (for example "B19"), creating it and growing the grid as needed.
    /// </summary>
    public ExcelCell this[string address]
    {
        get
        {
            var parsed = ExcelCellAddress.Parse(address);
            return this[parsed.Row, parsed.Column];
        }
    }

    /// <summary>
    /// Gets a cell by address, creating it and growing the grid as needed.
    /// </summary>
    public ExcelCell this[ExcelCellAddress address] => this[address.Row, address.Column];

    /// <summary>
    /// Gets a cell by address, or null if it is outside the grid. Never grows the grid.
    /// </summary>
    public ExcelCell? Find(ExcelCellAddress address) => Find(address.Row, address.Column);

    /// <summary>
    /// Gets a cell by A1-style address, or null if it is outside the grid. Never grows the grid.
    /// </summary>
    public ExcelCell? Find(string address)
    {
        var parsed = ExcelCellAddress.Parse(address);
        return Find(parsed.Row, parsed.Column);
    }

    /// <summary>
    /// Gets a cell, or null if it is outside the grid. Never grows the grid.
    /// </summary>
    public ExcelCell? Find(int row, int column)
        => (uint)row < (uint)RowCount && (uint)column < (uint)ColumnCount ? GetCell(row, column) : null;

    /// <summary>
    /// Gets the alignment of a cell without creating it, or <see cref="ExcelAlignment.General"/> outside the grid.
    /// Blank cells have an alignment too.
    /// </summary>
    public ExcelAlignment GetAlignment(int row, int column)
    {
        if ((uint)row >= (uint)RowCount || (uint)column >= (uint)ColumnCount)
        {
            return ExcelAlignment.General;
        }

        var cell = _rows?[row]?[column];
        if (cell is not null)
        {
            return cell.Alignment;
        }

        return IsSourcePosition(row, column) ? CellWords.Alignment(SourceWord1(row, column)) : ExcelAlignment.General;
    }

    /// <summary>
    /// Gets the type of a cell without creating it, or <see cref="ExcelCellType.Empty"/> outside the grid.
    /// </summary>
    public ExcelCellType GetCellType(int row, int column)
    {
        if ((uint)row >= (uint)RowCount || (uint)column >= (uint)ColumnCount)
        {
            return ExcelCellType.Empty;
        }

        var cell = _rows?[row]?[column];
        if (cell is not null)
        {
            return cell.Type;
        }

        return IsSourcePosition(row, column) ? CellWords.Type(SourceWord1(row, column)) : ExcelCellType.Empty;
    }

    /// <summary>
    /// Gets the cells that have a value, in row-major order, as <see cref="ExcelCellRef"/> views read straight from
    /// the cell table: a <c>foreach</c> over them allocates nothing, and no <see cref="ExcelCell"/> is created unless
    /// you call <see cref="ExcelCellRef.GetCell"/>.
    /// </summary>
    public NonEmptyCellCollection NonEmptyCells => new(this);

    /// <summary>
    /// The cells of a grid that have a value. <c>foreach</c> uses its struct enumerator; LINQ works too, through
    /// <see cref="IEnumerable{T}"/>, at the cost of boxing the enumerator.
    /// </summary>
    public readonly struct NonEmptyCellCollection : IEnumerable<ExcelCellRef>
    {
        private readonly ExcelCellGrid _grid;

        internal NonEmptyCellCollection(ExcelCellGrid grid) => _grid = grid;

        /// <summary>Gets an enumerator over the cells that have a value.</summary>
        public NonEmptyEnumerator GetEnumerator() => new(_grid);

        IEnumerator<ExcelCellRef> IEnumerable<ExcelCellRef>.GetEnumerator() => GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// An allocation-free enumerator over the cells that have a value.
    /// </summary>
    public struct NonEmptyEnumerator : IEnumerator<ExcelCellRef>
    {
        private readonly ExcelCellGrid _grid;
        private int _row;
        private int _column;

        internal NonEmptyEnumerator(ExcelCellGrid grid)
        {
            _grid = grid;
            _row = 0;
            _column = -1;
            Current = default;
        }

        /// <summary>Gets the current cell.</summary>
        public ExcelCellRef Current { get; private set; }

        readonly object System.Collections.IEnumerator.Current => Current;

        /// <summary>Advances to the next cell that has a value.</summary>
        public bool MoveNext()
        {
            int columns = _grid.ColumnCount;
            int rows = _grid.RowCount;
            while (_row < rows)
            {
                while (++_column < columns)
                {
                    _grid.GetWords(_row, _column, out ushort word0, out ushort word1, out ushort word2);
                    if (CellWords.Type(word1) != ExcelCellType.Empty)
                    {
                        Current = new ExcelCellRef(_grid, _row, _column, word0, word1, word2);
                        return true;
                    }
                }

                _row++;
                _column = -1;
            }

            return false;
        }

        /// <summary>Starts again from the first cell.</summary>
        public void Reset()
        {
            _row = 0;
            _column = -1;
            Current = default;
        }

        /// <inheritdoc/>
        public readonly void Dispose()
        {
        }
    }

    /// <summary>
    /// Gets a cell's table words, from the created cell if there is one, otherwise from the file.
    /// </summary>
    internal void GetWords(int row, int column, out ushort word0, out ushort word1, out ushort word2)
    {
        var cell = _rows?[row]?[column];
        if (cell is not null)
        {
            word0 = cell.Word0;
            word1 = cell.TableWord1;
            word2 = cell.Word2;
        }
        else if (IsSourcePosition(row, column))
        {
            var entry = _source.AsSpan(EntryOffset(row, column), EntrySize);
            word0 = BinaryPrimitives.ReadUInt16BigEndian(entry);
            word1 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(2));
            word2 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(4));
        }
        else
        {
            word0 = 0;
            word1 = 0;
            word2 = 0x4000;
        }
    }

    /// <summary>
    /// Gets the file bytes starting at an untouched cell's data, or returns false if the cell has been
    /// created or is not from the file.
    /// </summary>
    internal bool TryGetRawData(int row, int column, out ReadOnlySpan<byte> data)
    {
        data = default;
        if (_rows?[row]?[column] is not null || !IsSourcePosition(row, column))
        {
            return false;
        }

        var entry = _source.AsSpan(EntryOffset(row, column), EntrySize);
        int offset = DataOffset(BinaryPrimitives.ReadUInt16BigEndian(entry), BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(2)));
        if ((uint)offset >= (uint)_source!.Length)
        {
            return false;
        }

        data = _source.AsSpan(offset);
        return true;
    }

    /// <summary>
    /// Counts the cells that have a value, without creating any.
    /// </summary>
    public int CountNonEmpty()
    {
        int count = 0;
        for (int r = 0; r < _rowCount; r++)
        {
            for (int c = 0; c < ColumnCount; c++)
            {
                if (GetCellType(r, c) != ExcelCellType.Empty)
                {
                    count++;
                }
            }
        }

        return count;
    }

    /// <summary>
    /// Grows the grid to at least the given size. The grid never shrinks.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The size is negative or larger than the sheet.</exception>
    public void EnsureSize(int rows, int columns)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rows);
        ArgumentOutOfRangeException.ThrowIfNegative(columns);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(rows, MaxRows);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(columns, MaxColumns);
        Grow(rows, columns);
    }

    // Grows the table, which can hold one row and column past the sheet's last: the blank margin Excel keeps after
    // the last row and column in use, even after row 16384 (spec §6.1).
    internal void Grow(int rows, int columns)
    {
        Debug.Assert(rows <= MaxRows + 1 && columns <= MaxColumns + 1);
        if (columns > ColumnCount)
        {
            // A row's array can be longer than the grid is wide. Doubling it keeps filling a sheet column by column
            // from copying every row each time a column is added.
            for (int r = 0; r < _rowCount; r++)
            {
                if (_rows?[r] is { } row && row.Length < columns)
                {
                    var grown = new ExcelCell?[Math.Max(columns, Math.Min(MaxColumns + 1, row.Length * 2))];
                    row.CopyTo(grown, 0);
                    _rows![r] = grown;
                }
            }

            ColumnCount = columns;
        }

        if (_rowCount < rows)
        {
            if (_rows is not null && _rows.Length < rows)
            {
                Array.Resize(ref _rows, Math.Max(rows, Math.Min(MaxRows, _rows.Length * 2)));
            }

            _rowCount = rows;
        }
    }

    private ExcelCell GetCell(int row, int column)
    {
        _rows ??= new ExcelCell?[]?[_rowCount];
        var cells = _rows[row] ??= new ExcelCell?[ColumnCount];
        return cells[column] ??= CreateCell(row, column);
    }

    private ExcelCell CreateCell(int row, int column)
    {
        if (!IsSourcePosition(row, column))
        {
            return new ExcelCell();
        }

        var entry = _source.AsSpan(EntryOffset(row, column), EntrySize);
        ushort word0 = BinaryPrimitives.ReadUInt16BigEndian(entry);
        ushort word1 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(2));
        var cell = new ExcelCell(word0, word1, BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(4)));
        if (HasCellData(cell.Type))
        {
            int offset = DataOffset(word0, word1);
            if (offset >= _source!.Length)
            {
                throw new ExcelFormatException($"The data of cell {new ExcelCellAddress(row, column)} lies outside the file.");
            }

            cell.LoadLazy(_source, offset);
        }

        return cell;
    }

    internal static bool HasCellData(ExcelCellType type)
        => type is ExcelCellType.Number or ExcelCellType.Text or ExcelCellType.NumericFormula or ExcelCellType.Formula;

    private bool IsSourcePosition(int row, int column) => _source is not null && row < _sourceRows && column < _sourceColumns;

    private int EntryOffset(int row, int column) => _tableStart + (row * _sourceColumns + column) * EntrySize;

    private ushort SourceWord1(int row, int column) => BinaryPrimitives.ReadUInt16BigEndian(_source.AsSpan(EntryOffset(row, column) + 2));

    private int DataOffset(ushort word0, ushort word1) => _dataStart + (word0 << 4) + (word1 >> 12);

    /// <summary>
    /// Attaches the grid to a file's cell table. Rows and cells are created on demand.
    /// </summary>
    internal void AttachSource(byte[] source, int tableStart, int dataStart, int rows, int columns)
    {
        if (tableStart + (long)rows * columns * EntrySize > source.Length)
        {
            throw new ExcelFormatException("The cell table runs past the end of the file.");
        }

        _rows = null;
        _rowCount = rows;

        _source = source;
        _tableStart = tableStart;
        _dataStart = dataStart;
        _sourceRows = rows;
        _sourceColumns = columns;
        ColumnCount = columns;
    }

    /// <summary>
    /// Gets the file's table entries for a row none of whose cells has been created, so the writer can copy them in
    /// one piece, or returns false. The entries cover the file's columns, which may be fewer than the grid's.
    /// </summary>
    internal bool TryGetSourceRow(int row, out ReadOnlySpan<byte> entries)
    {
        entries = default;
        if (_source is null || row >= _sourceRows || _rows?[row] is not null)
        {
            return false;
        }

        entries = _source.AsSpan(EntryOffset(row, 0), _sourceColumns * EntrySize);
        return true;
    }

    /// <summary>
    /// Gets the cell data of a cell read from the file, given its table words.
    /// </summary>
    internal ReadOnlySpan<byte> GetSourceData(int row, int column, ushort word0, ushort word1, ExcelCellType type)
    {
        int offset = DataOffset(word0, word1);
        if (offset >= _source!.Length)
        {
            throw new ExcelFormatException($"The data of cell {new ExcelCellAddress(row, column)} lies outside the file.");
        }

        var tail = _source.AsSpan(offset);
        return tail.Slice(0, ExcelFormula.MeasureCellData(type, tail));
    }

    /// <summary>
    /// Gets an untouched cell's raw table entry and cell data so the writer can copy them, or returns
    /// false if the cell has been created (and may have changed) or is not from the file.
    /// </summary>
    internal bool TryGetRawCell(int row, int column, out ushort word0, out ushort word1, out ushort word2, out ReadOnlySpan<byte> data)
    {
        word0 = word1 = word2 = 0;
        data = default;
        if (_rows?[row]?[column] is not null || !IsSourcePosition(row, column))
        {
            return false;
        }

        var entry = _source.AsSpan(EntryOffset(row, column), EntrySize);
        word0 = BinaryPrimitives.ReadUInt16BigEndian(entry);
        word1 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(2));
        word2 = BinaryPrimitives.ReadUInt16BigEndian(entry.Slice(4));
        var type = CellWords.Type(word1);
        if (HasCellData(type))
        {
            data = GetSourceData(row, column, word0, word1, type);
        }

        return true;
    }
}
