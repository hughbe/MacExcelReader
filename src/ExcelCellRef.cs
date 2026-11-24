using System.Buffers.Binary;
using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A lightweight, allocation-free view of a cell, produced by <see cref="ExcelCellGrid.NonEmptyCells"/>.
/// It reads directly from the file's cell table and only creates an <see cref="ExcelCell"/> when
/// <see cref="GetCell"/> is called.
/// </summary>
/// <example>
/// <code>
/// double total = 0;
/// foreach (var cell in document.Cells.NonEmptyCells)
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
    public ExcelCellType Type => CellWords.Type(_word1);

    /// <summary>Gets the number format index.</summary>
    public int NumberFormat => CellWords.NumberFormat(_word1);

    /// <summary>Gets the alignment.</summary>
    public ExcelAlignment Alignment => CellWords.Alignment(_word1);

    /// <summary>Gets the raw attribute word (bold, italic, borders, protection).</summary>
    public ushort RawAttributes => _word2;

    /// <summary>Gets whether the cell is bold.</summary>
    public bool Bold => CellWords.Has(_word2, CellWords.BoldBit);

    /// <summary>Gets whether the cell is italic.</summary>
    public bool Italic => CellWords.Has(_word2, CellWords.ItalicBit);

    /// <summary>Gets whether the cell has a left border.</summary>
    public bool LeftBorder => CellWords.Has(_word2, CellWords.LeftBorderBit);

    /// <summary>Gets whether the cell has a right border.</summary>
    public bool RightBorder => CellWords.Has(_word2, CellWords.RightBorderBit);

    /// <summary>Gets whether the cell has a top border.</summary>
    public bool TopBorder => CellWords.Has(_word2, CellWords.TopBorderBit);

    /// <summary>Gets whether the cell has a bottom border.</summary>
    public bool BottomBorder => CellWords.Has(_word2, CellWords.BottomBorderBit);

    /// <summary>Gets whether the cell is locked when the document is protected.</summary>
    public bool IsLocked => CellWords.Has(_word2, CellWords.LockedBit);

    /// <summary>Gets whether the cell is hidden when the document is protected.</summary>
    public bool IsHidden => CellWords.Has(_word2, CellWords.HiddenBit);

    /// <summary>
    /// Gets the cell's value, read from the file without creating the cell: allocation-free apart from text.
    /// </summary>
    public ExcelValue Value
    {
        get
        {
            switch (Type)
            {
                case ExcelCellType.Integer:
                    return ExcelValue.FromNumber(_word0);
                case ExcelCellType.BooleanOrError:
                    return (_word0 >> 8) != 0 ? ExcelValue.FromErrorCode((byte)_word0) : ExcelValue.FromBoolean((byte)_word0 != 0);
                case ExcelCellType.Number:
                    return _grid.TryGetRawData(Row, Column, out var number) && number.Length >= 8
                        ? ExcelValue.FromNumber(BinaryPrimitives.ReadDoubleBigEndian(number))
                        : GetCell().Value;
                case ExcelCellType.Text:
                    return _grid.TryGetRawData(Row, Column, out var text) && text.Length >= 1 + text[0]
                        ? ExcelValue.FromText(MacRoman.GetString(text.Slice(1, text[0])))
                        : GetCell().Value;
                case ExcelCellType.NumericFormula:
                case ExcelCellType.Formula:
                    return _grid.TryGetRawData(Row, Column, out var record) ? ExcelFormula.ReadResult(record) : GetCell().Value;
                default:
                    return ExcelValue.Empty;
            }
        }
    }

    /// <summary>
    /// Gets the numeric value of an integer, number or numeric-formula cell without allocating.
    /// </summary>
    public bool TryGetNumber(out double value)
    {
        switch (Type)
        {
            case ExcelCellType.Integer:
                value = _word0;
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
    /// Copies the text of a text cell into <paramref name="destination"/> without allocating a string. Text is at
    /// most 255 characters, so a 255-character buffer always has room.
    /// </summary>
    /// <returns>False if the cell is not a text cell or the text does not fit.</returns>
    public bool TryGetText(Span<char> destination, out int charsWritten)
    {
        charsWritten = 0;
        if (Type != ExcelCellType.Text)
        {
            return false;
        }

        if (_grid.TryGetRawData(Row, Column, out var data) && data.Length >= 1 + data[0])
        {
            var bytes = data.Slice(1, data[0]);
            if (bytes.Length > destination.Length)
            {
                return false;
            }

            charsWritten = MacRoman.GetChars(bytes, destination);
            return true;
        }

        string text = GetCell().TextValue;
        if (!text.AsSpan().TryCopyTo(destination))
        {
            return false;
        }

        charsWritten = text.Length;
        return true;
    }

    /// <summary>
    /// Gets the full <see cref="ExcelCell"/>, creating it if needed.
    /// </summary>
    public ExcelCell GetCell() => _grid.Find(Row, Column)!;

    /// <inheritdoc/>
    public override string ToString() => $"{Address} ({Type})";
}
