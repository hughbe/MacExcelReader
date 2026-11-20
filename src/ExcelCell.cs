namespace MacExcelReader;

/// <summary>
/// A cell: one entry of the cell table (spec §7) and, for numbers, text and formulas, its cell data (spec §8).
/// </summary>
public sealed class ExcelCell
{
    // Raw cell-table words. For types with cell data, word 0 and the top nibble of word 1 hold the
    // data offset and are recomputed on write; for other types they are kept as read.
    internal ushort Word0;
    internal ushort Word1;
    internal ushort Word2 = 0x4000;

    private double _number;
    private string _text = string.Empty;
    private ExcelFormula? _formula;

    /// <summary>
    /// Creates an empty, locked cell.
    /// </summary>
    public ExcelCell()
    {
    }

    internal ExcelCell(ushort word0, ushort word1, ushort word2)
    {
        Word0 = word0;
        Word1 = word1;
        Word2 = word2;
    }

    /// <summary>
    /// Gets the cell type.
    /// </summary>
    public ExcelCellType Type
    {
        get => (ExcelCellType)((Word1 >> 9) & 7);
        private set => Word1 = (ushort)((Word1 & ~0x0E00) | ((int)value << 9));
    }

    /// <summary>
    /// Gets whether the cell has a value.
    /// </summary>
    public bool HasValue => Type != ExcelCellType.Empty;

    /// <summary>
    /// Gets the integer value of an <see cref="ExcelCellType.Integer"/> cell.
    /// </summary>
    public short IntegerValue => Type == ExcelCellType.Integer
        ? unchecked((short)Word0)
        : throw new InvalidOperationException("The cell is not an integer cell.");

    /// <summary>
    /// Gets the value of a <see cref="ExcelCellType.Number"/> cell.
    /// </summary>
    public double NumberValue => Type == ExcelCellType.Number
        ? _number
        : throw new InvalidOperationException("The cell is not a number cell.");

    /// <summary>
    /// Gets the value of a <see cref="ExcelCellType.Text"/> cell.
    /// </summary>
    public string TextValue => Type == ExcelCellType.Text
        ? _text
        : throw new InvalidOperationException("The cell is not a text cell.");

    /// <summary>
    /// Gets the formula of a formula cell, or null.
    /// </summary>
    public ExcelFormula? Formula => Type is ExcelCellType.Formula or ExcelCellType.NumericFormula ? _formula : null;

    /// <summary>
    /// Gets the cell's value: a <see cref="short"/>, <see cref="double"/>, <see cref="string"/>, the
    /// formula's current value, or null for an empty cell.
    /// </summary>
    public object? Value => Type switch
    {
        ExcelCellType.Integer => IntegerValue,
        ExcelCellType.Number => _number,
        ExcelCellType.Text => _text,
        ExcelCellType.NumericFormula or ExcelCellType.Formula => _formula!.Result.Value,
        _ => null,
    };

    /// <summary>
    /// Makes the cell an integer cell.
    /// </summary>
    public void SetInteger(short value)
    {
        ClearValue();
        Type = ExcelCellType.Integer;
        Word0 = unchecked((ushort)value);
    }

    /// <summary>
    /// Makes the cell a number cell.
    /// </summary>
    public void SetNumber(double value)
    {
        ClearValue();
        Type = ExcelCellType.Number;
        _number = value;
    }

    /// <summary>
    /// Sets a numeric value, using the compact integer form when the value is a whole number that fits in 16 bits.
    /// </summary>
    public void SetValue(double value)
    {
        if (value == Math.Floor(value) && value >= short.MinValue && value <= short.MaxValue && !(value == 0 && double.IsNegative(value)))
        {
            SetInteger((short)value);
        }
        else
        {
            SetNumber(value);
        }
    }

    /// <summary>
    /// Makes the cell a text cell. The text must encode to at most 255 Mac OS Roman bytes.
    /// </summary>
    public void SetText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (MacRoman.GetBytes(value).Length > 255)
        {
            throw new ArgumentException("Cell text cannot be longer than 255 characters.", nameof(value));
        }

        ClearValue();
        Type = ExcelCellType.Text;
        _text = value;
    }

    /// <summary>
    /// Makes the cell a formula cell.
    /// </summary>
    public void SetFormula(ExcelFormula formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ClearValue();
        _formula = formula;
        Type = formula.Result.Kind == ExcelFormulaResultKind.Number ? ExcelCellType.NumericFormula : ExcelCellType.Formula;
    }

    /// <summary>
    /// Removes the cell's value, keeping its formatting.
    /// </summary>
    public void Clear()
    {
        ClearValue();
        Word0 = 0;
        Word1 &= 0x01FF;
    }

    private void ClearValue()
    {
        _formula = null;
        _text = string.Empty;
        _number = 0;
        Type = ExcelCellType.Empty;
        Word0 = 0;
        Word1 &= 0x0FFF;
    }

    internal void LoadNumber(double value) => _number = value;

    internal void LoadText(string value) => _text = value;

    internal void LoadFormula(ExcelFormula formula) => _formula = formula;

    /// <summary>
    /// Gets or sets the number format index (0 = General; 0–63).
    /// </summary>
    public int NumberFormat
    {
        get => Word1 & 0x3F;
        set
        {
            ArgumentOutOfRangeException.ThrowIfNegative(value);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(value, 63);
            Word1 = (ushort)((Word1 & ~0x3F) | value);
        }
    }

    /// <summary>
    /// Gets or sets the alignment.
    /// </summary>
    public ExcelAlignment Alignment
    {
        get => (ExcelAlignment)((Word1 >> 6) & 7);
        set => Word1 = (ushort)((Word1 & ~0x01C0) | (((int)value & 7) << 6));
    }

    private bool GetFlag(int bit) => (Word2 & (1 << bit)) != 0;

    private void SetFlag(int bit, bool value)
        => Word2 = (ushort)(value ? Word2 | (1 << bit) : Word2 & ~(1 << bit));

    /// <summary>Gets or sets whether the cell is hidden when the document is protected.</summary>
    public bool IsHidden { get => GetFlag(15); set => SetFlag(15, value); }

    /// <summary>Gets or sets whether the cell is locked when the document is protected.</summary>
    public bool IsLocked { get => GetFlag(14); set => SetFlag(14, value); }

    /// <summary>Gets or sets whether the cell's formula uses an array-valued function.</summary>
    public bool UsesArrayFunction { get => GetFlag(13); set => SetFlag(13, value); }

    /// <summary>Gets or sets the bottom border.</summary>
    public bool BottomBorder { get => GetFlag(5); set => SetFlag(5, value); }

    /// <summary>Gets or sets the top border.</summary>
    public bool TopBorder { get => GetFlag(4); set => SetFlag(4, value); }

    /// <summary>Gets or sets the right border.</summary>
    public bool RightBorder { get => GetFlag(3); set => SetFlag(3, value); }

    /// <summary>Gets or sets the left border.</summary>
    public bool LeftBorder { get => GetFlag(2); set => SetFlag(2, value); }

    /// <summary>Gets or sets italic.</summary>
    public bool Italic { get => GetFlag(1); set => SetFlag(1, value); }

    /// <summary>Gets or sets bold.</summary>
    public bool Bold { get => GetFlag(0); set => SetFlag(0, value); }

    /// <summary>
    /// Gets the raw attribute word (the third cell-table word).
    /// </summary>
    public ushort RawAttributes => Word2;

    /// <inheritdoc/>
    public override string ToString() => Type switch
    {
        ExcelCellType.Empty => string.Empty,
        _ => Value?.ToString() ?? string.Empty,
    };
}
