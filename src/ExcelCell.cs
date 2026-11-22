using MacExcelReader.Internal;

namespace MacExcelReader;

/// <summary>
/// A cell: one entry of the cell table (spec §9) and, for numbers, text and formulas, its cell data (spec §12).
/// </summary>
public sealed class ExcelCell
{
    // Raw cell-table words. For types with cell data, word 0 and the top nibble of word 1 hold the
    // data offset and are recomputed on write; for other types they are kept as read.
    internal ushort Word0;
    internal ushort Word1;
    internal ushort Word2 = 0x4000;

    // The value: a boxed double, a string or an ExcelFormula; or, until first access, the file buffer
    // the number, text or formula record is read from at _dataOffset. Numbers are never cached, so
    // reading them does not allocate.
    private object? _payload;
    private int _dataOffset;

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
        get => CellWords.Type(Word1);
        private set => Word1 = CellWords.WithType(Word1, value);
    }

    /// <summary>
    /// Gets whether the cell has a value.
    /// </summary>
    public bool HasValue => Type != ExcelCellType.Empty;

    /// <summary>
    /// Gets the integer value of an <see cref="ExcelCellType.Integer"/> cell: 0–65535.
    /// </summary>
    public ushort IntegerValue => Type == ExcelCellType.Integer
        ? Word0
        : throw new InvalidOperationException("The cell is not an integer cell.");

    /// <summary>
    /// Gets whether a <see cref="ExcelCellType.BooleanOrError"/> cell holds an error rather than TRUE or FALSE.
    /// </summary>
    public bool IsError => Type == ExcelCellType.BooleanOrError && (Word0 >> 8) != 0;

    /// <summary>
    /// Gets the value of a <see cref="ExcelCellType.BooleanOrError"/> cell that holds TRUE or FALSE.
    /// </summary>
    public bool BooleanValue => Type == ExcelCellType.BooleanOrError && !IsError
        ? (Word0 & 0xFF) != 0
        : throw new InvalidOperationException("The cell is not a TRUE or FALSE cell.");

    /// <summary>
    /// Gets the error held by a <see cref="ExcelCellType.BooleanOrError"/> cell that holds one.
    /// </summary>
    public ExcelError Error => IsError
        ? (ExcelError)(byte)Word0
        : throw new InvalidOperationException("The cell is not an error cell.");

    /// <summary>
    /// Gets the value of a <see cref="ExcelCellType.Number"/> cell.
    /// </summary>
    public double NumberValue => Type == ExcelCellType.Number
        ? ReadNumber()
        : throw new InvalidOperationException("The cell is not a number cell.");

    private double ReadNumber() => _payload switch
    {
        byte[] source => ReadLazyDouble(source),
        double value => value,
        _ => 0,
    };

    private double ReadLazyDouble(byte[] source)
    {
        if ((uint)_dataOffset > (uint)source.Length || source.Length - _dataOffset < 8)
        {
            throw new ExcelFormatException($"A numeric cell's data at file offset {_dataOffset} runs past the end of the file.");
        }

        return System.Buffers.Binary.BinaryPrimitives.ReadDoubleBigEndian(source.AsSpan(_dataOffset));
    }

    /// <summary>
    /// Gets the numeric value of an integer, number or numeric-formula cell without allocating.
    /// </summary>
    /// <returns>True if the cell has a numeric value.</returns>
    public bool TryGetNumber(out double value)
    {
        switch (Type)
        {
            case ExcelCellType.Integer:
                value = Word0;
                return true;
            case ExcelCellType.Number:
                value = ReadNumber();
                return true;
            case ExcelCellType.NumericFormula:
                // The formula's value is the first 8 bytes of its record.
                value = _payload is byte[] source
                    ? ReadLazyDouble(source)
                    : DecodeFormula().Result.Number;
                return true;
            default:
                value = 0;
                return false;
        }
    }

    /// <summary>
    /// Gets whether a text or formula value is still undecoded (it is decoded on first access).
    /// </summary>
    internal bool IsLazy => _payload is byte[];

    /// <summary>
    /// Gets the value of a <see cref="ExcelCellType.Text"/> cell.
    /// </summary>
    public string TextValue => Type == ExcelCellType.Text
        ? DecodeText()
        : throw new InvalidOperationException("The cell is not a text cell.");

    private string DecodeText()
    {
        if (_payload is byte[] source)
        {
            if (_dataOffset >= source.Length || _dataOffset + 1 + source[_dataOffset] > source.Length)
            {
                throw new ExcelFormatException($"A text cell's data at file offset {_dataOffset} runs past the end of the file.");
            }

            _payload = MacRoman.GetString(source.AsSpan(_dataOffset + 1, source[_dataOffset]));
        }

        return (string?)_payload ?? string.Empty;
    }

    private ExcelFormula DecodeFormula()
    {
        if (_payload is byte[] source)
        {
            _payload = ExcelFormula.Read(source, _dataOffset);
        }

        return (ExcelFormula)_payload!;
    }

    /// <summary>
    /// Gets the formula of a formula cell, or null.
    /// </summary>
    public ExcelFormula? Formula => Type is ExcelCellType.Formula or ExcelCellType.NumericFormula ? DecodeFormula() : null;

    /// <summary>
    /// Gets the cell's value: a <see cref="ushort"/>, <see cref="double"/>, <see cref="string"/>,
    /// <see cref="bool"/>, an <see cref="ExcelError"/>, the formula's current value, or null for an empty cell.
    /// </summary>
    public object? Value => Type switch
    {
        ExcelCellType.Integer => IntegerValue,
        ExcelCellType.BooleanOrError => IsError ? Error : BooleanValue,
        ExcelCellType.Number => ReadNumber(),
        ExcelCellType.Text => DecodeText(),
        ExcelCellType.NumericFormula or ExcelCellType.Formula => DecodeFormula().Result.Value,
        _ => null,
    };

    /// <summary>
    /// Makes the cell an integer cell.
    /// </summary>
    public void SetInteger(ushort value)
    {
        ClearValue();
        Type = ExcelCellType.Integer;
        Word0 = value;
    }

    /// <summary>
    /// Makes the cell a constant TRUE or FALSE, as Excel stores one typed into a cell.
    /// </summary>
    public void SetBoolean(bool value)
    {
        ClearValue();
        Type = ExcelCellType.BooleanOrError;
        Word0 = value ? (ushort)1 : (ushort)0;
    }

    /// <summary>
    /// Makes the cell a constant error value, as Excel stores one typed into a cell.
    /// </summary>
    public void SetError(ExcelError error)
    {
        if (!Enum.IsDefined(error))
        {
            throw new ArgumentOutOfRangeException(nameof(error), error, "Not an Excel error.");
        }

        ClearValue();
        Type = ExcelCellType.BooleanOrError;
        Word0 = (ushort)(0x100 | (byte)error);
    }

    /// <summary>
    /// Makes the cell a number cell.
    /// </summary>
    public void SetNumber(double value)
    {
        ClearValue();
        Type = ExcelCellType.Number;
        _payload = value;
    }

    /// <summary>
    /// Sets a numeric value as Excel does: a whole number from 0 to 65535 in the compact integer form, anything else,
    /// negative numbers included, as a number.
    /// </summary>
    public void SetValue(double value)
    {
        if (value == Math.Floor(value) && value >= 0 && value <= ushort.MaxValue && !double.IsNegative(value))
        {
            SetInteger((ushort)value);
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
        if (value.Length > 255)   // Mac OS Roman has one byte per character
        {
            throw new ArgumentException("Cell text cannot be longer than 255 characters.", nameof(value));
        }

        ClearValue();
        Type = ExcelCellType.Text;
        _payload = value;
    }

    /// <summary>
    /// Makes the cell a formula cell.
    /// </summary>
    public void SetFormula(ExcelFormula formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ClearValue();
        _payload = formula;
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
        _payload = null;
        Type = ExcelCellType.Empty;
        Word0 = 0;
        Word1 &= 0x0FFF;
    }

    /// <summary>
    /// Points a number, text or formula cell at its data in the file buffer; it is decoded on first access.
    /// </summary>
    internal void LoadLazy(byte[] source, int dataOffset)
    {
        _payload = source;
        _dataOffset = dataOffset;
    }

    /// <summary>
    /// Gets or sets the number format index (0 = General; 0–63).
    /// </summary>
    public int NumberFormat
    {
        get => CellWords.NumberFormat(Word1);
        set
        {
            ArgumentOutOfRangeException.ThrowIfNegative(value);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(value, 63);
            Word1 = CellWords.WithNumberFormat(Word1, value);
        }
    }

    /// <summary>
    /// Gets or sets the alignment.
    /// </summary>
    public ExcelAlignment Alignment
    {
        get => CellWords.Alignment(Word1);
        set => Word1 = CellWords.WithAlignment(Word1, value);
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
        _ => Value switch
        {
            bool b => BooleanText.Of(b),
            ExcelError e => e.ToText(),
            var other => other?.ToString() ?? string.Empty,
        },
    };
}
