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

    // The second word as written: for a formula, with the type its current value gives.
    internal ushort TableWord1
    {
        get
        {
            _ = Type;
            return Word1;
        }
    }

    // The value: a string or an ExcelFormula; or, until first access, the file buffer
    // the number, text or formula record is read from at _dataOffset. Numbers are never cached, so
    // reading them does not allocate.
    private object? _payload;
    private int _dataOffset;

    // The value of a number cell set with SetNumber, kept out of _payload so that setting it does not box.
    private double _number;

    // An empty, locked cell. Cells belong to a grid, which creates them (ExcelCellGrid's indexer).
    internal ExcelCell()
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
        get
        {
            // A formula's type follows its current value, which ExcelFormula.SetResult can change after the formula
            // is set: type 5 for a number, 6 for anything else (spec §9, §12.4).
            if (_payload is ExcelFormula formula)
            {
                var type = formula.HasNumericResult ? ExcelCellType.NumericFormula : ExcelCellType.Formula;
                if (CellWords.Type(Word1) != type)
                {
                    Word1 = CellWords.WithType(Word1, type);
                }

                return type;
            }

            return CellWords.Type(Word1);
        }

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

    private double ReadNumber() => _payload is byte[] source ? ReadLazyDouble(source) : _number;

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
    /// Gets the cell's value: a number (integer and number cells), text, TRUE or FALSE, an error, a formula's
    /// current value, or <see cref="ExcelValue.Empty"/>. Reading it does not allocate, apart from decoding text.
    /// </summary>
    public ExcelValue Value => Type switch
    {
        ExcelCellType.Integer => ExcelValue.FromNumber(Word0),
        ExcelCellType.BooleanOrError => IsError ? ExcelValue.FromErrorCode((byte)Word0) : ExcelValue.FromBoolean(BooleanValue),
        ExcelCellType.Number => ExcelValue.FromNumber(ReadNumber()),
        ExcelCellType.Text => ExcelValue.FromText(DecodeText()),
        ExcelCellType.NumericFormula or ExcelCellType.Formula => _payload is byte[] source
            ? ExcelFormula.ReadResult(source.AsSpan(_dataOffset))
            : DecodeFormula().Result,
        _ => ExcelValue.Empty,
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
    /// <exception cref="ArgumentOutOfRangeException">The value is infinite or not a number, which Excel never
    /// stores: a formula that overflows gives <c>#NUM!</c> (spec §12.1).</exception>
    public void SetNumber(double value)
    {
        if (!double.IsFinite(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Excel numbers are finite: it gives #NUM! for an overflow.");
        }

        ClearValue();
        Type = ExcelCellType.Number;
        _number = value;
    }

    /// <summary>
    /// Sets a constant value as Excel stores one typed into the cell: a whole number from 0 to 65535 in the compact
    /// integer form, any other number (negative ones included) as a number, and text, TRUE, FALSE or an error as
    /// themselves. The empty value clears the cell, keeping its formatting.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is an infinite or not-a-number number.</exception>
    public void SetValue(ExcelValue value)
    {
        switch (value.Kind)
        {
            case ExcelValueKind.Number:
                double number = value.Number;
                if (number == Math.Floor(number) && number >= 0 && number <= ushort.MaxValue && !double.IsNegative(number))
                {
                    SetInteger((ushort)number);
                }
                else
                {
                    SetNumber(number);
                }

                break;
            case ExcelValueKind.Text:
                SetText(value.Text);
                break;
            case ExcelValueKind.Boolean:
                SetBoolean(value.Boolean);
                break;
            case ExcelValueKind.Error:
                SetError(value.Error);
                break;
            default:
                Clear();
                break;
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
        Type = formula.HasNumericResult ? ExcelCellType.NumericFormula : ExcelCellType.Formula;
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
        _number = 0;
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

    private bool GetFlag(int bit) => CellWords.Has(Word2, bit);

    private void SetFlag(int bit, bool value)
        => Word2 = (ushort)(value ? Word2 | (1 << bit) : Word2 & ~(1 << bit));

    /// <summary>Gets or sets whether the cell is hidden when the document is protected.</summary>
    public bool IsHidden { get => GetFlag(CellWords.HiddenBit); set => SetFlag(CellWords.HiddenBit, value); }

    /// <summary>Gets or sets whether the cell is locked when the document is protected.</summary>
    public bool IsLocked { get => GetFlag(CellWords.LockedBit); set => SetFlag(CellWords.LockedBit, value); }

    /// <summary>Gets or sets whether the cell's formula uses an array-valued function.</summary>
    public bool UsesArrayFunction { get => GetFlag(CellWords.ArrayFunctionBit); set => SetFlag(CellWords.ArrayFunctionBit, value); }

    /// <summary>Gets or sets the bottom border.</summary>
    public bool BottomBorder { get => GetFlag(CellWords.BottomBorderBit); set => SetFlag(CellWords.BottomBorderBit, value); }

    /// <summary>Gets or sets the top border.</summary>
    public bool TopBorder { get => GetFlag(CellWords.TopBorderBit); set => SetFlag(CellWords.TopBorderBit, value); }

    /// <summary>Gets or sets the right border.</summary>
    public bool RightBorder { get => GetFlag(CellWords.RightBorderBit); set => SetFlag(CellWords.RightBorderBit, value); }

    /// <summary>Gets or sets the left border.</summary>
    public bool LeftBorder { get => GetFlag(CellWords.LeftBorderBit); set => SetFlag(CellWords.LeftBorderBit, value); }

    /// <summary>Gets or sets italic.</summary>
    public bool Italic { get => GetFlag(CellWords.ItalicBit); set => SetFlag(CellWords.ItalicBit, value); }

    /// <summary>Gets or sets bold.</summary>
    public bool Bold { get => GetFlag(CellWords.BoldBit); set => SetFlag(CellWords.BoldBit, value); }

    /// <summary>
    /// Gets the raw attribute word (the third cell-table word).
    /// </summary>
    public ushort RawAttributes => Word2;

    /// <inheritdoc/>
    public override string ToString() => Type switch
    {
        ExcelCellType.Empty => string.Empty,
        _ => Value.ToString(),
    };
}
