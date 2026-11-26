using MacExcelReader.Internal;

namespace MacExcelReader.Rendering;

/// <summary>
/// Draws a worksheet or macro sheet's window as Excel 1.00 shows it on a Macintosh screen: the column and row
/// headings, gridlines, the cells' values laid out as spec Appendix D gives them, borders, the selection, the scroll
/// bars with Excel's split boxes, and optionally the window's frame. Only the visible part of the sheet is laid out,
/// so drawing costs the same for a sheet of any size.
/// </summary>
public static class ExcelRenderer
{
    // The largest window taken from a file's rectangle, and the largest that can be asked for: QuickDraw's
    // coordinates are 16-bit.
    private const int MaxSavedSize = 2048;
    private const int MaxSize = 16384;

    // QuickDraw's gray and ltGray patterns, a byte per row from the top, the top bit leftmost.
    private const ulong Gray = 0xAA55AA55AA55AA55;
    private const ulong LightGray = 0x8822882288228822;

    // The scroll bars' parts, 16 pixels square with their borders, as the System's control definition draws them.
    private static readonly ushort[] s_upArrow = Art(
        "################",
        "#..............#",
        "#......#.......#",
        "#.....#.#......#",
        "#....#...#.....#",
        "#...#.....#....#",
        "#..#.......#...#",
        "#.#.........#..#",
        "#####.....####.#",
        "#...#.....#....#",
        "#...#.....#....#",
        "#...#.....#....#",
        "#...#######....#",
        "#..............#",
        "#..............#",
        "################");

    private static readonly ushort[] s_leftArrow = Art(
        "################",
        "#......#.......#",
        "#.....##.......#",
        "#....#.#.......#",
        "#...#..#####...#",
        "#..#.......#...#",
        "#.#........#...#",
        "##.........#...#",
        "#.#........#...#",
        "#..#.......#...#",
        "#...#..#####...#",
        "#....#.#.......#",
        "#.....##.......#",
        "#......#.......#",
        "#..............#",
        "################");

    private static readonly ushort[] s_growBox = Art(
        "################",
        "#..............#",
        "#..............#",
        "#..#######.....#",
        "#..#.....#.....#",
        "#..#.....#####.#",
        "#..#.....#...#.#",
        "#..#.....#...#.#",
        "#..#.....#...#.#",
        "#..#######...#.#",
        "#....#.......#.#",
        "#....#.......#.#",
        "#....#.......#.#",
        "#....#########.#",
        "#..............#",
        "################");

    private static readonly ushort[] s_downArrow = [.. s_upArrow.Reverse()];
    private static readonly ushort[] s_rightArrow = [.. s_leftArrow.Select(Mirror)];

    /// <summary>
    /// Draws one of a document's windows.
    /// </summary>
    /// <param name="document">The worksheet or macro sheet.</param>
    /// <param name="fonts">The System's fonts, which must include the sheet's standard font family.</param>
    /// <param name="options">What to draw; by default the front window, framed.</param>
    /// <exception cref="InvalidOperationException">The document has no such window, or the fonts lack the sheet's
    /// standard font family.</exception>
    public static MonoBitmap Render(ExcelDocument document, MacFontSet fonts, ExcelRenderOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(fonts);
        options ??= new ExcelRenderOptions();
        if ((uint)options.Window >= (uint)document.Windows.Count)
        {
            throw new InvalidOperationException($"The document has no window {options.Window}.");
        }

        var window = document.Windows[options.Window];

        // Excel fits a window's saved rectangle to the screen when it opens the file (spec §13), so a damaged one
        // is drawn at most screen-sized.
        int width = options.Width ?? Math.Clamp(window.Right - window.Left, 64, MaxSavedSize);
        int height = options.Height ?? Math.Clamp(window.Bottom - window.Top, 64, MaxSavedSize);
        ArgumentOutOfRangeException.ThrowIfLessThan(width, 64, nameof(options));
        ArgumentOutOfRangeException.ThrowIfLessThan(height, 64, nameof(options));
        ArgumentOutOfRangeException.ThrowIfGreaterThan(width, MaxSize, nameof(options));
        ArgumentOutOfRangeException.ThrowIfGreaterThan(height, MaxSize, nameof(options));

        var state = document.State;
        var font = fonts.GetFont(state.StandardFont, state.StandardFontSize)
            ?? throw new InvalidOperationException($"The fonts have no font of family {state.StandardFont}, the sheet's standard font.");

        // The frame: a line round the contents, 18 pixels of title bar above, and a one-pixel shadow right and below.
        int x0 = options.Frame ? 1 : 0, y0 = options.Frame ? 19 : 0;
        var bitmap = new MonoBitmap(width + (options.Frame ? 3 : 0), height + (options.Frame ? 21 : 0));
        var view = new View(bitmap, x0, y0, width, height);
        var sheet = new SheetLayout(document, window, font, view, options);
        sheet.Draw();
        if (options.Frame)
        {
            DrawFrame(bitmap, width, height, options.Title ?? "Worksheet", fonts.GetFont(0, 12) ?? font);
        }

        return bitmap;
    }

    // The window's contents in the bitmap: an origin and a size, local coordinates measured from the origin as
    // QuickDraw's window coordinates are. Patterns line up with the origin.
    private readonly record struct View(MonoBitmap Bitmap, int X, int Y, int Width, int Height)
    {
        public void Fill(int left, int top, int right, int bottom, bool black)
            => Bitmap.FillRect(X + Math.Max(left, -X), Y + Math.Max(top, -Y), X + right, Y + bottom, black);

        public void Pattern(int left, int top, int right, int bottom, ulong pattern, int phase = 0)
            => Bitmap.FillPattern(X + left, Y + top, X + right, Y + bottom, pattern, -X, phase - Y);

        public void Invert(int left, int top, int right, int bottom) => Bitmap.InvertRect(X + left, Y + top, X + right, Y + bottom);

        public void Art(ushort[] art, int left, int top)
        {
            for (int row = 0; row < art.Length; row++)
            {
                Bitmap.FillRect(X + left, Y + top + row, X + left + 16, Y + top + row + 1, false);
                Bitmap.OrBits(X + left, Y + top + row, (ulong)art[row] << 48, 16, 0, Bitmap.Width);
            }
        }

        public int Text(MacBitmapFont font, int x, int baseline, ReadOnlySpan<char> text, bool bold, bool italic, int clipLeft, int clipTop, int clipRight, int clipBottom)
            => font.DrawText(Bitmap, X + x, Y + baseline, text, bold, italic, X + clipLeft, Y + clipTop, X + clipRight, Y + clipBottom);
    }

    // Everything about the sheet's part of the window: which rows and columns show, and where.
    private sealed class SheetLayout
    {
        private readonly ExcelDocument _document;
        private readonly ExcelWindow _window;
        private readonly MacBitmapFont _font;
        private readonly View _view;
        private readonly ExcelRenderOptions _options;
        private readonly ExcelSheetState _state;
        private readonly bool _formulas;
        private readonly int _rowHeight;
        private readonly int _ascent;

        // The pane's pattern phase (window state offset 19, spec §13): Excel rotates its pattern down by this many rows
        // before drawing in the pane.
        private readonly int _phase;

        // The cell area: right of the row headings, below the column headings, and left of and above the scroll bars.
        private readonly int _left, _top, _right, _bottom;

        // The visible columns and rows, with the left edge of each column and the top of each row.
        private readonly List<(int Column, int Left, int Width)> _columns = [];
        private readonly List<(int Row, int Top)> _rows = [];

        // Each column's left edge in pixels from column A, from the state block's widths (spec Appendix D).
        private readonly int[] _sheetLefts = new int[ExcelCellGrid.MaxColumns + 2];

        public SheetLayout(ExcelDocument document, ExcelWindow window, MacBitmapFont font, View view, ExcelRenderOptions options)
        {
            _document = document;
            _window = window;
            _font = font;
            _view = view;
            _options = options;
            _state = document.State;
            _formulas = window.DisplayFormulas;
            _phase = window.State[19] & 7;
            var metrics = MacFont.Find(_state.StandardFont, _state.StandardFontSize);
            _rowHeight = metrics?.RowHeight ?? Math.Max(1, (int)_state.CachedRowHeight);
            _ascent = metrics?.Ascent ?? font.Ascent;

            bool headings = window.DisplayHeadings;
            _left = headings ? window.RowHeadingWidth : 0;
            _top = headings ? _rowHeight : 0;
            _right = view.Width - 15;
            _bottom = view.Height - 15;

            for (int c = 0; c <= ExcelCellGrid.MaxColumns; c++)
            {
                _sheetLefts[c + 1] = _sheetLefts[c] + (c < ExcelCellGrid.MaxColumns ? _state.GetColumnWidth(c) : 0);
            }

            // Excel shows formulas in columns twice as wide (spec §13).
            int x = _left;
            for (int c = window.LeftColumn; c < ExcelCellGrid.MaxColumns && x < _right; c++)
            {
                int w = _state.GetColumnWidth(c) * (_formulas ? 2 : 1);
                if (w > 0)
                {
                    _columns.Add((c, x, w));
                    x += w;
                }
            }

            for (int r = window.TopRow, y = _top; r < ExcelCellGrid.MaxRows && y < _bottom; r++, y += _rowHeight)
            {
                _rows.Add((r, y));
            }
        }

        public void Draw()
        {
            _view.Fill(0, 0, _view.Width, _view.Height, false);
            if (_window.DisplayGridlines)
            {
                DrawGridlines();
            }

            DrawCells();
            DrawBorders();
            DrawSheetEdges();
            if (_window.DisplayHeadings)
            {
                DrawHeadings();
            }

            if (_options.Selection)
            {
                DrawSelection();
            }

            DrawPastSheet();

            DrawScrollBars();
        }

        private void DrawGridlines()
        {
            // Dotted in QuickDraw's gray, along each column's right edge and each row's bottom edge.
            foreach (var (_, left, width) in _columns)
            {
                _view.Pattern(left + width - 1, _top, left + width, _bottom, Gray, _phase);
            }

            foreach (var (_, top) in _rows)
            {
                _view.Pattern(_left, top + _rowHeight - 1, _right, top + _rowHeight, Gray, _phase);
            }
        }

        // The sheet's last row and column end in a solid line instead of a dotted gridline.
        private void DrawSheetEdges()
        {
            if (_columns.Count > 0 && _columns[^1].Column == ExcelCellGrid.MaxColumns - 1)
            {
                int edge = _columns[^1].Left + _columns[^1].Width - 1;
                _view.Fill(edge, _top, edge + 1, _bottom, true);
            }

            if (_rows.Count > 0 && _rows[^1].Row == ExcelCellGrid.MaxRows - 1)
            {
                int edge = _rows[^1].Top + _rowHeight - 1;
                _view.Fill(_left, edge, _right, edge + 1, true);
            }
        }

        // Below row 16384 and right of column IV there is no sheet: Excel fills the rest of the window with grey.
        private void DrawPastSheet()
        {
            if (_columns.Count > 0 && _columns[^1].Column == ExcelCellGrid.MaxColumns - 1)
            {
                int end = _columns[^1].Left + _columns[^1].Width;
                _view.Pattern(end, 0, _right, _bottom, Gray, _phase);
            }

            if (_rows.Count > 0 && _rows[^1].Row == ExcelCellGrid.MaxRows - 1)
            {
                _view.Pattern(0, _rows[^1].Top + _rowHeight, _right, _bottom, Gray, _phase);
            }
        }

        private void DrawHeadings()
        {
            // The column headings: a bold letter centred over each column, solid lines between them and below.
            Span<char> name = stackalloc char[2];
            foreach (var (column, left, width) in _columns)
            {
                int length = ColumnName(column, name);
                CenterText(name[..length], left, width - 1, 0);
                _view.Fill(left + width - 1, 0, left + width, _top, true);
            }

            _view.Fill(0, _top - 1, _right, _top, true);

            // The row headings: a bold number centred in the strip, solid lines between them and to the right.
            Span<char> number = stackalloc char[5];
            foreach (var (row, top) in _rows)
            {
                (row + 1).TryFormat(number, out int length);
                CenterText(number[..length], 0, _left - 1, top);
                _view.Fill(0, top + _rowHeight - 1, _left, top + _rowHeight, true);
            }

            _view.Fill(_left - 1, 0, _left, _bottom, true);
        }

        private void CenterText(ReadOnlySpan<char> text, int left, int width, int top)
        {
            int textWidth = _font.MeasureText(text, bold: true);
            int x = left + (width - textWidth) / 2;
            _view.Text(_font, x, top + _ascent, text, true, false, left, top, left + width, top + _rowHeight - 1);
        }

        private static int ColumnName(int column, Span<char> name)
        {
            if (column < 26)
            {
                name[0] = (char)('A' + column);
                return 1;
            }

            name[0] = (char)('A' + column / 26 - 1);
            name[1] = (char)('A' + column % 26);
            return 2;
        }

        private void DrawCells()
        {
            var grid = _document.Cells;
            int scroll = _sheetLefts[_window.LeftColumn];
            foreach (var (row, top) in _rows)
            {
                if (row >= grid.RowCount || _columns.Count == 0)
                {
                    continue;
                }

                if (_formulas)
                {
                    foreach (var (column, left, width) in _columns)
                    {
                        DrawFormulaCell(row, column, left, width, top);
                    }

                    continue;
                }

                // Text can run in from a cell either side of the visible columns, so lay out the nearest cell with a
                // value beyond each edge too (spec Appendix D: text running into the next cells).
                int first = _columns[0].Column, last = _columns[^1].Column;
                int before = first - 1;
                while (before >= 0 && grid.GetCellType(row, before) == ExcelCellType.Empty)
                {
                    before--;
                }

                int after = last + 1;
                while (after < grid.ColumnCount && grid.GetCellType(row, after) == ExcelCellType.Empty)
                {
                    after++;
                }

                if (before >= 0)
                {
                    DrawCell(row, before, top, scroll);
                }

                foreach (var (column, _, _) in _columns)
                {
                    DrawCell(row, column, top, scroll);
                }

                if (after < grid.ColumnCount)
                {
                    DrawCell(row, after, top, scroll);
                }
            }
        }

        private void DrawCell(int row, int column, int top, int scroll)
        {
            if (ExcelValueFormatter.Layout(_document, row, column, _sheetLefts[column], _options.Version) is not { } display)
            {
                return;
            }

            _document.Cells.GetWords(row, column, out _, out _, out ushort attributes);
            bool bold = CellWords.Has(attributes, CellWords.BoldBit);
            bool italic = CellWords.Has(attributes, CellWords.ItalicBit);
            int offset = _left - scroll;
            int clipLeft = Math.Max(_left, display.ClipLeft + offset);
            int clipRight = Math.Min(_right, display.ClipRight + offset);

            // Excel clears the stretch the text is clipped to first, so text running on into its neighbours covers the
            // gridlines between them.
            _view.Fill(clipLeft, top, clipRight, top + _rowHeight - 1, false);
            _view.Text(_font, display.X + offset, top + _ascent, display.Text, bold, italic, clipLeft, top, clipRight, top + _rowHeight - 1);
        }

        // With formulas shown, every cell shows what was typed into it, from the left of its own cell.
        private void DrawFormulaCell(int row, int column, int left, int width, int top)
        {
            var grid = _document.Cells;
            if (grid.GetCellType(row, column) == ExcelCellType.Empty)
            {
                return;
            }

            grid.GetWords(row, column, out ushort word0, out ushort word1, out ushort word2);
            var cell = new ExcelCellRef(grid, row, column, word0, word1, word2);
            string text = cell.Type is ExcelCellType.Formula or ExcelCellType.NumericFormula
                ? cell.GetCell().Formula!.ToText(_document)
                : cell.Value.Kind == ExcelValueKind.Number ? ExcelValueFormatter.FormatGeneral(cell.Value.Number) : cell.Value.ToString();
            _view.Text(_font, left + 2, top + _ascent, text, cell.Bold, cell.Italic, left, top, left + width - 1, top + _rowHeight - 1);
        }

        private void DrawBorders()
        {
            var grid = _document.Cells;
            foreach (var (row, top) in _rows)
            {
                if (row >= grid.RowCount)
                {
                    break;
                }

                foreach (var (column, left, width) in _columns)
                {
                    if (column >= grid.ColumnCount)
                    {
                        break;
                    }

                    grid.GetWords(row, column, out _, out _, out ushort attributes);
                    if ((attributes & 0x3C) == 0)
                    {
                        continue;
                    }

                    int right = left + width - 1, bottom = top + _rowHeight - 1;
                    if (CellWords.Has(attributes, CellWords.LeftBorderBit))
                    {
                        _view.Fill(left - 1, top - 1, left, bottom + 1, true);
                    }

                    if (CellWords.Has(attributes, CellWords.RightBorderBit))
                    {
                        _view.Fill(right, top - 1, right + 1, bottom + 1, true);
                    }

                    if (CellWords.Has(attributes, CellWords.TopBorderBit))
                    {
                        _view.Fill(left - 1, top - 1, right + 1, top, true);
                    }

                    if (CellWords.Has(attributes, CellWords.BottomBorderBit))
                    {
                        _view.Fill(left - 1, bottom, right + 1, bottom + 1, true);
                    }
                }
            }
        }

        private void DrawSelection()
        {
            // Each selected area is inverted from the pixel outside the gridline before it to the pixel outside the
            // gridline after it. The active cell is then inverted back inside a one-pixel frame, so it shows as it is
            // with an inverted outline.
            foreach (var range in _window.Selection)
            {
                if (Bounds(range) is { } r)
                {
                    _view.Invert(Math.Max(r.Left - 2, _left), Math.Max(r.Top - 2, _top), Math.Min(r.Right + 1, _right), Math.Min(r.Bottom + 1, _bottom));
                }
            }

            if (Bounds(new ExcelRange(_window.ActiveRow, _window.ActiveColumn, _window.ActiveRow, _window.ActiveColumn)) is { } active)
            {
                _view.Invert(Math.Max(active.Left + 1, _left), Math.Max(active.Top + 1, _top), Math.Min(active.Right - 2, _right), Math.Min(active.Bottom - 2, _bottom));
            }
        }

        // The pixels a range covers in the cell area, from its first cell's left and top edges to just past the
        // gridlines after its last cell, or null if none of it shows.
        private (int Left, int Top, int Right, int Bottom)? Bounds(ExcelRange range)
        {
            int left = int.MaxValue, right = int.MinValue, top = int.MaxValue, bottom = int.MinValue;
            foreach (var (column, x, width) in _columns)
            {
                if (column >= range.FirstColumn && column <= range.LastColumn)
                {
                    left = Math.Min(left, x);
                    right = Math.Max(right, x + width);
                }
            }

            foreach (var (row, y) in _rows)
            {
                if (row >= range.FirstRow && row <= range.LastRow)
                {
                    top = Math.Min(top, y);
                    bottom = Math.Max(bottom, y + _rowHeight);
                }
            }

            return left > right || top > bottom ? null : (left, top, right, bottom);
        }

        private void DrawScrollBars()
        {
            int w = _view.Width, h = _view.Height;

            // Vertical: Excel's black split box, the up arrow, the thumb, a light grey track and the down arrow.
            _view.Fill(w - 15, -1, w + 1, h - 14, true);
            _view.Fill(w - 14, 4, w, h - 15, false);
            _view.Pattern(w - 14, 20, w, h - 30, LightGray);
            _view.Art(s_upArrow, w - 15, 4);
            int down = VerticalThumbOffset(_window.TopRow, h - 30 - 20 - 16);
            Thumb(w - 14, 20 + down, w, 36 + down);
            _view.Art(s_downArrow, w - 15, h - 30);

            // Horizontal, the same on its side.
            _view.Fill(-1, h - 15, w - 14, h + 1, true);
            _view.Fill(4, h - 14, w - 15, h, false);
            _view.Pattern(20, h - 14, w - 30, h, LightGray);
            _view.Art(s_leftArrow, 4, h - 15);
            int across = HorizontalThumbOffset(_window.LeftColumn, w - 30 - 20 - 16);
            Thumb(20 + across, h - 14, 36 + across, h);
            _view.Art(s_rightArrow, w - 30, h - 15);

            _view.Art(s_growBox, w - 15, h - 15);
        }

        // The thumb: a white square inside the track, at its start for a window scrolled to the top or left.
        private void Thumb(int left, int top, int right, int bottom)
        {
            _view.Fill(left, top, right, bottom, true);
            _view.Fill(left + 1, top + 1, right - 1, bottom - 1, false);
        }
    }

    // The vertical scroll bar's value for a top row (18.12 in Excel 1.00): 0 to 17982, with the first and last 20
    // rows taking 819 each and the rows between sharing the rest, so that the thumb moves for every one of the first
    // rows.
    internal const int VerticalScrollMax = 17982;

    internal static int VerticalScrollValue(int topRow)
    {
        if (topRow < 20)
        {
            return (topRow * 819 + 10) / 20;
        }

        if (topRow < 16363)
        {
            return (int)(((long)(topRow - 20) * 16344 + 8171) / 16343) + 819;
        }

        return ((topRow - 16363) * 819 + 10) / 20 + 17163;
    }

    // How far the vertical thumb sits from the top of its travel: the Control Manager's position for the top row's
    // scroll value out of 17982.
    internal static int VerticalThumbOffset(int topRow, int travel) =>
        ThumbOffset(VerticalScrollValue(Math.Clamp(topRow, 0, ExcelCellGrid.MaxRows - 1)), VerticalScrollMax, travel);

    // How far the horizontal thumb sits from the left of its travel: the Control Manager's position for the left
    // column out of 255.
    internal static int HorizontalThumbOffset(int leftColumn, int travel) => ThumbOffset(Math.Clamp(leftColumn, 0, 255), 255, travel);

    // The System's scroll bar (CDEF 1) places the thumb value × travel ÷ max pixels along its travel, rounding up only
    // when the remainder is more than half of max.
    private static int ThumbOffset(int value, int max, int travel)
    {
        if (travel <= 0)
        {
            return 0;
        }

        int product = value * travel;
        int offset = product / max;
        return product % max > max / 2 ? offset + 1 : offset;
    }

    private static void DrawFrame(MonoBitmap bitmap, int width, int height, string title, MacBitmapFont titleFont)
    {
        int right = width + 1, bottom = height + 19;

        // The border, the title bar's bottom line and the shadow.
        bitmap.FillRect(0, 0, right + 1, 1, true);
        bitmap.FillRect(0, 0, 1, bottom + 1, true);
        bitmap.FillRect(right, 0, right + 1, bottom + 1, true);
        bitmap.FillRect(0, bottom, right + 1, bottom + 1, true);
        bitmap.FillRect(0, 18, right + 1, 19, true);
        bitmap.FillRect(right + 1, 1, right + 2, bottom + 2, true);
        bitmap.FillRect(1, bottom + 1, right + 2, bottom + 2, true);

        // The active window's stripes, broken round the close box and the title.
        bitmap.FillRect(1, 1, right, 18, false);
        for (int y = 4; y <= 14; y += 2)
        {
            bitmap.FillRect(2, y, right - 1, y + 1, true);
        }

        bitmap.FillRect(8, 3, 21, 16, false);
        bitmap.FillRect(9, 4, 20, 15, true);
        bitmap.FillRect(10, 5, 19, 14, false);

        int titleWidth = titleFont.MeasureText(title);
        int titleLeft = (right + 1 - titleWidth) / 2;
        bitmap.FillRect(titleLeft - 6, 3, titleLeft + titleWidth + 6, 16, false);
        titleFont.DrawText(bitmap, titleLeft, 14, title, false, false, titleLeft - 6, 1, titleLeft + titleWidth + 6, 18);
    }

    private static ushort[] Art(params string[] rows)
        => [.. rows.Select(r => (ushort)r.Select((c, i) => c == '#' ? 1 << (15 - i) : 0).Sum())];

    private static ushort Mirror(ushort row)
    {
        int mirrored = 0;
        for (int i = 0; i < 16; i++)
        {
            if ((row & (1 << i)) != 0)
            {
                mirrored |= 1 << (15 - i);
            }
        }

        return (ushort)mirrored;
    }
}
