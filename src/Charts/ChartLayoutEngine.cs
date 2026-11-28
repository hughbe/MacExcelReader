using MacExcelReader.Rendering;

namespace MacExcelReader.Charts;

/// <summary>
/// How Excel 1.00 lays a chart out when it draws it (<c>37.03</c>–<c>37.10</c>; docs/Display.md, "Laying out a
/// chart"): the chart area inside the drawing, the border, title and ticks, the provisional plot area it counts
/// tick labels against (<c>37.0a</c>), and the room the labels then take (<c>37.09</c>, <c>37.0e</c>). It works either
/// forward from the chart area, or back from the plot area Excel saved. Lengths are in twips, 20 a pixel; vertical
/// coordinates run upwards, so a box's top field is its lower edge.
/// </summary>
internal sealed class ChartLayoutEngine
{
    private const int TwipsPerPixel = 20;

    private readonly ChartDocument _document;
    private readonly Chart _chart;
    private readonly MacFontSet _fonts;
    private readonly LabelFont _valueFont;
    private readonly LabelFont _categoryFont;
    private readonly bool _scatter;

    public ChartLayoutEngine(ChartDocument document, Chart chart, MacFontSet fonts)
    {
        _document = document;
        _chart = chart;
        _fonts = fonts;
        _valueFont = GetFont(chart.ValueAxis);
        _categoryFont = GetFont(chart.CategoryAxis);
        _scatter = chart.Type == ChartType.Scatter;
    }

    // A rectangle as 37.0f shrinks it: vertical coordinates upwards from the chart area's foot.
    private record struct Rect(int Top, int Left, int Bottom, int Right)
    {
        public readonly int Height => Top - Bottom;

        public readonly int Width => Right - Left;
    }

    // The size of an axis's largest label box.
    private readonly record struct Extent(int Height, int Width);

    // An axis's label font as Excel sets it for measuring (33.22).
    private readonly record struct LabelFont(MacBitmapFont Font, bool Bold)
    {
        public int Height => Font.Ascent + Font.Descent;

        public int LineHeight => Font.Ascent + Font.Descent + Font.Leading;

        public int Measure(ReadOnlySpan<char> text) => Font.MeasureText(text, Bold);
    }

    /// <summary>The layout Excel draws the chart with, worked forward from its chart area (or the default one).</summary>
    public ChartLayout LayOut()
    {
        var (area, top, left) = GetChartArea();
        var rect = new Rect(area.Height, 0, 0, area.Width);
        Border(ref rect, area);
        Title(ref rect, area);
        int tick = Math.Min(rect.Height, rect.Width) / 70;
        var provisional = rect;
        int valueSpaces = Count(category: false, provisional.Height, provisional.Width);
        int categorySpaces = Count(category: true, provisional.Height, provisional.Width);

        // 37.08: the scales are worked out first; the labels are measured from them in the plot area's width, and the
        // left and right taken off, then the bottom and top.
        var valueLabels = MeasureNumbers(_chart.ValueAxis, _chart.ValueScale, _valueFont, _chart.IsBar, provisional.Width, area);
        var categoryLabels = MeasureCategoryLabels(provisional.Width, area);
        int leftRoom = Room(vertical: false, valueLabels, categoryLabels, provisional.Width, tick, highSide: false);
        int rightRoom = Room(vertical: false, valueLabels, categoryLabels, provisional.Width, tick, highSide: true);
        Shrink(ref rect, area, leftRoom, vertical: false, highSide: false);
        Shrink(ref rect, area, rightRoom, vertical: false, highSide: true);
        // 37.09 measures again the labels along the top and bottom: a bar chart's value labels, otherwise the category
        // labels, in the plot area's new width.
        if (_chart.IsBar)
        {
            valueLabels = MeasureNumbers(_chart.ValueAxis, _chart.ValueScale, _valueFont, horizontal: true, rect.Width, area);
        }
        else
        {
            categoryLabels = MeasureCategoryLabels(rect.Width, area);
        }

        int bottomRoom = Room(vertical: true, valueLabels, categoryLabels, provisional.Height, tick, highSide: false);
        int topRoom = Room(vertical: true, valueLabels, categoryLabels, provisional.Height, tick, highSide: true);
        Shrink(ref rect, area, bottomRoom, vertical: true, highSide: false, raiseTop: true);
        Shrink(ref rect, area, topRoom, vertical: true, highSide: true);
        var plot = new ChartBox((short)rect.Bottom, (short)rect.Left, (short)rect.Height, (short)rect.Width);
        return new ChartLayout(area, top, left, plot, (short)tick, valueSpaces, categorySpaces);
    }

    /// <summary>The number of tick labels Excel reckoned fit, worked back from the plot area it saved.</summary>
    public int GetLabelSpacesFromSaved(bool category)
    {
        var area = _chart.ChartArea.Box;
        var valueLabels = MeasureNumbers(_chart.ValueAxis, _chart.ValueScale, _valueFont, _chart.IsBar, _chart.PlotArea.Box.Width, area);
        var categoryLabels = MeasureCategoryLabels(_chart.PlotArea.Box.Width, area);
        bool horizontal = _chart.IsBar ^ category;
        return Count(category,
            horizontal ? 0 : ProvisionalFromSaved(vertical: true, valueLabels, categoryLabels),
            horizontal ? ProvisionalFromSaved(vertical: false, valueLabels, categoryLabels) : 0);
    }

    /// <summary>The size of an axis's largest tick label box, as the chart was saved.</summary>
    public (int Height, int Width) GetTickLabelSize(bool category)
    {
        var extent = category
            ? MeasureCategoryLabels(_chart.PlotArea.Box.Width, _chart.ChartArea.Box)
            : MeasureNumbers(_chart.ValueAxis, _chart.ValueScale, _valueFont, _chart.IsBar, _chart.PlotArea.Box.Width, _chart.ChartArea.Box);
        return (extent.Height, extent.Width);
    }

    // 37.0a: along a vertical axis the height less two of the other axis's label heights over a label height; along a
    // horizontal one (a bar chart's value axis, a scatter chart's X axis) the width over three widths of 0.
    private int Count(bool category, int height, int width)
    {
        var (own, other) = category ? (_categoryFont, _valueFont) : (_valueFont, _categoryFont);
        return _chart.IsBar ^ category
            ? width / (3 * own.Measure("0") * TwipsPerPixel)
            : (height - 2 * other.Height * TwipsPerPixel) / (own.Height * TwipsPerPixel);
    }

    // 37.03: the saved chart area and margins, or for a chart Excel has not drawn, 15/16 of the drawing's width by
    // 23/25 of its height, centred.
    private (ChartBox Area, int Top, int Left) GetChartArea()
    {
        var box = _chart.ChartArea.Box;
        if (box.Height > 0 && box.Width > 0)
        {
            var (top, left) = _chart.DrawingMargins;
            return (box, top, left);
        }

        int width = _document.ChartWidth * 15 / 16, height = _document.ChartHeight * 23 / 25;
        return (new ChartBox(0, 0, (short)height, (short)width), (_document.ChartHeight - height) / 2, (_document.ChartWidth - width) / 2);
    }

    // 37.04: a border takes its weight plus two pixels off the left and top, and the right and bottom; a shadow a
    // fortieth of the chart area's smaller side more off the right and bottom (inferred: no sample has either).
    private static void Border(ChartAreaFormat format, ref Rect rect, ChartBox area)
    {
        bool border = format.Border.Pattern != ChartPattern.None;
        if (!border && !format.HasShadow)
        {
            return;
        }

        int thickness = ((border ? (int)format.Border.Weight : -1) + 2) * TwipsPerPixel;
        int across = thickness, down = thickness;
        if (border)
        {
            Shrink(ref rect, area, across, vertical: false, highSide: false);
            Shrink(ref rect, area, down, vertical: true, highSide: true);
        }

        if (format.HasShadow)
        {
            int shadow = Math.Min(area.Width, area.Height) / 40;
            across += shadow;
            down += shadow;
        }

        Shrink(ref rect, area, across, vertical: false, highSide: true);
        Shrink(ref rect, area, down, vertical: true, highSide: false);
    }

    private void Border(ref Rect rect, ChartBox area) => Border(_chart.ChartArea, ref rect, area);

    // 37.07: the title, laid out four-fifths of the chart area wide, takes its box and 5% of the chart area's height
    // off the top. With no title text of its own, a chart of one named series is titled with the series' name.
    private void Title(ref Rect rect, ChartBox area)
    {
        ChartText? title = null;
        foreach (var text in _chart.Texts)
        {
            if (text.Kind == ChartTextKind.ChartTitle)
            {
                title = text;
                break;
            }
        }

        ReadOnlySpan<char> content;
        LabelFont font;
        if (title is { IsAutomaticText: false })
        {
            content = title.Text;
            font = GetTextFont(title);
        }
        else
        {
            var series = _document.SeriesOf(_chart);
            if (series.Length != 1 || series[0].Name.Length == 0)
            {
                return;
            }

            // The automatic title's font is Geneva 9 in the samples (inferred).
            content = series[0].Name;
            font = title is null ? GetFont(3, 9, bold: false) : GetTextFont(title);
        }

        var (lines, _) = Wrap(content, area.Width * 4 / 5 / TwipsPerPixel, font);
        int box = (lines * font.LineHeight + 2) * TwipsPerPixel;
        Shrink(ref rect, area, box + area.Height * 5 / 100, vertical: true, highSide: true);
    }

    // The plot length along one dimension before the labels were laid out: the saved length plus the room 37.0e took
    // on each side, which depends on where the axis crosses, so it is worked out until it settles.
    private int ProvisionalFromSaved(bool vertical, Extent valueLabels, Extent categoryLabels)
    {
        var box = _chart.PlotArea.Box;
        int saved = vertical ? box.Height : box.Width;
        int length = saved;
        for (int i = 0; i < 8; i++)
        {
            int next = saved + Room(vertical, valueLabels, categoryLabels, length, tick: null, highSide: false)
                + Room(vertical, valueLabels, categoryLabels, length, tick: null, highSide: true);
            if (next == length)
            {
                break;
            }

            length = next;
        }

        return length;
    }

    // 37.0e: the room taken on one side of the plot area for the labels and ticks along it. With no tick length given,
    // the saved one is taken, and the axis's position as it was saved.
    private int Room(bool vertical, Extent valueLabels, Extent categoryLabels, int length, int? tick, bool highSide)
    {
        int margin = vertical ? 0 : 100;
        ChartAxis axis;
        int labels;
        int? overhang;
        bool reverse;
        bool categoryHere = _chart.IsBar ^ vertical;
        if (categoryHere)
        {
            // The category axis's labels lie along these sides, the value axis's overhang the ends.
            axis = _chart.CategoryAxis;
            labels = Size(categoryLabels) + margin;
            overhang = Size(valueLabels);
            reverse = _chart.ValueScale.ValuesInReverseOrder;
        }
        else
        {
            axis = _chart.ValueAxis;
            labels = Size(valueLabels) + margin;
            overhang = _chart.ValueAxisCrossesBetweenCategories && !_scatter ? null : Size(categoryLabels);
            reverse = _scatter && _chart.XScale is { } x ? x.ValuesInReverseOrder : _chart.CategoriesInReverseOrder;
            if (axis.TickLabelPosition == TickLabelPosition.NextToAxis)
            {
                reverse ^= _chart.ValueAxisCrossesAtMaximumCategory;
            }
        }

        int room = 0;
        var position = axis.TickLabelPosition;
        if (reverse == highSide)
        {
            if (position == TickLabelPosition.NextToAxis)
            {
                // The labels sit by the axis line, so they take only what reaches past this side.
                int distance = AxisDistance(axis, categoryHere, vertical, length, fromSaved: tick is null);
                room = Math.Max(labels - (highSide ? length - distance : distance), 0);
            }
            else if (position == TickLabelPosition.Low)
            {
                room = labels;
            }
        }
        else if (position == TickLabelPosition.High)
        {
            room = labels;
        }

        if (overhang is { } other)
        {
            room = Math.Max(room, other / 2);
        }

        return room + 2 * (tick ?? axis.TickLength);

        int Size(Extent extent) => vertical ? extent.Height : extent.Width;
    }

    // How far an axis line lies from the low side of the plot area: in proportion as it was saved, or worked out
    // from where the other axis's scale crosses (37.22).
    private int AxisDistance(ChartAxis axis, bool categoryHere, bool vertical, int length, bool fromSaved)
    {
        if (fromSaved)
        {
            var box = _chart.PlotArea.Box;
            int savedLength = vertical ? box.Height : box.Width;
            int savedDistance = axis.LineAt - (vertical ? box.Top : box.Left);
            return savedLength == 0 ? 0 : (int)((long)savedDistance * length / savedLength);
        }

        double fraction;
        if (categoryHere)
        {
            // The category axis lies where the value axis crosses its scale.
            fraction = Fraction(_chart.ValueScale);
        }
        else if (_scatter && _chart.XScale is { } x)
        {
            fraction = Fraction(x);
        }
        else if (_chart.ValueAxisCrossesAtMaximumCategory)
        {
            fraction = 1;
        }
        else
        {
            int categories = CategoryCount();
            int at = Math.Max(1, (int)_chart.ValueAxisCrossesAtCategory) - 1;
            int spans = _chart.ValueAxisCrossesBetweenCategories ? categories : categories - 1;
            fraction = spans <= 0 ? 0 : Math.Min(1, (double)at / spans);
            if (_chart.CategoriesInReverseOrder)
            {
                fraction = 1 - fraction;
            }
        }

        return (int)(fraction * length);

        static double Fraction(ChartScale scale)
        {
            double range = scale.Maximum - scale.Minimum;
            double f = range == 0 ? 0 : Math.Clamp((scale.CrossesAt - scale.Minimum) / range, 0, 1);
            return scale.ValuesInReverseOrder ? 1 - f : f;
        }
    }

    // 37.0f: takes room off one side of the rectangle, but never below a quarter of the chart area. Below the plot
    // area (raiseTop) the top first moves up as far as the chart area allows.
    private static void Shrink(ref Rect rect, ChartBox area, int amount, bool vertical, bool highSide, bool raiseTop = false)
    {
        if (vertical)
        {
            int least = area.Height / 4;
            if (highSide)
            {
                rect.Top -= amount;
                if (rect.Height < least)
                {
                    rect.Top = rect.Bottom + least;
                }

                return;
            }

            rect.Bottom += amount;
            if (rect.Height < least && raiseTop)
            {
                rect.Top = Math.Min(rect.Bottom + least, area.Height);
            }

            if (rect.Height < least)
            {
                rect.Bottom = rect.Top - least;
            }

            return;
        }

        int narrowest = area.Width / 4;
        if (highSide)
        {
            rect.Right -= amount;
            if (rect.Width < narrowest)
            {
                rect.Right = rect.Left + narrowest;
            }

            return;
        }

        rect.Left += amount;
        if (rect.Width < narrowest)
        {
            rect.Left = rect.Right - narrowest;
        }
    }

    // The number of categories: as Excel last counted them, or the most points of any series (37.13).
    private int CategoryCount()
    {
        if (_chart.CategoryCount > 0)
        {
            return _chart.CategoryCount;
        }

        int count = 0;
        foreach (var series in _document.SeriesOf(_chart))
        {
            count = Math.Max(count, series.PointCount);
        }

        return count;
    }

    // 37.2a, 37.2c: number labels. Excel formats the scale's minimum and maximum, and when the major unit is not
    // whole the minimum plus a unit and the maximum less one, in General with the room of the label's slot (at least
    // four 0s; 33.01, 19.0d), and lays the longest out in the slot. A horizontal axis shares the plot area's width
    // among its intervals; a vertical one has a quarter of the chart area's width. A logarithmic scale's labels are
    // the powers of ten at its ends.
    private Extent MeasureNumbers(ChartAxis axis, ChartScale scale, LabelFont font, bool horizontal, int plotWidth, ChartBox area)
    {
        if (!HasLabels(axis))
        {
            return default;
        }

        double min = scale.Minimum, max = scale.Maximum, unit = scale.MajorUnit;
        int slot;
        if (horizontal)
        {
            double intervals = scale.IsLogarithmic ? Math.Log10(ChartAutomaticScale.Pow10(Floor(max)) / ChartAutomaticScale.Pow10(Floor(min)))
                : unit == 0 ? 1 : (max - min) / unit;
            slot = plotWidth / Math.Max(1, (int)intervals);
        }
        else
        {
            slot = area.Width / 4;
        }

        if (scale.IsLogarithmic)
        {
            min = ChartAutomaticScale.Pow10(Floor(min));
            max = ChartAutomaticScale.Pow10(Floor(max));
        }

        return MeasureLongestNumber(font, slot, min, max, scale.IsLogarithmic ? 1 : unit);

        static int Floor(double value) => (int)Math.Floor(value);
    }

    // 37.2c: of the ends (and, for a unit that is not whole, the values a unit inside them), the label with the most
    // characters, laid out in the slot.
    private static Extent MeasureLongestNumber(LabelFont font, int slot, double min, double max, double unit)
    {
        int room = Math.Max(slot / TwipsPerPixel, 4 * font.Measure("0"));
        Span<char> best = stackalloc char[ExcelValueFormatter.GeneralSize(64)];
        Span<char> text = stackalloc char[ExcelValueFormatter.GeneralSize(64)];
        int bestLength = FormatInRoom(min, room, font, best);
        bool fractional = Math.Truncate(unit) - unit != 0;
        ReadOnlySpan<double> others = fractional ? [max, min + unit, max - unit] : [max];
        foreach (double value in others)
        {
            // A power of ten past the largest number has no label (a fixed logarithmic maximum of 500).
            if (!double.IsFinite(value))
            {
                continue;
            }

            int length = FormatInRoom(value, room, font, text);
            if (length > bestLength)
            {
                text[..length].CopyTo(best);
                bestLength = length;
            }
        }

        var (lines, width) = Wrap(best[..bestLength], slot / TwipsPerPixel, font);
        return Box(lines, width, font);
    }

    // General in a cell `room` pixels wide: as many characters as 0s fit, less a minus sign (19.0d).
    private static int FormatInRoom(double value, int room, LabelFont font, Span<char> destination)
    {
        int digits = (room - (value < 0 ? font.Measure("-") : 0)) / font.Measure("0");
        return ExcelValueFormatter.FormatGeneral(value, Math.Clamp(digits, 0, 64), destination);
    }

    // 37.2b: category labels, every so many as the axis says, each wrapped to its category's slot (33.1d); a scatter
    // chart's X axis has number labels.
    private Extent MeasureCategoryLabels(int plotWidth, ChartBox area)
    {
        if (_scatter && _chart.XScale is { } x)
        {
            return MeasureNumbers(_chart.CategoryAxis, x, _categoryFont, horizontal: !_chart.IsBar, plotWidth, area);
        }

        var series = _document.SeriesOf(_chart);
        if (!HasLabels(_chart.CategoryAxis) || series.IsEmpty)
        {
            return default;
        }

        int step = Math.Max(1, (int)_chart.CategoriesBetweenTickLabels);
        var first = series[0];
        if (first.SeriesType != 3)
        {
            // Number categories are labelled 1 to the number of categories, in slots of at least 210 twips.
            int labels = Math.Max(1, ((_chart.ValueAxisCrossesBetweenCategories ? 1 : 0) + CategoryCount() - 1) / step);
            int numberSlot = _chart.IsBar ? area.Width / 3 : Math.Max(plotWidth / labels, 210);
            return MeasureLongestNumber(_categoryFont, numberSlot, 1, CategoryCount() / step * step, 1);
        }

        int slot = SlotWidth(plotWidth, area, step) / TwipsPerPixel;
        int lines = 0, widest = 0;
        for (int i = 0; i < first.PointCount; i += step)
        {
            var text = i < first.CategoryTexts.Count ? first.CategoryTexts[i].AsSpan() : default;
            var (labelLines, labelWidth) = Wrap(text, slot, _categoryFont);
            lines = Math.Max(lines, labelLines);
            widest = Math.Max(widest, labelWidth);
        }

        return Box(lines, widest, _categoryFont);
    }

    // The width a category label wraps to: a third of the chart area's width in a bar chart; otherwise the plot
    // area's width shared among the labels, but never less than a 0.
    private int SlotWidth(int plotWidth, ChartBox area, int step)
    {
        if (_chart.IsBar)
        {
            return area.Width / 3;
        }

        int labels = Math.Max(1, ((_chart.ValueAxisCrossesBetweenCategories ? 1 : 0) + CategoryCount() - 1) / step);
        return Math.Max(plotWidth / labels, _categoryFont.Measure("0") * TwipsPerPixel);
    }

    // 33.1e and 33.1f: lays text out in lines no wider than `width` pixels, a word at a time (a word runs to and
    // includes the next space or hyphen), breaking a first word too wide for its line between characters. A return
    // or line feed ends a line. Gives the number of lines and the widest.
    private static (int Lines, int Width) Wrap(ReadOnlySpan<char> text, int width, LabelFont font)
    {
        int lines = 0, widest = 0, start = 0;
        while (start < text.Length)
        {
            int used = 0, words = 0, at = start;
            while (at < text.Length)
            {
                int length = WordLength(text[at..]);
                int wordWidth = font.Measure(text.Slice(at, length));
                if (used + wordWidth <= width)
                {
                    used += wordWidth;
                    at += length;
                    words++;
                    if (at < text.Length && text[at] is '\r' or '\n')
                    {
                        at++;
                        break;
                    }

                    continue;
                }

                if (words == 0)
                {
                    // A word wider than the line: as many characters as fit, and at least one.
                    do
                    {
                        int charWidth = font.Measure(text.Slice(at, 1));
                        if (used + charWidth > width)
                        {
                            if (used == 0)
                            {
                                used = width;
                                at++;
                            }

                            break;
                        }

                        used += charWidth;
                        at++;
                    }
                    while (at < text.Length);
                }

                break;
            }

            lines++;
            widest = Math.Max(widest, used);
            start = at;
        }

        return (lines, widest);
    }

    // 33.21: the characters up to and including the next space or hyphen; a return ends a word before it.
    private static int WordLength(ReadOnlySpan<char> text)
    {
        int count = 1;
        while (count < text.Length)
        {
            char c = text[count - 1];
            if (c is ' ' or '-')
            {
                break;
            }

            if (c == '\r')
            {
                return count - 1;
            }

            count++;
        }

        return count;
    }

    // 33.1d: a label's box is its lines and widest line with a pixel of margin all round.
    private static Extent Box(int lines, int width, LabelFont font)
        => new((lines * font.LineHeight + 2) * TwipsPerPixel, (width + 2) * TwipsPerPixel);

    // 37.09: a hidden axis, or one without labels, takes no room.
    private static bool HasLabels(ChartAxis axis) => axis.Line.Pattern != ChartPattern.None && axis.TickLabelPosition != TickLabelPosition.None;

    private LabelFont GetFont(ChartAxis axis)
    {
        var font = _document.GetLabelFont(axis);
        return GetFont(font.Family, font.Size, font.Bold);
    }

    // A text object's font: its saved number and size, the System font at 12 when they are 0 (inferred).
    private LabelFont GetTextFont(ChartText text) => GetFont(text.Record[20], text.Record[21] == 0 ? 12 : text.Record[21], bold: false);

    // Font 1 is the application font, which the Font Manager draws in Geneva; a font the System has no bitmap family
    // for is drawn in the System font.
    private LabelFont GetFont(int family, int size, bool bold)
    {
        if (family == 1)
        {
            family = 3;
        }

        var font = _fonts.GetFont(family, size) ?? _fonts.GetFont(0, size) ?? _fonts.GetFont(0, 12)
            ?? throw new ArgumentException("The font set has no System font (Chicago).", "fonts");
        return new LabelFont(font, bold);
    }
}

/// <summary>
/// The layout Excel 1.00 draws a chart with, in twips: the chart area inside the drawing and its margins, the plot area
/// (chart coordinates run upwards, so its <see cref="ChartBox.Top"/> is its lower edge), the axes' tick length, and how
/// many tick labels Excel reckons fit along the value axis and the category (scatter X) axis.
/// </summary>
public readonly record struct ChartLayout(ChartBox ChartArea, int TopMargin, int LeftMargin, ChartBox PlotArea, short TickLength, int ValueLabelSpaces, int CategoryLabelSpaces);
