using MacExcelReader.Rendering;

namespace MacExcelReader.Charts;

/// <summary>
/// How many tick labels Excel 1.00 reckons fit along an axis (<c>37.0a</c>), worked back from the layout a chart was
/// saved with: the plot area Excel counts against is the saved one plus the room the axis labels and ticks then took
/// (<c>37.08</c>, <c>37.09</c>, <c>37.0e</c>; docs/Display.md, "Automatic axis scales"). Lengths are in twips, 20 a
/// pixel.
/// </summary>
internal static class ChartLabelLayout
{
    private const int TwipsPerPixel = 20;

    // The size of an axis's largest label box.
    private readonly record struct Extent(int Height, int Width);

    // An axis's label font as Excel sets it for measuring (33.22).
    private readonly record struct LabelFont(MacBitmapFont Font, bool Bold)
    {
        public int Height => Font.Ascent + Font.Descent;

        public int LineHeight => Font.Ascent + Font.Descent + Font.Leading;

        public int Measure(ReadOnlySpan<char> text) => Font.MeasureText(text, Bold);
    }

    public static (int Height, int Width) GetTickLabelSize(ChartDocument document, Chart chart, MacFontSet fonts, bool category)
    {
        if (!category)
        {
            var e = MeasureNumbers(chart.ValueAxis, chart.ValueScale, GetFont(document, fonts, chart.ValueAxis));
            return (e.Height, e.Width);
        }

        var font = GetFont(document, fonts, chart.CategoryAxis);
        var extent = chart.Type == ChartType.Scatter && chart.XScale is { } x
            ? MeasureNumbers(chart.CategoryAxis, x, font)
            : MeasureCategories(document, chart, font);
        return (extent.Height, extent.Width);
    }

    public static int GetLabelSpaces(ChartDocument document, Chart chart, MacFontSet fonts, bool xAxis)
    {
        bool scatter = chart.Type == ChartType.Scatter;
        var valueFont = GetFont(document, fonts, chart.ValueAxis);
        var categoryFont = GetFont(document, fonts, chart.CategoryAxis);
        var valueLabels = MeasureNumbers(chart.ValueAxis, chart.ValueScale, valueFont);
        var categoryLabels = scatter && chart.XScale is { } x
            ? MeasureNumbers(chart.CategoryAxis, x, categoryFont)
            : MeasureCategories(document, chart, categoryFont);

        // The value axis runs across in a bar chart; a scatter chart's X axis always does.
        if (chart.IsBar ^ xAxis)
        {
            var font = xAxis ? categoryFont : valueFont;
            int width = Provisional(chart, vertical: false, valueLabels, categoryLabels, scatter);
            return width / (3 * font.Measure("0") * TwipsPerPixel);
        }

        var (own, other) = xAxis ? (categoryFont, valueFont) : (valueFont, categoryFont);
        int height = Provisional(chart, vertical: true, valueLabels, categoryLabels, scatter);
        return (height - 2 * other.Height * TwipsPerPixel) / (own.Height * TwipsPerPixel);
    }

    // The provisional plot length along one dimension: the saved length plus the room 37.0e took on each side, which
    // depends on where the axis crosses, so it is worked out until it settles.
    private static int Provisional(Chart chart, bool vertical, Extent valueLabels, Extent categoryLabels, bool scatter)
    {
        var box = chart.PlotArea.Box;
        int saved = vertical ? box.Height : box.Width;
        int length = saved;
        for (int i = 0; i < 8; i++)
        {
            int next = saved + Room(chart, vertical, valueLabels, categoryLabels, scatter, length, highSide: false)
                + Room(chart, vertical, valueLabels, categoryLabels, scatter, length, highSide: true);
            if (next == length)
            {
                break;
            }

            length = next;
        }

        return length;
    }

    // 37.0e: the room taken on one side of the plot area for the labels and ticks along it.
    private static int Room(Chart chart, bool vertical, Extent valueLabels, Extent categoryLabels, bool scatter, int length, bool highSide)
    {
        int margin = vertical ? 0 : 100;
        ChartAxis axis;
        int labels;
        int? overhang;
        bool reverse;
        if (chart.IsBar ^ vertical)
        {
            // The category axis's labels lie along these sides, the value axis's overhang the ends.
            axis = chart.CategoryAxis;
            labels = Size(categoryLabels) + margin;
            overhang = Size(valueLabels);
            reverse = chart.ValueScale.ValuesInReverseOrder;
        }
        else
        {
            axis = chart.ValueAxis;
            labels = Size(valueLabels) + margin;
            overhang = chart.ValueAxisCrossesBetweenCategories && !scatter ? null : Size(categoryLabels);
            reverse = scatter && chart.XScale is { } x ? x.ValuesInReverseOrder : chart.CategoriesInReverseOrder;
            if (axis.TickLabelPosition == TickLabelPosition.NextToAxis)
            {
                reverse ^= chart.ValueAxisCrossesAtMaximumCategory;
            }
        }

        int room = 0;
        var position = axis.TickLabelPosition;
        if (reverse == highSide)
        {
            if (position == TickLabelPosition.NextToAxis)
            {
                // The labels sit by the axis line, so they take only what reaches past this side.
                var box = chart.PlotArea.Box;
                int savedLength = vertical ? box.Height : box.Width;
                int savedDistance = axis.LineAt - (vertical ? box.Top : box.Left);
                int distance = savedLength == 0 ? 0 : (int)((long)savedDistance * length / savedLength);
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

        return room + 2 * axis.TickLength;

        int Size(Extent extent) => vertical ? extent.Height : extent.Width;
    }

    // 37.2a: number labels from the scale's minimum to its maximum in major units, in General, each on one line.
    private static Extent MeasureNumbers(ChartAxis axis, ChartScale scale, LabelFont font)
    {
        if (!HasLabels(axis))
        {
            return default;
        }

        double min = scale.Minimum, max = scale.Maximum, unit = scale.MajorUnit;
        int widest = 0;
        if (unit > 0 && max >= min)
        {
            Span<char> text = stackalloc char[ExcelValueFormatter.GeneralSize(20)];
            int count = (int)Math.Min(500, Math.Round((max - min) / unit));
            for (int i = 0; i <= count; i++)
            {
                double value = min + i * unit;
                if (scale.IsLogarithmic)
                {
                    value = ChartAutomaticScale.Pow10((int)Math.Round(value));
                }

                int length = ExcelValueFormatter.FormatGeneral(value, 20, text);
                widest = Math.Max(widest, font.Measure(text[..length]));
            }
        }

        return Box(lines: 1, widest, font);
    }

    // 37.2b: category labels, every so many as the axis says, each wrapped to its category's slot (33.1d).
    private static Extent MeasureCategories(ChartDocument document, Chart chart, LabelFont font)
    {
        var axis = chart.CategoryAxis;
        var series = document.SeriesOf(chart);
        if (!HasLabels(axis) || series.IsEmpty)
        {
            return default;
        }

        int step = Math.Max(1, (int)chart.CategoriesBetweenTickLabels);
        int slot = SlotWidth(chart, font, step) / TwipsPerPixel;
        var first = series[0];
        Span<char> number = stackalloc char[ExcelValueFormatter.GeneralSize(20)];
        int lines = 0, widest = 0;
        for (int i = 0; i < first.PointCount; i += step)
        {
            scoped ReadOnlySpan<char> text;
            if (first.SeriesType == 3)
            {
                text = i < first.CategoryTexts.Count ? first.CategoryTexts[i] : default;
            }
            else
            {
                double value = i < first.CategoryValues.Count ? first.CategoryValues[i].ToDouble() : i + 1;
                text = number[..ExcelValueFormatter.FormatGeneral(value, 20, number)];
            }

            var (labelLines, labelWidth) = Wrap(text, slot, font);
            lines = Math.Max(lines, labelLines);
            widest = Math.Max(widest, labelWidth);
        }

        return Box(lines, widest, font);
    }

    // The width a category label wraps to: a third of the chart area's width in a bar chart; otherwise the plot
    // area's width shared among the labels, but never less than a 0.
    private static int SlotWidth(Chart chart, LabelFont font, int step)
    {
        if (chart.IsBar)
        {
            return chart.ChartArea.Box.Width / 3;
        }

        int labels = Math.Max(1, ((chart.ValueAxisCrossesBetweenCategories ? 1 : 0) + chart.CategoryCount - 1) / step);
        return Math.Max(chart.PlotArea.Box.Width / labels, font.Measure("0") * TwipsPerPixel);
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

    // Font 1 is the application font, which the Font Manager draws in Geneva; a font the System has no bitmap family
    // for is drawn in the System font.
    private static LabelFont GetFont(ChartDocument document, MacFontSet fonts, ChartAxis axis)
    {
        var (family, size, bold, _) = document.GetLabelFont(axis);
        if (family == 1)
        {
            family = 3;
        }

        var font = fonts.GetFont(family, size) ?? fonts.GetFont(0, size) ?? fonts.GetFont(0, 12)
            ?? throw new ArgumentException("The font set has no System font (Chicago).", nameof(fonts));
        return new LabelFont(font, bold);
    }
}
