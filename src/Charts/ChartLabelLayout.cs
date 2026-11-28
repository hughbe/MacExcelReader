using MacExcelReader.Rendering;

namespace MacExcelReader.Charts;

/// <summary>
/// How many tick labels Excel 1.00 reckons fit along an axis (<c>37.0a</c>), worked back from the layout a chart was
/// saved with: the plot area Excel counts against is the saved one plus the room the axis labels and ticks then took
/// (<c>37.08</c>, <c>37.09</c>, <c>37.0e</c>; docs/Display.md, "Automatic axis scales").
/// </summary>
internal static class ChartLabelLayout
{
    private readonly record struct Extent(int Height, int Width);

    public static int GetLabelSpaces(ChartDocument document, Chart chart, MacFontSet fonts, bool xAxis)
    {
        bool scatter = chart.Type == ChartType.Scatter;
        var valueLabels = Measure(document, chart, fonts, chart.ValueAxis, chart.ValueScale, numeric: true);
        var categoryLabels = scatter && chart.XScale is { } x
            ? Measure(document, chart, fonts, chart.CategoryAxis, x, numeric: true)
            : Measure(document, chart, fonts, chart.CategoryAxis, null, numeric: false);

        // The value axis runs across in a bar chart; a scatter chart's X axis always does.
        bool horizontal = chart.IsBar ^ xAxis;
        var axis = xAxis ? chart.CategoryAxis : chart.ValueAxis;
        var (family, size) = axis.GetFont(document.Version);
        var font = GetFont(fonts, family, size);
        if (horizontal)
        {
            int width = Provisional(chart, vertical: false, valueLabels, categoryLabels, scatter);
            int zero = font.MeasureText("0", axis.IsBold) * 20;
            return width / (3 * zero);
        }

        var other = xAxis ? chart.ValueAxis : chart.CategoryAxis;
        var (otherFamily, otherSize) = other.GetFont(document.Version);
        var otherFont = GetFont(fonts, otherFamily, otherSize);
        int height = Provisional(chart, vertical: true, valueLabels, categoryLabels, scatter);
        return (height - 2 * (otherFont.Ascent + otherFont.Descent) * 20) / ((font.Ascent + font.Descent) * 20);
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
            int next = saved + Room(chart, vertical, valueLabels, categoryLabels, scatter, length, side: 0)
                + Room(chart, vertical, valueLabels, categoryLabels, scatter, length, side: 1);
            if (next == length)
            {
                break;
            }

            length = next;
        }

        return length;
    }

    // 37.0e: the room taken on one side of the plot area for the labels and ticks along it.
    private static int Room(Chart chart, bool vertical, Extent valueLabels, Extent categoryLabels, bool scatter, int length, int side)
    {
        int Size(Extent e) => vertical ? e.Height : e.Width;
        int margin = vertical ? 0 : 100;
        bool categoryHere = chart.IsBar ^ vertical;
        ChartAxis axis;
        int labels;
        int? overhang;
        bool reverse;
        if (categoryHere)
        {
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
        if (reverse ^ (side == 0))
        {
            if (axis.TickLabelPosition == TickLabelPosition.NextToAxis)
            {
                // How far the axis line is from this side, in proportion as it was saved.
                var box = chart.PlotArea.Box;
                int savedLength = vertical ? box.Height : box.Width;
                int savedDistance = axis.LineAt - (vertical ? box.Top : box.Left);
                int distance = savedLength == 0 ? 0 : (int)((long)savedDistance * length / savedLength);
                if (side == 1)
                {
                    distance = length - distance;
                }

                room = Math.Max(labels - distance, 0);
            }

            if (axis.TickLabelPosition == TickLabelPosition.Low)
            {
                room = labels;
            }
        }
        else if (axis.TickLabelPosition == TickLabelPosition.High)
        {
            room = labels;
        }

        if (overhang is { } other)
        {
            room = Math.Max(room, other / 2);
        }

        return room + 2 * axis.TickLength;
    }

    // 37.09: the height and width of an axis's largest label in twips, a pixel of margin round its text; nothing
    // for a hidden axis or one without labels. Category labels are taken on one line each.
    private static Extent Measure(ChartDocument document, Chart chart, MacFontSet fonts, ChartAxis axis, ChartScale? scale, bool numeric)
    {
        if (axis.Line.Pattern == ChartPattern.None || axis.TickLabelPosition == TickLabelPosition.None)
        {
            return default;
        }

        var (family, size) = axis.GetFont(document.Version);
        var font = GetFont(fonts, family, size);
        int width = 0;
        foreach (string text in numeric ? NumberLabels(scale!) : CategoryLabels(document, chart))
        {
            width = Math.Max(width, font.MeasureText(text, axis.IsBold));
        }

        return new Extent((font.Ascent + font.Descent + font.Leading + 2) * 20, (width + 2) * 20);
    }

    private static IEnumerable<string> NumberLabels(ChartScale scale)
    {
        double min = scale.Minimum, max = scale.Maximum, unit = scale.MajorUnit;
        if (!(unit > 0) || !(max >= min))
        {
            yield break;
        }

        int count = (int)Math.Min(500, Math.Round((max - min) / unit));
        for (int i = 0; i <= count; i++)
        {
            double value = min + i * unit;
            yield return ExcelValueFormatter.FormatGeneral(scale.IsLogarithmic ? ChartAutomaticScale.Pow10((int)Math.Round(value)) : value);
        }
    }

    private static IEnumerable<string> CategoryLabels(ChartDocument document, Chart chart)
    {
        var series = document.GetSeries(chart);
        if (series.Count == 0)
        {
            yield break;
        }

        var first = series[0];
        int step = Math.Max(1, (int)chart.CategoriesBetweenTickLabels);
        for (int i = 0; i < first.PointCount; i += step)
        {
            yield return first.SeriesType == 3 && i < first.CategoryTexts.Count ? first.CategoryTexts[i]
                : i < first.CategoryValues.Count ? ExcelValueFormatter.FormatGeneral(first.CategoryValues[i].ToDouble())
                : (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    // A font the System has no bitmap family for is drawn in the System font.
    private static MacBitmapFont GetFont(MacFontSet fonts, int family, int size)
        => fonts.GetFont(family, size) ?? fonts.GetFont(0, size) ?? fonts.GetFont(0, 12)
            ?? throw new ArgumentException("The font set has no System font (Chicago).", nameof(fonts));
}
