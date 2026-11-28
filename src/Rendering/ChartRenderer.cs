using MacExcelReader.Charts;

namespace MacExcelReader.Rendering;

/// <summary>
/// Draws a chart's window as Excel 1.00 shows it on a Macintosh screen (docs/Display.md, "Drawing a chart"): the frame,
/// the axes and their ticks, labels and titles, the legend and the series, laid out as
/// <see cref="ChartDocument.LayOut"/> gives.
/// </summary>
public static class ChartRenderer
{
    private const int TwipsPerPixel = 20;

    /// <summary>
    /// Draws a chart document's window at the size it was saved, as Excel 1.00 draws it.
    /// </summary>
    /// <param name="document">The chart.</param>
    /// <param name="fonts">The System's fonts.</param>
    /// <param name="frame">Whether to draw the window's frame and title bar.</param>
    /// <param name="title">The window's title.</param>
    public static MonoBitmap Render(ChartDocument document, MacFontSet fonts, bool frame = true, string? title = null)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(fonts);
        var (width, height) = WindowSize(document);
        int x0 = frame ? 1 : 0, y0 = frame ? 19 : 0;
        var bitmap = new MonoBitmap(width + (frame ? 3 : 0), height + (frame ? 21 : 0));
        var canvas = new Canvas(bitmap, x0, y0, document, width, height);
        foreach (var chart in document.Charts)
        {
            if (document.GetSeries(chart).Count > 0 || chart == document.MainChart)
            {
                new ChartPainter(document, chart, fonts, canvas).Paint();
            }
        }

        DrawGrowBox(canvas, width, height);
        if (frame)
        {
            ExcelRenderer.DrawWindowFrame(bitmap, width, height, title ?? "Chart", fonts.GetFont(0, 12)!);
        }

        return bitmap;
    }

    // The window's content size, from the chart file's window block (spec §16.1): its size words.
    private static (int Width, int Height) WindowSize(ChartDocument document)
    {
        var w = document.Window;
        int height = (short)((w[4] << 8) | w[5]), width = (short)((w[6] << 8) | w[7]);
        return (Math.Clamp(width, 64, 2048), Math.Clamp(height, 64, 2048));
    }

    // A chart window has no scroll bars, only the grow box in the corner where they would meet.
    private static void DrawGrowBox(Canvas canvas, int width, int height)
        => ExcelRenderer.DrawGrowBox(canvas.Bitmap, canvas.OriginX + width - 15, canvas.OriginY + height - 15);

    // The window's contents: the chart's drawing, the window less 16 pixels each way, mapped from twips to pixels
    // as 32.04 maps points (truncating; vertical coordinates run upwards) and 32.02 lengths.
    internal sealed class Canvas
    {
        private readonly MonoBitmap _bitmap;
        private readonly int _x, _y;
        private readonly int _pixelWidth, _pixelHeight, _twipWidth, _twipHeight;

        public Canvas(MonoBitmap bitmap, int x, int y, ChartDocument document, int width, int height)
        {
            _bitmap = bitmap;
            _x = x;
            _y = y;
            _pixelWidth = width - 16;
            _pixelHeight = height - 16;
            _twipWidth = Math.Max(1, (int)document.ChartWidth);
            _twipHeight = Math.Max(1, (int)document.ChartHeight);
        }

        public MonoBitmap Bitmap => _bitmap;

        public int OriginX => _x;

        public int OriginY => _y;

        /// <summary>A drawing point in twips (from the drawing's lower left) as a window pixel.</summary>
        public (int X, int Y) Point(int x, int y)
            => ((int)((long)x * _pixelWidth / _twipWidth), (int)((long)(_twipHeight - y) * _pixelHeight / _twipHeight));

        /// <summary>A horizontal length in pixels as twips (32.01).</summary>
        public int TwipsAcross(int pixels) => (int)((long)pixels * _twipWidth / _pixelWidth);

        /// <summary>A horizontal length in twips as pixels.</summary>
        public int Across(int twips) => (int)((long)twips * _pixelWidth / _twipWidth);

        /// <summary>A vertical length in twips as pixels.</summary>
        public int Down(int twips) => (int)((long)twips * _pixelHeight / _twipHeight);

        // The clip rectangle, in window pixels.
        private (int Left, int Top, int Right, int Bottom) _clip = NoClip;

        private static readonly (int, int, int, int) NoClip = (int.MinValue / 2, int.MinValue / 2, int.MaxValue / 2, int.MaxValue / 2);

        /// <summary>Clips what follows to a rectangle, or (with no rectangle) stops clipping.</summary>
        public void Clip((int Left, int Top, int Right, int Bottom)? clip) => _clip = clip ?? NoClip;

        private bool Clipped(ref int left, ref int top, ref int right, ref int bottom)
        {
            left = Math.Max(left, _clip.Left);
            top = Math.Max(top, _clip.Top);
            right = Math.Min(right, _clip.Right);
            bottom = Math.Min(bottom, _clip.Bottom);
            return right <= left || bottom <= top;
        }

        public void Fill(int left, int top, int right, int bottom, bool black)
        {
            if (!Clipped(ref left, ref top, ref right, ref bottom))
            {
                _bitmap.FillRect(_x + left, _y + top, _x + right, _y + bottom, black);
            }
        }

        /// <summary>QuickDraw's FrameRect with a one-pixel black pen: the rectangle's edge pixels inside it.</summary>
        public void Frame(int left, int top, int right, int bottom)
        {
            if (right <= left || bottom <= top)
            {
                return;
            }

            Fill(left, top, right, top + 1, true);
            Fill(left, bottom - 1, right, bottom, true);
            Fill(left, top, left + 1, bottom, true);
            Fill(right - 1, top, right, bottom, true);
        }

        public void Pattern(int left, int top, int right, int bottom, ulong pattern)
        {
            if (!Clipped(ref left, ref top, ref right, ref bottom))
            {
                _bitmap.FillPattern(_x + left, _y + top, _x + right, _y + bottom, pattern, -_x, -_y);
            }
        }

        /// <summary>A QuickDraw line with a one-pixel pen, both ends included.</summary>
        public const ulong Solid = 0xFFFFFFFFFFFFFFFF;

        // The pen's pattern in QuickDraw's patCopy mode, aligned to the window's contents.
        private void PenFill(int left, int top, int right, int bottom, ulong pen)
        {
            if (pen == Solid)
            {
                Fill(left, top, right, bottom, true);
            }
            else
            {
                Pattern(left, top, right, bottom, pen);
            }
        }

        public void Line(int x1, int y1, int x2, int y2, ulong pen = Solid)
        {
            if (x1 == x2)
            {
                PenFill(x1, Math.Min(y1, y2), x1 + 1, Math.Max(y1, y2) + 1, pen);
                return;
            }

            if (y1 == y2)
            {
                PenFill(Math.Min(x1, x2), y1, Math.Max(x1, x2) + 1, y1 + 1, pen);
                return;
            }

            // QuickDraw's DrawLine (DrawLine.a): from the top end down, one horizontal slab a scanline, its edges
            // in 16.16 fixed point starting half a pixel in and half a slope along, stepping by the slope.
            if (y1 > y2)
            {
                (x1, x2) = (x2, x1);
                (y1, y2) = (y2, y1);
            }

            int slope = FixRatio(x2 - x1, y2 - y1);
            long leftEdge = ((long)x1 << 16) + 0x8000, rightEdge = leftEdge + 0x10000;
            leftEdge += slope >> 1;
            rightEdge += slope >> 1;
            if (slope >= 0)
            {
                leftEdge -= slope;
                if (slope >= 0x10000)
                {
                    rightEdge -= 0x10000;
                }
                else
                {
                    leftEdge += slope;
                }
            }
            else
            {
                rightEdge -= slope;
                if (slope < -0x10000)
                {
                    leftEdge += 0x10000;
                }
                else
                {
                    rightEdge += slope;
                }
            }

            int minX = Math.Min(x1, x2), maxX = Math.Max(x1, x2) + 1;
            for (int y = y1; y <= y2; y++)
            {
                int left = Math.Max((int)(leftEdge >> 16), minX), right = Math.Min((int)(rightEdge >> 16), maxX);
                if (right > left)
                {
                    PenFill(left, y, right, y + 1, pen);
                }

                leftEdge += slope;
                rightEdge += slope;
            }
        }

        /// <summary>
        /// QuickDraw's FillPoly: the region the polygon's edges make. Each edge that is not upright adds PutLine's
        /// inversion points, one step a scanline where its left edge (as DrawLine works it out) moves to another
        /// pixel; a pixel is inside when an odd number of points lie at or above and left of it.
        /// </summary>
        public void FillPolygon(ReadOnlySpan<(int X, int Y)> points, ulong pattern)
        {
            if (points.Length < 3)
            {
                return;
            }

            int top = int.MaxValue, bottom = int.MinValue, left = int.MaxValue, right = int.MinValue;
            foreach (var (x, y) in points)
            {
                top = Math.Min(top, y);
                bottom = Math.Max(bottom, y);
                left = Math.Min(left, x);
                right = Math.Max(right, x);
            }

            int width = right - left + 1, rows = bottom - top + 1;
            var flips = new List<(int Y, int X)>();
            for (int i = 0; i < points.Length; i++)
            {
                var (x1, y1) = points[i];
                var (x2, y2) = points[(i + 1) % points.Length];
                PutLine(x1, y1, x2, y2, flips);
            }

            var parity = new bool[width + 1];
            for (int y = top; y <= bottom; y++)
            {
                foreach (var (fy, fx) in flips)
                {
                    if (fy == y)
                    {
                        for (int x = Math.Max(fx, left) - left; x <= width; x++)
                        {
                            parity[x] = !parity[x];
                        }
                    }
                }

                int run = -1;
                for (int x = 0; x <= width; x++)
                {
                    bool inside = x < width && parity[x];
                    if (inside && run < 0)
                    {
                        run = x;
                    }
                    else if (!inside && run >= 0)
                    {
                        Pattern(left + run, y, left + x, y + 1, pattern);
                        run = -1;
                    }
                }
            }
        }

        // PutLine: a line as inversion points, in pairs on the scanlines where it steps.
        private static void PutLine(int h1, int v1, int h2, int v2, List<(int Y, int X)> flips)
        {
            if (h1 == h2)
            {
                return;
            }

            if (v1 == v2)
            {
                flips.Add((v1, h1));
                flips.Add((v2, h2));
                return;
            }

            if (v2 < v1)
            {
                (h1, h2) = (h2, h1);
                (v1, v2) = (v2, v1);
            }

            int slope = FixRatio(h2 - h1, v2 - v1);
            long h = ((long)h1 << 16) + 0x8000 + (slope >> 1);
            if (slope >= 0 && slope < 0x10000)
            {
                h += slope;
            }
            else if (slope < -0x10000)
            {
                h += 0x10000;
            }

            int old = h1;
            for (int v = v1; v < v2; v++)
            {
                int now = (int)(h >> 16);
                if (now != old)
                {
                    flips.Add((v, old));
                    flips.Add((v, now));
                    old = now;
                }

                h += slope;
            }

            if (old != h2)
            {
                flips.Add((v2, old));
                flips.Add((v2, h2));
            }
        }

        /// <summary>
        /// QuickDraw's DrawArc (FrameArc, FillArc and the ovals) with a one-pixel pen. Each scanline's oval runs
        /// from the centre out by half pixels while (2x/w)² + (2y/h)² stays within 1, y measured to the row's middle
        /// (InitOval, BumpOval); a frame takes away the oval a pixel inside. An arc keeps what lies between the rays
        /// at its start and stop angles, clockwise from 12 o'clock, each ray's position on a row worked out in 16.16
        /// fixed point from its slope.
        /// </summary>
        public void Arc(int left, int top, int right, int bottom, int startAngle, int arcAngle, bool hollow, ulong pattern)
        {
            int width = right - left, height = bottom - top;
            if (width <= 0 || height <= 0 || arcAngle == 0)
            {
                return;
            }

            if (arcAngle < 0)
            {
                startAngle += arcAngle;
                arcAngle = -arcAngle;
            }

            bool arc = arcAngle < 360;
            startAngle %= 360;
            if (startAngle < 0)
            {
                startAngle += 360;
            }

            int stopAngle = startAngle + arcAngle;
            if (stopAngle >= 360)
            {
                stopAngle -= 360;
            }

            bool inner = hollow && right - left > 2 && bottom - top > 2;
            int midH = (left + right) >> 1, midV = (top + bottom) >> 1;
            int aspect = FixRatio(width, height);
            long slope1 = FixMul(SlopeFromAngle(startAngle), aspect), slope2 = FixMul(SlopeFromAngle(stopAngle), aspect);
            long line1 = ((long)midH << 16) - slope1 * (height >> 1), line2 = ((long)midH << 16) - slope2 * (height >> 1);
            int flag1 = startAngle < 180 ? startAngle - 90 : 270 - startAngle;
            int flag2 = stopAngle < 180 ? stopAngle - 90 : 270 - stopAngle;
            bool obtuse = arcAngle >= 180;
            for (int v = top; v < bottom; v++)
            {
                if (arc && v == midV)
                {
                    // Below the middle the rays' roles swap.
                    (flag1, flag2) = (-flag2, -flag1);
                    (line1, line2) = (line2, line1);
                    (slope1, slope2) = (slope2, slope1);
                }

                var (outerLeft, outerRight) = OvalRow(left, top, width, height, v);
                bool hasInner = inner && v >= top + 1 && v < bottom - 1;
                var (innerLeft, innerRight) = hasInner ? OvalRow(left + 1, top + 1, width - 2, height - 2, v) : (0, 0);
                if (!arc)
                {
                    Slabs(v, outerLeft, outerRight, hasInner, innerLeft, innerRight, pattern);
                }
                else
                {
                    int x1 = (int)(line1 >> 16), x2 = (int)(line2 >> 16);
                    bool active1 = flag1 < 0, active2 = flag2 < 0;
                    if (active1 && active2 && obtuse)
                    {
                        // Both rays in this half of an arc of 180° or more: everything but the wedge between them.
                        ArcSlabs(v, outerLeft, Math.Min(outerRight, x2), hasInner, innerLeft, Math.Min(innerRight, x2), pattern);
                        ArcSlabs(v, Math.Max(outerLeft, x1), outerRight, hasInner, Math.Max(innerLeft, x1), innerRight, pattern);
                    }
                    else if (active1 || active2 || obtuse)
                    {
                        int l = active1 ? Math.Max(outerLeft, x1) : outerLeft, r = active2 ? Math.Min(outerRight, x2) : outerRight;
                        int il = active2 ? Math.Min(innerLeft, x2) : innerLeft, ir = active1 ? Math.Max(innerRight, x1) : innerRight;
                        ArcSlabs(v, l, r, hasInner, il, ir, pattern);
                    }
                }

                line1 += slope1;
                line2 += slope2;
            }
        }

        private void ArcSlabs(int v, int left, int right, bool hasInner, int innerLeft, int innerRight, ulong pattern)
        {
            if (left >= right)
            {
                return;
            }

            if (hasInner)
            {
                PenFill(left, v, Math.Min(innerLeft, right), v + 1, pattern);
                PenFill(Math.Max(innerRight, left), v, right, v + 1, pattern);
            }
            else
            {
                PenFill(left, v, right, v + 1, pattern);
            }
        }

        private void Slabs(int v, int left, int right, bool hasInner, int innerLeft, int innerRight, ulong pattern)
        {
            if (hasInner)
            {
                PenFill(left, v, innerLeft, v + 1, pattern);
                PenFill(innerRight, v, right, v + 1, pattern);
            }
            else
            {
                PenFill(left, v, right, v + 1, pattern);
            }
        }

        // An oval's pixels on a scanline: half-pixel steps k out from the centre while k²(h/w)² ≤ h² − (2y)².
        private static (int Left, int Right) OvalRow(int left, int top, int width, int height, int v)
        {
            long y = 2 * (v - top) + 1 - height;
            double limit = (double)height * height - y * y;
            double ratio = (double)height / width;
            long k = (long)Math.Floor(Math.Sqrt(Math.Max(0, limit)) / ratio);
            while (k > 0 && k * k * ratio * ratio > limit)
            {
                k--;
            }

            while ((k + 1) * (k + 1) * ratio * ratio <= limit)
            {
                k++;
            }

            long centre = ((long)left << 16) + width * 32768L;
            return ((int)((centre - k * 32768) >> 16), (int)((centre + 32768 + k * 32768) >> 16));
        }

        // SlopeFromAngle: −tan(angle) in 16.16, the run across for each row down along a ray from the centre.
        private static int SlopeFromAngle(int angle)
        {
            angle %= 180;
            if (angle == 90)
            {
                return -0x7FFFFFFF;
            }

            return (int)Math.Round(-Math.Tan(angle * Math.PI / 180) * 65536);
        }

        private static long FixMul(long a, long b) => (a * b + 0x8000) >> 16;

        // FixRatio, truncating.
        private static int FixRatio(int numerator, int denominator) => (int)(((long)numerator << 16) / denominator);

        public int Text(MacBitmapFont font, int x, int baseline, ReadOnlySpan<char> text, bool bold)
            => font.DrawText(_bitmap, _x + x, _y + baseline, text, bold, false);
    }

    // Draws one chart (the main chart or the overlay) into the canvas.
    private sealed class ChartPainter(ChartDocument document, Chart chart, MacFontSet fonts, Canvas canvas)
    {
        // Chart coordinates are the chart area's, inside the drawing's margins.
        private readonly int _left = chart.DrawingMargins.Left;
        private readonly int _bottom = document.ChartHeight - chart.DrawingMargins.Top - chart.ChartArea.Box.Height;

        public void Paint()
        {
            if (chart.Type == ChartType.Pie)
            {
                _engine = new ChartLayoutEngine(document, chart, fonts);
                DrawPie();
                DrawTexts();
                return;
            }

            _engine = new ChartLayoutEngine(document, chart, fonts);
            DrawGridlines();
            DrawAxes();
            DrawValueLabels();
            DrawCategoryLabels();
            // Columns and areas are clipped to the plot area; a line chart's lines and markers are not.
            var plot = chart.PlotArea.Box;
            var (clipLeft, clipTop) = TextPoint(plot.Left, plot.Top + plot.Height);
            var (clipRight, clipBottom) = TextPoint(plot.Left + plot.Width, plot.Top);
            canvas.Clip((clipLeft, clipTop, clipRight, clipBottom));
            DrawColumns();
            DrawAreas();
            canvas.Clip(null);
            DrawSeries();
            DrawTexts();
        }

        // The titles and other text objects in their saved boxes: each line centred (or aligned) a pixel inside the
        // box, the first line's baseline a pixel and an ascent below its top. A chart without a title of its own but with
        // one named series has its name as the title, centred over the plot area at the top of the chart area.
        private void DrawTexts()
        {
            bool titled = false;
            foreach (var text in chart.Texts)
            {
                if (text.Kind is ChartTextKind.ChartTitle or ChartTextKind.ValueAxisTitle or ChartTextKind.CategoryAxisTitle
                    && !text.IsAutomaticText && text.Text is { Length: > 0 } content && !text.IsVertical)
                {
                    DrawTextBox(content, text.Area.Box, _engine.TextFont(text), text.HorizontalAlignment);
                }

                titled |= text.Kind == ChartTextKind.ChartTitle;
            }

            var series = document.GetSeries(chart);
            if (!titled && series.Count == 1 && series[0].Name.Length > 0)
            {
                var (font, bold) = _engine.TextFont(null);
                var area = chart.ChartArea.Box;
                var plot = chart.PlotArea.Box;
                Span<Range> lines = stackalloc Range[16];
                var (count, width) = ChartLayoutEngine.Wrap(series[0].Name, area.Width * 4 / 5 / TwipsPerPixel, font, bold, lines);
                int boxWidth = (width + 2) * TwipsPerPixel, boxHeight = (count * (font.Ascent + font.Descent + font.Leading) + 2) * TwipsPerPixel;
                var box = new ChartBox((short)(area.Height - boxHeight), (short)(plot.Left + (plot.Width - boxWidth) / 2), (short)boxHeight, (short)boxWidth);
                DrawTextBox(series[0].Name, box, (font, bold), ChartHorizontalAlignment.Center);
            }
        }

        private void DrawTextBox(string content, ChartBox box, (MacBitmapFont Font, bool Bold) textFont, ChartHorizontalAlignment alignment)
        {
            var (font, bold) = textFont;
            var (left, top) = TextPoint(box.Left, box.Top + box.Height);
            int inner = box.Width / TwipsPerPixel - 2;
            Span<Range> lines = stackalloc Range[64];
            var (count, _) = ChartLayoutEngine.Wrap(content, inner, font, bold, lines);
            int lineHeight = font.Ascent + font.Descent + font.Leading;
            for (int i = 0; i < Math.Min(count, lines.Length); i++)
            {
                var line = content.AsSpan()[lines[i]].TrimEnd(' ');
                int width = font.MeasureText(line, bold);
                int x = alignment switch
                {
                    ChartHorizontalAlignment.Left => left + 1,
                    ChartHorizontalAlignment.Right => left + 1 + inner - width,
                    _ => left + 1 + (inner - width) / 2,
                };

                canvas.Text(font, x, top + 1 + font.Ascent + i * lineHeight, line, bold);
            }
        }

        private ChartLayoutEngine _engine = null!;

        // A chart point (twips) as a pixel, as text is placed: no pen offset.
        private (int X, int Y) TextPoint(int x, int y) => canvas.Point(_left + x, _bottom + y);

        // The value at a fraction of the way up the value axis's scale.
        private double Fraction(double value)
        {
            var scale = chart.ValueScale;
            double range = scale.Maximum - scale.Minimum;
            // A logarithmic scale holds its ends as powers of ten.
            if (scale.IsLogarithmic)
            {
                value = value > 0 ? Math.Log10(value) : scale.Minimum;
            }

            double f = range == 0 ? 0 : (value - scale.Minimum) / range;
            return scale.ValuesInReverseOrder ? 1 - f : f;
        }

        // Value labels at each major unit. Up an upright axis each is right-aligned a pixel inside the widest label's
        // box (37.09), which ends two tick lengths left of the plot area, and the box (lines and a pixel's margin) is
        // centred on the tick; along an axis across, each label's box is centred under its tick, its top two tick
        // lengths below the axis.
        private void DrawValueLabels()
        {
            var scale = chart.ValueScale;
            if (chart.ValueAxis.TickLabelPosition == TickLabelPosition.None || chart.ValueAxis.Line.Pattern == ChartPattern.None
                || scale.MajorUnit <= 0 || scale.Maximum <= scale.Minimum)
            {
                return;
            }

            var plot = chart.PlotArea.Box;
            var (font, bold) = _engine.ValueFont;
            int room = _engine.ValueLabelRoom(chart.ChartArea.Box, plot.Width);
            Span<char> text = stackalloc char[96];
            int count = TickCount(value: true);
            for (int i = 0; i <= count; i++)
            {
                double value = scale.Minimum + i * scale.MajorUnit;
                int length = ChartLayoutEngine.FormatInRoom(scale.IsLogarithmic ? Math.Pow(10, value) : value, room, font, bold, text);
                if (chart.IsBar)
                {
                    DrawLabelBelow(text[..length], plot.Left + (int)((long)plot.Width * i / count), chart.ValueAxis, font, bold);
                }
                else
                {
                    DrawLabelBeside(text[..length], plot.Top + (int)((long)plot.Height * i / count), ChartAxisKind.Value, font, bold);
                }
            }
        }

        // Category labels, centred on their categories: under them in a column chart, beside them in a bar chart.
        private void DrawCategoryLabels()
        {
            if (chart.CategoryAxis.TickLabelPosition == TickLabelPosition.None || chart.CategoryAxis.Line.Pattern == ChartPattern.None)
            {
                return;
            }

            var series = document.GetSeries(chart);
            if (series.Count == 0)
            {
                return;
            }

            var (font, bold) = _engine.CategoryFont;
            if (chart.XScale is { } xScale)
            {
                // A scatter chart's X labels, at each major unit, as the value labels are.
                var plot = chart.PlotArea.Box;
                int ticks = TickCount(value: false);
                Span<char> number = stackalloc char[96];
                for (int i = 0; i <= ticks && ticks > 0; i++)
                {
                    int length = ChartLayoutEngine.FormatInRoom(xScale.Minimum + i * xScale.MajorUnit, plot.Width / Math.Max(1, ticks) / TwipsPerPixel, font, bold, number);
                    DrawLabelBelow(number[..length], plot.Left + (int)((long)plot.Width * i / ticks), chart.CategoryAxis, font, bold);
                }

                return;
            }

            int categories = Math.Max(1, (int)chart.CategoryCount);
            var first = series[0];
            for (int i = 0; i < categories; i++)
            {
                string label = first.SeriesType == 3 && i < first.CategoryTexts.Count ? first.CategoryTexts[i] : (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (chart.IsBar)
                {
                    DrawLabelBeside(label, CategoryAt(i, categories), ChartAxisKind.Category, font, bold);
                }
                else
                {
                    DrawLabelBelow(label, CategoryAt(i, categories), chart.CategoryAxis, font, bold);
                }
            }
        }

        // A label beside an upright axis: right-aligned a pixel inside the axis's widest label box, which ends two tick
        // lengths left of the plot area, the box a line and a pixel's margin high, centred on the label's place.
        private void DrawLabelBeside(ReadOnlySpan<char> label, int y, ChartAxisKind axis, MacBitmapFont font, bool bold)
        {
            var plot = chart.PlotArea.Box;
            var (_, boxWidth) = document.GetTickLabelSize(chart, fonts, axis);
            int tick = (axis == ChartAxisKind.Value ? chart.ValueAxis : chart.CategoryAxis).TickLength;
            var (boxLeft, _) = TextPoint(plot.Left - 2 * tick - boxWidth, 0);
            int right = boxLeft + boxWidth / TwipsPerPixel - 1;
            int boxHeight = (font.Ascent + font.Descent + font.Leading + 2) * TwipsPerPixel;
            var (_, top) = TextPoint(0, y + boxHeight / 2);
            canvas.Text(font, right - font.MeasureText(label, bold), top + 1 + font.Ascent, label, bold);
        }

        // A label along an axis across: its box (the text and a pixel's margin) centred on its place, the text a pixel
        // inside, the box's top two tick lengths below the axis.
        private void DrawLabelBelow(ReadOnlySpan<char> label, int x, ChartAxis axis, MacBitmapFont font, bool bold)
        {
            var (_, top) = TextPoint(0, axis.LineAt - 2 * axis.TickLength);
            int width = font.MeasureText(label, bold);
            var (boxLeft, _) = TextPoint(x - (width + 2) * TwipsPerPixel / 2, 0);
            canvas.Text(font, boxLeft + 1, top + 1 + font.Ascent, label, bold);
        }

        // The middle of a category along the category axis: centred in its share when the value axis crosses between
        // categories, otherwise at its share's edges.
        private int CategoryAt(int index, int categories)
        {
            var plot = chart.PlotArea.Box;
            int start = chart.IsBar ? plot.Top : plot.Left, length = chart.IsBar ? plot.Height : plot.Width;
            return chart.ValueAxisCrossesBetweenCategories
                ? start + (int)((long)length * (2 * index + 1) / (2 * categories))
                : start + (categories <= 1 ? 0 : (int)((long)length * index / (categories - 1)));
        }

        // The 15 patterns, as QuickDraw patterns (spec §16.4).
        private static ReadOnlySpan<ulong> Patterns =>
        [
            0xFFFFFFFFFFFFFFFF, 0xEEBBEEBBEEBBEEBB, 0x55AA55AA55AA55AA, 0x8822882288228822, 0x0000000000000000,
            0xFFFF0000FFFF0000, 0xCCCCCCCCCCCCCCCC, 0xCC663399CC663399, 0x993366CC993366CC, 0x808080FF80808080,
            0x8041221408142241, 0xFCF9F3E7CF9F3F7E, 0x03060C183060C081, 0x3F9FCFE7F3F9FC7E, 0xC06030180C060381,
        ];

        // A series' format (or a point's, in a chart with one series): its own point-format record, or the chart's
        // default series format.
        private ChartPointFormat SeriesFormat(int series)
        {
            for (int i = 0; i < chart.PointFormats.Count; i++)
            {
                var format = chart.GetPointFormat(i);
                if (format.SeriesOrPoint == series)
                {
                    return format;
                }
            }

            return chart.DefaultSeriesFormat;
        }

        // An automatic area takes the patterns in turn, series by series (3c.03).
        private ChartPattern AreaPattern(int series)
        {
            var format = SeriesFormat(series);
            return format.AutomaticArea ? (ChartPattern)(series % 15 + 1) : format.AreaPattern;
        }

        // Column charts (39.03, 39.04, 39.00): each category's share of the plot area holds its series' columns, a
        // column 100 units wide, the next series' a column and the overlap further on, and half the gap between
        // clusters either side. Each column runs from where the category axis crosses to its value; its corners are
        // mapped to pixels, a column no pixel wide made one wide, and it is filled with its pattern and framed.
        private void DrawColumns()
        {
            if (chart.Type != ChartType.ColumnOrBar || chart.IsStacked)
            {
                return;
            }

            // A bar chart's categories run up the plot area (37.24), its values across.
            bool bar = chart.IsBar;
            var plot = chart.PlotArea.Box;
            var all = document.GetSeries(chart);
            int categories = Math.Max(1, (int)chart.CategoryCount);
            bool between = chart.ValueAxisCrossesBetweenCategories;
            int spans = TickCount(value: false);
            int start = bar ? plot.Top : plot.Left, length = bar ? plot.Height : plot.Width;
            int category = length / spans;
            int units = chart.ClusterSpacing + (all.Count - 1) * (100 + chart.Overlap) + 100;
            int width = units == 0 ? category : (int)((long)category * 100 / units);
            int halfGap = (int)((long)width * chart.ClusterSpacing / 100) / 2;
            int step = (int)((long)width * chart.Overlap / 100) + width;
            var scale = chart.ValueScale;
            int baseAt = ValueAt(Math.Clamp(scale.CrossesAt, scale.Minimum, scale.Maximum));
            for (int s = 0; s < all.Count; s++)
            {
                var values = all[s].Values;
                var pattern = AreaPattern(s);
                var format = SeriesFormat(s);
                for (int i = 0; i < values.Count && i < categories; i++)
                {
                    int from = start + (int)((long)length * i / spans) + s * step + halfGap - (between ? 0 : category / 2);
                    int at = ValueAt(values[i].ToDouble());
                    int low = Math.Min(at, baseAt), high = Math.Max(at, baseAt);
                    var (x1, y1) = bar ? TextPoint(low, from + width) : TextPoint(from, high);
                    var (x2, y2) = bar ? TextPoint(high, from) : TextPoint(from + width, low);
                    // A column no pixel wide is drawn a pixel wide; one no pixel high is not drawn.
                    if (bar)
                    {
                        if (y2 == y1)
                        {
                            y2++;
                        }

                        if (x2 == x1)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        if (x2 == x1)
                        {
                            x2++;
                        }

                        if (y2 == y1)
                        {
                            continue;
                        }
                    }

                    if (pattern != ChartPattern.None)
                    {
                        canvas.Pattern(x1, y1, x2, y2, Patterns[(int)pattern - 1]);
                    }

                    if (format.AutomaticLine || format.Line.Pattern != ChartPattern.None)
                    {
                        canvas.Frame(x1, y1, x2, y2);
                    }
                }
            }
        }

        // Where a value falls along the value axis, in twips from the chart area's foot or left (37.23).
        private int ValueAt(double value)
        {
            var plot = chart.PlotArea.Box;
            return chart.IsBar ? plot.Left + (int)(plot.Width * Fraction(value)) : plot.Top + (int)(plot.Height * Fraction(value));
        }

        // Area charts (55.00, 55.01): each series a polygon along its tops and back along its base (the series
        // below's tops when stacked, which adds the values' sizes; otherwise where the category axis crosses),
        // filled with its pattern, then its border drawn through the polygon's points.
        private void DrawAreas()
        {
            if (chart.Type != ChartType.Area)
            {
                return;
            }

            var all = document.GetSeries(chart);
            int categories = Math.Max(1, (int)chart.CategoryCount);
            var scale = chart.ValueScale;
            var below = new double[categories];
            var points = new List<(int X, int Y)>();
            for (int s = 0; s < all.Count; s++)
            {
                var values = all[s].Values;
                int count = Math.Min(values.Count, categories);
                var tops = new double[categories];
                for (int i = 0; i < count; i++)
                {
                    double value = values[i].ToDouble();
                    tops[i] = chart.IsStacked ? below[i] + Math.Abs(value) : value;
                }

                points.Clear();
                double baseValue = Math.Clamp(scale.CrossesAt, scale.Minimum, scale.Maximum);
                for (int i = 0; i < count; i++)
                {
                    points.Add(TextPoint(CategoryAt(i, categories), ValueAt(tops[i])));
                }

                for (int i = count - 1; i >= 0; i--)
                {
                    points.Add(TextPoint(CategoryAt(i, categories), ValueAt(chart.IsStacked && s > 0 ? below[i] : baseValue)));
                }

                var pattern = AreaPattern(s);
                if (pattern != ChartPattern.None)
                {
                    canvas.FillPolygon(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(points), Patterns[(int)pattern - 1]);
                }

                var format = SeriesFormat(s);
                if (format.AutomaticLine || format.Line.Pattern != ChartPattern.None)
                {
                    // The border runs along the tops only.
                    for (int i = 0; i < count - 1; i++)
                    {
                        var (x1, y1) = points[i];
                        var (x2, y2) = points[(i + 1) % points.Count];
                        canvas.Line(x1, y1, x2, y2);
                    }
                }

                if (chart.IsStacked)
                {
                    Array.Copy(tops, below, count);
                }
            }
        }

        // Excel's sines of whole degrees from 0 to 90, scaled by 32767 (in its resources at 0x6548 of 1.00).
        private static ReadOnlySpan<short> Sines =>
        [
            0, 572, 1144, 1715, 2286, 2856, 3425, 3993, 4560, 5126, 5690, 6252, 6813,
            7371, 7927, 8481, 9032, 9580, 10126, 10668, 11207, 11743, 12275, 12803, 13328, 13848,
            14364, 14876, 15383, 15886, 16384, 16876, 17364, 17846, 18323, 18794, 19260, 19720, 20173,
            20621, 21062, 21497, 21925, 22347, 22762, 23170, 23571, 23964, 24351, 24730, 25101, 25465,
            25821, 26169, 26509, 26841, 27165, 27481, 27788, 28087, 28377, 28659, 28932, 29196, 29451,
            29697, 29934, 30162, 30381, 30591, 30791, 30982, 31163, 31335, 31498, 31650, 31794, 31927,
            32051, 32165, 32269, 32364, 32448, 32523, 32587, 32642, 32687, 32722, 32747, 32762, 32767,
        ];

        // 56.06: the sine of a whole number of degrees, from the table.
        private static int Sine(int angle)
        {
            angle %= 360;
            if (angle < 0)
            {
                angle += 360;
            }

            bool positive = angle <= 180;
            if (!positive)
            {
                angle -= 180;
            }

            if (angle > 90)
            {
                angle = 180 - angle;
            }

            return positive ? Sines[angle] : -Sines[angle];
        }

        // Pie charts (56.00–56.05): the pie a circle in the middle of the plot area, as wide as its smaller side
        // allows; each slice from the running total of the values' sizes so far, as a share of 360°, turned by the
        // first slice's angle. A slice is filled (FillArc), then its edge drawn (FrameArc) with its two radii.
        private void DrawPie()
        {
            var all = document.GetSeries(chart);
            if (all.Count == 0)
            {
                return;
            }

            var values = all[0].Values;
            var plot = chart.PlotArea.Box;
            int diameter = Math.Min(plot.Width, plot.Height);
            int left = plot.Left + (plot.Width - diameter) / 2, bottom = plot.Top + (plot.Height - diameter) / 2;
            var (x1, y1) = TextPoint(left, bottom + diameter);
            var (x2, y2) = TextPoint(left + diameter, bottom);
            double total = 0;
            foreach (var value in values)
            {
                total += Math.Abs(value.ToDouble());
            }

            if (total == 0)
            {
                total = 1;
            }

            double sum = 0;
            int previous = 0;
            for (int i = 0; i < values.Count; i++)
            {
                sum += Math.Abs(values[i].ToDouble());
                double exact = sum * (360 / total);
                // SANE's conversion to an integer rounds to the nearest, even on a tie (inferred: no capture has a tie).
                int angle = i == values.Count - 1 ? 360 : (int)Math.Round(exact);
                int start = previous, arc = angle - previous;
                previous = angle;
                if (arc < 0)
                {
                    arc += 360;
                }

                if (arc <= 0)
                {
                    continue;
                }

                start += chart.FirstSliceAngle;
                var format = SeriesFormat(i);
                // An automatic slice is filled only when the chart varies by category (3c.03).
                var pattern = !format.AutomaticArea ? format.AreaPattern : chart.VaryByCategory ? (ChartPattern)(i % 15 + 1) : ChartPattern.None;
                if (pattern != ChartPattern.None)
                {
                    canvas.Arc(x1, y1, x2, y2, start, arc, hollow: false, Patterns[(int)pattern - 1]);
                }

                if (format.AutomaticLine || format.Line.Pattern != ChartPattern.None)
                {
                    DrawSlice(x1, y1, x2, y2, start, arc);
                }
            }
        }

        // 56.04: a slice's radii and edge. Each radius runs from the centre to the edge, both nudged by the pen
        // according to the radius's quadrant (56.05).
        private void DrawSlice(int left, int top, int right, int bottom, int start, int arc)
        {
            int width = right - left, height = bottom - top;
            int sumH = left + right, sumV = top + bottom;
            Radius(start);
            canvas.Arc(left, top, right, bottom, start, arc, hollow: true, Canvas.Solid);
            Radius(start + arc);

            void Radius(int angle)
            {
                int ch = sumH / 2, cv = sumV / 2;
                int eh = (int)((sumH + (long)width * Sine(angle) / 0x7FFF) / 2);
                int ev = (int)((sumV - (long)height * Sine(angle + 90) / 0x7FFF) / 2);
                switch (angle)
                {
                    case 0: cv--; break;
                    case 90: eh--; break;
                    case 180: ev--; break;
                    case 270: ch--; break;
                    case > 0 and < 90: cv--; eh--; break;
                    case > 90 and < 180: eh--; break;
                    case > 180 and < 270: ch--; ev--; break;
                    case > 270 and < 360: cv--; ch--; ev++; break;
                }

                canvas.Line(ch, cv, eh, ev);
            }
        }

        // Line charts: each series' points joined by lines, a marker on each.
        private void DrawSeries()
        {
            bool scatter = chart.Type == ChartType.Scatter;
            if (chart.Type != ChartType.Line && !scatter)
            {
                return;
            }

            int categories = Math.Max(1, (int)chart.CategoryCount);
            var all = document.GetSeries(chart);
            var points = new List<(int X, int Y)>();
            for (int s = 0; s < all.Count; s++)
            {
                var series = all[s];
                var values = series.Values;
                var format = SeriesFormat(s);
                (int X, int Y)? previous = null;
                points.Clear();
                int count = scatter ? values.Count : Math.Min(values.Count, categories);
                for (int i = 0; i < count; i++)
                {
                    // A scatter chart's points stand at their X values; a line chart's at their categories.
                    int x = scatter ? XAt(i < series.CategoryValues.Count ? series.CategoryValues[i].ToDouble() : i + 1) : CategoryAt(i, categories);
                    var twips = (X: x, Y: ValueAt(values[i].ToDouble()));
                    if (!scatter && previous is { } p)
                    {
                        // Each segment starts half a pen from its point and runs for its length in pixels (32.07).
                        var start = LinePoint(p.X, p.Y);
                        int dx = canvas.Across(twips.X - p.X), dy = -canvas.Down(twips.Y - p.Y);
                        canvas.Line(start.X, start.Y, start.X + dx, start.Y + dy);
                    }

                    previous = twips;
                    points.Add(TextPoint(twips.X, twips.Y));
                }

                var marker = format.AutomaticMarker ? (ChartMarker)(s % 8 + 1) : format.Marker;
                foreach (var (x, y) in points)
                {
                    DrawMarker(x, y, marker);
                }
            }
        }

        // Where an X value falls across a scatter chart's plot area.
        private int XAt(double value)
        {
            var plot = chart.PlotArea.Box;
            var scale = chart.XScale!;
            double range = scale.Maximum - scale.Minimum;
            double f = range == 0 ? 0 : (value - scale.Minimum) / range;
            return plot.Left + (int)(plot.Width * (scale.ValuesInReverseOrder ? 1 - f : f));
        }

        // Excel's marker bitmaps, eight rows of eight pixels each for the markers 1 to 8 (in its resources at 0x6428
        // and 0x64b8 of 1.00): the mask is cleared, then the image drawn (53.08's CopyBits, srcBic then srcOr). The
        // masks reach a pixel beyond the images, so lines stop short of the marker. The star has no mask.
        private static ReadOnlySpan<byte> MarkerMasks =>
        [
            0x18, 0x3C, 0x7E, 0xFF, 0xFF, 0x7E, 0x3C, 0x18,
            0x18, 0x3C, 0x7E, 0xFF, 0xFF, 0x7E, 0x3C, 0x18,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0x18, 0x3C, 0x3C, 0x7E, 0x7E, 0xFF, 0xFF, 0x7E,
            0x18, 0x3C, 0x3C, 0x7E, 0x7E, 0xFF, 0xFF, 0x7E,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];

        private static ReadOnlySpan<byte> MarkerImages =>
        [
            0x00, 0x18, 0x3C, 0x7E, 0x7E, 0x3C, 0x18, 0x00,
            0x00, 0x18, 0x24, 0x42, 0x42, 0x24, 0x18, 0x00,
            0x00, 0x7E, 0x7E, 0x7E, 0x7E, 0x7E, 0x7E, 0x00,
            0x00, 0x7E, 0x42, 0x42, 0x42, 0x42, 0x7E, 0x00,
            0x00, 0x18, 0x18, 0x3C, 0x3C, 0x7E, 0x7E, 0x00,
            0x00, 0x18, 0x18, 0x24, 0x24, 0x42, 0x7E, 0x00,
            0x00, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x00,
            0x00, 0x54, 0x38, 0x10, 0x38, 0x54, 0x00, 0x00,
        ];

        // A marker in the 8-pixel square from 4 left of and above the pen to 3 right of and below it.
        private void DrawMarker(int x, int y, ChartMarker marker)
        {
            int index = (int)marker - 1;
            if ((uint)index >= 8)
            {
                return;
            }

            for (int row = 0; row < 8; row++)
            {
                int mask = MarkerMasks[index * 8 + row], image = MarkerImages[index * 8 + row];
                for (int column = 0; column < 8; column++)
                {
                    int bit = 0x80 >> column;
                    if ((mask & bit) != 0)
                    {
                        canvas.Fill(x - 4 + column, y - 4 + row, x - 3 + column, y - 3 + row, false);
                    }

                    if ((image & bit) != 0)
                    {
                        canvas.Fill(x - 4 + column, y - 4 + row, x - 3 + column, y - 3 + row, true);
                    }
                }
            }
        }

        // A chart point (twips) as a pixel, for a line of a one-pixel pen centred on it: the pen hangs half a pixel
        // left of and above the point.
        private (int X, int Y) LinePoint(int x, int y) => canvas.Point(_left + x - TwipsPerPixel / 2, _bottom + y + TwipsPerPixel / 2);

        // Gridlines, before the axes (3a.00): the value axis's minor then major, then the category axis's, each across
        // the plot area from its start, half a pen off, in the gridline's pattern. Minor gridlines on a logarithmic
        // scale fall at 2 to 9 times each power of ten.
        private void DrawGridlines()
        {
            var valueAxis = chart.ValueAxis;
            var scale = chart.ValueScale;
            int count = TickCount(value: true);
            if (count > 0)
            {
                if (valueAxis.MinorGridlines.Pattern != ChartPattern.None)
                {
                    var pen = Patterns[(int)valueAxis.MinorGridlines.Pattern - 1];
                    double range = scale.Maximum - scale.Minimum;
                    if (scale.IsLogarithmic)
                    {
                        for (double decade = Math.Floor(scale.Minimum); decade < scale.Maximum; decade++)
                        {
                            for (int k = 2; k <= 9; k++)
                            {
                                double f = (decade + Math.Log10(k) - scale.Minimum) / range;
                                if (f > 0 && f < 1)
                                {
                                    ValueGridline(scale.ValuesInReverseOrder ? 1 - f : f, pen);
                                }
                            }
                        }
                    }
                    else if (scale.MinorUnit > 0)
                    {
                        int minors = (int)Math.Round(range / scale.MinorUnit);
                        for (int i = 0; i <= minors && minors <= 1000; i++)
                        {
                            ValueGridline((double)i / minors, pen);
                        }
                    }
                }

                if (valueAxis.MajorGridlines.Pattern != ChartPattern.None)
                {
                    var pen = Patterns[(int)valueAxis.MajorGridlines.Pattern - 1];
                    for (int i = 0; i <= count; i++)
                    {
                        ValueGridline((double)i / count, pen);
                    }
                }
            }

            var categoryAxis = chart.CategoryAxis;
            int spans = TickCount(value: false);
            if (categoryAxis.MinorGridlines.Pattern != ChartPattern.None)
            {
                var pen = Patterns[(int)categoryAxis.MinorGridlines.Pattern - 1];
                for (int i = 0; i <= 2 * spans; i++)
                {
                    if (i % 2 != 0 || categoryAxis.MajorGridlines.Pattern == ChartPattern.None)
                    {
                        CategoryGridline((double)i / (2 * spans), pen);
                    }
                }
            }

            if (categoryAxis.MajorGridlines.Pattern != ChartPattern.None)
            {
                var pen = Patterns[(int)categoryAxis.MajorGridlines.Pattern - 1];
                for (int i = 0; i <= spans; i++)
                {
                    CategoryGridline((double)i / spans, pen);
                }
            }
        }

        // A gridline at a fraction of the way along the value axis.
        private void ValueGridline(double fraction, ulong pen)
        {
            var plot = chart.PlotArea.Box;
            if (chart.IsBar)
            {
                var (px, py) = LinePoint(plot.Left + (int)(plot.Width * fraction), plot.Top);
                canvas.Line(px, py, px, py - canvas.Down(plot.Height), pen);
            }
            else
            {
                var (px, py) = LinePoint(plot.Left, plot.Top + (int)(plot.Height * fraction));
                canvas.Line(px, py, px + canvas.Across(plot.Width), py, pen);
            }
        }

        // A gridline at a fraction of the way along the category axis.
        private void CategoryGridline(double fraction, ulong pen)
        {
            var plot = chart.PlotArea.Box;
            if (chart.IsBar)
            {
                var (px, py) = LinePoint(plot.Left, plot.Top + (int)(plot.Height * fraction));
                canvas.Line(px, py, px + canvas.Across(plot.Width), py, pen);
            }
            else
            {
                var (px, py) = LinePoint(plot.Left + (int)(plot.Width * fraction), plot.Top);
                canvas.Line(px, py, px, py - canvas.Down(plot.Height), pen);
            }
        }

        private void DrawAxes()
        {
            var plot = chart.PlotArea.Box;
            int left = plot.Left, right = plot.Left + plot.Width, bottom = plot.Top, top = plot.Top + plot.Height;
            // The value axis is upright in a column chart and across in a bar chart; the category axis the other way.
            var upright = chart.IsBar ? chart.CategoryAxis : chart.ValueAxis;
            var across = chart.IsBar ? chart.ValueAxis : chart.CategoryAxis;

            // Each line runs from its start, half a pen off, for its length in pixels (3a.03, 32.07).
            if (upright.Line.Pattern != ChartPattern.None)
            {
                int x = upright.LineAt;
                var (px, py) = LinePoint(x, bottom);
                canvas.Line(px, py, px, py - canvas.Down(top - bottom));
                int count = TickCount(!chart.IsBar);
                for (int i = 0; i <= count && count > 0; i++)
                {
                    TickAcross(x, bottom + (int)((long)(top - bottom) * i / count), upright.MajorTickMarks);
                }
            }

            if (across.Line.Pattern != ChartPattern.None)
            {
                int y = across.LineAt;
                var (px, py) = LinePoint(left, y);
                canvas.Line(px, py, px + canvas.Across(right - left), py);
                int count = TickCount(chart.IsBar);
                for (int i = 0; i <= count && count > 0; i++)
                {
                    TickUp(left + (int)((long)(right - left) * i / count), y, across.MajorTickMarks);
                }
            }
        }

        // The number of spaces between major ticks: the value axis's major units, or the category axis's categories
        // (or the gaps between them when the value axis crosses at a category).
        private int TickCount(bool value)
        {
            if (value)
            {
                var scale = chart.ValueScale;
                return scale.MajorUnit <= 0 || scale.Maximum <= scale.Minimum ? 0 : (int)Math.Round((scale.Maximum - scale.Minimum) / scale.MajorUnit);
            }

            if (chart.XScale is { } x)
            {
                return x.MajorUnit <= 0 || x.Maximum <= x.Minimum ? 0 : (int)Math.Round((x.Maximum - x.Minimum) / x.MajorUnit);
            }

            int categories = Math.Max(1, (int)chart.CategoryCount);
            return chart.ValueAxisCrossesBetweenCategories ? categories : Math.Max(1, categories - 1);
        }

        // The length ticks are drawn (3a.00): a 70th of the plot area's smaller side, and at least a pixel and a half.
        // The saved tick length, a 70th of the chart area's, only lays the chart out.
        private int MajorTickLength
        {
            get
            {
                var plot = chart.PlotArea.Box;
                return Math.Max(Math.Min(plot.Width, plot.Height) / 70, canvas.TwipsAcross(1) * 3 / 2);
            }
        }

        // A tick on an upright axis (3a.04, 3a.05): from its start, half a pen off, for its length in pixels; a tick
        // outside or across starts its length left of the axis.
        private void TickAcross(int x, int y, TickMarkType marks)
        {
            if (marks == TickMarkType.Invisible)
            {
                return;
            }

            int length = MajorTickLength;
            int start = marks is TickMarkType.Outside or TickMarkType.Cross ? x - length : x;
            var (px, py) = LinePoint(start, y);
            canvas.Line(px, py, px + canvas.Across(marks == TickMarkType.Cross ? 2 * length : length), py);
        }

        // A tick on an axis across: one outside or across starts its length below the axis and runs up.
        private void TickUp(int x, int y, TickMarkType marks)
        {
            if (marks == TickMarkType.Invisible)
            {
                return;
            }

            int length = MajorTickLength;
            int start = marks is TickMarkType.Outside or TickMarkType.Cross ? y - length : y;
            var (px, py) = LinePoint(x, start);
            canvas.Line(px, py, px, py - canvas.Down(marks == TickMarkType.Cross ? 2 * length : length));
        }
    }
}
