namespace MacExcelReader.Charts;

/// <summary>
/// How Excel 1.00 works out a value axis's automatic minimum, maximum, units and crossing point when it draws a
/// chart (<c>37.0a</c>, docs/Display.md "Automatic axis scales"), step for step in the same floating-point
/// operations, so the results match Excel's to the bit.
/// </summary>
internal static class ChartAutomaticScale
{
    // The major unit's choices, as percentages of the power of ten below the range (global g24fa in 1.00).
    private static readonly int[] s_majorUnits = [10, 20, 50, 100, 200, 500, 1000];

    // The powers of ten 4a.01 reads from its table (global g11fa) instead of multiplying.
    private static readonly double[] s_smallPowers = [1, 10, 100, 1000, 10000];

    /// <summary>Sets the automatic fields of <paramref name="scale"/> as <c>37.0a</c> does.</summary>
    public static void Apply(ChartScale scale, double dataMinimum, double dataMaximum, int labelSpaces, bool scatter)
    {
        double min = dataMinimum, max = dataMaximum;

        // A single value: the minimum moves a tenth of its size away from zero (and for a scatter chart the
        // maximum the other way); all zeros give 0 to 1.
        if (max - min == 0)
        {
            if (min == 0)
            {
                max = 1;
            }
            else
            {
                bool negative = min < 0;
                double tenth = min / 10;
                min = (negative ? tenth : -tenth) + min;
                if (scatter)
                {
                    tenth = max / 10;
                    max = (negative ? -tenth : tenth) + max;
                }
            }
        }

        bool log = scale.IsLogarithmic;
        if (!scale.AutomaticMaximum)
        {
            max = scale.Maximum;
            if (log)
            {
                max = Pow10(Floor(CeilingLog10(max)));
            }
        }

        if (!scale.AutomaticMinimum)
        {
            min = scale.Minimum;
            if (log)
            {
                // A fixed minimum on a logarithmic axis is held as its exponent.
                min = Pow10(Floor(min));
            }
        }

        double range = max - min;
        if (log)
        {
            ApplyLogarithmic(scale, min, max, labelSpaces + 1);
        }
        else
        {
            // Start from zero when the data does not straddle it and its range is at least 8% of the value
            // furthest from zero. Scatter charts never do.
            if (scale.AutomaticMinimum && !scatter)
            {
                bool maxNotNegative = max >= 0;
                if (min >= 0 || !maxNotNegative)
                {
                    double share = Math.Abs(100.0 * range / (maxNotNegative ? max : min)) - 10 + 2;
                    if (share >= 0)
                    {
                        if (maxNotNegative)
                        {
                            min = 0;
                        }
                        else
                        {
                            max = 0;
                        }

                        range = max - min;
                    }
                }
            }

            if (scale.AutomaticMajorUnit)
            {
                Span<char> digits = stackalloc char[15];
                ExcelValueFormatter.Num2Str(Math.Abs(range), digits, out int exponent);
                double power = Pow10(exponent - 1);
                int intervals = Math.Clamp(labelSpaces + 1, 2, 11) - 1;
                double percent = range / intervals / power * 100;
                int needed = (int)percent;
                if (Math.Truncate(percent) - percent != 0)
                {
                    needed++;
                }

                int choice = 0;
                while (choice < s_majorUnits.Length - 1 && needed > s_majorUnits[choice])
                {
                    choice++;
                }

                scale.MajorUnit = s_majorUnits[choice] * power / 100;
            }

            if (scale.AutomaticMinorUnit)
            {
                scale.MinorUnit = scale.MajorUnit / 5;
            }

            if (scale.AutomaticMaximum)
            {
                scale.Maximum = max;
            }

            if (scale.AutomaticMinimum)
            {
                scale.Minimum = min;
            }

            scale.MajorUnit = LimitUnit(scale, scale.MajorUnit);
            scale.MinorUnit = LimitUnit(scale, scale.MinorUnit);
            if (scale.AutomaticMaximum)
            {
                scale.Maximum = RoundToUnit(max, scale.MajorUnit, up: true);
            }

            if (scale.AutomaticMinimum)
            {
                scale.Minimum = RoundToUnit(min, scale.MajorUnit, up: false);
            }
        }

        // The other axis crosses at zero, or at the end of the scale nearest to it.
        if (scale.AutomaticCrossesAt)
        {
            double crosses = 0;
            if (scale.Maximum < 0)
            {
                crosses = scale.Maximum;
            }

            if (scale.Minimum > 0)
            {
                crosses = scale.Minimum;
            }

            scale.CrossesAt = crosses;
        }

        if (scale.AutomaticMinimum && scale.Minimum - scale.Maximum >= 0)
        {
            scale.Minimum = scale.Maximum - scale.MajorUnit;
        }

        if (scale.AutomaticMaximum && scale.Maximum - scale.Minimum <= 0)
        {
            scale.Maximum = scale.Minimum + scale.MajorUnit;
        }
    }

    // 37.2d: a logarithmic axis holds powers of ten as their exponents, and its units are whole exponents.
    private static void ApplyLogarithmic(ChartScale scale, double min, double max, int labels)
    {
        double maxExponent = CeilingLog10(max);
        if (scale.AutomaticMaximum)
        {
            scale.Maximum = maxExponent;
        }

        if (scale.AutomaticMinorUnit)
        {
            scale.MinorUnit = 1;
        }

        double minExponent = (int)(100 + Math.Log10(Math.Abs(min))) - 100;
        if (scale.AutomaticMinimum)
        {
            scale.Minimum = minExponent;
        }

        double range = maxExponent - minExponent;
        if (scale.AutomaticMajorUnit)
        {
            scale.MajorUnit = (Floor(range) - 1) / Math.Max(1, labels - 1) + 1;
        }

        if (scale.AutomaticMaximum)
        {
            scale.Maximum = RoundToUnit(range, scale.MajorUnit, up: true) + minExponent;
        }
    }

    // 37.0c: a unit too small to change the ends becomes the whole range, and no scale has more than 500 units.
    private static double LimitUnit(ChartScale scale, double unit)
    {
        double max = scale.Maximum, min = scale.Minimum;
        if ((max + unit) - max == 0 || (min + unit) - min == 0)
        {
            return max - min;
        }

        double range = max - min;
        return range / unit - 500 > 0 ? range / 500 : unit;
    }

    // 37.0d: a value moved out to a whole number of units, away from zero in the direction asked for.
    private static double RoundToUnit(double value, double unit, bool up)
    {
        bool notNegative = value >= 0;
        double rounded = Math.Truncate(value / unit) * unit;
        if (rounded - value == 0)
        {
            return rounded;
        }

        if (notNegative)
        {
            return up ? rounded + unit : rounded;
        }

        return up ? rounded : rounded - unit;
    }

    // 37.2e: the exponent of the power of ten at or above the value, allowing 1E-10 for rounding in the logarithm.
    private static double CeilingLog10(double value)
    {
        double log = Math.Log10(value);
        int below = Floor(log);
        return 1e-10 - Math.Abs(below - log) >= 0 ? below : below + 1;
    }

    // 13.0f: rounds down through 32768.
    private static int Floor(double value) => value + 32768.0 < 0 ? -32768 : (int)(value + 32768.0) - 32768;

    // 4a.01: ten to a whole power, by squaring.
    internal static double Pow10(int exponent)
    {
        if (exponent < 0)
        {
            return exponent == int.MinValue ? 0 : 1 / Pow10(-exponent);
        }

        if (exponent <= 4)
        {
            return s_smallPowers[exponent];
        }

        double result = 1, power = 10;
        while (true)
        {
            if ((exponent & 1) != 0)
            {
                result = power * result;
            }

            exponent >>= 1;
            if (exponent == 0)
            {
                return result;
            }

            power *= power;
        }
    }
}
