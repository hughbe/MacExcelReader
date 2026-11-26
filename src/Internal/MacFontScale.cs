namespace MacExcelReader.Internal;

/// <summary>
/// How the Mac Plus ROM scales a font size the System has no bitmap for (spec Appendix D). <c>FMSwapFont</c> picks
/// another size of the family and returns the scale as numer ÷ denom, with denom 256 and numer the size ratio as a
/// <c>Fixed</c> shifted right 8 bits: Geneva 11 drawn from Geneva 12 is 234/256, not 11/12. QuickDraw measures and draws
/// with that ratio.
/// </summary>
internal static class MacFontScale
{
    /// <summary>The denominator <c>FMSwapFont</c> returns for a scaled size.</summary>
    public const int Denom = 256;

    /// <summary>
    /// Gets the numerator <c>FMSwapFont</c> returns for <paramref name="size"/> drawn from a bitmap of
    /// <paramref name="bitmapSize"/>: <c>FixRatio(size, bitmapSize)</c> shifted right 8 bits.
    /// </summary>
    public static int Numer(int size, int bitmapSize) => (int)(FixRatio(size, bitmapSize) >> 8);

    /// <summary>
    /// The Toolbox's <c>FixRatio</c> for positive arguments: numer ÷ denom as a 16.16 fixed-point number, rounded down.
    /// </summary>
    public static long FixRatio(int numer, int denom) => ((long)numer << 16) / denom;

    /// <summary>
    /// Scales a width in the bitmap's pixels as <c>TextWidth</c> and <c>CharWidth</c> do: times numer ÷ denom, rounded
    /// down.
    /// </summary>
    public static int Width(long unscaled, int numer) => (int)(unscaled * numer / Denom);

    /// <summary>
    /// Scales a font metric as <c>GetFontInfo</c> does, and a distance as <c>MapRect</c> does: times numer ÷ denom,
    /// rounded to the nearest pixel (halves away from zero).
    /// </summary>
    public static int Round(long value, int numer) => value < 0
        ? -(int)((-value * numer + Denom / 2) / Denom)
        : (int)((value * numer + Denom / 2) / Denom);
}
