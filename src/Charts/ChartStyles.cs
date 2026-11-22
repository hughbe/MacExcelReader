namespace MacExcelReader.Charts;

/// <summary>
/// A fill or line pattern: an index into Excel's 15 QuickDraw patterns, in the order the Patterns dialogs show
/// them (spec §16.4). 0 means invisible.
/// </summary>
public enum ChartPattern : byte
{
    /// <summary>Invisible: the line, border or area is not drawn.</summary>
    None = 0,

    /// <summary>Solid black.</summary>
    Solid = 1,

    /// <summary>75% grey.</summary>
    DarkGray = 2,

    /// <summary>50% grey.</summary>
    Gray = 3,

    /// <summary>25% grey.</summary>
    LightGray = 4,

    /// <summary>Solid white.</summary>
    White = 5,

    /// <summary>Horizontal stripes.</summary>
    HorizontalStripes = 6,

    /// <summary>Vertical stripes.</summary>
    VerticalStripes = 7,

    /// <summary>Diagonal stripes running down to the right (\).</summary>
    DiagonalDown = 8,

    /// <summary>Diagonal stripes running up to the right (/).</summary>
    DiagonalUp = 9,

    /// <summary>A square grid.</summary>
    Grid = 10,

    /// <summary>A diagonal lattice.</summary>
    DiagonalGrid = 11,

    /// <summary>Mostly black, with thin white lines running up to the right.</summary>
    ThickDiagonalUp = 12,

    /// <summary>Thin black lines running up to the right.</summary>
    ThinDiagonalUp = 13,

    /// <summary>Mostly black, with thin white lines running down to the right.</summary>
    ThickDiagonalDown = 14,

    /// <summary>Thin black lines running down to the right.</summary>
    ThinDiagonalDown = 15,
}

/// <summary>
/// A data-point marker, in the order of the Marker Pattern dialog (spec §16.8).
/// </summary>
public enum ChartMarker : byte
{
    /// <summary>No marker.</summary>
    None = 0,

    /// <summary>A filled diamond.</summary>
    FilledDiamond = 1,

    /// <summary>A hollow diamond.</summary>
    Diamond = 2,

    /// <summary>A filled square.</summary>
    FilledSquare = 3,

    /// <summary>A hollow square.</summary>
    Square = 4,

    /// <summary>A filled triangle.</summary>
    FilledTriangle = 5,

    /// <summary>A hollow triangle.</summary>
    Triangle = 6,

    /// <summary>An X.</summary>
    X = 7,

    /// <summary>An asterisk.</summary>
    Asterisk = 8,

    /// <summary>A short horizontal dash.</summary>
    Dash = 9,
}

/// <summary>The weight of a line, border or arrow shaft (spec §16.4).</summary>
public enum ChartLineWeight
{
    /// <summary>Hairline.</summary>
    Hairline = 0,

    /// <summary>Medium.</summary>
    Medium = 1,

    /// <summary>Thick.</summary>
    Thick = 2,
}

/// <summary>The width of an arrowhead (spec §16.7).</summary>
public enum ArrowHeadWidth
{
    /// <summary>Narrow.</summary>
    Narrow = 0,

    /// <summary>Medium.</summary>
    Medium = 1,

    /// <summary>Wide.</summary>
    Wide = 2,
}

/// <summary>The length of an arrowhead: bits 9–8 of the arrow's head word, 0 meaning no head (spec §16.7).</summary>
public enum ArrowHeadLength
{
    /// <summary>No head.</summary>
    None = 0,

    /// <summary>Short.</summary>
    Short = 1,

    /// <summary>Medium.</summary>
    Medium = 2,

    /// <summary>Long.</summary>
    Long = 3,
}

/// <summary>The style of an arrowhead, as the Patterns dialog offers it (spec §16.7).</summary>
public enum ArrowHeadStyle
{
    /// <summary>No head: the head length is 0.</summary>
    None = 0,

    /// <summary>A filled head: bit 7 clear.</summary>
    Filled = 1,

    /// <summary>An open head: bit 7 set.</summary>
    Open = 2,
}

/// <summary>The horizontal alignment of a chart text object (spec §16.7).</summary>
public enum ChartHorizontalAlignment
{
    /// <summary>Centred.</summary>
    Center = 0,

    /// <summary>Left.</summary>
    Left = 1,

    /// <summary>Right.</summary>
    Right = 2,
}

/// <summary>The vertical alignment of a chart text object (spec §16.7).</summary>
public enum ChartVerticalAlignment
{
    /// <summary>Bottom.</summary>
    Bottom = 0,

    /// <summary>Centred.</summary>
    Center = 1,

    /// <summary>Top.</summary>
    Top = 2,
}
