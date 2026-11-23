namespace MacExcelReader;

/// <summary>
/// How Excel 1.00 draws a cell's value in its row (spec Appendix D), from
/// <see cref="ExcelValueFormatter.LayoutCell(ExcelDocument, int, int)"/>. Positions are in pixels from the left edge
/// of column A, with the sheet's columns side by side and hidden ones taking no room.
/// </summary>
/// <param name="Text">The text drawn: the formatted value, a row of <c>#</c> for one too wide for its column, or a
/// fill's text repeated.</param>
/// <param name="Alignment">The alignment it is drawn with: never General, and Center for a row of <c>#</c>.</param>
/// <param name="X">Where the text starts. It can be left of <paramref name="ClipLeft"/>.</param>
/// <param name="Width">The text's width, with the room Excel allows for italic.</param>
/// <param name="ClipLeft">The first pixel the text may cover: the left edge of the cell, or of the leftmost blank
/// cell it runs into.</param>
/// <param name="ClipRight">The pixel the text stops before: the right gridline of the cell, or of the rightmost cell
/// it runs into.</param>
public readonly record struct ExcelCellDisplay(string Text, ExcelAlignment Alignment, int X, int Width, int ClipLeft, int ClipRight);
