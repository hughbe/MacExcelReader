namespace MacExcelReader;

/// <summary>
/// Excel 1.x releases that differ in the formulas they can store or in how they encode them (spec §2.1, §15).
/// </summary>
public enum ExcelVersion
{
    /// <summary>
    /// Excel 1.00 to 1.03 (including 1.01): 208 functions (0–207). Their function tables are identical, and for
    /// formulas this value stands for all of them.
    /// </summary>
    Excel100 = 100,

    /// <summary>
    /// Excel 1.03: encodes formulas exactly as 1.00 does (<see cref="Excel100"/>), but draws text in a font size the
    /// System has no bitmap for with a pixel more per character (spec Appendix D).
    /// </summary>
    Excel103 = 103,

    /// <summary>
    /// Excel 1.04 and 1.06: add CHANGE.LINK (208) and OPEN's third argument; otherwise encode formulas as 1.00–1.03 do.
    /// </summary>
    Excel104 = 104,

    /// <summary>
    /// Excel 1.5: adds functions 209–291, changes the argument rules of twelve older ones and the token class of
    /// function calls inside IF and CHOOSE, and can write version-18 files.
    /// </summary>
    Excel150 = 150,
}

/// <summary>
/// The format version in a document's header (spec §2.1).
/// </summary>
public enum ExcelFormatVersion : ushort
{
    /// <summary>Version 17: written by every Excel 1.x, and readable by every one.</summary>
    Version17 = 17,

    /// <summary>
    /// Version 18: written by Excel 1.5 when a document has frozen panes, a gridline colour or a number format only
    /// it can read. Only Excel 1.5 opens it.
    /// </summary>
    Version18 = 18,
}
