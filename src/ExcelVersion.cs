namespace MacExcelReader;

/// <summary>
/// Excel 1.x releases that differ in the formulas they can store or in how they encode them (spec §9).
/// </summary>
public enum ExcelVersion
{
    /// <summary>Excel 1.00 to 1.03 (including 1.01): 208 functions (0–207). Their function tables are identical.</summary>
    Excel100 = 100,

    /// <summary>Excel 1.04: adds CHANGE.LINK (208) and OPEN's third argument; otherwise encodes formulas as 1.00–1.03 do.</summary>
    Excel104 = 104,

    /// <summary>
    /// Excel 1.5: adds functions 209–291, changes the argument rules of twelve older ones and the token class of
    /// function calls inside IF and CHOOSE, and can write version-18 files.
    /// </summary>
    Excel150 = 150,
}
