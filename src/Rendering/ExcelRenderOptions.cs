namespace MacExcelReader.Rendering;

/// <summary>
/// What <see cref="ExcelRenderer"/> draws.
/// </summary>
public sealed class ExcelRenderOptions
{
    /// <summary>Gets or sets which of the document's saved windows to draw (0, the front one, by default).</summary>
    public int Window { get; set; }

    /// <summary>
    /// Gets or sets whether to draw the window's frame: its title bar with the title, close box and stripes of an
    /// active window, and its one-pixel border and shadow. Without it only the window's contents are drawn.
    /// </summary>
    public bool Frame { get; set; } = true;

    /// <summary>Gets or sets the title in the title bar; Excel shows the document's file name. "Worksheet" by default.</summary>
    public string? Title { get; set; }

    /// <summary>
    /// Gets or sets the size of the window's contents in pixels, scroll bars included: 64 to 16384. By default it is
    /// the size of the window's saved rectangle (spec §13), at most 2048 pixels each way, as Excel fits a window to the
    /// screen.
    /// </summary>
    public int? Width { get; set; }

    /// <inheritdoc cref="Width"/>
    public int? Height { get; set; }

    /// <summary>Gets or sets whether to show the window's selection and active cell, as the front window does.</summary>
    public bool Selection { get; set; } = true;

    /// <summary>Gets or sets the release whose drawing to follow (spec Appendix D). Excel 1.00 by default.</summary>
    public ExcelVersion Version { get; set; } = ExcelVersion.Excel100;
}
