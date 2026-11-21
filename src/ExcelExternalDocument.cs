namespace MacExcelReader;

/// <summary>
/// An entry of the external-document table (spec §6.2): a document this sheet refers to.
/// </summary>
public sealed class ExcelExternalDocument
{
    /// <summary>
    /// Creates an entry.
    /// </summary>
    public ExcelExternalDocument(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Name = name;
        Reserved = new byte[6];
    }

    /// <summary>
    /// Gets or sets the document name. An empty name means "this document".
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets the six unidentified bytes that precede the name.
    /// </summary>
    public byte[] Reserved { get; internal set; }

    /// <summary>
    /// Gets the names this sheet uses from the document (spec §10). A name token inside an external
    /// reference to this document holds the 1-based index of its entry here.
    /// </summary>
    public List<string> Names { get; } = [];

    /// <inheritdoc/>
    public override string ToString() => Name;
}
