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

    /// <inheritdoc/>
    public override string ToString() => Name;
}
