namespace MacExcelReader.Internal;

/// <summary>
/// Reads the rest of a stream into an array of exactly that length, for the readers, which work on whole files.
/// </summary>
internal static class StreamBytes
{
    public static byte[] ReadToEnd(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (stream.CanSeek)
        {
            var buffer = new byte[Math.Max(0, stream.Length - stream.Position)];
            stream.ReadExactly(buffer);
            return buffer;
        }

        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    public static async Task<byte[]> ReadToEndAsync(Stream stream, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (stream.CanSeek)
        {
            var buffer = new byte[Math.Max(0, stream.Length - stream.Position)];
            await stream.ReadExactlyAsync(buffer, cancellationToken).ConfigureAwait(false);
            return buffer;
        }

        using var copy = new MemoryStream();
        await stream.CopyToAsync(copy, cancellationToken).ConfigureAwait(false);
        return copy.ToArray();
    }
}
