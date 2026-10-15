using System.IO.Compression;
using System.Text;

namespace Passage.Web.Services;

/// <param name="Text">The script text, byte-for-byte as it was in the bundle.</param>
/// <param name="IsMarkdown">True when the bundle held text.md rather than text.fountain.</param>
/// <param name="Skipped">Every zip entry that was not imported, by its path in the zip.</param>
public sealed record HighlandImport(string Text, bool IsMarkdown, IReadOnlyList<string> Skipped);

/// <summary>
/// Reads and writes Highland 2 <c>.highland</c> files: a zip holding one
/// TextBundle directory (<c>Name.textbundle/text.fountain</c> plus JSON
/// sidecars). Only the script text crosses over; every other entry is
/// reported back to the caller rather than silently dropped.
/// </summary>
public static class HighlandBundle
{
    public const string Extension = ".highland";

    // Highland's sidecars hold app state Passage has no model for. A text
    // file larger than this is refused rather than inflated into memory.
    private const long MaxTextBytes = 5 * 1024 * 1024;
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private const string InfoJson = """
        {
          "version" : 2,
          "type" : "com.quoteunquoteapps.fountain",
          "transient" : false,
          "app.passage" : {
          }
        }
        """;

    public static byte[] Write(string text, string baseName)
    {
        var bundle = baseName + ".textbundle/";
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(zip, bundle + "text.fountain", text);
            WriteEntry(zip, bundle + "info.json", InfoJson);
            WriteEntry(zip, bundle + "characters.json", "{}");
            WriteEntry(zip, bundle + "sprints.json", "[]");
        }

        return buffer.ToArray();
    }

    /// <exception cref="InvalidDataException">
    /// The stream is not a zip, holds no text.fountain or text.md, or the text
    /// is too large.
    /// </exception>
    public static HighlandImport Read(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        // Prefer Fountain over Markdown; within each, the first bundle wins.
        var textEntry = FindText(zip, "text.fountain") ?? FindText(zip, "text.md")
            ?? throw new InvalidDataException("No text.fountain or text.md inside a .textbundle.");

        if (textEntry.Length > MaxTextBytes)
        {
            throw new InvalidDataException($"The script text is larger than {MaxTextBytes / (1024 * 1024)} MB.");
        }

        string text;
        using (var reader = new StreamReader(textEntry.Open(), Encoding.UTF8))
        {
            text = reader.ReadToEnd();
        }

        var skipped = zip.Entries
            .Where(entry => entry != textEntry && !IsDirectory(entry))
            .Select(entry => entry.FullName)
            .ToList();

        return new HighlandImport(text, textEntry.Name.Equals("text.md", StringComparison.OrdinalIgnoreCase), skipped);
    }

    // The text file sits directly inside a *.textbundle directory, which is
    // either at the zip root or (zipped from a parent folder) one level down.
    private static ZipArchiveEntry? FindText(ZipArchive zip, string fileName) =>
        zip.Entries.FirstOrDefault(entry =>
        {
            if (!entry.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var parts = entry.FullName.Split('/');
            return parts.Length is 2 or 3
                && parts[^2].EndsWith(".textbundle", StringComparison.OrdinalIgnoreCase);
        });

    private static bool IsDirectory(ZipArchiveEntry entry) =>
        entry.FullName.EndsWith('/') && entry.Length == 0;

    private static void WriteEntry(ZipArchive zip, string path, string content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), Utf8NoBom);
        writer.Write(content);
    }
}
