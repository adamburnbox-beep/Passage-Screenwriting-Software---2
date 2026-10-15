using Passage.Export;
using Passage.Parser;

namespace Passage.Web.Services;

public readonly record struct ExportResult(byte[] Bytes, string ContentType, string FileName);

/// <summary>
/// Bridges the file-based <see cref="IExporter"/> catalog to HTTP downloads by
/// exporting to a temp file and streaming the bytes back.
/// </summary>
public sealed class ExportService
{
    private readonly FountainParser _parser = new();

    // The catalog's PDF exporter is shared with the desktop apps; swap in one that uses web pagination.
    public IReadOnlyList<IExporter> Exporters { get; } = ExporterCatalog.GetDefaultExporters()
        .Select(exporter => exporter is ScreenplayPdfExporter ? new ScreenplayPdfExporter(WebPagination.Options) : exporter)
        .ToList();

    public ExportResult? Export(string content, string format, string? documentName)
    {
        var extension = format.StartsWith('.') ? format : "." + format;
        var baseName = string.IsNullOrWhiteSpace(documentName)
            ? "Untitled"
            : Path.GetFileNameWithoutExtension(documentName.Trim());

        // .highland is a zip of the source text, not a rendering of the parse,
        // so it stays out of the shared exporter catalog.
        if (extension.Equals(HighlandBundle.Extension, StringComparison.OrdinalIgnoreCase))
        {
            return new ExportResult(HighlandBundle.Write(content, baseName), "application/zip", baseName + HighlandBundle.Extension);
        }

        var exporter = Exporters.FirstOrDefault(candidate =>
            candidate.DefaultExtension.Equals(extension, StringComparison.OrdinalIgnoreCase));
        if (exporter is null)
        {
            return null;
        }

        var parsed = _parser.Parse(content);
        var tempPath = Path.Combine(Path.GetTempPath(), $"passage-export-{Guid.NewGuid():N}{exporter.DefaultExtension}");
        try
        {
            exporter.Export(parsed, tempPath);
            var bytes = File.ReadAllBytes(tempPath);

            var contentType = exporter.DefaultExtension.ToLowerInvariant() switch
            {
                ".pdf" => "application/pdf",
                ".txt" => "text/plain",
                _ => "application/octet-stream"
            };

            return new ExportResult(bytes, contentType, baseName + exporter.DefaultExtension);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
