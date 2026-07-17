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

    public IReadOnlyList<IExporter> Exporters { get; } = ExporterCatalog.GetDefaultExporters();

    public ExportResult? Export(string content, string format, string? documentName)
    {
        var extension = format.StartsWith('.') ? format : "." + format;
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
            var baseName = string.IsNullOrWhiteSpace(documentName)
                ? "Untitled"
                : Path.GetFileNameWithoutExtension(documentName.Trim());

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
