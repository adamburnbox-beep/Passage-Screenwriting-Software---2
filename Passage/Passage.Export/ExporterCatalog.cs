namespace Passage.Export;

public static class ExporterCatalog
{
    private static readonly IExporter[] DefaultExporters =
    {
        new ScreenplayExporter(),
        new ScreenplayPdfExporter()
    };

    public static IReadOnlyList<IExporter> GetDefaultExporters()
    {
        return DefaultExporters;
    }
}
