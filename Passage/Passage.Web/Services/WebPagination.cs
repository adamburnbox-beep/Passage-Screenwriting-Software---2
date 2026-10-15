using Passage.Export;

namespace Passage.Web.Services;

/// <summary>
/// The pagination the web preview and PDF export share. The desktop apps call
/// <see cref="ScreenplayLayoutBuilder.BuildPages"/> without options and keep the fixed cut.
/// </summary>
public static class WebPagination
{
    public static PaginationOptions Options { get; } = new();
}
