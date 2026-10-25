using Passage.Export;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>Where pagination put one scene: its heading, its 1-based source
/// line, and its page, or 0 when its heading could not be found on a page.</summary>
public sealed record ScenePlacement(string Heading, int LineNumber, int Page);

/// <summary>
/// Where the script's own scenes land on the paginated pages, which the Arc's
/// Pages axis plots against. Reads the script; never writes it.
/// </summary>
public static class BeatSheet
{
    public const int DefaultTargetPages = 110;

    /// <summary>
    /// Finds the page each scene heading lands on. The layout does not carry
    /// source lines, but scene headings are the only bold-left lines it
    /// produces and it emits them in script order, so each heading is matched
    /// to the next bold line its rendered text begins with. Wrapped headings
    /// span several bold lines, which are consumed with it. A heading that
    /// cannot be matched gets page 0 rather than a guess.
    /// </summary>
    public static IReadOnlyList<ScenePlacement> MapScenes(
        IReadOnlyList<ScreenplayElement> elements, IReadOnlyList<PreviewPage> pages)
    {
        var headingLines = pages
            .Where(page => !page.IsTitlePage)
            .SelectMany(page => page.Lines
                .Where(line => line.Style == LayoutTextStyle.LeftBold)
                .Select(line => (page.PageNumber, Text: Squash(line.Text))))
            .Where(line => line.Text.Length > 0)
            .ToList();

        var placements = new List<ScenePlacement>();
        var cursor = 0;
        foreach (var heading in elements.OfType<SceneHeadingElement>())
        {
            var label = FountainMarkup.StripOmissions(heading.Text).Trim();
            var rendered = Squash(label.ToUpperInvariant());
            var page = 0;
            if (rendered.Length > 0)
            {
                var match = cursor;
                while (match < headingLines.Count && !rendered.StartsWith(headingLines[match].Text, StringComparison.Ordinal))
                {
                    match++;
                }

                if (match < headingLines.Count)
                {
                    page = headingLines[match].PageNumber;
                    var consumed = headingLines[match].Text;
                    cursor = match + 1;
                    while (cursor < headingLines.Count
                        && consumed.Length < rendered.Length
                        && rendered.StartsWith(consumed + headingLines[cursor].Text, StringComparison.Ordinal))
                    {
                        consumed += headingLines[cursor].Text;
                        cursor++;
                    }
                }
            }

            placements.Add(new ScenePlacement(label, heading.LineNumber, page));
        }

        return placements;
    }

    private static string Squash(string text) =>
        string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
}
