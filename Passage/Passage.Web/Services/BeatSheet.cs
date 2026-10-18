using Passage.Export;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>Where pagination put one scene: its heading, its 1-based source
/// line, and its page, or 0 when its heading could not be found on a page.</summary>
public sealed record ScenePlacement(string Heading, int LineNumber, int Page);

/// <summary>An act (a <c>#</c> section) and the page its first scene lands on.</summary>
public sealed record ActBreak(string Heading, int Page);

/// <summary>
/// The rules behind the Beat Sheet view: the size of the page grid, where a
/// beat may sit on it, how overlapping beats stack, and where the script's
/// own scenes land. Beats are planning and never read or write the script.
/// </summary>
public static class BeatSheet
{
    public const int DefaultTargetPages = 110;
    public const int MaxTargetPages = 999;

    /// <summary>The grid is the larger of the script's length and the target.</summary>
    public static int Columns(int pageCount, int targetPages) =>
        Math.Max(1, Math.Max(pageCount, targetPages));

    /// <summary>
    /// Keeps a beat on the grid: it starts no earlier than page 1 and no later
    /// than the last column, spans at least one page, and never runs past the
    /// last column.
    /// </summary>
    public static (int Page, int Span) Clamp(int page, int span, int columns)
    {
        columns = Math.Max(1, columns);
        page = Math.Clamp(page, 1, columns);
        span = Math.Clamp(span, 1, columns - page + 1);
        return (page, span);
    }

    /// <summary>
    /// Where a moved beat lands. A move keeps the beat's length where the grid
    /// can hold it, so a beat dropped near the end slides back to fit rather
    /// than being cut short.
    /// </summary>
    public static (int Page, int Span) ClampMove(int page, int span, int columns)
    {
        columns = Math.Max(1, columns);
        span = Math.Clamp(span, 1, columns);
        page = Math.Clamp(page, 1, columns - span + 1);
        return (page, span);
    }

    /// <summary>The last page any placed beat covers, or 0 when none is placed.</summary>
    public static int LastBeatPage(IEnumerable<Beat> beats) =>
        beats.Where(beat => beat.Page > 0).Select(beat => beat.Page + beat.Span - 1).DefaultIfEmpty(0).Max();

    /// <summary>
    /// Assigns each placed beat a row so overlapping beats stack instead of
    /// covering each other. Earlier pages take the first free row; beats on
    /// the same page keep their list order.
    /// </summary>
    public static IReadOnlyList<(Beat Beat, int Row)> Stack(IReadOnlyList<Beat> beats)
    {
        var rowEnds = new List<int>();
        var placed = new List<(Beat, int)>();
        var ordered = beats
            .Select((beat, index) => (beat, index))
            .Where(item => item.beat.Page > 0)
            .OrderBy(item => item.beat.Page)
            .ThenBy(item => item.index);

        foreach (var (beat, _) in ordered)
        {
            var row = rowEnds.FindIndex(end => end < beat.Page);
            if (row < 0)
            {
                row = rowEnds.Count;
                rowEnds.Add(0);
            }

            rowEnds[row] = beat.Page + Math.Max(1, beat.Span) - 1;
            placed.Add((beat, row));
        }

        return placed;
    }

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

    /// <summary>Each top-level section that has a placed scene after it, at
    /// that scene's page.</summary>
    public static IReadOnlyList<ActBreak> ActBreaks(
        IReadOnlyList<ScreenplayElement> elements, IReadOnlyList<ScenePlacement> scenes)
    {
        var breaks = new List<ActBreak>();
        foreach (var section in elements.OfType<SectionElement>().Where(section => section.SectionDepth == 1))
        {
            var first = scenes.FirstOrDefault(scene => scene.Page > 0 && scene.LineNumber > section.LineNumber);
            if (first is not null)
            {
                breaks.Add(new ActBreak(section.Text.Trim(), first.Page));
            }
        }

        return breaks;
    }

    private static string Squash(string text) =>
        string.Concat(text.Where(character => !char.IsWhiteSpace(character)));
}
