using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>
/// Turning an inbox beat into a Board card (docs/STORY-PLAN.md, decision E).
/// Pure functions, so the insert point and the tag transfer are testable
/// without a browser.
/// </summary>
public static class BeatInbox
{
    private const string UntitledHeading = "UNTITLED BEAT";

    /// <summary>The Fountain lines for the new scene: <c>. TITLE [[id:…]]</c> and
    /// one <c>=</c> line per non-blank description line, the same as the Board's
    /// own scene cards. A beat with no title gets a placeholder heading, since
    /// a bare "." is not a scene.</summary>
    public static List<string> SceneLines(Beat beat, Guid id) =>
        BeatBoardText.BuildCardLines("Scene", Heading(beat), beat.Description, id);

    public static string Heading(Beat beat) =>
        beat.Title.Trim() is { Length: > 0 } title ? title : UntitledHeading;

    /// <summary>Where the card's lines go, and the text to hand to
    /// <c>passage.insertLinesAt</c>. Before a card is its first line; after is
    /// the line past its whole block (nested content included), the ranges
    /// <see cref="BeatBoardText.TryPlanMove"/> uses for a drag. False when the
    /// target is no longer in the script.</summary>
    public static bool TryPlan(
        IReadOnlyList<ScreenplayElement> elements,
        IReadOnlyList<string> scriptLines,
        Guid targetId,
        bool after,
        IReadOnlyList<string> cardLines,
        out int insertAt,
        out string text)
    {
        insertAt = 0;
        text = string.Empty;

        var range = BeatBoardText.GetCardLineRange(elements, targetId, scriptLines.Count, includeNestedBlock: true);
        if (!range.IsFound)
        {
            return false;
        }

        insertAt = after ? Math.Min(range.EndLineIndex + 1, scriptLines.Count) : range.StartLineIndex;
        text = Padded(scriptLines, insertAt, cardLines);
        return true;
    }

    /// <summary>The end of the script, for "Send to Board" with no drop target.</summary>
    public static string TextAtEnd(IReadOnlyList<string> scriptLines, IReadOnlyList<string> cardLines) =>
        Padded(scriptLines, scriptLines.Count, cardLines);

    // A blank line separates the new card from what it touches: before it when
    // the line above has text, after it unless it ends the script (a block
    // normally ends in its own blank line, so this rarely doubles up).
    private static string Padded(IReadOnlyList<string> scriptLines, int insertAt, IReadOnlyList<string> cardLines)
    {
        var padded = new List<string>();
        if (insertAt > 0 && insertAt <= scriptLines.Count && scriptLines[insertAt - 1].Trim().Length > 0)
        {
            padded.Add(string.Empty);
        }

        padded.AddRange(cardLines);
        if (insertAt < scriptLines.Count)
        {
            padded.Add(string.Empty);
        }

        return string.Join("\n", padded);
    }

    /// <summary>Re-keys the beat's tags and star to the new card. False when the
    /// beat carried none.</summary>
    public static bool MoveTags(StoryData story, Guid beatId, Guid cardId, string label)
    {
        var item = story.Items.FirstOrDefault(entry => entry.Kind == StoryItems.BeatKind && entry.TargetId == beatId);
        if (item is null)
        {
            return false;
        }

        item.TargetId = cardId;
        item.Kind = StoryItems.CardKind;
        item.Label = label;
        return true;
    }
}
