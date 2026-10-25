namespace Passage.Web.Services;

/// <summary>One line of the Beside pane's Board list: a lane or group header,
/// or a card. Line is the heading's line in the script, 0 for the untitled
/// "Script" lane that has no heading.</summary>
public sealed record BesideRow(string Role, string Text, int Line, int Depth, BoardCard? Card);

/// <summary>One line of the Beside pane's Beats list. Line is the scene to jump
/// to (0 for an unassigned beat, which has no place in the script).</summary>
public sealed record BesideBeat(Beat Beat, int Line);

/// <summary>
/// The compact lists behind SCRIPT's Beside pane (docs/STORY-PLAN.md S2.3).
/// Pure, so the order and the caret lookups are testable without a browser.
/// </summary>
public static class BesideList
{
    public const string LaneRole = "lane";
    public const string GroupRole = "group";
    public const string CardRole = "card";

    /// <summary>The Board's lanes in the Board's own order: lane header, then
    /// each group header and its cards, nested cards after their parent.</summary>
    public static List<BesideRow> Rows(IReadOnlyList<BoardLane> lanes)
    {
        var rows = new List<BesideRow>();
        foreach (var lane in lanes)
        {
            rows.Add(new BesideRow(LaneRole, lane.ActCard?.Heading ?? "Script", lane.ActCard?.LineNumber ?? 0, 0, lane.ActCard));
            foreach (var group in lane.Groups)
            {
                if (group.SequenceCard is { } sequence)
                {
                    rows.Add(new BesideRow(GroupRole, sequence.Heading, sequence.LineNumber, 1, sequence));
                }

                foreach (var card in group.Cards)
                {
                    AddCard(rows, card, 2);
                }
            }
        }

        return rows;
    }

    private static void AddCard(List<BesideRow> rows, BoardCard card, int depth)
    {
        rows.Add(new BesideRow(CardRole, card.Heading, card.LineNumber, depth, card));
        foreach (var child in card.Children)
        {
            AddCard(rows, child, depth + 1);
        }
    }

    /// <summary>The row the caret is in: the last row whose heading is at or
    /// above the caret line. -1 before any heading.</summary>
    public static int RowAt(IReadOnlyList<BesideRow> rows, int caretLine)
    {
        var found = -1;
        for (var i = 0; i < rows.Count; i++)
        {
            if (rows[i].Line > 0 && rows[i].Line <= caretLine)
            {
                found = i;
            }
        }

        return found;
    }

    /// <summary>Beats in page order (unassigned first, in their tray order),
    /// each with the scene it points at: the first scene on its page, or the
    /// last scene before it when its page starts none.</summary>
    public static List<BesideBeat> Beats(IReadOnlyList<Beat> beats, IReadOnlyList<ScenePlacement> scenes) =>
        beats.Select((beat, index) => (beat, index))
            .OrderBy(pair => pair.beat.Page)
            .ThenBy(pair => pair.index)
            .Select(pair => new BesideBeat(pair.beat, SceneFor(pair.beat.Page, scenes)))
            .ToList();

    private static int SceneFor(int page, IReadOnlyList<ScenePlacement> scenes)
    {
        if (page <= 0)
        {
            return 0;
        }

        var onPage = scenes.FirstOrDefault(scene => scene.Page == page);
        if (onPage is not null)
        {
            return onPage.LineNumber;
        }

        var before = scenes.Where(scene => scene.Page > 0 && scene.Page < page).LastOrDefault();
        return before?.LineNumber ?? 0;
    }

    /// <summary>The beat under the caret: the first placed beat whose pages
    /// cover the page of the scene the caret is in. -1 when the caret is before
    /// every scene, the scene has no page, or no beat covers it.</summary>
    public static int BeatAt(IReadOnlyList<BesideBeat> beats, IReadOnlyList<ScenePlacement> scenes, int caretLine)
    {
        var scene = scenes.Where(candidate => candidate.LineNumber <= caretLine).LastOrDefault();
        if (scene is null || scene.Page <= 0)
        {
            return -1;
        }

        for (var i = 0; i < beats.Count; i++)
        {
            var beat = beats[i].Beat;
            if (beat.Page > 0 && scene.Page >= beat.Page && scene.Page < beat.Page + Math.Max(1, beat.Span))
            {
                return i;
            }
        }

        return -1;
    }
}
