namespace Passage.Web.Services;

/// <summary>One line of the Beside pane's Board list: a lane or group header,
/// or a card. Line is the heading's line in the script, 0 for the untitled
/// "Script" lane that has no heading.</summary>
public sealed record BesideRow(string Role, string Text, int Line, int Depth, BoardCard? Card);

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
}
