using System.Text;

namespace Passage.Web.Services;

/// <summary>
/// The shape line from How the Slate Works: fifteen slots, eight sequences
/// with the seven turns between them. Derived from the Beat Board lanes on
/// every parse and stored nowhere. It shows what is known, which is the only
/// thing worth displaying — deliberately not a progress bar.
/// </summary>
public static class ShapeLine
{
    public const string Legend =
        "Shape line — eight sequences, seven turns between them. ▫ sequence not written · ▪ sequence written · · turn unknown · ● turn known · – turn dropped (write – as its value)";

    /// <summary>Null when the script has no Sequence heading at all.</summary>
    public static string? Derive(IReadOnlyList<BoardLane> lanes)
    {
        var sequences = lanes
            .SelectMany(lane => lane.Groups)
            .Where(group => group.SequenceCard is not null)
            .Take(8)
            .ToList();
        if (sequences.Count == 0)
        {
            return null;
        }

        var turns = SplitScript.Turns.ToArray();
        var line = new StringBuilder(15);
        for (var i = 0; i < 8; i++)
        {
            var group = i < sequences.Count ? sequences[i] : null;
            var written = group is not null && group.Cards.Any(card => card.Kind is "Scene" or "Section");
            line.Append(written ? '▪' : '▫');
            if (i == 7)
            {
                break;
            }

            line.Append(TurnGlyph(group?.SequenceCard?.Description, turns[i]));
        }

        return line.ToString();
    }

    private static char TurnGlyph(string? description, SplitScript.Slot turn)
    {
        if (description is null)
        {
            return '·';
        }

        foreach (var line in description.Split('\n'))
        {
            if (!SplitScript.TryParse(line, out var label, out var value) || label != turn.Label)
            {
                continue;
            }

            if (SplitScript.IsDropped(value))
            {
                return '–';
            }

            return SplitScript.IsKnown(value) ? '●' : '·';
        }

        return '·';
    }
}
