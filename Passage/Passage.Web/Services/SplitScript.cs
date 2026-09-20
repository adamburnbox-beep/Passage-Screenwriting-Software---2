using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>
/// How an A→Z→Split run is written into the script, and read back out of it.
/// Eight sequences with the seven turns between them (How the Slate Works,
/// "the shape line"): each turn is a <c>= Name: text</c> synopsis on the
/// sequence it closes, and a turn not found yet is a bracket, so Fill hands
/// it back later. Pure, so the shape line and the write-back can be tested
/// without the editor.
/// </summary>
public static class SplitScript
{
    /// <summary>A line the run owns: its label, the sequence it sits on (1–8)
    /// and the placeholder written while the turn is unknown.</summary>
    public sealed record Slot(string Label, int Sequence, string Bracket);

    // Story order. The ends are not turns, but they are lines the run owns.
    public static readonly Slot[] Slots =
    {
        new("Starts", 1, "[START EVENT]"),
        new("Inciting incident", 1, "[INCITING INCIDENT]"),
        new("Plot point 1", 2, "[PLOT POINT 1]"),
        new("Pinch 1", 3, "[PINCH 1]"),
        new("Midpoint", 4, "[MIDPOINT]"),
        new("Pinch 2", 5, "[PINCH 2]"),
        new("Crisis / Plot point 2", 6, "[CRISIS]"),
        new("Last obstacle", 7, "[LAST OBSTACLE]"),
        new("Ends", 8, "[END EVENT]")
    };

    /// <summary>The seven turns, left to right, as the shape line orders them.</summary>
    public static IEnumerable<Slot> Turns => Slots.Skip(1).Take(7);

    public static readonly string[] Acts = { "Act 1", "Act 2A", "Act 2B", "Act 3" };

    public static string[] Values(SplitRun run) => new[]
    {
        run.A, run.Inciting, run.PP1, run.Pinch1, run.Midpoint, run.Pinch2, run.Crisis, run.LastObstacle, run.Z
    };

    /// <summary>The synopsis line for a slot: the text as one line, or the bracket.</summary>
    public static string Line(Slot slot, string value)
    {
        var text = OneLine(value);
        return $"= {slot.Label}: {(text.Length == 0 ? slot.Bracket : text)}";
    }

    /// <summary>
    /// The whole structure for a script that has none yet: four acts, eight
    /// sequences, every slot's line. Blank lines separate headings so the
    /// text reads as Fountain, not as a dump.
    /// </summary>
    public static List<string> BuildLines(SplitRun run)
    {
        var values = Values(run);
        var lines = new List<string>();
        for (var sequence = 1; sequence <= 8; sequence++)
        {
            if (sequence % 2 == 1)
            {
                lines.Add(string.Empty);
                lines.AddRange(BeatBoardText.BuildCardLines("Act", Acts[(sequence - 1) / 2], string.Empty, Guid.NewGuid()));
            }

            lines.Add(string.Empty);
            lines.AddRange(BeatBoardText.BuildCardLines("Sequence", $"Sequence {sequence}", string.Empty, Guid.NewGuid()));
            for (var i = 0; i < Slots.Length; i++)
            {
                if (Slots[i].Sequence == sequence)
                {
                    lines.Add(Line(Slots[i], values[i]));
                }
            }
        }

        return lines;
    }

    /// <summary>Index of the line that carries a slot, or -1.</summary>
    public static int FindLine(string[] lines, Slot slot)
    {
        for (var i = 0; i < lines.Length; i++)
        {
            if (TryParse(lines[i], out var label, out _) && label == slot.Label)
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Reads a slot line — <c>= Label: value</c>, or the same without the
    /// <c>=</c> as the board's card descriptions carry it. False for any
    /// other line.
    /// </summary>
    public static bool TryParse(string line, out string label, out string value)
    {
        label = string.Empty;
        value = string.Empty;
        var text = line.Trim();
        if (text.StartsWith('='))
        {
            text = text[1..].TrimStart();
        }

        var colon = text.IndexOf(':');
        if (colon < 0)
        {
            return false;
        }

        var candidate = text[..colon].Trim();
        foreach (var slot in Slots)
        {
            if (string.Equals(candidate, slot.Label, StringComparison.OrdinalIgnoreCase))
            {
                label = slot.Label;
                value = text[(colon + 1)..].Trim();
                return true;
            }
        }

        return false;
    }

    /// <summary>A turn is known when its value is real text: not empty, not a
    /// bracket, not the dash that marks a dropped turn.</summary>
    public static bool IsKnown(string value) => value.Length > 0 && !value.StartsWith('[') && !IsDropped(value);

    public static bool IsDropped(string value) => value is "–" or "-" or "—";

    public static string OneLine(string value) =>
        string.Join(' ', (value ?? string.Empty).Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(part => part.Trim()));
}
