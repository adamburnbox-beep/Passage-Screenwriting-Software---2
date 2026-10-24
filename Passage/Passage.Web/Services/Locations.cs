using System.Text.RegularExpressions;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>One place in the script, from first appearance. Scenes are the
/// headings that normalise to it, in script order.</summary>
public sealed record LocationEntry(string Name, int FirstLine, IReadOnlyList<LocationScene> Scenes);

public sealed record LocationScene(string Heading, int LineNumber);

/// <summary>
/// Locations derived from scene headings (docs/STORY-PLAN.md, decision C).
/// Pure, so the table of headings is testable. Not <see cref="Characters.BareName"/>:
/// that normalises cues, this normalises places.
/// </summary>
public static class Locations
{
    private static readonly Regex IdMarker = new(@"\[\[id:[a-f\d\-]+\]\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SceneNumber = new(@"\s*#[\w.\-]+#\s*$", RegexOptions.Compiled);
    private static readonly Regex Spaces = new(@"\s+", RegexOptions.Compiled);

    // INT/EXT is ignored on purpose: the list groups by place, not by side.
    private static readonly Regex Prefix = new(
        @"^(?:INT\.?/EXT\.?|EXT\.?/INT\.?|I/E\.?|EST\.?|INT\.?|EXT\.?)(?=\s|$)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    // A dash standing alone between spaces (or at an edge), so HYPHEN-ATED survives.
    private static readonly Regex Separator = new(@"(?<=^|\s)[-—](?=\s|$)", RegexOptions.Compiled);

    private static readonly HashSet<string> TimesOfDay = new(StringComparer.Ordinal)
    {
        "DAY", "NIGHT", "MORNING", "AFTERNOON", "EVENING", "DAWN", "DUSK", "SUNRISE", "SUNSET",
        "LATER", "MOMENTS LATER", "CONTINUOUS", "SAME", "SAME TIME", "MAGIC HOUR",
    };

    /// <summary>The place a scene heading names, uppercased; empty when it names none.</summary>
    public static string Normalise(string heading)
    {
        var text = SceneNumber.Replace(IdMarker.Replace(heading, " "), string.Empty);
        text = Spaces.Replace(text, " ").Trim().ToUpperInvariant();

        text = text.StartsWith('.') ? text[1..] : Prefix.Replace(text, string.Empty, 1);

        var segments = Separator.Split(text)
            .Select(segment => segment.Trim())
            .Where(segment => segment.Length > 0)
            .ToList();
        while (segments.Count > 0 && TimesOfDay.Contains(segments[^1]))
        {
            segments.RemoveAt(segments.Count - 1);
        }

        return string.Join(" - ", segments);
    }

    /// <summary>Every location in order of first appearance. Headings that
    /// name no place are left out.</summary>
    public static IReadOnlyList<LocationEntry> Build(IEnumerable<ScreenplayElement> elements)
    {
        var order = new List<string>();
        var scenes = new Dictionary<string, List<LocationScene>>(StringComparer.Ordinal);
        foreach (var scene in elements.OfType<SceneHeadingElement>())
        {
            var name = Normalise(scene.Text);
            if (name.Length == 0)
            {
                continue;
            }

            if (!scenes.TryGetValue(name, out var list))
            {
                scenes[name] = list = new List<LocationScene>();
                order.Add(name);
            }

            list.Add(new LocationScene(scene.Heading, scene.LineNumber));
        }

        return order.Select(name => new LocationEntry(name, scenes[name][0].LineNumber, scenes[name])).ToList();
    }
}
