using System.Globalization;
using System.Text.RegularExpressions;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>
/// Add / Remove Scene Numbers (HIGHLAND-PLAN H1.3). Both commands rewrite the
/// <c>#…#</c> on scene heading lines and nothing else. Pure: they return the
/// line edits, and the editor applies them as one undoable change, checking
/// each line still reads <see cref="LineEdit.Expected"/> first.
/// </summary>
public static class SceneNumbers
{
    public sealed record LineEdit(int LineIndex, string Expected, string Replacement);

    // The number and the space before it, when it is the last thing on the line
    // apart from an optional [[id:…]] comment (which the parser also strips).
    private static readonly Regex NumberOnLine = new(
        @"\s*#[A-Za-z0-9.\-]+#(?=\s*(?:\[\[id:[a-f\d\-]+\]\])?\s*$)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex TrailingId = new(
        @"\s*\[\[id:[a-f\d\-]+\]\]\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex LeadingDigits = new(@"^\d+", RegexOptions.CultureInvariant);

    public static IReadOnlyList<LineEdit> Remove(string content)
    {
        var lines = content.Split('\n');
        var edits = new List<LineEdit>();
        foreach (var heading in Headings(content))
        {
            if (heading.SceneNumber is null || heading.LineIndex >= lines.Length) continue;
            var line = lines[heading.LineIndex];
            var stripped = NumberOnLine.Replace(line, string.Empty, 1);
            if (stripped != line) edits.Add(new LineEdit(heading.LineIndex, line, stripped));
        }

        return edits;
    }

    /// <summary>
    /// Numbers every unnumbered scene and leaves numbered ones alone. A scene
    /// takes the number after the one before it; when that number is taken or
    /// would run into the next numbered scene, it gets a letter instead
    /// (5, new, 6 → 5A), the usual locked-script convention.
    /// </summary>
    public static IReadOnlyList<LineEdit> Add(string content)
    {
        var lines = content.Split('\n');
        var headings = Headings(content);
        var used = new HashSet<string>(
            headings.Where(h => h.SceneNumber is not null).Select(h => h.SceneNumber!),
            StringComparer.OrdinalIgnoreCase);
        var edits = new List<LineEdit>();
        string? previous = null;

        for (var i = 0; i < headings.Count; i++)
        {
            var heading = headings[i];
            if (heading.SceneNumber is { } existing)
            {
                previous = existing;
                continue;
            }

            if (heading.LineIndex >= lines.Length) continue;
            var next = headings.Skip(i + 1).FirstOrDefault(h => h.SceneNumber is not null)?.SceneNumber;
            var number = NextNumber(previous, next, used);
            used.Add(number);
            previous = number;

            var line = lines[heading.LineIndex];
            var id = TrailingId.Match(line);
            var body = id.Success ? line[..id.Index] : line.TrimEnd();
            edits.Add(new LineEdit(heading.LineIndex, line, $"{body} #{number}#{(id.Success ? id.Value : string.Empty)}"));
        }

        return edits;
    }

    private static List<SceneHeadingElement> Headings(string content) =>
        new FountainParser().Parse(content).Elements.OfType<SceneHeadingElement>().ToList();

    private static string NextNumber(string? previous, string? next, HashSet<string> used)
    {
        var previousValue = previous is null ? 0 : LeadingValue(previous);
        var candidate = (previousValue + 1).ToString(CultureInfo.InvariantCulture);
        var fitsBeforeNext = next is null || LeadingValue(next) > previousValue + 1 || LeadingValue(next) < 0;
        if (previousValue >= 0 && fitsBeforeNext && !used.Contains(candidate)) return candidate;

        // Letter after the previous number: 5 → 5A, 5A → 5B.
        var stem = previous ?? "0";
        var last = stem[^1];
        string letterBase;
        char letter;
        if (char.IsAsciiLetter(last))
        {
            letterBase = stem[..^1];
            letter = (char)(char.ToUpperInvariant(last) + 1);
        }
        else
        {
            letterBase = stem;
            letter = 'A';
        }

        for (; letter <= 'Z'; letter++)
        {
            var lettered = letterBase + letter;
            if (!used.Contains(lettered)) return lettered;
        }

        // Past Z: a counter keeps it unique rather than reusing a number.
        for (var suffix = 1; ; suffix++)
        {
            var fallback = $"{stem}{suffix.ToString(CultureInfo.InvariantCulture)}";
            if (!used.Contains(fallback)) return fallback;
        }
    }

    // The leading integer of a scene number ("12A" → 12); -1 when it has none.
    private static int LeadingValue(string number)
    {
        var match = LeadingDigits.Match(number);
        return match.Success && int.TryParse(match.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value)
            ? value
            : -1;
    }
}
