using System.Text.RegularExpressions;
using Passage.Core;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>One speaker in the character panel. Lines are speeches — one per
/// cue — and words are spoken words, parentheticals excluded.</summary>
public sealed record CharacterStat(string Name, int Lines, int Words);

/// <summary>Dialogue balance by gender tag (HIGHLAND-PLAN H3.2). Totals and
/// pairs follow <see cref="Characters.Tags"/> order; rows follow the panel.</summary>
public sealed record GenderReport(
    IReadOnlyList<GenderTotal> Totals,
    IReadOnlyList<GenderPair> Pairs,
    IReadOnlyList<GenderRow> Rows);

public sealed record GenderTotal(string Tag, int Characters, int Lines, int Words);

/// <summary>Scenes adds one for every two different speakers who both speak
/// in the same scene, filed under their two tags.</summary>
public sealed record GenderPair(string First, string Second, int Scenes);

public sealed record GenderRow(string Name, string Tag, int Lines, int Words);

/// <summary>
/// The character panel (HIGHLAND-PLAN H3.1). Pure: stats and per-line
/// speakers come from the parsed elements, and rename returns line edits the
/// editor applies as one undoable change, as <see cref="SceneNumbers"/> does.
/// </summary>
public static class Characters
{
    // Pastel highlight swatches. Stored as plain hex; the editor tints with
    // them translucently so dialogue stays readable in both themes.
    public static readonly string[] Palette =
    {
        "#9ec5f0", "#f2c46d", "#a8d08d", "#d3a6dc", "#f0a0a0",
        "#8fd6cf", "#f5b38a", "#c9c9a0", "#a9b4f2", "#e8a3c4",
    };

    /// <summary>The gender tags, in report order. Anything stored that is
    /// not Female or Male reads as Unspecified.</summary>
    public static readonly string[] Tags = { "Female", "Male", "Unspecified" };

    public static string Tag(string? stored) =>
        stored is "Female" or "Male" ? stored : "Unspecified";

    // A cue line: indent, optional forcing @, the name, then any trailing
    // extensions — (V.O.), (CONT'D), (ON PHONE)… — and the dual-dialogue ^.
    private static readonly Regex CueLine = new(
        @"^(?<lead>\s*@?)(?<name>.*?)(?<tail>(?:\s*\([^()]*\))*\s*\^?\s*)$",
        RegexOptions.CultureInvariant);

    /// <summary>The speaker a cue names, uppercased, with @, ^ and the
    /// parenthetical extensions removed.</summary>
    public static string BareName(string cue)
    {
        var match = CueLine.Match(cue);
        return (match.Success ? match.Groups["name"].Value : cue).Trim().ToUpperInvariant();
    }

    /// <summary>
    /// Every speaker with a cue, most speeches first, and per line the speaker
    /// whose cue or dialogue block covers it ("" elsewhere).
    /// </summary>
    public static (IReadOnlyList<CharacterStat> Stats, string[] LineSpeakers) Build(
        IReadOnlyList<ScreenplayElement> elements, int lineCount)
    {
        var lines = new Dictionary<string, int>(StringComparer.Ordinal);
        var words = new Dictionary<string, int>(StringComparer.Ordinal);
        var speakers = new string[lineCount];
        Array.Fill(speakers, string.Empty);
        CharacterElement? cue = null;

        foreach (var element in elements)
        {
            switch (element)
            {
                case CharacterElement character:
                {
                    cue = character;
                    var name = BareName(character.CharacterName);
                    if (name.Length > 0) lines.TryAdd(name, 0);
                    break;
                }

                case DialogueElement dialogue:
                {
                    var name = BareName(dialogue.CharacterName);
                    if (name.Length == 0) break;
                    lines[name] = lines.GetValueOrDefault(name) + 1;
                    words[name] = words.GetValueOrDefault(name) + TextAnalysis.CountWords(string.Join('\n', dialogue.Lines));

                    var start = cue is not null && BareName(cue.CharacterName) == name ? cue.LineIndex : dialogue.LineIndex;
                    for (var i = Math.Max(0, start); i <= dialogue.EndLineIndex && i < lineCount; i++)
                    {
                        speakers[i] = name;
                    }

                    cue = null;
                    break;
                }
            }
        }

        var stats = lines
            .Select(entry => new CharacterStat(entry.Key, entry.Value, words.GetValueOrDefault(entry.Key)))
            .OrderByDescending(stat => stat.Lines)
            .ThenBy(stat => stat.Name, StringComparer.Ordinal)
            .ToList();
        return (stats, speakers);
    }

    /// <summary>
    /// Rewrites the name on every cue line for <paramref name="oldName"/> and
    /// nothing else, keeping @, extensions and ^. The name is uppercased: the
    /// shared parser only reads all-caps lines as cues, @ or not.
    /// </summary>
    public static IReadOnlyList<SceneNumbers.LineEdit> Rename(
        string content, IReadOnlyList<ScreenplayElement> elements, string oldName, string newName)
    {
        var lines = content.Split('\n');
        var edits = new List<SceneNumbers.LineEdit>();
        var trimmed = newName.Trim();
        if (trimmed.Length == 0) return edits;

        foreach (var cue in elements.OfType<CharacterElement>())
        {
            if (cue.LineIndex >= lines.Length || BareName(cue.CharacterName) != oldName) continue;
            var line = lines[cue.LineIndex];
            var match = CueLine.Match(line);
            if (!match.Success) continue;

            var replacement = match.Groups["lead"].Value + trimmed.ToUpperInvariant() + match.Groups["tail"].Value;
            if (replacement != line) edits.Add(new SceneNumbers.LineEdit(cue.LineIndex, line, replacement));
        }

        return edits;
    }

    /// <summary>
    /// The dialogue balance report. Totals are summed from the same rows the
    /// table shows, so the two always reconcile. A scene starts at each
    /// heading; dialogue before the first heading is a scene of its own.
    /// </summary>
    public static GenderReport Report(
        IReadOnlyList<ScreenplayElement> elements,
        IReadOnlyList<CharacterStat> stats,
        IReadOnlyDictionary<string, string> genders)
    {
        var rows = stats
            .Select(stat => new GenderRow(stat.Name, Tag(genders.GetValueOrDefault(stat.Name)), stat.Lines, stat.Words))
            .ToList();
        var totals = Tags
            .Select(tag =>
            {
                var tagged = rows.Where(row => row.Tag == tag).ToList();
                return new GenderTotal(tag, tagged.Count, tagged.Sum(row => row.Lines), tagged.Sum(row => row.Words));
            })
            .ToList();

        var scenes = new List<HashSet<string>> { new(StringComparer.Ordinal) };
        foreach (var element in elements)
        {
            if (element is SceneHeadingElement)
            {
                scenes.Add(new HashSet<string>(StringComparer.Ordinal));
            }
            else if (element is DialogueElement dialogue && BareName(dialogue.CharacterName) is { Length: > 0 } name)
            {
                scenes[^1].Add(name);
            }
        }

        var tagOf = rows.ToDictionary(row => row.Name, row => row.Tag, StringComparer.Ordinal);
        var pairCounts = new int[Tags.Length, Tags.Length];
        foreach (var scene in scenes)
        {
            var speakers = scene.Order(StringComparer.Ordinal).ToList();
            for (var i = 0; i < speakers.Count; i++)
            {
                for (var j = i + 1; j < speakers.Count; j++)
                {
                    var a = Array.IndexOf(Tags, tagOf.GetValueOrDefault(speakers[i], "Unspecified"));
                    var b = Array.IndexOf(Tags, tagOf.GetValueOrDefault(speakers[j], "Unspecified"));
                    pairCounts[Math.Min(a, b), Math.Max(a, b)]++;
                }
            }
        }

        var pairs = new List<GenderPair>();
        for (var a = 0; a < Tags.Length; a++)
        {
            for (var b = a; b < Tags.Length; b++)
            {
                pairs.Add(new GenderPair(Tags[a], Tags[b], pairCounts[a, b]));
            }
        }

        return new GenderReport(totals, pairs, rows);
    }
}
