using System.Text.RegularExpressions;
using Passage.Core;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>One speaker in the character panel. Lines are speeches — one per
/// cue — and words are spoken words, parentheticals excluded.</summary>
public sealed record CharacterStat(string Name, int Lines, int Words);

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
}
