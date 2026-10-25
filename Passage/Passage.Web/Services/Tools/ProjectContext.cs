using Passage.Parser;

namespace Passage.Web.Services.Tools;

/// <summary>A character in the Cast, for the picker: the highlight colour (empty for none) and how many speeches.</summary>
public sealed record CharacterChoice(string Name, string Colour, int Lines);

/// <summary>A scene heading, for the picker. Id is set only when the heading carries an [[id:…]] marker; an unmarked
/// heading gets a fresh id on every parse, so it can only be found again by its text.</summary>
public sealed record SceneChoice(Guid? Id, string Heading, int LineNumber, string Synopsis);

/// <summary>A bracket in the script, for the picker: the text as written, its line, whether it reads like prose
/// rather than a placeholder (rank 1), and whether the writer has left it out.</summary>
public sealed record BracketChoice(string Text, int LineNumber, bool Prose, bool Ignored);

/// <summary>
/// The lists the anchor pickers draw from (docs/SLATE-REDO-PLAN.md §4.4), built on demand from the parse and the
/// sidecar and never stored. Pure, so the lookups are testable without the editor.
/// </summary>
public static class ProjectContext
{
    /// <summary>Cast order (most lines first, as the panel shows them), coloured from the sidecar.</summary>
    public static List<CharacterChoice> Cast(IEnumerable<CharacterStat> stats, IEnumerable<CharacterSettings> settings)
    {
        var colours = settings
            .Where(setting => setting.Name.Length > 0)
            .GroupBy(setting => setting.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Colour, StringComparer.OrdinalIgnoreCase);
        return stats
            .OrderByDescending(stat => stat.Lines)
            .ThenBy(stat => stat.Name, StringComparer.OrdinalIgnoreCase)
            .Select(stat => new CharacterChoice(stat.Name, colours.GetValueOrDefault(stat.Name) ?? string.Empty, stat.Lines))
            .ToList();
    }

    /// <summary>Every scene heading in script order, with the synopsis line under it if there is one.</summary>
    public static List<SceneChoice> Scenes(IReadOnlyList<ScreenplayElement> elements)
    {
        var scenes = new List<SceneChoice>();
        for (var i = 0; i < elements.Count; i++)
        {
            if (elements[i] is not SceneHeadingElement scene)
            {
                continue;
            }

            var synopsis = string.Empty;
            for (var j = i + 1; j < elements.Count && elements[j] is SynopsisElement or NoteElement; j++)
            {
                if (elements[j] is SynopsisElement found)
                {
                    synopsis = found.Text.Trim();
                    break;
                }
            }

            scenes.Add(new SceneChoice(
                StoryItems.HasIdMarker(scene.RawText) ? scene.Id : null,
                scene.Text.Trim(),
                scene.LineNumber,
                synopsis));
        }

        return scenes;
    }

    /// <summary>The scene the caret is in: the last heading at or above it.</summary>
    public static SceneChoice? SceneAt(IReadOnlyList<SceneChoice> scenes, int caretLine) =>
        scenes.LastOrDefault(scene => scene.LineNumber <= caretLine);

    /// <summary>Finds an anchored scene again: by its id when it has one, otherwise by heading text when exactly one
    /// scene has it. Null means it is gone (or ambiguous), and the caller says so rather than guessing.</summary>
    public static SceneChoice? Resolve(IReadOnlyList<SceneChoice> scenes, Anchor? anchor)
    {
        if (anchor is null || (anchor.SceneId is null && anchor.SceneText.Length == 0))
        {
            return null;
        }

        if (anchor.SceneId is { } id && scenes.FirstOrDefault(scene => scene.Id == id) is { } byId)
        {
            return byId;
        }

        var matches = scenes.Where(scene => string.Equals(scene.Heading, anchor.SceneText, StringComparison.OrdinalIgnoreCase)).ToList();
        return matches.Count == 1 ? matches[0] : null;
    }

    /// <summary>Every bracket in the script as the picker lists it (ranked order), marked if ignored.</summary>
    public static List<BracketChoice> Brackets(IEnumerable<Bracket> brackets, IReadOnlyCollection<string> ignored) =>
        brackets.Select(bracket => new BracketChoice(bracket.Text, bracket.LineNumber, bracket.Rank == 1,
            ignored.Contains(bracket.Text, StringComparer.Ordinal))).ToList();

    /// <summary>Who speaks in the scene, most speeches in the script first (the Cast order).</summary>
    public static List<string> SpeakersOf(
        SceneChoice scene,
        IReadOnlyDictionary<int, HashSet<string>> sceneSpeakers,
        IReadOnlyList<CharacterChoice> cast)
    {
        if (!sceneSpeakers.TryGetValue(scene.LineNumber, out var names))
        {
            return new List<string>();
        }

        return cast.Where(character => names.Contains(character.Name)).Select(character => character.Name).ToList();
    }

    /// <summary>A character's note page: titled with their name, case-insensitively.</summary>
    public static NotePage? NoteFor(StoryData story, string character) =>
        story.Notes.FirstOrDefault(note => string.Equals(note.Title.Trim(), character.Trim(), StringComparison.OrdinalIgnoreCase));
}
