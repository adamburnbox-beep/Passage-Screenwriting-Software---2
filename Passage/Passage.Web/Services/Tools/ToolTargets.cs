using System.Text;

namespace Passage.Web.Services.Tools;

/// <summary>
/// What the "Send to" targets build (docs/SLATE-REDO-PLAN.md §4.4). Pure: the host decides whether a target can be
/// reached and does the writing; these decide the text and the sidecar edits, so they are testable without the editor.
/// </summary>
public static class ToolTargets
{
    private const int TitleLimit = 80;

    /// <summary>A "=" synopsis line from an answer: one line, lines joined with spaces.</summary>
    public static string SynopsisLine(string text) =>
        "= " + string.Join(' ', text.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(part => part.Trim()).Where(part => part.Length > 0));

    /// <summary>An inbox beat from an answer. The first line titles it (cut at a word where it is long); anything
    /// that did not fit in the title goes in the description, so nothing sent is lost.</summary>
    public static Beat BuildBeat(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(line => line.Trim()).Where(line => line.Length > 0).ToList();
        var first = lines.Count > 0 ? lines[0] : string.Empty;
        var title = first;
        if (title.Length > TitleLimit)
        {
            var cut = title.LastIndexOf(' ', TitleLimit - 1);
            title = (cut > TitleLimit / 2 ? title[..cut] : title[..(TitleLimit - 1)]).TrimEnd() + "…";
        }

        var description = string.Join('\n', lines);
        return new Beat { Title = title, Description = description == title ? string.Empty : description };
    }

    /// <summary>The sequence card a turn sits on: the one whose synopsis lines carry that turn's slot line
    /// (<c>= Pinch 1: …</c>, or its bracket). Null when no sequence in the script has it, so the caller says so
    /// rather than guessing a sequence by position.</summary>
    public static BoardCard? SequenceForTurn(IEnumerable<BoardCard> sequences, string turn) =>
        sequences.FirstOrDefault(card => card.Description.Split('\n')
            .Any(line => SplitScript.TryParse(line, out var label, out _) && label == turn.Trim()));

    /// <summary>Tags the beat with the character unless it already has it (ToggleTag alone would take it off).</summary>
    public static void TagBeat(StoryData story, Beat beat, string character)
    {
        var name = Characters.BareName(character);
        if (name.Length == 0)
        {
            return;
        }

        var item = story.Items.FirstOrDefault(entry => entry.TargetId == beat.Id);
        if (item is not null && item.Characters.Contains(name))
        {
            return;
        }

        StoryItems.ToggleTag(story, beat.Id, StoryItems.BeatKind, beat.Title, "character:" + name);
    }

    /// <summary>
    /// Adds "- label: answer" to the character's note page, under a heading named after the exercise. The page is
    /// made from the built-in Character template when the character has none. Returns the page and whether it was new.
    /// </summary>
    public static (NotePage Page, bool Created) AppendToNote(StoryData story, string character, string exercise, string label, string text)
    {
        var name = character.Trim();
        var page = ProjectContext.NoteFor(story, name);
        var created = page is null;
        if (page is null)
        {
            var template = NoteTemplates.BuiltIns[0];
            page = new NotePage { Title = name, Body = template.Body, Template = template.Name };
            story.Notes.Add(page);
        }

        page.Body = InsertUnderHeading(page.Body, "## " + exercise, Bullet(label, text));
        return (page, created);
    }

    private static string Bullet(string label, string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n').Select(line => line.TrimEnd()).Where(line => line.Trim().Length > 0).ToList();
        var builder = new StringBuilder("- ");
        if (label.Length > 0)
        {
            builder.Append(label).Append(": ");
        }

        for (var i = 0; i < lines.Count; i++)
        {
            builder.Append(i == 0 ? lines[i].TrimStart() : "\n  " + lines[i].TrimStart());
        }

        return builder.ToString();
    }

    // Puts the bullet at the end of the heading's section (before the next heading), or adds the section at the end.
    private static string InsertUnderHeading(string body, string heading, string bullet)
    {
        var lines = body.Replace("\r\n", "\n").Split('\n').ToList();
        var at = lines.FindIndex(line => string.Equals(line.TrimEnd(), heading, StringComparison.OrdinalIgnoreCase));
        if (at < 0)
        {
            var trimmed = body.TrimEnd();
            return (trimmed.Length == 0 ? string.Empty : trimmed + "\n\n") + heading + "\n" + bullet + "\n";
        }

        var end = at + 1;
        while (end < lines.Count && !lines[end].StartsWith('#'))
        {
            end++;
        }

        var last = end - 1;
        while (last > at && lines[last].Trim().Length == 0)
        {
            last--;
        }

        lines.Insert(last + 1, bullet);
        if (last + 2 == lines.Count)
        {
            lines.Add(string.Empty);
        }

        return string.Join('\n', lines);
    }
}
