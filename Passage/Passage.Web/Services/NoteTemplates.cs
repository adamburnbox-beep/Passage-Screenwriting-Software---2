using System.IO.Compression;
using System.Text;

namespace Passage.Web.Services;

/// <summary>A starting body for a new project note. Built-ins live in code;
/// custom ones are stored in <c>.slate/templates.json</c> (docs/STORY-PLAN.md S4.2).</summary>
public sealed class NoteTemplate
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public bool BuiltIn { get; set; }
}

/// <summary>The three built-in project-note templates, and the pure helpers
/// for downloads and the prune message, so they are testable without the editor.</summary>
public static class NoteTemplates
{
    public static IReadOnlyList<NoteTemplate> BuiltIns { get; } = new[]
    {
        Make("00000000-0000-0000-0000-00000000a001", "Character", """
            # Character

            ## Who they are
            Name, age, job, the one-line impression they give a stranger.

            ## Want and need
            What are they chasing? What do they actually need that they don't see?

            ## Flaw
            What gets in their own way?

            ## How they talk
            Rhythm, vocabulary, what they never say out loud.

            ## Relationships
            Who matters to them, and what does each person cost or give?

            ## Arc
            Who are they at the start, and who at the end?
            """),
        Make("00000000-0000-0000-0000-00000000a002", "Location", """
            # Location

            ## What it looks like and sounds like
            Shape, light, materials, the noise under the dialogue.

            ## Who owns it
            Who belongs here, and who is a guest?

            ## What it means to the story
            Why this place and not another?

            ## What can only happen here
            The scene that would not work anywhere else.
            """),
        Make("00000000-0000-0000-0000-00000000a003", "Logline & premise", """
            # Logline & premise

            ## The one sentence
            Someone, wanting something, against something, or else.

            ## The question the story asks
            What will the ending answer?

            ## The protagonist
            Who is at the centre, and why them?

            ## The pressure
            What forces the issue, and what is the clock?

            ## The ending it points at
            Where does it land, even if the route is not clear yet?
            """)
    };

    private static NoteTemplate Make(string id, string name, string body) => new()
    {
        Id = Guid.Parse(id),
        Name = name,
        Body = Dedent(body),
        BuiltIn = true
    };

    private static string Dedent(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n');
        var indent = lines.Where(line => line.Trim().Length > 0).Min(line => line.Length - line.TrimStart().Length);
        return string.Join('\n', lines.Select(line => line.Length >= indent ? line[indent..] : line.TrimStart())) + "\n";
    }

    /// <summary>A file name for a note: unsafe characters dropped, empty titles
    /// become "Untitled", and a name already taken (case-insensitively) gets
    /// " (2)", " (3)"…</summary>
    public static string FileName(string title, ISet<string> taken)
    {
        // Windows' set, not the host's: the zip is opened elsewhere.
        const string invalid = "\\/:*?\"<>|";
        var clean = new string(title.Where(c => !invalid.Contains(c) && !char.IsControl(c)).ToArray()).Trim().Trim('.').Trim();
        if (clean.Length > 80)
        {
            clean = clean[..80].TrimEnd();
        }

        if (clean.Length == 0)
        {
            clean = "Untitled";
        }

        var name = clean;
        for (var n = 2; !taken.Add(name); n++)
        {
            name = $"{clean} ({n})";
        }

        return name + ".md";
    }

    /// <summary>One <c>.md</c> per page, named from its title.</summary>
    public static byte[] BuildZip(IEnumerable<NotePage> notes)
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var note in notes)
            {
                var name = FileName(note.Title, taken);
                var entry = zip.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.Write(note.Body);
            }
        }

        return buffer.ToArray();
    }

    /// <summary>The status line after the orphan sweep (rule 12: never silent),
    /// naming how many project notes went with each deleted sidecar.</summary>
    public static string PruneMessage(IReadOnlyList<(string Script, int Notes)> pruned)
    {
        var message = $"Removed Slate data for missing script{(pruned.Count == 1 ? "" : "s")}: "
            + string.Join(", ", pruned.Select(entry => entry.Script));
        var withNotes = pruned.Where(entry => entry.Notes > 0).ToList();
        if (withNotes.Count > 0)
        {
            var total = withNotes.Sum(entry => entry.Notes);
            message += $" — including {total} project note{(total == 1 ? "" : "s")} ("
                + string.Join(", ", withNotes.Select(entry => $"{entry.Script}: {entry.Notes}")) + ")";
        }

        return message;
    }
}
