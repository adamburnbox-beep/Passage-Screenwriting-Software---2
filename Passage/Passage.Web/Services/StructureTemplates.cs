using System.Text.RegularExpressions;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>One heading in a template: an Act (<c>#</c>), Sequence (<c>##</c>)
/// or Section (<c>###</c>), the <c>=</c> prompt under it, and the page it
/// conventionally sits at on a 110-page reference (0 for none).</summary>
public sealed class TemplateNode
{
    public string Kind { get; set; } = "Section";
    public string Heading { get; set; } = string.Empty;
    public string Prompt { get; set; } = string.Empty;
    public int Page { get; set; }
}

public sealed class StructureTemplate
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public bool BuiltIn { get; set; }
    public List<TemplateNode> Nodes { get; set; } = new();
}

/// <summary>The library-wide file, <c>.slate/templates.json</c>: custom
/// templates only. The built-ins live in code.</summary>
public sealed class TemplatesDocument
{
    public int SlateVersion { get; set; } = SlateStore.CurrentVersion;
    public List<StructureTemplate> Templates { get; set; } = new();

    // S4.2: custom project-note templates, a sibling list so structure
    // templates are untouched. An older file simply has none.
    public List<NoteTemplate> NoteTemplates { get; set; } = new();
}

/// <summary>
/// Structure templates (docs/STORY-PLAN.md, decision F). A template writes only
/// <c>#</c>, <c>##</c>, <c>###</c> and <c>=</c> lines, so applying one adds no
/// pages. Pure, so the lines, the page scaling and the seeding are testable
/// without the editor.
/// </summary>
public static partial class StructureTemplates
{
    /// <summary>The page count the templates' conventional pages are given on.</summary>
    public const int ReferencePages = 110;

    private static readonly Regex IdTag = IdTagPattern();

    [GeneratedRegex(@"\s*\[\[id:[0-9a-fA-F-]+\]\]")]
    private static partial Regex IdTagPattern();

    // The 110-page position of each of Passage's eight sequences. Acts do not
    // seed: the sequences already say where they fall.
    private static readonly int[] SequencePages = { 1, 14, 28, 41, 55, 69, 83, 96 };

    public static IReadOnlyList<StructureTemplate> BuiltIns { get; } = new[]
    {
        Passage(),
        Paradigm(),
        SaveTheCat(),
        FiveActTv()
    };

    /// <summary>The lines to insert: a blank line before each heading, then the
    /// heading with a fresh id and a <c>=</c> line per prompt line.</summary>
    public static List<string> BuildLines(StructureTemplate template)
    {
        var lines = new List<string>();
        foreach (var node in template.Nodes)
        {
            lines.Add(string.Empty);
            lines.AddRange(BeatBoardText.BuildCardLines(node.Kind, node.Heading, node.Prompt, Guid.NewGuid()));
        }

        return lines;
    }

    /// <summary>True when the script already has a <c>#</c> or <c>##</c> heading,
    /// which is when a template is not offered (<c>###</c> alone does not count).</summary>
    public static bool HasStructure(IEnumerable<string> scriptLines) =>
        scriptLines.Any(line => HeadingDepth(line) is 1 or 2);

    /// <summary>
    /// Reads <c>#</c>/<c>##</c>/<c>###</c> headings and the <c>=</c> lines that
    /// belong to them. A <c>=</c> line after a scene or other printing line
    /// belongs to that scene, so it is left out. Ids are dropped: a template
    /// gets fresh ones each time it is written.
    /// </summary>
    public static List<TemplateNode> ReadNodes(IEnumerable<string> scriptLines)
    {
        var nodes = new List<TemplateNode>();
        var prompts = new List<string>();
        TemplateNode? current = null;

        void Finish()
        {
            if (current is not null)
            {
                current.Prompt = string.Join('\n', prompts);
            }

            prompts.Clear();
        }

        foreach (var line in scriptLines)
        {
            var depth = HeadingDepth(line);
            if (depth > 0)
            {
                Finish();
                var heading = IdTag.Replace(line.TrimStart('#'), string.Empty).Trim();
                current = new TemplateNode { Kind = depth == 1 ? "Act" : depth == 2 ? "Sequence" : "Section", Heading = heading };
                nodes.Add(current);
            }
            else if (line.StartsWith('=') && !line.StartsWith("==", StringComparison.Ordinal))
            {
                if (current is not null)
                {
                    var text = IdTag.Replace(line[1..], string.Empty).Trim();
                    if (text.Length > 0)
                    {
                        prompts.Add(text);
                    }
                }
            }
            else if (line.Trim().Length > 0)
            {
                // Something that prints, or a note: whatever follows is its.
                Finish();
                current = null;
            }
        }

        Finish();
        return nodes;
    }

    /// <summary>A page on the 110-page reference, scaled to the Beat Sheet's
    /// target and kept on the grid.</summary>
    public static int ScalePage(int referencePage, int targetPages)
    {
        targetPages = Math.Max(1, targetPages);
        var scaled = (int)Math.Round(referencePage * (double)targetPages / ReferencePages, MidpointRounding.AwayFromZero);
        return Math.Clamp(scaled, 1, targetPages);
    }

    /// <summary>
    /// Adds a Beat Sheet beat for each node with a page. A node whose title is
    /// already a beat is skipped and counted, so applying twice never doubles.
    /// </summary>
    public static (int Added, int Skipped) SeedBeats(StructureTemplate template, IList<Beat> beats, int targetPages, int columns)
    {
        var added = 0;
        var skipped = 0;
        foreach (var node in template.Nodes.Where(node => node.Page > 0))
        {
            if (beats.Any(beat => string.Equals(beat.Title.Trim(), node.Heading, StringComparison.OrdinalIgnoreCase)))
            {
                skipped++;
                continue;
            }

            var (page, span) = BeatSheet.Clamp(ScalePage(node.Page, targetPages), 1, columns);
            beats.Add(new Beat { Title = node.Heading, Page = page, Span = span });
            added++;
        }

        return (added, skipped);
    }

    /// <summary>The template that the script's current structure makes.</summary>
    public static StructureTemplate FromScript(string name, IEnumerable<string> scriptLines) =>
        new() { Name = name.Trim(), Nodes = ReadNodes(scriptLines) };

    private static StructureTemplate Passage()
    {
        // The Shape with empty values (decision F): one copy of the eight sequences.
        var nodes = ReadNodes(SplitScript.BuildLines(new SplitRun()));
        var sequence = 0;
        foreach (var node in nodes.Where(node => node.Kind == "Sequence"))
        {
            node.Page = SequencePages[sequence++];
        }

        return new StructureTemplate
        {
            Id = new Guid("8a3f6b1e-0000-4000-8000-000000000001"),
            Name = "8 sequences / 7 turns (Passage)",
            Description = "Passage's own: four acts, eight sequences and the seven turns between them, as the Shape tool "
                + "writes them. Apply it, then fill the turns in the Shape tool. Built on the sequence approach "
                + "(Paul Gulino, Screenwriting: The Sequence Approach).",
            BuiltIn = true,
            Nodes = nodes
        };
    }

    private static TemplateNode Act(string heading, string prompt, int page = 0) =>
        new() { Kind = "Act", Heading = heading, Prompt = prompt, Page = page };

    private static TemplateNode Sequence(string heading, string prompt, int page = 0) =>
        new() { Kind = "Sequence", Heading = heading, Prompt = prompt, Page = page };

    private static TemplateNode Section(string heading, string prompt, int page = 0) =>
        new() { Kind = "Section", Heading = heading, Prompt = prompt, Page = page };

    private static StructureTemplate Paradigm() => new()
    {
        Id = new Guid("8a3f6b1e-0000-4000-8000-000000000002"),
        Name = "Paradigm (Syd Field)",
        Description = "Three acts with two plot points, a midpoint and a pinch in each half of the middle, after Syd Field's "
            + "paradigm (Screenplay, 1979). Page positions are conventional, for a 110-page script.",
        BuiltIn = true,
        Nodes =
        {
            Act("Act 1 - Setup", "Who the story is about, the world they are in, and what the story is about.", 1),
            Section("Plot point 1", "The event that turns the story and sends it into the middle.", 25),
            Act("Act 2 - Confrontation", "Obstacles in the way of what the character needs.", 30),
            Section("Pinch 1", "A reminder of the opposing force, a quarter of the way through the script.", 45),
            Section("Midpoint", "The turn in the middle of the script that changes what the story is about.", 55),
            Section("Pinch 2", "A second reminder of the opposing force.", 75),
            Section("Plot point 2", "The event that turns the story and sends it into the ending.", 85),
            Act("Act 3 - Resolution", "How the story resolves.", 90)
        }
    };

    private static StructureTemplate SaveTheCat() => new()
    {
        Id = new Guid("8a3f6b1e-0000-4000-8000-000000000003"),
        Name = "Save the Cat (Blake Snyder)",
        Description = "The fifteen beats of Blake Snyder's beat sheet (Save the Cat!, 2005), with Passage's own prompts. "
            + "Page positions are Snyder's, for a 110-page script.",
        BuiltIn = true,
        Nodes =
        {
            Act("Act 1", "The world as it is, and the first push out of it."),
            Section("Opening image", "A picture of the story's starting point.", 1),
            Section("Theme stated", "Someone says what the story is really about, and the hero doesn't hear it.", 5),
            Section("Set-up", "The hero's world, what is missing, and who is in it.", 1),
            Section("Catalyst", "The event that upsets the world.", 12),
            Section("Debate", "The hero resists. Should they go?", 13),
            Section("Break into two", "The hero chooses, and enters the new world.", 25),
            Act("Act 2", "The new world, and the cost of staying in it."),
            Section("B story", "The relationship that carries the theme.", 30),
            Section("Fun and games", "The promise of the premise.", 30),
            Section("Midpoint", "A false victory or false defeat; the stakes rise.", 55),
            Section("Bad guys close in", "Outside pressure and inside doubt both tighten.", 55),
            Section("All is lost", "The lowest point; something dies.", 75),
            Section("Dark night of the soul", "The hero faces what the loss means.", 75),
            Act("Act 3", "The hero uses what they learned."),
            Section("Break into three", "The answer arrives, often from the B story.", 85),
            Section("Finale", "The hero acts, changed, against the odds.", 85),
            Section("Final image", "The opposite of the opening image.", 110)
        }
    };

    private static StructureTemplate FiveActTv() => new()
    {
        Id = new Guid("8a3f6b1e-0000-4000-8000-000000000004"),
        Name = "Five-act TV",
        Description = "A teaser and five acts, the common shape of a broadcast one-hour drama. It is a convention of the form, "
            + "not one writer's method. Pages are spread over a 110-page reference.",
        BuiltIn = true,
        Nodes =
        {
            Act("Teaser", "A hook before the titles.", 1),
            Act("Act 1", "The problem of the week, and the first complication.", 5),
            Section("Act 1 out", "The turn that makes the audience stay.", 25),
            Act("Act 2", "The first attempt, and it goes wrong.", 27),
            Section("Act 2 out", "The turn that raises the stakes.", 47),
            Act("Act 3", "A new plan, and a cost.", 49),
            Section("Act 3 out", "The turn that takes something away.", 69),
            Act("Act 4", "Everything at risk.", 71),
            Section("Act 4 out", "The darkest turn.", 91),
            Act("Act 5", "The resolution, and the tag.", 93)
        }
    };

    private static int HeadingDepth(string line)
    {
        var depth = 0;
        while (depth < line.Length && line[depth] == '#')
        {
            depth++;
        }

        return depth is >= 1 and <= 3 && depth < line.Length && line[depth] == ' ' ? depth : 0;
    }
}
