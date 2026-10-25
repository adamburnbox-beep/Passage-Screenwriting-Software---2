using System.Text.Json;
using System.Text.Json.Serialization;

namespace Passage.Web.Services;

/// <summary>
/// The per-script Slate sidecar. Worksheet working-out lives here, never in
/// the fountain file, so exports, page counts and the desktop apps' parse are
/// untouched. The sidecar is a staging area: once a runner promotes a result
/// into the script, the script is the truth (docs/SLATE-PLAN.md, decision B).
/// </summary>
public sealed class SlateDocument
{
    public int SlateVersion { get; set; } = SlateStore.CurrentVersion;

    // Bracket spans dismissed as prose, matched by text rather than line so a
    // reference survives the script being edited underneath it.
    public List<string> IgnoredBrackets { get; set; } = new();

    // Every worksheet run (docs/SLATE-REDO-PLAN.md R-E). Version 1 sidecars kept
    // one list per tool; SlateMigration converts them on load.
    public List<Passage.Web.Services.Tools.ExerciseRun> Exercises { get; set; } = new();

    public BurstSettings Burst { get; set; } = new();

    // Character panel settings (HIGHLAND-PLAN H3.1), keyed by bare cue name.
    // Only characters given something are kept.
    public List<CharacterSettings> Characters { get; set; } = new();

    // The Beat Sheet view. Planning only: beats never touch the script, and
    // list order is the tray order and the tie-break when beats stack.
    public List<Beat> Beats { get; set; } = new();
    public int BeatSheetPages { get; set; } = BeatSheet.DefaultTargetPages;

    // Permanent story-planning data (docs/STORY-PLAN.md, decision A).
    public StoryData Story { get; set; } = new();
}

public sealed class StoryData
{
    public List<Storyline> Storylines { get; set; } = new();
    public List<LocationSettings> Locations { get; set; } = new(); // S1.3, keyed by normalised name
    public List<ItemTags> Items { get; set; } = new();
    public List<ArcPoint> Arc { get; set; } = new(); // S3.1: only dragged points are stored
    public List<NotePage> Notes { get; set; } = new(); // S4.2: reference pages, never parsed or exported
}

/// <summary>A project note (S4.2): a titled page of Markdown the writer keeps
/// beside the script. Template is the id it started from, informational only.</summary>
public sealed class NotePage
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
}

/// <summary>A dragged value on one storyline's arc (S3.1), keyed by item rather
/// than position so Board order and Pages show the same values. Value is 0–100;
/// a tagged item with no point sits at 50, drawn hollow. Label is the last
/// heading seen, so a detached point stays recognisable.</summary>
public sealed class ArcPoint
{
    public Guid StorylineId { get; set; }
    public Guid TargetId { get; set; }
    public string Kind { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public int Value { get; set; } = 50;
}

/// <summary>Story data on one Board card or beat. Kind is "card" or "beat";
/// TargetId is the card's [[id:…]] or the Beat.Id. Label is the last heading
/// seen, so a detached item stays recognisable.</summary>
public sealed class ItemTags
{
    public Guid TargetId { get; set; }
    public string Kind { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public bool Key { get; set; }
    public List<Guid> Storylines { get; set; } = new();
    /// <summary>Bare cue names, as <see cref="CharacterSettings"/> (S1.4).</summary>
    public List<string> Characters { get; set; } = new();
    /// <summary>Normalised location names, decision C (S1.4).</summary>
    public List<string> Locations { get; set; } = new();

    public bool IsEmpty => !Key && Storylines.Count == 0 && Characters.Count == 0 && Locations.Count == 0;
}

/// <summary>A location's optional colour (S1.3), keyed by normalised name
/// (decision C). Only locations given a colour are kept.</summary>
public sealed class LocationSettings
{
    public string Name { get; set; } = string.Empty;
    public string Colour { get; set; } = string.Empty;
}

/// <summary>A named thread through the story (S1.2). Colour is a
/// <see cref="Characters.Palette"/> hex, or empty. List order is its order.</summary>
public sealed class Storyline
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public string Colour { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    /// <summary>What the arc's vertical axis measures for this storyline (S3.1).</summary>
    public string ArcAxis { get; set; } = "Tension";
}

/// <summary>One beat on the Beat Sheet. Page 0 means unassigned (the tray);
/// otherwise the beat covers Page through Page + Span - 1. Colour is a
/// <see cref="Characters.Palette"/> hex, or empty.</summary>
public sealed class Beat
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = string.Empty;
    public string Colour { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int Page { get; set; }
    public int Span { get; set; } = 1;
}

/// <summary>Per-character settings for one script. Colour is a palette hex,
/// or empty for no highlight. Gender is a <see cref="Characters.Tags"/> entry
/// (H3.2), or empty for Unspecified.</summary>
public sealed class CharacterSettings
{
    public string Name { get; set; } = string.Empty;
    public string Colour { get; set; } = string.Empty;
    public string Gender { get; set; } = string.Empty;
}

/// <summary>
/// The per-answer burst timer (SLATE-PLAN decision E). Off unless the writer
/// turns it on; expiry only ever advances to the next answer.
/// </summary>
public sealed class BurstSettings
{
    public bool Enabled { get; set; }
    public int Seconds { get; set; } = 30;

    // The Lens is applied timed, 3–5 minutes in the worksheet.
    public int LensMinutes { get; set; } = 4;
}

/// <summary>
/// Ideation burst ("Generate — Ideation or Bridge", Path A). Deliberately
/// about no script at all, so it lives in its own file. The rounds are
/// disposable; the one kept line per sitting is the output.
/// </summary>
public sealed class IdeationDocument
{
    public int SlateVersion { get; set; } = SlateStore.CurrentVersion;
    public int RoundSeconds { get; set; } = 90;

    // The sitting in progress, so closing the overlay or reloading keeps it.
    public IdeationSitting? Current { get; set; }

    // One line per finished sitting. Lines, not sessions: no dates, no lane.
    public List<string> Kept { get; set; } = new();
}

public sealed class IdeationSitting
{
    // 1–10, locked for the sitting.
    public int Lane { get; set; }
    public List<IdeationRound> Rounds { get; set; } = new() { new() };
    public string KeptLine { get; set; } = string.Empty;
}

public sealed class IdeationRound
{
    public string Input { get; set; } = string.Empty;
    public string Output { get; set; } = string.Empty;
}

/// <summary>
/// Reads and writes <c>&lt;root&gt;/.slate/&lt;script name&gt;.json</c>. Keys
/// go through <see cref="ScriptLibrary.TryValidateName"/> so browser input can
/// no more escape this directory than it can the library root. The
/// dot-directory has no script extension, so <see cref="ScriptLibrary.List"/>
/// never surfaces it.
/// </summary>
public sealed class SlateStore
{
    public const int CurrentVersion = 2;

    // Ideation's file. Not a script sidecar, so the orphan sweep leaves it.
    public const string IdeationFile = "ideation.json";

    // Saved structure templates (S2.2): library-wide, so also not a sidecar.
    public const string TemplatesFile = "templates.json";

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _root;
    private readonly object _gate = new();

    public SlateStore(ScriptLibrary library)
    {
        _root = Path.Combine(library.RootPath, ".slate");
    }

    /// <summary>
    /// Returns the sidecar for a script, or null when it has none. A corrupt
    /// file throws <see cref="JsonException"/> rather than being replaced with
    /// an empty document: the caller decides how to surface lost working-out.
    /// </summary>
    public SlateDocument? Load(string scriptName)
    {
        var path = ResolvePath(scriptName);
        lock (_gate)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            var json = File.ReadAllText(path);
            var document = JsonSerializer.Deserialize<SlateDocument>(json, JsonOptions);
            if (document is not null)
            {
                // A sidecar below the current version kept one record list per tool. They
                // become generic runs (docs/SLATE-REDO-PLAN.md R-E), in memory only; the file
                // is rewritten at the current version on the next save. The version is read
                // from the raw JSON: a sidecar that lacks the key would otherwise pass for current.
                Passage.Web.Services.Tools.SlateMigration.Migrate(document, json);
            }

            return document;
        }
    }

    public void Save(string scriptName, SlateDocument document)
    {
        var path = ResolvePath(scriptName);
        document.SlateVersion = CurrentVersion;
        lock (_gate)
        {
            Directory.CreateDirectory(_root);
            File.WriteAllText(path, JsonSerializer.Serialize(document, JsonOptions));
        }
    }

    public bool Exists(string scriptName)
    {
        if (!ScriptLibrary.TryValidateName(scriptName, out var validated))
        {
            return false;
        }

        lock (_gate)
        {
            return File.Exists(PathFor(validated));
        }
    }

    /// <summary>Cascade for the library's delete. No-op when there is no sidecar.</summary>
    public void Delete(string scriptName)
    {
        var path = ResolvePath(scriptName);
        lock (_gate)
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Deletes every sidecar whose script is not in <paramref name="scriptNames"/>
    /// and returns the script names they belonged to, so the caller can say
    /// so rather than pruning silently. Scripts can be removed from the volume
    /// outside the app, so this runs on load, like the recent-files list.
    /// </summary>
    public IReadOnlyList<string> PruneOrphans(IEnumerable<string> scriptNames) =>
        PruneOrphansCounted(scriptNames).Select(entry => entry.Script).ToList();

    /// <summary>As <see cref="PruneOrphans"/>, also reporting how many project
    /// notes each deleted sidecar held (0 for one that cannot be read).</summary>
    public IReadOnlyList<(string Script, int Notes)> PruneOrphansCounted(IEnumerable<string> scriptNames)
    {
        var present = scriptNames.ToHashSet(StringComparer.Ordinal);
        var pruned = new List<(string, int)>();
        lock (_gate)
        {
            if (!Directory.Exists(_root))
            {
                return pruned;
            }

            foreach (var path in Directory.EnumerateFiles(_root, "*.json"))
            {
                if (Path.GetFileName(path) is IdeationFile or TemplatesFile)
                {
                    continue;
                }

                var scriptName = Path.GetFileNameWithoutExtension(path);
                if (present.Contains(scriptName))
                {
                    continue;
                }

                var notes = 0;
                try
                {
                    notes = JsonSerializer.Deserialize<SlateDocument>(File.ReadAllText(path), JsonOptions)?.Story.Notes.Count ?? 0;
                }
                catch (JsonException)
                {
                }

                File.Delete(path);
                pruned.Add((scriptName, notes));
            }
        }

        return pruned;
    }

    /// <summary>The ideation file, or null when there is none yet. A corrupt
    /// file throws, as <see cref="Load"/> does.</summary>
    public IdeationDocument? LoadIdeation()
    {
        var path = Path.Combine(_root, IdeationFile);
        lock (_gate)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            return JsonSerializer.Deserialize<IdeationDocument>(File.ReadAllText(path), JsonOptions);
        }
    }

    public void SaveIdeation(IdeationDocument document)
    {
        document.SlateVersion = CurrentVersion;
        lock (_gate)
        {
            Directory.CreateDirectory(_root);
            File.WriteAllText(Path.Combine(_root, IdeationFile), JsonSerializer.Serialize(document, JsonOptions));
        }
    }

    /// <summary>The saved custom templates, or null when there are none yet. A
    /// corrupt file throws, as <see cref="Load"/> does.</summary>
    public TemplatesDocument? LoadTemplates()
    {
        var path = Path.Combine(_root, TemplatesFile);
        lock (_gate)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            return JsonSerializer.Deserialize<TemplatesDocument>(File.ReadAllText(path), JsonOptions);
        }
    }

    public void SaveTemplates(TemplatesDocument document)
    {
        document.SlateVersion = CurrentVersion;
        lock (_gate)
        {
            Directory.CreateDirectory(_root);
            File.WriteAllText(Path.Combine(_root, TemplatesFile), JsonSerializer.Serialize(document, JsonOptions));
        }
    }

    private string ResolvePath(string scriptName)
    {
        if (!ScriptLibrary.TryValidateName(scriptName, out var validated))
        {
            throw new ArgumentException($"Invalid script name '{scriptName}'.", nameof(scriptName));
        }

        return PathFor(validated);
    }

    private string PathFor(string validatedName) => Path.Combine(_root, validatedName + ".json");
}
