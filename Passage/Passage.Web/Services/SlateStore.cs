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

    // Fill working-out, keyed by bracket text. A run is dropped on accept: the
    // filled line in the script is then the only truth.
    public List<FillRun> Fills { get; set; } = new();

    // Forward-chain working-out (WOAC and Character Flaw Brainstorm), kept by
    // what each chain is about, never by when it was run.
    public List<ChainRun> Chains { get; set; } = new();

    public BurstSettings Burst { get; set; } = new();
}

/// <summary>
/// The per-answer burst timer (SLATE-PLAN decision E). Off unless the writer
/// turns it on; expiry only ever advances to the next answer.
/// </summary>
public sealed class BurstSettings
{
    public bool Enabled { get; set; }
    public int Seconds { get; set; } = 30;
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ChainPath { Woac, Flaw }

/// <summary>
/// One forward chain (the "Place or Build — Forward Chain" worksheet). WOAC
/// stores its rounds; the Flaw brainstorm stores its eight answers. Every
/// field is optional and a chain can be left at any point.
/// </summary>
public sealed class ChainRun
{
    public ChainPath Path { get; set; }

    // "Character + starting want" or "Character + the flaw".
    public string Seed { get; set; } = string.Empty;

    // WOAC rounds. Only the first round's Want is stored: every later Want
    // is the previous round's Consequence, read live, so it cannot drift.
    public List<ChainRound> Rounds { get; set; } = new() { new() };

    // Character Flaw Brainstorm, one answer per fixed question.
    public List<string> Answers { get; set; } = Enumerable.Repeat(string.Empty, ChainRun.FlawQuestionCount).ToList();

    // "Read it back — what's the scene now, in one line?"
    public string ReadBack { get; set; } = string.Empty;

    public const int FlawQuestionCount = 8;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Seed) && string.IsNullOrWhiteSpace(ReadBack)
        && Rounds.All(round => round.IsEmpty)
        && Answers.All(string.IsNullOrWhiteSpace);
}

public sealed class ChainRound
{
    public string Want { get; set; } = string.Empty;
    public string Obstacle { get; set; } = string.Empty;
    public string Action { get; set; } = string.Empty;
    public string Consequence { get; set; } = string.Empty;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Want) && string.IsNullOrWhiteSpace(Obstacle)
        && string.IsNullOrWhiteSpace(Action) && string.IsNullOrWhiteSpace(Consequence);
}

/// <summary>The Fill worksheet (Path B) for one bracket. Every field is optional.</summary>
public sealed class FillRun
{
    public string Bracket { get; set; } = string.Empty;
    public List<FillOption> Options { get; set; } = new() { new(), new(), new() };
    public string PointsTo { get; set; } = string.Empty;
    public string Answer { get; set; } = string.Empty;
}

public sealed class FillOption
{
    public string Text { get; set; } = string.Empty;
    public string WhyWrong { get; set; } = string.Empty;
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
    public const int CurrentVersion = 1;

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

            return JsonSerializer.Deserialize<SlateDocument>(File.ReadAllText(path), JsonOptions);
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
    public IReadOnlyList<string> PruneOrphans(IEnumerable<string> scriptNames)
    {
        var present = scriptNames.ToHashSet(StringComparer.Ordinal);
        var pruned = new List<string>();
        lock (_gate)
        {
            if (!Directory.Exists(_root))
            {
                return pruned;
            }

            foreach (var path in Directory.EnumerateFiles(_root, "*.json"))
            {
                var scriptName = Path.GetFileNameWithoutExtension(path);
                if (present.Contains(scriptName))
                {
                    continue;
                }

                File.Delete(path);
                pruned.Add(scriptName);
            }
        }

        return pruned;
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
