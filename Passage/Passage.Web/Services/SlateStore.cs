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

    // The Split family ("Generate — Split, Belief, or Extend Backward"). One
    // of each per script: the worksheet runs them once per story, and the
    // belief cuts are shared between Split's optional layer and Belief Split.
    public SplitRun Split { get; set; } = new();
    public BeliefRun Belief { get; set; } = new();
    public ExtendRun Extend { get; set; } = new();

    // Bridge and Position runs, kept by what each is about — the two ends,
    // the moment — never by when it was run.
    public List<BridgeRun> Bridges { get; set; } = new();
    public List<PositionRun> Positions { get; set; } = new();

    // Push/Pull and Lens runs, kept by the scene or stretch they check.
    public List<ReviseRun> Revisions { get; set; } = new();
}

/// <summary>Push/Pull and Lens ("Revise — Push-Pull and Lens"): six checks
/// on one scene or stretch, a Lens where a check is flagged, and the
/// optional diagnostic belief layer and read-back beneath.</summary>
public sealed class ReviseRun
{
    public static readonly string[] CheckNames = { "Polarity", "Linkage", "Third rail", "Migration", "Escalation", "Two-test" };

    public string Scene { get; set; } = string.Empty;
    public string Purpose { get; set; } = string.Empty;
    public List<ReviseCheck> Checks { get; set; } = CheckNames.Select(name => new ReviseCheck { Name = name }).ToList();

    // Going deeper still — Belief Split, diagnostic.
    public string DiagA { get; set; } = string.Empty;
    public string DiagZ { get; set; } = string.Empty;
    public string DiagShape { get; set; } = string.Empty;
    public string DiagMatch { get; set; } = string.Empty;
    public string DiagSeam { get; set; } = string.Empty;
    public string DiagLens { get; set; } = string.Empty;
    public string DiagFragment { get; set; } = string.Empty;

    // Read-back and crit.
    public string Working { get; set; } = string.Empty;
    public string NotSitting { get; set; } = string.Empty;
    public string Ready { get; set; } = string.Empty;

    public ReviseCheck Check(string name)
    {
        var check = Checks.FirstOrDefault(existing => existing.Name == name);
        if (check is null)
        {
            check = new ReviseCheck { Name = name };
            Checks.Add(check);
        }

        return check;
    }

    public bool HasDiagnostic =>
        !string.IsNullOrWhiteSpace(DiagA) || !string.IsNullOrWhiteSpace(DiagZ) || !string.IsNullOrWhiteSpace(DiagShape)
        || !string.IsNullOrWhiteSpace(DiagMatch) || !string.IsNullOrWhiteSpace(DiagSeam)
        || !string.IsNullOrWhiteSpace(DiagLens) || !string.IsNullOrWhiteSpace(DiagFragment);

    public bool HasReadBack =>
        !string.IsNullOrWhiteSpace(Working) || !string.IsNullOrWhiteSpace(NotSitting) || !string.IsNullOrWhiteSpace(Ready);

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Scene) && string.IsNullOrWhiteSpace(Purpose)
        && Checks.All(check => check.IsEmpty) && !HasDiagnostic && !HasReadBack;
}

/// <summary>One check: its answer and Y/N flag, and — only where flagged —
/// the Lens applied to that flagged thing.</summary>
public sealed class ReviseCheck
{
    public string Name { get; set; } = string.Empty;
    public string Answer { get; set; } = string.Empty;
    public string Flag { get; set; } = string.Empty;
    public string Lens { get; set; } = string.Empty;
    public string Fragment { get; set; } = string.Empty;
    public string Moved { get; set; } = string.Empty;
    public string MovedWhy { get; set; } = string.Empty;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Answer) && string.IsNullOrWhiteSpace(Flag) && string.IsNullOrWhiteSpace(Lens)
        && string.IsNullOrWhiteSpace(Fragment) && string.IsNullOrWhiteSpace(Moved) && string.IsNullOrWhiteSpace(MovedWhy);
}

/// <summary>Bridge ("Generate — Ideation or Bridge", Path B): two fixed
/// points, rounds of three candidate links between them.</summary>
public sealed class BridgeRun
{
    public string A { get; set; } = string.Empty;
    public string Z { get; set; } = string.Empty;
    public List<BridgeRound> Rounds { get; set; } = new() { new() };
    public string Wire { get; set; } = string.Empty;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(A) && string.IsNullOrWhiteSpace(Z) && string.IsNullOrWhiteSpace(Wire)
        && Rounds.All(round => round.IsEmpty);
}

public sealed class BridgeRound
{
    public List<string> Candidates { get; set; } = new() { string.Empty, string.Empty, string.Empty };
    public string Picked { get; set; } = string.Empty;

    public bool IsEmpty => string.IsNullOrWhiteSpace(Picked) && Candidates.All(string.IsNullOrWhiteSpace);
}

/// <summary>Position ("Place or Build — Position or Fill", Path A): one
/// moment tried in the story's seven slots, one turn at a time.</summary>
public sealed class PositionRun
{
    public string Moment { get; set; } = string.Empty;
    public List<PositionTurn> Turns { get; set; } = new() { new() };
    public List<string> Shortlist { get; set; } = new() { string.Empty, string.Empty };

    public const int SlotCount = 7;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Moment) && Shortlist.All(string.IsNullOrWhiteSpace) && Turns.All(turn => turn.IsEmpty);
}

public sealed class PositionTurn
{
    // One of the seven turn labels (SplitScript.Turns), or empty.
    public string Slot { get; set; } = string.Empty;
    public string Before { get; set; } = string.Empty;
    public string After { get; set; } = string.Empty;
    public string Alive { get; set; } = string.Empty;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Slot) && string.IsNullOrWhiteSpace(Before)
        && string.IsNullOrWhiteSpace(After) && string.IsNullOrWhiteSpace(Alive);
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

/// <summary>Path A — A→Z→Split. Rung 0 fixes the ends; each rung after is one split.</summary>
public sealed class SplitRun
{
    public string A { get; set; } = string.Empty;
    public string Z { get; set; } = string.Empty;
    public List<string> Candidates { get; set; } = new() { string.Empty, string.Empty, string.Empty };
    public string Midpoint { get; set; } = string.Empty;
    public string MidpointAlive { get; set; } = string.Empty;
    public string PP1 { get; set; } = string.Empty;
    public string Crisis { get; set; } = string.Empty;
    public string Inciting { get; set; } = string.Empty;
    public string Pinch1 { get; set; } = string.Empty;
    public string Pinch2 { get; set; } = string.Empty;
    public string LastObstacle { get; set; } = string.Empty;

    // The optional matched belief cut, asked inline after each plot turn.
    public bool BeliefLayer { get; set; }
    public string Pinch1ThirdRail { get; set; } = string.Empty;
    public string Pinch2ThirdRail { get; set; } = string.Empty;

    public bool HasRung2 => !string.IsNullOrWhiteSpace(PP1) || !string.IsNullOrWhiteSpace(Crisis);
    public bool HasRung3 => !string.IsNullOrWhiteSpace(Inciting) || !string.IsNullOrWhiteSpace(Pinch1)
        || !string.IsNullOrWhiteSpace(Pinch2) || !string.IsNullOrWhiteSpace(LastObstacle);
}

/// <summary>Path B — Belief Split. The five named cuts, by ratio.</summary>
public sealed class BeliefRun
{
    public string Shape { get; set; } = string.Empty;
    public string A { get; set; } = string.Empty;
    public string Z { get; set; } = string.Empty;
    public string Ghost { get; set; } = string.Empty;
    public string GhostSize { get; set; } = string.Empty;
    public string GhostCommensurate { get; set; } = string.Empty;
    public string GhostMismatch { get; set; } = string.Empty;
    public string Cut50Change { get; set; } = string.Empty;
    public List<BeliefCut> Cuts { get; set; } = BeliefCut.Ratios.Select(ratio => new BeliefCut { Ratio = ratio }).ToList();

    public BeliefCut Cut(int ratio)
    {
        var cut = Cuts.FirstOrDefault(existing => existing.Ratio == ratio);
        if (cut is null)
        {
            cut = new BeliefCut { Ratio = ratio };
            Cuts.Add(cut);
        }

        return cut;
    }

    public bool HasRung3 => !string.IsNullOrWhiteSpace(Cut(25).Text) || !string.IsNullOrWhiteSpace(Cut(75).Text);
    public bool HasRung4 => !string.IsNullOrWhiteSpace(Cut(12).Text) || !string.IsNullOrWhiteSpace(Cut(88).Text);
}

public sealed class BeliefCut
{
    public static readonly int[] Ratios = { 12, 25, 50, 75, 88 };

    public int Ratio { get; set; }
    public string Text { get; set; } = string.Empty;
    public string Alive { get; set; } = string.Empty;
}

/// <summary>Path C — Extend Backward. Links run from Z backwards, one at a time.</summary>
public sealed class ExtendRun
{
    public string Z { get; set; } = string.Empty;
    public List<ExtendLink> Links { get; set; } = new() { new() };
    public string ChainAlive { get; set; } = string.Empty;
    public string WeakestLink { get; set; } = string.Empty;
}

public sealed class ExtendLink
{
    public string Text { get; set; } = string.Empty;

    // The mystery-lens variant: Z as a crime scene, what would a detective
    // need to find — a different question for this link only.
    public bool Mystery { get; set; }
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
