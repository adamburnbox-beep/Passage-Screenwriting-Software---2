namespace Passage.Web.Services.Tools;

/// <summary>
/// The answers of an A→Z→Split or Belief Split run as the old records the Shape write and the
/// belief cuts read (docs/SLATE-REDO-PLAN.md T5). The keys are the step keys in the catalog;
/// nothing here is stored, so the run stays the only copy.
/// </summary>
public static class ShapeAnswers
{
    // Split's step keys, in the order SplitScript.Slots lists the lines they feed.
    public const string A = "a";
    public const string Z = "z";
    public const string Midpoint = "mid";
    public const string Stress = "stress";
    public const string PP1 = "pp1";
    public const string Crisis = "pp2";
    public const string Inciting = "inc";
    public const string Pinch1 = "p1";
    public const string Pinch2 = "p2";
    public const string LastObstacle = "last";
    public const string BeliefLayer = "layer";
    public const string Pinch1ThirdRail = "rail1";
    public const string Pinch2ThirdRail = "rail2";
    public const string Candidates = "cands";

    // The Split step that holds each Shape line, by the line's label.
    public static readonly (string Label, string Key)[] SplitKeys =
    {
        ("Starts", A), ("Inciting incident", Inciting), ("Plot point 1", PP1), ("Pinch 1", Pinch1), ("Midpoint", Midpoint),
        ("Pinch 2", Pinch2), ("Crisis / Plot point 2", Crisis), ("Last obstacle", LastObstacle), ("Ends", Z)
    };

    // Belief steps: the cut at a ratio, what changed to reach the 50% cut, and whether each cut is alive.
    public static string CutKey(int ratio) => $"cut{ratio}";

    public static string AliveKey(int ratio) => $"al{ratio}";

    // Split's own optional belief layer asks the same cuts under these keys.
    public static string LayerCutKey(int ratio) => $"b{ratio}";

    private static string Get(ExerciseRun run, string key) => run.Answers.TryGetValue(key, out var value) ? value : string.Empty;

    /// <summary>
    /// A fresh run starts from the script: A and Z are the Shape's own Starts / Ends lines when it has them.
    /// Only a run with no answers is touched (clearing a field later must not bring the script's text back),
    /// and a run about a named scene has its own Z, so it is left to the writer. True when anything was filled.
    /// </summary>
    public static bool Prefill(Exercise exercise, ExerciseRun run, Func<string, string?> scriptValue)
    {
        if (run.Answers.Count > 0 || run.Anchor.SceneId is not null || run.Anchor.SceneText.Length > 0)
        {
            return false;
        }

        var filled = false;
        foreach (var step in exercise.Steps.Where(step => step.ScriptSlot is "Starts" or "Ends"))
        {
            if (scriptValue(step.ScriptSlot!) is { } text)
            {
                run.Answers[step.Key] = text;
                filled = true;
            }
        }

        return filled;
    }

    /// <summary>A Split run's answers as a <see cref="SplitRun"/>, the shape <see cref="SplitScript"/> writes from.</summary>
    public static SplitRun ToSplitRun(ExerciseRun run)
    {
        var split = new SplitRun
        {
            A = Get(run, A),
            Z = Get(run, Z),
            Midpoint = Get(run, Midpoint),
            MidpointAlive = Get(run, Stress),
            PP1 = Get(run, PP1),
            Crisis = Get(run, Crisis),
            Inciting = Get(run, Inciting),
            Pinch1 = Get(run, Pinch1),
            Pinch2 = Get(run, Pinch2),
            LastObstacle = Get(run, LastObstacle),
            BeliefLayer = Get(run, BeliefLayer) == "y",
            Pinch1ThirdRail = Get(run, Pinch1ThirdRail),
            Pinch2ThirdRail = Get(run, Pinch2ThirdRail),
            Candidates = new List<string>()
        };

        for (var i = 0; run.Answers.ContainsKey($"{Candidates}[{i}]") || i < 3; i++)
        {
            split.Candidates.Add(Get(run, $"{Candidates}[{i}]"));
        }

        return split;
    }

    /// <summary>A Belief run's answers as a <see cref="BeliefRun"/>: the shape, the ends, the origin and the five cuts.</summary>
    public static BeliefRun ToBeliefRun(ExerciseRun run)
    {
        var belief = new BeliefRun
        {
            Shape = Get(run, "shape"),
            A = Get(run, A),
            Z = Get(run, Z),
            Ghost = Get(run, "ghost"),
            GhostSize = Get(run, "size"),
            GhostCommensurate = Get(run, "fits"),
            GhostMismatch = Get(run, "resize"),
            Cut50Change = Get(run, "chg50")
        };

        foreach (var ratio in BeliefCut.Ratios)
        {
            var cut = belief.Cut(ratio);
            cut.Text = Get(run, CutKey(ratio));
            cut.Alive = Get(run, AliveKey(ratio));
        }

        return belief;
    }

    /// <summary>
    /// What a Split run holds for a Shape line, for the plot turn a belief cut is matched to.
    /// The newest kept run wins when there is more than one; null when it holds nothing real.
    /// </summary>
    public static string? SplitTurn(SlateDocument slate, string label)
    {
        var key = SplitKeys.FirstOrDefault(pair => pair.Label == label).Key;
        var run = slate.Exercises.LastOrDefault(candidate => candidate.ExerciseId == SlateMigration.SplitId);
        if (key is null || run is null)
        {
            return null;
        }

        var value = Get(run, key).Trim();
        return value.Length == 0 ? null : value;
    }
}
