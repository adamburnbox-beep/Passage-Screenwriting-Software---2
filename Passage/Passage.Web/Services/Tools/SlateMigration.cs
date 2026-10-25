namespace Passage.Web.Services.Tools;

/// <summary>
/// Moves old per-tool records into <see cref="SlateDocument.Exercises"/>
/// (docs/SLATE-REDO-PLAN.md R-E). A move, not a copy: each record type is
/// converted and its old list cleared by the row that deletes that type's dock
/// view, so a second run finds nothing to do. Not gated by version; the
/// version bump waits for T8. Handles Chains (T2), Fills (T4) and the Split family (T5); it runs from
/// <c>SlateStore.Load</c>.
/// </summary>
public static class SlateMigration
{
    public const string WoacId = "woac";
    public const string FlawId = "flaw";
    public const string FillId = "fill";
    public const string SceneDialsId = "scenedials";
    public const string SplitId = "split";
    public const string BeliefId = "belief";
    public const string ExtendId = "extend";

    /// <summary>Moves Chains, Fills and the Split family in place. Returns the number of runs moved.</summary>
    public static int Migrate(SlateDocument document)
    {
        var moved = document.Chains.Count;
        foreach (var chain in document.Chains)
        {
            document.Exercises.Add(FromChain(chain, Guid.NewGuid().ToString("N")));
        }

        document.Chains.Clear();

        // A Fill with nothing written was never worth keeping (the old view
        // dropped it too), but it is still cleared so a second run finds none.
        foreach (var fill in document.Fills.Where(fill => !IsBlank(fill)))
        {
            document.Exercises.Add(FromFill(fill, Guid.NewGuid().ToString("N")));
            moved++;
        }

        document.Fills.Clear();

        // The Split family was one run each per script. An untouched one is not worth a kept run.
        // Split's belief layer asked the same five cuts Belief Split holds, so when the layer was on
        // the cuts are copied into the Split run as well (they are two exercises now, not one record).
        if (!IsBlank(document.Split))
        {
            document.Exercises.Add(FromSplit(document.Split, document.Belief, Guid.NewGuid().ToString("N")));
            moved++;
        }

        if (!IsBlank(document.Belief))
        {
            document.Exercises.Add(FromBelief(document.Belief, Guid.NewGuid().ToString("N")));
            moved++;
        }

        if (!IsBlank(document.Extend))
        {
            document.Exercises.Add(FromExtend(document.Extend, Guid.NewGuid().ToString("N")));
            moved++;
        }

        document.Split = new SplitRun();
        document.Belief = new BeliefRun();
        document.Extend = new ExtendRun();
        return moved;
    }

    private static bool IsBlank(SplitRun split) =>
        SplitScript.Values(split).All(string.IsNullOrWhiteSpace) && split.Candidates.All(string.IsNullOrWhiteSpace)
        && string.IsNullOrWhiteSpace(split.MidpointAlive)
        && string.IsNullOrWhiteSpace(split.Pinch1ThirdRail) && string.IsNullOrWhiteSpace(split.Pinch2ThirdRail);

    private static bool IsBlank(BeliefRun belief) =>
        new[] { belief.Shape, belief.A, belief.Z, belief.Ghost, belief.GhostSize, belief.GhostCommensurate, belief.GhostMismatch, belief.Cut50Change }
            .All(string.IsNullOrWhiteSpace)
        && belief.Cuts.All(cut => string.IsNullOrWhiteSpace(cut.Text) && string.IsNullOrWhiteSpace(cut.Alive));

    private static bool IsBlank(ExtendRun extend) =>
        string.IsNullOrWhiteSpace(extend.Z) && extend.Links.All(link => string.IsNullOrWhiteSpace(link.Text))
        && string.IsNullOrWhiteSpace(extend.ChainAlive) && string.IsNullOrWhiteSpace(extend.WeakestLink);

    /// <summary>A Split run keeps its answers under the keys <see cref="ShapeAnswers"/> reads back.</summary>
    public static ExerciseRun FromSplit(SplitRun split, BeliefRun belief, string id)
    {
        var run = new ExerciseRun { Id = id, ExerciseId = SplitId, Anchor = new Anchor { Kind = AnchorKind.Shape } };
        Put(run, ShapeAnswers.A, 0, split.A);
        Put(run, ShapeAnswers.Z, 0, split.Z);
        for (var i = 0; i < split.Candidates.Count; i++)
        {
            Put(run, $"{ShapeAnswers.Candidates}[{i}]", 0, split.Candidates[i]);
        }

        Put(run, ShapeAnswers.Midpoint, 0, split.Midpoint);
        Put(run, ShapeAnswers.Stress, 0, split.MidpointAlive);
        Put(run, ShapeAnswers.PP1, 0, split.PP1);
        Put(run, ShapeAnswers.Crisis, 0, split.Crisis);
        Put(run, ShapeAnswers.Inciting, 0, split.Inciting);
        Put(run, ShapeAnswers.Pinch1, 0, split.Pinch1);
        Put(run, ShapeAnswers.Pinch2, 0, split.Pinch2);
        Put(run, ShapeAnswers.LastObstacle, 0, split.LastObstacle);
        if (!split.BeliefLayer)
        {
            // Not asked in the old tool unless it was switched on.
            run.Passed.Add(ShapeAnswers.BeliefLayer);
        }

        if (split.BeliefLayer)
        {
            Put(run, ShapeAnswers.BeliefLayer, 0, "y");
            Put(run, ShapeAnswers.Pinch1ThirdRail, 0, split.Pinch1ThirdRail);
            Put(run, ShapeAnswers.Pinch2ThirdRail, 0, split.Pinch2ThirdRail);
            foreach (var cut in belief.Cuts)
            {
                Put(run, ShapeAnswers.LayerCutKey(cut.Ratio), 0, cut.Text);
            }
        }

        run.Passed.AddRange(new[] { "checks", "mirror" });
        PassStopsBeforeAnswers(ExerciseCatalog.Split, run);
        return run;
    }

    // A stop-callout the old tool showed above text the writer had already written is behind them.
    private static void PassStopsBeforeAnswers(Exercise exercise, ExerciseRun run)
    {
        bool Answered(Step step) => run.Answers.Any(pair => !string.IsNullOrWhiteSpace(pair.Value)
            && (pair.Key == step.Key || pair.Key.StartsWith(step.Key + "[", StringComparison.Ordinal)
                || pair.Key.StartsWith(step.Key + "/", StringComparison.Ordinal)));
        var last = exercise.Steps.ToList().FindLastIndex(step => step.Kind != StepKind.Stop && Answered(step));
        for (var i = 0; i < last; i++)
        {
            if (exercise.Steps[i].Kind == StepKind.Stop && !run.Passed.Contains(exercise.Steps[i].Key))
            {
                run.Passed.Add(exercise.Steps[i].Key);
            }
        }
    }

    public static ExerciseRun FromBelief(BeliefRun belief, string id)
    {
        var run = new ExerciseRun { Id = id, ExerciseId = BeliefId, Anchor = new Anchor { Kind = AnchorKind.Character } };
        Put(run, "shape", 0, belief.Shape);
        Put(run, ShapeAnswers.A, 0, belief.A);
        Put(run, ShapeAnswers.Z, 0, belief.Z);
        Put(run, "ghost", 0, belief.Ghost);
        Put(run, "size", 0, belief.GhostSize);
        Put(run, "fits", 0, belief.GhostCommensurate);
        Put(run, "resize", 0, belief.GhostMismatch);
        Put(run, "chg50", 0, belief.Cut50Change);
        foreach (var cut in belief.Cuts)
        {
            Put(run, ShapeAnswers.CutKey(cut.Ratio), 0, cut.Text);
            Put(run, ShapeAnswers.AliveKey(cut.Ratio), 0, cut.Alive);
        }

        PassStopsBeforeAnswers(ExerciseCatalog.Belief, run);
        return run;
    }

    /// <summary>Extend's links become the rounds of one loop. The old record always ended on a blank
    /// link waiting to be written, which is the round the engine would have appended anyway.</summary>
    public static ExerciseRun FromExtend(ExtendRun extend, string id)
    {
        var run = new ExerciseRun { Id = id, ExerciseId = ExtendId, Anchor = new Anchor { Kind = AnchorKind.Shape } };
        Put(run, ShapeAnswers.Z, 0, extend.Z);
        var written = extend.Links.FindLastIndex(link => !string.IsNullOrWhiteSpace(link.Text)) + 1;
        run.Rounds = Math.Max(1, written + (written > 0 ? 1 : 0));
        for (var r = 0; r < extend.Links.Count; r++)
        {
            var link = extend.Links[r];
            if (r >= run.Rounds)
            {
                break;
            }

            Put(run, "link", r, link.Text);
            if (link.Mystery)
            {
                run.Answers[r == 0 ? "link~alt" : $"link#{r}~alt"] = "y";
            }
        }

        Put(run, "chain", 0, extend.ChainAlive);
        Put(run, "weak", 0, extend.WeakestLink);
        run.Passed.Add("obit");
        return run;
    }

    private static bool IsBlank(FillRun fill) =>
        string.IsNullOrWhiteSpace(fill.PointsTo) && string.IsNullOrWhiteSpace(fill.Answer)
        && fill.Options.All(option => string.IsNullOrWhiteSpace(option.Text) && string.IsNullOrWhiteSpace(option.WhyWrong));

    /// <summary>A Fill run becomes a run anchored on its bracket text: candidates and their
    /// reasons are two Lines steps, in the same order.</summary>
    public static ExerciseRun FromFill(FillRun fill, string id)
    {
        var run = new ExerciseRun
        {
            Id = id,
            ExerciseId = FillId,
            Anchor = new Anchor { Kind = AnchorKind.Bracket, Text = fill.Bracket }
        };

        for (var i = 0; i < fill.Options.Count; i++)
        {
            Put(run, $"cands[{i}]", 0, fill.Options[i].Text);
            Put(run, $"why[{i}]", 0, fill.Options[i].WhyWrong);
        }

        Put(run, "points", 0, fill.PointsTo);
        Put(run, "answer", 0, fill.Answer);
        return run;
    }

    public static ExerciseRun FromChain(ChainRun chain, string id)
    {
        var run = new ExerciseRun
        {
            Id = id,
            ExerciseId = chain.Path == ChainPath.Woac ? WoacId : FlawId,
            Anchor = new Anchor { Kind = AnchorKind.None, Text = chain.Seed }
        };

        if (chain.Path == ChainPath.Woac)
        {
            run.Rounds = Math.Max(1, chain.Rounds.Count);
            for (var r = 0; r < chain.Rounds.Count; r++)
            {
                var round = chain.Rounds[r];
                // Later wants were never stored: they are the carry.
                Put(run, "want", r, r == 0 ? round.Want : string.Empty);
                Put(run, "obstacle", r, round.Obstacle);
                Put(run, "action", r, round.Action);
                Put(run, "consequence", r, round.Consequence);
            }
        }
        else
        {
            for (var q = 0; q < chain.Answers.Count; q++)
            {
                Put(run, $"q{q}", 0, chain.Answers[q]);
            }
        }

        Put(run, "readback", 0, chain.ReadBack);
        return run;
    }

    private static void Put(ExerciseRun run, string key, int round, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            run.Answers[round == 0 ? key : $"{key}#{round}"] = value;
        }
    }
}
