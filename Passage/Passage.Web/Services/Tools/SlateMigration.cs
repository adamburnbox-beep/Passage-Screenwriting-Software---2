namespace Passage.Web.Services.Tools;

/// <summary>
/// Moves old per-tool records into <see cref="SlateDocument.Exercises"/>
/// (docs/SLATE-REDO-PLAN.md R-E). A move, not a copy: each record type is
/// converted and its old list cleared by the row that deletes that type's dock
/// view, so a second run finds nothing to do. Not gated by version; the
/// version bump waits for T8. Handles Chains (T2) and Fills (T4); it runs from
/// <c>SlateStore.Load</c>.
/// </summary>
public static class SlateMigration
{
    public const string WoacId = "woac";
    public const string FlawId = "flaw";
    public const string FillId = "fill";
    public const string SceneDialsId = "scenedials";

    /// <summary>Moves Chains and Fills in place. Returns the number of runs moved.</summary>
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
        return moved;
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
