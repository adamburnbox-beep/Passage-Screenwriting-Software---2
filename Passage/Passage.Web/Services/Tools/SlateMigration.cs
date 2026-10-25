namespace Passage.Web.Services.Tools;

/// <summary>
/// Moves old per-tool records into <see cref="SlateDocument.Exercises"/>
/// (docs/SLATE-REDO-PLAN.md R-E). A move, not a copy: each record type is
/// converted and its old list cleared by the row that deletes that type's dock
/// view, so a second run finds nothing to do. Not gated by version; the
/// version bump waits for T8. Row T1 handles Chains only and is not yet called
/// from <c>SlateStore.Load</c> (T2 wires it in).
/// </summary>
public static class SlateMigration
{
    public const string WoacId = "woac";
    public const string FlawId = "flaw";

    /// <summary>Moves Chains in place. Returns the number of runs moved.</summary>
    public static int Migrate(SlateDocument document)
    {
        var moved = document.Chains.Count;
        foreach (var chain in document.Chains)
        {
            document.Exercises.Add(FromChain(chain, Guid.NewGuid().ToString("N")));
        }

        document.Chains.Clear();
        return moved;
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
