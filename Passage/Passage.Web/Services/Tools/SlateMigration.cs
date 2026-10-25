namespace Passage.Web.Services.Tools;

/// <summary>
/// Sidecar v1 to v2 (docs/SLATE-REDO-PLAN.md R-E). Each old per-tool record
/// becomes an <see cref="ExerciseRun"/> in <see cref="SlateDocument.Exercises"/>.
/// Row T1 migrates Chains only; later rows add the rest. The old lists are
/// left in place until the row that deletes their dock views. Idempotent:
/// run ids are derived from the source, and a run already present is kept.
/// </summary>
public static class SlateMigration
{
    public const int TargetVersion = 2;

    public const string WoacId = "woac";
    public const string FlawId = "flaw";

    /// <summary>Migrates in place. Returns the number of runs added.</summary>
    public static int Migrate(SlateDocument document)
    {
        var added = 0;
        if (document.SlateVersion < TargetVersion)
        {
            for (var i = 0; i < document.Chains.Count; i++)
            {
                var id = $"migrated-chain-{i}";
                if (document.Exercises.Any(run => run.Id == id))
                {
                    continue;
                }

                document.Exercises.Add(FromChain(document.Chains[i], id));
                added++;
            }

            document.SlateVersion = TargetVersion;
        }

        return added;
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
