namespace Passage.Web.Services.Tools;

/// <summary>
/// Moves old per-tool records into <see cref="SlateDocument.Exercises"/>
/// (docs/SLATE-REDO-PLAN.md R-E). A move, not a copy: each record type is
/// converted and its old list cleared by the row that deletes that type's dock
/// view, so a second run finds nothing to do. Not gated by version; the
/// version bump waits for T8. Handles Chains (T2), Fills (T4), the Split family (T5), Bridges and Positions (T6) and Revisions (T7); it runs from
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
    public const string BridgeId = "bridge";
    public const string PositionId = "position";
    public const string PushPullId = "pushpull";
    public const string UnderstudyId = "understudy";

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

        // Bridges and Positions: an untouched one is cleared, not moved.
        foreach (var bridge in document.Bridges.Where(bridge => !bridge.IsEmpty))
        {
            document.Exercises.Add(FromBridge(bridge, Guid.NewGuid().ToString("N")));
            moved++;
        }

        document.Bridges.Clear();

        foreach (var position in document.Positions.Where(position => !position.IsEmpty))
        {
            document.Exercises.Add(FromPosition(position, Guid.NewGuid().ToString("N")));
            moved++;
        }

        document.Positions.Clear();

        // Revisions (Push/Pull): an untouched one is cleared, not moved.
        foreach (var revision in document.Revisions.Where(revision => !revision.IsEmpty))
        {
            document.Exercises.Add(FromRevision(revision, Guid.NewGuid().ToString("N")));
            moved++;
        }

        document.Revisions.Clear();

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

    /// <summary>A Bridge run: A and Z as text, each round's candidates and pick as one loop round. Trailing rounds
    /// with nothing written are not carried over; a last round that was picked gets the empty round the engine
    /// would have appended. The two cards were never stored, so the run names its ends by their text.</summary>
    public static ExerciseRun FromBridge(BridgeRun bridge, string id)
    {
        var a = bridge.A.Trim();
        var z = bridge.Z.Trim();
        var run = new ExerciseRun
        {
            Id = id,
            ExerciseId = BridgeId,
            Anchor = new Anchor { Kind = AnchorKind.TwoPoints, Text = a.Length == 0 && z.Length == 0 ? string.Empty : $"{(a.Length == 0 ? "?" : a)} → {(z.Length == 0 ? "?" : z)}" }
        };

        Put(run, "a", 0, bridge.A);
        Put(run, "z", 0, bridge.Z);
        var written = bridge.Rounds.FindLastIndex(round => !round.IsEmpty) + 1;
        var picked = written > 0 && !string.IsNullOrWhiteSpace(bridge.Rounds[written - 1].Picked);
        run.Rounds = Math.Max(1, written + (picked ? 1 : 0));
        for (var r = 0; r < written; r++)
        {
            var round = bridge.Rounds[r];
            for (var i = 0; i < round.Candidates.Count; i++)
            {
                Put(run, r == 0 ? $"links[{i}]" : $"links#{r}[{i}]", 0, round.Candidates[i]);
            }

            Put(run, "pick", r, round.Picked);
        }

        Put(run, "wire", 0, bridge.Wire);
        return run;
    }

    /// <summary>A Position run: the moment, then each tried turn's before / after / alive-or-flat under its turn
    /// label, and the shortlist. A turn with text but no slot named has nowhere to go in the seven, so its text
    /// is kept on the moment rather than dropped.</summary>
    public static ExerciseRun FromPosition(PositionRun position, string id)
    {
        var run = new ExerciseRun
        {
            Id = id,
            ExerciseId = PositionId,
            Anchor = new Anchor { Kind = AnchorKind.None, Text = string.Join(' ', position.Moment.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)).Trim() }
        };

        var moment = position.Moment.Trim();
        foreach (var turn in position.Turns.Where(turn => !turn.IsEmpty))
        {
            var slot = ExerciseCatalog.TurnLabels.FirstOrDefault(label => label == turn.Slot);
            if (slot is null)
            {
                var loose = string.Join("; ", new[] { turn.Before, turn.After, turn.Alive }.Where(text => !string.IsNullOrWhiteSpace(text)));
                if (loose.Length > 0)
                {
                    moment += (moment.Length == 0 ? string.Empty : "\n") + $"(a turn with no slot named: {loose})";
                }

                continue;
            }

            // A slot tried twice keeps both writings; the mark is the later one.
            Join(run, $"before/{slot}", turn.Before);
            Join(run, $"after/{slot}", turn.After);
            Put(run, $"alive/{slot}", 0, turn.Alive);
        }

        Put(run, "moment", 0, moment);
        for (var i = 0; i < position.Shortlist.Count; i++)
        {
            Put(run, $"shortlist[{i}]", 0, position.Shortlist[i]);
        }

        return run;
    }

    /// <summary>
    /// A Push/Pull run, kept by the scene it named. GSU rows land under their character (a row with no name and some
    /// text has no one to sit under, so it joins the purpose rather than vanishing); each check's answer becomes the
    /// "problem" line behind its flag, its Lens answers keep their place. A Y/N the writer never set stays unset.
    /// Whatever the old view had open (the deeper sections) has no marker, so the gates are opened only where there is text.
    /// </summary>
    public static ExerciseRun FromRevision(ReviseRun revision, string id)
    {
        var scene = Flatten(revision.Scene);
        var run = new ExerciseRun
        {
            Id = id,
            ExerciseId = PushPullId,
            Anchor = new Anchor { Kind = AnchorKind.Scene, Text = scene, SceneText = scene }
        };

        var purpose = revision.Purpose.Trim();
        foreach (var row in revision.Gsu.Where(row => !row.IsEmpty))
        {
            var who = row.Character.Trim();
            if (who.Length == 0)
            {
                var loose = string.Join("; ", new[] { row.Goal, row.Stakes, row.Urgency }.Where(text => !string.IsNullOrWhiteSpace(text)));
                if (loose.Length > 0)
                {
                    purpose += (purpose.Length == 0 ? string.Empty : "\n") + $"(a GSU row with no character named: {loose})";
                }

                continue;
            }

            Put(run, $"gsu-goal/{who}", 0, row.Goal);
            Put(run, $"gsu-stakes/{who}", 0, row.Stakes);
            Put(run, $"gsu-urgency/{who}", 0, row.Urgency);
        }

        Put(run, "purpose", 0, purpose);
        Put(run, "gsu", 0, YesNo(revision.GsuFlag));
        for (var i = 0; i < ExerciseCatalog.CheckNames.Length; i++)
        {
            var check = revision.Check(ExerciseCatalog.CheckNames[i]);
            var key = "c" + (i + 1);
            Put(run, key, 0, YesNo(check.Flag));
            Put(run, key + "-why", 0, check.Answer);
            Put(run, key + "-lens", 0, check.Lens);
            Put(run, key + "-frag", 0, check.Fragment);
            Put(run, key + "-moved", 0, YesNo(check.Moved));
            Put(run, key + "-movedwhy", 0, check.MovedWhy);
        }

        if (revision.HasDiagnostic)
        {
            Put(run, "diag", 0, "y");
            Put(run, "diag-a", 0, revision.DiagA);
            Put(run, "diag-z", 0, revision.DiagZ);
            Put(run, "diag-shape", 0, revision.DiagShape);
            Put(run, "diag-match", 0, revision.DiagMatch);
            Put(run, "diag-seam", 0, revision.DiagSeam);
            Put(run, "diag-lens", 0, revision.DiagLens);
            Put(run, "diag-frag", 0, revision.DiagFragment);
        }

        Put(run, "working", 0, revision.Working);
        Put(run, "notsitting", 0, revision.NotSitting);
        Put(run, "ready", 0, revision.Ready);
        PassStopsBeforeAnswers(ExerciseCatalog.PushPull, run);
        return run;
    }

    private static string Flatten(string text) =>
        string.Join(' ', text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)).Trim();

    // The old Y / N flags were upper case; the engine's yes / no marks are lower case.
    private static string YesNo(string flag) => flag.Trim().ToLowerInvariant() switch { "y" => "y", "n" => "n", _ => string.Empty };

    private static void Join(ExerciseRun run, string key, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        run.Answers[key] = run.Answers.TryGetValue(key, out var existing) ? existing + "\n" + value : value;
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
