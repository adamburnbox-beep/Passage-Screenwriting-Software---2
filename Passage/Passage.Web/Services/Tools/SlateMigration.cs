using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Passage.Web.Services.Tools;

/// <summary>
/// Version 1 sidecars kept one record list per tool (Chains, Fills, Split, Belief,
/// Extend, Bridges, Positions, Revisions). This turns them into
/// <see cref="SlateDocument.Exercises"/> (docs/SLATE-REDO-PLAN.md R-E) when the sidecar
/// is loaded. <see cref="SlateDocument"/> has no such properties any more, so the old
/// lists are read from the raw JSON into the private DTOs at the bottom of this class,
/// which are the old record types kept only for that.
/// A sidecar already at version 2 is left alone. A version 1 sidecar written by one of
/// the T2–T7 builds already holds some runs in <c>exercises</c> and empty old lists;
/// those lists convert to nothing, so its runs are kept as they are.
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

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    /// <summary>
    /// Appends a run to <paramref name="document"/> for every old record in <paramref name="json"/>
    /// (the sidecar as read from disk) and stamps the document with the current version.
    /// Returns the number of runs added; 0 and no change when the sidecar is already current.
    /// </summary>
    public static int Migrate(SlateDocument document, string json)
    {
        var root = JsonNode.Parse(json) as JsonObject;
        var version = root?.FirstOrDefault(pair => pair.Key.Equals("slateVersion", StringComparison.OrdinalIgnoreCase)).Value is JsonValue value
            && value.TryGetValue<int>(out var read) ? read : 1;
        if (root is null || version >= SlateStore.CurrentVersion)
        {
            return 0;
        }

        var legacy = root.Deserialize<OldDocument>(JsonOptions) ?? new OldDocument();
        legacy.Fills ??= new();
        legacy.Chains ??= new();
        legacy.Split ??= new();
        legacy.Belief ??= new();
        legacy.Extend ??= new();
        legacy.Bridges ??= new();
        legacy.Positions ??= new();
        legacy.Revisions ??= new();
        var moved = 0;
        foreach (var chain in legacy.Chains.Where(chain => !chain.IsEmpty))
        {
            document.Exercises.Add(FromChain(chain, NewId()));
            moved++;
        }

        // A Fill with nothing written was never worth keeping (the old view dropped it too).
        foreach (var fill in legacy.Fills.Where(fill => !IsBlank(fill)))
        {
            document.Exercises.Add(FromFill(fill, NewId()));
            moved++;
        }

        // The Split family was one run each per script. An untouched one is not worth a kept run.
        // Split's belief layer asked the same five cuts Belief Split holds, so when the layer was on
        // the cuts are copied into the Split run as well (they are two exercises now, not one record).
        if (!IsBlank(legacy.Split))
        {
            document.Exercises.Add(FromSplit(legacy.Split, legacy.Belief, NewId()));
            moved++;
        }

        if (!IsBlank(legacy.Belief))
        {
            document.Exercises.Add(FromBelief(legacy.Belief, NewId()));
            moved++;
        }

        if (!IsBlank(legacy.Extend))
        {
            document.Exercises.Add(FromExtend(legacy.Extend, NewId()));
            moved++;
        }

        foreach (var bridge in legacy.Bridges.Where(bridge => !bridge.IsEmpty))
        {
            document.Exercises.Add(FromBridge(bridge, NewId()));
            moved++;
        }

        foreach (var position in legacy.Positions.Where(position => !position.IsEmpty))
        {
            document.Exercises.Add(FromPosition(position, NewId()));
            moved++;
        }

        foreach (var revision in legacy.Revisions.Where(revision => !revision.IsEmpty))
        {
            document.Exercises.Add(FromRevision(revision, NewId()));
            moved++;
        }

        document.SlateVersion = SlateStore.CurrentVersion;
        return moved;
    }

    private static string NewId() => Guid.NewGuid().ToString("N");

    private static bool IsBlank(OldSplit split) =>
        new[] { split.A, split.Z, split.Midpoint, split.PP1, split.Crisis, split.Inciting, split.Pinch1, split.Pinch2, split.LastObstacle }.All(string.IsNullOrWhiteSpace) && split.Candidates.All(string.IsNullOrWhiteSpace)
        && string.IsNullOrWhiteSpace(split.MidpointAlive)
        && string.IsNullOrWhiteSpace(split.Pinch1ThirdRail) && string.IsNullOrWhiteSpace(split.Pinch2ThirdRail);

    private static bool IsBlank(OldBelief belief) =>
        new[] { belief.Shape, belief.A, belief.Z, belief.Ghost, belief.GhostSize, belief.GhostCommensurate, belief.GhostMismatch, belief.Cut50Change }
            .All(string.IsNullOrWhiteSpace)
        && belief.Cuts.All(cut => string.IsNullOrWhiteSpace(cut.Text) && string.IsNullOrWhiteSpace(cut.Alive));

    private static bool IsBlank(ExtendRun extend) =>
        string.IsNullOrWhiteSpace(extend.Z) && extend.Links.All(link => string.IsNullOrWhiteSpace(link.Text))
        && string.IsNullOrWhiteSpace(extend.ChainAlive) && string.IsNullOrWhiteSpace(extend.WeakestLink);

    /// <summary>A Split run keeps its answers under the keys <see cref="ShapeAnswers"/> reads back.</summary>
    private static ExerciseRun FromSplit(OldSplit split, OldBelief belief, string id)
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

    private static ExerciseRun FromBelief(OldBelief belief, string id)
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
    private static ExerciseRun FromExtend(ExtendRun extend, string id)
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
    private static ExerciseRun FromBridge(BridgeRun bridge, string id)
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
    private static ExerciseRun FromPosition(PositionRun position, string id)
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
    private static ExerciseRun FromRevision(ReviseRun revision, string id)
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
    private static ExerciseRun FromFill(FillRun fill, string id)
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

    private static ExerciseRun FromChain(ChainRun chain, string id)
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

    // ---- The version 1 records, read from the raw JSON and nowhere else ----

    private sealed class OldDocument
    {
        public List<FillRun> Fills { get; set; } = new();
        public List<ChainRun> Chains { get; set; } = new();
        public OldSplit Split { get; set; } = new();
        public OldBelief Belief { get; set; } = new();
        public ExtendRun Extend { get; set; } = new();
        public List<BridgeRun> Bridges { get; set; } = new();
        public List<PositionRun> Positions { get; set; } = new();
        public List<ReviseRun> Revisions { get; set; } = new();
    }

    /// <summary>Push/Pull and Lens ("Revise — Push-Pull and Lens"): check 0
    /// (GSU) and six checks on one scene or stretch, a Lens where a check is
    /// flagged, and the optional diagnostic belief layer and read-back beneath.</summary>
    private sealed class ReviseRun
    {
        public static readonly string[] CheckNames = { "Polarity", "Linkage", "Third rail", "Migration", "Escalation", "Two-test" };

        public string Scene { get; set; } = string.Empty;
        public string Purpose { get; set; } = string.Empty;

        // Check 0 — GSU. Not one of CheckNames: a flag here points at a Scene
        // Dial, never a Lens. Two rows to start, as the worksheet prints them.
        public List<GsuRow> Gsu { get; set; } = new() { new GsuRow(), new GsuRow() };
        public string GsuFlag { get; set; } = string.Empty;

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
            && Gsu.All(row => row.IsEmpty) && string.IsNullOrWhiteSpace(GsuFlag)
            && Checks.All(check => check.IsEmpty) && !HasDiagnostic && !HasReadBack;
    }

    /// <summary>One character's line in check 0: goal, stakes, urgency.</summary>
    private sealed class GsuRow
    {
        public string Character { get; set; } = string.Empty;
        public string Goal { get; set; } = string.Empty;
        public string Stakes { get; set; } = string.Empty;
        public string Urgency { get; set; } = string.Empty;

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(Character) && string.IsNullOrWhiteSpace(Goal)
            && string.IsNullOrWhiteSpace(Stakes) && string.IsNullOrWhiteSpace(Urgency);
    }

    /// <summary>One check: its answer and Y/N flag, and — only where flagged —
    /// the Lens applied to that flagged thing.</summary>
    private sealed class ReviseCheck
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
    private sealed class BridgeRun
    {
        public string A { get; set; } = string.Empty;
        public string Z { get; set; } = string.Empty;
        public List<BridgeRound> Rounds { get; set; } = new() { new() };
        public string Wire { get; set; } = string.Empty;

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(A) && string.IsNullOrWhiteSpace(Z) && string.IsNullOrWhiteSpace(Wire)
            && Rounds.All(round => round.IsEmpty);
    }

    private sealed class BridgeRound
    {
        public List<string> Candidates { get; set; } = new() { string.Empty, string.Empty, string.Empty };
        public string Picked { get; set; } = string.Empty;

        public bool IsEmpty => string.IsNullOrWhiteSpace(Picked) && Candidates.All(string.IsNullOrWhiteSpace);
    }

    /// <summary>Position ("Place or Build — Position or Fill", Path A): one
    /// moment tried in the story's seven slots, one turn at a time.</summary>
    private sealed class PositionRun
    {
        public string Moment { get; set; } = string.Empty;
        public List<PositionTurn> Turns { get; set; } = new() { new() };
        public List<string> Shortlist { get; set; } = new() { string.Empty, string.Empty };

        public const int SlotCount = 7;

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(Moment) && Shortlist.All(string.IsNullOrWhiteSpace) && Turns.All(turn => turn.IsEmpty);
    }

    private sealed class PositionTurn
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

    [JsonConverter(typeof(JsonStringEnumConverter))]
    private enum ChainPath { Woac, Flaw }

    /// <summary>
    /// One forward chain (the "Place or Build — Forward Chain" worksheet). WOAC
    /// stores its rounds; the Flaw brainstorm stores its eight answers. Every
    /// field is optional and a chain can be left at any point.
    /// </summary>
    private sealed class ChainRun
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
    private sealed class OldSplit
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
    }

    /// <summary>Path B — Belief Split. The five named cuts, by ratio.</summary>
    private sealed class OldBelief
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

    private sealed class BeliefCut
    {
        public static readonly int[] Ratios = { 12, 25, 50, 75, 88 };

        public int Ratio { get; set; }
        public string Text { get; set; } = string.Empty;
        public string Alive { get; set; } = string.Empty;
    }

    /// <summary>Path C — Extend Backward. Links run from Z backwards, one at a time.</summary>
    private sealed class ExtendRun
    {
        public string Z { get; set; } = string.Empty;
        public List<ExtendLink> Links { get; set; } = new() { new() };
        public string ChainAlive { get; set; } = string.Empty;
        public string WeakestLink { get; set; } = string.Empty;
    }

    private sealed class ExtendLink
    {
        public string Text { get; set; } = string.Empty;

        // The mystery-lens variant: Z as a crime scene, what would a detective
        // need to find — a different question for this link only.
        public bool Mystery { get; set; }
    }

    private sealed class ChainRound
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
    private sealed class FillRun
    {
        public string Bracket { get; set; } = string.Empty;
        public List<FillOption> Options { get; set; } = new() { new(), new(), new() };
        public string PointsTo { get; set; } = string.Empty;
        public string Answer { get; set; } = string.Empty;
    }

    private sealed class FillOption
    {
        public string Text { get; set; } = string.Empty;
        public string WhyWrong { get; set; } = string.Empty;
    }
}
