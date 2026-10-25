namespace Passage.Web.Services.Tools;

/// <summary>
/// The deck (docs/SLATE-REDO-PLAN.md §4.5): one static definition per
/// exercise. Questions keep the vault worksheets' words (R-B); anything
/// longer sits in Why. Row T2 defines WOAC and Flaw, T4 Fill and Scene Dials,
/// T5 the Split family; later rows add the rest.
/// </summary>
public static class ExerciseCatalog
{
    /// <summary>The Slots source that lists the people who speak in the run's scene.</summary>
    public const string SpeakersSource = "speakers";

    /// <summary>The Slots source that lists the story's seven turns, in order.</summary>
    public const string TurnsSource = "turns";

    /// <summary>The Position step whose answer is the turn Send to the Board files the beat under.</summary>
    public const string PositionTurnKey = "turn";

    /// <summary>The seven turns' labels, as the Shape names them.</summary>
    public static IReadOnlyList<string> TurnLabels { get; } = SplitScript.Turns.Select(slot => slot.Label).ToList();

    /// <summary>The Fill step whose answer replaces the bracket.</summary>
    public const string FillAnswerKey = "answer";

    public static readonly Exercise Woac = new()
    {
        Id = SlateMigration.WoacId,
        Name = "WOAC chain",
        Tagline = "What happens next, round by round",
        Shelf = Shelf.SceneGoing,
        Anchor = AnchorKind.Character,
        SceneToo = true,
        AnchorLabel = "Whose want?",
        Targets = new[] { TargetKind.InboxBeat, TargetKind.Synopsis, TargetKind.NotePage, TargetKind.Copy },
        Why = "Want → Obstacle → Action → Consequence. Each round: what does the character want right now, "
            + "what's in their way, what do they do about it, and what happens as a result? That consequence "
            + "becomes next round's want automatically, so you never have to invent a new question, just carry "
            + "it down. Do one round or ten, stop whenever, and read it back before you close. A scene usually "
            + "finds its shape somewhere around 5–7 rounds, but that's a pattern, not a rule.",
        Steps = new[]
        {
            new Step { Key = "want", Label = "Want", Group = "round", Offer = true,
                Question = "What does the character want right now?", Hint = "Round one only; after that it's carried." },
            new Step { Key = "obstacle", Label = "Obstacle", Group = "round", Offer = true,
                Question = "What's in their way?" },
            new Step { Key = "action", Label = "Action", Group = "round", Offer = true,
                Question = "What do they do about it?", Hint = "Action, not feeling." },
            new Step { Key = "consequence", Label = "Consequence", Group = "round", Offer = true,
                Question = "What happens as a result?", Hint = "This becomes the next want." },
            new Step { Key = "loop", Kind = StepKind.Loop, LoopGroup = "round", CarryFrom = "consequence", CarryTo = "want" },
            new Step { Key = "readback", Label = "Read it back", Multiline = true,
                Question = "Read it back — what's the scene now, in one line?" }
        }
    };

    private static readonly string[] FlawQuestions =
    {
        "Worst situation the flaw causes",
        "The character's first (wrong) response",
        "How that backfires",
        "Who unexpectedly helps",
        "The next (also wrong) action",
        "The obstacle that creates",
        "The flaw reframed as a hidden skill",
        "A surprising final action"
    };

    public static readonly Exercise Flaw = new()
    {
        Id = SlateMigration.FlawId,
        Name = "Character Flaw Brainstorm",
        Tagline = "One flaw, eight questions, a surprise",
        Shelf = Shelf.SceneGoing,
        Anchor = AnchorKind.Character,
        AnchorLabel = "Whose flaw?",
        Targets = new[] { TargetKind.NotePage, TargetKind.InboxBeat, TargetKind.Copy },
        Why = "Eight fixed questions that take one character flaw from the worst situation it causes to a "
            + "surprising final action. Same chain grammar as WOAC, but a fixed-length loop, each question "
            + "different from the last. You can stop after any question — a partial chain is still real "
            + "material — but the eight are designed to land on a shape together, so if you're mid-flow, "
            + "finishing the set in one sitting is worth more here than with WOAC's open-ended rounds.",
        Steps = FlawQuestions
            .Select((question, i) => new Step { Key = $"q{i}", Label = question, Question = question, Offer = true })
            .Append(new Step { Key = "readback", Label = "Read it back", Multiline = true,
                Question = "Read it back — what's the scene now, in one line?" })
            .ToList()
    };

    // The six dials, "Name — what you do", from the vault's Scene Dials table. Fill's
    // "Stage it" group and the standalone Scene Dials roll the same table.
    private static readonly string[] Dials =
    {
        "Opposed wants — Two people want different things out of this scene.",
        "Pre-agitator — The conflict starts before the scene does. They walk in mid-argument, or one has just had news the other doesn't know.",
        "Agitator — Something drops into the middle and makes the goal harder: a kid who won't stop talking, a phone that keeps ringing, rain.",
        "Hidden thing — One character is hiding something from the other and can't mention it. The subtext writes itself.",
        "Wrong room — Move it to the last place they'd want to have this conversation, or somewhere they can't talk freely: a funeral, a lift, a church pew.",
        "Collision ahead — Show us something coming that they can't see: a car on its way, someone walking up the drive. Everything before it now plays twice."
    };

    private const string DialsWhy = "The event is right but the scene around it is two people saying what they mean in a kitchen. "
        + "That's a different problem from a scene that's gone wrong: this builds a scene around an event that's right. "
        + "If anyone in the room wants nothing, that's usually the whole problem, and giving them a want is the fix before any dial. "
        + "Then turn just one dial.";

    // The three Scene Dials steps under a key prefix, so Fill's optional "Stage it" and the standalone
    // exercise are the same questions.
    private static IEnumerable<Step> DialSteps(string prefix) => new[]
    {
        new Step { Key = prefix + "wants", Kind = StepKind.Slots, Source = SpeakersSource, Label = "Wants",
            Question = "What does each character want out of this scene, right now?",
            Hint = "This scene, not their story goal. If anyone wants nothing, give them something first." },
        new Step { Key = prefix + "dial", Kind = StepKind.Dice, Table = Dials, Label = "Dial",
            Question = "Roll one die and turn that one dial.", Hint = "Just the one." },
        new Step { Key = prefix + "scene", Label = "The scene", Multiline = true, Offer = true,
            Question = "The scene with that dial turned — a few lines or a paragraph.",
            Hint = "A few lines is enough; this is a sketch, not the scene." }
    };

    public static readonly Exercise Fill = new()
    {
        Id = SlateMigration.FillId,
        Name = "Fill",
        Tagline = "Turn a [BRACKET] into the real event",
        Shelf = Shelf.Moment,
        Anchor = AnchorKind.Bracket,
        AnchorLabel = "Which bracket?",
        Targets = new[] { TargetKind.Synopsis, TargetKind.Copy },
        ShortcutKey = FillAnswerKey,
        Why = "A bracket is a placeholder like [SOMETHING happens] sitting in your outline or draft — a gap you marked "
            + "instead of stalling on it. Filling it means swapping it for the real, concrete event, not a better-worded "
            + "placeholder. Invent three candidate events, each one you expect to be wrong, then write down why each is "
            + "wrong. Those reasons are the actual work: together they narrow down what the real answer has to do, even "
            + "before you know what it is. Most of the time the real answer shows up before the third candidate; when it "
            + "does, write it down — you don't owe the questions anything else. Filling brackets is meant to chain, one "
            + "after another, for as long as you want; the stop is yours to call.",
        Steps = new[]
        {
            new Step { Key = "cands", Kind = StepKind.Lines, Count = 3, Label = "Candidates", Offer = true,
                Question = "Three things that could happen here, each one you expect to be wrong.",
                Hint = "The real answer often shows up before the third." },
            new Step { Key = "why", Kind = StepKind.Lines, Count = 3, Label = "Why wrong",
                Question = "Why is each one probably wrong? One line each, in the same order." },
            new Step { Key = "points", Label = "Points to", Multiline = true,
                Question = "What does the wrongness point to?",
                Hint = "What do your reasons agree the real event needs to do, or needs to avoid?" },
            new Step { Key = FillAnswerKey, Label = "The real answer", Multiline = true,
                Question = "The real answer — if it showed up already, write it here.",
                Hint = "A concrete, filmable event: [SOMETHING forces her to stay] becomes her car won't start, and the last bus already left." },
            new Step { Key = "stage-stop", Kind = StepKind.Stop,
                Text = "You know what happens. If the scene around it will play flat — two people in a room saying what they mean — stage it. Skip this if the scene's already alive in your head." }
        }.Concat(DialSteps("stage-")).ToList()
    };

    public static readonly Exercise SceneDials = new()
    {
        Id = SlateMigration.SceneDialsId,
        Name = "Scene Dials",
        Tagline = "The event is right; the scene is dead",
        Shelf = Shelf.SceneGoing,
        Anchor = AnchorKind.Scene,
        AnchorLabel = "Which scene?",
        Targets = new[] { TargetKind.Synopsis, TargetKind.Copy },
        Why = DialsWhy,
        Steps = DialSteps(string.Empty).ToList()
    };


    // ---- The Split family (T5): the worksheet's Path A, B and C ----

    private static Step Gate(string key, string question, string hint, string label) =>
        new() { Key = key, Kind = StepKind.YesNo, Label = label, Question = question, Hint = hint };

    private static Step IfYes(string key, string group) =>
        new() { Key = key + "-if", Kind = StepKind.Branch, BranchGroup = group, OnKey = key, OnValue = "y" };

    private static Step LayerCut(string key, int ratio, string question) =>
        new() { Key = key, Group = "belief", Label = $"Belief {ratio}%", Question = question,
            Hint = "What does she believe now, and how has that moved since the last cut?" };

    public static readonly Exercise Split = new()
    {
        Id = SlateMigration.SplitId,
        Name = "A→Z→Split",
        Tagline = "Find the middle of the story",
        Shelf = Shelf.Ending,
        Anchor = AnchorKind.Shape,
        AnchorLabel = "The Shape",
        Targets = new[] { TargetKind.Copy },
        Why = "Never “what happens next”. Always: I'm at A, I need Z — what single event makes Z possible that wasn't before? "
            + "Fix the ends, then do one split at a time. Stop after any one of them — that's a complete session. "
            + "Keep going and the next rung is right there. Write the shape into the script at any point: a rung-1 write "
            + "gives the Board a split story with six brackets to fill.",
        Steps = new[]
        {
            new Step { Key = ShapeAnswers.A, Label = "A", ScriptSlot = "Starts",
                Question = "A — the start event.", Hint = "Something happening, not a feeling." },
            new Step { Key = ShapeAnswers.Z, Label = "Z", ScriptSlot = "Ends",
                Question = "Z — the end event." },
            Gate("checks", "Two ten-second checks on the ends — run them?", "Optional. Either, both or neither.", "Checks"),
            IfYes("checks", "checks"),
            new Step { Key = "irony", Group = "checks", Label = "Irony",
                Question = "Say the story's condition as an ironic noun phrase.",
                Hint = "A genius janitor. A king who can't speak.",
                Why = "If one comes easily, the condition has a charge. If nothing comes, it may be stable but not yet wrong." },
            new Step { Key = "objection", Group = "checks", Label = "Objection",
                Question = "Say your logline back as the harshest executive in the room, trying to kill it.",
                Hint = "“Why would anyone greenlight this?”" },
            new Step { Key = "answer", Group = "checks", Label = "Answer",
                Question = "Now answer the objection in one line, no hedging." },
            Gate("mirror", "Mirror Z — write the opposite ending?", "Optional. Before or after the midpoint, either works.", "Mirror Z"),
            IfYes("mirror", "mirror"),
            new Step { Key = "opp", Group = "mirror", Label = "Opposite Z", Multiline = true,
                Question = "Write the opposite ending, in full, as if it were the real one.",
                Hint = "Not a note: an actual paragraph." },
            new Step { Key = "diff", Group = "mirror", Label = "What changes",
                Question = "What would the middle need to do differently to earn this Z instead of the real one?",
                Why = "If the honest answer is “not much”, the middle isn't pointed anywhere yet, it's just walking toward whichever Z "
                    + "you happened to pick. Worth knowing now, before three rungs of splitting, not after." },
            new Step { Key = ShapeAnswers.Candidates, Kind = StepKind.Lines, Count = 3, Label = "Candidates", Offer = true,
                Question = "The one event that, if it happened, makes the second half inevitable given the first.",
                Hint = "Not the best midpoint — the first one that holds." },
            new Step { Key = ShapeAnswers.Midpoint, Kind = StepKind.Pick, FromKey = ShapeAnswers.Candidates, Label = "Midpoint", ScriptSlot = "Midpoint",
                Question = "Which one is the midpoint?", Hint = "Tap one, or write your own." },
            new Step { Key = ShapeAnswers.Stress, Kind = StepKind.AliveFlat, Label = "Stress",
                Question = "Given A, does the midpoint feel inevitable or invented? Given the midpoint, does Z?" },
            Gate(ShapeAnswers.BeliefLayer, "Also ask what she believes at each turn?", "Optional. The pinches take a lighter third-rail check.", "Belief layer"),
            IfYes(ShapeAnswers.BeliefLayer, "belief"),
            LayerCut(ShapeAnswers.LayerCutKey(50), 50, "Belief at the midpoint (50%)."),
            new Step { Key = "stop1", Kind = StepKind.Stop,
                Text = "Stop here and you've got a real shape — A, midpoint, Z. That's enough to carry into Fill or a later session. Nothing below is owed." },
            new Step { Key = ShapeAnswers.PP1, Label = "PP1", ScriptSlot = "Plot point 1",
                Question = "PP1 — the event that splits A → Midpoint.", Hint = "Same move as the midpoint, run twice." },
            LayerCut(ShapeAnswers.LayerCutKey(25), 25, "Belief at PP1 (25%)."),
            new Step { Key = ShapeAnswers.Crisis, Label = "PP2", ScriptSlot = "Crisis / Plot point 2",
                Question = "PP2 (crisis) — the event that splits Midpoint → Z." },
            LayerCut(ShapeAnswers.LayerCutKey(75), 75, "Belief at PP2 (75%)."),
            new Step { Key = "stop2", Kind = StepKind.Stop, Text = "Stop here and four points are fixed — still a complete, usable shape." },
            new Step { Key = ShapeAnswers.Inciting, Label = "Inciting incident", ScriptSlot = "Inciting incident",
                Question = "Inciting incident — the event that splits A → PP1." },
            LayerCut(ShapeAnswers.LayerCutKey(12), 12, "Belief at the inciting incident (12%)."),
            new Step { Key = ShapeAnswers.Pinch1, Label = "Pinch 1", ScriptSlot = "Pinch 1",
                Question = "Pinch 1 — the event that splits PP1 → Midpoint." },
            new Step { Key = ShapeAnswers.Pinch1ThirdRail, Group = "belief", Label = "Third rail",
                Question = "Third rail — what does this mean to her?", Hint = "What does she believe now that she didn't five minutes ago?" },
            new Step { Key = ShapeAnswers.Pinch2, Label = "Pinch 2", ScriptSlot = "Pinch 2",
                Question = "Pinch 2 — the event that splits Midpoint → PP2." },
            new Step { Key = ShapeAnswers.Pinch2ThirdRail, Group = "belief", Label = "Third rail",
                Question = "Third rail — what does this mean to her?", Hint = "What does she believe now that she didn't five minutes ago?" },
            new Step { Key = ShapeAnswers.LastObstacle, Label = "Pre-climax", ScriptSlot = "Last obstacle",
                Question = "Pre-climax — the event that splits PP2 → Z." },
            LayerCut(ShapeAnswers.LayerCutKey(88), 88, "Belief at the pre-climax (88%)."),
            new Step { Key = "stop3", Kind = StepKind.Stop,
                Text = "All seven turns down. That's the full shape; there's no rung 4. If it still doesn't hold, that's Fill and Push/Pull's job, not another rung here." }
        }
    };

    // The worksheet's five shapes: "Name — A → Z". The answer stored is the Name.
    public static readonly string[] BeliefShapes =
    {
        "Positive Change — Lie → Truth",
        "Flat — Truth held → used to overcome the world's Lie",
        "Disillusionment — Lie → Truth (tragic)",
        "Fall — Lie → clings, then a worse Lie",
        "Corruption — sees the Truth → rejects it, embraces the Lie"
    };

    private static Step BeliefCut(string key, int ratio, string turn, string question) =>
        new() { Key = ShapeAnswers.CutKey(ratio), Label = $"Belief {ratio}%", PlotTurn = turn, Question = question };

    private static Step CastStep(string key, string label, string question, string hint = "", string why = "") =>
        new() { Key = key, Group = "cast", Label = label, Question = question, Hint = hint, Why = why, WhoKey = "who" };

    public static readonly Exercise Belief = new()
    {
        Id = SlateMigration.BeliefId,
        Name = "Belief Split",
        Tagline = "Split on what a character believes",
        Shelf = Shelf.Ending,
        Anchor = AnchorKind.Character,
        AnchorLabel = "Whose belief?",
        Targets = new[] { TargetKind.NotePage, TargetKind.Copy },
        Why = "The “Lie” is the false belief the character organises their life around at the start; the “Truth” is what they "
            + "actually need to believe by the end. Same fix-the-ends move as A→Z→Split, run on belief instead of plot events. "
            + "It only defines five named cuts (12, 25, 50, 75, 88%); the two pinches take the lighter third-rail question in A→Z→Split.",
        Steps = new[]
        {
            new Step { Key = "shape", Kind = StepKind.Pick, Options = BeliefShapes, Label = "Shape",
                Question = "Which shape is her story?", Hint = "The shape fixes what A and Z are.",
                Why = "Positive Change: Lie → Truth. Flat: Truth held, used to overcome the world's Lie. Disillusionment: Lie → Truth (tragic). "
                    + "Fall: Lie → clings, then a worse Lie. Corruption: sees the Truth, rejects it, embraces the Lie." },
            new Step { Key = ShapeAnswers.A, Label = "A", Question = "A — Lie / Truth at the start." },
            new Step { Key = ShapeAnswers.Z, Label = "Z", Question = "Z — Truth / worse Lie at the end." },
            new Step { Key = "ghost", Label = "Origin",
                Question = "What happened before the story that made A feel true?",
                Hint = "An event, a streak, a habit, or “none needed”. Take the first answer that feels true." },
            new Step { Key = "size", Label = "Size", Question = "Size it, 1–10." },
            Gate("fits", "Gut check: does it make A's behaviour in Act 1 feel right?", string.Empty, "Gut check"),
            new Step { Key = "fits-if", Kind = StepKind.Branch, BranchGroup = "resize", OnKey = "fits", OnValue = "n" },
            new Step { Key = "resize", Group = "resize", Label = "Resize",
                Question = "Resize it, or drop it and let A stand on its own." },
            new Step { Key = "stop-origin", Kind = StepKind.Stop,
                Text = "Run this rung once per story, not once per session. Don't dig below the first answer: if you're asking “but why, really?” a second time, write the next scene instead." },
            BeliefCut("cut50", 50, "Midpoint", "What does she believe right now — and what's changed since A?"),
            new Step { Key = "chg50", Label = "What changed", Question = "What changed to get here?" },
            new Step { Key = ShapeAnswers.AliveKey(50), Kind = StepKind.AliveFlat, Label = "Stress",
                Question = "Does this feel like a real step from A, or a jump?" },
            new Step { Key = "stop-50", Kind = StepKind.Stop,
                Text = "Stop here and you've got A, an origin if you wanted one, and the midpoint belief — real, usable, enough to carry forward." },
            BeliefCut("cut25", 25, "Plot point 1", "Belief at 25%."),
            new Step { Key = ShapeAnswers.AliveKey(25), Kind = StepKind.AliveFlat, Label = "Alive?", Question = "Alive or flat?" },
            BeliefCut("cut75", 75, "Crisis / Plot point 2", "Belief at 75%."),
            new Step { Key = ShapeAnswers.AliveKey(75), Kind = StepKind.AliveFlat, Label = "Alive?", Question = "Alive or flat?" },
            BeliefCut("cut12", 12, "Inciting incident", "Belief at 12%."),
            new Step { Key = ShapeAnswers.AliveKey(12), Kind = StepKind.AliveFlat, Label = "Alive?", Question = "Alive or flat?" },
            BeliefCut("cut88", 88, "Last obstacle", "Belief at 88%."),
            new Step { Key = ShapeAnswers.AliveKey(88), Kind = StepKind.AliveFlat, Label = "Alive?", Question = "Alive or flat?",
                Hint = "Mark alive or flat — don't fix flat ones now, just flag them." },
            new Step { Key = "stop-cuts", Kind = StepKind.Stop,
                Text = "Belief Split only defines these five named cuts — the two pinches take the lighter third-rail question instead, not a sixth or seventh. Manufacturing more precision than that isn't the point." },
            new Step { Key = "stop-cast", Kind = StepKind.Stop,
                Text = "Optional — the Cast card, for everyone who isn't the protagonist: the opponent and each main supporting character, once per story. Short answers are fine." },
            CastStep("who", "Who", "Who is this?", "The opponent, or a main supporting character.") with { WhoKey = null },
            CastStep("now", "Wants now", "What do they want right now?"),
            CastStep("life", "Wants from life", "What do they want from life?"),
            CastStep("secret", "Secret", "A secret."),
            CastStep("resolve", "Resolving", "What are they trying to resolve before the story ends?", "Stands in for a flaw if they don't change."),
            CastStep("trait", "Against type", "One trait against type."),
            CastStep("risk", "Risk", "Opponent only: why would they risk everything for this?",
                why: "If the opponent can't answer this and the next one, they're a strawman. The fix is a stronger reason or a real edge, never a nastier personality."),
            CastStep("edge", "Edge", "Opponent only: what edge do they have over the protagonist?"),
            CastStep("watch", "Why we'd watch", "Why would we keep watching them?", "Not: would we like them."),
            new Step { Key = "cast-loop", Kind = StepKind.Loop, LoopGroup = "cast" }
        }
    };

    public static readonly Exercise Extend = new()
    {
        Id = SlateMigration.ExtendId,
        Name = "Extend Backward",
        Tagline = "Work back from the ending",
        Shelf = Shelf.Ending,
        Anchor = AnchorKind.Shape,
        SceneToo = true,
        AnchorLabel = "The Shape's end, or a scene",
        Targets = new[] { TargetKind.InboxBeat, TargetKind.Copy },
        Why = "For when you have an ending — yours, or one you're stealing to reverse-engineer — and nothing that leads to it yet. "
            + "Answer, then ask the same question of your own answer. No jumping ahead to fill the middle from the front — only the "
            + "most recent link you wrote generates the next one. Keep going until the chain reaches something that can plausibly be "
            + "a beginning, or until you're done for now. Either is a valid place to stop.",
        Steps = new[]
        {
            new Step { Key = ShapeAnswers.Z, Label = "Z", ScriptSlot = "Ends",
                Question = "Z — the one thing you already have: a final image, a last line, an ending." },
            Gate("obit", "Is Z more a feeling than a picture? Try an Obituary Z.", "Optional.", "Obituary Z"),
            IfYes("obit", "obit"),
            new Step { Key = "eulogy", Group = "obit", Label = "Eulogy Z", Multiline = true,
                Question = "Write a eulogy for a character who's already died by the end, spoken by someone else in the story.",
                Hint = "What do they say happened? What do they conspicuously not say?",
                Why = "Use that as Z instead — richer, and already implies things left unsaid." },
            new Step { Key = "link", Group = "link", Label = "Link", Offer = true, Echo = true,
                EndKeys = new[] { "eulogy", ShapeAnswers.Z },
                Question = "What had to be true right before this, for this to happen?",
                AltQuestion = "Treat it as a crime scene: what would a detective need to find to explain it?",
                AltHint = "Not what caused it — what would be found.", AltLabel = "Mystery lens on this link",
                Why = "Stuck partway? Try the mystery lens on the next link only. Different question, sometimes unsticks the same chain." },
            new Step { Key = "link-loop", Kind = StepKind.Loop, LoopGroup = "link" },
            new Step { Key = "chain", Kind = StepKind.AliveFlat, Label = "Read it forward", ReadForwardKey = "link",
                EndKeys = new[] { "eulogy", ShapeAnswers.Z },
                Question = "Read the chain forward, start to Z. Does each link earn the next, or did one get invented to patch a gap?" },
            new Step { Key = "weak", Label = "Weakest link", Question = "If flat, name the weakest link." }
        }
    };

    // ---- Bridge and Position (T6): the worksheet's Path B and Path A ----

    public static readonly Exercise Bridge = new()
    {
        Id = SlateMigration.BridgeId,
        Name = "Bridge",
        Tagline = "Connect two moments you know",
        Shelf = Shelf.Nothing,
        Anchor = AnchorKind.TwoPoints,
        AnchorLabel = "Which two cards?",
        Targets = new[] { TargetKind.InboxBeat, TargetKind.SceneCard, TargetKind.Copy },
        Why = "Two fixed points — a planted detail and its payoff, an opening image and the climax, two characters who haven't "
            + "met yet but will — and nothing between them. Bridge fixes both ends before anything else; that's the whole trick. "
            + "Then rounds of three candidate links: pick the one that works, or half-pick one that's close but not quite right. "
            + "Move to the next round only if the gap still needs more than one link. Stop after any round — a picked or "
            + "half-picked link is a complete session. A gap that needs more than three or four links is probably a whole "
            + "story's gap, not a scene-to-scene one: that's the slower card version, which is Treatment-stage work and isn't "
            + "built here. A→Z→Split is the tool for a whole story's shape.",
        Steps = new[]
        {
            new Step { Key = "a", Label = "A", Question = "A — the point you're starting from.",
                Hint = "Don't skip this — the whole tool depends on it." },
            new Step { Key = "z", Label = "Z", Question = "Z — the point you're building to." },
            new Step { Key = "links", Kind = StepKind.Lines, Count = 3, Group = "round", Label = "Candidates", Offer = true,
                Question = "Three candidate links between A and Z.",
                Hint = "About ten seconds each if you time it." },
            new Step { Key = "pick", Kind = StepKind.Pick, FromKey = "links", Group = "round", Label = "Link",
                Question = "Which one works — or is close but not quite right?",
                Hint = "Tap one, or write your own. A half-pick counts." },
            new Step { Key = "loop", Kind = StepKind.Loop, LoopGroup = "round" },
            new Step { Key = "wire", Label = "The wire", Multiline = true, ReadForwardKey = "pick", LeadKey = "a", EndKeys = new[] { "z" },
                Question = "The wire, read start to finish.",
                Hint = "A, the links you picked, Z — as one run.",
                Why = "If it's a whole story's gap rather than a scene-to-scene one, this burst version isn't the right tool — "
                    + "that's the slower card version, which is Treatment-stage work and isn't built here. A→Z→Split is the tool "
                    + "for a whole story's shape." }
        }
    };

    public static readonly Exercise Position = new()
    {
        Id = SlateMigration.PositionId,
        Name = "Position",
        Tagline = "Try a scene in each story slot",
        Shelf = Shelf.Moment,
        Anchor = AnchorKind.InboxBeat,
        AnchorLabel = "Which beat or scene?",
        Targets = new[] { TargetKind.Board, TargetKind.Copy },
        Why = "For a scene, image or line you already have and already love, but don't know where it goes. Run it through the "
            + "story's seven turns, as few or as many as you want: if this sat here, what does it need before it, and what does "
            + "it force after it? An answer that comes easily and feels like it belongs is alive; one you have to force, or "
            + "nothing at all, is flat. Flat is information — wrong slot, not a wasted turn — and more than one alive means "
            + "you've found candidates, not committed to an answer. Carry the strongest one or two into Split or Fill and let "
            + "the surrounding structure make the final call.",
        Steps = new[]
        {
            new Step { Key = "moment", Label = "Moment", Multiline = true,
                Question = "The moment — the scene, image or line you already have and love, with nowhere to put it yet." },
            new Step { Key = "before", Kind = StepKind.Slots, Source = TurnsSource, Label = "Before",
                Question = "If this moment sat right here, what would it need to be true right before it?",
                Hint = "Fill in the turns you want to try; leave the rest.",
                Why = "If an answer comes easily and feels like it belongs, that's alive. If you have to force it, or nothing comes, "
                    + "that's flat — not a failure, just proof this isn't the slot." },
            new Step { Key = "after", Kind = StepKind.Slots, Source = TurnsSource, Label = "After",
                Question = "…and what would it force to happen right after?" },
            new Step { Key = "alive", Kind = StepKind.Slots, Source = TurnsSource, Pills = true, Label = "Alive or flat",
                Question = "Alive or flat, in each turn you tried?",
                Hint = "Flat is information — wrong slot. More than one alive isn't a problem." },
            new Step { Key = "ladder-if", Kind = StepKind.Branch, BranchGroup = "ladder", OnKey = "alive", OnValue = "flat", AllSlots = true },
            new Step { Key = "at", Group = "ladder", Label = "At it", Multiline = true,
                Question = "Flat everywhere? Write the moment again, as it happens.",
                Hint = "Time-Skip Ladder: three times — as it happens, a year on, a minute on.",
                Why = "A moment that's actually structural earns its keep in all three; one that only ever worked as mood usually "
                    + "can't survive the jump either direction. It cuts a moment down to what's structural about it versus what's "
                    + "just atmosphere — sometimes the moment was never flat, it was just timed wrong." },
            new Step { Key = "year", Group = "ladder", Label = "One year later", Multiline = true,
                Question = "One year later — has anything about it survived?" },
            new Step { Key = "minute", Group = "ladder", Label = "One minute later", Multiline = true,
                Question = "One minute later — what's the very next domino it knocks over?" },
            new Step { Key = "structural", Kind = StepKind.YesNo, Group = "ladder", Label = "Still structural?",
                Question = "Still structural?",
                Hint = "Yes: it wasn't flat, it was timed wrong — go back and try the remaining slots. No: it may be a mood beat, and that's fine to know." },
            new Step { Key = "shortlist", Kind = StepKind.Lines, Count = 2, Label = "Shortlist",
                Question = "Shortlist the strongest one or two alive answers.",
                Hint = "Split or Fill picks the winner another day, not this." },
            new Step { Key = PositionTurnKey, Kind = StepKind.Pick, FromKey = "alive", Label = "The turn",
                Question = "Which turn does it go in?",
                Hint = "Optional — it's what Send to the Board files the beat under." }
        }
    };

    public static readonly IReadOnlyList<Exercise> All = new[] { Woac, Flaw, Fill, SceneDials, Split, Belief, Extend, Bridge, Position };

    public static Exercise? Find(string id) => All.FirstOrDefault(exercise => exercise.Id == id);
}
