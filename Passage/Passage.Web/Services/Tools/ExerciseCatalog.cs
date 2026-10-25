namespace Passage.Web.Services.Tools;

/// <summary>
/// The deck (docs/SLATE-REDO-PLAN.md §4.5): one static definition per
/// exercise. Questions keep the vault worksheets' words (R-B); anything
/// longer sits in Why. Row T2 defines WOAC and Flaw, T4 Fill and Scene Dials;
/// later rows add the rest.
/// </summary>
public static class ExerciseCatalog
{
    /// <summary>The Slots source that lists the people who speak in the run's scene.</summary>
    public const string SpeakersSource = "speakers";

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

    public static readonly IReadOnlyList<Exercise> All = new[] { Woac, Flaw, Fill, SceneDials };

    public static Exercise? Find(string id) => All.FirstOrDefault(exercise => exercise.Id == id);
}
