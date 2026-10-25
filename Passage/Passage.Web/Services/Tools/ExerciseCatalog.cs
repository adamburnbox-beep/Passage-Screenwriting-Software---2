namespace Passage.Web.Services.Tools;

/// <summary>
/// The deck (docs/SLATE-REDO-PLAN.md §4.5): one static definition per
/// exercise. Questions keep the vault worksheets' words (R-B); anything
/// longer sits in Why. Row T2 defines WOAC and Flaw; later rows add the rest.
/// </summary>
public static class ExerciseCatalog
{
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

    public static readonly IReadOnlyList<Exercise> All = new[] { Woac, Flaw };

    public static Exercise? Find(string id) => All.FirstOrDefault(exercise => exercise.Id == id);
}
