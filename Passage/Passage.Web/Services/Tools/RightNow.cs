namespace Passage.Web.Services.Tools;

/// <summary>One "right now" line on the launcher: tapping it starts a run with the anchor filled in.</summary>
public sealed record RightNowSuggestion(string ExerciseId, string Title, string Detail, Anchor Anchor);

/// <summary>What <see cref="RightNow.Suggest"/> looks at, gathered by the host from the parse and the sidecar.</summary>
public sealed record RightNowInput
{
    /// <summary>The first live bracket on the caret's line, if any.</summary>
    public string? BracketUnderCaret { get; init; }

    public SceneChoice? SceneUnderCaret { get; init; }

    /// <summary>Who speaks in that scene, most speeches first.</summary>
    public IReadOnlyList<string> SceneSpeakers { get; init; } = Array.Empty<string>();

    /// <summary>The newest beat sitting in the inbox (Page 0).</summary>
    public Beat? NewestInboxBeat { get; init; }

    /// <summary>No Shape written and no "#" or "##" headings in the script.</summary>
    public bool NoStructure { get; init; }

    /// <summary>Cast members with no note page, most speeches first.</summary>
    public IReadOnlyList<CharacterChoice> CastWithoutNote { get; init; } = Array.Empty<CharacterChoice>();

    /// <summary>Bumped each time the dock opens so the scene suggestion rotates through its candidates.</summary>
    public int Rotation { get; init; }
}

/// <summary>
/// The launcher's "Right now" row (docs/SLATE-REDO-PLAN.md §4.1): at most three suggestions from where the caret is and
/// what the project holds. Pure. Only exercises in <c>available</c> are offered, so a rule whose exercise is not ported
/// yet stays quiet instead of leading to a card that does nothing.
/// </summary>
public static class RightNow
{
    public const int Max = 3;

    // The ids the later rows will add; kept here so a rule never has to change when its exercise lands.
    public const string FillId = "fill";
    public const string PushPullId = "pushpull";
    public const string SceneDialsId = "scenedials";
    public const string UnderstudyId = "understudy";
    public const string PositionId = "position";
    public const string SplitId = "split";

    public static List<RightNowSuggestion> Suggest(RightNowInput input, IReadOnlySet<string> available)
    {
        var found = new List<RightNowSuggestion>();

        if (input.BracketUnderCaret is { Length: > 0 } bracket && available.Contains(FillId))
        {
            found.Add(new RightNowSuggestion(FillId, "Fill", bracket, new Anchor { Kind = AnchorKind.Bracket, Text = bracket }));
        }

        if (input.SceneUnderCaret is { } scene)
        {
            // One of them, rotating on each open (§4.1). WOAC is the one that exists before the rest are ported.
            var candidates = new[] { PushPullId, SceneDialsId, UnderstudyId, ExerciseCatalog.Woac.Id }
                .Where(available.Contains).ToList();
            if (candidates.Count > 0)
            {
                var id = candidates[Math.Abs(input.Rotation) % candidates.Count];
                var anchor = new Anchor { Kind = AnchorKind.Scene, TargetId = scene.Id, Text = scene.Heading, SceneId = scene.Id, SceneText = scene.Heading };
                if (id == ExerciseCatalog.Woac.Id)
                {
                    // WOAC is about a character; the scene's busiest speaker is the default (changeable in the strip).
                    anchor = new Anchor { Kind = AnchorKind.Character, Text = input.SceneSpeakers.FirstOrDefault() ?? string.Empty, SceneId = scene.Id, SceneText = scene.Heading };
                }

                found.Add(new RightNowSuggestion(id, NameOf(id), scene.Heading, anchor));
            }
        }

        if (input.NewestInboxBeat is { } beat && available.Contains(PositionId))
        {
            var title = beat.Title.Trim().Length > 0 ? beat.Title.Trim() : "Untitled beat";
            found.Add(new RightNowSuggestion(PositionId, "Position", $"“{title}”", new Anchor { Kind = AnchorKind.InboxBeat, TargetId = beat.Id, Text = title }));
        }

        if (input.NoStructure && available.Contains(SplitId))
        {
            found.Add(new RightNowSuggestion(SplitId, "A→Z→Split", "Find the shape of the story", new Anchor { Kind = AnchorKind.Shape }));
        }

        if (input.CastWithoutNote.FirstOrDefault() is { } character && available.Contains(ExerciseCatalog.Flaw.Id))
        {
            found.Add(new RightNowSuggestion(ExerciseCatalog.Flaw.Id, "Flaw Brainstorm", character.Name,
                new Anchor { Kind = AnchorKind.Character, Text = character.Name }));
        }

        return found.Take(Max).ToList();
    }

    private static string NameOf(string id) => ExerciseCatalog.Find(id)?.Name ?? id switch
    {
        PushPullId => "Push/Pull",
        SceneDialsId => "Scene Dials",
        UnderstudyId => "Understudy",
        _ => id
    };
}
