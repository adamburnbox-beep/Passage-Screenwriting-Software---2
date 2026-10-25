namespace Passage.Web.Services.Tools;

/// <summary>
/// Starts the answers an anchor can supply: the moment a Position run is about, the two ends of a Bridge. Only blank answers
/// are filled, so naming a beat or a card never overwrites what the writer has written. Pure over <see cref="IToolHost"/>'s lists.
/// </summary>
public static class AnchorPrefill
{
    /// <summary>True when an answer was filled.</summary>
    public static bool Apply(Exercise exercise, ExerciseRun run, IToolHost host)
    {
        switch (exercise.Id)
        {
            case SlateMigration.BridgeId:
                var start = host.ResolveScene(run.Anchor);
                var end = ProjectContext.ResolveEnd(host.Scenes, run.Anchor);
                return Fill(run, "a", start is null ? null : Describe(start))
                    | Fill(run, "z", end is null ? null : Describe(end));
            case SlateMigration.PositionId:
                return Fill(run, "moment", MomentOf(run.Anchor, host));
            default:
                return false;
        }
    }

    private static string? MomentOf(Anchor anchor, IToolHost host)
    {
        if (anchor.Kind == AnchorKind.InboxBeat)
        {
            var beat = host.Slate.Beats.FirstOrDefault(entry => entry.Id == anchor.TargetId);
            if (beat is null)
            {
                return null;
            }

            var title = beat.Title.Trim();
            var description = beat.Description.Trim();
            return title.Length > 0 && description.Length > 0 ? title + " — " + description : title.Length > 0 ? title : description;
        }

        return anchor.Kind == AnchorKind.Scene && host.ResolveScene(anchor) is { } scene ? Describe(scene) : null;
    }

    /// <summary>A scene as a point in a story: its heading, then its synopsis if it has one.</summary>
    public static string Describe(SceneChoice scene) =>
        scene.Synopsis.Length > 0 ? $"{scene.Heading} — {scene.Synopsis}" : scene.Heading;

    private static bool Fill(ExerciseRun run, string key, string? text)
    {
        if (string.IsNullOrWhiteSpace(text) || run.Answers.TryGetValue(key, out var existing) && !string.IsNullOrWhiteSpace(existing))
        {
            return false;
        }

        run.Answers[key] = text;
        return true;
    }
}
