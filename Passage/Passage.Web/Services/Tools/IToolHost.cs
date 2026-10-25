namespace Passage.Web.Services.Tools;

/// <summary>
/// What the tools dock needs from the page that hosts it (docs/SLATE-REDO-PLAN.md
/// R-D). The runner never touches CodeMirror, the script text or the analysis
/// itself; every script write and status line goes through here, so the traps
/// that guard them (ranged writes only, mirror and re-analyse, fail visibly)
/// live in one place.
/// </summary>
public interface IToolHost
{
    /// <summary>The sidecar the dock edits. Replaced when a different script loads.</summary>
    SlateDocument Slate { get; }

    /// <summary>A named script is open, so the working-out is kept.</summary>
    bool HasScript { get; }

    /// <summary>Persists <see cref="Slate"/> after a change.</summary>
    void Save();

    /// <summary>Shows a line in the status bar.</summary>
    void SetStatus(string message);

    /// <summary>The Cast, most speeches first, with highlight colours (the character picker).</summary>
    IReadOnlyList<CharacterChoice> Cast { get; }

    /// <summary>Every scene heading in the script (the scene picker).</summary>
    IReadOnlyList<SceneChoice> Scenes { get; }

    /// <summary>The scene the caret is in, if any.</summary>
    SceneChoice? SceneUnderCaret { get; }

    /// <summary>Who speaks in the scene, most speeches first.</summary>
    IReadOnlyList<string> SpeakersOf(SceneChoice scene);

    /// <summary>Who speaks in the scene the anchor is about (the scene itself, the one a character is
    /// named in, or the one a bracket sits in), most speeches first. Empty when there is none.</summary>
    IReadOnlyList<string> SpeakersFor(Anchor anchor);

    /// <summary>The speaker with the fewest speeches in the scene the anchor is about (Understudy's default). Null when
    /// the scene is gone or nobody speaks in it.</summary>
    string? LeastSpeakingFor(Anchor anchor);

    /// <summary>Opens another exercise on the same scene (GSU's flag points at Scene Dials): the run already kept for
    /// that scene, or a new one.</summary>
    Task OpenOnSceneAsync(string exerciseId, Anchor anchor);

    /// <summary>Every bracket in the script, ranked, with the ones the writer left out marked.</summary>
    IReadOnlyList<BracketChoice> Brackets { get; }

    /// <summary>The bracket "hand me one" offers now: the top of the queue at the cursor. Null when none is left.</summary>
    Anchor? HandBracket();

    /// <summary>Opens a run on this bracket (the one already kept for it, or a new one).</summary>
    Task OpenBracketAsync(string text);

    /// <summary>Puts the run's bracket back unfilled and hands on the next. Closes to the launcher when none is left.</summary>
    Task NextBracketAsync();

    /// <summary>"Not a bracket": adds it to the ignore list, drops its run, and hands on the next.</summary>
    Task NotABracketAsync(string text);

    /// <summary>Takes a bracket off the ignore list.</summary>
    void RestoreBracket(string text);

    /// <summary>The anchor's scene as it is in the script now; null when it is gone.</summary>
    SceneChoice? ResolveScene(Anchor anchor);

    /// <summary>The up-to-three suggestions for the launcher's Right now row.</summary>
    IReadOnlyList<RightNowSuggestion> RightNow();

    /// <summary>Puts an answer where the writer sent it. The host does the write (ranged, so undo
    /// holds) and reports what it did or why it could not on the status line. True when it landed.</summary>
    Task<bool> SendAsync(TargetKind target, ToolSend send);

    /// <summary>Where a synopsis would land for this anchor, for the Send to menu's label.</summary>
    string SynopsisTargetLabel(Anchor anchor);

    /// <summary>The script already holds a Shape (a sequence with a Starts / Midpoint / … line).</summary>
    bool ScriptHasShape { get; }

    /// <summary>What the script holds for a Shape line (a <see cref="SplitScript.Slot"/> label) when it is real
    /// text: not empty, not a bracket, not a dropped turn. Null otherwise.</summary>
    string? ScriptShapeValue(string label);

    /// <summary>Writes a Split run into the script: the whole structure the first time (a ranged insert), then
    /// only the turns still held as brackets. Reports what it did on the status line. True when the script changed.</summary>
    Task<bool> WriteShapeAsync(SplitRun run);

    /// <summary>Scrolls the script to the anchor without taking focus from the runner.</summary>
    Task JumpToAsync(Anchor anchor);

    /// <summary>A burst timer is running (it lives in passage.js).</summary>
    bool BurstRunning { get; }

    /// <summary>Starts the timer over the inputs under <paramref name="rootId"/>. <paramref name="seconds"/> of 0 uses the
    /// per-answer setting; the Lens and the freewrite pass their own, in minutes.</summary>
    Task StartBurstAsync(string rootId, int seconds = 0);

    Task StopBurstAsync();
}

/// <summary>One thing being sent: the text, what it answers, and the run it came from.</summary>
public sealed record ToolSend(Anchor Anchor, string Exercise, string Label, string Text);
