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

    /// <summary>The anchor's scene as it is in the script now; null when it is gone.</summary>
    SceneChoice? ResolveScene(Anchor anchor);

    /// <summary>The up-to-three suggestions for the launcher's Right now row.</summary>
    IReadOnlyList<RightNowSuggestion> RightNow();

    /// <summary>Puts an answer where the writer sent it. The host does the write (ranged, so undo
    /// holds) and reports what it did or why it could not on the status line. True when it landed.</summary>
    Task<bool> SendAsync(TargetKind target, ToolSend send);

    /// <summary>Where a synopsis would land for this anchor, for the Send to menu's label.</summary>
    string SynopsisTargetLabel(Anchor anchor);

    /// <summary>Scrolls the script to the anchor without taking focus from the runner.</summary>
    Task JumpToAsync(Anchor anchor);

    /// <summary>A burst timer is running (it lives in passage.js).</summary>
    bool BurstRunning { get; }

    Task StartBurstAsync(string rootId);

    Task StopBurstAsync();
}

/// <summary>One thing being sent: the text, what it answers, and the run it came from.</summary>
public sealed record ToolSend(Anchor Anchor, string Exercise, string Label, string Text);
