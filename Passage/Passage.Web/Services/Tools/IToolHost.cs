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

    /// <summary>Adds a "=" synopsis line under the heading the caret is in. Reports its own status.</summary>
    Task PutSynopsisAsync(string text);

    /// <summary>Where a synopsis would land, for the end panel's label.</summary>
    string SynopsisTargetLabel { get; }

    /// <summary>A burst timer is running (it lives in passage.js).</summary>
    bool BurstRunning { get; }

    Task StartBurstAsync(string rootId);

    Task StopBurstAsync();
}
