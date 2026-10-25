using System.Text.Json.Serialization;

namespace Passage.Web.Services.Tools;

/// <summary>
/// What a run is about. Never a line number: a Board/beat Guid where one
/// exists, otherwise text matched at use time (R-E).
/// </summary>
public sealed class Anchor
{
    public AnchorKind Kind { get; set; }
    public Guid? TargetId { get; set; }
    public string Text { get; set; } = string.Empty;

    // The scene a character-anchored run is about, when the writer named one.
    // By id where the heading has a marker, otherwise by heading text.
    public Guid? SceneId { get; set; }
    public string SceneText { get; set; } = string.Empty;
}

/// <summary>An answer that was sent somewhere, so the trail can tag it.</summary>
public sealed class SentMark
{
    public string Key { get; set; } = string.Empty;
    public TargetKind Target { get; set; }
}

/// <summary>
/// One run of an exercise. Generic: answers keyed by step. A loop round r&gt;0
/// stores under "key#r", a Lines slot under "key[i]", a Slots item under
/// "key/item". Kept by what it is about, never by when it was run (rule 0.3),
/// so there is deliberately no timestamp.
/// </summary>
public sealed class ExerciseRun
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string ExerciseId { get; set; } = string.Empty;
    public Anchor Anchor { get; set; } = new();
    public Dictionary<string, string> Answers { get; set; } = new();

    // Loop rounds shown so far (at least one). One loop per exercise.
    public int Rounds { get; set; } = 1;

    // Steps skipped and stop points continued past, by step id ("key" or
    // "key#r"). Skipping stores nothing about the step's answer.
    public List<string> Passed { get; set; } = new();

    public List<SentMark> Sent { get; set; } = new();

    // A bracket or scene is handed over rather than typed, so a run that only
    // holds one is as empty as a blank page: handing out brackets must not
    // leave a kept run behind for each.
    [JsonIgnore]
    public bool IsEmpty => Answers.Values.All(string.IsNullOrWhiteSpace)
        && (Anchor.Kind is AnchorKind.Bracket or AnchorKind.Scene || string.IsNullOrWhiteSpace(Anchor.Text));
}
