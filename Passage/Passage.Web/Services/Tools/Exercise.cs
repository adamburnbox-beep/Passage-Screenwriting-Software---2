using System.Text.Json.Serialization;

namespace Passage.Web.Services.Tools;

/// <summary>Where an exercise sits on the launcher's "what have you got?" chips.</summary>
public enum Shelf { Nothing, Ending, Moment, SceneGoing, DeadScene }

/// <summary>What an exercise is about (docs/SLATE-REDO-PLAN.md §4.4).</summary>
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum AnchorKind { None, Character, Scene, Bracket, InboxBeat, TwoPoints, Shape }

/// <summary>Where an answer can be sent (§4.4).</summary>
public enum TargetKind { Bracket, Synopsis, InboxBeat, SceneCard, NotePage, Shape, Copy }

/// <summary>The whole step vocabulary (§4.3). New behaviour means a new kind with a test.</summary>
public enum StepKind
{
    Text,
    Lines,
    Pick,
    YesNo,
    AliveFlat,
    Branch,
    Dice,
    Slots,
    Stop,
    Loop
}

/// <summary>
/// One step of an exercise. A flat record with the fields each kind reads:
/// the engine, not the step, decides what is current, visible or carried.
/// </summary>
public sealed record Step
{
    // Unique within the exercise. Answers are stored under it.
    public required string Key { get; init; }
    public StepKind Kind { get; init; } = StepKind.Text;

    // The worksheet's own words, a muted hint, and the folded "Why?" (R-B).
    public string Question { get; init; } = string.Empty;
    // The short name the trail shows for an answered step ("Want", "Q1").
    public string Label { get; init; } = string.Empty;
    public string Hint { get; init; } = string.Empty;
    public string Why { get; init; } = string.Empty;
    public bool Multiline { get; init; }

    // Branch and Loop groups: a step belongs to at most one group. A group's
    // steps appear only through the Branch or Loop that names it.
    public string? Group { get; init; }

    // Lines: how many slots; Open lets the writer add another.
    public int Count { get; init; } = 1;
    public bool Open { get; init; }

    // Pick: the Lines step whose slots become the options.
    public string? FromKey { get; init; }

    // Branch: show BranchGroup only while the step OnKey (same round) holds
    // OnValue. With AllSlots, OnKey names a Slots step and every one of its
    // items must hold OnValue ("flat everywhere").
    public string? BranchGroup { get; init; }
    public string? OnKey { get; init; }
    public string? OnValue { get; init; }
    public bool AllSlots { get; init; }

    // Dice: the table the roll is drawn from.
    public IReadOnlyList<string> Table { get; init; } = Array.Empty<string>();

    // Slots: the name the engine's slot source is asked for (the seven turns,
    // the speakers in a scene).
    public string? Source { get; init; }

    // Stop: the stop-callout's words, shown with Stop here / Keep going.
    public string Text { get; init; } = string.Empty;

    // Where an answer to this step can be put in the script, if anywhere.
    public TargetKind? Send { get; init; }

    // Partner seam (§4.2): a site where PartnerLines may render.
    public bool Offer { get; init; }

    // Loop: repeat LoopGroup. Each later round's CarryTo step shows the
    // previous round's CarryFrom answer, read live and never stored.
    public string? LoopGroup { get; init; }
    public string? CarryFrom { get; init; }
    public string? CarryTo { get; init; }
}

/// <summary>An exercise definition: data, rendered by one runner (R-C).</summary>
public sealed record Exercise
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public string Tagline { get; init; } = string.Empty;
    // The long description, folded behind the intro's "Why?" (R-B).
    public string Why { get; init; } = string.Empty;
    public Shelf Shelf { get; init; }
    public AnchorKind Anchor { get; init; }
    // What the free-text anchor asks for, until T3 adds real pickers.
    public string AnchorLabel { get; init; } = string.Empty;
    public IReadOnlyList<TargetKind> Targets { get; init; } = Array.Empty<TargetKind>();
    public required IReadOnlyList<Step> Steps { get; init; }
}
