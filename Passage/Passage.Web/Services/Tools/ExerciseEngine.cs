namespace Passage.Web.Services.Tools;

/// <summary>One step as it appears in a run: the step and which loop round.</summary>
public readonly record struct StepSlot(Step Step, int Round)
{
    // The id used by Passed; the answer key for single-value steps.
    public string Id => ExerciseEngine.Suffix(Step.Key, Round);
}

/// <summary>A line of the trail: an answered step, one line, no number.</summary>
public sealed record TrailEntry(StepSlot Slot, string Summary);

/// <summary>
/// The pure rules of a run (docs/SLATE-REDO-PLAN.md §4.3): which steps are
/// visible, which is current, what a loop's carry shows, what a branch hides.
/// Exposes nothing that reads as progress (rule 0.2): no counts, no fraction.
/// </summary>
public sealed class ExerciseEngine
{
    private readonly Exercise _exercise;
    private readonly Func<string, IReadOnlyList<string>> _slotSource;

    public ExerciseEngine(Exercise exercise, Func<string, IReadOnlyList<string>>? slotSource = null)
    {
        _exercise = exercise;
        _slotSource = slotSource ?? (_ => Array.Empty<string>());
    }

    internal static string Suffix(string key, int round) => round == 0 ? key : $"{key}#{round}";

    private static string LineKey(string key, int round, int index) => $"{Suffix(key, round)}[{index}]";

    private static string ItemKey(string key, int round, string item) => $"{Suffix(key, round)}/{item}";

    private Step? Find(string key) => _exercise.Steps.FirstOrDefault(step => step.Key == key);

    private static bool Blank(string? value) => string.IsNullOrWhiteSpace(value);

    private static string Read(ExerciseRun run, string key) =>
        run.Answers.TryGetValue(key, out var value) ? value : string.Empty;

    /// <summary>
    /// Every step the writer can see, in order: loop groups expanded to the
    /// run's rounds, branch groups present only while their branch holds,
    /// Branch and Loop markers themselves never listed.
    /// </summary>
    public IReadOnlyList<StepSlot> Sequence(ExerciseRun run)
    {
        var slots = new List<StepSlot>();
        foreach (var step in _exercise.Steps)
        {
            if (step.Kind == StepKind.Branch)
            {
                continue;
            }

            if (step.Kind == StepKind.Loop)
            {
                for (var round = 0; round < Math.Max(1, run.Rounds); round++)
                {
                    foreach (var inner in _exercise.Steps.Where(s => s.Group == step.LoopGroup))
                    {
                        slots.Add(new StepSlot(inner, round));
                    }
                }

                continue;
            }

            if (step.Group is null)
            {
                slots.Add(new StepSlot(step, 0));
                continue;
            }

            // A step in a branch group shows only while its branch holds. A
            // step in a loop group is emitted by the Loop above.
            if (_exercise.Steps.Any(s => s.Kind == StepKind.Loop && s.LoopGroup == step.Group))
            {
                continue;
            }

            var branch = _exercise.Steps.FirstOrDefault(s => s.Kind == StepKind.Branch && s.BranchGroup == step.Group);
            if (branch is not null && BranchHolds(run, branch))
            {
                slots.Add(new StepSlot(step, 0));
            }
        }

        return slots;
    }

    private bool BranchHolds(ExerciseRun run, Step branch)
    {
        if (branch.OnKey is null || branch.OnValue is null)
        {
            return false;
        }

        if (!branch.AllSlots)
        {
            return string.Equals(Read(run, branch.OnKey), branch.OnValue, StringComparison.Ordinal);
        }

        var on = Find(branch.OnKey);
        var items = on?.Source is { } source ? _slotSource(source) : Array.Empty<string>();
        return items.Count > 0
            && items.All(item => string.Equals(Read(run, ItemKey(branch.OnKey, 0, item)), branch.OnValue, StringComparison.Ordinal));
    }

    private (Step Loop, Step Member)? CarryTarget(StepSlot slot)
    {
        if (slot.Round == 0 || slot.Step.Group is null)
        {
            return null;
        }

        var loop = _exercise.Steps.FirstOrDefault(s => s.Kind == StepKind.Loop && s.LoopGroup == slot.Step.Group);
        return loop is not null && loop.CarryTo == slot.Step.Key ? (loop, slot.Step) : null;
    }

    /// <summary>True when the slot shows a carried value: read-only, edited at its source.</summary>
    public bool IsCarried(StepSlot slot) => CarryTarget(slot) is not null;

    /// <summary>
    /// The value a single-value step shows. A loop's carry target in a later
    /// round shows the previous round's source answer, read at call time.
    /// </summary>
    public string GetAnswer(ExerciseRun run, string key, int round = 0)
    {
        var step = Find(key);
        if (step is not null && CarryTarget(new StepSlot(step, round)) is { } carry && carry.Loop.CarryFrom is { } from)
        {
            return Read(run, Suffix(from, round - 1));
        }

        return Read(run, Suffix(key, round));
    }

    /// <summary>
    /// Stores an answer. Returns false for a carried value, which is never
    /// stored. Answering a loop group's last step in the last round appends a
    /// round (the next one starts from the carry).
    /// </summary>
    public bool SetAnswer(ExerciseRun run, string key, int round, string value)
    {
        var step = Find(key);
        if (step is null || CarryTarget(new StepSlot(step, round)) is not null)
        {
            return false;
        }

        run.Answers[Suffix(key, round)] = value;
        run.Passed.Remove(Suffix(key, round));
        AppendRoundIfLooped(run, step, round, value);
        return true;
    }

    public void SetLine(ExerciseRun run, string key, int round, int index, string value)
    {
        run.Answers[LineKey(key, round, index)] = value;
        run.Passed.Remove(Suffix(key, round));
        if (Find(key) is { } step)
        {
            AppendRoundIfLooped(run, step, round, value);
        }
    }

    public string GetLine(ExerciseRun run, string key, int round, int index) =>
        Read(run, LineKey(key, round, index));

    public void SetItem(ExerciseRun run, string key, int round, string item, string value)
    {
        run.Answers[ItemKey(key, round, item)] = value;
        run.Passed.Remove(Suffix(key, round));
    }

    public string GetItem(ExerciseRun run, string key, int round, string item) =>
        Read(run, ItemKey(key, round, item));

    private void AppendRoundIfLooped(ExerciseRun run, Step step, int round, string value)
    {
        if (Blank(value) || step.Group is null || round != Math.Max(1, run.Rounds) - 1)
        {
            return;
        }

        var group = _exercise.Steps.Where(s => s.Group == step.Group).ToList();
        if (_exercise.Steps.Any(s => s.Kind == StepKind.Loop && s.LoopGroup == step.Group)
            && group.Count > 0 && group[^1].Key == step.Key)
        {
            run.Rounds = round + 2;
        }
    }

    /// <summary>
    /// The options a Pick offers: the non-blank slots of its Lines step in the same round, or
    /// its fixed list.
    /// </summary>
    public IReadOnlyList<string> PickOptions(ExerciseRun run, StepSlot slot)
    {
        var from = slot.Step.FromKey is { } key ? Find(key) : null;
        if (from is null)
        {
            return slot.Step.Options;
        }

        var options = new List<string>();
        for (var i = 0; i < SlotCount(run, from, slot.Round); i++)
        {
            var text = GetLine(run, from.Key, slot.Round, i);
            if (!Blank(text))
            {
                options.Add(text);
            }
        }

        return options;
    }

    // A fixed option is "Name — gloss" and stores its Name; a line from a Lines step stores as written.
    private static string PickValue(Step step, string option)
    {
        if (step.FromKey is not null)
        {
            return option;
        }

        var at = option.IndexOf(" — ", StringComparison.Ordinal);
        return at < 0 ? option : option[..at];
    }

    /// <summary>Taps an option of a Pick: its text becomes the answer (the writer can still edit it).</summary>
    public void Pick(ExerciseRun run, StepSlot slot, string option) =>
        SetAnswer(run, slot.Step.Key, slot.Round, PickValue(slot.Step, option));

    /// <summary>True when the Pick's answer is this option.</summary>
    public bool IsPicked(ExerciseRun run, StepSlot slot, string option) =>
        GetAnswer(run, slot.Step.Key, slot.Round) is { Length: > 0 } answer
        && string.Equals(answer, PickValue(slot.Step, option), StringComparison.Ordinal);

    private static string AltKey(StepSlot slot) => slot.Id + "~alt";

    /// <summary>True while a step's second wording is on for this round (Extend's mystery lens).</summary>
    public bool IsAlt(ExerciseRun run, StepSlot slot) =>
        slot.Step.AltQuestion is not null && Read(run, AltKey(slot)) == "y";

    /// <summary>Switches a step's second wording on or off for this round only. Not an answer.</summary>
    public void SetAlt(ExerciseRun run, StepSlot slot, bool on)
    {
        if (slot.Step.AltQuestion is null)
        {
            return;
        }

        if (on)
        {
            run.Answers[AltKey(slot)] = "y";
        }
        else
        {
            run.Answers.Remove(AltKey(slot));
        }
    }

    /// <summary>The question as asked this round: the second wording while it is on.</summary>
    public string Question(ExerciseRun run, StepSlot slot) =>
        IsAlt(run, slot) && slot.Step.AltQuestion is { } alt ? alt : slot.Step.Question;

    public string Hint(ExerciseRun run, StepSlot slot) =>
        IsAlt(run, slot) ? slot.Step.AltHint : slot.Step.Hint;

    private string FirstEnd(ExerciseRun run, Step step) =>
        step.EndKeys.Select(key => GetAnswer(run, key)).FirstOrDefault(text => !Blank(text)) ?? string.Empty;

    /// <summary>
    /// What an Echo step is asking about: the chain's end in the first round, then the previous
    /// round's answer to the same step. Read at call time, never stored. Empty when there is none.
    /// </summary>
    public string EchoOf(ExerciseRun run, StepSlot slot)
    {
        if (!slot.Step.Echo)
        {
            return string.Empty;
        }

        return slot.Round == 0 ? FirstEnd(run, slot.Step) : GetAnswer(run, slot.Step.Key, slot.Round - 1);
    }

    /// <summary>
    /// A chain written backward, read forward: the loop's answers from the last round to the
    /// first, then the chain's end. Blank links are left out.
    /// </summary>
    public IReadOnlyList<string> ReadForward(ExerciseRun run, StepSlot slot)
    {
        if (slot.Step.ReadForwardKey is not { } key)
        {
            return Array.Empty<string>();
        }

        var chain = Enumerable.Range(0, Math.Max(1, run.Rounds)).Reverse()
            .Select(round => GetAnswer(run, key, round)).Where(text => !Blank(text)).ToList();
        if (FirstEnd(run, slot.Step) is { Length: > 0 } end)
        {
            chain.Add(end);
        }

        return chain;
    }

    /// <summary>Slots a Lines step shows: declared, or more when an Open step has been given more.</summary>
    public int LineCount(ExerciseRun run, StepSlot slot) => SlotCount(run, slot.Step, slot.Round);

    // Declared slots, or more when an Open step has been given more lines.
    private int SlotCount(ExerciseRun run, Step step, int round)
    {
        var count = Math.Max(1, step.Count);
        if (step.Open)
        {
            while (run.Answers.ContainsKey(LineKey(step.Key, round, count)))
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>The items a Slots step has a field for: the project list its source names, in order.</summary>
    public IReadOnlyList<string> SlotItems(StepSlot slot) =>
        slot.Step.Kind == StepKind.Slots && slot.Step.Source is { } source ? _slotSource(source) : Array.Empty<string>();

    /// <summary>
    /// Deals a Dice step a row and stores it. Rolling again replaces it with a
    /// different row (a re-roll that could land on the same one would feel
    /// broken); a one-row table can only repeat itself.
    /// </summary>
    public string? Roll(ExerciseRun run, StepSlot slot, Random random)
    {
        var table = slot.Step.Table;
        if (slot.Step.Kind != StepKind.Dice || table.Count == 0)
        {
            return null;
        }

        var current = Read(run, Suffix(slot.Step.Key, slot.Round));
        var candidates = table.Count > 1 ? table.Where(row => row != current).ToList() : table.ToList();
        var row = candidates[random.Next(candidates.Count)];
        SetAnswer(run, slot.Step.Key, slot.Round, row);
        return row;
    }

    public bool IsAnswered(ExerciseRun run, StepSlot slot)
    {
        var step = slot.Step;
        switch (step.Kind)
        {
            case StepKind.Stop:
                return run.Passed.Contains(slot.Id);
            case StepKind.Lines:
                return Enumerable.Range(0, SlotCount(run, step, slot.Round))
                    .Any(i => !Blank(GetLine(run, step.Key, slot.Round, i)));
            case StepKind.Slots:
                return (step.Source is { } source ? _slotSource(source) : Array.Empty<string>())
                    .Any(item => !Blank(GetItem(run, step.Key, slot.Round, item)));
            default:
                return !Blank(GetAnswer(run, step.Key, slot.Round));
        }
    }

    /// <summary>
    /// The step to show now: the first visible one neither answered nor
    /// passed. A Stop is current until continued past, but never hides or
    /// locks a later step that already has text. Null when nothing is left.
    /// </summary>
    public StepSlot? Current(ExerciseRun run)
    {
        foreach (var slot in Sequence(run))
        {
            if (!run.Passed.Contains(slot.Id) && !IsAnswered(run, slot))
            {
                return slot;
            }
        }

        return null;
    }

    /// <summary>Answered steps in order, one line each. Stops are not listed.</summary>
    public IReadOnlyList<TrailEntry> Trail(ExerciseRun run)
    {
        var trail = new List<TrailEntry>();
        foreach (var slot in Sequence(run))
        {
            if (slot.Step.Kind != StepKind.Stop && IsAnswered(run, slot))
            {
                trail.Add(new TrailEntry(slot, Summarise(run, slot)));
            }
        }

        return trail;
    }

    private string Summarise(ExerciseRun run, StepSlot slot)
    {
        var step = slot.Step;
        switch (step.Kind)
        {
            case StepKind.Lines:
                return string.Join(" / ", Enumerable.Range(0, SlotCount(run, step, slot.Round))
                    .Select(i => GetLine(run, step.Key, slot.Round, i)).Where(text => !Blank(text)));
            case StepKind.Slots:
                var items = step.Source is { } source ? _slotSource(source) : Array.Empty<string>();
                return string.Join("; ", items.Select(item => (item, text: GetItem(run, step.Key, slot.Round, item)))
                    .Where(pair => !Blank(pair.text)).Select(pair => $"{pair.item}: {pair.text}"));
            default:
                return GetAnswer(run, step.Key, slot.Round);
        }
    }

    /// <summary>Skip a step: it stops being current, nothing is stored or scored.</summary>
    public void Skip(ExerciseRun run, StepSlot slot)
    {
        if (!run.Passed.Contains(slot.Id))
        {
            run.Passed.Add(slot.Id);
        }
    }

    /// <summary>
    /// Jumping to a step ("got the real answer already? write it here and stop"):
    /// everything before it that has no answer is passed, so the run carries on from
    /// it rather than back to the questions it skipped. Stops are left for the writer.
    /// </summary>
    public void SkipUnansweredBefore(ExerciseRun run, StepSlot target)
    {
        foreach (var slot in Sequence(run))
        {
            if (slot.Id == target.Id)
            {
                return;
            }

            if (slot.Step.Kind != StepKind.Stop && !IsAnswered(run, slot) && !IsCarried(slot))
            {
                Skip(run, slot);
            }
        }
    }

    /// <summary>"Keep going" on a stop point. The same mark as a skip.</summary>
    public void Continue(ExerciseRun run, StepSlot stop) => Skip(run, stop);
}
