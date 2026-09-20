using Passage.Core.Extensibility;

namespace Passage.Web.Services;

/// <summary>
/// Builds the ask for each place a partner can offer lines (SLATE-PLAN
/// Phase 7). Every one has the same shape: the worksheet so far, in its own
/// words, and one field to offer <see cref="Count"/> lines for. A builder
/// returns null when the worksheet has nothing left to ask for.
/// </summary>
public static class Suggestions
{
    public const int Count = 3;

    private static SuggestionLine Line(string label, string text) => new(label, text.Trim());

    private static bool Has(string text) => !string.IsNullOrWhiteSpace(text);

    /// <summary>Fill: three candidate events for the bracket, each one the writer expects to reject.</summary>
    public static SuggestionRequest ForFill(FillRun run, string scriptLine)
    {
        var context = new List<SuggestionLine> { Line("Bracket", run.Bracket) };
        if (Has(scriptLine)) context.Add(Line("The line it's on", scriptLine));
        for (var i = 0; i < run.Options.Count; i++)
        {
            var option = run.Options[i];
            if (Has(option.Text)) context.Add(Line($"Candidate {i + 1}", option.Text));
            if (Has(option.WhyWrong)) context.Add(Line($"Why candidate {i + 1} is probably wrong", option.WhyWrong));
        }
        if (Has(run.PointsTo)) context.Add(Line("What the wrongness points to", run.PointsTo));
        return new SuggestionRequest("Fill", "Candidate — what could happen here", context, Count);
    }

    /// <summary>A→Z→Split, rung 1: candidate midpoints between the fixed ends.</summary>
    public static SuggestionRequest ForSplitMidpoint(SplitRun run)
    {
        var context = new List<SuggestionLine>();
        if (Has(run.A)) context.Add(Line("A — start event", run.A));
        if (Has(run.Z)) context.Add(Line("Z — end event", run.Z));
        for (var i = 0; i < run.Candidates.Count; i++)
        {
            if (Has(run.Candidates[i])) context.Add(Line($"Candidate {i + 1}", run.Candidates[i]));
        }
        return new SuggestionRequest("A→Z→Split",
            "The midpoint — the one event that, if it happened, makes the second half inevitable given the first half",
            context, Count);
    }

    /// <summary>Bridge: candidate links for the round the writer is on (the last one).</summary>
    public static SuggestionRequest ForBridgeRound(BridgeRun run)
    {
        var context = new List<SuggestionLine>();
        if (Has(run.A)) context.Add(Line("A — the point you're starting from", run.A));
        if (Has(run.Z)) context.Add(Line("Z — the point you're building to", run.Z));
        var last = run.Rounds.Count - 1;
        for (var r = 0; r < last; r++)
        {
            if (Has(run.Rounds[r].Picked)) context.Add(Line($"Round {r + 1}, picked / half-picked", run.Rounds[r].Picked));
        }
        var round = run.Rounds[last];
        for (var c = 0; c < round.Candidates.Count; c++)
        {
            if (Has(round.Candidates[c])) context.Add(Line($"Round {last + 1}, candidate {c + 1}", round.Candidates[c]));
        }
        return new SuggestionRequest("Bridge", "A candidate link between A and Z", context, Count);
    }

    /// <summary>
    /// WOAC: the next empty answer of the round the writer is on, in the
    /// chain's own order. Character Flaw Brainstorm: the first unanswered
    /// question. Null when every field is written.
    /// </summary>
    public static SuggestionRequest? ForChain(ChainRun run, IReadOnlyList<string> flawQuestions)
    {
        var context = new List<SuggestionLine>();
        if (run.Path == ChainPath.Woac)
        {
            if (Has(run.Seed)) context.Add(Line("Character + starting want", run.Seed));
            for (var r = 0; r < run.Rounds.Count; r++)
            {
                var round = run.Rounds[r];
                // Round 1 writes its want; every later round reads the consequence above.
                if (r > 0 && Has(run.Rounds[r - 1].Consequence)) context.Add(Line($"Round {r + 1} — want", run.Rounds[r - 1].Consequence));
                var fields = new List<(string Label, string Ask, string Text)>();
                if (r == 0) fields.Add(("want", "Want — what does the character want right now?", round.Want));
                fields.Add(("obstacle", "Obstacle — what's in their way?", round.Obstacle));
                fields.Add(("action", "Action — what do they do about it?", round.Action));
                fields.Add(("consequence", "Consequence — what happens as a result?", round.Consequence));
                foreach (var (label, ask, text) in fields)
                {
                    if (Has(text))
                    {
                        context.Add(Line($"Round {r + 1} — {label}", text));
                        continue;
                    }
                    if (r != run.Rounds.Count - 1) continue;
                    return new SuggestionRequest("WOAC chain", ask, context, Count);
                }
            }
            return null;
        }

        if (Has(run.Seed)) context.Add(Line("Character + the flaw", run.Seed));
        for (var i = 0; i < flawQuestions.Count && i < run.Answers.Count; i++)
        {
            if (Has(run.Answers[i]))
            {
                context.Add(Line($"{i + 1}. {flawQuestions[i]}", run.Answers[i]));
                continue;
            }
            return new SuggestionRequest("Character Flaw Brainstorm", $"{i + 1}. {flawQuestions[i]}", context, Count);
        }
        return null;
    }

    /// <summary>Extend Backward: the open link — what had to be true right before the one above it.</summary>
    public static SuggestionRequest ForExtendLink(ExtendRun run)
    {
        var context = new List<SuggestionLine>();
        if (Has(run.Z)) context.Add(Line("Z — the ending", run.Z));
        var open = run.Links.Count - 1;
        for (var i = 0; i < open; i++)
        {
            if (Has(run.Links[i].Text)) context.Add(Line($"Link {i + 1} — right before {(i == 0 ? "Z" : $"link {i}")}", run.Links[i].Text));
        }
        var ask = run.Links[open].Mystery
            ? "Mystery lens: treat it as a crime scene — what would a detective need to find to explain it, rather than what caused it?"
            : "What had to be true right before this, for this to happen?";
        return new SuggestionRequest("Extend Backward", ask, context, Count);
    }
}
