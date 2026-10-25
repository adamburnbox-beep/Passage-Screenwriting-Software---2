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

    /// <summary>
    /// Any exercise step flagged Offer (docs/SLATE-REDO-PLAN.md §4.2): the
    /// trail so far as context, the step's own question as the ask.
    /// </summary>
    public static SuggestionRequest ForStep(string tool, string ask, IEnumerable<(string Label, string Text)> trail)
    {
        var context = trail.Where(entry => Has(entry.Text)).Select(entry => Line(entry.Label, entry.Text)).ToList();
        return new SuggestionRequest(tool, ask, context, Count);
    }
}
