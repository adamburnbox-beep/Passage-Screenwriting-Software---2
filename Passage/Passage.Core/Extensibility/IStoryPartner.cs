namespace Passage.Core.Extensibility;

/// <summary>
/// The assist seam for the Writer's Tools (docs/SLATE-PLAN.md, Phase 7).
/// A partner is handed a partially filled worksheet and asked for a few
/// candidate lines for one field. It never writes anything itself; the
/// lines come back and the writer takes one or none.
/// </summary>
public interface IStoryPartner
{
    Task<IReadOnlyList<string>> SuggestAsync(SuggestionRequest request, CancellationToken cancellationToken);
}

/// <summary>
/// One ask. <paramref name="Tool"/> is the worksheet's name as the writer
/// sees it, <paramref name="Ask"/> the field's own question in the
/// worksheet's words, <paramref name="Context"/> the lines already on the
/// worksheet in reading order, <paramref name="Count"/> how many lines to
/// offer.
/// </summary>
public sealed record SuggestionRequest(string Tool, string Ask, IReadOnlyList<SuggestionLine> Context, int Count);

public sealed record SuggestionLine(string Label, string Text);

/// <summary>The default partner: nothing is asked, nothing comes back.</summary>
public sealed class NullStoryPartner : IStoryPartner
{
    public Task<IReadOnlyList<string>> SuggestAsync(SuggestionRequest request, CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<string>>(Array.Empty<string>());
}
