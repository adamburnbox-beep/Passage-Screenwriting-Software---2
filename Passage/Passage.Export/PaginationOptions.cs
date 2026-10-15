namespace Passage.Export;

/// <summary>
/// Opt-in body pagination behaviour for <see cref="ScreenplayLayoutBuilder.BuildPages"/>.
/// Passing <c>null</c> keeps the original fixed 54-line cut, which the desktop apps rely on.
/// </summary>
/// <param name="KeepTogether">
/// Fit groups instead of lines: a scene heading stays with the next element's first two
/// lines, a character cue and parenthetical stay with the first two dialogue lines, and a
/// transition stays with the last two lines of the element before it.
/// </param>
/// <param name="OverflowLines">
/// How far a glued group, or the last line of a block already broken on the page, may run
/// past the bottom margin rather than move to the next page.
/// </param>
/// <param name="MoreAndContinued">
/// With <paramref name="KeepTogether"/>, a speech that overflows the page splits at the last
/// sentence end that fits (a line end if no sentence ends in the space), with
/// <paramref name="MoreText"/> under it and the cue repeated with <paramref name="ContinuedText"/>
/// on the next page. Parentheticals and dual dialogue never split.
/// </param>
/// <param name="MoreText">Written at the cue indent at the bottom of the page.</param>
/// <param name="ContinuedText">Appended, after a space, to the cue repeated on the next page.</param>
public sealed record PaginationOptions(
    bool KeepTogether = true,
    int OverflowLines = 1,
    bool MoreAndContinued = true,
    string MoreText = "(MORE)",
    string ContinuedText = "(CONT'D)");
