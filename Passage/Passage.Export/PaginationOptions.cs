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
public sealed record PaginationOptions(bool KeepTogether = true, int OverflowLines = 1);
