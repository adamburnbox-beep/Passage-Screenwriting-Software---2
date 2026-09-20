namespace Passage.Web.Services;

/// <summary>
/// Where a runner's one-line read-back goes when promoted into the script as
/// a <c>=</c> synopsis: under the nearest section or scene heading at or
/// above the caret, after any synopsis lines already there. Pure so it can be
/// tested without the editor.
/// </summary>
public static class SynopsisPlacement
{
    /// <summary>
    /// Returns the line index to insert before, and the index of the heading
    /// it sits under (or -1 when no heading precedes the caret, in which case
    /// the line goes in at the caret itself).
    /// </summary>
    public static (int InsertAt, int HeadingIndex) Find(string[] lineClasses, int caretIndex)
    {
        var caret = Math.Clamp(caretIndex, 0, Math.Max(0, lineClasses.Length));
        for (var index = Math.Min(caret, lineClasses.Length - 1); index >= 0; index--)
        {
            if (lineClasses[index] is not ("sx-section" or "sx-scene" or "md-heading"))
            {
                continue;
            }

            var insertAt = index + 1;
            while (insertAt < lineClasses.Length && lineClasses[insertAt] == "sx-synopsis")
            {
                insertAt++;
            }

            return (insertAt, index);
        }

        return (caret, -1);
    }
}
