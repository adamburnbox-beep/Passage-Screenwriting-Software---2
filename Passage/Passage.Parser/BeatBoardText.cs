using System;
using System.Collections.Generic;
using System.Linq;

namespace Passage.Parser;

/// <summary>
/// Line-range and splicing helpers for the Beat Board, shared by every
/// frontend. Lifted verbatim in behaviour from the Avalonia
/// MainWindowViewModel, where they were untested and unreachable from the web
/// app. Pure string and list logic: no UI type appears in the signatures.
///
/// These live in Passage.Parser rather than Passage.Core because they work on
/// <see cref="ScreenplayElement"/> and its subclasses, and Parser already
/// depends on Core — putting them in Core would make that reference circular.
/// </summary>
public static class BeatBoardText
{
    /// <summary>A half-open-free, inclusive line range; both indices are 0-based.</summary>
    public readonly record struct LineRange(int StartLineIndex, int EndLineIndex)
    {
        public static LineRange NotFound { get; } = new(-1, -1);

        public bool IsFound => StartLineIndex >= 0;

        public int LineCount => IsFound ? EndLineIndex - StartLineIndex + 1 : 0;
    }

    /// <summary>
    /// The lines a card owns. With <paramref name="includeNestedBlock"/> false
    /// this is the heading/note element plus its trailing synopsis lines only.
    /// With it true, an Act/Sequence card also covers everything nested beneath
    /// it up to the next section of the same or higher level, and a Scene card
    /// covers up to the next scene heading or section — which is what dragging
    /// a card on the board has to move.
    /// </summary>
    public static LineRange GetCardLineRange(
        IReadOnlyList<ScreenplayElement> elements,
        Guid cardId,
        int totalLineCount,
        bool includeNestedBlock)
    {
        if (elements is null)
        {
            return LineRange.NotFound;
        }

        var elementIndex = -1;
        for (var i = 0; i < elements.Count; i++)
        {
            if (elements[i].Id == cardId)
            {
                elementIndex = i;
                break;
            }
        }

        if (elementIndex < 0)
        {
            return LineRange.NotFound;
        }

        var element = elements[elementIndex];
        var startLineIdx = element.LineIndex;
        var endLineIdx = element.EndLineIndex;

        // Trailing synopsis lines belong to the card that precedes them — but
        // only once something has marked them suppressed, which is what the
        // board build does when it folds a synopsis into a card's description.
        // An unclaimed synopsis is still a card in its own right.
        for (var i = elementIndex + 1; i < elements.Count; i++)
        {
            var next = elements[i];
            if (next.IsSuppressed && next.Type == ScreenplayElementType.Synopsis)
            {
                endLineIdx = next.EndLineIndex;
            }
            else
            {
                break;
            }
        }

        if (includeNestedBlock && element is SectionElement section)
        {
            var blockEndLineIdx = totalLineCount - 1;
            for (var i = elementIndex + 1; i < elements.Count; i++)
            {
                if (elements[i] is SectionElement nextSection && nextSection.SectionDepth <= section.SectionDepth)
                {
                    blockEndLineIdx = nextSection.LineIndex - 1;
                    break;
                }
            }

            endLineIdx = Math.Max(endLineIdx, blockEndLineIdx);
        }

        if (includeNestedBlock && element is SceneHeadingElement)
        {
            var blockEndLineIdx = totalLineCount - 1;
            for (var i = elementIndex + 1; i < elements.Count; i++)
            {
                var next = elements[i];
                if (next.Type is ScreenplayElementType.SceneHeading or ScreenplayElementType.Section)
                {
                    blockEndLineIdx = next.LineIndex - 1;
                    break;
                }
            }

            endLineIdx = Math.Max(endLineIdx, blockEndLineIdx);
        }

        return new LineRange(startLineIdx, endLineIdx);
    }

    /// <summary>
    /// The Fountain lines that represent a card: its heading in the syntax its
    /// type requires, carrying the id so the card survives a reparse, followed
    /// by one "= " synopsis line per non-blank description line.
    /// </summary>
    public static List<string> BuildCardLines(string type, string heading, string description, Guid id)
    {
        var lines = new List<string>();
        var trimmedHeading = (heading ?? string.Empty).Trim();

        switch (type)
        {
            case "Act":
            case "Sequence":
            case "Section":
            {
                var depth = type == "Act" ? 1 : type == "Sequence" ? 2 : 3;
                lines.Add($"{new string('#', depth)} {trimmedHeading} [[id:{id}]]");
                break;
            }

            case "Scene":
            {
                var isSceneSyntax =
                    trimmedHeading.StartsWith("INT.", StringComparison.OrdinalIgnoreCase) ||
                    trimmedHeading.StartsWith("EXT.", StringComparison.OrdinalIgnoreCase) ||
                    trimmedHeading.StartsWith("I/E.", StringComparison.OrdinalIgnoreCase) ||
                    trimmedHeading.StartsWith(".", StringComparison.Ordinal);

                lines.Add(isSceneSyntax
                    ? $"{trimmedHeading} [[id:{id}]]"
                    : $". {trimmedHeading} [[id:{id}]]");
                break;
            }

            case "Note":
                lines.Add($"[[{trimmedHeading} id:{id}]]");
                break;
        }

        if (!string.IsNullOrWhiteSpace(description))
        {
            foreach (var line in description.Replace("\r\n", "\n").Split('\n'))
            {
                var trimmed = line.Trim();
                if (trimmed.Length > 0)
                {
                    lines.Add($"= {trimmed}");
                }
            }
        }

        return lines;
    }

    /// <summary>
    /// Swaps an inclusive line range for new lines. Returns the text unchanged
    /// when the range starts outside the document, and clamps a range that runs
    /// off the end — both cases the desktop code already tolerated.
    /// </summary>
    public static string ReplaceLines(
        string text,
        int startLineIndex,
        int endLineIndex,
        IReadOnlyList<string> replacementLines)
    {
        var lines = (text ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();

        if (startLineIndex < 0 || startLineIndex >= lines.Count)
        {
            return text ?? string.Empty;
        }

        var countToRemove = endLineIndex - startLineIndex + 1;
        if (countToRemove > 0)
        {
            lines.RemoveRange(startLineIndex, Math.Min(countToRemove, lines.Count - startLineIndex));
        }

        if (replacementLines is { Count: > 0 })
        {
            lines.InsertRange(startLineIndex, replacementLines);
        }

        return string.Join("\n", lines);
    }

    /// <summary>
    /// Works out the single contiguous splice that moves <paramref name="source"/>
    /// to sit before or after <paramref name="target"/>. Ports
    /// MoveBeatBoardCardText, but returns the edit rather than applying it, so a
    /// caller can splice one range and keep the undo history — a move never
    /// changes the line count, so the affected region maps one-to-one onto its
    /// replacement.
    ///
    /// Returns false for the moves the desktop refuses: onto itself, or an
    /// Act/Sequence dropped onto a card nested inside its own block.
    /// </summary>
    public static bool TryPlanMove(
        IReadOnlyList<string> lines,
        LineRange source,
        LineRange target,
        bool insertAfter,
        out LineRange splice,
        out List<string> replacement)
    {
        splice = LineRange.NotFound;
        replacement = new List<string>();

        if (lines is null || !source.IsFound || !target.IsFound)
        {
            return false;
        }

        if (source.StartLineIndex == target.StartLineIndex)
        {
            return false;
        }

        if (target.StartLineIndex >= source.StartLineIndex && target.EndLineIndex <= source.EndLineIndex)
        {
            return false;
        }

        var sourceCount = source.EndLineIndex - source.StartLineIndex + 1;
        if (source.StartLineIndex < 0 || sourceCount <= 0 ||
            source.StartLineIndex + sourceCount > lines.Count)
        {
            return false;
        }

        var targetInsert = insertAfter ? target.EndLineIndex + 1 : target.StartLineIndex;

        var working = new List<string>(lines);
        var moved = working.GetRange(source.StartLineIndex, sourceCount);
        working.RemoveRange(source.StartLineIndex, sourceCount);

        var adjustedTarget = targetInsert;
        if (adjustedTarget > source.StartLineIndex)
        {
            adjustedTarget = Math.Max(0, adjustedTarget - sourceCount);
        }

        adjustedTarget = Math.Min(adjustedTarget, working.Count);
        working.InsertRange(adjustedTarget, moved);

        // Only the span between the old and new homes actually changed.
        var lo = Math.Min(source.StartLineIndex, targetInsert);
        var hi = Math.Max(source.EndLineIndex, targetInsert - 1);
        lo = Math.Max(0, lo);
        hi = Math.Min(hi, working.Count - 1);

        if (hi < lo)
        {
            return false;
        }

        splice = new LineRange(lo, hi);
        replacement = working.GetRange(lo, hi - lo + 1);
        return true;
    }
}
