using System.Text;
using Passage.Core.Goals;
using Passage.Core;
using Passage.Parser;

namespace Passage.Parser.Goals;

public sealed class ScreenplayPageEstimator
{
    private const int MaxLinesPerBodyPage = 42;
    private const int MaxLinesPerTitlePage = 38;
    private const int TitlePageMaxChars = 60;

    public int EstimatePageCount(ParsedScreenplay screenplay)
    {
        ArgumentNullException.ThrowIfNull(screenplay);

        var titlePageLines = BuildTitlePageLines(screenplay.TitlePage.Entries);
        var bodyLines = BuildBodyLines(screenplay.Elements);

        var pageCount = 0;

        if (titlePageLines.Count > 0)
        {
            pageCount += CountPages(titlePageLines.Count, MaxLinesPerTitlePage);
        }

        if (bodyLines.Count > 0)
        {
            pageCount += CountPages(bodyLines.Count, MaxLinesPerBodyPage);
        }

        return Math.Max(1, pageCount);
    }

    public Goal UpdatePageCountGoal(Goal goal, ParsedScreenplay screenplay)
    {
        ArgumentNullException.ThrowIfNull(goal);
        ArgumentNullException.ThrowIfNull(screenplay);

        if (goal.Type != GoalType.PageCount)
        {
            throw new ArgumentException("Goal must be a page count goal.", nameof(goal));
        }

        var currentValue = EstimatePageCount(screenplay);
        return new Goal(
            GoalType.PageCount,
            goal.TargetValue,
            currentValue,
            currentValue >= goal.TargetValue);
    }

    private static int CountPages(int lineCount, int maxLinesPerPage)
    {
        if (lineCount <= 0)
        {
            return 0;
        }

        return (lineCount + maxLinesPerPage - 1) / maxLinesPerPage;
    }

    private static List<string> BuildTitlePageLines(IReadOnlyList<TitlePageEntry> entries)
    {
        var lines = new List<string>();
        if (entries.Count == 0)
        {
            return lines;
        }

        AddBlankLines(lines, 8);

        var titleEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Title);
        var authorEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Author);
        var creditEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Credit);
        var draftDateEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.DraftDate);
        var contactEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Contact);

        if (titleEntry is not null)
        {
            AddWrappedBlock(lines, titleEntry.Value.ToUpperInvariant(), TitlePageMaxChars);
            AddBlankLines(lines, 2);
        }

        if (authorEntry is not null)
        {
            AddWrappedBlock(lines, authorEntry.Value, TitlePageMaxChars);
            AddBlankLines(lines, 4);
        }

        foreach (var entry in new[] { creditEntry, draftDateEntry, contactEntry }.Where(entry => entry is not null))
        {
            AddWrappedBlock(lines, $"{entry!.Label}: {entry.Value}", TitlePageMaxChars);
            AddBlankLines(lines, 1);
        }

        TrimTrailingBlankLines(lines);
        return lines;
    }

    private static List<string> BuildBodyLines(IReadOnlyList<ScreenplayElement> elements)
    {
        var lines = new List<string>();
        var renderableBlocks = BuildRenderableBlocks(elements);

        if (renderableBlocks.Count == 0)
        {
            return lines;
        }

        AddBlankLines(lines, ScreenplayFormatting.GetBlankLinesBefore(ToSpacingType(renderableBlocks[0].Type)));

        RenderBlock? previous = null;
        foreach (var block in renderableBlocks)
        {
            if (previous is not null)
            {
                AddBlankLines(
                    lines,
                    ScreenplayFormatting.GetGapBlankLines(ToSpacingType(previous.Value.Type), ToSpacingType(block.Type)));
            }

            AddWrappedBlock(lines, block.Text, block.MaxChars);
            previous = block;
        }

        AddBlankLines(lines, ScreenplayFormatting.GetBlankLinesAfter(ToSpacingType(renderableBlocks[^1].Type)));
        return lines;
    }

    private static void AddWrappedBlock(ICollection<string> lines, string text, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        foreach (var wrappedLine in WrapText(text, maxChars))
        {
            lines.Add(wrappedLine);
        }
    }

    private static List<RenderBlock> BuildRenderableBlocks(IReadOnlyList<ScreenplayElement> elements)
    {
        var blocks = new List<RenderBlock>(elements.Count);

        foreach (var element in elements)
        {
            if (element is NoteElement or BoneyardElement or SectionElement or SynopsisElement)
            {
                continue;
            }

            if (TryCreateRenderableBlock(element, out var block))
            {
                blocks.Add(block);
            }
        }

        return blocks;
    }

    private static bool TryCreateRenderableBlock(ScreenplayElement element, out RenderBlock block)
    {
        block = default;

        string text;
        int maxChars;

        switch (element)
        {
            case SceneHeadingElement sceneHeading:
                text = FountainMarkup.StripOmissions(sceneHeading.Text).ToUpperInvariant();
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
            case ActionElement action:
                text = FountainMarkup.StripOmissions(action.Text);
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
            case CharacterElement character:
                text = FountainMarkup.StripOmissions(character.CharacterName);
                maxChars = ScreenplayFormatting.CharacterWrapChars;
                break;
            case ParentheticalElement parenthetical:
                text = FountainMarkup.StripOmissions($"({parenthetical.Text})");
                maxChars = ScreenplayFormatting.ParentheticalWrapChars;
                break;
            case DialogueElement dialogue:
                text = FountainMarkup.StripOmissions(dialogue.Text);
                maxChars = ScreenplayFormatting.DialogueWrapChars;
                break;
            case TransitionElement transition:
                text = FountainMarkup.StripOmissions(transition.Text);
                maxChars = ScreenplayFormatting.TransitionWrapChars;
                break;
            default:
                text = FountainMarkup.StripOmissions(element.Text);
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        block = new RenderBlock(element.Type, text, maxChars);
        return true;
    }

    private static void AddBlankLines(ICollection<string> lines, int count)
    {
        for (var index = 0; index < count; index++)
        {
            lines.Add(string.Empty);
        }
    }

    private static List<string> WrapText(string text, int maxChars)
    {
        var result = new List<string>();

        foreach (var paragraph in text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n'))
        {
            var trimmed = paragraph.Trim();
            if (trimmed.Length == 0)
            {
                result.Add(string.Empty);
                continue;
            }

            var words = trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var current = new StringBuilder();

            foreach (var word in words)
            {
                if (current.Length == 0)
                {
                    current.Append(word);
                    continue;
                }

                if (current.Length + 1 + word.Length <= maxChars)
                {
                    current.Append(' ').Append(word);
                    continue;
                }

                result.Add(current.ToString());
                current.Clear();
                current.Append(word);
            }

            if (current.Length > 0)
            {
                result.Add(current.ToString());
            }
        }

        if (result.Count == 0)
        {
            result.Add(string.Empty);
        }

        return result;
    }

    private static void TrimTrailingBlankLines(ICollection<string> lines)
    {
        if (lines is not List<string> list)
        {
            return;
        }

        while (list.Count > 0 && string.IsNullOrEmpty(list[^1]))
        {
            list.RemoveAt(list.Count - 1);
        }
    }

    private static ScreenplayFormatting.ElementSpacingType ToSpacingType(ScreenplayElementType elementType)
    {
        return elementType switch
        {
            ScreenplayElementType.SceneHeading => ScreenplayFormatting.ElementSpacingType.SceneHeading,
            ScreenplayElementType.Action => ScreenplayFormatting.ElementSpacingType.Action,
            ScreenplayElementType.Character => ScreenplayFormatting.ElementSpacingType.Character,
            ScreenplayElementType.Dialogue => ScreenplayFormatting.ElementSpacingType.Dialogue,
            ScreenplayElementType.Parenthetical => ScreenplayFormatting.ElementSpacingType.Parenthetical,
            ScreenplayElementType.Transition => ScreenplayFormatting.ElementSpacingType.Transition,
            _ => ScreenplayFormatting.ElementSpacingType.Other
        };
    }

    private readonly record struct RenderBlock(
        ScreenplayElementType Type,
        string Text,
        int MaxChars);
}
