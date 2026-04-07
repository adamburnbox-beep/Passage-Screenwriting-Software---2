using System.Globalization;
using System.Text;
using Passage.Core;
using Passage.Parser;

namespace Passage.Export;

public enum LayoutTextStyle
{
    Left,
    LeftBold,
    CenterWithinBody,
    CenterWithinBodyBold,
    RightWithinBody
}

public readonly record struct LayoutLine(string Text, LayoutTextStyle Style, double X)
{
    public bool IsBlank => Text.Length == 0;
    public static LayoutLine Blank() => new(string.Empty, LayoutTextStyle.Left, 0.0);
}

public readonly record struct LayoutPage(IReadOnlyList<LayoutLine> Lines, int PageNumber, bool IsTitlePage);

public static class ScreenplayLayoutBuilder
{
    public const double PageWidth = 612.0;
    public const double PageHeight = 792.0;
    public const double MarginLeft = 108.0;
    public const double MarginRight = 72.0;
    public const double MarginTop = 72.0;
    public const double MarginBottom = 72.0;
    public const double FontSize = 12.0;
    public const double LineHeight = FontSize; // Standard screenplay is 6 lines per inch (12pt lines)
    public const double PageNumberTopMargin = 36.0;
    public const double CharWidth = FontSize * 0.6;
    private const double DialogueIndent = 72.0;
    private const double CharacterIndent = 144.0;
    private const double ParentheticalIndent = 108.0;
    private static readonly int MaxLinesPerBodyPage = (int)Math.Floor((PageHeight - MarginTop - MarginBottom) / LineHeight);
    private static readonly int MaxLinesPerTitlePage = MaxLinesPerBodyPage;

    public static (List<LayoutPage> titlePages, List<LayoutPage> bodyPages) BuildPages(ParsedScreenplay screenplay)
    {
        var titlePages = new List<LayoutPage>();
        var bodyPages = new List<LayoutPage>();

        if (screenplay.TitlePage.Entries.Count > 0)
        {
            var titlePageLines = BuildTitlePage(screenplay.TitlePage.Entries);
            int pageNum = 0;
            foreach (var pageLines in Paginate(titlePageLines, MaxLinesPerTitlePage))
            {
                titlePages.Add(new LayoutPage(pageLines, pageNum++, true));
            }
        }

        var bodyLines = BuildBodyLines(screenplay.Elements);
        int bodyPageNum = 1;
        var paginatedBody = Paginate(bodyLines, MaxLinesPerBodyPage).ToList();
        
        if (paginatedBody.Count == 0)
        {
            paginatedBody.Add(new List<LayoutLine>());
        }

        foreach (var pageLines in paginatedBody)
        {
            bodyPages.Add(new LayoutPage(pageLines, bodyPageNum++, false));
        }

        return (titlePages, bodyPages);
    }

    private static List<LayoutLine> BuildTitlePage(IReadOnlyList<TitlePageEntry> entries)
    {
        var lines = new List<LayoutLine>();

        var titleEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Title);
        var authorEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Author);
        var creditEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Credit);
        var draftDateEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.DraftDate);
        var contactEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Contact);
        var sourceEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Source);
        var episodeEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Episode);
        var revisionEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Revision);
        var notesEntry = entries.FirstOrDefault(entry => entry.FieldType == TitlePageFieldType.Notes);

        var contactAlignRight = entries.Any(e => e.Label.Equals("Contact-Align", StringComparison.OrdinalIgnoreCase) && e.Value.Equals("Right", StringComparison.OrdinalIgnoreCase));
        var contactStyle = contactAlignRight ? LayoutTextStyle.RightWithinBody : LayoutTextStyle.Left;

        var topBlock = new List<LayoutLine>();
        if (titleEntry is not null)
        {
            AddWrappedBlock(topBlock, titleEntry.Value.ToUpperInvariant(), LayoutTextStyle.CenterWithinBodyBold, 0.0, ScreenplayFormatting.ActionWrapChars);
            if (episodeEntry is not null) AddWrappedBlock(topBlock, episodeEntry.Value, LayoutTextStyle.CenterWithinBody, 0.0, ScreenplayFormatting.ActionWrapChars);
        }

        if (topBlock.Count > 0) AddBlankLines(topBlock, 3);

        if (creditEntry is not null)
        {
            AddWrappedBlock(topBlock, creditEntry.Value, LayoutTextStyle.CenterWithinBody, 0.0, ScreenplayFormatting.ActionWrapChars);
            AddBlankLines(topBlock, 1);
        }
        else if (authorEntry is not null)
        {
            AddWrappedBlock(topBlock, "WRITTEN BY", LayoutTextStyle.CenterWithinBody, 0.0, ScreenplayFormatting.ActionWrapChars);
            AddBlankLines(topBlock, 1);
        }

        if (authorEntry is not null)
        {
            AddWrappedBlock(topBlock, authorEntry.Value, LayoutTextStyle.CenterWithinBody, 0.0, ScreenplayFormatting.ActionWrapChars);
        }

        if (sourceEntry is not null)
        {
            AddBlankLines(topBlock, 2);
            AddWrappedBlock(topBlock, sourceEntry.Value, LayoutTextStyle.CenterWithinBody, 0.0, ScreenplayFormatting.ActionWrapChars);
        }

        var bottomBlock = new List<LayoutLine>();
        if (revisionEntry is not null) AddWrappedBlock(bottomBlock, revisionEntry.Value, contactStyle, contactAlignRight ? 0.0 : MarginLeft, ScreenplayFormatting.ActionWrapChars);
        if (draftDateEntry is not null) AddWrappedBlock(bottomBlock, draftDateEntry.Value, contactStyle, contactAlignRight ? 0.0 : MarginLeft, ScreenplayFormatting.ActionWrapChars);
        if (notesEntry is not null) AddWrappedBlock(bottomBlock, notesEntry.Value, contactStyle, contactAlignRight ? 0.0 : MarginLeft, ScreenplayFormatting.ActionWrapChars);
        if (contactEntry is not null) AddWrappedBlock(bottomBlock, contactEntry.Value, contactStyle, contactAlignRight ? 0.0 : MarginLeft, ScreenplayFormatting.ActionWrapChars);

        int totalUsableLines = MaxLinesPerTitlePage;
        int topBlockCount = topBlock.Count;
        int bottomBlockCount = bottomBlock.Count;

        int leadingPadding = (totalUsableLines - topBlockCount - bottomBlockCount) / 2;
        if (leadingPadding < 0) leadingPadding = 0;

        AddBlankLines(lines, leadingPadding);
        lines.AddRange(topBlock);

        int middlePadding = totalUsableLines - lines.Count - bottomBlockCount;
        if (middlePadding > 0) AddBlankLines(lines, middlePadding);
        
        lines.AddRange(bottomBlock);

        return lines;
    }

    private static List<LayoutLine> BuildBodyLines(IReadOnlyList<ScreenplayElement> elements)
    {
        var lines = new List<LayoutLine>();
        var renderableBlocks = BuildRenderableBlocks(elements);

        if (renderableBlocks.Count == 0)
        {
            return lines;
        }

        AddBlankLines(lines, GetBlankLinesBefore(renderableBlocks[0].Type));

        RenderBlock? previous = null;
        foreach (var block in renderableBlocks)
        {
            if (previous is not null)
            {
                AddBlankLines(lines, GetGapBlankLines(previous.Value.Type, block.Type));
            }

            AddWrappedBlock(lines, block.Text, block.Style, block.X, block.MaxChars);
            previous = block;
        }

        AddBlankLines(lines, GetBlankLinesAfter(renderableBlocks[^1].Type));
        return lines;
    }

    private static void AddWrappedBlock(ICollection<LayoutLine> lines, string text, LayoutTextStyle style, double x, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        foreach (var wrappedLine in WrapText(text, maxChars))
        {
            lines.Add(new LayoutLine(wrappedLine, style, x));
        }
    }

    private static List<RenderBlock> BuildRenderableBlocks(IReadOnlyList<ScreenplayElement> elements)
    {
        var blocks = new List<RenderBlock>(elements.Count);
        foreach (var element in elements)
        {
            if (element is NoteElement or BoneyardElement or SectionElement or SynopsisElement) continue;

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
        LayoutTextStyle style;
        double x;
        int maxChars;

        switch (element)
        {
            case SceneHeadingElement sceneHeading:
                text = FountainMarkup.StripOmissions(sceneHeading.Text).Trim().ToUpperInvariant();
                style = LayoutTextStyle.LeftBold;
                x = MarginLeft;
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
            case ActionElement action:
                text = FountainMarkup.StripOmissions(action.Text);
                style = LayoutTextStyle.Left;
                x = MarginLeft;
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
            case CharacterElement character:
                text = FountainMarkup.StripOmissions(character.CharacterName).ToUpperInvariant();
                style = LayoutTextStyle.Left;
                x = MarginLeft + CharacterIndent;
                maxChars = ScreenplayFormatting.CharacterWrapChars;
                break;
            case ParentheticalElement parenthetical:
                text = FountainMarkup.StripOmissions($"({parenthetical.Text})");
                style = LayoutTextStyle.Left;
                x = MarginLeft + ParentheticalIndent;
                maxChars = ScreenplayFormatting.ParentheticalWrapChars;
                break;
            case DialogueElement dialogue:
                text = FountainMarkup.StripOmissions(dialogue.Text);
                style = LayoutTextStyle.Left;
                x = MarginLeft + DialogueIndent;
                maxChars = ScreenplayFormatting.DialogueWrapChars;
                break;
            case TransitionElement transition:
                text = FountainMarkup.StripOmissions(transition.Text).ToUpperInvariant();
                style = LayoutTextStyle.RightWithinBody;
                x = 0.0;
                maxChars = ScreenplayFormatting.TransitionWrapChars;
                break;
            default:
                text = FountainMarkup.StripOmissions(element.Text);
                style = LayoutTextStyle.Left;
                x = MarginLeft;
                maxChars = ScreenplayFormatting.ActionWrapChars;
                break;
        }

        if (string.IsNullOrWhiteSpace(text)) return false;

        block = new RenderBlock(element.Type, text, style, x, maxChars);
        return true;
    }

    private static void AddBlankLines(ICollection<LayoutLine> lines, int count)
    {
        for (var i = 0; i < count; i++)
        {
            lines.Add(LayoutLine.Blank());
        }
    }

    private static IEnumerable<List<LayoutLine>> Paginate(IReadOnlyList<LayoutLine> lines, int maxLinesPerPage)
    {
        if (lines.Count == 0)
        {
            yield return new List<LayoutLine>();
            yield break;
        }

        for (var index = 0; index < lines.Count; index += maxLinesPerPage)
        {
            var pageLines = new List<LayoutLine>();
            for (var lineIndex = index; lineIndex < Math.Min(index + maxLinesPerPage, lines.Count); lineIndex++)
            {
                pageLines.Add(lines[lineIndex]);
            }
            yield return pageLines;
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

    private static int GetBlankLinesBefore(ScreenplayElementType elementType) => ScreenplayFormatting.GetBlankLinesBefore(ToSpacingType(elementType));
    private static int GetBlankLinesAfter(ScreenplayElementType elementType) => ScreenplayFormatting.GetBlankLinesAfter(ToSpacingType(elementType));

    private static int GetGapBlankLines(ScreenplayElementType previousType, ScreenplayElementType currentType)
    {
        return ScreenplayFormatting.GetGapBlankLines(ToSpacingType(previousType), ToSpacingType(currentType));
    }

    private readonly record struct RenderBlock(ScreenplayElementType Type, string Text, LayoutTextStyle Style, double X, int MaxChars);
}
