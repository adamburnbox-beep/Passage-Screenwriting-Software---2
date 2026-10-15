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

    public static (List<LayoutPage> titlePages, List<LayoutPage> bodyPages) BuildPages(ParsedScreenplay screenplay, PaginationOptions? options = null)
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

        int bodyPageNum = 1;
        var paginatedBody = options is { KeepTogether: true }
            ? PaginateKeepTogether(screenplay.Elements, MaxLinesPerBodyPage, options)
            : Paginate(BuildBodyLines(screenplay.Elements), MaxLinesPerBodyPage).ToList();
        
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

        block = new RenderBlock(element.Type, text, style, x, maxChars, element is CharacterElement { IsDualDialogue: true });
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

    // Keep-together pagination (PaginationOptions.KeepTogether). Lines are laid out exactly as
    // BuildBodyLines does; what changes is where a page may end. canBreakBefore[i] marks the
    // lines a new page may start at, and the glue rules below clear it.
    private static List<List<LayoutLine>> PaginateKeepTogether(IReadOnlyList<ScreenplayElement> elements, int maxLinesPerPage, PaginationOptions options)
    {
        var overflowLines = Math.Max(0, options.OverflowLines);
        var pages = new List<List<LayoutLine>>();
        var renderableBlocks = BuildRenderableBlocks(elements);
        if (renderableBlocks.Count == 0)
        {
            return pages;
        }

        var lines = new List<LayoutLine>();
        var canBreakBefore = new List<bool>();
        var endsBlock = new List<bool>();
        var blockOf = new List<int>();
        var leadStarts = new int[renderableBlocks.Count];
        var contentStarts = new int[renderableBlocks.Count];
        var contentCounts = new int[renderableBlocks.Count];

        for (var blockIndex = 0; blockIndex < renderableBlocks.Count; blockIndex++)
        {
            var block = renderableBlocks[blockIndex];
            var lead = blockIndex == 0
                ? GetBlankLinesBefore(block.Type)
                : GetGapBlankLines(renderableBlocks[blockIndex - 1].Type, block.Type);

            leadStarts[blockIndex] = lines.Count;
            AddBlankLines(lines, lead);
            contentStarts[blockIndex] = lines.Count;
            AddWrappedBlock(lines, block.Text, block.Style, block.X, block.MaxChars);
            contentCounts[blockIndex] = lines.Count - contentStarts[blockIndex];

            // A page may start before a block's leading gap or between its lines, never
            // inside the gap or between the gap and the block.
            for (var index = canBreakBefore.Count; index < lines.Count; index++)
            {
                canBreakBefore.Add(index == leadStarts[blockIndex] || index > contentStarts[blockIndex]);
                endsBlock.Add(index == lines.Count - 1);
                blockOf.Add(blockIndex);
            }
        }

        void Glue(int fromIndex, int toIndexExclusive)
        {
            for (var index = Math.Max(fromIndex, 0); index < Math.Min(toIndexExclusive, lines.Count); index++)
            {
                canBreakBefore[index] = false;
            }
        }

        for (var blockIndex = 0; blockIndex < renderableBlocks.Count; blockIndex++)
        {
            var type = renderableBlocks[blockIndex].Type;
            var contentEnd = contentStarts[blockIndex] + contentCounts[blockIndex];
            switch (type)
            {
                case ScreenplayElementType.SceneHeading:
                case ScreenplayElementType.Character:
                case ScreenplayElementType.Parenthetical:
                    // Unbreakable itself, and holds the next element's first two lines.
                    Glue(contentStarts[blockIndex] + 1, contentEnd);
                    if (blockIndex + 1 < renderableBlocks.Count)
                    {
                        Glue(leadStarts[blockIndex + 1], contentStarts[blockIndex + 1] + Math.Min(2, contentCounts[blockIndex + 1]));
                    }
                    break;
                case ScreenplayElementType.Transition:
                    // Unbreakable itself, and held by the previous element's last two lines.
                    Glue(contentStarts[blockIndex] + 1, contentEnd);
                    Glue(leadStarts[blockIndex], contentStarts[blockIndex] + 1);
                    if (blockIndex > 0)
                    {
                        var previousEnd = contentStarts[blockIndex - 1] + contentCounts[blockIndex - 1];
                        Glue(Math.Max(previousEnd - 1, contentStarts[blockIndex - 1] + 1), previousEnd);
                    }
                    break;
            }
        }

        // (MORE)/(CONT'D): a speech is a cue and the parentheticals and dialogue under it. Each
        // speech is glued whole and split explicitly by TrySplitSpeech when it overflows. A
        // dual-dialogue cue joins its speech to the one before it; that pair never splits.
        var speechCue = new int[renderableBlocks.Count];
        var splittable = new bool[renderableBlocks.Count];
        var blockText = renderableBlocks.Select(block => block.Text).ToArray();
        Array.Fill(speechCue, -1);
        if (options.MoreAndContinued)
        {
            var previousCue = -1;
            for (var blockIndex = 0; blockIndex < renderableBlocks.Count; blockIndex++)
            {
                var type = renderableBlocks[blockIndex].Type;
                if (type != ScreenplayElementType.Character) continue;

                var last = blockIndex;
                while (last + 1 < renderableBlocks.Count && renderableBlocks[last + 1].Type is ScreenplayElementType.Dialogue or ScreenplayElementType.Parenthetical)
                {
                    last++;
                }

                var isDual = renderableBlocks[blockIndex].IsDualCue && previousCue >= 0 && speechCue[blockIndex - 1] == previousCue;
                var glueFrom = isDual ? previousCue : blockIndex;
                Glue(contentStarts[glueFrom] + 1, contentStarts[last] + contentCounts[last]);
                for (var member = blockIndex; member <= last; member++)
                {
                    speechCue[member] = blockIndex;
                    splittable[member] = !isDual;
                }
                if (isDual)
                {
                    for (var member = previousCue; member < blockIndex; member++) splittable[member] = false;
                }

                previousCue = blockIndex;
                blockIndex = last;
            }
        }

        var current = new List<LayoutLine>();
        var start = 0;
        while (start < lines.Count)
        {
            var end = start + 1;
            while (end < lines.Count && !canBreakBefore[end]) end++;

            // Blank lines never open a page after the first.
            var unitStart = start;
            if (current.Count == 0 && pages.Count > 0)
            {
                while (unitStart < end && lines[unitStart].IsBlank) unitStart++;
            }

            var height = end - unitStart;
            if (height == 0 || current.Count + height <= maxLinesPerPage)
            {
                current.AddRange(lines.Skip(unitStart).Take(height));
                start = end;
                continue;
            }

            var contentLines = 0;
            for (var index = unitStart; index < end; index++)
            {
                if (!lines[index].IsBlank) contentLines++;
            }

            // A glued group, or the last line of a block already broken on this page, may hang
            // past the bottom margin rather than leave a one-line orphan.
            var isGlued = contentLines > 1;
            var isOrphanTail = contentLines == 1 && end - start == 1 && endsBlock[start] && !lines[start - 1].IsBlank && current.Count > 0;
            if ((isGlued || isOrphanTail) && current.Count + height <= maxLinesPerPage + overflowLines)
            {
                current.AddRange(lines.Skip(unitStart).Take(height));
                pages.Add(current);
                current = new List<LayoutLine>();
                start = end;
                continue;
            }

            if (options.MoreAndContinued && TrySplitSpeech(unitStart, end, maxLinesPerPage - current.Count, out var split))
            {
                // The page takes the head of the speech and (MORE); the lines from start through
                // the split are replaced by the repeated cue and the rest, which open the next page.
                current.AddRange(lines.Skip(unitStart).Take(split.ContentStart - unitStart));
                current.AddRange(split.Head);
                current.Add(new LayoutLine(options.MoreText, LayoutTextStyle.Left, MarginLeft + CharacterIndent));
                pages.Add(current);
                current = new List<LayoutLine>();

                var cue = renderableBlocks[speechCue[split.Block]];
                var cueText = cue.Text.EndsWith(options.ContinuedText, StringComparison.OrdinalIgnoreCase)
                    ? cue.Text
                    : $"{cue.Text} {options.ContinuedText}";
                var cueLines = new List<LayoutLine>();
                AddWrappedBlock(cueLines, cueText, cue.Style, cue.X, cue.MaxChars);
                if (split.TailText is not null) blockText[split.Block] = split.TailText;

                var removed = split.ResumeAt - start;
                lines.RemoveRange(start, removed);
                canBreakBefore.RemoveRange(start, removed);
                endsBlock.RemoveRange(start, removed);
                blockOf.RemoveRange(start, removed);
                var inserted = cueLines.Concat(split.Tail).ToList();
                lines.InsertRange(start, inserted);
                canBreakBefore.InsertRange(start, inserted.Select((_, index) => index == 0));
                endsBlock.InsertRange(start, inserted.Select((_, index) => index == cueLines.Count - 1 || (split.Tail.Count > 0 && index == inserted.Count - 1)));
                blockOf.InsertRange(start, inserted.Select((_, index) => index < cueLines.Count ? speechCue[split.Block] : split.Block));
                continue;
            }

            if (current.Count == 0)
            {
                // Taller than a page: nothing to keep it with, so cut it like the fixed paginator.
                current.AddRange(lines.Skip(unitStart).Take(maxLinesPerPage));
                start = unitStart + maxLinesPerPage;
            }

            pages.Add(current);
            current = new List<LayoutLine>();
        }

        var trailing = Math.Min(GetBlankLinesAfter(renderableBlocks[^1].Type), maxLinesPerPage - current.Count);
        if (current.Count > 0)
        {
            AddBlankLines(current, Math.Max(0, trailing));
            pages.Add(current);
        }

        return pages;

        // Picks where a speech in lines[unitStart, unitEnd) breaks so the head and (MORE) fit in
        // space lines: the most text that ends on a sentence end (or at the end of a dialogue
        // block followed by more of the speech), else on a line end. Only dialogue breaks, and at
        // least two of its lines, or all of them, stay on the page.
        bool TrySplitSpeech(int unitStart, int unitEnd, int space, out SpeechSplit best)
        {
            best = default;
            SpeechSplit? sentenceSplit = null;
            SpeechSplit? lineSplit = null;

            var lastBlock = -1;
            for (var index = unitStart; index < unitEnd; index++)
            {
                var blockIndex = blockOf[index];
                if (lines[index].IsBlank || blockIndex == lastBlock) continue;
                lastBlock = blockIndex;
                if (!splittable[blockIndex] || renderableBlocks[blockIndex].Type != ScreenplayElementType.Dialogue) continue;

                var contentStart = index;
                var contentEnd = contentStart;
                while (contentEnd < unitEnd && blockOf[contentEnd] == blockIndex) contentEnd++;
                var count = contentEnd - contentStart;
                var room = space - (contentStart - unitStart) - 1;
                var minimum = Math.Min(2, count);
                if (room < minimum) continue;

                var block = renderableBlocks[blockIndex];
                var text = blockText[blockIndex];
                foreach (var offset in SentenceEnds(text))
                {
                    var head = new List<LayoutLine>();
                    var tail = new List<LayoutLine>();
                    AddWrappedBlock(head, text[..offset], block.Style, block.X, block.MaxChars);
                    AddWrappedBlock(tail, text[offset..], block.Style, block.X, block.MaxChars);
                    if (head.Count > room) break;
                    if (head.Count < minimum) continue;
                    sentenceSplit = new SpeechSplit(blockIndex, contentStart, head, tail, text[offset..].Trim(), contentEnd);
                }

                // The end of a dialogue block with more of the speech after it.
                if (contentEnd < unitEnd && count <= room && blockOf[contentEnd] != blockIndex && speechCue[blockOf[contentEnd]] == speechCue[blockIndex])
                {
                    var next = contentEnd;
                    while (lines[next].IsBlank) next++;
                    sentenceSplit = new SpeechSplit(blockIndex, contentStart, lines.GetRange(contentStart, count), [], null, next);
                }

                if (sentenceSplit is null && count > minimum)
                {
                    var take = Math.Min(room, count - 1);
                    if (take >= minimum)
                    {
                        var tail = lines.GetRange(contentStart + take, count - take);
                        lineSplit = new SpeechSplit(blockIndex, contentStart, lines.GetRange(contentStart, take), tail,
                            string.Join('\n', tail.Select(line => line.Text)), contentEnd);
                    }
                }
            }

            best = sentenceSplit ?? lineSplit ?? default;
            return sentenceSplit is not null || lineSplit is not null;
        }
    }

    // Character offsets just past each sentence end in text that more text follows: '.', '!' or
    // '?' (with any closing quotes or brackets) before whitespace and a word that does not start
    // lowercase. "Mr." and similar titles are not sentence ends.
    private static IEnumerable<int> SentenceEnds(string text)
    {
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] is not ('.' or '!' or '?')) continue;

            var end = index + 1;
            while (end < text.Length && text[end] is '.' or '!' or '?' or '"' or '\'' or ')' or '\u2019' or '\u201D') end++;
            var next = end;
            while (next < text.Length && char.IsWhiteSpace(text[next])) next++;
            var isEnd = next > end && next < text.Length && !char.IsLower(text[next]) && !IsTitleAbbreviation(text, index);
            index = end - 1;
            if (isEnd) yield return end;
        }
    }

    private static bool IsTitleAbbreviation(string text, int periodIndex)
    {
        var wordStart = periodIndex;
        while (wordStart > 0 && char.IsLetter(text[wordStart - 1])) wordStart--;
        return text[periodIndex] == '.' && text[wordStart..periodIndex] is "Mr" or "Mrs" or "Ms" or "Dr" or "St" or "Jr" or "Sr" or "MR" or "MRS" or "MS" or "DR" or "ST" or "JR" or "SR";
    }

    // Head: the dialogue lines that end the page. Tail: what follows the repeated cue, ending at
    // ResumeAt (exclusive) in the original lines; TailText replaces the block's text when the
    // block itself was cut, so a later split re-wraps from it.
    private readonly record struct SpeechSplit(int Block, int ContentStart, List<LayoutLine> Head, List<LayoutLine> Tail, string? TailText, int ResumeAt);

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

    private readonly record struct RenderBlock(ScreenplayElementType Type, string Text, LayoutTextStyle Style, double X, int MaxChars, bool IsDualCue);
}
