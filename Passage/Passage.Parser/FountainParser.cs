using System;
using System.Text;
using System.Text.RegularExpressions;
using Passage.Core;

namespace Passage.Parser;

public sealed class FountainParser
{
    private static readonly Regex KnownTransitionRegex = new(
        @"^(?:FADE IN|FADE OUT|CUT TO|DISSOLVE TO|SMASH CUT TO|MATCH CUT TO|WIPE TO|JUMP CUT TO|FADE TO BLACK)(?:[:.])?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex SectionRegex = new(
        @"^(#{1,3})\s*(.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex SynopsisRegex = new(
        @"^=\s*(.*)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly string[] SceneHeadingPrefixes =
    [
        "INT",
        "INT.",
        "EXT",
        "EXT.",
        "INT/EXT",
        "INT/EXT.",
        "INT./EXT.",
        "EXT/INT",
        "EXT/INT.",
        "EXT./INT.",
        "I/E",
        "I.E.",
        "EST",
        "EST."
    ];

    public ParsedScreenplay Parse(string? screenplayText)
    {
        return Parse(screenplayText, null);
    }

    public ParsedScreenplay Parse(string? screenplayText, IReadOnlyDictionary<int, ScreenplayElementType>? lineTypeOverrides)
    {
        screenplayText ??= string.Empty;
        lineTypeOverrides ??= new Dictionary<int, ScreenplayElementType>();

        var lines = ReadLines(screenplayText);
        var titlePage = ParseTitlePage(lines, out var bodyStartIndex);
        var elements = new List<ScreenplayElement>();
        var actionLineIndices = new List<int>();
        var currentSectionDepth = (int?)null;

        void FlushAction()
        {
            if (actionLineIndices.Count == 0)
            {
                return;
            }

            var startIndex = actionLineIndices[0];
            var endIndex = actionLineIndices[^1];
            var rawText = BuildRawText(lines, startIndex, endIndex);
            var textWithId = BuildDisplayText(lines, startIndex, endIndex);
            
            var text = ExtractId(textWithId, out var id);
            var action = new ActionElement(text, rawText, startIndex, endIndex);
            if (id != Guid.Empty) action.Id = id;
            
            elements.Add(action);
            actionLineIndices.Clear();
        }

        var index = bodyStartIndex;
        while (index < lines.Count)
        {
            var line = lines[index];
            var trimmed = line.Text.Trim();

            if (trimmed.Length == 0)
            {
                FlushAction();
                index++;
                continue;
            }

            if (TryParseExplicitLine(line, lineTypeOverrides, insideDialogueBlock: false, characterContext: null, out var forcedElement, out _))
            {
                FlushAction();

                if (forcedElement is CharacterElement forcedCharacterElement)
                {
                    elements.Add(forcedCharacterElement);

                    index++;
                    index = ParseDialogueBlock(lines, index, forcedCharacterElement, elements, lineTypeOverrides);
                    continue;
                }

                elements.Add(forcedElement);
                index++;
                continue;
            }

            if (TryParseBoneyard(line, lines, index, out var boneyardElement, out var boneyardNextIndex))
            {
                FlushAction();
                elements.Add(boneyardElement);
                index = boneyardNextIndex;
                continue;
            }

            if (TryParseNote(line, lines, index, out var noteElement, out var noteNextIndex))
            {
                FlushAction();
                elements.Add(noteElement);
                index = noteNextIndex;
                continue;
            }

            if (TryParseCenteredText(line, out var centeredTextElement))
            {
                FlushAction();
                elements.Add(centeredTextElement);
                index++;
                continue;
            }

            if (TryParseLyrics(line, out var lyricsElement))
            {
                FlushAction();
                elements.Add(lyricsElement);
                index++;
                continue;
            }

            if (TryParseSection(line, out var sectionElement))
            {
                FlushAction();
                currentSectionDepth = sectionElement.SectionDepth;
                elements.Add(sectionElement);
                index++;
                continue;
            }

            if (TryParseSynopsis(line, currentSectionDepth, out var synopsisElement))
            {
                FlushAction();
                elements.Add(synopsisElement);
                index++;
                continue;
            }

            if (TryParseTransition(line, out var transitionElement))
            {
                FlushAction();
                elements.Add(transitionElement);
                index++;
                continue;
            }

            if (TryParseSceneHeading(line, out var sceneHeadingElement))
            {
                FlushAction();
                elements.Add(sceneHeadingElement);
                index++;
                continue;
            }

            if (TryParseCharacterCue(line, lines, index, out var cueCharacterElement))
            {
                FlushAction();
                elements.Add(cueCharacterElement);

                index++;
                index = ParseDialogueBlock(lines, index, cueCharacterElement, elements, lineTypeOverrides);
                continue;
            }

            actionLineIndices.Add(index);
            index++;
        }

        FlushAction();

        return new ParsedScreenplay(screenplayText, titlePage, elements, lineTypeOverrides);
    }

    private static TitlePageData ParseTitlePage(IReadOnlyList<SourceLine> lines, out int bodyStartIndex)
    {
        var entries = new List<TitlePageEntry>();
        var index = 0;

        while (index < lines.Count && lines[index].Text.Trim().Length == 0)
        {
            index++;
        }

        while (index < lines.Count)
        {
            var line = lines[index];
            var trimmed = line.Text.Trim();

            if (trimmed.Length == 0)
            {
                // A blank line doesn't immediately end the title page unless the next item is not a title page field.
                index++;
                continue;
            }

            if ((line.Text.StartsWith("   ", StringComparison.Ordinal) || line.Text.StartsWith('\t')) && entries.Count > 0)
            {
                var prev = entries[^1];
                var appendedValue = prev.Value.Length > 0 ? prev.Value + "\n" + trimmed : trimmed;
                entries[^1] = new TitlePageEntry(prev.FieldType, prev.RawText + "\n" + line.Text, appendedValue, prev.LineIndex, prev.CustomLabel);
                index++;
                continue;
            }

            if (TryParseTitlePageField(line, out var entry))
            {
                entries.Add(entry);
                index++;
                continue;
            }

            break;
        }

        while (index < lines.Count && lines[index].Text.Trim().Length == 0)
        {
            index++;
        }

        bodyStartIndex = index;
        return new TitlePageData(entries, bodyStartIndex);
    }

    private static int ParseDialogueBlock(
        IReadOnlyList<SourceLine> lines,
        int startIndex,
        CharacterElement dialogueCharacterElement,
        ICollection<ScreenplayElement> elements,
        IReadOnlyDictionary<int, ScreenplayElementType> lineTypeOverrides)
    {
        var blockLineIndices = new List<int>();
        var dialogueLineIndices = new List<int>();
        var parentheticalTexts = new List<string>();

        var index = startIndex;
        while (index < lines.Count)
        {
            var line = lines[index];
            var trimmed = line.Text.Trim();

            if (trimmed.Length == 0)
            {
                break;
            }

            if (TryParseExplicitLine(line, lineTypeOverrides, insideDialogueBlock: true, dialogueCharacterElement, out var forcedElement, out var shouldBreakBlock))
            {
                if (forcedElement is DialogueElement)
                {
                    blockLineIndices.Add(index);
                    dialogueLineIndices.Add(index);
                    index++;
                    continue;
                }

                if (forcedElement is ParentheticalElement)
                {
                    elements.Add(forcedElement);
                    blockLineIndices.Add(index);
                    parentheticalTexts.Add(forcedElement.Text);
                    index++;
                    continue;
                }
            }

            if (shouldBreakBlock)
            {
                break;
            }

            if (TryParseSection(line, out _) ||
                TryParseSynopsis(line, null, out _) ||
                TryParseTransition(line, out _) ||
                TryParseSceneHeading(line, out _) ||
                TryParseCharacterCue(line, lines, index, out _) ||
                TryParseNote(line, lines, index, out _, out _) ||
                TryParseBoneyard(line, lines, index, out _, out _) ||
                TryParseCenteredText(line, out _) ||
                TryParseLyrics(line, out _))
            {
                break;
            }

            blockLineIndices.Add(index);

            if (IsParentheticalLine(trimmed))
            {
                var parentheticalText = trimmed[1..^1].Trim();
                parentheticalTexts.Add(parentheticalText);

                var rawText = BuildRawText(lines, index, index);
                elements.Add(new ParentheticalElement(
                    parentheticalText,
                    rawText,
                    index,
                    dialogueCharacterElement.CharacterName,
                    dialogueCharacterElement.IsDualDialogue));
            }
            else
            {
                dialogueLineIndices.Add(index);
            }

            index++;
        }

        if (dialogueLineIndices.Count > 0)
        {
            var startLineIndex = blockLineIndices[0];
            var endLineIndex = blockLineIndices[^1];
            var rawText = BuildRawText(lines, startLineIndex, endLineIndex);
            var text = BuildDisplayText(lines, dialogueLineIndices[0], dialogueLineIndices[^1]);
            var spokenLines = dialogueLineIndices
                .Select(lineIndex => lines[lineIndex].Text)
                .ToArray();

            elements.Add(new DialogueElement(
                dialogueCharacterElement.CharacterName,
                parentheticalTexts,
                spokenLines,
                text,
                rawText,
                startLineIndex,
                endLineIndex,
                dialogueCharacterElement.IsDualDialogue));
        }

        return index;
    }

    private static bool TryParseExplicitLine(
        SourceLine line,
        IReadOnlyDictionary<int, ScreenplayElementType> lineTypeOverrides,
        bool insideDialogueBlock,
        CharacterElement? characterContext,
        out ScreenplayElement element,
        out bool shouldBreakBlock)
    {
        element = default!;
        shouldBreakBlock = false;

        var trimmed = line.Text.Trim();
        if (trimmed.Length == 0)
        {
            return false;
        }

        if (!lineTypeOverrides.TryGetValue(line.Index + 1, out var forcedType))
        {
            return false;
        }

        switch (forcedType)
        {
            case ScreenplayElementType.Action:
                if (insideDialogueBlock)
                {
                    shouldBreakBlock = true;
                    return false;
                }

                element = new ActionElement(trimmed, line.Text, line.Index, line.Index);
                return true;
            case ScreenplayElementType.SceneHeading:
                if (insideDialogueBlock)
                {
                    shouldBreakBlock = true;
                    return false;
                }

                element = new SceneHeadingElement(trimmed, line.Text, line.Index, true);
                return true;
            case ScreenplayElementType.Character:
                if (insideDialogueBlock)
                {
                    shouldBreakBlock = true;
                    return false;
                }

                var isDualDialogue = trimmed.EndsWith('^');
                var characterName = isDualDialogue ? trimmed[..^1].TrimEnd() : trimmed;
                element = new CharacterElement(characterName, line.Text, line.Index, isDualDialogue);
                return true;
            case ScreenplayElementType.Dialogue:
                element = new DialogueElement(
                    characterContext?.CharacterName ?? string.Empty,
                    Array.Empty<string>(),
                    new[] { line.Text },
                    trimmed,
                    line.Text,
                    line.Index,
                    line.Index,
                    characterContext?.IsDualDialogue ?? false);
                return true;
            case ScreenplayElementType.Parenthetical:
                element = new ParentheticalElement(
                    trimmed,
                    line.Text,
                    line.Index,
                    characterContext?.CharacterName,
                    characterContext?.IsDualDialogue ?? false);
                return true;
            case ScreenplayElementType.Transition:
                if (insideDialogueBlock)
                {
                    shouldBreakBlock = true;
                    return false;
                }

                element = new TransitionElement(trimmed, line.Text, line.Index, true);
                return true;
            default:
                return false;
        }
    }

    private static bool TryParseSection(SourceLine line, out SectionElement sectionElement)
    {
        sectionElement = default!;

        if (!LooksLikeSection(line.Text))
        {
            return false;
        }

        var match = SectionRegex.Match(line.Text);
        var sectionDepth = match.Groups[1].Value.Length;
        var text = ExtractId(match.Groups[2].Value.Trim(), out var id);
        sectionElement = new SectionElement(text, line.Text, line.Index, sectionDepth);
        if (id != Guid.Empty) sectionElement.Id = id;
        return true;
    }

    private static bool TryParseTitlePageField(SourceLine line, out TitlePageEntry titlePageEntry)
    {
        var trimmed = line.Text.Trim();
        titlePageEntry = default!;

        if (IdCommentRegex.IsMatch(trimmed))
        {
            return false;
        }

        var colonIndex = trimmed.IndexOf(':');
        if (colonIndex <= 0)
        {
            return false;
        }

        var label = trimmed[..colonIndex].Trim();
        var value = trimmed[(colonIndex + 1)..].Trim();

        var fieldType = label.ToLowerInvariant() switch
        {
            "title" => TitlePageFieldType.Title,
            "episode" => TitlePageFieldType.Episode,
            "author" => TitlePageFieldType.Author,
            "credit" => TitlePageFieldType.Credit,
            "draft date" => TitlePageFieldType.DraftDate,
            "contact" => TitlePageFieldType.Contact,
            "source" => TitlePageFieldType.Source,
            "revision" => TitlePageFieldType.Revision,
            "notes" => TitlePageFieldType.Notes,
            _ => TitlePageFieldType.Custom
        };

        titlePageEntry = new TitlePageEntry(fieldType, line.Text, value, line.Index, label);
        return true;
    }

    private static bool TryParseSynopsis(SourceLine line, int? currentSectionDepth, out SynopsisElement synopsisElement)
    {
        synopsisElement = default!;

        if (!LooksLikeSynopsis(line.Text))
        {
            return false;
        }

        var match = SynopsisRegex.Match(line.Text);
        var text = ExtractId(match.Groups[1].Value.Trim(), out var id);
        synopsisElement = new SynopsisElement(text, line.Text, line.Index, currentSectionDepth);
        if (id != Guid.Empty) synopsisElement.Id = id;
        return true;
    }

    private static bool TryParseTransition(SourceLine line, out TransitionElement transitionElement)
    {
        var trimmed = line.Text.Trim();
        transitionElement = default!;

        if (trimmed.Length == 0)
        {
            return false;
        }

        if (trimmed.EndsWith(':'))
        {
            transitionElement = new TransitionElement(trimmed, line.Text, line.Index, false);
            return true;
        }

        if (KnownTransitionRegex.IsMatch(trimmed) || (trimmed.EndsWith("TO:", StringComparison.OrdinalIgnoreCase) && TextAnalysis.IsUppercaseLike(trimmed)))
        {
            transitionElement = new TransitionElement(trimmed, line.Text, line.Index, false);
            return true;
        }

        return false;
    }

    private static bool TryParseSceneHeading(SourceLine line, out SceneHeadingElement sceneHeadingElement)
    {
        var trimmed = line.Text.Trim();
        sceneHeadingElement = default!;

        if (trimmed.Length == 0)
        {
            return false;
        }

        if (trimmed.StartsWith('.'))
        {
            var rawForcedText = trimmed[1..].TrimStart();
            if (rawForcedText.Length > 0)
            {
                var forcedText = ExtractId(rawForcedText, out var forcedId);
                sceneHeadingElement = new SceneHeadingElement(forcedText, line.Text, line.Index, true);
                if (forcedId != Guid.Empty) sceneHeadingElement.Id = forcedId;
                return true;
            }

            return false;
        }

        if (!LooksLikeSceneHeading(trimmed))
        {
            return false;
        }

        var text = ExtractId(trimmed, out var id);
        sceneHeadingElement = new SceneHeadingElement(text, line.Text, line.Index, false);
        if (id != Guid.Empty) sceneHeadingElement.Id = id;
        return true;
    }

    private static bool TryParseCharacterCue(
        SourceLine line,
        IReadOnlyList<SourceLine> lines,
        int index,
        out CharacterElement cueCharacterElement)
    {
        var trimmed = line.Text.Trim();
        cueCharacterElement = default!;

        if (trimmed.Length == 0 || trimmed[0] == '#' || trimmed[0] == '=')
        {
            return false;
        }

        if (LooksLikeSceneHeading(trimmed) || KnownTransitionRegex.IsMatch(trimmed))
        {
            return false;
        }

        var isDualDialogue = trimmed.EndsWith('^');
        if (isDualDialogue)
        {
            trimmed = trimmed[..^1].TrimEnd();
        }

        if (!LooksLikeCharacterCue(trimmed))
        {
            return false;
        }

        var nextNonBlankLine = PeekNextNonBlankLine(lines, index + 1);
        if (nextNonBlankLine is null)
        {
            return false;
        }

        var nextText = nextNonBlankLine.Value.Value.Text;
        var nextTrimmed = nextText.Trim();
        if (LooksLikeSceneHeading(nextTrimmed) ||
            KnownTransitionRegex.IsMatch(nextTrimmed) ||
            LooksLikeSection(nextText) ||
            LooksLikeSynopsis(nextText) ||
            LooksLikeCharacterCue(nextTrimmed) ||
            LooksLikeNote(nextTrimmed) ||
            LooksLikeBoneyard(nextTrimmed) ||
            LooksLikeCenteredText(nextTrimmed) ||
            LooksLikeLyrics(nextTrimmed))
        {
            return false;
        }

        cueCharacterElement = new CharacterElement(trimmed, line.Text, line.Index, isDualDialogue);
        return true;
    }

    private static bool LooksLikeSection(string trimmedLine)
    {
        return SectionRegex.IsMatch(trimmedLine);
    }

    private static bool LooksLikeSynopsis(string trimmedLine)
    {
        return SynopsisRegex.IsMatch(trimmedLine);
    }

    private static bool LooksLikeNote(string trimmedLine)
    {
        var trimmed = trimmedLine.TrimStart();
        return trimmed.StartsWith("[[", StringComparison.Ordinal);
    }

    private static bool LooksLikeBoneyard(string trimmedLine)
    {
        var trimmed = trimmedLine.TrimStart();
        return trimmed.StartsWith("/*", StringComparison.Ordinal);
    }

    private static bool LooksLikeCenteredText(string trimmedLine)
    {
        return trimmedLine.Length > 2
            && trimmedLine.StartsWith('>')
            && trimmedLine.EndsWith('<');
    }

    private static bool LooksLikeLyrics(string trimmedLine)
    {
        return trimmedLine.Length > 1 && trimmedLine[0] == '~';
    }

    private static bool LooksLikeSceneHeading(string trimmedLine)
    {
        if (trimmedLine.Length == 0)
        {
            return false;
        }

        var tokenEnd = trimmedLine.IndexOfAny([' ', '\t']);
        var token = tokenEnd < 0 ? trimmedLine.AsSpan() : trimmedLine.AsSpan(0, tokenEnd);
        return TextAnalysis.StartsWithAnyPrefix(token, SceneHeadingPrefixes);
    }

    private static bool LooksLikeCharacterCue(string trimmedLine)
    {
        if (trimmedLine.Length == 0)
        {
            return false;
        }

        if (trimmedLine.IndexOf(':') >= 0)
        {
            return false;
        }

        return TextAnalysis.IsLikelyCharacterCue(trimmedLine, 45, 6);
    }

    private static bool IsParentheticalLine(string trimmedLine)
    {
        return trimmedLine.Length > 2
            && trimmedLine[0] == '('
            && trimmedLine[^1] == ')';
    }

    private static bool TryParseNote(
        SourceLine line,
        IReadOnlyList<SourceLine> lines,
        int index,
        out NoteElement noteElement,
        out int nextIndex)
    {
        noteElement = default!;
        nextIndex = index;

        if (!LooksLikeNote(line.Text.Trim()))
        {
            return false;
        }

        var endIndex = FindClosingLine(lines, index, "]]");
        var rawText = BuildRawText(lines, index, endIndex);
        var text = ExtractId(FountainMarkup.ExtractNoteText(rawText), out var id);
        var closed = ContainsClosingMarker(lines, index, endIndex, "]]");

        noteElement = new NoteElement(text, rawText, line.Index, lines[endIndex].Index, closed);
        if (id != Guid.Empty) noteElement.Id = id;
        
        nextIndex = endIndex + 1;
        return true;
    }

    private static bool TryParseBoneyard(
        SourceLine line,
        IReadOnlyList<SourceLine> lines,
        int index,
        out BoneyardElement boneyardElement,
        out int nextIndex)
    {
        boneyardElement = default!;
        nextIndex = index;

        if (!LooksLikeBoneyard(line.Text.Trim()))
        {
            return false;
        }

        var endIndex = FindClosingLine(lines, index, "*/");
        var rawText = BuildRawText(lines, index, endIndex);
        var text = FountainMarkup.ExtractBoneyardText(rawText);
        var closed = ContainsClosingMarker(lines, index, endIndex, "*/");

        boneyardElement = new BoneyardElement(text, rawText, line.Index, lines[endIndex].Index, closed);
        nextIndex = endIndex + 1;
        return true;
    }

    private static bool TryParseCenteredText(SourceLine line, out CenteredTextElement centeredTextElement)
    {
        centeredTextElement = default!;

        var trimmed = line.Text.Trim();
        if (!LooksLikeCenteredText(trimmed))
        {
            return false;
        }

        var text = trimmed[1..^1].Trim();
        centeredTextElement = new CenteredTextElement(text, line.Text, line.Index);
        return true;
    }

    private static bool TryParseLyrics(SourceLine line, out LyricsElement lyricsElement)
    {
        lyricsElement = default!;

        var trimmed = line.Text.TrimStart();
        if (!LooksLikeLyrics(trimmed))
        {
            return false;
        }

        var text = trimmed[1..].TrimStart();
        lyricsElement = new LyricsElement(text, line.Text, line.Index);
        return true;
    }

    private static int FindClosingLine(IReadOnlyList<SourceLine> lines, int startIndex, string closingMarker)
    {
        for (var i = startIndex; i < lines.Count; i++)
        {
            if (lines[i].Text.Contains(closingMarker, StringComparison.Ordinal))
            {
                return i;
            }
        }

        return lines.Count - 1;
    }

    private static bool ContainsClosingMarker(IReadOnlyList<SourceLine> lines, int startIndex, int endIndex, string closingMarker)
    {
        for (var i = startIndex; i <= endIndex && i < lines.Count; i++)
        {
            if (lines[i].Text.Contains(closingMarker, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<SourceLine> ReadLines(string text)
    {
        if (text.Length == 0)
        {
            return Array.Empty<SourceLine>();
        }

        var lines = new List<SourceLine>(Math.Max(16, text.Length / 32));
        var start = 0;

        while (start <= text.Length)
        {
            var index = start;
            while (index < text.Length && text[index] != '\r' && text[index] != '\n')
            {
                index++;
            }

            var lineText = text[start..index];
            var lineEnding = string.Empty;

            if (index < text.Length)
            {
                if (text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                {
                    lineEnding = "\r\n";
                    index += 2;
                }
                else
                {
                    lineEnding = text[index].ToString();
                    index++;
                }
            }

            lines.Add(new SourceLine(lines.Count, lineText, lineEnding));

            if (index >= text.Length)
            {
                break;
            }

            start = index;
        }

        return lines;
    }

    private static string BuildRawText(IReadOnlyList<SourceLine> lines, int startIndex, int endIndex)
    {
        var builder = new StringBuilder(Math.Max(16, (endIndex - startIndex + 1) * 24));

        for (var i = startIndex; i <= endIndex; i++)
        {
            builder.Append(lines[i].Text);
            if (i < endIndex)
            {
                builder.Append(lines[i].LineEnding);
            }
        }

        return builder.ToString();
    }

    private static string BuildDisplayText(IReadOnlyList<SourceLine> lines, int startIndex, int endIndex)
    {
        var builder = new StringBuilder(Math.Max(16, (endIndex - startIndex + 1) * 18));

        for (var i = startIndex; i <= endIndex; i++)
        {
            if (i > startIndex)
            {
                builder.Append('\n');
            }

            builder.Append(lines[i].Text);
        }

        return builder.ToString();
    }

    private static NextNonBlankLine? PeekNextNonBlankLine(IReadOnlyList<SourceLine> lines, int startIndex)
    {
        for (var i = startIndex; i < lines.Count; i++)
        {
            if (lines[i].Text.Trim().Length == 0)
            {
                continue;
            }

            return new NextNonBlankLine(i, lines[i]);
        }

        return null;
    }

    private static readonly Regex IdCommentRegex = new(@"\s*\[\[id:([a-f\d\-]+)\]\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static string ExtractId(string text, out Guid id)
    {
        id = Guid.Empty;
        if (string.IsNullOrWhiteSpace(text)) return text;

        var match = IdCommentRegex.Match(text);
        if (match.Success)
        {
            if (Guid.TryParse(match.Groups[1].Value, out var parsedId))
            {
                id = parsedId;
            }
            return text.Replace(match.Value, "").Trim();
        }
        return text;
    }

    private readonly record struct SourceLine(int Index, string Text, string LineEnding);

    private readonly record struct NextNonBlankLine(int Index, SourceLine Line)
    {
        public SourceLine Value => Line;
    }
}
