using System;
using System.Collections.Generic;
using System.Linq;
using AvaloniaEdit.Document;
using Passage.Core;
using Passage.Parser;

namespace Passage.App.Views;

/// <summary>
/// Classifies each editor line into a screenplay element type, reusing the same
/// parser the rest of the app relies on. The result is cached per document version
/// so the editor's live colouring (<see cref="FountainSyntaxColorizer"/>) and live
/// indentation (<see cref="FountainIndentationGenerator"/>) can share a single parse
/// pass instead of each re-parsing the whole document on every redraw.
/// </summary>
public sealed class FountainLineClassifier
{
    // Heuristic limits shared with the Windows editor's live classification.
    private const int CharacterCueMaxLength = 45;
    private const int CharacterCueMaxWords = 6;

    private readonly FountainParser _parser = new();
    private Dictionary<int, ScreenplayElementType> _lineTypes = new();
    private ITextSourceVersion? _cachedVersion;

    /// <summary>
    /// Resolves the element type for a single line. Lines the full parser was able
    /// to classify come straight from the cached map; anything it couldn't classify
    /// yet (for example a character cue that has no dialogue typed beneath it) falls
    /// back to the same live heuristics the Windows editor uses, so formatting lights
    /// up as the user types rather than only once a block is complete.
    /// </summary>
    public ScreenplayElementType Classify(TextDocument document, DocumentLine line)
    {
        EnsureLineTypes(document);

        if (_lineTypes.TryGetValue(line.LineNumber, out var mappedType))
        {
            return mappedType;
        }

        var lineText = document.GetText(line.Offset, line.Length);
        var trimmed = lineText.Trim();
        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        if (TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()))
        {
            return ScreenplayElementType.SceneHeading;
        }

        if (trimmed.StartsWith('(') && trimmed.EndsWith(')'))
        {
            return ScreenplayElementType.Parenthetical;
        }

        return TextAnalysis.IsLiveCharacterCueCandidate(lineText.AsSpan(), CharacterCueMaxLength, CharacterCueMaxWords)
            ? ScreenplayElementType.Character
            : ScreenplayElementType.Action;
    }

    private void EnsureLineTypes(TextDocument document)
    {
        var version = document.Version;
        if (version is not null && ReferenceEquals(version, _cachedVersion))
        {
            return;
        }

        _cachedVersion = version;

        var map = new Dictionary<int, ScreenplayElementType>();
        try
        {
            var parsed = _parser.Parse(document.Text, null);

            // Parenthetical lines live inside a dialogue element's line range, and the
            // dialogue element is emitted after them. Track them so the dialogue span
            // doesn't overwrite their classification (mirrors the Windows lookup).
            var parentheticalLines = parsed.Elements
                .OfType<ParentheticalElement>()
                .Select(element => element.LineNumber)
                .ToHashSet();

            foreach (var element in parsed.Elements)
            {
                if (element is DialogueElement dialogue)
                {
                    for (var lineNumber = dialogue.StartLine; lineNumber <= dialogue.EndLine; lineNumber++)
                    {
                        if (parentheticalLines.Contains(lineNumber))
                        {
                            continue;
                        }

                        map[lineNumber] = ScreenplayElementType.Dialogue;
                    }

                    continue;
                }

                for (var lineNumber = element.StartLine; lineNumber <= element.EndLine; lineNumber++)
                {
                    map[lineNumber] = element.Type;
                }
            }
        }
        catch
        {
            // Parsing is best-effort; on any failure we simply skip classification.
        }

        _lineTypes = map;
    }
}
