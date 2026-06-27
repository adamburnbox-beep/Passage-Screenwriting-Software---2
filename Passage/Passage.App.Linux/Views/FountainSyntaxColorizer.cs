using System.Collections.Generic;
using System.Linq;
using Avalonia.Media;
using AvaloniaEdit.Document;
using AvaloniaEdit.Rendering;
using Passage.Core;
using Passage.Parser;

namespace Passage.App.Views;

/// <summary>
/// Live Fountain syntax highlighting for the AvaloniaEdit editor. Each document
/// line is classified into a screenplay element type (reusing the same parser the
/// rest of the app relies on) and rendered with a distinct colour / weight / style
/// so the editor visibly "follows" Fountain syntax as the user types, matching the
/// formatted feel of the Windows version.
/// </summary>
public sealed class FountainSyntaxColorizer : DocumentColorizingTransformer
{
    // Heuristic limits shared with the Windows editor's live classification.
    private const int CharacterCueMaxLength = 45;
    private const int CharacterCueMaxWords = 6;

    private readonly FountainParser _parser = new();
    private Dictionary<int, ScreenplayElementType> _lineTypes = new();
    private ITextSourceVersion? _cachedVersion;

    private static readonly IBrush SceneHeadingBrush = new SolidColorBrush(Color.Parse("#2C5AA0"));
    private static readonly IBrush CharacterBrush = new SolidColorBrush(Color.Parse("#A23B72"));
    private static readonly IBrush DialogueBrush = new SolidColorBrush(Color.Parse("#2F6F6A"));
    private static readonly IBrush ParentheticalBrush = new SolidColorBrush(Color.Parse("#7A7A7A"));
    private static readonly IBrush TransitionBrush = new SolidColorBrush(Color.Parse("#6B4FA0"));
    private static readonly IBrush SectionBrush = new SolidColorBrush(Color.Parse("#4A7A3A"));
    private static readonly IBrush SynopsisBrush = new SolidColorBrush(Color.Parse("#6E8A4A"));
    private static readonly IBrush NoteBrush = new SolidColorBrush(Color.Parse("#999999"));

    protected override void ColorizeLine(DocumentLine line)
    {
        if (line.Length == 0)
        {
            return;
        }

        var document = CurrentContext.Document;
        EnsureLineTypes(document);

        var type = ResolveLineType(document, line);

        var (brush, style, weight) = ResolveStyle(type);
        if (brush is null && style == FontStyle.Normal && weight == FontWeight.Normal)
        {
            return;
        }

        ChangeLinePart(line.Offset, line.EndOffset, element =>
        {
            if (brush is not null)
            {
                element.TextRunProperties.SetForegroundBrush(brush);
            }

            if (weight != FontWeight.Normal || style != FontStyle.Normal)
            {
                var typeface = element.TextRunProperties.Typeface;
                element.TextRunProperties.SetTypeface(new Typeface(typeface.FontFamily, style, weight));
            }
        });
    }

    /// <summary>
    /// Resolves the element type for a single line. Lines the full parser was able
    /// to classify come straight from the cached map; anything it couldn't classify
    /// yet (for example a character cue that has no dialogue typed beneath it) falls
    /// back to the same live heuristics the Windows editor uses, so syntax lights up
    /// as the user types rather than only once a block is complete.
    /// </summary>
    private ScreenplayElementType ResolveLineType(TextDocument document, DocumentLine line)
    {
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
            // Parsing is best-effort; on any failure we simply skip highlighting.
        }

        _lineTypes = map;
    }

    private static (IBrush? Brush, FontStyle Style, FontWeight Weight) ResolveStyle(ScreenplayElementType type)
    {
        return type switch
        {
            ScreenplayElementType.SceneHeading => (SceneHeadingBrush, FontStyle.Normal, FontWeight.Bold),
            ScreenplayElementType.Character => (CharacterBrush, FontStyle.Normal, FontWeight.Bold),
            ScreenplayElementType.Dialogue => (DialogueBrush, FontStyle.Normal, FontWeight.Normal),
            ScreenplayElementType.Parenthetical => (ParentheticalBrush, FontStyle.Italic, FontWeight.Normal),
            ScreenplayElementType.Transition => (TransitionBrush, FontStyle.Normal, FontWeight.Bold),
            ScreenplayElementType.Section => (SectionBrush, FontStyle.Normal, FontWeight.Bold),
            ScreenplayElementType.Synopsis => (SynopsisBrush, FontStyle.Italic, FontWeight.Normal),
            ScreenplayElementType.Note => (NoteBrush, FontStyle.Italic, FontWeight.Normal),
            ScreenplayElementType.Boneyard => (NoteBrush, FontStyle.Italic, FontWeight.Normal),
            ScreenplayElementType.CenteredText => (null, FontStyle.Normal, FontWeight.Bold),
            ScreenplayElementType.Lyrics => (null, FontStyle.Italic, FontWeight.Normal),
            _ => (null, FontStyle.Normal, FontWeight.Normal)
        };
    }
}
