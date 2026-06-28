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
    private readonly FountainLineClassifier _classifier;

    public FountainSyntaxColorizer(FountainLineClassifier classifier)
    {
        _classifier = classifier;
    }

    private static readonly IBrush SceneHeadingBrush = new SolidColorBrush(Color.Parse("#6797FF"));
    private static readonly IBrush CharacterBrush = new SolidColorBrush(Color.Parse("#A23B72"));
    private static readonly IBrush DialogueBrush = new SolidColorBrush(Color.Parse("#CEB2C9"));
    private static readonly IBrush ParentheticalBrush = new SolidColorBrush(Color.Parse("#7A7A7A"));
    private static readonly IBrush TransitionBrush = new SolidColorBrush(Color.Parse("#6B4FA0"));
    private static readonly IBrush SectionBrush = new SolidColorBrush(Color.Parse("#4FC3F7"));
    private static readonly IBrush SynopsisBrush = new SolidColorBrush(Color.Parse("#FFB74D"));
    private static readonly IBrush NoteBrush = new SolidColorBrush(Color.Parse("#81C784"));

    protected override void ColorizeLine(DocumentLine line)
    {
        if (line.Length == 0)
        {
            return;
        }

        var type = _classifier.Classify(CurrentContext.Document, line);

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
