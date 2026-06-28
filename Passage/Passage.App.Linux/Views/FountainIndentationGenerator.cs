using System;
using System.Globalization;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using AvaloniaEdit.Document;
using AvaloniaEdit.Rendering;
using Passage.Parser;

namespace Passage.App.Views;

/// <summary>
/// Live screenplay indentation for the AvaloniaEdit editor. AvaloniaEdit is a
/// plain-text control, so — unlike the Windows WPF editor's FlowDocument paragraphs —
/// it has no notion of per-line margins or alignment. This generator recreates that
/// behaviour by injecting a zero-document-length, fixed-pixel-width "blank space"
/// element at the start of each line, sized so the line's visible text lands at the
/// industry-standard screenplay position. The spacer consumes no document characters,
/// so the underlying Fountain text — and every caret/offset calculation the rest of
/// the app performs — is completely unaffected.
///
/// Offsets are measured in real pixels at 96 DPI (the page is a fixed 8.5" / 816px
/// sheet, see EditorDocumentHost in MainWindow.axaml, whose 1.5" left + 1.0" right
/// padding establishes the screenplay margins). Working in pixels rather than
/// character columns keeps the layout correct even when the editor falls back from
/// the monospace screenplay font to a proportional one, where a space glyph is far
/// narrower than a typical character.
/// </summary>
public sealed class FountainIndentationGenerator : VisualLineElementGenerator
{
    private const double Dpi = 96.0;

    // Indents measured from the left text margin (where action / scene headings sit),
    // matching standard US screenplay layout:
    //   Dialogue      2.5" from page edge → 1.0" past the 1.5" left margin
    //   Parenthetical 3.0" from page edge → 1.5" past the left margin
    //   Character     3.7" from page edge → 2.2" past the left margin
    private const double DialogueIndent = 1.0 * Dpi;
    private const double ParentheticalIndent = 1.5 * Dpi;
    private const double CharacterIndent = 2.2 * Dpi;

    // The text block is 6.0" wide (8.5" page − 1.5" left − 1.0" right). Used only as a
    // fallback before the editor has been laid out; once it has, we read the real
    // width (which also accounts for the scrollbar) from the TextView.
    private const double FallbackTextWidth = 6.0 * Dpi;

    // Anything below this is treated as "no indent" so we never emit a useless spacer.
    private const double MinimumIndent = 0.5;

    // For right-aligned / centred lines the spacer plus the text would otherwise fill
    // the text block to its exact width, and word wrap breaks a line the instant its
    // content reaches (>=) the available width — pushing the last word onto a second
    // line. Holding the text a few pixels short of the margin keeps it on one line; the
    // gap is well under a hundredth of an inch and not visible.
    private const double EdgeGuard = 8.0;

    private readonly FountainLineClassifier _classifier;

    public FountainIndentationGenerator(FountainLineClassifier classifier)
    {
        _classifier = classifier;
    }

    public override int GetFirstInterestedOffset(int startOffset)
    {
        var document = CurrentContext.Document;
        var line = document.GetLineByOffset(startOffset);

        // Only insert a spacer at the very start of a non-empty line. Returning the
        // line start exactly once is what keeps the construction loop from looping:
        // after the zero-length spacer is built the loop advances past the start
        // offset, so this method is never re-interested in the same position.
        if (startOffset != line.Offset || line.Length == 0)
        {
            return -1;
        }

        return GetIndentWidth(document, line) >= MinimumIndent ? line.Offset : -1;
    }

    public override VisualLineElement? ConstructElement(int offset)
    {
        var document = CurrentContext.Document;
        var line = document.GetLineByOffset(offset);
        var width = GetIndentWidth(document, line);
        return width >= MinimumIndent ? new BlankSpaceElement(width) : null;
    }

    private double GetIndentWidth(TextDocument document, DocumentLine line)
    {
        return _classifier.Classify(document, line) switch
        {
            ScreenplayElementType.Dialogue => DialogueIndent,
            ScreenplayElementType.Parenthetical => ParentheticalIndent,
            ScreenplayElementType.Character => CharacterIndent,
            // Transitions are right-aligned to the right margin; centred text is
            // centred within the text block. Both depend on the line's rendered width.
            ScreenplayElementType.Transition => RightAlignIndent(document, line),
            ScreenplayElementType.CenteredText => CenterIndent(document, line),
            _ => 0.0
        };
    }

    private double RightAlignIndent(TextDocument document, DocumentLine line)
    {
        return Math.Max(0.0, AvailableWidth() - MeasureLineWidth(document, line) - EdgeGuard);
    }

    private double CenterIndent(TextDocument document, DocumentLine line)
    {
        return Math.Max(0.0, (AvailableWidth() - EdgeGuard - MeasureLineWidth(document, line)) / 2.0);
    }

    private double AvailableWidth()
    {
        var width = CurrentContext.TextView.Bounds.Width;
        return width > 0.0 ? width : FallbackTextWidth;
    }

    private double MeasureLineWidth(TextDocument document, DocumentLine line)
    {
        var text = document.GetText(line.Offset, line.Length);
        var properties = CurrentContext.GlobalTextRunProperties;
        var formatted = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            properties.Typeface,
            properties.FontRenderingEmSize,
            foreground: null);

        return formatted.Width;
    }

    /// <summary>
    /// A blank run of an exact pixel width that occupies a single visual column but no
    /// document characters, used purely to shift the line's real text to its screenplay
    /// position. Caret navigation skips over it so the indent never becomes a place the
    /// cursor can land.
    /// </summary>
    private sealed class BlankSpaceElement : VisualLineElement
    {
        private readonly double _width;

        public BlankSpaceElement(double width)
            : base(1, 0)
        {
            _width = width;
        }

        public override TextRun CreateTextRun(int startVisualColumn, ITextRunConstructionContext context)
        {
            return new FixedWidthRun(_width, TextRunProperties);
        }

        public override bool IsWhitespace(int visualColumn) => true;

        public override int GetNextCaretPosition(int visualColumn, AvaloniaEdit.Document.LogicalDirection direction, CaretPositioningMode mode)
        {
            // The spacer is not a caret stop; navigation falls through to the real text.
            return -1;
        }
    }

    private sealed class FixedWidthRun : DrawableTextRun
    {
        private readonly double _width;

        public FixedWidthRun(double width, TextRunProperties? properties)
        {
            _width = width;
            Properties = properties;
        }

        public override int Length => 1;

        public override TextRunProperties? Properties { get; }

        public override Size Size => new(_width, 0.0);

        public override double Baseline => 0.0;

        public override void Draw(DrawingContext drawingContext, Point origin)
        {
            // Intentionally blank — the run is pure horizontal spacing.
        }
    }
}
