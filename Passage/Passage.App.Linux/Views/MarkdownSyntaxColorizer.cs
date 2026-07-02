using System.Text.RegularExpressions;
using Avalonia.Media;
using AvaloniaEdit.Document;
using AvaloniaEdit.Rendering;

namespace Passage.App.Views;

/// <summary>
/// Live Markdown syntax highlighting for the AvaloniaEdit editor, active while the
/// app is in Markdown mode (Ctrl+M). Works line-by-line: block constructs (headings,
/// blockquotes, code fences) style the whole line, and inline constructs (bold,
/// italic, code spans, links) style just their span. Colours mirror the tone of the
/// Fountain colorizer so switching modes feels like the same editor.
/// </summary>
public sealed class MarkdownSyntaxColorizer : DocumentColorizingTransformer
{
    // Theme-aware brushes (see SyntaxTheme); constants are dark-theme fallbacks.
    private static IBrush HeadingBrush => SyntaxTheme.Brush("SyntaxSceneHeading", "#6797FF");
    private static IBrush QuoteBrush => SyntaxTheme.Brush("SyntaxParenthetical", "#7A7A7A");
    private static IBrush CodeBrush => SyntaxTheme.Brush("SyntaxSynopsis", "#FFB74D");
    private static IBrush ListMarkerBrush => SyntaxTheme.Brush("SyntaxSection", "#4FC3F7");
    private static IBrush LinkBrush => SyntaxTheme.Brush("SyntaxNote", "#81C784");

    private static readonly Regex HeadingRegex = new(@"^\s{0,3}#{1,6}\s", RegexOptions.Compiled);
    private static readonly Regex QuoteRegex = new(@"^\s{0,3}>", RegexOptions.Compiled);
    private static readonly Regex FenceRegex = new(@"^\s{0,3}(```|~~~)", RegexOptions.Compiled);
    private static readonly Regex ListMarkerRegex = new(@"^\s*(?:[-*+]|\d{1,3}[.)])\s", RegexOptions.Compiled);
    private static readonly Regex BoldRegex = new(@"(\*\*|__)(?!\s)(?:.+?)(?<!\s)\1", RegexOptions.Compiled);
    private static readonly Regex ItalicRegex = new(@"(?<![*_])([*_])(?![*_\s])(?:[^*_]+?)(?<!\s)\1(?![*_])", RegexOptions.Compiled);
    private static readonly Regex CodeSpanRegex = new(@"`[^`\n]+`", RegexOptions.Compiled);
    private static readonly Regex LinkRegex = new(@"!?\[[^\]\n]*\]\([^)\n]*\)", RegexOptions.Compiled);

    protected override void ColorizeLine(DocumentLine line)
    {
        if (line.Length == 0)
        {
            return;
        }

        var text = CurrentContext.Document.GetText(line.Offset, line.Length);

        if (HeadingRegex.IsMatch(text))
        {
            StyleSpan(line.Offset, line.EndOffset, HeadingBrush, FontStyle.Normal, FontWeight.Bold);
            return;
        }

        if (FenceRegex.IsMatch(text))
        {
            StyleSpan(line.Offset, line.EndOffset, CodeBrush, FontStyle.Normal, FontWeight.Normal);
            return;
        }

        if (QuoteRegex.IsMatch(text))
        {
            StyleSpan(line.Offset, line.EndOffset, QuoteBrush, FontStyle.Italic, FontWeight.Normal);
            return;
        }

        var listMarker = ListMarkerRegex.Match(text);
        if (listMarker.Success)
        {
            StyleSpan(line.Offset + listMarker.Index, line.Offset + listMarker.Index + listMarker.Length, ListMarkerBrush, FontStyle.Normal, FontWeight.Bold);
        }

        foreach (Match match in CodeSpanRegex.Matches(text))
        {
            StyleSpan(line.Offset + match.Index, line.Offset + match.Index + match.Length, CodeBrush, FontStyle.Normal, FontWeight.Normal);
        }

        foreach (Match match in LinkRegex.Matches(text))
        {
            StyleSpan(line.Offset + match.Index, line.Offset + match.Index + match.Length, LinkBrush, FontStyle.Normal, FontWeight.Normal);
        }

        foreach (Match match in BoldRegex.Matches(text))
        {
            StyleSpan(line.Offset + match.Index, line.Offset + match.Index + match.Length, null, FontStyle.Normal, FontWeight.Bold);
        }

        foreach (Match match in ItalicRegex.Matches(text))
        {
            StyleSpan(line.Offset + match.Index, line.Offset + match.Index + match.Length, null, FontStyle.Italic, FontWeight.Normal);
        }
    }

    private void StyleSpan(int startOffset, int endOffset, IBrush? brush, FontStyle style, FontWeight weight)
    {
        ChangeLinePart(startOffset, endOffset, element =>
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
}
