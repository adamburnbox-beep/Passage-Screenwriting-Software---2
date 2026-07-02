using System;
using System.Globalization;
using Avalonia;
using Avalonia.Media;
using AvaloniaEdit.Rendering;

namespace Passage.App.Views;

/// <summary>
/// Draws page-break rules across the script editor so the single text surface
/// reads as consecutive screenplay pages: every 55 lines (the industry-standard
/// page) a dashed rule snapped to the nearest line boundary is drawn, with a
/// "PAGE N" pill at the right margin. Purely visual — the document itself is
/// untouched, and boundaries are the same line-count approximation the page
/// estimator uses; the Page Preview tab remains the exact print layout.
/// </summary>
public sealed class ScreenplayPageRuler : IBackgroundRenderer
{
    private const int LinesPerPage = 55;

    public KnownLayer Layer => KnownLayer.Background;

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (textView.Document is null)
        {
            return;
        }

        var lineHeight = textView.DefaultLineHeight;
        if (lineHeight <= 0 || double.IsNaN(lineHeight))
        {
            return;
        }

        var pageBodyHeight = LinesPerPage * lineHeight;
        var documentHeight = textView.DocumentHeight;
        if (documentHeight <= pageBodyHeight)
        {
            return;
        }

        var chipBrush = SyntaxTheme.Brush("SurfaceBackground", "#111111");
        var lineBrush = SyntaxTheme.Brush("SurfaceBorder", "#242424");
        var labelBrush = SyntaxTheme.Brush("MutedText", "#70706C");
        var rulePen = new Pen(lineBrush, 1, new DashStyle(new double[] { 4, 4 }, 0));
        var chipPen = new Pen(lineBrush, 1);

        var viewWidth = textView.Bounds.Width;
        var viewHeight = textView.Bounds.Height;
        var scrollY = textView.ScrollOffset.Y;
        var totalPages = (int)Math.Ceiling(documentHeight / pageBodyHeight);

        for (var page = 1; page < totalPages; page++)
        {
            var documentY = page * pageBodyHeight;
            var y = documentY - scrollY;
            if (y < -lineHeight)
            {
                continue;
            }

            if (y > viewHeight + lineHeight)
            {
                break;
            }

            // Snap the rule to the top of the visual line at the boundary so it
            // sits between two text lines instead of cutting through glyphs.
            var visualLine = textView.GetVisualLineFromVisualTop(documentY);
            if (visualLine != null)
            {
                y = visualLine.VisualTop - scrollY;
            }

            drawingContext.DrawLine(rulePen, new Point(0, y), new Point(viewWidth, y));

            var label = new FormattedText(
                $"PAGE {page + 1}",
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(FontFamily.Default, FontStyle.Normal, FontWeight.SemiBold),
                9.5,
                labelBrush);

            var chipWidth = label.Width + 16;
            var chipHeight = label.Height + 6;
            var chipRect = new Rect(viewWidth - chipWidth - 6, y - chipHeight / 2, chipWidth, chipHeight);
            drawingContext.DrawRectangle(chipBrush, chipPen, new RoundedRect(chipRect, chipHeight / 2));
            drawingContext.DrawText(label, new Point(chipRect.X + 8, chipRect.Y + 3));
        }
    }
}
