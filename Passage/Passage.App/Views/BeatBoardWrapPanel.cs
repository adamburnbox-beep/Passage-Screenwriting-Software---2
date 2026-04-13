using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Passage.Parser;

namespace Passage.App.Views;

public sealed class BeatBoardWrapPanel : Panel
{
    private const double MinActWidth = 180.0;
    private const double SequenceWidth = 200.0;
    private const double SceneWidth = 180.0;
    private const double HorizontalGap = 16.0;
    private const double VerticalGap = 24.0;
    private const double TrailingGutterWidth = 18.0;
    private const double BoardPadding = 32.0;
    private const double LanePadding = 16.0;

    private readonly List<Rect> _arrangedChildBounds = new();
    private readonly List<double> _actBoundariesY = new();
    private readonly List<(double X, double Y)> _sequenceBoundaries = new();
    private readonly List<Rect> _actLaneRects = new();
    private readonly List<Rect> _sequenceLaneRects = new();

    protected override Size MeasureOverride(Size availableSize)
    {
        var totalMeasuredHeight = 0.0;
        var maxMeasuredWidth = 0.0;

        // First, measure all children with their hierarchical widths
        foreach (UIElement child in InternalChildren)
        {
            if (child is null) continue;
            if (!TryGetScreenplayElement(child, out var element))
            {
                child.Measure(new Size(availableSize.Width, double.PositiveInfinity));
                continue;
            }

            double constrainedWidth = element.Level switch
            {
                0 => double.PositiveInfinity, // Acts grow based on title
                1 => SequenceWidth,
                _ => SceneWidth
            };

            child.Measure(new Size(constrainedWidth, double.PositiveInfinity));
        }

        // Structural Measurement Pass
        for (var i = 0; i < InternalChildren.Count; i++)
        {
            if (InternalChildren[i] is not UIElement actChild || !TryGetScreenplayElement(actChild, out var actElement))
            {
                if (InternalChildren[i] is UIElement nonElement)
                {
                    totalMeasuredHeight += nonElement.DesiredSize.Height + VerticalGap;
                }
                continue;
            }

            // Group into Acts (Vertical Spines)
            if (actElement.Level <= 0)
            {
                double currentActWidth = Math.Max(MinActWidth, actChild.DesiredSize.Width);
                double actTotalHeight = 0;
                double actMaxRowWidth = currentActWidth + HorizontalGap;
                int actEndIndex = i + 1;

                // Iterate through children belonging to this Act
                while (actEndIndex < InternalChildren.Count && 
                       (!TryGetScreenplayElement(InternalChildren[actEndIndex], out var nextElement) || nextElement.Level > 0))
                {
                    if (InternalChildren[actEndIndex] is not UIElement seqChild || !TryGetScreenplayElement(seqChild, out var seqElement))
                    {
                        actEndIndex++;
                        continue;
                    }

                    // Group into Sequence Rows
                    if (seqElement.Level == 1)
                    {
                        double rowX = currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        double rowHeight = seqChild.DesiredSize.Height;
                        int sceneIndex = actEndIndex + 1;

                        // Iterate through Scenes in this Sequence Row
                        while (sceneIndex < InternalChildren.Count && 
                               TryGetScreenplayElement(InternalChildren[sceneIndex], out var sceneElement) && 
                               sceneElement.Level >= 2)
                        {
                            var sceneChild = InternalChildren[sceneIndex];
                            rowHeight = Math.Max(rowHeight, sceneChild.DesiredSize.Height);
                            rowX += SceneWidth + HorizontalGap;
                            sceneIndex++;
                        }

                        actTotalHeight += rowHeight + VerticalGap;
                        actMaxRowWidth = Math.Max(actMaxRowWidth, rowX);
                        actEndIndex = sceneIndex;
                    }
                    else
                    {
                        actTotalHeight += seqChild.DesiredSize.Height + VerticalGap;
                        actEndIndex++;
                    }
                }

                actTotalHeight = Math.Max(actTotalHeight, actChild.DesiredSize.Height + VerticalGap);
                totalMeasuredHeight += actTotalHeight;
                maxMeasuredWidth = Math.Max(maxMeasuredWidth, actMaxRowWidth);
                i = actEndIndex - 1;
            }
            else
            {
                totalMeasuredHeight += actChild.DesiredSize.Height + VerticalGap;
                maxMeasuredWidth = Math.Max(maxMeasuredWidth, actChild.DesiredSize.Width);
            }
        }

        return new Size(maxMeasuredWidth + TrailingGutterWidth + (BoardPadding * 2), totalMeasuredHeight + (BoardPadding * 2));
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        EnsureArrangedChildBoundsCapacity();
        _actBoundariesY.Clear();
        _sequenceBoundaries.Clear();
        _actLaneRects.Clear();
        _sequenceLaneRects.Clear();
        var currentY = BoardPadding;

        for (var i = 0; i < InternalChildren.Count; i++)
        {
            if (InternalChildren[i] is not UIElement actChild || !TryGetScreenplayElement(actChild, out var actElement))
            {
                if (InternalChildren[i] is UIElement nonElement)
                {
                    var bounds = new Rect(BoardPadding, currentY, nonElement.DesiredSize.Width, nonElement.DesiredSize.Height);
                    nonElement.Arrange(bounds);
                    _arrangedChildBounds[i] = bounds;
                    currentY += bounds.Height + VerticalGap;
                }
                continue;
            }

            if (actElement.Level <= 0) // Act Spine
            {
                double currentActWidth = Math.Max(MinActWidth, actChild.DesiredSize.Width);
                double actTop = currentY;
                double actRunningHeight = 0;
                int actEndIndex = i + 1;
                
                // Track row boundaries within this Act for OnRender
                var internalRowBottoms = new List<double>();

                while (actEndIndex < InternalChildren.Count && 
                       (!TryGetScreenplayElement(InternalChildren[actEndIndex], out var nextElement) || nextElement.Level > 0))
                {
                    if (InternalChildren[actEndIndex] is not UIElement seqChild || !TryGetScreenplayElement(seqChild, out var seqElement))
                    {
                        actEndIndex++;
                        continue;
                    }

                    if (seqElement.Level == 1)
                    {
                        double rowTop = actTop + actRunningHeight;
                        double rowHeight = seqChild.DesiredSize.Height;
                        int sceneIndex = actEndIndex + 1;
                        while (sceneIndex < InternalChildren.Count && 
                               TryGetScreenplayElement(InternalChildren[sceneIndex], out var sceneElement) && 
                               sceneElement.Level >= 2)
                        {
                            var sceneChild = InternalChildren[sceneIndex];
                            rowHeight = Math.Max(rowHeight, sceneChild.DesiredSize.Height);
                            sceneIndex++;
                        }

                        var seqBounds = new Rect(BoardPadding + currentActWidth + HorizontalGap, rowTop, SequenceWidth, seqChild.DesiredSize.Height);
                        seqChild.Arrange(seqBounds);
                        _arrangedChildBounds[actEndIndex] = seqBounds;

                        // Track Sequence Lane Background
                        _sequenceLaneRects.Add(new Rect(BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop - LanePadding, finalSize.Width - BoardPadding - currentActWidth - (HorizontalGap / 2) - LanePadding, rowHeight + (LanePadding * 2)));

                        double currentSceneX = BoardPadding + currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        for (int k = actEndIndex + 1; k < sceneIndex; k++)
                        {
                            var sceneChild = InternalChildren[k];
                            var sceneBounds = new Rect(currentSceneX, rowTop, SceneWidth, sceneChild.DesiredSize.Height);
                            sceneChild.Arrange(sceneBounds);
                            _arrangedChildBounds[k] = sceneBounds;
                            currentSceneX += SceneWidth + HorizontalGap;
                        }

                        _sequenceBoundaries.Add((BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop + rowHeight + (VerticalGap / 2)));
                        actRunningHeight += rowHeight + VerticalGap;
                        internalRowBottoms.Add(actTop + actRunningHeight - (VerticalGap / 2));
                        actEndIndex = sceneIndex;
                    }
                    else
                    {
                        var bounds = new Rect(BoardPadding + currentActWidth + HorizontalGap, actTop + actRunningHeight, seqChild.DesiredSize.Width, seqChild.DesiredSize.Height);
                        seqChild.Arrange(bounds);
                        _arrangedChildBounds[actEndIndex] = bounds;
                        actRunningHeight += bounds.Height + VerticalGap;
                        internalRowBottoms.Add(actTop + actRunningHeight - (VerticalGap / 2));
                        actEndIndex++;
                    }
                }

                // to exactly the end of the content (minus trailing gap).
                double contentHeight = actRunningHeight > 0 ? actRunningHeight - VerticalGap : actChild.DesiredSize.Height;
                double cardHeight = Math.Max(actChild.DesiredSize.Height, contentHeight);
                var actBounds = new Rect(BoardPadding, actTop, currentActWidth, cardHeight);
                actChild.Arrange(actBounds);
                _arrangedChildBounds[i] = actBounds;

                // Add the absolute bottom of the Act block to the Act boundaries
                double actBlockBottom = actTop + Math.Max(actRunningHeight, actChild.DesiredSize.Height + VerticalGap);
                double actBoundaryY = actBlockBottom - (VerticalGap / 2);
                _actBoundariesY.Add(actBoundaryY);

                // Track Act Lane Background
                _actLaneRects.Add(new Rect(BoardPadding - LanePadding, actTop - LanePadding, finalSize.Width - (BoardPadding - LanePadding) * 2, actBoundaryY - actTop + (VerticalGap / 4) + LanePadding));

                currentY = actBlockBottom;
                i = actEndIndex - 1;
            }
            else // Lone element
            {
                var bounds = new Rect(BoardPadding, currentY, actChild.DesiredSize.Width, actChild.DesiredSize.Height);
                actChild.Arrange(bounds);
                _arrangedChildBounds[i] = bounds;
                currentY += bounds.Height + VerticalGap;
                _actBoundariesY.Add(currentY - (VerticalGap / 2));
            }
        }

        InvalidateVisual();
        return finalSize;
    }

    protected override void OnRender(DrawingContext dc)
    {
        base.OnRender(dc);

        var actLaneBrush = TryFindResource("BeatBoardActLaneBackground") as Brush;
        var seqLaneBrush = TryFindResource("BeatBoardSequenceLaneBackground") as Brush;
        var separatorBrush = TryFindResource("BeatBoardLaneSeparator") as Brush;

        // Draw Act Lane Backgrounds
        if (actLaneBrush != null)
        {
            foreach (var rect in _actLaneRects)
            {
                dc.DrawRoundedRectangle(actLaneBrush, null, rect, 8, 8);
            }
        }

        // Draw Sequence Lane Backgrounds
        if (seqLaneBrush != null)
        {
            foreach (var rect in _sequenceLaneRects)
            {
                dc.DrawRoundedRectangle(seqLaneBrush, null, rect, 8, 8);
            }
        }

        // Draw Separator Lines
        if (separatorBrush != null)
        {
            var separatorPen = new Pen(separatorBrush, 1.0);
            separatorPen.StartLineCap = separatorPen.EndLineCap = PenLineCap.Round;
            separatorPen.Freeze();

            // Glow pen (minimal spread and more transparent)
            var glowBrush = separatorBrush.Clone();
            glowBrush.Opacity *= 0.3; // Reduce opacity for the minimal glow layer
            var glowPen = new Pen(glowBrush, 2.5);
            glowPen.StartLineCap = glowPen.EndLineCap = PenLineCap.Round;
            glowPen.Freeze();

            // Act Boundaries
            foreach (var y in _actBoundariesY)
            {
                var startPoint = new Point(BoardPadding, y);
                var endPoint = new Point(ActualWidth - (BoardPadding - LanePadding), y);
                dc.DrawLine(glowPen, startPoint, endPoint);
                dc.DrawLine(separatorPen, startPoint, endPoint);
            }

            // Sequence Separators
            foreach (var boundary in _sequenceBoundaries)
            {
                var startPoint = new Point(boundary.X, boundary.Y);
                var endPoint = new Point(ActualWidth - (BoardPadding - LanePadding), boundary.Y);
                dc.DrawLine(glowPen, startPoint, endPoint);
                dc.DrawLine(separatorPen, startPoint, endPoint);
            }
        }
    }

    protected override void OnVisualChildrenChanged(DependencyObject visualAdded, DependencyObject visualRemoved)
    {
        base.OnVisualChildrenChanged(visualAdded, visualRemoved);
        InvalidateMeasure();
        InvalidateArrange();
        InvalidateVisual();
    }

    public static readonly DependencyProperty IsFullWidthProperty = DependencyProperty.RegisterAttached(
        "IsFullWidth",
        typeof(bool),
        typeof(BeatBoardWrapPanel),
        new FrameworkPropertyMetadata(
            false,
            FrameworkPropertyMetadataOptions.AffectsParentArrange | FrameworkPropertyMetadataOptions.AffectsParentMeasure));

    public static bool GetIsFullWidth(DependencyObject element)
    {
        ArgumentNullException.ThrowIfNull(element);
        return (bool)element.GetValue(IsFullWidthProperty);
    }

    public static void SetIsFullWidth(DependencyObject element, bool value)
    {
        ArgumentNullException.ThrowIfNull(element);
        element.SetValue(IsFullWidthProperty, value);
    }

    private double ResolveViewportWidth(Size availableSize)
    {
        if (!double.IsInfinity(availableSize.Width) && availableSize.Width > 0.0)
        {
            return availableSize.Width;
        }

        if (Parent is FrameworkElement parent && parent.ActualWidth > 0.0)
        {
            return parent.ActualWidth;
        }

        if (RenderSize.Width > 0.0)
        {
            return RenderSize.Width;
        }

        return 960.0;
    }

    private static bool TryGetScreenplayElement(UIElement child, out ScreenplayElement element)
    {
        if (child is FrameworkElement { DataContext: ScreenplayElement dataContextElement })
        {
            element = dataContextElement;
            return true;
        }

        element = null!;
        return false;
    }


    private void EnsureArrangedChildBoundsCapacity()
    {
        while (_arrangedChildBounds.Count < InternalChildren.Count)
        {
            _arrangedChildBounds.Add(Rect.Empty);
        }

        if (_arrangedChildBounds.Count > InternalChildren.Count)
        {
            _arrangedChildBounds.RemoveRange(InternalChildren.Count, _arrangedChildBounds.Count - InternalChildren.Count);
        }

        for (var index = 0; index < _arrangedChildBounds.Count; index++)
        {
            _arrangedChildBounds[index] = Rect.Empty;
        }
    }
}
