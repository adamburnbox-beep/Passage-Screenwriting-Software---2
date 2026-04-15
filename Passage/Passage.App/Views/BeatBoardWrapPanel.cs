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
            if (InternalChildren[i] is not UIElement currentChild || !TryGetScreenplayElement(currentChild, out var currentElement))
            {
                if (InternalChildren[i] is UIElement nonElement)
                {
                    totalMeasuredHeight += nonElement.DesiredSize.Height + VerticalGap;
                }
                continue;
            }

            // Group into Acts (Vertical Spines)
            if (currentElement.Level <= 0)
            {
                double currentActWidth = Math.Max(MinActWidth, currentChild.DesiredSize.Width);
                double actTotalHeight = 0;
                double actMaxRowWidth = currentActWidth + HorizontalGap;
                int actEndIndex = i + 1;

                // Iterate through children belonging to this Act
                while (actEndIndex < InternalChildren.Count && 
                       (!TryGetScreenplayElement(InternalChildren[actEndIndex], out var nextElement) || nextElement.Level > 0))
                {
                    if (InternalChildren[actEndIndex] is not UIElement rowChild || !TryGetScreenplayElement(rowChild, out var rowElement))
                    {
                        actEndIndex++;
                        continue;
                    }

                    // Group into Rows
                    double rowX;
                    double rowHeight;
                    int nextIndex;
                    
                    if (rowElement.Level == 1) // Sequence-started row
                    {
                        rowX = currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        rowHeight = rowChild.DesiredSize.Height;
                        nextIndex = actEndIndex + 1;
                    }
                    else // Scene-started row (inside Act)
                    {
                        rowX = currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        rowHeight = 0; // Will be set by scenes
                        nextIndex = actEndIndex;
                    }

                    // Collect Scenes
                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        var sceneChild = InternalChildren[nextIndex];
                        rowHeight = Math.Max(rowHeight, sceneChild.DesiredSize.Height);
                        rowX += SceneWidth + HorizontalGap;
                        nextIndex++;
                    }

                    actTotalHeight += rowHeight + VerticalGap;
                    actMaxRowWidth = Math.Max(actMaxRowWidth, rowX);
                    actEndIndex = nextIndex;
                }

                actTotalHeight = Math.Max(actTotalHeight, currentChild.DesiredSize.Height + VerticalGap);
                totalMeasuredHeight += actTotalHeight;
                maxMeasuredWidth = Math.Max(maxMeasuredWidth, actMaxRowWidth);
                i = actEndIndex - 1;
            }
            else
            {
                // Lone Group (No Act)
                double rowWidth = BoardPadding;
                double rowHeight;
                int nextIndex;

                if (currentElement.Level == 1) // Sequence-started lone row
                {
                    rowWidth += MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    rowHeight = currentChild.DesiredSize.Height;
                    nextIndex = i + 1;
                }
                else // Scene-started lone row
                {
                    rowWidth += MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    rowHeight = 0;
                    nextIndex = i;
                }

                // Collect Scenes
                while (nextIndex < InternalChildren.Count && 
                       TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                       sceneElement.Level >= 2)
                {
                    var sceneChild = InternalChildren[nextIndex];
                    rowHeight = Math.Max(rowHeight, sceneChild.DesiredSize.Height);
                    rowWidth += SceneWidth + HorizontalGap;
                    nextIndex++;
                }

                totalMeasuredHeight += rowHeight + VerticalGap;
                maxMeasuredWidth = Math.Max(maxMeasuredWidth, rowWidth);
                i = nextIndex - 1;
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
            if (InternalChildren[i] is not UIElement currentChild || !TryGetScreenplayElement(currentChild, out var currentElement))
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

            if (currentElement.Level <= 0) // Act Spine
            {
                double currentActWidth = Math.Max(MinActWidth, currentChild.DesiredSize.Width);
                double actTop = currentY;
                double actRunningHeight = 0;
                int actEndIndex = i + 1;

                while (actEndIndex < InternalChildren.Count && 
                       (!TryGetScreenplayElement(InternalChildren[actEndIndex], out var nextElement) || nextElement.Level > 0))
                {
                    if (InternalChildren[actEndIndex] is not UIElement rowChild || !TryGetScreenplayElement(rowChild, out var rowElement))
                    {
                        actEndIndex++;
                        continue;
                    }

                    double rowTop = actTop + actRunningHeight;
                    double rowHeight;
                    int nextIndex;

                    if (rowElement.Level == 1) // Sequence row
                    {
                        rowHeight = rowChild.DesiredSize.Height;
                        nextIndex = actEndIndex + 1;

                        while (nextIndex < InternalChildren.Count && 
                               TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                               sceneElement.Level >= 2)
                        {
                            rowHeight = Math.Max(rowHeight, InternalChildren[nextIndex].DesiredSize.Height);
                            nextIndex++;
                        }

                        var seqBounds = new Rect(BoardPadding + currentActWidth + HorizontalGap, rowTop, SequenceWidth, rowChild.DesiredSize.Height);
                        rowChild.Arrange(seqBounds);
                        _arrangedChildBounds[actEndIndex] = seqBounds;

                        // Track Sequence Lane Background
                        _sequenceLaneRects.Add(new Rect(BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop - LanePadding, finalSize.Width - BoardPadding - currentActWidth - (HorizontalGap / 2) - LanePadding, rowHeight + (LanePadding * 2)));

                        double currentSceneX = BoardPadding + currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        for (int k = actEndIndex + 1; k < nextIndex; k++)
                        {
                            var sceneChild = InternalChildren[k];
                            var sceneBounds = new Rect(currentSceneX, rowTop, SceneWidth, sceneChild.DesiredSize.Height);
                            sceneChild.Arrange(sceneBounds);
                            _arrangedChildBounds[k] = sceneBounds;
                            currentSceneX += SceneWidth + HorizontalGap;
                        }

                        _sequenceBoundaries.Add((BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop + rowHeight + (VerticalGap / 2)));
                        actRunningHeight += rowHeight + VerticalGap;
                        actEndIndex = nextIndex;
                    }
                    else // Scene row (inside Act)
                    {
                        rowHeight = 0;
                        nextIndex = actEndIndex;

                         while (nextIndex < InternalChildren.Count && 
                                TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                                sceneElement.Level >= 2)
                        {
                            rowHeight = Math.Max(rowHeight, InternalChildren[nextIndex].DesiredSize.Height);
                            nextIndex++;
                        }

                        double currentSceneX = BoardPadding + currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        for (int k = actEndIndex; k < nextIndex; k++)
                        {
                            var sceneChild = InternalChildren[k];
                            var sceneBounds = new Rect(currentSceneX, rowTop, SceneWidth, sceneChild.DesiredSize.Height);
                            sceneChild.Arrange(sceneBounds);
                            _arrangedChildBounds[k] = sceneBounds;
                            currentSceneX += SceneWidth + HorizontalGap;
                        }

                        actRunningHeight += rowHeight + VerticalGap;
                        actEndIndex = nextIndex;
                    }
                }

                double contentHeight = actRunningHeight > 0 ? actRunningHeight - VerticalGap : currentChild.DesiredSize.Height;
                double cardHeight = Math.Max(currentChild.DesiredSize.Height, contentHeight);
                var actBounds = new Rect(BoardPadding, actTop, currentActWidth, cardHeight);
                currentChild.Arrange(actBounds);
                _arrangedChildBounds[i] = actBounds;

                double actBlockBottom = actTop + Math.Max(actRunningHeight, currentChild.DesiredSize.Height + VerticalGap);
                double actBoundaryY = actBlockBottom - (VerticalGap / 2);
                _actBoundariesY.Add(actBoundaryY);

                _actLaneRects.Add(new Rect(BoardPadding - LanePadding, actTop - LanePadding, finalSize.Width - (BoardPadding - LanePadding) * 2, actBoundaryY - actTop + (VerticalGap / 4) + LanePadding));

                currentY = actBlockBottom;
                i = actEndIndex - 1;
            }
            else // Lone Group (No Act)
            {
                double rowTop = currentY;
                double rowHeight;
                int nextIndex;

                if (currentElement.Level == 1) // Sequence row
                {
                    rowHeight = currentChild.DesiredSize.Height;
                    nextIndex = i + 1;

                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        rowHeight = Math.Max(rowHeight, InternalChildren[nextIndex].DesiredSize.Height);
                        nextIndex++;
                    }

                    var seqBounds = new Rect(BoardPadding + MinActWidth + HorizontalGap, rowTop, SequenceWidth, currentChild.DesiredSize.Height);
                    currentChild.Arrange(seqBounds);
                    _arrangedChildBounds[i] = seqBounds;

                    _sequenceLaneRects.Add(new Rect(BoardPadding + MinActWidth + (HorizontalGap / 2), rowTop - LanePadding, finalSize.Width - BoardPadding - MinActWidth - (HorizontalGap / 2) - LanePadding, rowHeight + (LanePadding * 2)));

                    double currentSceneX = BoardPadding + MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    for (int k = i + 1; k < nextIndex; k++)
                    {
                        var sceneChild = InternalChildren[k];
                        var sceneBounds = new Rect(currentSceneX, rowTop, SceneWidth, sceneChild.DesiredSize.Height);
                        sceneChild.Arrange(sceneBounds);
                        _arrangedChildBounds[k] = sceneBounds;
                        currentSceneX += SceneWidth + HorizontalGap;
                    }

                    _sequenceBoundaries.Add((BoardPadding + MinActWidth + (HorizontalGap / 2), rowTop + rowHeight + (VerticalGap / 2)));
                }
                else // Scene row
                {
                    rowHeight = 0;
                    nextIndex = i;

                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        rowHeight = Math.Max(rowHeight, InternalChildren[nextIndex].DesiredSize.Height);
                        nextIndex++;
                    }

                    double currentSceneX = BoardPadding + MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    for (int k = i; k < nextIndex; k++)
                    {
                        var sceneChild = InternalChildren[k];
                        var sceneBounds = new Rect(currentSceneX, rowTop, SceneWidth, sceneChild.DesiredSize.Height);
                        sceneChild.Arrange(sceneBounds);
                        _arrangedChildBounds[k] = sceneBounds;
                        currentSceneX += SceneWidth + HorizontalGap;
                    }
                    
                    _actBoundariesY.Add(rowTop + rowHeight + (VerticalGap / 2));
                }

                currentY += rowHeight + VerticalGap;
                i = nextIndex - 1;
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

        // Draw Act Lane Backgrounds (Removed as per request)
        /*
        if (actLaneBrush != null)
        {
            foreach (var rect in _actLaneRects)
            {
                dc.DrawRoundedRectangle(actLaneBrush, null, rect, 8, 8);
            }
        }
        */

        // Draw Sequence Lane Backgrounds (Removed as per request)
        /*
        if (seqLaneBrush != null)
        {
            foreach (var rect in _sequenceLaneRects)
            {
                dc.DrawRoundedRectangle(seqLaneBrush, null, rect, 8, 8);
            }
        }
        */

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
