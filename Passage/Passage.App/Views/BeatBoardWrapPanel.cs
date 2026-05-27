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
        var viewportWidth = ResolveViewportWidth(availableSize);

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
                    double startX = currentActWidth + SequenceWidth + (HorizontalGap * 2);
                    double currentSceneX = startX;
                    double rowHeight;
                    int nextIndex;
                    double sequenceTotalHeight = 0.0;
                    double sceneRowHeight = 0.0;
                    int sceneCountInRow = 0;
                    
                    if (rowElement.Level == 1) // Sequence-started row
                    {
                        rowHeight = rowChild.DesiredSize.Height;
                        nextIndex = actEndIndex + 1;
                        actMaxRowWidth = Math.Max(actMaxRowWidth, currentActWidth + HorizontalGap + SequenceWidth);
                    }
                    else // Scene-started row (inside Act)
                    {
                        rowHeight = 0; // Will be set by scenes
                        nextIndex = actEndIndex;
                    }

                    // Collect and wrap Scenes
                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        var sceneChild = InternalChildren[nextIndex];
                        
                        // Check if we need to wrap
                        if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                        {
                            sequenceTotalHeight += Math.Max(sceneRowHeight, rowHeight) + VerticalGap;
                            rowHeight = 0.0; // only applies to first row
                            sceneRowHeight = sceneChild.DesiredSize.Height;
                            currentSceneX = startX;
                            sceneCountInRow = 0;
                        }
                        else
                        {
                            sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                        }

                        currentSceneX += SceneWidth + HorizontalGap;
                        actMaxRowWidth = Math.Max(actMaxRowWidth, currentSceneX);
                        sceneCountInRow++;
                        nextIndex++;
                    }

                    sequenceTotalHeight += Math.Max(sceneRowHeight, rowHeight);
                    actTotalHeight += sequenceTotalHeight + VerticalGap;
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
                double startX = MinActWidth + SequenceWidth + (HorizontalGap * 2);
                double currentSceneX = startX;
                double rowHeight;
                int nextIndex;
                double groupTotalHeight = 0.0;
                double sceneRowHeight = 0.0;
                int sceneCountInRow = 0;

                if (currentElement.Level == 1) // Sequence-started lone row
                {
                    rowHeight = currentChild.DesiredSize.Height;
                    nextIndex = i + 1;
                    maxMeasuredWidth = Math.Max(maxMeasuredWidth, BoardPadding + MinActWidth + HorizontalGap + SequenceWidth);
                }
                else // Scene-started lone row
                {
                    rowHeight = 0;
                    nextIndex = i;
                }

                // Collect and wrap Scenes
                while (nextIndex < InternalChildren.Count && 
                       TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                       sceneElement.Level >= 2)
                {
                    var sceneChild = InternalChildren[nextIndex];

                    if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                    {
                        groupTotalHeight += Math.Max(sceneRowHeight, rowHeight) + VerticalGap;
                        rowHeight = 0.0;
                        sceneRowHeight = sceneChild.DesiredSize.Height;
                        currentSceneX = startX;
                        sceneCountInRow = 0;
                    }
                    else
                    {
                        sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                    }

                    currentSceneX += SceneWidth + HorizontalGap;
                    maxMeasuredWidth = Math.Max(maxMeasuredWidth, currentSceneX);
                    sceneCountInRow++;
                    nextIndex++;
                }

                groupTotalHeight += Math.Max(sceneRowHeight, rowHeight);
                totalMeasuredHeight += groupTotalHeight + VerticalGap;
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
        var viewportWidth = ResolveViewportWidth(finalSize);

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
                    double sequenceTotalHeight = 0.0;
                    double currentLineTop = rowTop;
                    int nextIndex;

                    if (rowElement.Level == 1) // Sequence row
                    {
                        nextIndex = actEndIndex + 1;

                        // Place the Sequence card
                        var seqBounds = new Rect(BoardPadding + currentActWidth + HorizontalGap, rowTop, SequenceWidth, rowChild.DesiredSize.Height);
                        rowChild.Arrange(seqBounds);
                        _arrangedChildBounds[actEndIndex] = seqBounds;

                        double startX = BoardPadding + currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        double currentSceneX = startX;
                        double rowHeight = rowChild.DesiredSize.Height;
                        double sceneRowHeight = 0.0;
                        int sceneCountInRow = 0;
                        var currentLineScenes = new List<(UIElement Child, int Index, double OffsetX)>();

                        while (nextIndex < InternalChildren.Count && 
                               TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                               sceneElement.Level >= 2)
                        {
                            var sceneChild = InternalChildren[nextIndex];

                            if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                            {
                                // Arrange current line
                                double layoutHeight = Math.Max(sceneRowHeight, rowHeight);
                                foreach (var item in currentLineScenes)
                                {
                                    var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                                    item.Child.Arrange(itemBounds);
                                    _arrangedChildBounds[item.Index] = itemBounds;
                                }

                                sequenceTotalHeight += layoutHeight + VerticalGap;
                                currentLineTop += layoutHeight + VerticalGap;

                                rowHeight = 0.0;
                                sceneRowHeight = sceneChild.DesiredSize.Height;
                                currentSceneX = startX;
                                sceneCountInRow = 0;
                                currentLineScenes.Clear();
                            }
                            else
                            {
                                sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                            }

                            currentLineScenes.Add((sceneChild, nextIndex, currentSceneX));
                            currentSceneX += SceneWidth + HorizontalGap;
                            sceneCountInRow++;
                            nextIndex++;
                        }

                        // Arrange last line of scenes
                        double finalLayoutHeight = Math.Max(sceneRowHeight, rowHeight);
                        foreach (var item in currentLineScenes)
                        {
                            var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                            item.Child.Arrange(itemBounds);
                            _arrangedChildBounds[item.Index] = itemBounds;
                        }
                        sequenceTotalHeight += finalLayoutHeight;

                        // Track Sequence Lane Background
                        _sequenceLaneRects.Add(new Rect(BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop - LanePadding, finalSize.Width - BoardPadding - currentActWidth - (HorizontalGap / 2) - LanePadding, sequenceTotalHeight + (LanePadding * 2)));
                        _sequenceBoundaries.Add((BoardPadding + currentActWidth + (HorizontalGap / 2), rowTop + sequenceTotalHeight + (VerticalGap / 2)));

                        actRunningHeight += sequenceTotalHeight + VerticalGap;
                        actEndIndex = nextIndex;
                    }
                    else // Scene row (inside Act, no sequence card)
                    {
                        nextIndex = actEndIndex;
                        double startX = BoardPadding + currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        double currentSceneX = startX;
                        double sceneRowHeight = 0.0;
                        int sceneCountInRow = 0;
                        var currentLineScenes = new List<(UIElement Child, int Index, double OffsetX)>();

                        while (nextIndex < InternalChildren.Count && 
                               TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                               sceneElement.Level >= 2)
                        {
                            var sceneChild = InternalChildren[nextIndex];

                            if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                            {
                                // Arrange current line
                                foreach (var item in currentLineScenes)
                                {
                                    var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                                    item.Child.Arrange(itemBounds);
                                    _arrangedChildBounds[item.Index] = itemBounds;
                                }

                                sequenceTotalHeight += sceneRowHeight + VerticalGap;
                                currentLineTop += sceneRowHeight + VerticalGap;

                                sceneRowHeight = sceneChild.DesiredSize.Height;
                                currentSceneX = startX;
                                sceneCountInRow = 0;
                                currentLineScenes.Clear();
                            }
                            else
                            {
                                sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                            }

                            currentLineScenes.Add((sceneChild, nextIndex, currentSceneX));
                            currentSceneX += SceneWidth + HorizontalGap;
                            sceneCountInRow++;
                            nextIndex++;
                        }

                        // Arrange last line
                        foreach (var item in currentLineScenes)
                        {
                            var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                            item.Child.Arrange(itemBounds);
                            _arrangedChildBounds[item.Index] = itemBounds;
                        }
                        sequenceTotalHeight += sceneRowHeight;

                        actRunningHeight += sequenceTotalHeight + VerticalGap;
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
                double groupTotalHeight = 0.0;
                double currentLineTop = rowTop;
                int nextIndex;

                if (currentElement.Level == 1) // Sequence row
                {
                    nextIndex = i + 1;

                    // Arrange Sequence card
                    var seqBounds = new Rect(BoardPadding + MinActWidth + HorizontalGap, rowTop, SequenceWidth, currentChild.DesiredSize.Height);
                    currentChild.Arrange(seqBounds);
                    _arrangedChildBounds[i] = seqBounds;

                    double startX = BoardPadding + MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    double currentSceneX = startX;
                    double rowHeight = currentChild.DesiredSize.Height;
                    double sceneRowHeight = 0.0;
                    int sceneCountInRow = 0;
                    var currentLineScenes = new List<(UIElement Child, int Index, double OffsetX)>();

                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        var sceneChild = InternalChildren[nextIndex];

                        if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                        {
                            double layoutHeight = Math.Max(sceneRowHeight, rowHeight);
                            foreach (var item in currentLineScenes)
                            {
                                var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                                item.Child.Arrange(itemBounds);
                                _arrangedChildBounds[item.Index] = itemBounds;
                            }

                            groupTotalHeight += layoutHeight + VerticalGap;
                            currentLineTop += layoutHeight + VerticalGap;

                            rowHeight = 0.0;
                            sceneRowHeight = sceneChild.DesiredSize.Height;
                            currentSceneX = startX;
                            sceneCountInRow = 0;
                            currentLineScenes.Clear();
                        }
                        else
                        {
                            sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                        }

                        currentLineScenes.Add((sceneChild, nextIndex, currentSceneX));
                        currentSceneX += SceneWidth + HorizontalGap;
                        sceneCountInRow++;
                        nextIndex++;
                    }

                    double finalLayoutHeight = Math.Max(sceneRowHeight, rowHeight);
                    foreach (var item in currentLineScenes)
                    {
                        var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                        item.Child.Arrange(itemBounds);
                        _arrangedChildBounds[item.Index] = itemBounds;
                    }
                    groupTotalHeight += finalLayoutHeight;

                    _sequenceLaneRects.Add(new Rect(BoardPadding + MinActWidth + (HorizontalGap / 2), rowTop - LanePadding, finalSize.Width - BoardPadding - MinActWidth - (HorizontalGap / 2) - LanePadding, groupTotalHeight + (LanePadding * 2)));
                    _sequenceBoundaries.Add((BoardPadding + MinActWidth + (HorizontalGap / 2), rowTop + groupTotalHeight + (VerticalGap / 2)));
                }
                else // Scene row
                {
                    nextIndex = i;
                    double startX = BoardPadding + MinActWidth + SequenceWidth + (HorizontalGap * 2);
                    double currentSceneX = startX;
                    double sceneRowHeight = 0.0;
                    int sceneCountInRow = 0;
                    var currentLineScenes = new List<(UIElement Child, int Index, double OffsetX)>();

                    while (nextIndex < InternalChildren.Count && 
                           TryGetScreenplayElement(InternalChildren[nextIndex], out var sceneElement) && 
                           sceneElement.Level >= 2)
                    {
                        var sceneChild = InternalChildren[nextIndex];

                        if (currentSceneX + SceneWidth > viewportWidth - BoardPadding - TrailingGutterWidth && sceneCountInRow > 0)
                        {
                            foreach (var item in currentLineScenes)
                            {
                                var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                                item.Child.Arrange(itemBounds);
                                _arrangedChildBounds[item.Index] = itemBounds;
                            }

                            groupTotalHeight += sceneRowHeight + VerticalGap;
                            currentLineTop += sceneRowHeight + VerticalGap;

                            sceneRowHeight = sceneChild.DesiredSize.Height;
                            currentSceneX = startX;
                            sceneCountInRow = 0;
                            currentLineScenes.Clear();
                        }
                        else
                        {
                            sceneRowHeight = Math.Max(sceneRowHeight, sceneChild.DesiredSize.Height);
                        }

                        currentLineScenes.Add((sceneChild, nextIndex, currentSceneX));
                        currentSceneX += SceneWidth + HorizontalGap;
                        sceneCountInRow++;
                        nextIndex++;
                    }

                    foreach (var item in currentLineScenes)
                    {
                        var itemBounds = new Rect(item.OffsetX, currentLineTop, SceneWidth, item.Child.DesiredSize.Height);
                        item.Child.Arrange(itemBounds);
                        _arrangedChildBounds[item.Index] = itemBounds;
                    }
                    groupTotalHeight += sceneRowHeight;

                    _actBoundariesY.Add(rowTop + groupTotalHeight + (VerticalGap / 2));
                }

                currentY += groupTotalHeight + VerticalGap;
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
