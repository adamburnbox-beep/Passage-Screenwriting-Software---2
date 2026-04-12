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
    private const double VerticalGap = 16.0;
    private const double TrailingGutterWidth = 18.0;

    private readonly List<Rect> _arrangedChildBounds = new();
    private readonly List<double> _actBoundariesY = new();
    private readonly List<(double X, double Y)> _sequenceBoundaries = new();

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

        return new Size(maxMeasuredWidth + TrailingGutterWidth, totalMeasuredHeight);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        EnsureArrangedChildBoundsCapacity();
        _actBoundariesY.Clear();
        _sequenceBoundaries.Clear();
        var currentY = 0.0;

        for (var i = 0; i < InternalChildren.Count; i++)
        {
            if (InternalChildren[i] is not UIElement actChild || !TryGetScreenplayElement(actChild, out var actElement))
            {
                if (InternalChildren[i] is UIElement nonElement)
                {
                    var bounds = new Rect(0, currentY, nonElement.DesiredSize.Width, nonElement.DesiredSize.Height);
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

                        var seqBounds = new Rect(currentActWidth + HorizontalGap, rowTop, SequenceWidth, seqChild.DesiredSize.Height);
                        seqChild.Arrange(seqBounds);
                        _arrangedChildBounds[actEndIndex] = seqBounds;

                        double currentSceneX = currentActWidth + SequenceWidth + (HorizontalGap * 2);
                        for (int k = actEndIndex + 1; k < sceneIndex; k++)
                        {
                            var sChild = InternalChildren[k];
                            var sBounds = new Rect(currentSceneX, rowTop, SceneWidth, sChild.DesiredSize.Height);
                            sChild.Arrange(sBounds);
                            _arrangedChildBounds[k] = sBounds;
                            currentSceneX += SceneWidth + HorizontalGap;
                        }

                        actRunningHeight += rowHeight + VerticalGap;
                        internalRowBottoms.Add(actTop + actRunningHeight - (VerticalGap / 2));
                        actEndIndex = sceneIndex;
                    }
                    else
                    {
                        var bounds = new Rect(currentActWidth + HorizontalGap, actTop + actRunningHeight, seqChild.DesiredSize.Width, seqChild.DesiredSize.Height);
                        seqChild.Arrange(bounds);
                        _arrangedChildBounds[actEndIndex] = bounds;
                        actRunningHeight += bounds.Height + VerticalGap;
                        internalRowBottoms.Add(actTop + actRunningHeight - (VerticalGap / 2));
                        actEndIndex++;
                    }
                }

                // Perfect Spanning Logic: 
                // The height of the card is the distance from top of first row to bottom of last row.
                // If we have content, the card should span from actTop (same as first row) 
                // to exactly the end of the content (minus trailing gap).
                double cardHeight = actRunningHeight > 0 ? actRunningHeight - VerticalGap : actChild.DesiredSize.Height;
                var actBounds = new Rect(0, actTop, currentActWidth, cardHeight);
                actChild.Arrange(actBounds);
                _arrangedChildBounds[i] = actBounds;

                // Dividers: Add all internal row bottoms to sequence boundaries EXCEPT the last one
                if (internalRowBottoms.Count > 1)
                {
                    for (int j = 0; j < internalRowBottoms.Count - 1; j++)
                    {
                        _sequenceBoundaries.Add((currentActWidth + HorizontalGap, internalRowBottoms[j]));
                    }
                }

                // Add the absolute bottom of the Act block to the Act boundaries
                double actBlockBottom = actTop + Math.Max(actRunningHeight, actChild.DesiredSize.Height + VerticalGap);
                _actBoundariesY.Add(actBlockBottom - (VerticalGap / 2));

                currentY = actBlockBottom;
                i = actEndIndex - 1;
            }
            else
            {
                var bounds = new Rect(0, currentY, actChild.DesiredSize.Width, actChild.DesiredSize.Height);
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

        var actBrush = TryFindResource("BeatBoardActBoundaryBrush") as Brush;
        var seqBrush = TryFindResource("BeatBoardSequenceSeparatorBrush") as Brush;

        // Draw Act Boundaries (Red Lines) - Full Width
        if (actBrush != null)
        {
            var actPen = new Pen(actBrush, 2.0);
            actPen.Freeze();
            foreach (var y in _actBoundariesY)
            {
                dc.DrawLine(actPen, new Point(0, y), new Point(ActualWidth, y));
            }
        }

        // Draw Sequence Separators (Yellow Lines) - Starting from the correct column
        if (seqBrush != null)
        {
            var seqPen = new Pen(seqBrush, 1.0);
            seqPen.Freeze();
            foreach (var boundary in _sequenceBoundaries)
            {
                dc.DrawLine(seqPen, new Point(boundary.X, boundary.Y), new Point(ActualWidth, boundary.Y));
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
