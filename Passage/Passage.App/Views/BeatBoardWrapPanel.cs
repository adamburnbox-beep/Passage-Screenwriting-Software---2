using System.Windows;
using System.Windows.Controls;
using Passage.Parser;

namespace Passage.App.Views;

public sealed class BeatBoardWrapPanel : Panel
{
    private const double HierarchyIndentWidth = 48.0;
    private const double HorizontalChronologyGap = 16.0;
    private const double TrailingGutterWidth = 18.0;
    private const double FallbackViewportWidth = 960.0;

    private readonly List<Rect> _arrangedChildBounds = new();

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

    protected override Size MeasureOverride(Size availableSize)
    {
        var viewportWidth = ResolveViewportWidth(availableSize);
        var measuredWidth = 0.0;
        var measuredHeight = 0.0;

        foreach (UIElement child in InternalChildren)
        {
            if (child is null)
            {
                continue;
            }

            MeasureBeatBoardChild(child, availableSize.Height, viewportWidth);
        }

        for (var index = 0; index < InternalChildren.Count; index++)
        {
            if (InternalChildren[index] is not UIElement child)
            {
                continue;
            }

            if (!TryGetScreenplayElement(child, out var element))
            {
                measuredWidth = Math.Max(measuredWidth, child.DesiredSize.Width);
                measuredHeight += child.DesiredSize.Height;
                continue;
            }



            if (element.Level <= 0) // Act - Row Header
            {
                measuredWidth = Math.Max(measuredWidth, child.DesiredSize.Width + TrailingGutterWidth);
                measuredHeight += child.DesiredSize.Height;
                continue;
            }

            double laneWidth = ResolveHierarchyIndent(child);
            double laneHeight = 0;
            bool isFirstInLane = true;

            while (index < InternalChildren.Count &&
                   InternalChildren[index] is UIElement horizontalChild &&
                   TryGetScreenplayElement(horizontalChild, out var horizontalElement) &&
                   IsHorizontalChronologyElement(horizontalElement))
            {
                if (!isFirstInLane)
                {
                    laneWidth += HorizontalChronologyGap;
                }

                var rowChildSize = horizontalChild.DesiredSize;
                laneWidth += rowChildSize.Width;
                laneHeight = Math.Max(laneHeight, rowChildSize.Height);
                isFirstInLane = false;
                index++;
            }

            measuredWidth = Math.Max(measuredWidth, laneWidth + TrailingGutterWidth);
            measuredHeight += laneHeight;
            index--;
        }

        return new Size(
            Math.Max(viewportWidth, measuredWidth),
            measuredHeight);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        EnsureArrangedChildBoundsCapacity();

        var lineTop = 0.0;
        var contentWidth = Math.Max(finalSize.Width, ResolveViewportWidth(finalSize));

        for (var index = 0; index < InternalChildren.Count; index++)
        {
            if (InternalChildren[index] is not UIElement child)
            {
                continue;
            }

            if (!TryGetScreenplayElement(child, out var element))
            {
                var bounds = new Rect(0, lineTop, child.DesiredSize.Width, child.DesiredSize.Height);
                child.Arrange(bounds);
                _arrangedChildBounds[index] = bounds;
                lineTop += child.DesiredSize.Height;
                continue;
            }



            if (element.Level <= 0) // Act - Row Header
            {
                var rowHeaderIndent = ResolveHierarchyIndent(child);
                var childBounds = new Rect(rowHeaderIndent, lineTop, child.DesiredSize.Width, child.DesiredSize.Height);
                child.Arrange(childBounds);
                _arrangedChildBounds[index] = childBounds;
                lineTop += child.DesiredSize.Height;
                continue;
            }

            double lineLeft = ResolveHierarchyIndent(child);
            double rowHeight = 0;

            while (index < InternalChildren.Count &&
                   InternalChildren[index] is UIElement horizontalChild &&
                   TryGetScreenplayElement(horizontalChild, out var horizontalElement) &&
                   IsHorizontalChronologyElement(horizontalElement))
            {
                var rowChildSize = horizontalChild.DesiredSize;
                var rowBounds = new Rect(lineLeft, lineTop, rowChildSize.Width, rowChildSize.Height);
                horizontalChild.Arrange(rowBounds);
                _arrangedChildBounds[index] = rowBounds;

                rowHeight = Math.Max(rowHeight, rowChildSize.Height);
                lineLeft += rowChildSize.Width + HorizontalChronologyGap;
                index++;
            }

            lineTop += rowHeight;
            index--;
        }

        return new Size(contentWidth, Math.Max(finalSize.Height, lineTop));
    }

    protected override void OnVisualChildrenChanged(DependencyObject visualAdded, DependencyObject visualRemoved)
    {
        base.OnVisualChildrenChanged(visualAdded, visualRemoved);
        InvalidateMeasure();
        InvalidateArrange();
        InvalidateVisual();
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

        return FallbackViewportWidth;
    }

    private void MeasureBeatBoardChild(UIElement child, double availableHeight, double viewportWidth)
    {
        var indent = ResolveHierarchyIndent(child);
        var childMeasureWidth = GetIsFullWidth(child)
            ? Math.Max(280.0, viewportWidth - TrailingGutterWidth)
            : double.PositiveInfinity;
        child.Measure(new Size(childMeasureWidth, availableHeight));
    }

    private static double ResolveSingleRowWidth(UIElement child, double viewportWidth)
    {
        var desiredWidth = child.DesiredSize.Width;
        if (!GetIsFullWidth(child))
        {
            return desiredWidth;
        }

        var indent = ResolveHierarchyIndent(child);
        return Math.Max(desiredWidth, Math.Max(0.0, viewportWidth - indent - TrailingGutterWidth));
    }

    private static double ResolveHierarchyIndent(UIElement child)
    {
        if (child is not FrameworkElement { DataContext: ScreenplayElement element })
        {
            return 0.0;
        }

        return Math.Max(0, element.Level) * HierarchyIndentWidth;
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

    private static bool IsHorizontalChronologyElement(ScreenplayElement element)
    {
        return element.Level >= 1;
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
