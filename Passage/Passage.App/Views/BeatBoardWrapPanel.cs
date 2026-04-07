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

            if (TryGetScreenplayElement(child, out var element) && IsHorizontalChronologyElement(element))
            {
                var rowWidth = ResolveHierarchyIndent(child);
                var rowHeight = 0.0;
                var isFirstInRow = true;

                while (index < InternalChildren.Count &&
                       InternalChildren[index] is UIElement rowChild &&
                       TryGetScreenplayElement(rowChild, out var rowElement) &&
                       IsHorizontalChronologyElement(rowElement))
                {
                    if (!isFirstInRow)
                    {
                        rowWidth += HorizontalChronologyGap;
                    }

                    var rowChildSize = rowChild.DesiredSize;
                    rowWidth += rowChildSize.Width;
                    rowHeight = Math.Max(rowHeight, rowChildSize.Height);
                    isFirstInRow = false;
                    index++;
                }

                measuredWidth = Math.Max(measuredWidth, rowWidth + TrailingGutterWidth);
                measuredHeight += rowHeight;
                index--;
                continue;
            }

            var indent = ResolveHierarchyIndent(child);
            measuredWidth = Math.Max(
                measuredWidth,
                indent + ResolveSingleRowWidth(child, viewportWidth) + TrailingGutterWidth);
            measuredHeight += child.DesiredSize.Height;
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

            if (TryGetScreenplayElement(child, out var element) && IsHorizontalChronologyElement(element))
            {
                var lineLeft = ResolveHierarchyIndent(child);
                var rowHeight = 0.0;
                var rowStartIndex = index;
                var rowItems = new List<(UIElement Child, int Index)>();

                while (index < InternalChildren.Count &&
                       InternalChildren[index] is UIElement rowChild &&
                       TryGetScreenplayElement(rowChild, out var rowElement) &&
                       IsHorizontalChronologyElement(rowElement))
                {
                    rowItems.Add((rowChild, index));
                    rowHeight = Math.Max(rowHeight, rowChild.DesiredSize.Height);
                    index++;
                }

                foreach (var (rowChild, childIndex) in rowItems)
                {
                    var rowChildSize = rowChild.DesiredSize;
                    var rowBounds = new Rect(lineLeft, lineTop, rowChildSize.Width, rowChildSize.Height);
                    rowChild.Arrange(rowBounds);
                    _arrangedChildBounds[childIndex] = rowBounds;
                    lineLeft += rowChildSize.Width + HorizontalChronologyGap;
                }

                lineTop += rowHeight;
                index = rowStartIndex + rowItems.Count - 1;
                continue;
            }

            var childSize = child.DesiredSize;
            var indent = ResolveHierarchyIndent(child);
            var isFullWidth = GetIsFullWidth(child);
            var availableWidth = Math.Max(0.0, contentWidth - indent - TrailingGutterWidth);
            var arrangedWidth = isFullWidth
                ? Math.Max(childSize.Width, availableWidth)
                : Math.Min(childSize.Width, availableWidth);
            var arrangedBounds = new Rect(indent, lineTop, arrangedWidth, childSize.Height);
            child.Arrange(arrangedBounds);
            _arrangedChildBounds[index] = arrangedBounds;
            lineTop += childSize.Height;
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
            ? Math.Max(280.0, viewportWidth - indent - TrailingGutterWidth)
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
        return element.Level >= 2;
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
