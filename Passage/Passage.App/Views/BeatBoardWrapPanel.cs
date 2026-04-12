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
            if (child is null) continue;
            MeasureBeatBoardChild(child, availableSize.Height, viewportWidth);
        }

        for (var index = 0; index < InternalChildren.Count; index++)
        {
            if (InternalChildren[index] is not UIElement child) continue;

            if (!TryGetScreenplayElement(child, out var element))
            {
                measuredWidth = Math.Max(measuredWidth, child.DesiredSize.Width);
                measuredHeight += child.DesiredSize.Height;
                continue;
            }

            if (element.Level <= 0) // Act - Full width divider
            {
                measuredWidth = Math.Max(measuredWidth, child.DesiredSize.Width + TrailingGutterWidth);
                measuredHeight += child.DesiredSize.Height;
                continue;
            }

            // Sequence-initiated horizontal row
            double rowX = ResolveHierarchyIndent(child);
            double currentX = rowX;
            double currentLineHeight = 0;
            double rowHeightTotal = 0;
            double rowMaxWidth = 0;
            bool isFirstInRow = true;

            while (index < InternalChildren.Count &&
                   InternalChildren[index] is UIElement horizontalChild &&
                   TryGetScreenplayElement(horizontalChild, out var horizontalElement))
            {
                // Only the first element starts a row (Level 1 or 2).
                // Subsequent elements MUST be Level >= 2 to stay in this row.
                if (!isFirstInRow && horizontalElement.Level < 2) break;

                var childSize = horizontalChild.DesiredSize;
                double indent = ResolveHierarchyIndent(horizontalChild);

                if (!isFirstInRow)
                {
                    // Check for wrap
                    if (currentX + HorizontalChronologyGap + childSize.Width > viewportWidth - TrailingGutterWidth)
                    {
                        rowHeightTotal += currentLineHeight;
                        currentX = indent; // Respect hierarchy indent on wrap
                        currentLineHeight = 0;
                    }
                    else
                    {
                        currentX += HorizontalChronologyGap;
                    }
                }
                else
                {
                    currentX = indent;
                }

                currentX += childSize.Width;
                currentLineHeight = Math.Max(currentLineHeight, childSize.Height);
                rowMaxWidth = Math.Max(rowMaxWidth, currentX);
                isFirstInRow = false;
                index++;
            }

            rowHeightTotal += currentLineHeight;
            measuredWidth = Math.Max(measuredWidth, rowMaxWidth + TrailingGutterWidth);
            measuredHeight += rowHeightTotal;
            index--;
        }

        return new Size(Math.Max(viewportWidth, measuredWidth), measuredHeight);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        EnsureArrangedChildBoundsCapacity();

        var viewportWidth = Math.Max(finalSize.Width, ResolveViewportWidth(finalSize));
        var lineTop = 0.0;

        for (var index = 0; index < InternalChildren.Count; index++)
        {
            if (InternalChildren[index] is not UIElement child) continue;

            if (!TryGetScreenplayElement(child, out var element))
            {
                var bounds = new Rect(0, lineTop, child.DesiredSize.Width, child.DesiredSize.Height);
                child.Arrange(bounds);
                _arrangedChildBounds[index] = bounds;
                lineTop += child.DesiredSize.Height;
                continue;
            }

            if (element.Level <= 0) // Act - Full width divider
            {
                var indent = ResolveHierarchyIndent(child);
                var bounds = new Rect(indent, lineTop, child.DesiredSize.Width, child.DesiredSize.Height);
                child.Arrange(bounds);
                _arrangedChildBounds[index] = bounds;
                lineTop += child.DesiredSize.Height;
                continue;
            }

            // Sequence-initiated horizontal row
            double currentX = ResolveHierarchyIndent(child);
            double currentLineHeight = 0;
            bool isFirstInRow = true;

            // We need to look ahead to group them exactly as MeasureOverride did
            int startIndex = index;
            while (index < InternalChildren.Count &&
                   InternalChildren[index] is UIElement horizontalChild &&
                   TryGetScreenplayElement(horizontalChild, out var horizontalElement))
            {
                if (!isFirstInRow && horizontalElement.Level < 2) break;

                var childSize = horizontalChild.DesiredSize;
                double indent = ResolveHierarchyIndent(horizontalChild);

                if (!isFirstInRow)
                {
                    if (currentX + HorizontalChronologyGap + childSize.Width > viewportWidth - TrailingGutterWidth)
                    {
                        lineTop += currentLineHeight;
                        currentX = indent;
                        currentLineHeight = 0;
                    }
                    else
                    {
                        currentX += HorizontalChronologyGap;
                    }
                }
                else
                {
                    currentX = indent;
                }

                var bounds = new Rect(currentX, lineTop, childSize.Width, childSize.Height);
                horizontalChild.Arrange(bounds);
                _arrangedChildBounds[index] = bounds;

                currentLineHeight = Math.Max(currentLineHeight, childSize.Height);
                currentX += childSize.Width;
                isFirstInRow = false;
                index++;
            }

            lineTop += currentLineHeight;
            index--;
        }

        return new Size(viewportWidth, Math.Max(finalSize.Height, lineTop));
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
