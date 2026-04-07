namespace Passage.Parser;

public readonly record struct StoryBlockRange(int Index, int Count);

public static class StoryHierarchyHelper
{
    public static StoryBlockRange GetBlockRange(IReadOnlyList<ScreenplayElement> elements, ScreenplayElement parent)
    {
        ArgumentNullException.ThrowIfNull(elements);
        ArgumentNullException.ThrowIfNull(parent);

        var parentIndex = -1;
        for (var index = 0; index < elements.Count; index++)
        {
            var element = elements[index];
            if (ReferenceEquals(element, parent) || element.Id == parent.Id)
            {
                parentIndex = index;
                break;
            }
        }

        if (parentIndex < 0)
        {
            throw new ArgumentException("The supplied parent element is not part of the element list.", nameof(parent));
        }

        var endIndexExclusive = elements.Count;
        for (var index = parentIndex + 1; index < elements.Count; index++)
        {
            if (elements[index].Level <= parent.Level)
            {
                endIndexExclusive = index;
                break;
            }
        }

        return new StoryBlockRange(parentIndex, endIndexExclusive - parentIndex);
    }
}
