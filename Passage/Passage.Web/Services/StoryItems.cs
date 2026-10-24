using System.Text.RegularExpressions;
using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>
/// Story data attached to Board cards and Beat Sheet beats (docs/STORY-PLAN.md,
/// decisions A, A2 and B). Pure functions, so the stamping and detaching rules
/// are testable without a browser.
/// </summary>
public static class StoryItems
{
    public const string CardKind = "card";
    public const string BeatKind = "beat";

    private static readonly Regex IdMarker = new(@"\[\[id:([a-f\d\-]+)\]\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool HasIdMarker(string line) => IdMarker.IsMatch(line);

    /// <summary>Appends the marker to the line; the parser strips it wherever it
    /// sits and then reads a trailing #12#, so the scene number survives. A
    /// marker already on the line (a duplicate) is swapped for the new id.</summary>
    public static string StampLine(string line, Guid id)
    {
        var marker = $"[[id:{id}]]";
        return IdMarker.IsMatch(line)
            ? IdMarker.Replace(line, marker, 1)
            : $"{line.TrimEnd()} {marker}";
    }

    /// <summary>1-based line numbers of every heading whose id was already seen
    /// earlier in the document. The first occurrence is not flagged.</summary>
    public static HashSet<int> DuplicateIdLines(IEnumerable<ScreenplayElement> elements)
    {
        var seen = new HashSet<Guid>();
        var duplicates = new HashSet<int>();
        foreach (var element in elements)
        {
            if (element is not (SectionElement or SceneHeadingElement))
            {
                continue;
            }

            if (!seen.Add(element.Id))
            {
                duplicates.Add(element.LineNumber);
            }
        }

        return duplicates;
    }

    /// <summary>Card items whose id matches no heading in the current parse.
    /// Computed on each parse and never stored (decision B).</summary>
    public static List<ItemTags> Detached(IEnumerable<ItemTags> items, IEnumerable<ScreenplayElement> elements)
    {
        var live = elements
            .Where(element => element is SectionElement or SceneHeadingElement)
            .Select(element => element.Id)
            .ToHashSet();
        return items.Where(item => item.Kind == CardKind && !live.Contains(item.TargetId)).ToList();
    }

    public static bool IsKey(StoryData story, Guid targetId) =>
        story.Items.Any(item => item.TargetId == targetId && item.Key);

    /// <summary>Sets or clears the star. An item carrying nothing is removed.</summary>
    public static void SetKey(StoryData story, Guid targetId, string kind, string label, bool key)
    {
        var item = story.Items.FirstOrDefault(existing => existing.TargetId == targetId);
        if (!key)
        {
            if (item is not null)
            {
                item.Key = false;
                story.Items.RemoveAll(existing => existing.TargetId == targetId && !existing.Key);
            }

            return;
        }

        if (item is null)
        {
            item = new ItemTags { TargetId = targetId, Kind = kind };
            story.Items.Add(item);
        }

        item.Key = true;
        item.Label = label;
    }

    /// <summary>Deleting a beat is explicit and the beat is sidecar-owned, so its
    /// tags go with it. Deleting a card never calls this (decision B).</summary>
    public static void RemoveBeat(StoryData story, Guid beatId) =>
        story.Items.RemoveAll(item => item.Kind == BeatKind && item.TargetId == beatId);

    /// <summary>Refreshes labels of attached items whose heading changed.
    /// Returns whether anything changed; called only on saves that happen anyway.</summary>
    public static bool RefreshLabels(StoryData story, IReadOnlyDictionary<Guid, string> labels)
    {
        var changed = false;
        foreach (var item in story.Items)
        {
            if (labels.TryGetValue(item.TargetId, out var label) && item.Label != label)
            {
                item.Label = label;
                changed = true;
            }
        }

        return changed;
    }
}
