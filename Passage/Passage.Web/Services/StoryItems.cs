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

    /// <summary>Whether tagging this heading line must stamp it first: it has no
    /// marker, or its marker is a copy of an earlier heading's.</summary>
    public static bool NeedsStamp(string line, bool duplicate) => duplicate || !HasIdMarker(line);

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
                story.Items.RemoveAll(existing => existing.TargetId == targetId && existing.IsEmpty);
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

    /// <summary>A palette hex, or empty for anything else.</summary>
    public static string StorylineColour(string? colour) =>
        colour is not null && Characters.Palette.Contains(colour) ? colour : string.Empty;

    public static int StorylineUseCount(StoryData story, Guid storylineId) =>
        story.Items.Count(item => item.Storylines.Contains(storylineId));

    /// <summary>Removes the storyline and untags it from every item. Returns how
    /// many items were untagged. Items left carrying nothing are dropped.</summary>
    public static int DeleteStoryline(StoryData story, Guid storylineId)
    {
        var removed = story.Storylines.RemoveAll(line => line.Id == storylineId);
        story.Arc.RemoveAll(point => point.StorylineId == storylineId);
        var untagged = 0;
        foreach (var item in story.Items)
        {
            if (item.Storylines.RemoveAll(id => id == storylineId) > 0)
            {
                untagged++;
            }
        }

        if (removed > 0)
        {
            story.Items.RemoveAll(item => item.IsEmpty);
        }

        return untagged;
    }

    /// <summary>Tag keys are "storyline:&lt;guid&gt;", "character:&lt;name&gt;" or
    /// "location:&lt;name&gt;" (S1.4). Returns false for anything else.</summary>
    public static bool TryParseTag(string key, out string type, out string value)
    {
        type = value = string.Empty;
        var colon = key.IndexOf(':');
        if (colon <= 0)
        {
            return false;
        }

        type = key[..colon];
        value = key[(colon + 1)..];
        return type is "storyline" or "character" or "location" && value.Length > 0;
    }

    /// <summary>Adds the tag if the item lacks it, removes it if it has it.
    /// Characters are stored as bare cue names. Returns true when the tag is on
    /// afterwards, false when it was removed or the key was rejected. An item
    /// left carrying nothing is dropped.</summary>
    public static bool ToggleTag(StoryData story, Guid targetId, string kind, string label, string key)
    {
        if (!TryParseTag(key, out var type, out var value))
        {
            return false;
        }

        if (type == "character")
        {
            value = Characters.BareName(value);
            if (value.Length == 0)
            {
                return false;
            }
        }
        else if (type == "storyline" && !Guid.TryParse(value, out _))
        {
            return false;
        }

        var item = story.Items.FirstOrDefault(existing => existing.TargetId == targetId);
        if (item is null)
        {
            item = new ItemTags { TargetId = targetId, Kind = kind };
            story.Items.Add(item);
        }

        item.Label = label;
        bool on;
        switch (type)
        {
            case "storyline":
                var line = Guid.Parse(value);
                on = !item.Storylines.Remove(line);
                if (on) item.Storylines.Add(line);
                break;
            case "character":
                on = !item.Characters.Remove(value);
                if (on) item.Characters.Add(value);
                break;
            default:
                on = !item.Locations.Remove(value);
                if (on) item.Locations.Add(value);
                break;
        }

        if (item.IsEmpty)
        {
            story.Items.Remove(item);
        }

        return on;
    }

    /// <summary>The keys the popover shows as ticked for an item.</summary>
    public static List<string> TagKeys(ItemTags item) =>
        item.Storylines.Select(id => "storyline:" + id)
            .Concat(item.Characters.Select(name => "character:" + name))
            .Concat(item.Locations.Select(name => "location:" + name))
            .ToList();

    /// <summary>Rewrites a character name on every item, merging an item that
    /// already carries the new name. Returns how many items changed.</summary>
    public static int RenameCharacter(StoryData story, string oldName, string newName)
    {
        var changed = 0;
        foreach (var item in story.Items)
        {
            var index = item.Characters.IndexOf(oldName);
            if (index < 0)
            {
                continue;
            }

            changed++;
            if (item.Characters.Contains(newName))
            {
                item.Characters.RemoveAt(index);
            }
            else
            {
                item.Characters[index] = newName;
            }
        }

        return changed;
    }

    /// <summary>Deleting a beat is explicit and the beat is sidecar-owned, so its
    /// tags go with it. Deleting a card never calls this (decision B).</summary>
    public static void RemoveBeat(StoryData story, Guid beatId)
    {
        story.Items.RemoveAll(item => item.Kind == BeatKind && item.TargetId == beatId);
        story.Arc.RemoveAll(point => point.Kind == BeatKind && point.TargetId == beatId);
    }

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

        foreach (var point in story.Arc)
        {
            if (labels.TryGetValue(point.TargetId, out var label) && point.Label != label)
            {
                point.Label = label;
                changed = true;
            }
        }

        return changed;
    }
}
