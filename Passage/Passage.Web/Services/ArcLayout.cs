using Passage.Parser;

namespace Passage.Web.Services;

public enum ArcAxisMode { Board, Pages }

/// <summary>A point on a storyline's line. X is a Board index or a page number
/// depending on the mode; Touched is false for the hollow default of 50.</summary>
public sealed record ArcPlotPoint(Guid StorylineId, Guid TargetId, string Kind, string Label, int X, int Value, bool Touched);

/// <summary>A tagged item the current x-axis cannot place. Never guessed.</summary>
public sealed record ArcUnplaced(Guid TargetId, string Kind, string Label, string Reason);

/// <summary>A point whose card no longer parses (decision B): kept and listed.</summary>
public sealed record ArcDetached(Guid StorylineId, Guid TargetId, string Label, int Value);

public sealed record ArcChart(
    Dictionary<Guid, List<ArcPlotPoint>> Lines,
    List<ArcUnplaced> Unplaced,
    List<ArcDetached> Detached,
    int XMin,
    int XMax);

/// <summary>
/// The ARC view's data (docs/STORY-PLAN.md S3.1). Pure, so both x-axis
/// orderings, the clamp and the unplaced report are testable without a browser.
/// </summary>
public static class ArcLayout
{
    public const int DefaultValue = 50;

    public static int Clamp(int value) => Math.Clamp(value, 0, 100);

    /// <summary>Stores a dragged value. False (nothing written) when the
    /// storyline is unknown or the item is not tagged with it.</summary>
    public static bool SetValue(StoryData story, Guid storylineId, Guid targetId, string kind, string label, int value)
    {
        var tagged = story.Storylines.Any(line => line.Id == storylineId)
            && story.Items.Any(item => item.TargetId == targetId && item.Storylines.Contains(storylineId));
        if (!tagged)
        {
            return false;
        }

        var point = story.Arc.FirstOrDefault(existing => existing.StorylineId == storylineId && existing.TargetId == targetId);
        if (point is null)
        {
            point = new ArcPoint { StorylineId = storylineId, TargetId = targetId, Kind = kind };
            story.Arc.Add(point);
        }

        point.Label = label;
        point.Value = Clamp(value);
        return true;
    }

    public static ArcChart Build(
        StoryData story,
        IReadOnlyList<ScreenplayElement> elements,
        IReadOnlyList<ScenePlacement> scenes,
        IReadOnlyList<Beat> beats,
        ArcAxisMode mode)
    {
        var pageByLine = scenes.ToDictionary(scene => scene.LineNumber, scene => scene.Page);
        var duplicates = StoryItems.DuplicateIdLines(elements);

        // Board order: every card in script order. A duplicated id belongs to
        // its first heading only.
        var cards = new List<(Guid Id, string Label, int Page)>();
        var seen = new HashSet<Guid>();
        foreach (var element in elements)
        {
            if (element is not (SectionElement or SceneHeadingElement) || element.Id == Guid.Empty
                || duplicates.Contains(element.LineNumber) || !seen.Add(element.Id))
            {
                continue;
            }

            var page = element is SceneHeadingElement && pageByLine.TryGetValue(element.LineNumber, out var found) ? found : 0;
            cards.Add((element.Id, element.Heading, page));
        }

        var slots = new Dictionary<Guid, (string Kind, string Label, int X)>();
        var unplacedReason = new Dictionary<Guid, (string Kind, string Label, string Reason)>();
        for (var i = 0; i < cards.Count; i++)
        {
            var (id, label, page) = cards[i];
            if (mode == ArcAxisMode.Board)
            {
                slots[id] = (StoryItems.CardKind, label, i);
            }
            else if (page > 0)
            {
                slots[id] = (StoryItems.CardKind, label, page);
            }
            else
            {
                unplacedReason[id] = (StoryItems.CardKind, label, "not on a page");
            }
        }

        foreach (var beat in beats)
        {
            var label = BeatInbox.Heading(beat);
            if (mode == ArcAxisMode.Pages && beat.Page > 0)
            {
                slots[beat.Id] = (StoryItems.BeatKind, label, beat.Page);
            }
            else
            {
                unplacedReason[beat.Id] = (StoryItems.BeatKind, label, mode == ArcAxisMode.Board ? "not on the Board" : "not on a page");
            }
        }

        // Ties on one x keep script order, then beat-list order.
        var order = cards.Select(card => card.Id).Concat(beats.Select(beat => beat.Id))
            .Select((id, index) => (id, index)).ToDictionary(pair => pair.id, pair => pair.index);
        var values = story.Arc.GroupBy(point => (point.StorylineId, point.TargetId))
            .ToDictionary(group => group.Key, group => group.First().Value);

        var lines = new Dictionary<Guid, List<ArcPlotPoint>>();
        var unplaced = new Dictionary<Guid, ArcUnplaced>();
        foreach (var line in story.Storylines)
        {
            var points = new List<ArcPlotPoint>();
            foreach (var item in story.Items.Where(item => item.Storylines.Contains(line.Id)))
            {
                if (slots.TryGetValue(item.TargetId, out var slot))
                {
                    var touched = values.TryGetValue((line.Id, item.TargetId), out var value);
                    points.Add(new ArcPlotPoint(line.Id, item.TargetId, slot.Kind, slot.Label, slot.X, touched ? Clamp(value) : DefaultValue, touched));
                }
                else if (unplacedReason.TryGetValue(item.TargetId, out var reason))
                {
                    unplaced.TryAdd(item.TargetId, new ArcUnplaced(item.TargetId, reason.Kind, reason.Label, reason.Reason));
                }
            }

            lines[line.Id] = points.OrderBy(point => point.X).ThenBy(point => order.GetValueOrDefault(point.TargetId, int.MaxValue)).ToList();
        }

        var detached = new List<ArcDetached>();
        var live = cards.Select(card => card.Id).ToHashSet();
        foreach (var item in StoryItems.Detached(story.Items, elements))
        {
            foreach (var lineId in item.Storylines.Where(id => story.Storylines.Any(line => line.Id == id)))
            {
                detached.Add(new ArcDetached(lineId, item.TargetId, item.Label, values.GetValueOrDefault((lineId, item.TargetId), DefaultValue)));
            }
        }

        foreach (var point in story.Arc.Where(point => point.Kind == StoryItems.CardKind && !live.Contains(point.TargetId)
            && !detached.Any(entry => entry.StorylineId == point.StorylineId && entry.TargetId == point.TargetId)
            && story.Storylines.Any(line => line.Id == point.StorylineId)))
        {
            detached.Add(new ArcDetached(point.StorylineId, point.TargetId, point.Label, Clamp(point.Value)));
        }

        var xs = lines.Values.SelectMany(points => points).Select(point => point.X).ToList();
        var xMin = mode == ArcAxisMode.Board ? 0 : 1;
        var xMax = mode == ArcAxisMode.Board ? Math.Max(1, cards.Count - 1) : Math.Max(Math.Max(2, xs.DefaultIfEmpty(0).Max()), 2);
        return new ArcChart(lines, unplaced.Values.ToList(), detached, xMin, xMax);
    }
}
