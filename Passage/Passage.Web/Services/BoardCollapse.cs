namespace Passage.Web.Services;

/// <summary>
/// The keys the Board uses to remember which lanes and groups are collapsed
/// (docs/STORY-PLAN.md S2.4). Keyed by heading, since the web has no stable
/// card identity; untitled lanes and groups share one fallback key each.
/// </summary>
public static class BoardCollapse
{
    public static string LaneKey(BoardLane lane) => lane.ActCard?.Heading ?? "\u0000lane";

    public static string GroupKey(string laneKey, BoardGroup group) =>
        laneKey + "/" + (group.SequenceCard?.Heading ?? "\u0000group");

    /// <summary>Every lane key and every group key on the Board, untitled ones
    /// included, so "Collapse all" leaves nothing open.</summary>
    public static (HashSet<string> Lanes, HashSet<string> Groups) AllKeys(IReadOnlyList<BoardLane> lanes)
    {
        var laneKeys = new HashSet<string>(StringComparer.Ordinal);
        var groupKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var lane in lanes)
        {
            var laneKey = LaneKey(lane);
            laneKeys.Add(laneKey);
            foreach (var group in lane.Groups)
            {
                groupKeys.Add(GroupKey(laneKey, group));
            }
        }

        return (laneKeys, groupKeys);
    }
}
