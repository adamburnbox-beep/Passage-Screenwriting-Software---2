namespace Passage.Web.Services;

/// <summary>
/// Which outline kinds the OUTLINE panel shows (HIGHLAND-PLAN H4.3). Per browser,
/// in the session key. Notes start hidden so the default outline is unchanged.
/// Every kind that is not a scene, synopsis, note or marker is a section
/// (Act / Sequence / Beat, and the H1–H6 headings in Markdown mode).
/// </summary>
public sealed record OutlineFilters
{
    public bool Sections { get; init; } = true;

    public bool Scenes { get; init; } = true;

    public bool Synopses { get; init; } = true;

    public bool Notes { get; init; }

    public bool Markers { get; init; } = true;

    public bool Shows(string kind) => kind switch
    {
        "Scene" => Scenes,
        "Synopsis" => Synopses,
        "Note" => Notes,
        "Marker" => Markers,
        _ => Sections
    };

    /// <summary>
    /// The outline with hidden kinds removed. A hidden node's visible descendants take its
    /// place, so hiding sections still lists their scenes.
    /// </summary>
    public IReadOnlyList<OutlineNode> Apply(IReadOnlyList<OutlineNode> nodes)
    {
        var visible = new List<OutlineNode>();
        foreach (var node in nodes)
        {
            var children = Apply(node.Children);
            if (Shows(node.Kind))
            {
                visible.Add(node with { Children = children.ToList() });
            }
            else
            {
                visible.AddRange(children);
            }
        }
        return visible;
    }
}
