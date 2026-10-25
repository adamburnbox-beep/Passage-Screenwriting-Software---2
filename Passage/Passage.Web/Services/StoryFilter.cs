using Passage.Parser;

namespace Passage.Web.Services;

/// <summary>
/// The BOARD / ARC filter bar (docs/STORY-PLAN.md S1.5). Per browser, one set
/// per view, in the session key like <see cref="OutlineFilters"/>. Choices
/// combine as OR within a kind and AND across kinds; an empty kind is not
/// constrained, so an empty filter matches everything. Characters are bare
/// cue names and locations are normalised names (decision C).
/// </summary>
public sealed record StoryFilter
{
    public List<Guid> Storylines { get; init; } = new();

    public List<string> Characters { get; init; } = new();

    public List<string> Locations { get; init; } = new();

    public bool KeyOnly { get; init; }

    public bool IsActive => KeyOnly || Storylines.Count > 0 || Characters.Count > 0 || Locations.Count > 0;

    /// <summary>Whether an item survives the filter at full strength. Callers
    /// pass the item's tags plus anything derived (a speaking character, a
    /// scene's own location).</summary>
    public bool Matches(
        bool key,
        IReadOnlyCollection<Guid> storylines,
        IReadOnlyCollection<string> characters,
        IReadOnlyCollection<string> locations) =>
        (!KeyOnly || key)
        && (Storylines.Count == 0 || Storylines.Any(storylines.Contains))
        && (Characters.Count == 0 || Characters.Any(characters.Contains))
        && (Locations.Count == 0 || Locations.Any(locations.Contains));

    /// <summary>An act or sequence lane matches if it does, or any card in it does.</summary>
    public static bool LaneMatches(bool self, IEnumerable<bool> cards) => self || cards.Any(match => match);

    public StoryFilter ToggleStoryline(Guid id) =>
        this with { Storylines = Toggled(Storylines, id) };

    public StoryFilter ToggleCharacter(string name) =>
        this with { Characters = Toggled(Characters, name) };

    public StoryFilter ToggleLocation(string name) =>
        this with { Locations = Toggled(Locations, name) };

    private static List<T> Toggled<T>(List<T> list, T value) =>
        list.Contains(value) ? list.Where(entry => !EqualityComparer<T>.Default.Equals(entry, value)).ToList() : new List<T>(list) { value };

    /// <summary>Bare cue names speaking in each scene, keyed by the scene
    /// heading's line number. Derived on demand, never stored.</summary>
    public static Dictionary<int, HashSet<string>> SceneSpeakers(IEnumerable<ScreenplayElement> elements)
    {
        var speakers = new Dictionary<int, HashSet<string>>();
        HashSet<string>? current = null;
        foreach (var element in elements)
        {
            if (element is SceneHeadingElement scene)
            {
                speakers[scene.LineNumber] = current = new HashSet<string>(StringComparer.Ordinal);
            }
            else if (element is CharacterElement cue && current is not null)
            {
                var name = Passage.Web.Services.Characters.BareName(cue.CharacterName);
                if (name.Length > 0)
                {
                    current.Add(name);
                }
            }
        }

        return speakers;
    }
}
