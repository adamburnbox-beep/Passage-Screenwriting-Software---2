using System.Text.Json;
using System.Text.RegularExpressions;

namespace Passage.Web.Services;

/// <summary>
/// A named set of Fountain token colours, one palette per theme. Keys are the
/// token names app.css uses in <c>--syntax-{token}</c>.
/// </summary>
public sealed record SyntaxPreset(
    string Id,
    string Name,
    IReadOnlyDictionary<string, string> Dark,
    IReadOnlyDictionary<string, string> Light);

/// <summary>
/// What the browser remembers: which preset is chosen, and the user's own
/// colours when that choice is <see cref="SyntaxSchemes.CustomId"/>. Kept
/// per-browser in localStorage by passage.js, like the theme.
/// </summary>
public sealed class SyntaxSchemeState
{
    public string Preset { get; set; } = SyntaxSchemes.Classic.Id;
    public Dictionary<string, string>? CustomDark { get; set; }
    public Dictionary<string, string>? CustomLight { get; set; }
}

public static partial class SyntaxSchemes
{
    public const string CustomId = "custom";

    /// <summary>Token name → label shown next to its colour picker.</summary>
    public static readonly IReadOnlyList<(string Token, string Label, string Sample)> Tokens =
    [
        ("scene", "Scene heading", "INT."),
        ("character", "Character", "NAME"),
        ("dialogue", "Dialogue", "Aa"),
        ("paren", "Parenthetical", "(beat)"),
        ("transition", "Transition", "CUT TO:"),
        ("section", "Section", "#"),
        ("synopsis", "Synopsis", "="),
        ("note", "Note", "[[ ]]")
    ];

    // Palettes are listed in token order: scene, character, dialogue, paren,
    // transition, section, synopsis, note. "Classic" mirrors the values in
    // app.css and must be kept in step with it.
    public static readonly IReadOnlyList<SyntaxPreset> Presets =
    [
        Preset("classic", "Classic",
            ["#6797FF", "#C05587", "#CEB2C9", "#8A8A85", "#8A6FC9", "#4FC3F7", "#FFB74D", "#81C784"],
            ["#2F55B8", "#94306A", "#3C3C39", "#757570", "#5B4392", "#1D6FA5", "#9A6A1C", "#3E7D44"]),
        Preset("solarized", "Solarized",
            ["#268BD2", "#D33682", "#93A1A1", "#839496", "#6C71C4", "#2AA198", "#B58900", "#859900"],
            ["#1E6FA8", "#B0286B", "#586E75", "#839496", "#5A5FA8", "#1F8A80", "#8A6800", "#6B7B00"]),
        Preset("arctic", "Arctic",
            ["#88C0D0", "#B48EAD", "#D8DEE9", "#7B88A1", "#81A1C1", "#5E81AC", "#EBCB8B", "#A3BE8C"],
            ["#3B6E8F", "#7E5A7A", "#3B4252", "#6C7386", "#4C6A8F", "#2E5A87", "#9A7420", "#5C7D3E"]),
        Preset("ember", "Ember",
            ["#F0A868", "#E8734A", "#E6D5C3", "#A08C7A", "#C98A5A", "#FFD180", "#F5B942", "#B5C98A"],
            ["#B35A1E", "#A63D1A", "#3E332B", "#857262", "#8A5324", "#9A6200", "#8C6A00", "#5B7A2A"]),
        Preset("ink", "Ink",
            ["#F4F4F2", "#D8D8D5", "#C4C4C0", "#8A8A85", "#B2B2AD", "#FAFAF8", "#A0A09B", "#7C7C77"],
            ["#111110", "#2E2E2C", "#3C3C39", "#757570", "#4A4A47", "#000000", "#5C5C58", "#6E6E69"]),
        Preset("contrast", "High Contrast",
            ["#00D4FF", "#FF6EC7", "#FFFFFF", "#BFBFBF", "#D0A8FF", "#FFE600", "#FFA033", "#6CFF8A"],
            ["#0043B0", "#B8006E", "#000000", "#4D4D4D", "#5A00A8", "#8A5A00", "#A34A00", "#006B1E"]),
        // The presets below spread the eight tokens across the whole hue wheel
        // rather than one family, so neighbouring elements never share a colour.
        Preset("primary", "Primary",
            ["#4D8DFF", "#FF4D4D", "#FFFFFF", "#A0A0A0", "#FF9F1A", "#FFE14D", "#B65CFF", "#3DDC84"],
            ["#1F5FD6", "#D41F1F", "#000000", "#5A5A5A", "#C25E00", "#8A7400", "#7B22C9", "#12803F"]),
        Preset("monokai", "Monokai",
            ["#66D9EF", "#F92672", "#F8F8F2", "#A59F85", "#AE81FF", "#A6E22E", "#FD971F", "#E6DB74"],
            ["#0E7C9A", "#C4104F", "#272822", "#6E6A55", "#6F3FC7", "#4A7A00", "#B35900", "#7A6A00"]),
        Preset("dracula", "Dracula",
            ["#8BE9FD", "#FF79C6", "#F8F8F2", "#9AA5D0", "#BD93F9", "#50FA7B", "#FFB86C", "#F1FA8C"],
            ["#0A7E9A", "#B8247F", "#282A36", "#5B6590", "#6B3FBF", "#1B8A3E", "#A85A00", "#6F7A00"]),
        Preset("gruvbox", "Gruvbox",
            ["#83A598", "#FB4934", "#EBDBB2", "#A89984", "#D3869B", "#FABD2F", "#FE8019", "#B8BB26"],
            ["#327482", "#B3140F", "#3C3836", "#7C6F64", "#9D2A6B", "#A66A00", "#B04500", "#6B7000"]),
        Preset("jewel", "Jewel",
            ["#3AA0FF", "#FF4F6D", "#E8E4D8", "#9C9A90", "#B56BFF", "#19D3A2", "#FFB020", "#C8E04A"],
            ["#1665C1", "#C21E40", "#2A2824", "#6E6C62", "#7B2FD0", "#007A62", "#A8620A", "#5F7A00"]),
        // Okabe-Ito: chosen to stay distinguishable with common colour blindness.
        Preset("colourblind", "Colour-blind Safe",
            ["#56B4E9", "#E69F00", "#F0F0F0", "#999999", "#CC79A7", "#F0E442", "#FF7A33", "#00B386"],
            ["#0072B2", "#B36B00", "#111111", "#666666", "#A0457D", "#7A6F00", "#B84A00", "#007A5A"])
    ];

    public static SyntaxPreset Classic => Presets[0];

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    [GeneratedRegex("^#[0-9A-Fa-f]{6}$")]
    private static partial Regex HexColour();

    public static bool IsValidColour(string? value) => value is not null && HexColour().IsMatch(value);

    public static SyntaxPreset? FindPreset(string id) =>
        Presets.FirstOrDefault(preset => string.Equals(preset.Id, id, StringComparison.Ordinal));

    /// <summary>
    /// The colours the page should show for each theme. A custom set that is
    /// missing a token, or holds something that is not a colour, falls back to
    /// Classic for that token so a stale or hand-edited blob still renders.
    /// </summary>
    public static (IReadOnlyDictionary<string, string> Dark, IReadOnlyDictionary<string, string> Light) Resolve(
        SyntaxSchemeState state)
    {
        if (state.Preset != CustomId)
        {
            var preset = FindPreset(state.Preset) ?? Classic;
            return (preset.Dark, preset.Light);
        }

        return (Complete(state.CustomDark, Classic.Dark), Complete(state.CustomLight, Classic.Light));
    }

    private static IReadOnlyDictionary<string, string> Complete(
        Dictionary<string, string>? custom, IReadOnlyDictionary<string, string> fallback)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (token, _, _) in Tokens)
        {
            result[token] = custom is not null && custom.TryGetValue(token, out var value) && IsValidColour(value)
                ? value
                : fallback[token];
        }

        return result;
    }

    public static string Serialize(SyntaxSchemeState state) => JsonSerializer.Serialize(state, JsonOptions);

    /// <summary>Reads a stored blob; anything unreadable yields the default state.</summary>
    public static SyntaxSchemeState Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return new SyntaxSchemeState();
        }

        try
        {
            var state = JsonSerializer.Deserialize<SyntaxSchemeState>(json, JsonOptions) ?? new SyntaxSchemeState();
            if (state.Preset != CustomId && FindPreset(state.Preset) is null)
            {
                state.Preset = Classic.Id;
            }

            return state;
        }
        catch (JsonException)
        {
            return new SyntaxSchemeState();
        }
    }

    private static SyntaxPreset Preset(string id, string name, string[] dark, string[] light) =>
        new(id, name, Palette(dark), Palette(light));

    private static Dictionary<string, string> Palette(string[] colours)
    {
        var palette = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var i = 0; i < Tokens.Count; i++)
        {
            palette[Tokens[i].Token] = colours[i];
        }

        return palette;
    }
}
