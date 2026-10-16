using System.Text.RegularExpressions;

namespace Passage.Export;

/// <summary>
/// A page header or footer: left, centre and right slots. All three empty means "off".
/// </summary>
public sealed record PageSlots(string Left, string Center, string Right)
{
    public static PageSlots Off { get; } = new(string.Empty, string.Empty, string.Empty);

    public bool IsEmpty => Left.Length == 0 && Center.Length == 0 && Right.Length == 0;

    public PageSlots Resolve(int pageNumber, int totalPages)
    {
        string Fill(string slot) => slot
            .Replace("%pp", totalPages.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("%p", pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
        return new PageSlots(Fill(Left), Fill(Center), Fill(Right));
    }
}

/// <summary>
/// <c>{{HEADER: a | b | c}}</c> / <c>{{FOOTER: …}}</c> lines and <c>{{%m}}</c> markers. Used only by the opt-in
/// pagination path (<see cref="PaginationOptions.Directives"/>); without options they print
/// as the action lines the parser reads them as.
/// </summary>
public static partial class PageDirectives
{
    public readonly record struct Directive(bool IsFooter, PageSlots Slots);

    [GeneratedRegex(@"^\s*\{\{\s*(HEADER|FOOTER)\s*:(.*)\}\}\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex DirectivePattern();

    // {{%m}} or {{%m: name}}, with the spaces before it (after it at a line start), so
    // "a {{%m}} b" prints "a b" and "{{%m}} b" prints "b".
    [GeneratedRegex(@"^[ \t]*\{\{\s*%m\s*(?::[^}]*)?\}\}[ \t]*|[ \t]*\{\{\s*%m\s*(?::[^}]*)?\}\}", RegexOptions.Multiline)]
    private static partial Regex MarkerPattern();

    [GeneratedRegex(@"\{\{\s*%m\s*:\s*([^}]*?)\s*\}\}")]
    private static partial Regex NamedMarkerPattern();

    /// <summary>The names of the named markers in a line; <c>{{%m}}</c> has none.</summary>
    public static IEnumerable<string> MarkerNames(string line) => line.Contains("{{", StringComparison.Ordinal)
        ? NamedMarkerPattern().Matches(line).Select(match => match.Groups[1].Value).Where(name => name.Length > 0)
        : [];

    /// <summary>A line holding nothing but markers.</summary>
    public static bool IsMarkerLine(string line) => line.Contains("{{", StringComparison.Ordinal)
        && MarkerPattern().IsMatch(line) && MarkerPattern().Replace(line, string.Empty).Trim().Length == 0;

    /// <summary>Removes markers, which never print, from text.</summary>
    public static string StripMarkers(string text) => text.Contains("{{", StringComparison.Ordinal)
        ? MarkerPattern().Replace(text, string.Empty)
        : text;

    /// <summary>
    /// Parses a whole line. One slot goes right in a header and centre in a footer; two slots
    /// are left and right; three are left, centre, right. <c>%none</c> turns it off.
    /// </summary>
    public static bool TryParse(string line, out Directive directive)
    {
        directive = default;
        var match = DirectivePattern().Match(line);
        if (!match.Success) return false;

        var isFooter = match.Groups[1].Value.Equals("FOOTER", StringComparison.OrdinalIgnoreCase);
        var body = match.Groups[2].Value.Trim();
        if (body.Equals("%none", StringComparison.OrdinalIgnoreCase))
        {
            directive = new Directive(isFooter, PageSlots.Off);
            return true;
        }

        var slots = body.Split('|').Select(slot => slot.Trim()).ToArray();
        var parsed = slots.Length switch
        {
            1 => isFooter ? new PageSlots(string.Empty, slots[0], string.Empty) : new PageSlots(string.Empty, string.Empty, slots[0]),
            2 => new PageSlots(slots[0], string.Empty, slots[1]),
            _ => new PageSlots(slots[0], slots[1], slots[2])
        };
        directive = new Directive(isFooter, parsed);
        return true;
    }
}
