using System.Text.RegularExpressions;

namespace Passage.Parser;

public static class FountainMarkup
{
    private static readonly Regex NotesRegex = new(
        @"\[\[([\s\S]*?)\]\]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex BoneyardRegex = new(
        @"/\*([\s\S]*?)\*/",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static string ExtractNoteText(string text)
    {
        return ExtractDelimitedText(text, NotesRegex, "[[");
    }

    public static string ExtractBoneyardText(string text)
    {
        return ExtractDelimitedText(text, BoneyardRegex, "/*");
    }

    public static string StripOmissions(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return NotesRegex.Replace(BoneyardRegex.Replace(text, string.Empty), string.Empty);
    }

    private static string ExtractDelimitedText(string text, Regex regex, string openMarker)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var match = regex.Match(text);
        if (match.Success)
        {
            return match.Groups[1].Value.Trim();
        }

        var openIndex = text.IndexOf(openMarker, StringComparison.Ordinal);
        if (openIndex < 0)
        {
            return string.Empty;
        }

        return text[(openIndex + openMarker.Length)..].Trim();
    }
}
