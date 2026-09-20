using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Passage.Parser;

/// <summary>
/// A single-bracket placeholder in the script text, such as
/// <c>[SOMETHING forces her to stay]</c>. Positions are 0-based into the
/// original text; <see cref="Text"/> is the span exactly as written, brackets
/// included, which is what a fill replaces.
/// </summary>
public sealed record Bracket(int LineIndex, int Column, string Text, int Rank)
{
    public int LineNumber => LineIndex + 1;

    public string Content => Text[1..^1].Trim();
}

/// <summary>
/// Finds <c>[...]</c> placeholders. Pure string logic in the shape of
/// <see cref="BeatBoardText"/>: no UI types, fully testable.
///
/// Single brackets are also legal prose — <c>[sic]</c>, a citation — so the
/// scanner ranks rather than filters (docs/SLATE-PLAN.md, decision C): a
/// bracket whose content is all caps or opens with a placeholder keyword gets
/// rank 0 and sorts first; everything else is rank 1 and still listed. Every
/// line is scanned, not just Action and Dialogue, because a bracketed cue
/// like <c>[SOMEONE]</c> parses as a character and is still a placeholder.
/// Fountain's <c>[[notes]]</c> and boneyard are masked first so they cannot
/// double-match.
/// </summary>
public static class BracketScanner
{
    private static readonly string[] Keywords = ["SOMETHING", "CONDITION", "TODO"];

    // No nesting and no line breaks inside a bracket; an unclosed one is prose.
    private static readonly Regex BracketRegex = new(
        @"\[([^\[\]\r\n]*)\]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static IReadOnlyList<Bracket> Scan(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return Array.Empty<Bracket>();
        }

        var lines = FountainMarkup.MaskOmissions(text).Split('\n');
        var found = new List<Bracket>();
        for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            foreach (Match match in BracketRegex.Matches(lines[lineIndex]))
            {
                var content = match.Groups[1].Value.Trim();
                if (content.Length == 0)
                {
                    continue;
                }

                found.Add(new Bracket(lineIndex, match.Index, match.Value, RankOf(content)));
            }
        }

        return found
            .OrderBy(bracket => bracket.Rank)
            .ThenBy(bracket => bracket.LineIndex)
            .ThenBy(bracket => bracket.Column)
            .ToList();
    }

    public static int RankOf(string content)
    {
        var firstWord = new string(content.TakeWhile(char.IsLetter).ToArray());
        if (Keywords.Contains(firstWord, StringComparer.OrdinalIgnoreCase))
        {
            return 0;
        }

        return content.Any(char.IsLetter) && !content.Any(char.IsLower) ? 0 : 1;
    }
}
