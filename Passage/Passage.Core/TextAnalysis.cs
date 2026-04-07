namespace Passage.Core;

public static class TextAnalysis
{
    private static readonly string[] SceneHeadingPrefixes =
    [
        "INT",
        "INT.",
        "EXT",
        "EXT.",
        "INT/EXT",
        "INT/EXT.",
        "INT./EXT.",
        "EXT/INT",
        "EXT/INT.",
        "EXT./INT.",
        "I/E",
        "I.E.",
        "EST",
        "EST."
    ];

    public static bool ContainsLetter(ReadOnlySpan<char> value)
    {
        foreach (var ch in value)
        {
            if (char.IsLetter(ch))
            {
                return true;
            }
        }

        return false;
    }

    public static bool ContainsLowercaseLetter(ReadOnlySpan<char> value)
    {
        foreach (var ch in value)
        {
            if (char.IsLower(ch))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsUppercaseLike(ReadOnlySpan<char> value)
    {
        var hasLetter = false;

        foreach (var ch in value)
        {
            if (!char.IsLetter(ch))
            {
                continue;
            }

            hasLetter = true;
            if (!char.IsUpper(ch))
            {
                return false;
            }
        }

        return hasLetter;
    }

    public static bool StartsWithAnyPrefix(ReadOnlySpan<char> value, IReadOnlyList<string> prefixes)
    {
        foreach (var prefix in prefixes)
        {
            if (value.StartsWith(prefix.AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    public static int CountWords(ReadOnlySpan<char> value, int maxWords)
    {
        var wordCount = 0;
        var inWord = false;

        foreach (var ch in value)
        {
            if (char.IsWhiteSpace(ch))
            {
                inWord = false;
                continue;
            }

            if (!inWord)
            {
                wordCount++;
                if (wordCount > maxWords)
                {
                    return wordCount;
                }

                inWord = true;
            }
        }

        return wordCount;
    }

    public static int CountWords(ReadOnlySpan<char> value)
    {
        return CountWords(value, int.MaxValue);
    }

    public static bool LooksLikeSceneHeadingStart(ReadOnlySpan<char> value, bool allowPartialPrefixMatch = false)
    {
        var trimmed = value.TrimStart();
        if (trimmed.Length == 0)
        {
            return false;
        }

        var tokenEnd = trimmed.IndexOfAny([' ', '\t']);
        var token = tokenEnd < 0 ? trimmed : trimmed[..tokenEnd];
        if (token.Length == 0)
        {
            return false;
        }

        foreach (var prefix in SceneHeadingPrefixes)
        {
            var prefixSpan = prefix.AsSpan();
            if (token.Equals(prefixSpan, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (allowPartialPrefixMatch &&
                prefixSpan.StartsWith(token, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsLikelyCharacterCue(ReadOnlySpan<char> value, int maxLength, int maxWords)
    {
        if (value.Length == 0 || value.Length > maxLength)
        {
            return false;
        }

        if (value.IndexOf(':') >= 0)
        {
            return false;
        }

        if (ContainsLowercaseLetter(value) || !ContainsLetter(value))
        {
            return false;
        }

        var wordCount = CountWords(value, maxWords);
        return wordCount > 0 && wordCount <= maxWords && IsUppercaseLike(value);
    }

    public static bool IsLiveCharacterCueCandidate(
        ReadOnlySpan<char> value,
        int maxLength,
        int maxWords,
        int minimumLeadingWordLetterCount = 2)
    {
        var trimmedStart = value.TrimStart();
        if (trimmedStart.Length == 0 ||
            LooksLikeSceneHeadingStart(trimmedStart, allowPartialPrefixMatch: true))
        {
            return false;
        }

        var content = trimmedStart.TrimEnd();
        if (!HasMinimumLeadingWordLetterCount(content, minimumLeadingWordLetterCount) ||
            !IsLikelyCharacterCue(content, maxLength, maxWords))
        {
            return false;
        }

        return !(trimmedStart.Length > content.Length && CountWords(content, 2) == 1);
    }

    private static bool HasMinimumLeadingWordLetterCount(ReadOnlySpan<char> value, int minimumLetterCount)
    {
        if (minimumLetterCount <= 0)
        {
            return true;
        }

        var tokenEnd = value.IndexOfAny([' ', '\t']);
        var token = tokenEnd < 0 ? value : value[..tokenEnd];
        var letterCount = 0;

        foreach (var ch in token)
        {
            if (!char.IsLetter(ch))
            {
                continue;
            }

            letterCount++;
            if (letterCount >= minimumLetterCount)
            {
                return true;
            }
        }

        return false;
    }
}
