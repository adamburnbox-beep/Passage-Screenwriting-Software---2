using System;
using System.Collections.Generic;

namespace Passage.Parser;

public static class ScriptSanitizer
{
    public static string ExtractCleanScript(string rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
        {
            return rawText;
        }

        var parser = new FountainParser();
        var parsed = parser.Parse(rawText);
        var lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        var keepLine = new bool[lines.Length];
        for (int i = 0; i < keepLine.Length; i++)
        {
            keepLine[i] = true;
        }

        bool IsNukeElement(ScreenplayElement e) =>
            e.Type == ScreenplayElementType.Synopsis ||
            e.Type == ScreenplayElementType.Note ||
            e.Type == ScreenplayElementType.Boneyard ||
            (e.Type == ScreenplayElementType.Section && e.Level <= 1); // Act and Sequence

        bool IsStandardElement(ScreenplayElement e) =>
            e.Type == ScreenplayElementType.SceneHeading ||
            e.Type == ScreenplayElementType.Transition ||
            e.Type == ScreenplayElementType.Character ||
            e.Type == ScreenplayElementType.Dialogue;

        var droppedElements = new HashSet<ScreenplayElement>();

        // Identify elements to drop unconditionally
        foreach (var el in parsed.Elements)
        {
            if (IsNukeElement(el))
            {
                droppedElements.Add(el);
            }
        }

        // Apply Nested Action Rule
        for (int i = 0; i < parsed.Elements.Count; )
        {
            if (parsed.Elements[i].Type == ScreenplayElementType.Action)
            {
                int startIdx = i;
                while (i < parsed.Elements.Count && parsed.Elements[i].Type == ScreenplayElementType.Action)
                {
                    i++;
                }
                int endIdx = i - 1;

                bool followsNuke = false;
                if (startIdx > 0 && IsNukeElement(parsed.Elements[startIdx - 1]))
                {
                    followsNuke = true;
                }

                bool precedesStandard = false;
                if (endIdx + 1 < parsed.Elements.Count && IsStandardElement(parsed.Elements[endIdx + 1]))
                {
                    precedesStandard = true;
                }

                if (followsNuke && precedesStandard)
                {
                    for (int j = startIdx; j <= endIdx; j++)
                    {
                        droppedElements.Add(parsed.Elements[j]);
                    }
                }
            }
            else
            {
                i++;
            }
        }

        foreach (var el in droppedElements)
        {
            for (int l = el.LineIndex; l <= el.EndLineIndex && l < lines.Length; l++)
            {
                keepLine[l] = false;
            }
        }

        // Reconstruct the layout explicitly maintaining normal empty line breaks.
        var keptLines = new List<string>();
        for (int l = 0; l < lines.Length; l++)
        {
            if (keepLine[l])
            {
                keptLines.Add(lines[l]);
            }
        }

        return string.Join(Environment.NewLine, keptLines);
    }
}
