using System.Text;
using Passage.Core;
using Passage.Parser;
using System.IO;

namespace Passage.Export;

public sealed class ScreenplayExporter : IExporter
{
    public string DisplayName => "Text";

    public string DefaultExtension => ".txt";

    public void Export(ParsedScreenplay screenplay, string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        File.WriteAllText(filePath, GetExportText(screenplay), Encoding.UTF8);
    }

    public string GetExportText(ParsedScreenplay screenplay)
    {
        ArgumentNullException.ThrowIfNull(screenplay);

        var builder = new StringBuilder();

        foreach (var entry in screenplay.TitlePage.Entries)
        {
            builder.AppendLine($"{entry.Label}: {entry.Value}");
        }

        if (screenplay.TitlePage.Entries.Count > 0 && screenplay.Elements.Count > 0)
        {
            builder.AppendLine();
        }

        if (screenplay.Elements.Count > 0)
        {
            var firstSpacingType = ToSpacingType(screenplay.Elements[0].Type);
            AppendBlankLines(builder, ScreenplayFormatting.GetBlankLinesBefore(firstSpacingType));

            ScreenplayElement? previous = null;
            foreach (var element in screenplay.Elements)
            {
                AppendElement(builder, element, previous);
                previous = element;
            }

            AppendBlankLines(builder, ScreenplayFormatting.GetBlankLinesAfter(ToSpacingType(screenplay.Elements[^1].Type)));
        }

        return builder.ToString();
    }

    private static void AppendElement(StringBuilder builder, ScreenplayElement element, ScreenplayElement? previous)
    {
        var currentSpacingType = ToSpacingType(element.Type);
        var blankLines = previous is null
            ? 0
            : ScreenplayFormatting.GetGapBlankLines(ToSpacingType(previous.Type), currentSpacingType);

        AppendBlankLines(builder, blankLines);

        switch (element)
        {
            case SceneHeadingElement sceneHeading:
                AppendBlock(builder, sceneHeading.Text.ToUpperInvariant());
                break;
            case ActionElement action:
                AppendBlock(builder, action.Text);
                break;
            case CharacterElement character:
                AppendBlock(builder, CenterText(character.CharacterName));
                break;
            case ParentheticalElement parenthetical:
                AppendBlock(builder, IndentText($"({parenthetical.Text})", 2));
                break;
            case DialogueElement dialogue:
                AppendBlock(builder, IndentText(dialogue.Text, 4));
                break;
            case TransitionElement transition:
                AppendBlock(builder, AlignRight(transition.Text));
                break;
            case SectionElement section:
                AppendBlock(builder, $"{new string('#', section.SectionDepth)} {section.Text}");
                break;
            case SynopsisElement synopsis:
                AppendBlock(builder, $"= {synopsis.Text}");
                break;
            default:
                AppendBlock(builder, element.Text);
                break;
        }
    }

    private static void AppendBlock(StringBuilder builder, string text)
    {
        builder.AppendLine(text);
    }

    private static void AppendBlankLines(StringBuilder builder, int count)
    {
        for (var index = 0; index < count; index++)
        {
            builder.AppendLine();
        }
    }

    private static string IndentText(string text, int spaces)
    {
        return $"{new string(' ', spaces)}{text}";
    }

    private static string CenterText(string text)
    {
        var line = text.Trim();
        var padding = Math.Max(0, (60 - line.Length) / 2);
        return $"{new string(' ', padding)}{line}";
    }

    private static string AlignRight(string text)
    {
        var line = text.Trim();
        var padding = Math.Max(0, 60 - line.Length);
        return $"{new string(' ', padding)}{line}";
    }

    private static ScreenplayFormatting.ElementSpacingType ToSpacingType(ScreenplayElementType elementType)
    {
        return elementType switch
        {
            ScreenplayElementType.SceneHeading => ScreenplayFormatting.ElementSpacingType.SceneHeading,
            ScreenplayElementType.Action => ScreenplayFormatting.ElementSpacingType.Action,
            ScreenplayElementType.Character => ScreenplayFormatting.ElementSpacingType.Character,
            ScreenplayElementType.Dialogue => ScreenplayFormatting.ElementSpacingType.Dialogue,
            ScreenplayElementType.Parenthetical => ScreenplayFormatting.ElementSpacingType.Parenthetical,
            ScreenplayElementType.Transition => ScreenplayFormatting.ElementSpacingType.Transition,
            _ => ScreenplayFormatting.ElementSpacingType.Other
        };
    }
}
