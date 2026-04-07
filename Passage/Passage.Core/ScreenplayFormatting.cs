namespace Passage.Core;

public static class ScreenplayFormatting
{
    public const int ActionWrapChars = 60;
    public const int DialogueWrapChars = 36;
    public const int ParentheticalWrapChars = 30;
    public const int CharacterWrapChars = 24;
    public const int TransitionWrapChars = 24;

    public enum ElementSpacingType
    {
        SceneHeading,
        Section,
        Synopsis,
        Action,
        Character,
        Dialogue,
        Parenthetical,
        Transition,
        Other
    }

    public static int GetBlankLinesBefore(ElementSpacingType elementType)
    {
        return elementType switch
        {
            ElementSpacingType.SceneHeading => 2,
            ElementSpacingType.Section => 2,
            ElementSpacingType.Synopsis => 1,
            ElementSpacingType.Action => 1,
            ElementSpacingType.Character => 1,
            ElementSpacingType.Dialogue => 0,
            ElementSpacingType.Parenthetical => 0,
            ElementSpacingType.Transition => 1,
            _ => 1
        };
    }

    public static int GetBlankLinesAfter(ElementSpacingType elementType)
    {
        return elementType switch
        {
            ElementSpacingType.SceneHeading => 1,
            ElementSpacingType.Section => 1,
            ElementSpacingType.Synopsis => 1,
            ElementSpacingType.Action => 1,
            ElementSpacingType.Character => 0,
            ElementSpacingType.Dialogue => 1,
            ElementSpacingType.Parenthetical => 0,
            ElementSpacingType.Transition => 1,
            _ => 1
        };
    }

    public static int GetGapBlankLines(ElementSpacingType previousType, ElementSpacingType currentType)
    {
        return Math.Max(GetBlankLinesAfter(previousType), GetBlankLinesBefore(currentType));
    }
}
