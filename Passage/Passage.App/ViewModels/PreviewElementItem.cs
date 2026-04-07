using Passage.Parser;

namespace Passage.App.ViewModels;

public sealed class PreviewElementItem
{
    public PreviewElementItem(
        ScreenplayElementType elementType,
        string text,
        string rawText,
        int startLine,
        int endLine)
    {
        ElementType = elementType;
        Text = text;
        RawText = rawText;
        StartLine = startLine;
        EndLine = endLine;
    }

    public ScreenplayElementType ElementType { get; }

    public string Text { get; }

    public string RawText { get; }

    public int StartLine { get; }

    public int EndLine { get; }
}
