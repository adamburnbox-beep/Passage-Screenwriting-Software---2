namespace Passage.App.ViewModels;

public sealed class SceneHeadingItem
{
    public SceneHeadingItem(int lineNumber, string text, string rawText)
    {
        LineNumber = lineNumber;
        Text = text;
        RawText = rawText;
    }

    public int LineNumber { get; }

    public string Text { get; }

    public string RawText { get; }
}
