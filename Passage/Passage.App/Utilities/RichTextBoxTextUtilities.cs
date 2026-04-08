using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace Passage.App.Utilities;

internal static class RichTextBoxTextUtilities
{
    private const string ParagraphBreak = "\r\n";
    private static readonly ConditionalWeakTable<RichTextBox, DocumentCache> DocumentCaches = new();

    private sealed record ParagraphInfo(Paragraph Paragraph, int Start, int Length, string Text);

    private sealed class DocumentCache
    {
        public DocumentCache(RichTextBox editor)
        {
            Document = editor.Document;
            editor.TextChanged += (_, _) => IsDirty = true;
        }

        public FlowDocument? Document { get; set; }

        public bool IsDirty { get; set; } = true;

        public string PlainText { get; set; } = string.Empty;

        public int TextLength => PlainText.Length;

        public List<ParagraphInfo> Paragraphs { get; } = [];

        public Dictionary<Paragraph, int> ParagraphIndices { get; } = [];
    }

    public static string GetPlainText(RichTextBox editor)
    {
        ArgumentNullException.ThrowIfNull(editor);
        return GetDocumentCache(editor).PlainText;
    }

    public static int GetTextOffset(RichTextBox editor, TextPointer position)
    {
        ArgumentNullException.ThrowIfNull(editor);
        ArgumentNullException.ThrowIfNull(position);

        var cache = GetDocumentCache(editor);
        if (position.CompareTo(editor.Document.ContentStart) <= 0)
        {
            return 0;
        }

        if (position.CompareTo(editor.Document.ContentEnd) >= 0)
        {
            return cache.TextLength;
        }

        var paragraph = position.Paragraph;
        if (paragraph is null || !cache.ParagraphIndices.TryGetValue(paragraph, out var paragraphIndex))
        {
            return cache.TextLength;
        }

        var paragraphInfo = cache.Paragraphs[paragraphIndex];
        var paragraphOffset = GetParagraphOffset(paragraph, position, paragraphInfo.Length);
        return Math.Min(cache.TextLength, paragraphInfo.Start + paragraphOffset);
    }

    public static int GetCaretIndex(RichTextBox editor)
    {
        return GetTextOffset(editor, editor.CaretPosition);
    }

    public static int GetSelectionStart(RichTextBox editor)
    {
        var selectionStart = GetTextOffset(editor, editor.Selection.Start);
        var selectionEnd = GetTextOffset(editor, editor.Selection.End);
        return Math.Min(selectionStart, selectionEnd);
    }

    public static int GetSelectionLength(RichTextBox editor)
    {
        var selectionStart = GetTextOffset(editor, editor.Selection.Start);
        var selectionEnd = GetTextOffset(editor, editor.Selection.End);
        return Math.Abs(selectionEnd - selectionStart);
    }

    public static int GetSelectionAnchorIndex(RichTextBox editor)
    {
        var selectionStart = GetSelectionStart(editor);
        var selectionEnd = selectionStart + GetSelectionLength(editor);
        var caretIndex = GetCaretIndex(editor);

        if (selectionEnd <= selectionStart)
        {
            return caretIndex;
        }

        return caretIndex == selectionStart
            ? selectionEnd
            : selectionStart;
    }

    public static TextPointer GetTextPointerAtOffset(RichTextBox editor, int offset)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var cache = GetDocumentCache(editor);
        var safeOffset = Math.Clamp(offset, 0, cache.TextLength);
        if (safeOffset <= 0)
        {
            return editor.Document.ContentStart;
        }

        if (cache.Paragraphs.Count == 0)
        {
            return editor.Document.ContentEnd;
        }

        for (var paragraphIndex = 0; paragraphIndex < cache.Paragraphs.Count; paragraphIndex++)
        {
            var paragraphInfo = cache.Paragraphs[paragraphIndex];
            var lineEnd = paragraphInfo.Start + paragraphInfo.Length;
            if (safeOffset <= lineEnd)
            {
                return GetParagraphTextPointerAtOffset(paragraphInfo.Paragraph, safeOffset - paragraphInfo.Start, paragraphInfo.Length);
            }

            if (paragraphIndex >= cache.Paragraphs.Count - 1)
            {
                continue;
            }

            var nextLineStart = cache.Paragraphs[paragraphIndex + 1].Start;
            if (safeOffset < nextLineStart)
            {
                var nextParagraph = cache.Paragraphs[paragraphIndex + 1].Paragraph;
                return nextParagraph.ContentStart.GetInsertionPosition(LogicalDirection.Forward)
                    ?? nextParagraph.ContentStart;
            }
        }

        var lastParagraph = cache.Paragraphs[^1];
        return GetParagraphTextPointerAtOffset(lastParagraph.Paragraph, lastParagraph.Length, lastParagraph.Length);
    }

    public static void SetSelection(RichTextBox editor, int anchorIndex, int activeIndex)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var textLength = GetDocumentCache(editor).TextLength;
        var safeAnchor = Math.Clamp(anchorIndex, 0, textLength);
        var safeActive = Math.Clamp(activeIndex, 0, textLength);
        var anchor = GetTextPointerAtOffset(editor, safeAnchor);
        var active = GetTextPointerAtOffset(editor, safeActive);

        editor.Selection.Select(anchor, active);
    }

    public static void Select(RichTextBox editor, int start, int length)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var textLength = GetDocumentCache(editor).TextLength;
        var safeStart = Math.Clamp(start, 0, textLength);
        var safeLength = Math.Clamp(length, 0, textLength - safeStart);
        SetSelection(editor, safeStart, safeStart + safeLength);
    }

    public static string GetSelectedText(RichTextBox editor)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var selectedText = editor.Selection.Text;
        return editor.Selection.End.CompareTo(editor.Document.ContentEnd) == 0
            ? NormalizeTerminalParagraphBreak(selectedText)
            : selectedText;
    }

    public static bool TryGetCharacterIndexFromPoint(RichTextBox editor, Point point, out int characterIndex)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var position = editor.GetPositionFromPoint(point, true);
        if (position is null)
        {
            characterIndex = 0;
            return false;
        }

        characterIndex = GetTextOffset(editor, position);
        return true;
    }

    public static int GetLineCount(RichTextBox editor)
    {
        ArgumentNullException.ThrowIfNull(editor);
        var paragraphCount = GetDocumentCache(editor).Paragraphs.Count;
        return Math.Max(1, paragraphCount);
    }

    public static int GetLineIndexFromCharacterIndex(RichTextBox editor, int characterIndex)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var cache = GetDocumentCache(editor);
        if (cache.Paragraphs.Count == 0)
        {
            return 0;
        }

        var safeCharacterIndex = Math.Clamp(characterIndex, 0, cache.TextLength);
        var low = 0;
        var high = cache.Paragraphs.Count - 1;
        var result = 0;

        while (low <= high)
        {
            var mid = low + ((high - low) / 2);
            if (cache.Paragraphs[mid].Start <= safeCharacterIndex)
            {
                result = mid;
                low = mid + 1;
            }
            else
            {
                high = mid - 1;
            }
        }

        return result;
    }

    public static int GetCharacterIndexFromLineIndex(RichTextBox editor, int lineIndex)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var cache = GetDocumentCache(editor);
        if (lineIndex < 0 || lineIndex >= cache.Paragraphs.Count)
        {
            return -1;
        }

        return cache.Paragraphs[lineIndex].Start;
    }

    public static string GetLineText(RichTextBox editor, int lineIndex)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var cache = GetDocumentCache(editor);
        if (lineIndex < 0 || lineIndex >= cache.Paragraphs.Count)
        {
            return string.Empty;
        }

        return cache.Paragraphs[lineIndex].Text;
    }

    public static void SetPlainText(RichTextBox editor, string text)
    {
        ArgumentNullException.ThrowIfNull(editor);

        var normalizedText = NormalizeLineEndings(text ?? string.Empty);
        var lines = normalizedText.Split('\n', StringSplitOptions.None);
        if (lines.Length == 0)
        {
            lines = [string.Empty];
        }

        editor.Document.Blocks.Clear();
        foreach (var line in lines)
        {
            editor.Document.Blocks.Add(new Paragraph(new Run(line))
            {
                Margin = new Thickness(0),
                TextAlignment = TextAlignment.Left
            });
        }

        if (editor.Document.Blocks.Count == 0)
        {
            editor.Document.Blocks.Add(new Paragraph
            {
                Margin = new Thickness(0),
                TextAlignment = TextAlignment.Left
            });
        }
    }

    private static DocumentCache GetDocumentCache(RichTextBox editor)
    {
        var cache = DocumentCaches.GetValue(editor, static instance => new DocumentCache(instance));
        if (!ReferenceEquals(cache.Document, editor.Document))
        {
            cache.Document = editor.Document;
            cache.IsDirty = true;
        }

        if (!cache.IsDirty)
        {
            return cache;
        }

        cache.IsDirty = false;
        cache.Paragraphs.Clear();
        cache.ParagraphIndices.Clear();

        var builder = new StringBuilder();
        var paragraphStart = 0;
        var paragraphIndex = 0;

        foreach (var paragraph in editor.Document.Blocks.OfType<Paragraph>())
        {
            var paragraphText = GetParagraphText(paragraph);
            if (paragraphIndex > 0)
            {
                builder.Append(ParagraphBreak);
            }

            builder.Append(paragraphText);
            cache.Paragraphs.Add(new ParagraphInfo(paragraph, paragraphStart, paragraphText.Length, paragraphText));
            cache.ParagraphIndices[paragraph] = paragraphIndex;
            paragraphStart += paragraphText.Length + ParagraphBreak.Length;
            paragraphIndex++;
        }

        cache.PlainText = builder.ToString();
        return cache;
    }

    private static string GetParagraphText(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        AppendInlineText(paragraph.Inlines, builder);
        return builder.ToString();
    }

    private static void AppendInlineText(InlineCollection inlines, StringBuilder builder)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case Run run:
                    builder.Append(run.Text);
                    break;
                case LineBreak:
                    builder.Append('\n');
                    break;
                case Span span:
                    AppendInlineText(span.Inlines, builder);
                    break;
                case InlineUIContainer container:
                    if (container.Child is TextBlock tb)
                    {
                        builder.Append(tb.Text);
                    }
                    break;
            }
        }
    }

    private static int GetParagraphOffset(Paragraph paragraph, TextPointer position, int paragraphLength)
    {
        if (position.CompareTo(paragraph.ContentStart) <= 0)
        {
            return 0;
        }

        if (position.CompareTo(paragraph.ContentEnd) >= 0)
        {
            return paragraphLength;
        }

        var text = new TextRange(paragraph.ContentStart, position).Text;
        return Math.Min(paragraphLength, NormalizeTerminalParagraphBreak(text).Length);
    }

    private static TextPointer GetParagraphTextPointerAtOffset(Paragraph paragraph, int offset, int paragraphLength)
    {
        var safeOffset = Math.Clamp(offset, 0, paragraphLength);
        if (safeOffset <= 0)
        {
            return paragraph.ContentStart.GetInsertionPosition(LogicalDirection.Forward)
                ?? paragraph.ContentStart;
        }

        var remaining = safeOffset;
        var position = paragraph.ContentStart.GetInsertionPosition(LogicalDirection.Forward)
            ?? paragraph.ContentStart;

        while (position is not null && position.CompareTo(paragraph.ContentEnd) <= 0)
        {
            if (position.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
            {
                var runText = position.GetTextInRun(LogicalDirection.Forward);
                if (remaining <= runText.Length)
                {
                    return position.GetPositionAtOffset(remaining, LogicalDirection.Forward)
                        ?? paragraph.ContentEnd.GetInsertionPosition(LogicalDirection.Backward)
                        ?? paragraph.ContentEnd;
                }

                remaining -= runText.Length;
                position = position.GetPositionAtOffset(runText.Length, LogicalDirection.Forward);
                continue;
            }

            var nextPosition = position.GetNextContextPosition(LogicalDirection.Forward);
            if (nextPosition is null)
            {
                break;
            }

            position = nextPosition;
        }

        return paragraph.ContentEnd.GetInsertionPosition(LogicalDirection.Backward)
            ?? paragraph.ContentEnd;
    }

    private static string NormalizeLineEndings(string text)
    {
        return text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
    }

    private static string NormalizeTerminalParagraphBreak(string text)
    {
        return text.EndsWith(ParagraphBreak, StringComparison.Ordinal)
            ? text[..^ParagraphBreak.Length]
            : text;
    }
}
