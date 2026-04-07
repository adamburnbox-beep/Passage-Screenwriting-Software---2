using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using Passage.App.Utilities;
using Passage.Core;
using Passage.Parser;

namespace Passage.App.Visuals;

public sealed class EditorCueAdorner : Adorner
{
    private const double DipsPerInch = 96.0;
    private const double ScreenplayPageWidth = 8.5 * DipsPerInch;
    private const double ScreenplayPageHeight = 11.0 * DipsPerInch;
    private const double ScreenplayPageTopMargin = 1.0 * DipsPerInch;
    private const double ScreenplayPageBottomMargin = 1.0 * DipsPerInch;
    private const double ScreenplayPageRightMargin = 1.0 * DipsPerInch;
    private const double ScreenplayPageLeftMargin = 1.5 * DipsPerInch;
    private const double DialogueIndent = 1.0 * DipsPerInch;
    private const double CharacterIndent = 2.0 * DipsPerInch;
    private const double ParentheticalIndent = 1.5 * DipsPerInch;
    private const double PageGap = 26.0;
    private const double PageShadowOffset = 6.0;
    private const int InteractiveLayoutRefreshDelayMilliseconds = 60;
    private const int DeferredLayoutRefreshDelayMilliseconds = 300;
    private static readonly Regex MarkdownBoldRegex = new(@"\*\*(.+?)\*\*", RegexOptions.Compiled | RegexOptions.Multiline);
    private static readonly Regex IdTagStripRegex = new(@"\s*\[\[id:[a-f\d\-]+\]\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly record struct FormattingSpan(int Start, int Length);

    private sealed record SourceLineInfo(int LineIndex, int CharacterIndex, string Text);

    private sealed record WrappedLine(int StartIndex, int Length, string Text);

    private sealed record WrappedRow(
        int StartIndex,
        int Length,
        string Text,
        Point Origin,
        double Width,
        FormattedText FormattedText);

    private sealed record LineLayout(
        int SourceLineIndex,
        int CharacterStartIndex,
        int CharacterLength,
        ScreenplayElementType ScreenplayType,
        string VisualText,
        IReadOnlyList<WrappedRow> Rows,
        PageLayout Page,
        Rect Bounds,
        double RowHeight)
    {
        public double Top => Bounds.Top;

        public double Bottom => Bounds.Bottom;
    }

    private sealed record LayoutBlock(
        ScreenplayElementType ScreenplayType,
        IReadOnlyList<SourceLineInfo> SourceLines,
        bool IsSpacer)
    {
        public bool IsSyntheticSpacer => IsSpacer && SourceLines.All(line => line.LineIndex < 0);

        public double Height(double rowHeight)
        {
            if (SourceLines.Count == 0)
            {
                return rowHeight;
            }

            return SourceLines.Count * rowHeight;
        }
    }

    private sealed record PageLayout(int PageNumber, Rect PageRect, Rect BodyRect);

    private sealed record PageFlowResult(
        IReadOnlyList<PageLayout> Pages,
        IReadOnlyList<LineLayout> Lines,
        IReadOnlyDictionary<int, LineLayout> LineLookup,
        double DocumentHeight)
    {
        public static PageFlowResult Empty { get; } = new(
            Array.Empty<PageLayout>(),
            Array.Empty<LineLayout>(),
            new Dictionary<int, LineLayout>(),
            0.0);
    }

    private readonly Func<ParsedScreenplay> _getParsedScreenplay;
    private readonly Func<IReadOnlyDictionary<int, ScreenplayElementType>> _getLineTypeOverrides;
    private readonly Func<double> _getZoomScale;
    private readonly Func<bool> _useScreenplayLayout;
    private readonly FountainParser _interactiveParser = new();
    private readonly RichTextBox _editor;
    private readonly ScrollViewer _scrollHost;
    private readonly DispatcherTimer _caretTimer;
    private readonly DispatcherTimer _layoutRefreshTimer;

    private bool _isFormattingEnabled = true;
    private bool _isCaretVisible = true;
    private string? _lastLayoutText;
    private double _lastLayoutWidth = -1.0;
    private double _lastLayoutFontSize = -1.0;
    private ParsedScreenplay? _lastLayoutParsed;
    private bool _layoutRefreshPending;
    private bool _themeBrushesDirty = true;
    private bool _visualRefreshPending;
    private PageFlowResult _pageFlow = PageFlowResult.Empty;
    private string _currentText = string.Empty;
    private int _selectionStart;
    private int _selectionLength;
    private int _caretIndex;
    private IReadOnlyList<SourceLineInfo> _sourceLines = BuildSourceLines(string.Empty);
    private IReadOnlyList<SourceLineInfo>? _pendingSourceLines;

    private static readonly Brush FallbackPageFillBrush = Freeze(new SolidColorBrush(Colors.White));
    private static readonly Brush FallbackPageShadowBrush = Freeze(new SolidColorBrush(Color.FromArgb(34, 0, 0, 0)));
    private static readonly Brush FallbackEditorTextBrush = Freeze(new SolidColorBrush(Color.FromRgb(24, 21, 18)));
    private static readonly Brush FallbackSelectionBrush = Freeze(new SolidColorBrush(Color.FromArgb(88, 108, 129, 156)));
    private static readonly Brush FallbackCaretBrush = Freeze(new SolidColorBrush(Color.FromRgb(24, 21, 18)));
    private static readonly Brush FallbackPageBorderBrush = Freeze(new SolidColorBrush(Color.FromArgb(42, 0, 0, 0)));

    private Brush _pageFillBrush = FallbackPageFillBrush;
    private Brush _pageShadowBrush = FallbackPageShadowBrush;
    private Brush _editorTextBrush = FallbackEditorTextBrush;
    private Brush _selectionBrush = FallbackSelectionBrush;
    private Brush _caretBrush = FallbackCaretBrush;
    private Brush _pageBorderBrush = FallbackPageBorderBrush;

    public EditorCueAdorner(
        RichTextBox editor,
        ScrollViewer scrollHost,
        Func<ParsedScreenplay> getParsedScreenplay,
        Func<IReadOnlyDictionary<int, ScreenplayElementType>> getLineTypeOverrides,
        Func<double> getZoomScale,
        Func<bool> useScreenplayLayout)
        : base(editor)
    {
        _editor = editor;
        _scrollHost = scrollHost;
        _getParsedScreenplay = getParsedScreenplay;
        _getLineTypeOverrides = getLineTypeOverrides;
        _getZoomScale = getZoomScale;
        _useScreenplayLayout = useScreenplayLayout;
        IsHitTestVisible = false;
        _currentText = RichTextBoxTextUtilities.GetPlainText(_editor);
        _sourceLines = BuildSourceLines(_currentText);
        RefreshSelectionState();

        _caretTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(530)
        };
        _layoutRefreshTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(DeferredLayoutRefreshDelayMilliseconds)
        };
        _caretTimer.Tick += (_, _) =>
        {
            if (!_isFormattingEnabled)
            {
                return;
            }

            if (!_editor.IsKeyboardFocused || GetSelectionLength() > 0)
            {
                _isCaretVisible = true;
                return;
            }

            _isCaretVisible = !_isCaretVisible;
            ScheduleVisualRefresh();
        };
        _layoutRefreshTimer.Tick += (_, _) => ApplyPendingLayoutRefresh();

        _editor.GotKeyboardFocus += (_, _) =>
        {
            if (!_isFormattingEnabled)
            {
                return;
            }

            _isCaretVisible = true;
            _caretTimer.Start();
            ScheduleVisualRefresh();
        };
        _editor.LostKeyboardFocus += (_, _) =>
        {
            if (!_isFormattingEnabled)
            {
                return;
            }

            _isCaretVisible = false;
            _caretTimer.Stop();
            ScheduleVisualRefresh();
        };
        _editor.SelectionChanged += (_, _) =>
        {
            if (!_isFormattingEnabled)
            {
                return;
            }

            RefreshSelectionState();
            _isCaretVisible = true;
            ScheduleVisualRefresh();
        };
        _editor.TextChanged += (_, _) =>
        {
            RefreshTextState();
            RefreshSelectionState();
            _isCaretVisible = true;
            ScheduleLayoutRefresh(
                _pendingSourceLines ?? _sourceLines,
                immediate: _editor.IsKeyboardFocusWithin);
            ScheduleVisualRefresh();
        };
    }

    public event EventHandler? LayoutChanged;

    public bool HasPendingLayoutRefresh => _layoutRefreshPending;

    public void NotifyParsedLayoutChanged(bool immediate = false)
    {
        if (!_isFormattingEnabled)
        {
            return;
        }

        ScheduleLayoutRefresh(_pendingSourceLines ?? _sourceLines, immediate);
        ScheduleVisualRefresh();
    }

    public void SetFormattingEnabled(bool isEnabled)
    {
        if (_isFormattingEnabled == isEnabled)
        {
            return;
        }

        _isFormattingEnabled = isEnabled;
        _layoutRefreshTimer.Stop();
        _layoutRefreshPending = false;
        _pendingSourceLines = null;
        _lastLayoutText = null;
        _lastLayoutWidth = -1.0;
        _lastLayoutFontSize = -1.0;
        _lastLayoutParsed = null;
        _themeBrushesDirty = true;
        RefreshTextState();
        _sourceLines = _pendingSourceLines ?? BuildSourceLines(_currentText);
        _pendingSourceLines = null;
        RefreshSelectionState();

        if (isEnabled)
        {
            _isCaretVisible = _editor.IsKeyboardFocused;

            if (_editor.IsKeyboardFocused)
            {
                _caretTimer.Start();
            }
        }
        else
        {
            _isCaretVisible = false;
            _caretTimer.Stop();
            _pageFlow = PageFlowResult.Empty;
        }

        ScheduleVisualRefresh();
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        UpdateLayoutCache();
        DrawPageChrome(drawingContext);

        if (!_isFormattingEnabled)
        {
            return;
        }

        if (_pageFlow.Lines.Count == 0)
        {
            return;
        }

        var visibleBounds = GetVisibleBounds();
        var visibleLines = GetVisibleLineLayouts(visibleBounds);

        DrawSelection(drawingContext, visibleLines);

        foreach (var lineLayout in visibleLines)
        {
            foreach (var row in lineLayout.Rows)
            {
                drawingContext.DrawText(row.FormattedText, row.Origin);
            }
        }

        DrawCaret(drawingContext);
    }

    public double GetDocumentHeight()
    {
        UpdateLayoutCache();
        return Math.Max(_pageFlow.DocumentHeight, GetMinimumDocumentHeight());
    }

    public bool TryGetCaretRect(out Rect caretRect)
    {
        caretRect = Rect.Empty;

        if (!_isFormattingEnabled)
        {
            return false;
        }

        UpdateLayoutCache();

        var caretIndex = GetCaretIndex();
        if (caretIndex < 0)
        {
            return false;
        }

        var lineIndex = GetLineIndexFromCharacterIndex(caretIndex);
        if (lineIndex < 0)
        {
            return false;
        }

        var lineLayout = GetRenderableLineLayout(lineIndex);
        if (lineLayout is null)
        {
            return false;
        }

        caretRect = BuildCaretRect(lineLayout, caretIndex);
        return !caretRect.IsEmpty;
    }

    public bool TryGetCharacterIndexFromPoint(Point point, out int characterIndex)
    {
        characterIndex = 0;

        if (!_isFormattingEnabled)
        {
            return false;
        }

        if (_layoutRefreshPending &&
            _pendingSourceLines is not null &&
            _pendingSourceLines.Count != _sourceLines.Count)
        {
            ApplyPendingLayoutRefresh();
        }
        else
        {
            UpdateLayoutCache();
        }

        if (_pageFlow.Lines.Count == 0)
        {
            return false;
        }

        var lineLayout = GetClosestLineLayout(point);
        if (lineLayout is null)
        {
            return false;
        }

        characterIndex = GetCharacterIndexFromPoint(lineLayout, point);
        return true;
    }

    private string GetEditorText()
    {
        return _currentText;
    }

    private int GetCaretIndex()
    {
        return _caretIndex;
    }

    private int GetLineIndexFromCharacterIndex(int characterIndex)
    {
        var sourceLines = _pendingSourceLines ?? _sourceLines;
        if (sourceLines.Count == 0)
        {
            return 0;
        }

        var safeCharacterIndex = Math.Clamp(characterIndex, 0, _currentText.Length);
        var low = 0;
        var high = sourceLines.Count - 1;
        var result = 0;

        while (low <= high)
        {
            var mid = low + ((high - low) / 2);
            if (sourceLines[mid].CharacterIndex <= safeCharacterIndex)
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

    private int GetSelectionStart()
    {
        return _selectionStart;
    }

    private int GetSelectionLength()
    {
        return _selectionLength;
    }

    private void RefreshTextState()
    {
        _currentText = RichTextBoxTextUtilities.GetPlainText(_editor);
        _pendingSourceLines = BuildSourceLines(_currentText);
    }

    private void RefreshSelectionState()
    {
        var selectionStartOffset = RichTextBoxTextUtilities.GetTextOffset(_editor, _editor.Selection.Start);
        var selectionEndOffset = RichTextBoxTextUtilities.GetTextOffset(_editor, _editor.Selection.End);
        _selectionStart = Math.Min(selectionStartOffset, selectionEndOffset);
        _selectionLength = Math.Abs(selectionEndOffset - selectionStartOffset);
        _caretIndex = RichTextBoxTextUtilities.GetTextOffset(_editor, _editor.CaretPosition);
    }

    private void UpdateLayoutCache()
    {
        RefreshThemeBrushes();
        var themeBrushesDirty = _themeBrushesDirty;

        var currentText = _currentText;
        var currentWidth = _editor.ActualWidth;
        var currentFontSize = _editor.FontSize;
        var pendingSourceLines = _pendingSourceLines ?? _sourceLines;
        var requiresStructuralRefresh = _layoutRefreshPending &&
            _pendingSourceLines is not null &&
            _pendingSourceLines.Count != _sourceLines.Count;

        if (string.Equals(currentText, _lastLayoutText, StringComparison.Ordinal) &&
            Math.Abs(currentWidth - _lastLayoutWidth) <= 0.5 &&
            Math.Abs(currentFontSize - _lastLayoutFontSize) <= 0.01 &&
            !_layoutRefreshPending &&
            !themeBrushesDirty)
        {
            return;
        }

        if (_layoutRefreshPending)
        {
            if (themeBrushesDirty ||
                Math.Abs(currentWidth - _lastLayoutWidth) > 0.5 ||
                Math.Abs(currentFontSize - _lastLayoutFontSize) > 0.01 ||
                requiresStructuralRefresh)
            {
                RebuildLayoutCache(
                    GetRenderableParsedScreenplay(currentText),
                    currentText,
                    currentWidth,
                    currentFontSize,
                    pendingSourceLines);
                return;
            }

            return;
        }

        RebuildLayoutCache(
            GetRenderableParsedScreenplay(currentText),
            currentText,
            currentWidth,
            currentFontSize,
            pendingSourceLines);
    }

    private PageFlowResult CalculatePageFlow(ParsedScreenplay screenplay, IReadOnlyList<SourceLineInfo> sourceLines)
    {
        if (sourceLines.Count == 0)
        {
            return PageFlowResult.Empty;
        }

        var lineTypes = BuildLineTypeMap(screenplay, sourceLines);
        var blocks = BuildLayoutBlocks(sourceLines, lineTypes);
        if (_useScreenplayLayout())
        {
            blocks = ApplySpacingRules(blocks);
        }
        var lineLayouts = new List<LineLayout>(sourceLines.Count);
        var lineLookup = new Dictionary<int, LineLayout>();
        var rowHeight = GetLineHeight();
        var bodyHeight = GetPageBodyHeight();
        var pages = new List<PageLayout>();
        var page = CreatePageLayout(0);
        var currentY = page.BodyRect.Top;
        var hasContentOnPage = false;

        pages.Add(page);

        for (var blockIndex = 0; blockIndex < blocks.Count; blockIndex++)
        {
            var block = blocks[blockIndex];

            if (block.IsSyntheticSpacer && !hasContentOnPage)
            {
                continue;
            }

            var blockHeight = EstimateBlockHeight(block);
            var keepTogetherHeight = block.ScreenplayType == ScreenplayElementType.Character
                ? Math.Max(blockHeight, EstimateCharacterKeepTogetherHeight(blocks, blockIndex))
                : blockHeight;

            if (hasContentOnPage && keepTogetherHeight > GetRemainingPageSpace(page, currentY))
            {
                page = CreatePageLayout(pages.Count);
                pages.Add(page);
                currentY = page.BodyRect.Top;
                hasContentOnPage = false;

                if (block.IsSyntheticSpacer)
                {
                    continue;
                }
            }

            if (blockHeight > bodyHeight)
            {
                PlaceOversizedBlock(block, lineLayouts, lineLookup, pages, ref page, ref currentY, ref hasContentOnPage);
                continue;
            }

            foreach (var sourceLine in block.SourceLines)
            {
                var lineLayout = CreateLineLayout(sourceLine, block.ScreenplayType, page, currentY);
                lineLayouts.Add(lineLayout);
                lineLookup[sourceLine.LineIndex] = lineLayout;
                currentY += lineLayout.Bounds.Height;
                hasContentOnPage = true;
            }
        }

        var documentHeight = pages.Count == 0
            ? GetMinimumDocumentHeight()
            : pages[^1].PageRect.Bottom;

        return new PageFlowResult(pages, lineLayouts, lineLookup, documentHeight);
    }

    private void PlaceOversizedBlock(
        LayoutBlock block,
        ICollection<LineLayout> lineLayouts,
        IDictionary<int, LineLayout> lineLookup,
        IList<PageLayout> pages,
        ref PageLayout page,
        ref double currentY,
        ref bool hasContentOnPage)
    {
        foreach (var sourceLine in block.SourceLines)
        {
            if (hasContentOnPage && GetRemainingPageSpace(page, currentY) < GetLineHeight())
            {
                page = CreatePageLayout(pages.Count);
                pages.Add(page);
                currentY = page.BodyRect.Top;
                hasContentOnPage = false;
            }

            var lineLayout = CreateLineLayout(sourceLine, block.ScreenplayType, page, currentY);
            if (hasContentOnPage && lineLayout.Bounds.Height > GetRemainingPageSpace(page, currentY))
            {
                page = CreatePageLayout(pages.Count);
                pages.Add(page);
                currentY = page.BodyRect.Top;
                hasContentOnPage = false;
                lineLayout = CreateLineLayout(sourceLine, block.ScreenplayType, page, currentY);
            }

            lineLayouts.Add(lineLayout);
            lineLookup[sourceLine.LineIndex] = lineLayout;
            currentY += lineLayout.Bounds.Height;
            hasContentOnPage = true;
        }
    }

    private IReadOnlyDictionary<int, ScreenplayElementType> BuildLineTypeMap(
        ParsedScreenplay screenplay,
        IReadOnlyList<SourceLineInfo> sourceLines)
    {
        var lineCount = sourceLines.Count;
        var lineTypes = Enumerable.Range(0, lineCount)
            .ToDictionary(index => index, _ => ScreenplayElementType.Action);

        if (!_useScreenplayLayout())
        {
            return lineTypes;
        }

        var parentheticalLines = screenplay.Elements
            .OfType<ParentheticalElement>()
            .Select(element => element.LineIndex)
            .ToHashSet();

        foreach (var element in screenplay.Elements)
        {
            if (element is DialogueElement dialogue)
            {
                for (var lineIndex = dialogue.LineIndex; lineIndex <= dialogue.EndLineIndex && lineIndex < lineCount; lineIndex++)
                {
                    if (lineIndex < 0 || parentheticalLines.Contains(lineIndex))
                    {
                        continue;
                    }

                    lineTypes[lineIndex] = ScreenplayElementType.Dialogue;
                }

                continue;
            }

            for (var lineIndex = element.LineIndex; lineIndex <= element.EndLineIndex && lineIndex < lineCount; lineIndex++)
            {
                if (lineIndex < 0)
                {
                    continue;
                }

                lineTypes[lineIndex] = element.Type;
            }
        }

        var explicitOverrideLineIndices = screenplay.LineTypeOverrides.Keys
            .Select(lineNumber => lineNumber - 1)
            .Where(lineIndex => lineIndex >= 0 && lineIndex < lineCount)
            .ToHashSet();

        foreach (var lineTypeOverride in screenplay.LineTypeOverrides)
        {
            var lineIndex = lineTypeOverride.Key - 1;
            if (lineIndex < 0 || lineIndex >= lineCount)
            {
                continue;
            }

            lineTypes[lineIndex] = lineTypeOverride.Value;
        }

        foreach (var sourceLine in sourceLines)
        {
            if (explicitOverrideLineIndices.Contains(sourceLine.LineIndex) ||
                !lineTypes.TryGetValue(sourceLine.LineIndex, out var mappedType) ||
                !SupportsLiveLineTypeRefresh(mappedType))
            {
                continue;
            }

            lineTypes[sourceLine.LineIndex] = ResolveLiveScreenplayLineType(sourceLine.Text);
        }

        return lineTypes;
    }

    private static bool SupportsLiveLineTypeRefresh(ScreenplayElementType screenplayType)
    {
        return screenplayType is ScreenplayElementType.Action
            or ScreenplayElementType.SceneHeading
            or ScreenplayElementType.Section
            or ScreenplayElementType.Synopsis
            or ScreenplayElementType.Character
            or ScreenplayElementType.Parenthetical;
    }

    private static ScreenplayElementType ResolveLiveScreenplayLineType(string lineText)
    {
        var trimmed = (lineText ?? string.Empty).Trim();
        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        if (TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()))
        {
            return ScreenplayElementType.SceneHeading;
        }

        if (trimmed.StartsWith('#'))
        {
            return ScreenplayElementType.Section;
        }

        if (trimmed.StartsWith('='))
        {
            return ScreenplayElementType.Synopsis;
        }

        if (trimmed.StartsWith("(") && trimmed.EndsWith(")", StringComparison.Ordinal))
        {
            return ScreenplayElementType.Parenthetical;
        }

        return TextAnalysis.IsLiveCharacterCueCandidate(lineText.AsSpan(), 45, 6)
            ? ScreenplayElementType.Character
            : ScreenplayElementType.Action;
    }

    private IReadOnlyList<LayoutBlock> BuildLayoutBlocks(
        IReadOnlyList<SourceLineInfo> sourceLines,
        IReadOnlyDictionary<int, ScreenplayElementType> lineTypes)
    {
        var blocks = new List<LayoutBlock>();
        var index = 0;

        while (index < sourceLines.Count)
        {
            var line = sourceLines[index];
            var isBlankLine = string.IsNullOrWhiteSpace(line.Text);
            var type = lineTypes.TryGetValue(index, out var mappedType)
                ? NormalizeRenderableType(mappedType)
                : ScreenplayElementType.Action;

            if (isBlankLine)
            {
                var spacerType = type;
                var spacerLines = new List<SourceLineInfo>();
                while (index < sourceLines.Count &&
                       string.IsNullOrWhiteSpace(sourceLines[index].Text) &&
                       NormalizeRenderableType(
                           lineTypes.TryGetValue(index, out var blankLineType)
                               ? blankLineType
                               : ScreenplayElementType.Action) == spacerType)
                {
                    spacerLines.Add(sourceLines[index]);
                    index++;
                }

                blocks.Add(new LayoutBlock(spacerType, spacerLines, true));
                continue;
            }

            var blockLines = new List<SourceLineInfo> { line };
            index++;

            if (type is ScreenplayElementType.Action or ScreenplayElementType.Dialogue)
            {
                while (index < sourceLines.Count &&
                       !string.IsNullOrWhiteSpace(sourceLines[index].Text) &&
                       lineTypes.TryGetValue(index, out var nextType) &&
                       NormalizeRenderableType(nextType) == type)
                {
                    blockLines.Add(sourceLines[index]);
                    index++;
                }
            }

            blocks.Add(new LayoutBlock(type, blockLines, false));
        }

        return blocks;
    }

    private IReadOnlyList<LayoutBlock> ApplySpacingRules(IReadOnlyList<LayoutBlock> blocks)
    {
        if (blocks.Count == 0)
        {
            return blocks;
        }

        // Augment the rendered page with spacer rows so Fountain spacing is visible
        // even when the source text does not already contain the required blanks.
        var adjustedBlocks = new List<LayoutBlock>(blocks.Count * 2);
        LayoutBlock? previousContentBlock = null;

        foreach (var block in blocks)
        {
            if (block.IsSpacer)
            {
                adjustedBlocks.Add(block);
                continue;
            }

            var currentSpacingType = ToSpacingType(block.ScreenplayType);

            if (previousContentBlock is null)
            {
                var requiredLeadingSpacing = ScreenplayFormatting.GetBlankLinesBefore(currentSpacingType);
                if (requiredLeadingSpacing > 0)
                {
                    if (adjustedBlocks.Count > 0 && adjustedBlocks[^1].IsSpacer)
                    {
                        var existingSpacer = adjustedBlocks[^1];
                        if (existingSpacer.SourceLines.Count < requiredLeadingSpacing)
                        {
                            adjustedBlocks[^1] = new LayoutBlock(
                                ScreenplayElementType.Action,
                                CombineSpacerLines(existingSpacer.SourceLines, requiredLeadingSpacing - existingSpacer.SourceLines.Count),
                                true);
                        }
                    }
                    else
                    {
                        adjustedBlocks.Add(new LayoutBlock(
                            ScreenplayElementType.Action,
                            CreateSpacerLines(requiredLeadingSpacing),
                            true));
                    }
                }
            }
            else
            {
                var requiredSpacing = ScreenplayFormatting.GetGapBlankLines(
                    ToSpacingType(previousContentBlock.ScreenplayType),
                    currentSpacingType);

                if (requiredSpacing > 0 && adjustedBlocks.Count > 0 && adjustedBlocks[^1].IsSpacer)
                {
                    var existingSpacer = adjustedBlocks[^1];
                    if (existingSpacer.SourceLines.Count < requiredSpacing)
                    {
                        adjustedBlocks[^1] = new LayoutBlock(
                            ScreenplayElementType.Action,
                            CombineSpacerLines(existingSpacer.SourceLines, requiredSpacing - existingSpacer.SourceLines.Count),
                            true);
                    }
                }
                else if (requiredSpacing > 0)
                {
                    adjustedBlocks.Add(new LayoutBlock(
                        ScreenplayElementType.Action,
                        CreateSpacerLines(requiredSpacing),
                        true));
                }
            }

            adjustedBlocks.Add(block);
            previousContentBlock = block;
        }

        if (previousContentBlock is not null)
        {
            var trailingSpacing = ScreenplayFormatting.GetBlankLinesAfter(ToSpacingType(previousContentBlock.ScreenplayType));
            if (trailingSpacing > 0 && adjustedBlocks.Count > 0 && adjustedBlocks[^1].IsSpacer)
            {
                var existingSpacer = adjustedBlocks[^1];
                if (existingSpacer.SourceLines.Count < trailingSpacing)
                {
                    adjustedBlocks[^1] = new LayoutBlock(
                        ScreenplayElementType.Action,
                        CombineSpacerLines(existingSpacer.SourceLines, trailingSpacing - existingSpacer.SourceLines.Count),
                        true);
                }
            }
            else if (trailingSpacing > 0)
            {
                adjustedBlocks.Add(new LayoutBlock(
                    ScreenplayElementType.Action,
                    CreateSpacerLines(trailingSpacing),
                    true));
            }
        }

        return adjustedBlocks;
    }

    private static IReadOnlyList<SourceLineInfo> CombineSpacerLines(IReadOnlyList<SourceLineInfo> existingLines, int additionalBlankLines)
    {
        if (additionalBlankLines <= 0)
        {
            return existingLines;
        }

        var combined = new List<SourceLineInfo>(existingLines.Count + additionalBlankLines);
        combined.AddRange(existingLines);
        combined.AddRange(CreateSpacerLines(additionalBlankLines));
        return combined;
    }

    private static IReadOnlyList<SourceLineInfo> CreateSpacerLines(int count)
    {
        var spacerLines = new List<SourceLineInfo>(Math.Max(0, count));

        for (var index = 0; index < count; index++)
        {
            spacerLines.Add(new SourceLineInfo(-1, 0, string.Empty));
        }

        return spacerLines;
    }

    private static ScreenplayFormatting.ElementSpacingType ToSpacingType(ScreenplayElementType screenplayType)
    {
        return screenplayType switch
        {
            ScreenplayElementType.SceneHeading => ScreenplayFormatting.ElementSpacingType.SceneHeading,
            ScreenplayElementType.Section => ScreenplayFormatting.ElementSpacingType.Section,
            ScreenplayElementType.Synopsis => ScreenplayFormatting.ElementSpacingType.Synopsis,
            ScreenplayElementType.Action => ScreenplayFormatting.ElementSpacingType.Action,
            ScreenplayElementType.Character => ScreenplayFormatting.ElementSpacingType.Character,
            ScreenplayElementType.Dialogue => ScreenplayFormatting.ElementSpacingType.Dialogue,
            ScreenplayElementType.Parenthetical => ScreenplayFormatting.ElementSpacingType.Parenthetical,
            ScreenplayElementType.Transition => ScreenplayFormatting.ElementSpacingType.Transition,
            _ => ScreenplayFormatting.ElementSpacingType.Other
        };
    }

    private double EstimateCharacterKeepTogetherHeight(IReadOnlyList<LayoutBlock> blocks, int characterBlockIndex)
    {
        var totalHeight = EstimateBlockHeight(blocks[characterBlockIndex]);
        var foundDialogue = false;

        for (var index = characterBlockIndex + 1; index < blocks.Count; index++)
        {
            var candidate = blocks[index];

            if (candidate.IsSpacer)
            {
                totalHeight += EstimateBlockHeight(candidate);
                continue;
            }

            if (candidate.ScreenplayType == ScreenplayElementType.Parenthetical)
            {
                totalHeight += EstimateBlockHeight(candidate);
                continue;
            }

            if (candidate.ScreenplayType == ScreenplayElementType.Dialogue)
            {
                totalHeight += EstimateBlockHeight(candidate);
                foundDialogue = true;
            }

            break;
        }

        return foundDialogue
            ? totalHeight
            : EstimateBlockHeight(blocks[characterBlockIndex]);
    }

    private double EstimateBlockHeight(LayoutBlock block)
    {
        if (block.SourceLines.Count == 0)
        {
            return GetLineHeight();
        }

        return block.SourceLines.Sum(sourceLine => EstimateLineHeight(sourceLine.Text, block.ScreenplayType));
    }

    private double EstimateLineHeight(string text, ScreenplayElementType screenplayType)
    {
        var visualText = GetVisualText(text, screenplayType);
        var wrappedLines = WrapVisualText(visualText, GetWrapLimit(screenplayType));
        return Math.Max(1, wrappedLines.Count) * GetLineHeight();
    }

    private LineLayout CreateLineLayout(
        SourceLineInfo sourceLine,
        ScreenplayElementType screenplayType,
        PageLayout page,
        double top)
    {
        var rowHeight = GetLineHeight();
        var visualText = GetVisualText(sourceLine.Text, screenplayType);
        var wrappedLines = WrapVisualText(visualText, GetWrapLimit(screenplayType));
        var lineBoldSpans = _useScreenplayLayout()
            ? Array.Empty<FormattingSpan>()
            : GetMarkdownBoldSpans(visualText);
        var rows = new List<WrappedRow>(wrappedLines.Count);

        for (var rowIndex = 0; rowIndex < wrappedLines.Count; rowIndex++)
        {
            var wrappedLine = wrappedLines[rowIndex];
            var formattedText = CreateFormattedText(
                wrappedLine.Text,
                screenplayType,
                GetWrappedLineBoldSpans(lineBoldSpans, wrappedLine.StartIndex, wrappedLine.Length));
            var rowWidth = formattedText.WidthIncludingTrailingWhitespace;
            var rowOrigin = new Point(
                GetRowOriginX(page.BodyRect, screenplayType, rowWidth),
                top + rowIndex * rowHeight);

            rows.Add(new WrappedRow(
                wrappedLine.StartIndex,
                wrappedLine.Length,
                wrappedLine.Text,
                rowOrigin,
                rowWidth,
                formattedText));
        }

        var bounds = new Rect(page.PageRect.Left, top, page.PageRect.Width, Math.Max(rowHeight, rows.Count * rowHeight));

        return new LineLayout(
            sourceLine.LineIndex,
            sourceLine.CharacterIndex,
            sourceLine.Text.Length,
            screenplayType,
            visualText,
            rows,
            page,
            bounds,
            rowHeight);
    }

    private IReadOnlyList<LineLayout> GetVisibleLineLayouts(Rect visibleBounds)
    {
        var visibleLines = _pageFlow.Lines
            .Where(layout => layout.Bottom >= visibleBounds.Top - 4.0 && layout.Top <= visibleBounds.Bottom + 4.0)
            .ToArray();

        if (!_layoutRefreshPending || _pendingSourceLines is null)
        {
            return visibleLines;
        }

        var transformed = new LineLayout[visibleLines.Length];
        for (var index = 0; index < visibleLines.Length; index++)
        {
            transformed[index] = CreateTransientLineLayout(visibleLines[index]);
        }

        return transformed;
    }

    private LineLayout? GetRenderableLineLayout(int lineIndex)
    {
        if (!_pageFlow.LineLookup.TryGetValue(lineIndex, out var lineLayout))
        {
            return null;
        }

        return _layoutRefreshPending && _pendingSourceLines is not null
            ? CreateTransientLineLayout(lineLayout)
            : lineLayout;
    }

    private LineLayout CreateTransientLineLayout(LineLayout baseLayout)
    {
        if (_pendingSourceLines is null ||
            baseLayout.SourceLineIndex < 0 ||
            baseLayout.SourceLineIndex >= _pendingSourceLines.Count)
        {
            return baseLayout;
        }

        var currentSourceLine = _pendingSourceLines[baseLayout.SourceLineIndex];
        var currentVisualText = GetVisualText(currentSourceLine.Text, baseLayout.ScreenplayType);
        if (string.Equals(currentVisualText, baseLayout.VisualText, StringComparison.Ordinal) &&
            currentSourceLine.CharacterIndex == baseLayout.CharacterStartIndex &&
            currentSourceLine.Text.Length == baseLayout.CharacterLength)
        {
            return baseLayout;
        }

        return CreateLineLayout(
            currentSourceLine,
            baseLayout.ScreenplayType,
            baseLayout.Page,
            baseLayout.Bounds.Top);
    }

    private LineLayout? GetClosestLineLayout(Point point)
    {
        if (_pageFlow.Lines.Count == 0)
        {
            return null;
        }

        var candidateIndices = GetClosestLineCandidateIndices(point.Y);
        LineLayout? closestLine = null;
        var closestDistance = double.PositiveInfinity;

        foreach (var index in candidateIndices)
        {
            var layout = _pageFlow.Lines[index];
            var candidate = _layoutRefreshPending && _pendingSourceLines is not null
                ? CreateTransientLineLayout(layout)
                : layout;

            var horizontalDistance = point.X < candidate.Page.BodyRect.Left
                ? candidate.Page.BodyRect.Left - point.X
                : point.X > candidate.Page.BodyRect.Right
                    ? point.X - candidate.Page.BodyRect.Right
                    : 0.0;

            var verticalDistance = point.Y < candidate.Top
                ? candidate.Top - point.Y
                : point.Y > candidate.Bottom
                    ? point.Y - candidate.Bottom
                    : 0.0;

            var weightedDistance = (verticalDistance * verticalDistance * 4.0) + (horizontalDistance * horizontalDistance);
            if (weightedDistance >= closestDistance)
            {
                continue;
            }

            closestDistance = weightedDistance;
            closestLine = candidate;

            if (weightedDistance <= 0.0)
            {
                break;
            }
        }

        return closestLine;
    }

    private IEnumerable<int> GetClosestLineCandidateIndices(double y)
    {
        if (_pageFlow.Lines.Count == 0)
        {
            yield break;
        }

        var nearestIndex = FindNearestLineIndexByY(y);
        const int candidateRadius = 12;
        var start = Math.Max(0, nearestIndex - candidateRadius);
        var end = Math.Min(_pageFlow.Lines.Count - 1, nearestIndex + candidateRadius);

        for (var index = start; index <= end; index++)
        {
            yield return index;
        }
    }

    private int FindNearestLineIndexByY(double y)
    {
        var low = 0;
        var high = _pageFlow.Lines.Count - 1;

        while (low <= high)
        {
            var mid = low + ((high - low) / 2);
            var layout = _pageFlow.Lines[mid];

            if (y < layout.Top)
            {
                high = mid - 1;
                continue;
            }

            if (y > layout.Bottom)
            {
                low = mid + 1;
                continue;
            }

            return mid;
        }

        if (low >= _pageFlow.Lines.Count)
        {
            return _pageFlow.Lines.Count - 1;
        }

        if (high < 0)
        {
            return 0;
        }

        var lowDistance = Math.Min(
            Math.Abs(_pageFlow.Lines[low].Top - y),
            Math.Abs(_pageFlow.Lines[low].Bottom - y));
        var highDistance = Math.Min(
            Math.Abs(_pageFlow.Lines[high].Top - y),
            Math.Abs(_pageFlow.Lines[high].Bottom - y));

        return lowDistance < highDistance
            ? low
            : high;
    }

    private int GetCharacterIndexFromPoint(LineLayout lineLayout, Point point)
    {
        if (lineLayout.Rows.Count == 0)
        {
            return lineLayout.CharacterStartIndex;
        }

        var row = GetClosestRow(lineLayout, point.Y);
        if (row is null)
        {
            return lineLayout.CharacterStartIndex;
        }

        var localX = point.X - row.Origin.X;
        if (localX <= 0.0 || row.Length == 0)
        {
            return lineLayout.CharacterStartIndex + row.StartIndex;
        }

        if (localX >= row.Width)
        {
            return lineLayout.CharacterStartIndex + row.StartIndex + row.Length;
        }

        return lineLayout.CharacterStartIndex + row.StartIndex + GetClosestCaretOffset(row.Text, localX, lineLayout.ScreenplayType);
    }

    private WrappedRow? GetClosestRow(LineLayout lineLayout, double y)
    {
        WrappedRow? closestRow = null;
        var closestDistance = double.PositiveInfinity;

        foreach (var row in lineLayout.Rows)
        {
            var rowTop = row.Origin.Y;
            var rowBottom = row.Origin.Y + lineLayout.RowHeight;
            var distance = y < rowTop
                ? rowTop - y
                : y > rowBottom
                    ? y - rowBottom
                    : 0.0;

            if (distance >= closestDistance)
            {
                continue;
            }

            closestDistance = distance;
            closestRow = row;

            if (distance <= 0.0)
            {
                break;
            }
        }

        return closestRow;
    }

    private int GetClosestCaretOffset(string text, double localX, ScreenplayElementType screenplayType)
    {
        if (string.IsNullOrEmpty(text) || localX <= 0.0)
        {
            return 0;
        }

        if (localX >= MeasureTextWidth(text, screenplayType))
        {
            return text.Length;
        }

        var low = 0;
        var high = text.Length;

        while (low < high)
        {
            var mid = (low + high) / 2;
            var width = MeasureTextWidth(text[..mid], screenplayType);
            if (width < localX)
            {
                low = mid + 1;
            }
            else
            {
                high = mid;
            }
        }

        var rightIndex = Math.Clamp(low, 0, text.Length);
        var leftIndex = Math.Max(0, rightIndex - 1);
        var leftWidth = MeasureTextWidth(text[..leftIndex], screenplayType);
        var rightWidth = MeasureTextWidth(text[..rightIndex], screenplayType);

        return Math.Abs(localX - leftWidth) <= Math.Abs(rightWidth - localX)
            ? leftIndex
            : rightIndex;
    }

    private Rect BuildCaretRect(LineLayout lineLayout, int caretIndex)
    {
        if (lineLayout.Rows.Count == 0)
        {
            return Rect.Empty;
        }

        var caretColumn = Math.Clamp(caretIndex - lineLayout.CharacterStartIndex, 0, lineLayout.CharacterLength);
        var caretRow = lineLayout.Rows[^1];

        foreach (var row in lineLayout.Rows)
        {
            var rowEnd = row.StartIndex + row.Length;
            if (caretColumn < rowEnd || ReferenceEquals(row, lineLayout.Rows[^1]))
            {
                caretRow = row;
                break;
            }
        }

        var prefixLength = Math.Clamp(caretColumn - caretRow.StartIndex, 0, caretRow.Length);
        var prefixWidth = MeasureTextWidth(caretRow.Text[..prefixLength], lineLayout.ScreenplayType);
        return new Rect(
            caretRow.Origin.X + prefixWidth,
            caretRow.Origin.Y + 1.0,
            1.4,
            Math.Max(8.0, lineLayout.RowHeight - 2.0));
    }

    private void DrawSelection(DrawingContext drawingContext, IReadOnlyList<LineLayout> lineLayouts)
    {
        if (GetSelectionLength() <= 0)
        {
            return;
        }

        var selectionStart = GetSelectionStart();
        var selectionEnd = selectionStart + GetSelectionLength();

        foreach (var layout in lineLayouts)
        {
            var lineStart = layout.CharacterStartIndex;
            var lineEnd = lineStart + layout.CharacterLength;
            var overlapStart = Math.Max(selectionStart, lineStart);
            var overlapEnd = Math.Min(selectionEnd, lineEnd);

            if (overlapEnd <= overlapStart)
            {
                continue;
            }

            foreach (var row in layout.Rows)
            {
                var rowStart = lineStart + row.StartIndex;
                var rowEnd = rowStart + row.Length;
                var rowOverlapStart = Math.Max(overlapStart, rowStart);
                var rowOverlapEnd = Math.Min(overlapEnd, rowEnd);

                if (rowOverlapEnd <= rowOverlapStart)
                {
                    continue;
                }

                var localStart = Math.Clamp(rowOverlapStart - rowStart, 0, row.Length);
                var localLength = Math.Clamp(rowOverlapEnd - rowOverlapStart, 0, row.Length - localStart);

                if (localLength <= 0)
                {
                    continue;
                }

                var prefixWidth = MeasureTextWidth(row.Text[..localStart], layout.ScreenplayType);
                var selectedWidth = Math.Max(2.0, MeasureTextWidth(row.Text.Substring(localStart, localLength), layout.ScreenplayType));
                var selectionRect = new Rect(
                    row.Origin.X + prefixWidth,
                    row.Origin.Y + 1.0,
                    selectedWidth,
                    Math.Max(4.0, layout.RowHeight - 2.0));

                drawingContext.DrawRoundedRectangle(_selectionBrush, null, selectionRect, 2.0, 2.0);
            }
        }
    }

    private void DrawCaret(DrawingContext drawingContext)
    {
        if (!_editor.IsKeyboardFocused || GetSelectionLength() > 0 || !_isCaretVisible)
        {
            return;
        }

        if (!TryGetCaretRect(out var caretRect))
        {
            return;
        }

        drawingContext.DrawRectangle(_caretBrush, null, caretRect);
    }

    private void DrawPageChrome(DrawingContext drawingContext)
    {
        UpdateLayoutCache();

        if (_pageFlow.Pages.Count == 0)
        {
            return;
        }

        var visibleBounds = GetVisibleBounds();
        DrawPageGaps(drawingContext, visibleBounds);

        foreach (var page in _pageFlow.Pages)
        {
            if (page.PageRect.Bottom < visibleBounds.Top - PageGap || page.PageRect.Top > visibleBounds.Bottom + PageGap)
            {
                continue;
            }

            var shadowRect = page.PageRect;
            shadowRect.Offset(PageShadowOffset, PageShadowOffset);

            if (_isFormattingEnabled)
            {
                drawingContext.DrawRectangle(_pageShadowBrush, null, shadowRect);
                drawingContext.DrawRectangle(_pageFillBrush, new Pen(_pageBorderBrush, 1.0), page.PageRect);
            }
            else
            {
                DrawMarkdownPageShadow(drawingContext, page.PageRect);
                drawingContext.DrawRectangle(Brushes.Transparent, new Pen(_pageBorderBrush, 1.0), page.PageRect);
            }
        }
    }

    private void DrawPageGaps(DrawingContext drawingContext, Rect visibleBounds)
    {
        if (_pageFlow.Pages.Count < 2)
        {
            return;
        }

        var gapBrush = ResolveBrush("SurfaceRaisedBackground", _pageFillBrush);

        for (var i = 0; i < _pageFlow.Pages.Count - 1; i++)
        {
            var currentPage = _pageFlow.Pages[i];
            var nextPage = _pageFlow.Pages[i + 1];
            var gapHeight = nextPage.PageRect.Top - currentPage.PageRect.Bottom;

            if (gapHeight <= 0.0)
            {
                continue;
            }

            var gapRect = new Rect(
                currentPage.PageRect.Left,
                currentPage.PageRect.Bottom,
                currentPage.PageRect.Width,
                gapHeight);

            if (gapRect.Bottom < visibleBounds.Top || gapRect.Top > visibleBounds.Bottom)
            {
                continue;
            }

            drawingContext.DrawRectangle(gapBrush, null, gapRect);
        }
    }

    private void DrawMarkdownPageShadow(DrawingContext drawingContext, Rect pageRect)
    {
        var shadowThickness = PageShadowOffset;
        var rightShadowRect = new Rect(
            pageRect.Right,
            pageRect.Top + shadowThickness,
            shadowThickness,
            Math.Max(0.0, pageRect.Height));
        var bottomShadowRect = new Rect(
            pageRect.Left + shadowThickness,
            pageRect.Bottom,
            Math.Max(0.0, pageRect.Width),
            shadowThickness);

        drawingContext.DrawRectangle(_pageShadowBrush, null, rightShadowRect);
        drawingContext.DrawRectangle(_pageShadowBrush, null, bottomShadowRect);
    }

    private void RefreshThemeBrushes()
    {
        var pageFillBrush = ResolveBrush("EditorPageBackground", FallbackPageFillBrush);
        var pageShadowBrush = ResolveBrush("EditorPageShadow", FallbackPageShadowBrush);
        var editorTextBrush = ResolveBrush("EditorForeground", FallbackEditorTextBrush);
        var selectionBrush = ResolveBrush("EditorSelection", FallbackSelectionBrush);
        var caretBrush = ResolveBrush("EditorCaret", FallbackCaretBrush);
        var pageBorderBrush = ResolveBrush("EditorPageBorder", FallbackPageBorderBrush);

        var themeBrushesChanged =
            !ReferenceEquals(_pageFillBrush, pageFillBrush) ||
            !ReferenceEquals(_pageShadowBrush, pageShadowBrush) ||
            !ReferenceEquals(_editorTextBrush, editorTextBrush) ||
            !ReferenceEquals(_selectionBrush, selectionBrush) ||
            !ReferenceEquals(_caretBrush, caretBrush) ||
            !ReferenceEquals(_pageBorderBrush, pageBorderBrush);

        _pageFillBrush = pageFillBrush;
        _pageShadowBrush = pageShadowBrush;
        _editorTextBrush = editorTextBrush;
        _selectionBrush = selectionBrush;
        _caretBrush = caretBrush;
        _pageBorderBrush = pageBorderBrush;

        if (themeBrushesChanged)
        {
            _themeBrushesDirty = true;
        }
    }

    private Brush ResolveBrush(string resourceKey, Brush fallback)
    {
        if (Application.Current?.TryFindResource(resourceKey) is Brush applicationBrush)
        {
            return applicationBrush;
        }

        return TryFindResource(resourceKey) as Brush ?? fallback;
    }

    private Rect GetVisibleBounds()
    {
        if (!_editor.IsLoaded || _scrollHost.ViewportHeight <= 0 || _scrollHost.ViewportWidth <= 0)
        {
            return new Rect(new Point(0, 0), _editor.RenderSize);
        }

        var zoomScale = Math.Max(0.01, _getZoomScale());

        Point origin;
        try
        {
            origin = _editor.TranslatePoint(new Point(0, 0), _scrollHost);
        }
        catch (InvalidOperationException)
        {
            return new Rect(new Point(0, 0), _editor.RenderSize);
        }

        return new Rect(
            Math.Max(0.0, -origin.X / zoomScale),
            Math.Max(0.0, -origin.Y / zoomScale),
            Math.Max(0.0, _scrollHost.ViewportWidth / zoomScale),
            Math.Max(0.0, _scrollHost.ViewportHeight / zoomScale));
    }

    private PageLayout CreatePageLayout(int pageNumber)
    {
        var pageTop = pageNumber * (ScreenplayPageHeight + PageGap);
        var pageLeft = GetPageLeft();
        var pageRect = new Rect(pageLeft, pageTop, ScreenplayPageWidth, ScreenplayPageHeight);
        var bodyRect = new Rect(
            pageRect.Left + ScreenplayPageLeftMargin,
            pageRect.Top + ScreenplayPageTopMargin,
            ScreenplayPageWidth - ScreenplayPageLeftMargin - ScreenplayPageRightMargin,
            ScreenplayPageHeight - ScreenplayPageTopMargin - ScreenplayPageBottomMargin);

        return new PageLayout(pageNumber + 1, pageRect, bodyRect);
    }

    private double GetPageLeft()
    {
        return Math.Max(0.0, (_editor.ActualWidth - ScreenplayPageWidth) / 2.0);
    }

    private double GetPageBodyHeight()
    {
        return ScreenplayPageHeight - ScreenplayPageTopMargin - ScreenplayPageBottomMargin;
    }

    private double GetRemainingPageSpace(PageLayout page, double currentY)
    {
        return Math.Max(0.0, page.BodyRect.Bottom - currentY);
    }

    private double GetMinimumDocumentHeight()
    {
        return ScreenplayPageHeight + PageShadowOffset;
    }

    private double GetLineHeight()
    {
        return Math.Max(1.0, _editor.FontSize * 1.08);
    }

    private double GetRowOriginX(Rect bodyRect, ScreenplayElementType screenplayType, double rowWidth)
    {
        return screenplayType switch
        {
            ScreenplayElementType.Dialogue => bodyRect.Left + DialogueIndent,
            ScreenplayElementType.Character => bodyRect.Left + CharacterIndent,
            ScreenplayElementType.Parenthetical => bodyRect.Left + ParentheticalIndent,
            ScreenplayElementType.Transition => Math.Max(bodyRect.Left, bodyRect.Right - rowWidth),
            ScreenplayElementType.CenteredText => bodyRect.Left + Math.Max(0.0, (bodyRect.Width - rowWidth) / 2.0),
            _ => bodyRect.Left
        };
    }

    private int GetWrapLimit(ScreenplayElementType screenplayType)
    {
        var configuredLimit = screenplayType switch
        {
            ScreenplayElementType.Character => ScreenplayFormatting.CharacterWrapChars,
            ScreenplayElementType.Dialogue => ScreenplayFormatting.DialogueWrapChars,
            ScreenplayElementType.Parenthetical => ScreenplayFormatting.ParentheticalWrapChars,
            ScreenplayElementType.Transition => ScreenplayFormatting.TransitionWrapChars,
            _ => ScreenplayFormatting.ActionWrapChars
        };

        var sampleWidth = MeasureTextWidth("W", screenplayType);
        if (sampleWidth <= 0.0)
        {
            return configuredLimit;
        }

        var fittedLimit = (int)Math.Floor(GetWrapWidth(screenplayType) / sampleWidth);
        return Math.Max(1, Math.Min(configuredLimit, fittedLimit));
    }

    private double GetWrapWidth(ScreenplayElementType screenplayType)
    {
        var bodyWidth = ScreenplayPageWidth - ScreenplayPageLeftMargin - ScreenplayPageRightMargin;
        return screenplayType switch
        {
            ScreenplayElementType.Dialogue => Math.Max(0.0, bodyWidth - DialogueIndent),
            ScreenplayElementType.Character => Math.Max(0.0, bodyWidth - CharacterIndent),
            ScreenplayElementType.Parenthetical => Math.Max(0.0, bodyWidth - ParentheticalIndent),
            ScreenplayElementType.Transition => Math.Max(0.0, bodyWidth),
            _ => Math.Max(0.0, bodyWidth)
        };
    }

    private static ScreenplayElementType NormalizeRenderableType(ScreenplayElementType screenplayType)
    {
        return screenplayType switch
        {
            ScreenplayElementType.SceneHeading => ScreenplayElementType.SceneHeading,
            ScreenplayElementType.Section => ScreenplayElementType.Section,
            ScreenplayElementType.Synopsis => ScreenplayElementType.Synopsis,
            ScreenplayElementType.Action => ScreenplayElementType.Action,
            ScreenplayElementType.Character => ScreenplayElementType.Character,
            ScreenplayElementType.Dialogue => ScreenplayElementType.Dialogue,
            ScreenplayElementType.Parenthetical => ScreenplayElementType.Parenthetical,
            ScreenplayElementType.Transition => ScreenplayElementType.Transition,
            ScreenplayElementType.CenteredText => ScreenplayElementType.CenteredText,
            _ => ScreenplayElementType.Action
        };
    }

    private static IReadOnlyList<SourceLineInfo> BuildSourceLines(string text)
    {
        var lines = new List<SourceLineInfo>();

        if (text.Length == 0)
        {
            lines.Add(new SourceLineInfo(0, 0, string.Empty));
            return lines;
        }

        var lineStart = 0;
        var lineIndex = 0;

        for (var index = 0; index <= text.Length; index++)
        {
            var atEnd = index == text.Length;
            if (!atEnd && text[index] is not ('\r' or '\n'))
            {
                continue;
            }

            var lineLength = index - lineStart;
            var lineText = lineLength <= 0 ? string.Empty : text.Substring(lineStart, lineLength);
            lines.Add(new SourceLineInfo(lineIndex, lineStart, NormalizeLineText(lineText)));
            lineIndex++;

            if (atEnd)
            {
                break;
            }

            if (text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
            {
                index++;
            }

            lineStart = index + 1;
        }

        return lines;
    }

    private static IReadOnlyList<WrappedLine> WrapVisualText(string text, int maxChars)
    {
        var wrappedLines = new List<WrappedLine>();
        var normalizedText = text ?? string.Empty;

        if (normalizedText.Length == 0)
        {
            wrappedLines.Add(new WrappedLine(0, 0, string.Empty));
            return wrappedLines;
        }

        var startIndex = 0;
        var limit = Math.Max(1, maxChars);

        while (startIndex < normalizedText.Length)
        {
            var remaining = normalizedText.Length - startIndex;
            if (remaining <= limit)
            {
                wrappedLines.Add(new WrappedLine(startIndex, remaining, normalizedText.Substring(startIndex, remaining)));
                break;
            }

            var breakIndex = FindWrapBreakIndex(normalizedText, startIndex, limit);
            if (breakIndex <= startIndex)
            {
                breakIndex = Math.Min(normalizedText.Length, startIndex + limit);
            }

            var length = breakIndex - startIndex;
            wrappedLines.Add(new WrappedLine(startIndex, length, normalizedText.Substring(startIndex, length)));
            startIndex = breakIndex;
        }

        return wrappedLines;
    }

    private static int FindWrapBreakIndex(string text, int startIndex, int limit)
    {
        var searchEnd = Math.Min(text.Length, startIndex + limit);

        for (var index = searchEnd - 1; index > startIndex; index--)
        {
            if (char.IsWhiteSpace(text[index]))
            {
                return index + 1;
            }
        }

        return searchEnd;
    }

    private FormattedText CreateFormattedText(
        string text,
        ScreenplayElementType screenplayType,
        IReadOnlyList<FormattingSpan>? boldSpans = null)
    {
        var fontWeight = screenplayType == ScreenplayElementType.SceneHeading
            ? FontWeights.Bold
            : FontWeights.Normal;

        var formattedText = new FormattedText(
            text ?? string.Empty,
            CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            new Typeface(_editor.FontFamily, FontStyles.Normal, fontWeight, FontStretches.Normal),
            _editor.FontSize,
            _editorTextBrush,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);

        if (boldSpans is not null)
        {
            foreach (var boldSpan in boldSpans)
            {
                var start = Math.Clamp(boldSpan.Start, 0, formattedText.Text.Length);
                var length = Math.Clamp(boldSpan.Length, 0, formattedText.Text.Length - start);
                if (length > 0)
                {
                    formattedText.SetFontWeight(FontWeights.Bold, start, length);
                }
            }
        }

        return formattedText;
    }

    private double MeasureTextWidth(string text, ScreenplayElementType screenplayType)
    {
        if (string.IsNullOrEmpty(text))
        {
            return 0.0;
        }

        return CreateFormattedText(text, screenplayType).WidthIncludingTrailingWhitespace;
    }

    private static string NormalizeLineText(string text)
    {
        return text
            .Replace("\r", string.Empty, StringComparison.Ordinal)
            .Replace("\n", string.Empty, StringComparison.Ordinal);
    }

    private static string GetVisualText(string text, ScreenplayElementType screenplayType)
    {
        var normalizedText = NormalizeLineText(text);
        // Strip [[id:GUID]] tags so they contribute zero width/height to the visual layout.
        // The tag remains in the backing string but is never rendered or measured.
        var strippedText = IdTagStripRegex.Replace(normalizedText, string.Empty);
        return screenplayType switch
        {
            ScreenplayElementType.SceneHeading => strippedText.ToUpperInvariant(),
            ScreenplayElementType.Transition   => strippedText.ToUpperInvariant(),
            ScreenplayElementType.Character    => strippedText.ToUpperInvariant().TrimEnd('^'),
            _                                  => strippedText
        };
    }

    private static IReadOnlyList<FormattingSpan> GetMarkdownBoldSpans(string lineText)
    {
        var boldSpans = new List<FormattingSpan>();
        foreach (Match match in MarkdownBoldRegex.Matches(lineText ?? string.Empty))
        {
            if (!match.Success || match.Groups.Count < 2)
            {
                continue;
            }

            var boldGroup = match.Groups[1];
            if (boldGroup.Length <= 0)
            {
                continue;
            }

            boldSpans.Add(new FormattingSpan(boldGroup.Index, boldGroup.Length));
        }

        return boldSpans;
    }

    private static IReadOnlyList<FormattingSpan> GetWrappedLineBoldSpans(
        IReadOnlyList<FormattingSpan> lineBoldSpans,
        int lineStartIndex,
        int lineLength)
    {
        if (lineBoldSpans.Count == 0 || lineLength <= 0)
        {
            return Array.Empty<FormattingSpan>();
        }

        var lineEndIndex = lineStartIndex + lineLength;
        var wrappedLineBoldSpans = new List<FormattingSpan>();
        foreach (var boldSpan in lineBoldSpans)
        {
            var boldStart = Math.Max(lineStartIndex, boldSpan.Start);
            var boldEnd = Math.Min(lineEndIndex, boldSpan.Start + boldSpan.Length);
            if (boldEnd <= boldStart)
            {
                continue;
            }

            wrappedLineBoldSpans.Add(new FormattingSpan(
                boldStart - lineStartIndex,
                boldEnd - boldStart));
        }

        return wrappedLineBoldSpans;
    }

    private void ScheduleLayoutRefresh(IReadOnlyList<SourceLineInfo> sourceLines, bool immediate)
    {
        _pendingSourceLines = sourceLines;
        _layoutRefreshPending = true;
        _layoutRefreshTimer.Stop();

        if (immediate)
        {
            ApplyPendingLayoutRefresh();
            return;
        }

        _layoutRefreshTimer.Interval = TimeSpan.FromMilliseconds(
            _editor.IsKeyboardFocusWithin
                ? InteractiveLayoutRefreshDelayMilliseconds
                : DeferredLayoutRefreshDelayMilliseconds);
        _layoutRefreshTimer.Start();
    }

    private void ScheduleLayoutRefresh(bool immediate)
    {
        _pendingSourceLines = null;
        _layoutRefreshPending = true;
        _layoutRefreshTimer.Stop();

        if (immediate)
        {
            ApplyPendingLayoutRefresh();
            return;
        }

        _layoutRefreshTimer.Interval = TimeSpan.FromMilliseconds(
            _editor.IsKeyboardFocusWithin
                ? InteractiveLayoutRefreshDelayMilliseconds
                : DeferredLayoutRefreshDelayMilliseconds);
        _layoutRefreshTimer.Start();
    }

    private void ApplyPendingLayoutRefresh()
    {
        _layoutRefreshTimer.Stop();
        if (!_layoutRefreshPending)
        {
            return;
        }

        var currentText = _currentText;
        RebuildLayoutCache(
            GetRenderableParsedScreenplay(currentText),
            currentText,
            _editor.ActualWidth,
            _editor.FontSize,
            _pendingSourceLines ?? _sourceLines);

        LayoutChanged?.Invoke(this, EventArgs.Empty);
        ScheduleVisualRefresh();
    }

    private void RebuildLayoutCache(
        ParsedScreenplay parsed,
        string currentText,
        double currentWidth,
        double currentFontSize,
        IReadOnlyList<SourceLineInfo> sourceLines)
    {
        _sourceLines = sourceLines;
        _pendingSourceLines = null;
        _pageFlow = CalculatePageFlow(parsed, sourceLines);
        _lastLayoutText = currentText;
        _lastLayoutWidth = currentWidth;
        _lastLayoutFontSize = currentFontSize;
        _lastLayoutParsed = parsed;
        _layoutRefreshPending = false;
        _themeBrushesDirty = false;
    }

    private ParsedScreenplay GetRenderableParsedScreenplay(string currentText)
    {
        var parsed = _getParsedScreenplay();
        if (string.Equals(parsed.RawText, currentText, StringComparison.Ordinal))
        {
            return parsed;
        }

        return _interactiveParser.Parse(currentText, _getLineTypeOverrides());
    }

    private void ScheduleVisualRefresh()
    {
        if (_visualRefreshPending)
        {
            return;
        }

        _visualRefreshPending = true;
        Dispatcher.BeginInvoke(
            new Action(() =>
            {
                _visualRefreshPending = false;
                InvalidateVisual();
            }),
            DispatcherPriority.Render);
    }

    private static T Freeze<T>(T freezable)
        where T : Freezable
    {
        freezable.Freeze();
        return freezable;
    }
}
