using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using AvaloniaEdit;
using Passage.App.ViewModels;
using Passage.Parser;
using Passage.App.Services;

namespace Passage.App.Views;

public partial class MainWindow : Window
{
    public static OutlineNodeViewModel? DraggedOutlineNode { get; set; }
    public static BeatBoardCardViewModel? DraggedBeatBoardCard { get; set; }

    private static readonly System.Text.RegularExpressions.Regex TransitionRegex = new(
        @"^(?:FADE IN|FADE OUT|CUT TO|DISSOLVE TO|SMASH CUT TO|MATCH CUT TO|WIPE TO|JUMP CUT TO|FADE TO BLACK)(?:[:.])?$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private bool _isLeftDockCollapsed = false;
    private double _leftDockExpandedWidth = 360;
    // Guards against feedback loops while we mirror text between the editor control
    // and the view model's EditorContent string.
    private bool _suppressEditorSync;

    // Editor transformers swapped by write mode: Fountain colouring + screenplay
    // indentation in Screenplay mode, Markdown colouring in Markdown mode.
    private FountainSyntaxColorizer? _fountainColorizer;
    private FountainIndentationGenerator? _fountainIndentation;
    private MarkdownSyntaxColorizer? _markdownColorizer;
    private ScreenplayPageRuler? _pageRuler;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel(this);

        // Apply the configured workspace panel width at startup so that a single
        // field (_leftDockExpandedWidth) is the source of truth for BOTH the
        // initial width and the expanded width after a collapse/expand toggle.
        // Without this, startup uses the XAML ColumnDefinition width instead, and
        // changing the field alone has no visible effect on launch.
        var mainGrid = this.FindControl<Grid>("MainLayoutGrid");
        if (mainGrid != null && mainGrid.ColumnDefinitions.Count > 0)
        {
            mainGrid.ColumnDefinitions[0].Width = new GridLength(_leftDockExpandedWidth);
        }

        // Add keyboard shortcuts
        AddKeyboardShortcuts();

        // Zoom gestures. Ctrl+scroll is the standard Linux zoom gesture (it is
        // also what two-finger trackpad scrolling with Ctrl held produces), and
        // the pinch recognizer covers touchscreens/Wayland where real pinch
        // events are delivered. Tunneling so the editor's own scroll handling
        // can't swallow the Ctrl+wheel first.
        AddHandler(PointerWheelChangedEvent, OnGlobalPointerWheelChanged, RoutingStrategies.Tunnel);
        var mainGrid2 = this.FindControl<Grid>("MainLayoutGrid");
        if (mainGrid2 != null)
        {
            mainGrid2.GestureRecognizers.Add(new PinchGestureRecognizer());
            mainGrid2.AddHandler(InputElement.PinchEvent, OnPinch);
            mainGrid2.AddHandler(InputElement.PinchEndedEvent, OnPinchEnded);
        }

        var editorBox = this.FindControl<TextEditor>("EditorTextBox");
        if (editorBox != null)
        {
            // Add tunneling KeyDown handler for autocomplete
            editorBox.AddHandler(InputElement.KeyDownEvent, EditorBox_KeyDown, RoutingStrategies.Tunnel);

            // Live Fountain formatting — the editor re-colours each line by its
            // screenplay element type (colorizer) and shifts the line's text to its
            // proper screenplay position (indentation generator) as the user types.
            // Both share one classifier so the document is only parsed once per edit.
            var lineClassifier = new FountainLineClassifier();
            _fountainColorizer = new FountainSyntaxColorizer(lineClassifier);
            _fountainIndentation = new FountainIndentationGenerator(lineClassifier);
            _markdownColorizer = new MarkdownSyntaxColorizer();
            _pageRuler = new ScreenplayPageRuler();

            // Drive the editor content through code-behind (instead of a XAML binding)
            // so every existing EditorContent code path in the view model keeps working.
            if (DataContext is MainWindowViewModel vm)
            {
                _suppressEditorSync = true;
                editorBox.Text = vm.EditorContent ?? string.Empty;
                _suppressEditorSync = false;

                ApplyEditorWriteMode(editorBox, vm.CurrentMode);

                vm.PropertyChanged += (_, e) =>
                {
                    if (e.PropertyName == nameof(MainWindowViewModel.CurrentMode))
                    {
                        ApplyEditorWriteMode(editorBox, vm.CurrentMode);
                        return;
                    }

                    if (e.PropertyName != nameof(MainWindowViewModel.EditorContent))
                    {
                        return;
                    }

                    var desired = vm.EditorContent ?? string.Empty;
                    if (_suppressEditorSync || editorBox.Text == desired)
                    {
                        return;
                    }

                    _suppressEditorSync = true;
                    var caret = editorBox.CaretOffset;
                    // Replace through the document (not the Text property, which
                    // resets the undo stack) so VM-driven edits — card saves,
                    // drag-drop moves, title-page changes — stay undoable.
                    editorBox.Document.Replace(0, editorBox.Document.TextLength, desired);
                    editorBox.CaretOffset = Math.Min(caret, editorBox.Document.TextLength);
                    _suppressEditorSync = false;
                };
            }

            editorBox.TextChanged += EditorBox_TextChanged;

            // Listen to caret changes for the status bar.
            editorBox.TextArea.Caret.PositionChanged += (_, _) =>
            {
                if (DataContext is MainWindowViewModel vm2)
                {
                    vm2.UpdateCaretStatus(editorBox.CaretOffset);
                }
            };
        }
    }

    private void AddKeyboardShortcuts()
    {
        if (DataContext is not MainWindowViewModel vm) return;

        // File Menu Shortcuts
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.N, KeyModifiers.Control), Command = vm.NewCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.O, KeyModifiers.Control), Command = vm.OpenCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.S, KeyModifiers.Control), Command = vm.SaveCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.S, KeyModifiers.Control | KeyModifiers.Shift), Command = vm.SaveAsCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.W, KeyModifiers.Control), Command = vm.CloseCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.Q, KeyModifiers.Control), Command = vm.ExitCommand });

        // Edit Menu Shortcuts
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.Z, KeyModifiers.Control), Command = vm.UndoCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.Y, KeyModifiers.Control), Command = vm.RedoCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.F, KeyModifiers.Control), Command = vm.FindCommand });

        // View Menu Shortcuts
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.OemPlus, KeyModifiers.Control), Command = vm.ZoomInCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.OemMinus, KeyModifiers.Control), Command = vm.ZoomOutCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.D0, KeyModifiers.Control), Command = vm.ResetZoomCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.F1), Command = vm.ToggleSyntaxPanelCommand });

        // Navigate Menu Shortcuts
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.G, KeyModifiers.Control), Command = vm.GoToLineCommand });
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.G, KeyModifiers.Control | KeyModifiers.Shift), Command = vm.GoToSceneCommand });

        // Mode toggle shortcut
        KeyBindings.Add(new KeyBinding { Gesture = new KeyGesture(Key.M, KeyModifiers.Control), Command = vm.ToggleWriteModeCommand });
    }

    // Swaps the editor's live-formatting layer to match the write mode: Fountain
    // colouring + screenplay indentation for screenplays, Markdown colouring (and a
    // symmetric page margin, since there is no 1.5" screenplay gutter) for Markdown.
    private void ApplyEditorWriteMode(TextEditor editorBox, WriteMode mode)
    {
        if (_fountainColorizer == null || _fountainIndentation == null || _markdownColorizer == null || _pageRuler == null)
        {
            return;
        }

        var textView = editorBox.TextArea.TextView;
        textView.LineTransformers.Remove(_fountainColorizer);
        textView.LineTransformers.Remove(_markdownColorizer);
        textView.ElementGenerators.Remove(_fountainIndentation);
        textView.BackgroundRenderers.Remove(_pageRuler);

        if (mode == WriteMode.Screenplay)
        {
            textView.LineTransformers.Add(_fountainColorizer);
            textView.ElementGenerators.Add(_fountainIndentation);
            textView.BackgroundRenderers.Add(_pageRuler);
            editorBox.Padding = new Thickness(144, 96, 96, 96);
        }
        else
        {
            textView.LineTransformers.Add(_markdownColorizer);
            editorBox.Padding = new Thickness(96);
        }

        textView.Redraw();
    }

    private const double MinZoom = 0.5;
    private const double MaxZoom = 2.0;
    private double? _pinchStartZoom;

    private void OnGlobalPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            return;
        }

        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        // Delta is ±1 per notch for wheels and fractional for trackpad smooth
        // scrolling, so scaling by the delta keeps both feeling proportional.
        var newZoom = vm.EditorZoomScale + e.Delta.Y * 0.1;
        vm.EditorZoomScale = Math.Clamp(newZoom, MinZoom, MaxZoom);
        e.Handled = true;
    }

    private void OnPinch(object? sender, PinchEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        _pinchStartZoom ??= vm.EditorZoomScale;
        vm.EditorZoomScale = Math.Clamp(_pinchStartZoom.Value * e.Scale, MinZoom, MaxZoom);
        e.Handled = true;
    }

    private void OnPinchEnded(object? sender, PinchEndedEventArgs e)
    {
        _pinchStartZoom = null;
    }

    public void UndoEditor()
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Undo();
    }

    public void RedoEditor()
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Redo();
    }

    // Clears the editor's undo history. Called when a different document is
    // loaded so undo can't walk back into the previous document's text.
    public void ResetEditorUndoHistory()
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Document.UndoStack.ClearAll();
    }

    // Repaints the editor so theme-resolved syntax brushes (SyntaxTheme) pick up
    // a light/dark switch immediately.
    public void RedrawEditor()
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.TextArea.TextView.Redraw();
    }

    private void ToggleLeftDock_Click(object? sender, RoutedEventArgs e)
    {
        SetLeftDockCollapsed(!_isLeftDockCollapsed);
    }

    private void SetLeftDockCollapsed(bool collapsed)
    {
        _isLeftDockCollapsed = collapsed;

        var grid = this.FindControl<Grid>("MainLayoutGrid");
        var leftSplitter = this.FindControl<GridSplitter>("LeftSplitter");
        var leftDockTabs = this.FindControl<TabControl>("LeftDockTabs");
        var leftDockTitleText = this.FindControl<Border>("LeftDockTitleText");
        var leftDockToggleButton = this.FindControl<Button>("LeftDockToggleButton");

        if (grid != null && grid.ColumnDefinitions.Count >= 3)
        {
            var leftDockColumn = grid.ColumnDefinitions[0];
            var leftSplitterColumn = grid.ColumnDefinitions[1];

            // Save current width before collapsing if it is not collapsed
            if (collapsed)
            {
                _leftDockExpandedWidth = leftDockColumn.Width.Value > 80 ? leftDockColumn.Width.Value : 350;
            }

            leftDockColumn.Width = collapsed ? new GridLength(65) : new GridLength(_leftDockExpandedWidth);
            leftSplitterColumn.Width = collapsed ? new GridLength(0) : new GridLength(2);
        }

        if (leftSplitter != null)
        {
            leftSplitter.IsVisible = !collapsed;
        }

        if (leftDockTabs != null)
        {
            leftDockTabs.IsVisible = !collapsed;
        }

        if (leftDockTitleText != null)
        {
            leftDockTitleText.IsVisible = !collapsed;
        }

        if (leftDockToggleButton != null)
        {
            leftDockToggleButton.Content = collapsed ? ">" : "<";
            ToolTip.SetTip(leftDockToggleButton, collapsed ? "Expand left dock" : "Collapse left dock");
        }
    }

    private bool _isSyntaxPanelVisible = false;
    private double _syntaxPanelExpandedWidth = 280;

    public void ToggleSyntaxPanel()
    {
        SetSyntaxPanelVisible(!_isSyntaxPanelVisible);
    }

    private void SetSyntaxPanelVisible(bool visible)
    {
        _isSyntaxPanelVisible = visible;

        var grid = this.FindControl<Grid>("MainLayoutGrid");
        var rightSplitter = this.FindControl<GridSplitter>("RightSplitter");
        var rightDockBorder = this.FindControl<Border>("RightDockBorder");

        if (grid != null && grid.ColumnDefinitions.Count >= 5)
        {
            var splitterColumn = grid.ColumnDefinitions[3];
            var panelColumn = grid.ColumnDefinitions[4];

            if (visible)
            {
                splitterColumn.Width = GridLength.Auto;
                panelColumn.Width = new GridLength(_syntaxPanelExpandedWidth);
                panelColumn.MinWidth = 200;
            }
            else
            {
                // Save current width before hiding
                if (panelColumn.Width.Value > 0)
                {
                    _syntaxPanelExpandedWidth = panelColumn.Width.Value;
                }
                splitterColumn.Width = new GridLength(0);
                panelColumn.Width = new GridLength(0);
                panelColumn.MinWidth = 0;
            }
        }

        if (rightSplitter != null)
        {
            rightSplitter.IsVisible = visible;
        }

        if (rightDockBorder != null)
        {
            rightDockBorder.IsVisible = visible;
        }
    }


    private void OutlineNode_DoubleTapped(object? sender, Avalonia.Input.TappedEventArgs e)
    {
        if (sender is StackPanel panel && panel.DataContext is OutlineNodeViewModel node)
        {
            NavigateToLine(node.LineNumber);
        }
    }

    private void ScratchpadItem_DoubleTapped(object? sender, Avalonia.Input.TappedEventArgs e)
    {
        if (sender is Border border && border.DataContext is ScreenplayElement element)
        {
            NavigateToLine(element.LineNumber);
        }
    }

    private void BeatBoardCard_DoubleTapped(object? sender, Avalonia.Input.TappedEventArgs e)
    {
        if (sender is Border border && border.DataContext is BeatBoardCardViewModel card)
        {
            if (DataContext is MainWindowViewModel vm)
            {
                vm.IsBoardModeActive = false;
                NavigateToLine(card.LineNumber);
            }
        }
    }

    private void OnCutClick(object? sender, RoutedEventArgs e)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Cut();
    }

    private void OnCopyClick(object? sender, RoutedEventArgs e)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Copy();
    }

    private void OnPasteClick(object? sender, RoutedEventArgs e)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        editor?.Paste();
    }

    public void NavigateToLine(int lineNumber)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        if (editor == null || string.IsNullOrEmpty(editor.Text)) return;

        var text = editor.Text;
        var lineIndex = Math.Max(0, lineNumber - 1);

        var currentIndex = 0;
        var currentLine = 0;

        while (currentLine < lineIndex && currentIndex < text.Length)
        {
            var nextNewline = text.IndexOf('\n', currentIndex);
            if (nextNewline == -1) break;
            currentIndex = nextNewline + 1;
            currentLine++;
        }

        editor.TextArea.ClearSelection();
        editor.CaretOffset = currentIndex;
        editor.ScrollToLine(Math.Max(1, lineNumber));

        // Move focus to the editor's TextArea so the caret actually shows on the
        // page and typing continues from here. Deferred to Input priority so it
        // runs after the originating pointer event finishes — otherwise the
        // clicked card grabs focus back and the caret never appears.
        Dispatcher.UIThread.Post(() =>
        {
            editor.TextArea.Focus();
            editor.CaretOffset = currentIndex;
        }, DispatcherPriority.Input);
    }

    private void EditorBox_TextChanged(object? sender, EventArgs e)
    {
        var textBox = sender as TextEditor;
        if (textBox == null || DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        // Programmatic content sync (e.g. opening a file) — don't run live logic.
        if (_suppressEditorSync)
        {
            return;
        }

        // Mirror the edit back into the view model so parsing, preview, word counts
        // and saving all stay in sync with what the user typed.
        vm.EditorContent = textBox.Text ?? string.Empty;

        // A single edit can change how *other* lines are classified (typing a
        // character cue re-colours the dialogue block beneath it, etc.). AvaloniaEdit
        // only repaints the edited line by default, so force a full re-colour of the
        // visible text to match the Windows editor's whole-document reformatting.
        textBox.TextArea.TextView.Redraw();

        if (!vm.IsScreenplayMode)
        {
            vm.IsAutoCompleteOpen = false;
            return;
        }

        var caretIndex = textBox.CaretOffset;
        var text = textBox.Text ?? "";

        // Find line start and line text
        var lineStart = 0;
        var lineIndex = 0;
        var currentIndex = 0;
        while (currentIndex < caretIndex && currentIndex < text.Length)
        {
            var nextNewline = text.IndexOf('\n', currentIndex);
            if (nextNewline == -1 || nextNewline >= caretIndex)
            {
                lineStart = currentIndex;
                break;
            }
            currentIndex = nextNewline + 1;
            lineIndex++;
            lineStart = currentIndex;
        }

        var nextNewlineAfter = text.IndexOf('\n', lineStart);
        var lineEnd = nextNewlineAfter == -1 ? text.Length : nextNewlineAfter;
        var lineText = text.Substring(lineStart, lineEnd - lineStart).TrimEnd('\r');

        // Dynamic capitalization of Scene Headings and Transitions
        var trimmedLineText = lineText.TrimStart();
        if (trimmedLineText.Length > 0)
        {
            var spaceIndex = trimmedLineText.IndexOfAny([' ', '\t']);
            var firstToken = spaceIndex < 0 ? trimmedLineText : trimmedLineText.Substring(0, spaceIndex);

            var hasSpaceOrDot = firstToken.Contains('.') || spaceIndex >= 0;
            var isSceneHeading = (hasSpaceOrDot && Passage.Core.TextAnalysis.LooksLikeSceneHeadingStart(firstToken.AsSpan())) || trimmedLineText.StartsWith('.');

            var isTransition = TransitionRegex.IsMatch(trimmedLineText) ||
                               trimmedLineText.EndsWith("TO:", StringComparison.OrdinalIgnoreCase) ||
                               trimmedLineText.EndsWith("TO.", StringComparison.OrdinalIgnoreCase);

            if (isSceneHeading || isTransition)
            {
                if (lineText != lineText.ToUpperInvariant())
                {
                    var upperLine = lineText.ToUpperInvariant();

                    textBox.TextChanged -= EditorBox_TextChanged;
                    try
                    {
                        // Replace just the current line so undo history and caret stay intact.
                        textBox.Document.Replace(lineStart, lineEnd - lineStart, upperLine);
                        textBox.CaretOffset = caretIndex;
                    }
                    finally
                    {
                        textBox.TextChanged += EditorBox_TextChanged;
                    }
                    text = textBox.Text;
                    lineText = upperLine;
                    vm.EditorContent = text;
                }
            }
        }

        var prefixLen = caretIndex - lineStart;

        if (prefixLen < 0 || prefixLen > lineText.Length)
        {
            vm.IsAutoCompleteOpen = false;
            return;
        }

        var prefix = lineText.Substring(0, prefixLen);
        var elementType = vm.GetLatestEffectiveLineType(lineIndex + 1, lineText);
        var elementTypeName = elementType.ToString();

        // If it's an Action line but starts with INT. or EXT., treat it as a Scene Heading for suggestions
        if (elementTypeName == "Action" && prefix.Length >= 3)
        {
            var upperPrefix = prefix.ToUpperInvariant();
            if (upperPrefix.StartsWith("INT.") || upperPrefix.StartsWith("EXT.") || upperPrefix.StartsWith("I/E."))
            {
                elementTypeName = "SceneHeading";
            }
        }

        if (elementTypeName == "SceneHeading")
        {
            var isSceneHeadingPrefix = prefix.StartsWith(".") ||
                                       Passage.Core.TextAnalysis.LooksLikeSceneHeadingStart(prefix.AsSpan(), allowPartialPrefixMatch: true);
            if (!isSceneHeadingPrefix)
            {
                vm.IsAutoCompleteOpen = false;
                return;
            }
        }

        vm.UpdateSuggestions(prefix, elementTypeName);

        if (vm.IsAutoCompleteOpen)
        {
            Avalonia.Threading.Dispatcher.UIThread.Post(() => PositionAutoCompletePopup(), Avalonia.Threading.DispatcherPriority.Background);
        }
    }

    private void EditorBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm || !vm.IsAutoCompleteOpen)
        {
            return;
        }

        var count = vm.AutoCompleteSuggestions.Count;
        if (count == 0) return;

        switch (e.Key)
        {
            case Key.Up:
                vm.SelectedSuggestionIndex = (vm.SelectedSuggestionIndex - 1 + count) % count;
                e.Handled = true;
                break;
            case Key.Down:
                vm.SelectedSuggestionIndex = (vm.SelectedSuggestionIndex + 1) % count;
                e.Handled = true;
                break;
            case Key.Enter:
            case Key.Tab:
                if (vm.SelectedSuggestionIndex >= 0 && vm.SelectedSuggestionIndex < count)
                {
                    ApplyAutoCompleteSuggestion(vm.AutoCompleteSuggestions[vm.SelectedSuggestionIndex]);
                    e.Handled = true;
                }
                break;
            case Key.Escape:
                vm.IsAutoCompleteOpen = false;
                e.Handled = true;
                break;
        }
    }

    private void SuggestionsListBox_Tapped(object? sender, TappedEventArgs e)
    {
        var listBox = sender as ListBox;
        if (listBox != null && listBox.SelectedItem is string suggestion)
        {
            ApplyAutoCompleteSuggestion(suggestion);
        }
    }

    private void PositionAutoCompletePopup()
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        var popup = this.FindControl<Popup>("AutoCompletePopup");
        if (editor == null || popup == null) return;

        try
        {
            // The caret rectangle is in TextView coordinates; translate it into the
            // editor's coordinate space for the popup offsets.
            var textView = editor.TextArea.TextView;
            var caretRect = editor.TextArea.Caret.CalculateCaretRectangle();
            var point = textView.TranslatePoint(caretRect.Position, editor);
            if (point.HasValue)
            {
                popup.Placement = PlacementMode.Top;
                popup.PlacementTarget = editor;
                popup.HorizontalOffset = point.Value.X + 24;
                popup.VerticalOffset = point.Value.Y - 4;
            }
        }
        catch
        {
            // If caret geometry isn't available yet, fall back to the default placement.
        }
    }

    private void ApplyAutoCompleteSuggestion(string suggestion)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        if (editor == null || DataContext is not MainWindowViewModel vm) return;

        var caretIndex = editor.CaretOffset;
        var text = editor.Text ?? "";

        // Find line start and line end
        var lineStart = 0;
        var currentIndex = 0;
        while (currentIndex < caretIndex && currentIndex < text.Length)
        {
            var nextNewline = text.IndexOf('\n', currentIndex);
            if (nextNewline == -1 || nextNewline >= caretIndex)
            {
                lineStart = currentIndex;
                break;
            }
            currentIndex = nextNewline + 1;
            lineStart = currentIndex;
        }

        var nextNewlineAfter = text.IndexOf('\n', lineStart);
        var lineEnd = nextNewlineAfter == -1 ? text.Length : nextNewlineAfter;

        // Replace the current line with the suggestion.
        editor.Document.Replace(lineStart, lineEnd - lineStart, suggestion);
        editor.CaretOffset = lineStart + suggestion.Length;
        vm.EditorContent = editor.Text ?? string.Empty;

        vm.IsAutoCompleteOpen = false;
    }

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);

        if (RecoveryStorage.TryReadRecovery(out var pendingRecoveryDocument) && pendingRecoveryDocument != null)
        {
            var prompt = new RecoveryPromptDialog();
            _ = prompt.ShowDialog<bool>(this).ContinueWith(t =>
            {
                if (t.IsCompletedSuccessfully)
                {
                    var restore = t.Result;
                    if (restore)
                    {
                        if (DataContext is MainWindowViewModel vm)
                        {
                            vm.LoadRecoveryDocument(pendingRecoveryDocument);
                        }
                    }
                    else
                    {
                        RecoveryStorage.ClearRecoveryFile();
                    }
                }
                else
                {
                    RecoveryStorage.ClearRecoveryFile();
                }

                // Re-focus the editor after the dialog closes so the user can keep
                // typing immediately.
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    var editor = this.FindControl<TextEditor>("EditorTextBox");
                    editor?.Focus();
                }, Avalonia.Threading.DispatcherPriority.Input);
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }
        else
        {
            if (SessionStorage.TryLoadSession(out var sessionState) && sessionState != null)
            {
                if (DataContext is MainWindowViewModel vm)
                {
                    vm.LoadSessionState(sessionState);
                }

                if (sessionState.WindowWidth.HasValue && sessionState.WindowHeight.HasValue)
                {
                    this.Width = sessionState.WindowWidth.Value;
                    this.Height = sessionState.WindowHeight.Value;
                }

                if (sessionState.WindowX.HasValue && sessionState.WindowY.HasValue)
                {
                    this.Position = new PixelPoint(sessionState.WindowX.Value, sessionState.WindowY.Value);
                }

                if (!string.IsNullOrEmpty(sessionState.WindowState))
                {
                    if (Enum.TryParse<WindowState>(sessionState.WindowState, out var state))
                    {
                        this.WindowState = state;
                    }
                }
            }

            // Focus the EditorTextBox on startup (no dialog case)
            var editor = this.FindControl<TextEditor>("EditorTextBox");
            editor?.Focus();
        }
    }

    private bool _forceClose;

    protected override void OnClosing(WindowClosingEventArgs e)
    {
        base.OnClosing(e);

        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        // Unsaved changes: hold the close, ask Save / Discard / Cancel, then
        // re-close with the prompt bypassed if the user didn't cancel.
        if (!_forceClose && vm.IsDirty)
        {
            e.Cancel = true;
            _ = PromptThenCloseAsync(vm);
            return;
        }

        vm.SaveSessionNow();
        vm.StopRecoveryAutosave();
    }

    private async Task PromptThenCloseAsync(MainWindowViewModel vm)
    {
        if (await vm.ConfirmLoseChangesAsync())
        {
            _forceClose = true;
            Close();
        }
    }

    public async void ShowGoToLineDialog()
    {
        var dialog = new GoToLineDialog();
        var result = await dialog.ShowDialog<bool>(this);
        if (result)
        {
            NavigateToLine(dialog.LineNumber);
        }
    }

    public async void ShowGoToSceneDialog()
    {
        if (DataContext is not MainWindowViewModel vm) return;
        var dialog = new GoToSceneDialog();
        dialog.SetScenes(vm.OutlineRoots);
        var lineNumber = await dialog.ShowDialog<int?>(this);
        if (lineNumber.HasValue)
        {
            NavigateToLine(lineNumber.Value);
        }
    }

    private FindReplaceDialog? _findReplaceDialog;

    public void ShowFindReplaceDialog()
    {
        if (_findReplaceDialog != null)
        {
            _findReplaceDialog.Activate();
            return;
        }

        _findReplaceDialog = new FindReplaceDialog();
        _findReplaceDialog.Closed += (s, e) => _findReplaceDialog = null;

        _findReplaceDialog.FindNextRequested += () =>
        {
            FindNext(_findReplaceDialog.SearchText, _findReplaceDialog.MatchCase, _findReplaceDialog.WholeWord);
        };

        _findReplaceDialog.ReplaceRequested += () =>
        {
            ReplaceCurrent(_findReplaceDialog.SearchText, _findReplaceDialog.ReplaceText, _findReplaceDialog.MatchCase, _findReplaceDialog.WholeWord);
        };

        _findReplaceDialog.ReplaceAllRequested += () =>
        {
            ReplaceAll(_findReplaceDialog.SearchText, _findReplaceDialog.ReplaceText, _findReplaceDialog.MatchCase, _findReplaceDialog.WholeWord);
        };

        _findReplaceDialog.Show(this);
    }

    public bool FindNext(string searchText, bool matchCase, bool wholeWord)
    {
        return FindText(searchText, forward: true, matchCase, wholeWord);
    }

    public bool FindPrevious(string searchText, bool matchCase, bool wholeWord)
    {
        return FindText(searchText, forward: false, matchCase, wholeWord);
    }

    private bool FindText(string searchText, bool forward, bool matchCase, bool wholeWord)
    {
        var textBox = this.FindControl<TextEditor>("EditorTextBox");
        if (textBox == null || string.IsNullOrEmpty(searchText))
        {
            return false;
        }

        var documentText = textBox.Text ?? string.Empty;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        int matchIndex = -1;
        int start, end;
        if (textBox.SelectionLength > 0)
        {
            start = textBox.SelectionStart;
            end = textBox.SelectionStart + textBox.SelectionLength;
        }
        else
        {
            start = end = textBox.CaretOffset;
        }

        if (forward)
        {
            var startIndex = end;
            if (startIndex > documentText.Length)
            {
                startIndex = documentText.Length;
            }

            matchIndex = FindTextIndex(documentText, searchText, startIndex, forward, comparison, wholeWord);
            if (matchIndex < 0)
            {
                matchIndex = FindTextIndex(documentText, searchText, 0, forward, comparison, wholeWord);
            }
        }
        else
        {
            var startIndex = Math.Max(0, start - 1);
            matchIndex = FindTextIndex(documentText, searchText, startIndex, forward, comparison, wholeWord);
            if (matchIndex < 0)
            {
                matchIndex = FindTextIndex(documentText, searchText, documentText.Length - 1, forward, comparison, wholeWord);
            }
        }

        if (matchIndex < 0)
        {
            return false;
        }

        SelectEditorRange(textBox, matchIndex, searchText.Length);
        return true;
    }

    private int FindTextIndex(string documentText, string searchText, int startIndex, bool forward, StringComparison comparison, bool wholeWord)
    {
        int index = startIndex;
        while (true)
        {
            int foundIndex;
            if (forward)
            {
                if (index > documentText.Length - searchText.Length) return -1;
                foundIndex = documentText.IndexOf(searchText, index, comparison);
            }
            else
            {
                if (index < 0) return -1;
                foundIndex = documentText.LastIndexOf(searchText, index, comparison);
            }

            if (foundIndex < 0) return -1;

            if (wholeWord)
            {
                if (IsWholeWord(documentText, foundIndex, searchText.Length))
                {
                    return foundIndex;
                }
                else
                {
                    if (forward)
                    {
                        index = foundIndex + 1;
                    }
                    else
                    {
                        index = foundIndex - 1;
                    }
                }
            }
            else
            {
                return foundIndex;
            }
        }
    }

    private bool IsWholeWord(string text, int index, int length)
    {
        if (index > 0)
        {
            char prev = text[index - 1];
            if (char.IsLetterOrDigit(prev) || prev == '_') return false;
        }
        if (index + length < text.Length)
        {
            char next = text[index + length];
            if (char.IsLetterOrDigit(next) || next == '_') return false;
        }
        return true;
    }

    private void SelectEditorRange(TextEditor editor, int start, int length)
    {
        var textLength = (editor.Text ?? string.Empty).Length;
        var safeStart = Math.Max(0, Math.Min(start, textLength));
        var safeLength = Math.Max(0, Math.Min(length, textLength - safeStart));

        editor.Focus();
        editor.CaretOffset = safeStart + safeLength;
        editor.Select(safeStart, safeLength);
        var line = editor.Document.GetLineByOffset(safeStart).LineNumber;
        editor.ScrollToLine(line);
    }

    public bool ReplaceCurrent(string searchText, string replacementText, bool matchCase, bool wholeWord)
    {
        var editor = this.FindControl<TextEditor>("EditorTextBox");
        if (editor == null || string.IsNullOrWhiteSpace(searchText))
        {
            return false;
        }

        if (!SelectionMatchesSearch(editor, searchText) && !FindNext(searchText, matchCase, wholeWord))
        {
            return false;
        }

        var start = editor.SelectionStart;
        var length = editor.SelectionLength;
        var replacement = replacementText ?? string.Empty;

        editor.Document.Replace(start, length, replacement);
        if (DataContext is MainWindowViewModel vm)
        {
            vm.EditorContent = editor.Text ?? string.Empty;
        }

        SelectEditorRange(editor, start, replacement.Length);
        return true;
    }

    public int ReplaceAll(string searchText, string replacementText, bool matchCase, bool wholeWord)
    {
        var textBox = this.FindControl<TextEditor>("EditorTextBox");
        if (textBox == null || string.IsNullOrWhiteSpace(searchText))
        {
            return 0;
        }

        var source = textBox.Text ?? string.Empty;
        var replacement = replacementText ?? string.Empty;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var builder = new StringBuilder(source.Length);
        var index = 0;
        var replacements = 0;

        while (index < source.Length)
        {
            var matchIndex = source.IndexOf(searchText, index, comparison);
            if (matchIndex < 0)
            {
                builder.Append(source, index, source.Length - index);
                break;
            }

            if (wholeWord && !IsWholeWord(source, matchIndex, searchText.Length))
            {
                builder.Append(source, index, matchIndex - index + 1);
                index = matchIndex + 1;
                continue;
            }

            builder.Append(source, index, matchIndex - index);
            builder.Append(replacement);
            replacements++;
            index = matchIndex + searchText.Length;
        }

        if (replacements == 0)
        {
            return 0;
        }

        var caretIndex = textBox.CaretOffset;
        var mappedCaret = MapCaretAfterReplaceAll(source, searchText, replacement, caretIndex, matchCase, wholeWord);

        textBox.Document.Text = builder.ToString();
        if (DataContext is MainWindowViewModel vm)
        {
            vm.EditorContent = textBox.Text ?? string.Empty;
        }
        SelectEditorRange(textBox, mappedCaret, 0);
        return replacements;
    }

    private bool SelectionMatchesSearch(TextEditor editor, string searchText)
    {
        var selectedText = editor.SelectedText ?? string.Empty;
        return string.Equals(selectedText, searchText, StringComparison.OrdinalIgnoreCase);
    }

    private static int MapCaretAfterReplaceAll(string source, string searchText, string replacementText, int caretIndex, bool matchCase, bool wholeWord)
    {
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var sourceIndex = 0;
        var targetIndex = 0;

        while (sourceIndex < source.Length)
        {
            var matchIndex = source.IndexOf(searchText, sourceIndex, comparison);
            if (matchIndex < 0)
            {
                if (caretIndex <= source.Length)
                {
                    return targetIndex + Math.Max(0, caretIndex - sourceIndex);
                }

                return targetIndex;
            }

            if (wholeWord && !IsWholeWordStatic(source, matchIndex, searchText.Length))
            {
                var step = matchIndex - sourceIndex + 1;
                targetIndex += step;
                sourceIndex = matchIndex + 1;
                continue;
            }

            if (caretIndex < matchIndex)
            {
                return targetIndex + (caretIndex - sourceIndex);
            }

            targetIndex += matchIndex - sourceIndex;
            sourceIndex = matchIndex;

            if (caretIndex < matchIndex + searchText.Length)
            {
                return targetIndex + replacementText.Length;
            }

            targetIndex += replacementText.Length;
            sourceIndex += searchText.Length;
        }

        return targetIndex;
    }
    private static bool IsWholeWordStatic(string text, int index, int length)
    {
        if (index > 0)
        {
            char prev = text[index - 1];
            if (char.IsLetterOrDigit(prev) || prev == '_') return false;
        }
        if (index + length < text.Length)
        {
            char next = text[index + length];
            if (char.IsLetterOrDigit(next) || next == '_') return false;
        }
        return true;
    }

    // Drag and Drop Event Handlers for Outline Nodes
    // Distinguishes a click (navigate to the element) from a drag (reorder).
    private OutlineNodeViewModel? _pressedOutlineNode;
    private PointerPressedEventArgs? _outlinePressArgs;
    private Point _outlinePressPoint;
    private bool _outlineDragging;

    private void OutlineNode_PointerPressed(object? sender, PointerPressedEventArgs e)
    {
        var properties = e.GetCurrentPoint(this).Properties;
        if (!properties.IsLeftButtonPressed) return;

        if (sender is Visual visual && visual.DataContext is OutlineNodeViewModel node)
        {
            _pressedOutlineNode = node;
            _outlinePressArgs = e;
            _outlinePressPoint = e.GetPosition(this);
            _outlineDragging = false;
        }
    }

    private async void OutlineNode_PointerMoved(object? sender, PointerEventArgs e)
    {
        if (_pressedOutlineNode == null || _outlineDragging || _outlinePressArgs == null) return;
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;

        var delta = e.GetPosition(this) - _outlinePressPoint;
        // Only begin a drag once the pointer moves past a small threshold,
        // otherwise a plain click would be swallowed by the drag operation.
        if (Math.Abs(delta.X) < 4 && Math.Abs(delta.Y) < 4) return;

        _outlineDragging = true;
        DraggedOutlineNode = _pressedOutlineNode;
        var dragData = new DataTransfer();
        dragData.Add(DataTransferItem.Create(DataFormat.Text, "OutlineNode"));
        await DragDrop.DoDragDropAsync(_outlinePressArgs, dragData, DragDropEffects.Move);
    }

    private void OutlineNode_PointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        // A press with no meaningful movement is a click: move the caret to the
        // start of this element on the page.
        if (_pressedOutlineNode != null && !_outlineDragging)
        {
            NavigateToLine(_pressedOutlineNode.LineNumber);
        }

        _pressedOutlineNode = null;
        _outlinePressArgs = null;
        _outlineDragging = false;
    }

    private OutlineNodeViewModel? _lastDragOverNode;

    private void OutlineNode_DragOver(object? sender, DragEventArgs e)
    {
        if (DraggedOutlineNode == null)
        {
            e.DragEffects = DragDropEffects.None;
            return;
        }

        if (sender is Visual visual && visual.DataContext is OutlineNodeViewModel targetNode)
        {
            if (DraggedOutlineNode == targetNode)
            {
                e.DragEffects = DragDropEffects.None;
                return;
            }

            if (_lastDragOverNode != targetNode)
            {
                if (_lastDragOverNode != null)
                {
                    _lastDragOverNode.IsDragOver = false;
                }
                _lastDragOverNode = targetNode;
                targetNode.IsDragOver = true;
            }

            e.DragEffects = DragDropEffects.Move;
            e.Handled = true;
        }
    }

    private void OutlineNode_DragLeave(object? sender, RoutedEventArgs e)
    {
        if (sender is Visual visual && visual.DataContext is OutlineNodeViewModel targetNode)
        {
            targetNode.IsDragOver = false;
            if (_lastDragOverNode == targetNode)
            {
                _lastDragOverNode = null;
            }
        }
    }

    private void OutlineNode_Drop(object? sender, DragEventArgs e)
    {
        if (_lastDragOverNode != null)
        {
            _lastDragOverNode.IsDragOver = false;
            _lastDragOverNode = null;
        }

        if (DataContext is not MainWindowViewModel vm) return;

        if (DraggedOutlineNode != null)
        {
            if (sender is Visual visual && visual.DataContext is OutlineNodeViewModel targetNode)
            {
                if (DraggedOutlineNode == targetNode) return;

                var position = e.GetPosition(visual);
                double height = visual.Bounds.Height;
                WorkspaceDropPosition dropPos;

                if (position.Y < height * 0.25)
                {
                    dropPos = WorkspaceDropPosition.Above;
                }
                else if (position.Y > height * 0.75)
                {
                    dropPos = WorkspaceDropPosition.Below;
                }
                else
                {
                    dropPos = WorkspaceDropPosition.Onto;
                }

                vm.MoveOutlineNodeText(DraggedOutlineNode, targetNode, dropPos);
                e.Handled = true;
            }
        }
        DraggedOutlineNode = null;
    }

    // Drag and Drop Event Handlers for Beat Board Cards
    private async void BeatBoardCard_PointerPressed(object? sender, PointerPressedEventArgs e)
    {
        var properties = e.GetCurrentPoint(this).Properties;
        if (!properties.IsLeftButtonPressed) return;

        var source = e.Source as Visual;
        while (source != null)
        {
            if (source is TextBox || source is ComboBox || source is Button)
            {
                return;
            }
            source = source.GetVisualParent();
        }

        if (sender is Visual visual && visual.DataContext is BeatBoardCardViewModel card)
        {
            if (card.IsEditing) return;

            DraggedBeatBoardCard = card;
            var dragData = new DataTransfer();
            dragData.Add(DataTransferItem.Create(DataFormat.Text, "BeatBoardCard"));
            var result = await DragDrop.DoDragDropAsync(e, dragData, DragDropEffects.Move);
        }
    }

    private void BeatBoardCard_DragOver(object? sender, DragEventArgs e)
    {
        if (DraggedBeatBoardCard == null)
        {
            e.DragEffects = DragDropEffects.None;
            return;
        }

        if (sender is Visual visual && visual.DataContext is BeatBoardCardViewModel targetCard)
        {
            if (DraggedBeatBoardCard == targetCard)
            {
                e.DragEffects = DragDropEffects.None;
                return;
            }

            e.DragEffects = DragDropEffects.Move;
            e.Handled = true;
        }
    }

    private void BeatBoardCard_Drop(object? sender, DragEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        if (DraggedBeatBoardCard != null)
        {
            if (sender is Visual visual && visual.DataContext is BeatBoardCardViewModel targetCard)
            {
                if (DraggedBeatBoardCard == targetCard) return;

                var position = e.GetPosition(visual);
                bool insertAfter = position.X > visual.Bounds.Width / 2;

                vm.MoveBeatBoardCardText(DraggedBeatBoardCard, targetCard, insertAfter);
                e.Handled = true;
            }
        }
        DraggedBeatBoardCard = null;
    }
}
