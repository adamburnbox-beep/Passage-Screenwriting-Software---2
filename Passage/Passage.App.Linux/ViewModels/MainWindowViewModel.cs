using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.Threading;
using Passage.Parser;
using Passage.Parser.Goals;
using Passage.Export;
using Passage.Core.Goals;
using Passage.App.Services;
using Passage.Core.Services;

namespace Passage.App.ViewModels;

public enum WriteMode
{
    Screenplay,
    Markdown
}

public partial class MainWindowViewModel : ViewModelBase
{
    private static readonly GoalType[] GoalTypeOptions =
    [
        GoalType.WordCount,
        GoalType.PageCount
    ];
    private static readonly GoalType[] SessionGoalTypeOptions =
    [
        GoalType.WordCount,
        GoalType.PageCount,
        GoalType.Timer
    ];

    private readonly Window? _window;
    private string _currentFilePath = string.Empty;
    private string _editorContent = string.Empty;
    private double _editorZoomScale = 1.0;
    private bool _suppressDirtyTracking;
    private bool _isDirty;

    // Which main view is showing: 0 = Script, 1 = Beat Board, 2 = Page Preview.
    // A single index (rather than per-tab IsSelected bindings) so selecting one
    // tab can never re-select another.
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsBoardModeActive))]
    private int _mainViewIndex;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsScreenplayMode))]
    [NotifyPropertyChangedFor(nameof(IsMarkdownMode))]
    [NotifyPropertyChangedFor(nameof(ModeStatusText))]
    [NotifyPropertyChangedFor(nameof(BeatBoardEmptyMessage))]
    [NotifyPropertyChangedFor(nameof(OutlineEmptyMessage))]
    private WriteMode _currentMode = WriteMode.Screenplay;

    public bool IsScreenplayMode => CurrentMode == WriteMode.Screenplay;
    public bool IsMarkdownMode => CurrentMode == WriteMode.Markdown;
    public string ModeStatusText => CurrentMode == WriteMode.Screenplay ? "MODE: SCREENPLAY" : "MODE: MARKDOWN";

    [ObservableProperty] private int _currentLineNumber = 1;
    [ObservableProperty] private string _currentElementType = "Action";
    [ObservableProperty] private string _currentElementText = string.Empty;
    [ObservableProperty] private string _statusMessage = "Ready";

    public string EnterContinuationText => "NewLine";
    public string ActiveZoomDisplayText => $"{(int)Math.Round(EditorZoomScale * 100)}%";

    // Services and parser
    private readonly DispatcherTimer _refreshTimer;
    private readonly DispatcherTimer _recoveryTimer;
    private readonly FountainParser _parser = new();
    private readonly Dictionary<int, ScreenplayElementType> _lineTypeOverrides = new();
    private readonly ScreenplayPageEstimator _pageEstimator = new();
    private readonly GoalProgressCalculator _goalProgressCalculator = new();

    // Throttled Parsing state
    private int _outlineRefreshVersion;
    private bool _outlineRefreshPending;
    private bool _outlineRefreshInProgress;
    private TimeSpan _lastParseDuration = TimeSpan.Zero;
    private ParsedScreenplay _lastParsed = new ParsedScreenplay(string.Empty, Array.Empty<ScreenplayElement>());

    // Goal Configuration Backing fields
    private GoalType _selectedGoalType = GoalType.WordCount;
    private int _wordCountGoalTargetValue = 1000;
    private int _pageCountGoalTargetValue = 120;
    private int _timerGoalTargetMinutes = 25;
    private GoalType _sessionSelectedGoalType = GoalType.WordCount;
    private int _sessionWordCountTargetValue = 1000;
    private int _sessionPageCountTargetValue = 10;
    private int _sessionWordCountBaseline;
    private int _sessionPageCountBaseline;
    private bool _sessionGoalBaselineNeedsCapture = true;
    private GoalTimerRuntime? _goalTimerRuntime;

    // Goal Display backing fields
    [ObservableProperty] private double _goalProgressPercent;
    [ObservableProperty] private double _sessionGoalProgressPercent;
    [ObservableProperty] private string _goalCurrentDisplayText = "0 words";
    [ObservableProperty] private string _sessionGoalCurrentDisplayText = "0 words";
    [ObservableProperty] private string _goalTargetDisplayText = "1,000 words";
    [ObservableProperty] private string _sessionGoalTargetDisplayText = "1,000 words";
    [ObservableProperty] private string _goalStateText = "In progress";
    [ObservableProperty] private string _sessionGoalStateText = "In progress";
    [ObservableProperty] private string _goalTargetUnitLabel = "Words";
    [ObservableProperty] private string _sessionGoalTargetUnitLabel = "Words";
    [ObservableProperty] private string _goalTimerElapsedText = "00:00";
    [ObservableProperty] private string _goalTimerRemainingText = "25:00";
    [ObservableProperty] private string _goalProgressSummaryText = "Ready";

    // Sidebar tree fields
    [ObservableProperty] private ObservableCollection<OutlineNodeViewModel> _outlineRoots = new();
    [ObservableProperty] private ObservableCollection<OutlineNodeViewModel> _notesRoots = new();
    [ObservableProperty] private ObservableCollection<OutlineNodeViewModel> _outlineNodes = new();
    [ObservableProperty] private int? _selectedOutlineLineNumber;
    [ObservableProperty] private ObservableCollection<PreviewPageViewModel> _previewPages = new();

    // Scratchpad fields
    [ObservableProperty] private ObservableCollection<ScreenplayElement> _scratchpadElements = new();
    [ObservableProperty] private string _scratchpadSearchText = string.Empty;
    [ObservableProperty] private ScreenplayElement? _selectedScratchpadElement;

    // Autocomplete fields
    [ObservableProperty] private bool _isAutoCompleteOpen;
    [ObservableProperty] private int _selectedSuggestionIndex;
    public ObservableCollection<string> AutoCompleteSuggestions { get; } = new();

    private readonly HashSet<string> _uniqueSceneHeadings = new();
    private readonly HashSet<string> _uniqueCharacterNames = new();

    public MainWindowViewModel() : this(null) { }

    public MainWindowViewModel(Window? window)
    {
        _window = window;
        BeatBoardCards = new ObservableCollection<BeatBoardCardViewModel>();
        AvailableExporters = new ObservableCollection<IExporter>(ExporterCatalog.GetDefaultExporters());

        _refreshTimer = new DispatcherTimer();
        _refreshTimer.Tick += (_, _) => RefreshParsedDocument();

        _recoveryTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _recoveryTimer.Tick += (_, _) => SaveRecoverySnapshot();

        // Initial baseline capture
        CaptureSessionGoalBaseline();
        RefreshGoalState();
    }

    [ObservableProperty]
    private string _windowTitle = "Passage - Untitled";

    public bool IsDirty
    {
        get => _isDirty;
        set
        {
            if (_suppressDirtyTracking || !SetProperty(ref _isDirty, value))
            {
                return;
            }

            if (_isDirty)
            {
                StartRecoveryAutosave();
            }
            else
            {
                StopRecoveryAutosave();
            }

            UpdateWindowTitle();
        }
    }

    private void UpdateWindowTitle()
    {
        var name = string.IsNullOrEmpty(_currentFilePath) ? "Untitled" : Path.GetFileName(_currentFilePath);
        WindowTitle = IsDirty ? $"Passage - {name} *" : $"Passage - {name}";
    }

    public string EditorContent
    {
        get => _editorContent;
        set
        {
            if (SetProperty(ref _editorContent, value))
            {
                OnContentChanged();
            }
        }
    }

    public double EditorZoomScale
    {
        get => _editorZoomScale;
        set
        {
            if (SetProperty(ref _editorZoomScale, value))
            {
                OnPropertyChanged(nameof(ActiveZoomDisplayText));
            }
        }
    }

    public bool IsBoardModeActive
    {
        get => MainViewIndex == 1;
        set => MainViewIndex = value ? 1 : 0;
    }

    public ObservableCollection<BeatBoardCardViewModel> BeatBoardCards { get; }
    public ObservableCollection<BeatBoardLaneViewModel> BeatBoardLanes { get; } = new();
    public ObservableCollection<IExporter> AvailableExporters { get; }

    public bool HasBeatBoardCards => BeatBoardCards.Count > 0;
    public string BeatBoardEmptyMessage => IsMarkdownMode
        ? "The Beat Board works in Screenplay mode.\nPress Ctrl+M to switch back to Fountain."
        : "The board mirrors your script's structure.\nAdd Acts (#), Sequences (##), and Scenes to see nested index cards here,\nor click New Card to start.";

    // Most-recently-used files, newest first; persisted with the session.
    private const int MaxRecentFiles = 8;
    public ObservableCollection<string> RecentFiles { get; } = new();
    public bool HasRecentFiles => RecentFiles.Count > 0;

    private void AddRecentFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        RecentFiles.Remove(path);
        RecentFiles.Insert(0, path);
        while (RecentFiles.Count > MaxRecentFiles)
        {
            RecentFiles.RemoveAt(RecentFiles.Count - 1);
        }
        OnPropertyChanged(nameof(HasRecentFiles));
    }

    [ObservableProperty] private string _pageCountStatusText = "~1 page";

    // Sidebar properties
    public bool HasOutlineItems => OutlineRoots.Count > 0;
    public bool HasNoteItems => NotesRoots.Count > 0;
    public string OutlineEmptyMessage => IsMarkdownMode
        ? "Markdown headings (#, ##, ###) will appear here."
        : "Sections, synopses, and scene headings will appear here.";
    public string NotesEmptyMessage => "Notes will appear here.";

    // Scratchpad properties
    public bool HasScratchpadItems => ScratchpadElements.Count > 0;
    public string ScratchpadEmptyMessage => string.IsNullOrWhiteSpace(ScratchpadSearchText)
        ? "Moved scenes, notes, and loose ideas will appear here."
        : "No scratchpad cards match the current search.";

    // Goal Configuration and choices
    public GoalConfiguration GoalConfiguration => new(
        SelectedGoalType,
        _wordCountGoalTargetValue,
        _pageCountGoalTargetValue,
        _timerGoalTargetMinutes);

    public IReadOnlyList<GoalType> GoalTypes => GoalTypeOptions;

    public SessionGoalConfiguration SessionGoalConfiguration => new(
        SessionSelectedGoalType,
        _sessionWordCountTargetValue,
        _sessionPageCountTargetValue,
        _timerGoalTargetMinutes);

    public IReadOnlyList<GoalType> SessionGoalTypes => SessionGoalTypeOptions;

    public GoalType SelectedGoalType
    {
        get => _selectedGoalType;
        set
        {
            var normalizedValue = value == GoalType.Timer ? GoalType.WordCount : value;
            if (SetProperty(ref _selectedGoalType, normalizedValue))
            {
                OnPropertyChanged(nameof(GoalTargetValue));
                RefreshGoalState();
            }
        }
    }

    public GoalType SessionSelectedGoalType
    {
        get => _sessionSelectedGoalType;
        set
        {
            if (SetProperty(ref _sessionSelectedGoalType, value))
            {
                if (value != GoalType.Timer)
                {
                    _goalTimerRuntime?.Stop();
                }

                OnPropertyChanged(nameof(SessionGoalTargetValue));
                OnPropertyChanged(nameof(IsTimerGoal));
                RefreshGoalState();
            }
        }
    }

    public int SessionGoalTargetValue
    {
        get
        {
            return SessionSelectedGoalType switch
            {
                GoalType.WordCount => _sessionWordCountTargetValue,
                GoalType.PageCount => _sessionPageCountTargetValue,
                GoalType.Timer => _timerGoalTargetMinutes,
                _ => _sessionWordCountTargetValue
            };
        }
        set
        {
            var clampedValue = Math.Max(0, value);

            switch (SessionSelectedGoalType)
            {
                case GoalType.WordCount:
                    if (_sessionWordCountTargetValue == clampedValue) return;
                    _sessionWordCountTargetValue = clampedValue;
                    break;
                case GoalType.PageCount:
                    if (_sessionPageCountTargetValue == clampedValue) return;
                    _sessionPageCountTargetValue = clampedValue;
                    break;
                case GoalType.Timer:
                    if (_timerGoalTargetMinutes == clampedValue) return;
                    _timerGoalTargetMinutes = clampedValue;
                    RebuildGoalTimerRuntime();
                    break;
            }

            OnPropertyChanged(nameof(SessionGoalTargetValue));
            RefreshGoalState();
        }
    }

    public int GoalTargetValue
    {
        get
        {
            return SelectedGoalType switch
            {
                GoalType.WordCount => _wordCountGoalTargetValue,
                GoalType.PageCount => _pageCountGoalTargetValue,
                _ => _wordCountGoalTargetValue
            };
        }
        set
        {
            var clampedValue = Math.Max(0, value);

            switch (SelectedGoalType)
            {
                case GoalType.WordCount:
                    if (_wordCountGoalTargetValue == clampedValue) return;
                    _wordCountGoalTargetValue = clampedValue;
                    break;
                case GoalType.PageCount:
                    if (_pageCountGoalTargetValue == clampedValue) return;
                    _pageCountGoalTargetValue = clampedValue;
                    break;
            }

            OnPropertyChanged(nameof(GoalTargetValue));
            RefreshGoalState();
        }
    }

    public bool IsTimerGoal => SessionSelectedGoalType == GoalType.Timer;

    public TimerGoalState GoalTimerState => _goalTimerRuntime?.State ?? TimerGoalState.Idle;

    public string GoalTimerPrimaryButtonText =>
        GoalTimerState switch
        {
            TimerGoalState.Running => "Pause",
            TimerGoalState.Paused => "Resume",
            _ => "Start"
        };

    public string GoalTimerSecondaryButtonText =>
        GoalTimerState is TimerGoalState.Running or TimerGoalState.Paused ? "Stop" : "Reset";

    [RelayCommand]
    private async Task New()
    {
        if (!await ConfirmLoseChangesAsync()) return;
        ClearDocument();
    }

    private void ClearDocument()
    {
        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Empty;
            _currentFilePath = string.Empty;
            IsDirty = false;
            ResetSessionGoal();
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        UpdateWindowTitle();
        (_window as Views.MainWindow)?.ResetEditorUndoHistory();
    }

    /// <summary>
    /// Gate for any action that discards the current document. Returns true when
    /// it is safe to proceed: the document is clean, the user saved it, or the
    /// user explicitly chose to discard. Cancel (or a failed/cancelled save)
    /// returns false.
    /// </summary>
    public async Task<bool> ConfirmLoseChangesAsync()
    {
        if (!IsDirty || _window == null) return true;

        var dialog = new Views.UnsavedChangesDialog();
        var choice = await dialog.ShowDialog<Views.UnsavedChangesChoice>(_window);

        switch (choice)
        {
            case Views.UnsavedChangesChoice.Save:
                await Save();
                return !IsDirty;
            case Views.UnsavedChangesChoice.Discard:
                return true;
            default:
                return false;
        }
    }

    [RelayCommand]
    private async Task Open()
    {
        if (_window == null) return;
        if (!await ConfirmLoseChangesAsync()) return;

        var files = await _window.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open Screenplay",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Fountain Files") { Patterns = new[] { "*.fountain" } },
                new FilePickerFileType("Markdown Files") { Patterns = new[] { "*.md" } },
                new FilePickerFileType("All Files") { Patterns = new[] { "*.*" } }
            }
        });

        if (files.Count > 0)
        {
            await OpenDocumentFromPathAsync(files[0].Path.LocalPath);
        }
    }

    // Parameter is object: the style that wires the recent-file submenu also
    // matches its parent MenuItem (whose DataContext is this VM), and a
    // strongly-typed RelayCommand throws on that parameter during CanExecute.
    [RelayCommand]
    private async Task OpenRecent(object? parameter)
    {
        if (parameter is not string path) return;
        if (!await ConfirmLoseChangesAsync()) return;

        if (!File.Exists(path))
        {
            StatusMessage = $"File not found: {path}";
            RecentFiles.Remove(path);
            OnPropertyChanged(nameof(HasRecentFiles));
            return;
        }

        await OpenDocumentFromPathAsync(path);
    }

    private async Task OpenDocumentFromPathAsync(string path)
    {
        string text;
        try
        {
            text = await File.ReadAllTextAsync(path);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not open {Path.GetFileName(path)}: {ex.Message}";
            return;
        }

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = text;
            _currentFilePath = path;
            IsDirty = false;
            ResetSessionGoal();

            // Auto-detect mode based on file extension
            var ext = Path.GetExtension(path);
            CurrentMode = string.Equals(ext, ".md", StringComparison.OrdinalIgnoreCase)
                ? WriteMode.Markdown
                : WriteMode.Screenplay;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        UpdateWindowTitle();
        AddRecentFile(path);
        (_window as Views.MainWindow)?.ResetEditorUndoHistory();
        StatusMessage = $"Opened {Path.GetFileName(path)}";
        RefreshParsedDocument();
    }

    [RelayCommand]
    private async Task Save()
    {
        if (string.IsNullOrEmpty(_currentFilePath))
        {
            await SaveAs();
            return;
        }

        try
        {
            await File.WriteAllTextAsync(_currentFilePath, EditorContent);
            IsDirty = false;
            RecoveryStorage.ClearRecoveryFile();
            AddRecentFile(_currentFilePath);
            StatusMessage = $"Saved {Path.GetFileName(_currentFilePath)}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Save failed: {ex.Message}";
            System.Diagnostics.Debug.WriteLine($"Error saving file: {ex.Message}");
        }
    }

    [RelayCommand]
    private async Task SaveAs()
    {
        if (_window == null) return;

        var defaultExt = CurrentMode == WriteMode.Markdown ? ".md" : ".fountain";
        var fileTypeChoices = CurrentMode == WriteMode.Markdown
            ? new[] { new FilePickerFileType("Markdown Files") { Patterns = new[] { "*.md" } } }
            : new[] { new FilePickerFileType("Fountain Files") { Patterns = new[] { "*.fountain" } } };

        var file = await _window.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save Screenplay",
            DefaultExtension = defaultExt,
            FileTypeChoices = fileTypeChoices
        });

        if (file != null)
        {
            try
            {
                _currentFilePath = file.Path.LocalPath;
                await File.WriteAllTextAsync(_currentFilePath, EditorContent);
                IsDirty = false;
                RecoveryStorage.ClearRecoveryFile();
                AddRecentFile(_currentFilePath);
                StatusMessage = $"Saved {Path.GetFileName(_currentFilePath)}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Save failed: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Error saving file: {ex.Message}");
            }
        }
    }

    [RelayCommand]
    private async Task Close()
    {
        if (!await ConfirmLoseChangesAsync()) return;
        ClearDocument();
    }

    [RelayCommand]
    private void Exit()
    {
        _window?.Close();
    }

    [RelayCommand]
    private void Undo()
    {
        (_window as Views.MainWindow)?.UndoEditor();
    }

    [RelayCommand]
    private void Redo()
    {
        (_window as Views.MainWindow)?.RedoEditor();
    }

    [RelayCommand]
    private void Find()
    {
        if (_window is Views.MainWindow mainWin)
        {
            mainWin.ShowFindReplaceDialog();
        }
    }

    [RelayCommand]
    private void ZoomIn()
    {
        if (EditorZoomScale < 2.0)
        {
            EditorZoomScale += 0.1;
        }
    }

    [RelayCommand]
    private void ZoomOut()
    {
        if (EditorZoomScale > 0.5)
        {
            EditorZoomScale -= 0.1;
        }
    }

    [RelayCommand]
    private void ResetZoom()
    {
        EditorZoomScale = 1.0;
    }

    [RelayCommand]
    private void SetDarkTheme()
    {
        if (App.Current is App app)
        {
            app.RequestedThemeVariant = ThemeVariant.Dark;
            app.LoadThemeResources(isLight: false);
            (_window as Views.MainWindow)?.RedrawEditor();
        }
    }

    [RelayCommand]
    private void SetLightTheme()
    {
        if (App.Current is App app)
        {
            app.RequestedThemeVariant = ThemeVariant.Light;
            app.LoadThemeResources(isLight: true);
            (_window as Views.MainWindow)?.RedrawEditor();
        }
    }

    [RelayCommand]
    private void GoToLine()
    {
        if (_window is Views.MainWindow mainWin)
        {
            mainWin.ShowGoToLineDialog();
        }
    }

    [RelayCommand]
    private void GoToScene()
    {
        if (_window is Views.MainWindow mainWin)
        {
            mainWin.ShowGoToSceneDialog();
        }
    }

    [RelayCommand]
    private void ToggleSyntaxPanel()
    {
        if (_window is Views.MainWindow mainWin)
        {
            mainWin.ToggleSyntaxPanel();
        }
    }

    [RelayCommand]
    private void CreateNewCard()
    {
        var newId = Guid.NewGuid();
        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();

        var newCardLines = new List<string>
        {
            "",
            $". New Scene [[id:{newId}]]",
            "= Click the pencil to edit. Double-click to locate in script."
        };

        lines.AddRange(newCardLines);

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Join("\n", lines);
            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        RefreshParsedDocument();
    }

    [RelayCommand]
    private void StartCardEdit(BeatBoardCardViewModel card)
    {
        if (card == null) return;
        foreach (var c in BeatBoardCards)
        {
            c.IsEditing = false;
        }
        card.EditingHeading = card.Heading;
        card.EditingDescription = card.Description;
        card.EditingType = card.Type;
        card.IsEditing = true;
    }

    [RelayCommand]
    private void CancelCardEdit(BeatBoardCardViewModel card)
    {
        if (card != null)
        {
            card.IsEditing = false;
        }
    }

    [RelayCommand]
    private void SaveCard(BeatBoardCardViewModel card)
    {
        if (card == null) return;
        card.Heading = card.EditingHeading;
        card.Description = card.EditingDescription;
        card.Type = card.EditingType;
        card.IsEditing = false;

        UpdateCardInScript(card);
    }

    [RelayCommand]
    private void SyncBoardToScript()
    {
        RefreshParsedDocument();
    }

    [RelayCommand]
    private void ExpandAllBeatBoard() => SetBeatBoardExpansion(true);

    [RelayCommand]
    private void CollapseAllBeatBoard() => SetBeatBoardExpansion(false);

    private void SetBeatBoardExpansion(bool isExpanded)
    {
        foreach (var lane in BeatBoardLanes)
        {
            // Header-less (implicit) lanes/groups have no chevron to reopen
            // them, so collapse-all leaves them visible.
            if (lane.HasActCard)
            {
                lane.IsExpanded = isExpanded;
            }

            foreach (var group in lane.Groups)
            {
                if (group.HasSequenceCard)
                {
                    group.IsExpanded = isExpanded;
                }
            }
        }
    }

    // "+ Scene" on an act lane or sequence group: inserts a new scene at the end
    // of that container's block. A null container appends to the document.
    [RelayCommand]
    private void AddSceneToBlock(BeatBoardCardViewModel? container)
    {
        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();

        int insertAt = lines.Count;
        if (container != null)
        {
            var (start, end) = GetBeatBoardCardLineRange(container);
            if (start != -1)
            {
                insertAt = Math.Min(end + 1, lines.Count);
            }
        }

        var newId = Guid.NewGuid();
        lines.InsertRange(insertAt, new[]
        {
            "",
            $". NEW SCENE [[id:{newId}]]",
            "= Click Edit to describe this scene."
        });

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Join("\n", lines);
            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        RefreshParsedDocument();
    }

    // Delete from the card's edit header. Scenes/sections take their whole block
    // (nested content included), so this always confirms first.
    [RelayCommand]
    private async Task DeleteCard(BeatBoardCardViewModel card)
    {
        if (card == null) return;

        var (start, end) = card.Type == "Note"
            ? GetCardOwnLineRange(card)
            : GetBeatBoardCardLineRange(card);
        if (start == -1) return;

        if (_window != null)
        {
            var scope = card.Type is "Act" or "Sequence"
                ? $"the {card.Type.ToLowerInvariant()} \"{card.Heading}\" and everything nested inside it"
                : $"\"{card.Heading}\" and its contents";
            var dialog = new Views.ConfirmDialog(
                "Delete Card",
                $"This removes {scope} from the script ({end - start + 1} line(s)). Continue?",
                confirmText: "Delete");
            var confirmed = await dialog.ShowDialog<bool>(_window);
            if (!confirmed) return;
        }

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();
        int count = Math.Min(end - start + 1, lines.Count - start);
        if (start < 0 || start >= lines.Count || count <= 0) return;

        lines.RemoveRange(start, count);

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Join("\n", lines);
            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        card.IsEditing = false;
        StatusMessage = $"Deleted {card.Type.ToLowerInvariant()} \"{card.Heading}\"";
        RefreshParsedDocument();
    }

    private void UpdateCardInScript(BeatBoardCardViewModel card)
    {
        var (startLineIdx, endLineIdx) = GetCardOwnLineRange(card);
        if (startLineIdx == -1) return;

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();

        var replacementLines = new List<string>();
        string typeStr = card.Type;
        if (typeStr == "Act" || typeStr == "Sequence" || typeStr == "Section")
        {
            int depth = typeStr == "Act" ? 1 : typeStr == "Sequence" ? 2 : 3;
            var prefix = new string('#', depth);
            replacementLines.Add($"{prefix} {card.Heading.Trim()} [[id:{card.Id}]]");
        }
        else if (typeStr == "Scene")
        {
            var heading = card.Heading.Trim();
            if (heading.StartsWith("INT.", StringComparison.OrdinalIgnoreCase) ||
                heading.StartsWith("EXT.", StringComparison.OrdinalIgnoreCase) ||
                heading.StartsWith("I/E.", StringComparison.OrdinalIgnoreCase) ||
                heading.StartsWith("."))
            {
                replacementLines.Add($"{heading} [[id:{card.Id}]]");
            }
            else
            {
                replacementLines.Add($". {heading} [[id:{card.Id}]]");
            }
        }
        else if (typeStr == "Note")
        {
            replacementLines.Add($"[[{card.Heading.Trim()} id:{card.Id}]]");
        }

        if (!string.IsNullOrWhiteSpace(card.Description))
        {
            foreach (var descLine in card.Description.Split('\n'))
            {
                var trimmed = descLine.Trim();
                if (trimmed.Length > 0)
                {
                    replacementLines.Add($"= {trimmed}");
                }
            }
        }

        int countToRemove = endLineIdx - startLineIdx + 1;
        if (startLineIdx >= 0 && startLineIdx < lines.Count)
        {
            lines.RemoveRange(startLineIdx, Math.Min(countToRemove, lines.Count - startLineIdx));
            lines.InsertRange(startLineIdx, replacementLines);

            _suppressDirtyTracking = true;
            try
            {
                EditorContent = string.Join("\n", lines);
                IsDirty = true;
            }
            finally
            {
                _suppressDirtyTracking = false;
            }
        }
    }

    [RelayCommand]
    private void ClassifyAsAction() => SetLineTypeOverride(ScreenplayElementType.Action);

    [RelayCommand]
    private void ClassifyAsSceneHeading() => SetLineTypeOverride(ScreenplayElementType.SceneHeading);

    [RelayCommand]
    private void ClassifyAsCharacter() => SetLineTypeOverride(ScreenplayElementType.Character);

    [RelayCommand]
    private void ClassifyAsDialogue() => SetLineTypeOverride(ScreenplayElementType.Dialogue);

    [RelayCommand]
    private void ClassifyAsParenthetical() => SetLineTypeOverride(ScreenplayElementType.Parenthetical);

    [RelayCommand]
    private void ClassifyAsTransition() => SetLineTypeOverride(ScreenplayElementType.Transition);

    [RelayCommand]
    private void ClassifyAsNote() => SetLineTypeOverride(ScreenplayElementType.Note);

    private void SetLineTypeOverride(ScreenplayElementType type)
    {
        if (_window is Views.MainWindow mainWin)
        {
            var textBox = mainWin.FindControl<AvaloniaEdit.TextEditor>("EditorTextBox");
            if (textBox != null)
            {
                var caretIndex = textBox.CaretOffset;
                var text = textBox.Text ?? "";

                int lineNumber = GetLineNumberFromCaretIndex(text, caretIndex);
                _lineTypeOverrides[lineNumber] = type;

                OnContentChanged();
            }
        }
    }

    private int GetLineNumberFromCaretIndex(string text, int caretIndex)
    {
        int line = 1;
        for (int i = 0; i < caretIndex && i < text.Length; i++)
        {
            if (text[i] == '\n')
            {
                line++;
            }
        }
        return line;
    }

    [RelayCommand]
    private async Task EditTitlePage()
    {
        if (_window == null) return;

        var dialogViewModel = new TitlePageViewModel();
        if (_lastParsed != null && _lastParsed.TitlePage != null)
        {
            foreach (var entry in _lastParsed.TitlePage.Entries)
            {
                var val = entry.Value;
                switch (entry.FieldType)
                {
                    case TitlePageFieldType.Title: dialogViewModel.Title = val; break;
                    case TitlePageFieldType.Episode: dialogViewModel.Episode = val; break;
                    case TitlePageFieldType.Credit: dialogViewModel.Credit = val; break;
                    case TitlePageFieldType.Author: dialogViewModel.Author = val; break;
                    case TitlePageFieldType.Source: dialogViewModel.Source = val; break;
                    case TitlePageFieldType.Contact: dialogViewModel.Contact = val; break;
                    case TitlePageFieldType.DraftDate: dialogViewModel.DraftDate = val; break;
                    case TitlePageFieldType.Revision: dialogViewModel.Revision = val; break;
                    case TitlePageFieldType.Notes: dialogViewModel.Notes = val; break;
                }
            }
        }

        var dialog = new Views.TitlePageDialog
        {
            DataContext = dialogViewModel
        };

        var result = await dialog.ShowDialog<bool>(_window);
        if (result)
        {
            var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();
            int bodyStart = (_lastParsed != null && _lastParsed.TitlePage != null) ? _lastParsed.TitlePage.BodyStartLineIndex : 0;

            if (dialog.Deleted)
            {
                if (bodyStart > 0 && bodyStart <= lines.Count)
                {
                    lines.RemoveRange(0, bodyStart);
                }
            }
            else
            {
                var sb = new StringBuilder();
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Title)) sb.AppendLine($"Title: {dialogViewModel.Title}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Episode)) sb.AppendLine($"Episode: {dialogViewModel.Episode}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Author)) sb.AppendLine($"Author: {dialogViewModel.Author}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Credit)) sb.AppendLine($"Credit: {dialogViewModel.Credit}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Source)) sb.AppendLine($"Source: {dialogViewModel.Source}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.DraftDate)) sb.AppendLine($"Draft date: {dialogViewModel.DraftDate}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Revision)) sb.AppendLine($"Revision: {dialogViewModel.Revision}");
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Contact))
                {
                    sb.AppendLine("Contact:");
                    foreach (var l in dialogViewModel.Contact.Split('\n'))
                    {
                        sb.AppendLine($"    {l.Trim()}");
                    }
                }
                if (!string.IsNullOrWhiteSpace(dialogViewModel.Notes))
                {
                    sb.AppendLine("Notes:");
                    foreach (var l in dialogViewModel.Notes.Split('\n'))
                    {
                        sb.AppendLine($"    {l.Trim()}");
                    }
                }

                var newHeader = sb.ToString().TrimEnd();
                if (bodyStart > 0 && bodyStart <= lines.Count)
                {
                    lines.RemoveRange(0, bodyStart);
                }

                if (!string.IsNullOrEmpty(newHeader))
                {
                    lines.Insert(0, "");
                    lines.Insert(0, newHeader);
                }
            }

            _suppressDirtyTracking = true;
            try
            {
                EditorContent = string.Join("\n", lines);
                IsDirty = true;
            }
            finally
            {
                _suppressDirtyTracking = false;
            }
            RefreshParsedDocument();
        }
    }

    // Parameter is object for the same reason as OpenRecent: the menu style
    // also matches the parent "Export" MenuItem, whose DataContext is this VM.
    [RelayCommand]
    private async Task Export(object? parameter)
    {
        if (parameter is not IExporter exporter || _window == null) return;

        var file = await _window.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = $"Export to {exporter.DisplayName}",
            DefaultExtension = exporter.DefaultExtension,
            FileTypeChoices = new[]
            {
                new FilePickerFileType($"{exporter.DisplayName} Files") { Patterns = new[] { $"*{exporter.DefaultExtension}" } }
            }
        });

        if (file != null)
        {
            try
            {
                exporter.Export(_lastParsed, file.Path.LocalPath);
                StatusMessage = $"Exported to {file.Name}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export failed: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Error exporting: {ex.Message}");
            }
        }
    }

    // Scratchpad deletion command
    [RelayCommand]
    private void DeleteScratchpadCard()
    {
        if (SelectedScratchpadElement != null)
        {
            ScratchpadElements.Remove(SelectedScratchpadElement);
            OnPropertyChanged(nameof(HasScratchpadItems));
            OnPropertyChanged(nameof(ScratchpadEmptyMessage));
        }
    }

    // Parsing control
    private void OnContentChanged()
    {
        if (!_suppressDirtyTracking)
        {
            IsDirty = true;
        }
        ScheduleOutlineRefresh();
    }

    private void ScheduleOutlineRefresh()
    {
        _outlineRefreshVersion++;
        _outlineRefreshPending = true;
        _refreshTimer.Stop();
        _refreshTimer.Interval = TimeSpan.FromMilliseconds(250); // Throttle parses
        _refreshTimer.Start();
    }

    private void RefreshParsedDocument()
    {
        _refreshTimer.Stop();

        if (_outlineRefreshInProgress)
        {
            _outlineRefreshPending = true;
            return;
        }

        _outlineRefreshInProgress = true;
        _outlineRefreshPending = false;

        var version = _outlineRefreshVersion;
        var snapshotText = EditorContent ?? string.Empty;
        var parseStartedAt = Environment.TickCount64;

        try
        {
            if (CurrentMode == WriteMode.Markdown)
            {
                // Markdown documents get a heading-based outline instead of a
                // Fountain parse; the screenplay-only panels show placeholders.
                if (version == _outlineRefreshVersion)
                {
                    ApplyMarkdownDocument(snapshotText);
                }
            }
            else
            {
                // Parse the Fountain file content
                var parsed = _parser.Parse(snapshotText, _lineTypeOverrides);

                if (version == _outlineRefreshVersion)
                {
                    _lastParseDuration = TimeSpan.FromMilliseconds(Math.Max(0, Environment.TickCount64 - parseStartedAt));
                    ApplyParsedDocument(parsed);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error parsing document: {ex.Message}");
        }
        finally
        {
            _outlineRefreshInProgress = false;

            if (_outlineRefreshPending)
            {
                ScheduleOutlineRefresh();
            }
        }
    }

    private void ApplyParsedDocument(ParsedScreenplay parsed)
    {
        _lastParsed = parsed;

        var expandedOutline = CaptureExpandedIdentifiers(OutlineRoots);
        var expandedNotes = CaptureExpandedIdentifiers(NotesRoots);

        PopulateBodyText(parsed.Elements);
        UpdateUniqueScreenplayElements(parsed.Elements);

        var outlineRoots = BuildOutlineTree(parsed.Elements, expandedOutline, line => SelectedOutlineLineNumber = line);
        var noteRoots = BuildNotesTree(parsed.Elements, expandedNotes, line => SelectedOutlineLineNumber = line);

        OutlineRoots = new ObservableCollection<OutlineNodeViewModel>(outlineRoots);
        NotesRoots = new ObservableCollection<OutlineNodeViewModel>(noteRoots);

        // The tree was rebuilt with fresh node instances; re-apply the
        // caret-driven highlight so it lands on the new nodes.
        _currentOutlineNode = null;
        _currentNotesNode = null;
        UpdateCurrentOutlineNode(CurrentLineNumber);

        // Update flattened list of outline nodes
        var flatNodes = new List<OutlineNodeViewModel>();
        FlattenOutlineNodesRecursive(outlineRoots, flatNodes);
        OutlineNodes = new ObservableCollection<OutlineNodeViewModel>(flatNodes.OrderBy(n => n.LineNumber));

        UpdateBeatBoardCards(parsed.Elements);
        UpdatePreviewPages(parsed);

        OnPropertyChanged(nameof(HasOutlineItems));
        OnPropertyChanged(nameof(HasNoteItems));

        RefreshGoalState();
    }

    // Markdown mode: the Outline mirrors the document's #-heading structure and
    // the Fountain-only panels (Notes, Beat Board, Page Preview) are emptied so
    // their placeholders show.
    private void ApplyMarkdownDocument(string text)
    {
        var expandedOutline = CaptureExpandedIdentifiers(OutlineRoots);

        var roots = new List<OutlineNodeViewModel>();
        var stack = new Stack<OutlineNodeViewModel>();
        var lines = (text ?? string.Empty).Replace("\r\n", "\n").Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            int hashCount = 0;
            int j = 0;
            while (j < line.Length && line[j] == ' ' && j < 3) j++;
            while (j < line.Length && line[j] == '#') { hashCount++; j++; }
            if (hashCount is < 1 or > 6 || j >= line.Length || line[j] != ' ')
            {
                continue;
            }

            var title = line[(j + 1)..].Trim();
            if (title.Length == 0) continue;

            while (stack.Count > 0 && (stack.Peek().SectionLevel ?? 0) >= hashCount)
            {
                stack.Pop();
            }

            var node = new OutlineNodeViewModel(
                OutlineNodeKind.Section,
                title,
                lineNumber: i + 1,
                sectionLevel: hashCount,
                bodyText: null,
                navigateAction: line2 => SelectedOutlineLineNumber = line2,
                level: hashCount - 1,
                kindLabelOverride: $"H{hashCount}");

            if (HasExpandedOutlineKey(expandedOutline, OutlineNodeKind.Section, i + 1, hashCount, title))
            {
                node.IsExpanded = true;
            }

            if (stack.Count == 0)
            {
                roots.Add(node);
            }
            else
            {
                stack.Peek().Children.Add(node);
            }

            stack.Push(node);
        }

        OutlineRoots = new ObservableCollection<OutlineNodeViewModel>(roots);
        NotesRoots = new ObservableCollection<OutlineNodeViewModel>();

        _currentOutlineNode = null;
        _currentNotesNode = null;
        UpdateCurrentOutlineNode(CurrentLineNumber);

        var flatNodes = new List<OutlineNodeViewModel>();
        FlattenOutlineNodesRecursive(roots, flatNodes);
        OutlineNodes = new ObservableCollection<OutlineNodeViewModel>(flatNodes.OrderBy(n => n.LineNumber));

        BeatBoardCards.Clear();
        BeatBoardLanes.Clear();
        PreviewPages = new ObservableCollection<PreviewPageViewModel>();
        _lastParsed = new ParsedScreenplay(string.Empty, Array.Empty<ScreenplayElement>());

        OnPropertyChanged(nameof(HasOutlineItems));
        OnPropertyChanged(nameof(HasNoteItems));
        OnPropertyChanged(nameof(HasBeatBoardCards));

        RefreshGoalState();
    }

    private void UpdatePreviewPages(ParsedScreenplay screenplay)
    {
        var (titlePages, bodyPages) = ScreenplayLayoutBuilder.BuildPages(screenplay);
        var viewModels = new List<PreviewPageViewModel>();

        // Process title pages
        foreach (var page in titlePages)
        {
            var pageVm = new PreviewPageViewModel
            {
                PageNumberLabel = string.Empty
            };

            double y = ScreenplayLayoutBuilder.MarginTop;
            foreach (var line in page.Lines)
            {
                if (!line.IsBlank)
                {
                    var resolvedX = ComputeWpfX(line);
                    pageVm.Lines.Add(new PreviewLineViewModel
                    {
                        Text = line.Text,
                        X = resolvedX,
                        Y = y,
                        IsBold = line.Style == LayoutTextStyle.LeftBold || line.Style == LayoutTextStyle.CenterWithinBodyBold
                    });
                }
                y += ScreenplayLayoutBuilder.LineHeight;
            }

            viewModels.Add(pageVm);
        }

        // Process body pages
        foreach (var page in bodyPages)
        {
            var pageVm = new PreviewPageViewModel
            {
                PageNumberLabel = page.PageNumber > 1 ? $"{page.PageNumber}." : string.Empty
            };

            if (page.PageNumber > 1)
            {
                var pageNumText = $"{page.PageNumber}.";
                var pageNumLine = new LayoutLine(pageNumText, LayoutTextStyle.RightWithinBody, 0.0);
                var resolvedX = ComputeWpfX(pageNumLine);
                pageVm.Lines.Add(new PreviewLineViewModel
                {
                    Text = pageNumText,
                    X = resolvedX,
                    Y = ScreenplayLayoutBuilder.PageNumberTopMargin,
                    IsBold = false
                });
            }

            double y = ScreenplayLayoutBuilder.MarginTop;
            foreach (var line in page.Lines)
            {
                if (!line.IsBlank)
                {
                    var resolvedX = ComputeWpfX(line);
                    pageVm.Lines.Add(new PreviewLineViewModel
                    {
                        Text = line.Text,
                        X = resolvedX,
                        Y = y,
                        IsBold = line.Style == LayoutTextStyle.LeftBold || line.Style == LayoutTextStyle.CenterWithinBodyBold
                    });
                }
                y += ScreenplayLayoutBuilder.LineHeight;
            }

            viewModels.Add(pageVm);
        }

        PreviewPages = new ObservableCollection<PreviewPageViewModel>(viewModels);
    }

    [RelayCommand]
    private void ToggleWriteMode()
    {
        CurrentMode = CurrentMode == WriteMode.Screenplay ? WriteMode.Markdown : WriteMode.Screenplay;
        StatusMessage = CurrentMode == WriteMode.Screenplay
            ? "Switched to Screenplay mode (Fountain)"
            : "Switched to Markdown mode";
        RefreshParsedDocument();
    }

    [RelayCommand]
    private void ExpandAllOutlineNodes() => SetOutlineExpansionState(OutlineRoots, isExpanded: true);

    [RelayCommand]
    private void CollapseAllOutlineNodes() => SetOutlineExpansionState(OutlineRoots, isExpanded: false);

    [RelayCommand]
    private void ExpandAllNotesNodes() => SetOutlineExpansionState(NotesRoots, isExpanded: true);

    [RelayCommand]
    private void CollapseAllNotesNodes() => SetOutlineExpansionState(NotesRoots, isExpanded: false);

    private static void SetOutlineExpansionState(IEnumerable<OutlineNodeViewModel> nodes, bool isExpanded)
    {
        foreach (var node in nodes)
        {
            node.IsExpanded = isExpanded;

            if (node.Children.Count > 0)
            {
                SetOutlineExpansionState(node.Children, isExpanded);
            }
        }
    }

    public void UpdateCaretStatus(int caretIndex)
    {
        var text = EditorContent ?? string.Empty;
        if (caretIndex < 0 || caretIndex > text.Length) return;

        int line = 1;
        for (int i = 0; i < caretIndex && i < text.Length; i++)
        {
            if (text[i] == '\n') line++;
        }
        CurrentLineNumber = line;
        UpdateCurrentOutlineNode(line);

        if (_lastParsed != null && IsScreenplayMode)
        {
            var element = _lastParsed.Elements.FirstOrDefault(e => e.LineIndex <= line - 1 && line - 1 <= e.EndLineIndex);
            if (element != null)
            {
                CurrentElementType = element.Type.ToString();
                CurrentElementText = element.Text;
            }
            else
            {
                CurrentElementType = "Action";
                CurrentElementText = string.Empty;
            }
        }
        else
        {
            CurrentElementType = "Action";
            CurrentElementText = string.Empty;
        }
    }

    private OutlineNodeViewModel? _currentOutlineNode;
    private OutlineNodeViewModel? _currentNotesNode;

    // Highlights the deepest outline/notes card whose element contains the
    // given 1-based caret line, so the workspace indicator tracks the caret.
    public void UpdateCurrentOutlineNode(int line)
    {
        _currentOutlineNode = SetCurrentNode(OutlineRoots, line, _currentOutlineNode);
        _currentNotesNode = SetCurrentNode(NotesRoots, line, _currentNotesNode);
    }

    private OutlineNodeViewModel? SetCurrentNode(
        IEnumerable<OutlineNodeViewModel> roots,
        int line,
        OutlineNodeViewModel? previous)
    {
        var flat = new List<OutlineNodeViewModel>();
        FlattenOutlineNodesRecursive(roots, flat);

        // The element the caret is in is the node with the greatest line number
        // that still starts at or before the caret — i.e. the most specific
        // (deepest, latest-starting) element containing the caret.
        OutlineNodeViewModel? match = null;
        foreach (var node in flat)
        {
            if (node.LineNumber <= line && (match == null || node.LineNumber > match.LineNumber))
            {
                match = node;
            }
        }

        if (!ReferenceEquals(previous, match))
        {
            if (previous != null) previous.IsCurrent = false;
            if (match != null) match.IsCurrent = true;
        }

        return match;
    }

    private static double ComputeWpfX(LayoutLine line)
    {
        var usableWidth = ScreenplayLayoutBuilder.PageWidth - ScreenplayLayoutBuilder.MarginLeft - ScreenplayLayoutBuilder.MarginRight;
        var lineWidth = Math.Min(usableWidth, Math.Max(1, line.Text.Length) * ScreenplayLayoutBuilder.CharWidth);

        return line.Style switch
        {
            LayoutTextStyle.CenterWithinBody or LayoutTextStyle.CenterWithinBodyBold => ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, (usableWidth - lineWidth) / 2),
            LayoutTextStyle.RightWithinBody => ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, usableWidth - lineWidth),
            _ => line.X
        };
    }

    private void UpdateBeatBoardCards(IReadOnlyList<ScreenplayElement> elements)
    {
        var activeElements = elements.Where(e => !e.IsSuppressed &&
            (e.Type is ScreenplayElementType.Section or ScreenplayElementType.SceneHeading or ScreenplayElementType.Note)).ToList();

        var existingById = new Dictionary<Guid, BeatBoardCardViewModel>();
        foreach (var c in BeatBoardCards)
        {
            if (!existingById.ContainsKey(c.Id))
            {
                existingById.Add(c.Id, c);
            }
        }

        var newCards = new List<BeatBoardCardViewModel>();
        foreach (var element in activeElements)
        {
            var id = element.Id;
            var heading = element.Heading;
            var desc = element.BoardDescription;
            var typeStr = element.Type switch
            {
                ScreenplayElementType.Section => element.Level == 0 ? "Act" : element.Level == 1 ? "Sequence" : "Section",
                ScreenplayElementType.SceneHeading => "Scene",
                ScreenplayElementType.Note => "Note",
                _ => "Card"
            };

            if (existingById.TryGetValue(id, out var existingCard))
            {
                existingCard.Heading = heading;
                existingCard.Description = desc;
                existingCard.LineNumber = element.LineNumber;
                existingCard.Type = typeStr;
                existingCard.Level = element.Level;
                newCards.Add(existingCard);
            }
            else
            {
                newCards.Add(new BeatBoardCardViewModel
                {
                    Id = id,
                    Heading = heading,
                    Description = desc,
                    LineNumber = element.LineNumber,
                    Type = typeStr,
                    Level = element.Level
                });
            }
        }

        BeatBoardCards.Clear();
        foreach (var card in newCards)
        {
            BeatBoardCards.Add(card);
        }

        RebuildBeatBoardLanes(newCards);
        OnPropertyChanged(nameof(HasBeatBoardCards));
    }

    // Groups the flat, document-ordered card list into the story hierarchy the
    // board renders: an Act opens a new lane, a Sequence opens a new group inside
    // the current lane, and every other card (Scene / Note / deep Section) lands in
    // the current group. Cards that appear before any Act or Sequence fall into
    // implicit (header-less) lanes/groups so nothing is hidden from the board.
    private void RebuildBeatBoardLanes(IReadOnlyList<BeatBoardCardViewModel> cards)
    {
        // Lanes/groups are rebuilt from scratch on every parse; carry the user's
        // collapsed acts and sequences across the rebuild by header-card id.
        var collapsedActIds = BeatBoardLanes
            .Where(lane => lane.ActCard != null && !lane.IsExpanded)
            .Select(lane => lane.ActCard!.Id)
            .ToHashSet();
        var collapsedSequenceIds = BeatBoardLanes
            .SelectMany(lane => lane.Groups)
            .Where(group => group.SequenceCard != null && !group.IsExpanded)
            .Select(group => group.SequenceCard!.Id)
            .ToHashSet();

        BeatBoardLanes.Clear();

        BeatBoardLaneViewModel? currentLane = null;
        BeatBoardGroupViewModel? currentGroup = null;

        foreach (var card in cards)
        {
            if (card.Type == "Act")
            {
                currentLane = new BeatBoardLaneViewModel
                {
                    ActCard = card,
                    IsExpanded = !collapsedActIds.Contains(card.Id)
                };
                currentGroup = null;
                BeatBoardLanes.Add(currentLane);
            }
            else if (card.Type == "Sequence")
            {
                if (currentLane == null)
                {
                    currentLane = new BeatBoardLaneViewModel();
                    BeatBoardLanes.Add(currentLane);
                }

                currentGroup = new BeatBoardGroupViewModel
                {
                    SequenceCard = card,
                    IsExpanded = !collapsedSequenceIds.Contains(card.Id)
                };
                currentLane.Groups.Add(currentGroup);
            }
            else
            {
                if (currentLane == null)
                {
                    currentLane = new BeatBoardLaneViewModel();
                    BeatBoardLanes.Add(currentLane);
                }

                if (currentGroup == null)
                {
                    currentGroup = new BeatBoardGroupViewModel();
                    currentLane.Groups.Add(currentGroup);
                }

                currentGroup.Cards.Add(card);
            }
        }
    }

    // Goal operations
    private void RefreshGoalState()
    {
        CaptureSessionGoalBaselineIfNeeded();

        var currentWordCount = _goalProgressCalculator.CalculateWordCount(EditorContent);
        var currentPageCount = _pageEstimator.EstimatePageCount(_lastParsed);

        RefreshOverallGoal(currentWordCount, currentPageCount);
        RefreshSessionGoal(currentWordCount, currentPageCount);

        var pageCount = Math.Max(1, currentPageCount);
        PageCountStatusText = $"~{pageCount:n0} {FormatCountLabel(pageCount, "page", "pages")}";

        GoalProgressSummaryText = BuildGoalProgressSummaryText(currentWordCount, currentPageCount);

        OnPropertyChanged(nameof(IsTimerGoal));
        OnPropertyChanged(nameof(GoalTimerState));
        OnPropertyChanged(nameof(GoalTimerPrimaryButtonText));
        OnPropertyChanged(nameof(GoalTimerSecondaryButtonText));
    }

    private string BuildGoalProgressSummaryText(int currentWordCount, int currentPageCount)
    {
        var summaryParts = new List<string>(3)
        {
            $"Words {currentWordCount:n0}"
        };

        var overallRemainingText = BuildGoalRemainingText(
            SelectedGoalType,
            currentWordCount,
            currentPageCount,
            _wordCountGoalTargetValue,
            _pageCountGoalTargetValue,
            GoalTimerRemainingText);

        if (!string.IsNullOrWhiteSpace(overallRemainingText))
        {
            summaryParts.Add($"Overall {overallRemainingText}");
        }

        var sessionRemainingText = BuildGoalRemainingText(
            SessionSelectedGoalType,
            Math.Max(0, currentWordCount - _sessionWordCountBaseline),
            Math.Max(0, currentPageCount - _sessionPageCountBaseline),
            _sessionWordCountTargetValue,
            _sessionPageCountTargetValue,
            GoalTimerRemainingText);

        if (!string.IsNullOrWhiteSpace(sessionRemainingText))
        {
            summaryParts.Add($"Session {sessionRemainingText}");
        }

        return string.Join(" | ", summaryParts);
    }

    private void RefreshOverallGoal(int currentWordCount, int currentPageCount)
    {
        switch (SelectedGoalType)
        {
            case GoalType.WordCount:
                RefreshWordCountGoal(currentWordCount);
                break;
            case GoalType.PageCount:
                RefreshPageCountGoal(currentPageCount);
                break;
        }
    }

    private void RefreshSessionGoal(int currentWordCount, int currentPageCount)
    {
        switch (SessionSelectedGoalType)
        {
            case GoalType.WordCount:
                RefreshSessionWordCountGoal(currentWordCount);
                break;
            case GoalType.PageCount:
                RefreshSessionPageCountGoal(currentPageCount);
                break;
            case GoalType.Timer:
                RefreshSessionTimerGoal();
                break;
        }
    }

    private void RefreshWordCountGoal(int currentValue)
    {
        var targetValue = _wordCountGoalTargetValue;
        var completed = targetValue <= 0 || currentValue >= targetValue;
        var unitLabel = FormatCountLabel(targetValue, "Word", "Words");

        GoalTargetUnitLabel = unitLabel;
        GoalCurrentDisplayText = $"{currentValue:n0} {FormatCountLabel(currentValue, "word", "words").ToLowerInvariant()}";
        GoalTargetDisplayText = $"{targetValue:n0} {FormatCountLabel(targetValue, "word", "words").ToLowerInvariant()}";
        GoalStateText = completed ? "Completed" : "In progress";
        GoalProgressPercent = CalculateProgressPercent(currentValue, targetValue);
    }

    private void RefreshPageCountGoal(int currentValue)
    {
        var targetValue = _pageCountGoalTargetValue;
        var completed = targetValue <= 0 || currentValue >= targetValue;
        var unitLabel = FormatCountLabel(targetValue, "Page", "Pages");

        GoalTargetUnitLabel = unitLabel;
        GoalCurrentDisplayText = $"{currentValue:n0} {FormatCountLabel(currentValue, "page", "pages").ToLowerInvariant()}";
        GoalTargetDisplayText = $"{targetValue:n0} {FormatCountLabel(targetValue, "page", "pages").ToLowerInvariant()}";
        GoalStateText = completed ? "Completed" : "In progress";
        GoalProgressPercent = CalculateProgressPercent(currentValue, targetValue);
    }

    private void RefreshSessionWordCountGoal(int currentWordCount)
    {
        var targetValue = _sessionWordCountTargetValue;
        var currentValue = Math.Max(0, currentWordCount - _sessionWordCountBaseline);
        var completed = targetValue <= 0 || currentValue >= targetValue;
        var unitLabel = FormatCountLabel(targetValue, "Word", "Words");

        SessionGoalTargetUnitLabel = unitLabel;
        SessionGoalCurrentDisplayText = $"{currentValue:n0} {FormatCountLabel(currentValue, "word", "words").ToLowerInvariant()}";
        SessionGoalTargetDisplayText = $"{targetValue:n0} {FormatCountLabel(targetValue, "word", "words").ToLowerInvariant()}";
        SessionGoalStateText = completed ? "Completed" : "In progress";
        SessionGoalProgressPercent = CalculateProgressPercent(currentValue, targetValue);
    }

    private void RefreshSessionPageCountGoal(int currentPageCount)
    {
        var targetValue = _sessionPageCountTargetValue;
        var currentValue = Math.Max(0, currentPageCount - _sessionPageCountBaseline);
        var completed = targetValue <= 0 || currentValue >= targetValue;
        var unitLabel = FormatCountLabel(targetValue, "Page", "Pages");

        SessionGoalTargetUnitLabel = unitLabel;
        SessionGoalCurrentDisplayText = $"{currentValue:n0} {FormatCountLabel(currentValue, "page", "pages").ToLowerInvariant()}";
        SessionGoalTargetDisplayText = $"{targetValue:n0} {FormatCountLabel(targetValue, "page", "pages").ToLowerInvariant()}";
        SessionGoalStateText = completed ? "Completed" : "In progress";
        SessionGoalProgressPercent = CalculateProgressPercent(currentValue, targetValue);
    }

    private void RefreshSessionTimerGoal()
    {
        EnsureGoalTimerRuntime();

        if (_goalTimerRuntime is null)
        {
            SessionGoalTargetUnitLabel = FormatCountLabel(_timerGoalTargetMinutes, "Minute", "Minutes");
            SessionGoalCurrentDisplayText = "00:00";
            SessionGoalTargetDisplayText = FormatDuration(TimeSpan.FromMinutes(_timerGoalTargetMinutes));
            SessionGoalStateText = "Idle";
            SessionGoalProgressPercent = 0;
            GoalTimerElapsedText = "00:00";
            GoalTimerRemainingText = FormatDuration(TimeSpan.FromMinutes(_timerGoalTargetMinutes));
            return;
        }

        var targetDuration = _goalTimerRuntime.TargetDuration;
        var elapsed = _goalTimerRuntime.ElapsedTime;
        var remaining = _goalTimerRuntime.RemainingTime;
        var completed = _goalTimerRuntime.IsCompleted;

        SessionGoalTargetUnitLabel = FormatCountLabel(_timerGoalTargetMinutes, "Minute", "Minutes");
        SessionGoalCurrentDisplayText = FormatDuration(elapsed);
        SessionGoalTargetDisplayText = FormatDuration(targetDuration);
        SessionGoalStateText = _goalTimerRuntime.State.ToString();
        SessionGoalProgressPercent = CalculateProgressPercent(elapsed.TotalSeconds, targetDuration.TotalSeconds);
        GoalTimerElapsedText = FormatDuration(elapsed);
        GoalTimerRemainingText = FormatDuration(remaining);

        if (completed)
        {
            SessionGoalStateText = "Completed";
        }
    }

    private void CaptureSessionGoalBaselineIfNeeded()
    {
        if (!_sessionGoalBaselineNeedsCapture) return;
        CaptureSessionGoalBaseline();
    }

    private void CaptureSessionGoalBaseline()
    {
        var parsed = _parser.Parse(EditorContent ?? string.Empty, null);
        _lastParsed = parsed;
        _sessionWordCountBaseline = _goalProgressCalculator.CalculateWordCount(EditorContent);
        _sessionPageCountBaseline = _pageEstimator.EstimatePageCount(parsed);
        _sessionGoalBaselineNeedsCapture = false;
    }

    public void ResetSessionGoal()
    {
        CaptureSessionGoalBaseline();
        RefreshGoalState();
    }

    private void EnsureGoalTimerRuntime()
    {
        var targetDuration = TimeSpan.FromMinutes(Math.Max(0, _timerGoalTargetMinutes));

        if (_goalTimerRuntime is not null && _goalTimerRuntime.TargetDuration == targetDuration) return;

        if (_goalTimerRuntime is not null)
        {
            _goalTimerRuntime.PropertyChanged -= GoalTimerRuntime_PropertyChanged;
            _goalTimerRuntime.Dispose();
        }

        _goalTimerRuntime = new GoalTimerRuntime(targetDuration);
        _goalTimerRuntime.PropertyChanged += GoalTimerRuntime_PropertyChanged;
    }

    private void RebuildGoalTimerRuntime()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        EnsureGoalTimerRuntime();
        RefreshGoalState();
    }

    public void StartGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        EnsureGoalTimerRuntime();
        if (_goalTimerRuntime?.State == TimerGoalState.Completed)
        {
            _goalTimerRuntime.Reset();
        }
        _goalTimerRuntime?.Start();
        RefreshGoalState();
    }

    public void PauseGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        _goalTimerRuntime?.Pause();
        RefreshGoalState();
    }

    public void ResumeGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        EnsureGoalTimerRuntime();
        _goalTimerRuntime?.Resume();
        RefreshGoalState();
    }

    public void StopGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        _goalTimerRuntime?.Stop();
        RefreshGoalState();
    }

    public void ResetGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer) return;
        EnsureGoalTimerRuntime();
        _goalTimerRuntime?.Reset();
        RefreshGoalState();
    }

    private void GoalTimerRuntime_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!ReferenceEquals(sender, _goalTimerRuntime)) return;
        if (e.PropertyName is not nameof(GoalTimerRuntime.Snapshot)
            and not nameof(GoalTimerRuntime.State)
            and not nameof(GoalTimerRuntime.IsCompleted))
        {
            return;
        }

        RefreshGoalState();
    }

    private static double CalculateProgressPercent(double currentValue, double targetValue)
    {
        if (targetValue <= 0) return 100;
        return Math.Min(100, Math.Max(0, (currentValue / targetValue) * 100));
    }

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration < TimeSpan.Zero) duration = TimeSpan.Zero;
        if (duration.TotalHours >= 1) return duration.ToString(@"h\:mm\:ss");
        return duration.ToString(@"mm\:ss");
    }

    private static string FormatCountLabel(int value, string singular, string plural)
    {
        return value == 1 ? singular : plural;
    }

    private static string BuildGoalRemainingText(
        GoalType goalType,
        int currentWordCount,
        int currentPageCount,
        int wordTargetValue,
        int pageTargetValue,
        string? timerRemainingText)
    {
        return goalType switch
        {
            GoalType.WordCount => BuildCountRemainingText(currentWordCount, wordTargetValue, "word", "words"),
            GoalType.PageCount => BuildCountRemainingText(currentPageCount, pageTargetValue, "page", "pages"),
            GoalType.Timer => string.IsNullOrWhiteSpace(timerRemainingText) ? "00:00 left" : $"{timerRemainingText} left",
            _ => string.Empty
        };
    }

    private static string BuildCountRemainingText(int currentValue, int targetValue, string singular, string plural)
    {
        var remainingValue = Math.Max(0, targetValue - currentValue);
        return $"{remainingValue:n0} {FormatCountLabel(remainingValue, singular, plural).ToLowerInvariant()} left";
    }

    // Outline node generation
    private static ISet<string> CaptureExpandedIdentifiers(IEnumerable<OutlineNodeViewModel> nodes)
    {
        var expanded = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in nodes)
        {
            if (node.IsExpanded)
            {
                AddOutlineExpansionKeys(expanded, node.Kind, node.LineNumber, node.SectionLevel, node.Text);
            }
            if (node.Children.Count > 0)
            {
                CaptureExpandedIdentifiersRecursive(node.Children, expanded);
            }
        }
        return expanded;
    }

    private static void CaptureExpandedIdentifiersRecursive(IEnumerable<OutlineNodeViewModel> nodes, ISet<string> expanded)
    {
        foreach (var node in nodes)
        {
            if (node.IsExpanded)
            {
                AddOutlineExpansionKeys(expanded, node.Kind, node.LineNumber, node.SectionLevel, node.Text);
            }
            if (node.Children.Count > 0)
            {
                CaptureExpandedIdentifiersRecursive(node.Children, expanded);
            }
        }
    }

    private static void AddOutlineExpansionKeys(ISet<string> expanded, OutlineNodeKind kind, int lineNumber, int? sectionLevel, string text)
    {
        expanded.Add(BuildOutlineExpansionKey(kind, lineNumber, sectionLevel));
        expanded.Add(BuildOutlineLegacyExpansionKey(kind, text));
    }

    private static bool HasExpandedOutlineKey(ISet<string>? expandedKeys, OutlineNodeKind kind, int lineNumber, int? sectionLevel, string text)
    {
        if (expandedKeys is null) return false;
        return expandedKeys.Contains(BuildOutlineExpansionKey(kind, lineNumber, sectionLevel)) ||
               expandedKeys.Contains(BuildOutlineLegacyExpansionKey(kind, text));
    }

    private static string BuildOutlineExpansionKey(OutlineNodeKind kind, int lineNumber, int? sectionLevel)
    {
        return $"{kind}_{lineNumber}_{sectionLevel ?? -1}";
    }

    private static string BuildOutlineLegacyExpansionKey(OutlineNodeKind kind, string text)
    {
        return $"{kind}_{text}";
    }

    private void FlattenOutlineNodesRecursive(IEnumerable<OutlineNodeViewModel> nodes, List<OutlineNodeViewModel> flatList)
    {
        foreach (var node in nodes)
        {
            flatList.Add(node);
            if (node.Children.Count > 0)
            {
                FlattenOutlineNodesRecursive(node.Children, flatList);
            }
        }
    }

    private static IReadOnlyList<OutlineNodeViewModel> BuildOutlineTree(
        IReadOnlyList<ScreenplayElement> elements,
        ISet<string>? expandedKeys = null,
        Action<int>? navigateAction = null)
    {
        var roots = new List<OutlineNodeViewModel>();
        var sectionStack = new Stack<OutlineNodeViewModel>();

        foreach (var element in elements)
        {
            switch (element)
            {
                case SectionElement section:
                    {
                        while (sectionStack.Count > 0 && (sectionStack.Peek().SectionLevel ?? 0) >= section.SectionDepth)
                        {
                            sectionStack.Pop();
                        }

                        var sectionNode = new OutlineNodeViewModel(
                            OutlineNodeKind.Section,
                            section.Text,
                            section.StartLine,
                            section.SectionDepth,
                            section.BodyText,
                            navigateAction,
                            level: section.SectionDepth - 1);

                        if (HasExpandedOutlineKey(expandedKeys, OutlineNodeKind.Section, section.StartLine, section.SectionDepth, section.Text))
                        {
                            sectionNode.IsExpanded = true;
                        }

                        if (sectionStack.Count == 0)
                        {
                            roots.Add(sectionNode);
                        }
                        else
                        {
                            sectionStack.Peek().Children.Add(sectionNode);
                        }

                        sectionStack.Push(sectionNode);
                        break;
                    }

                case SceneHeadingElement sceneHeading:
                    {
                        var sceneNode = new OutlineNodeViewModel(
                            OutlineNodeKind.SceneHeading,
                            sceneHeading.Text,
                            sceneHeading.StartLine,
                            sectionLevel: null,
                            sceneHeading.BodyText,
                            navigateAction,
                            level: sectionStack.Count);

                        if (HasExpandedOutlineKey(expandedKeys, OutlineNodeKind.SceneHeading, sceneHeading.StartLine, sectionLevel: null, sceneHeading.Text))
                        {
                            sceneNode.IsExpanded = true;
                        }

                        if (sectionStack.Count == 0)
                        {
                            roots.Add(sceneNode);
                        }
                        else
                        {
                            sectionStack.Peek().Children.Add(sceneNode);
                        }

                        break;
                    }
            }
        }

        return roots;
    }

    private static IReadOnlyList<OutlineNodeViewModel> BuildNotesTree(
        IReadOnlyList<ScreenplayElement> elements,
        ISet<string>? expandedKeys = null,
        Action<int>? navigateAction = null)
    {
        var roots = new List<OutlineNodeViewModel>();

        foreach (var element in elements)
        {
            if (element is not NoteElement note) continue;
            if (string.IsNullOrWhiteSpace(note.Text)) continue;

            var noteNode = new OutlineNodeViewModel(
                OutlineNodeKind.Note,
                note.Text,
                note.StartLine,
                sectionLevel: null,
                note.BodyText,
                navigateAction,
                level: 0);

            if (HasExpandedOutlineKey(expandedKeys, OutlineNodeKind.Note, note.StartLine, sectionLevel: null, note.Text))
            {
                noteNode.IsExpanded = true;
            }
            roots.Add(noteNode);
        }

        return roots;
    }

    private static void PopulateBodyText(IEnumerable<ScreenplayElement> elements)
    {
        var elementList = elements.ToList();
        for (int i = 0; i < elementList.Count; i++)
        {
            var element = elementList[i];

            if (element.Type is ScreenplayElementType.Section or ScreenplayElementType.SceneHeading or ScreenplayElementType.Note)
            {
                var bodyLines = new List<string>();
                for (int j = i + 1; j < elementList.Count; j++)
                {
                    var next = elementList[j];
                    if (next.Type is ScreenplayElementType.Synopsis)
                    {
                        var lineText = StripGuidComment(next.Text);
                        bodyLines.Add(lineText);
                        next.IsSuppressed = true;
                    }
                    else
                    {
                        break;
                    }
                }

                element.BodyText = bodyLines.Count > 0 ? string.Join("\n", bodyLines) : null;
            }
        }
    }

    private static string StripGuidComment(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        var idMatch = System.Text.RegularExpressions.Regex.Match(text, @"\s*\[\[id:[a-f\d\-]+\]\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (idMatch.Success)
        {
            return text.Replace(idMatch.Value, "").Trim();
        }

        var bareMatch = System.Text.RegularExpressions.Regex.Match(text, @"\s*id:[a-f\d\-]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (bareMatch.Success)
        {
            return text.Replace(bareMatch.Value, "").Trim();
        }

        return text;
    }

    public ScreenplayElementType GetLatestEffectiveLineType(int lineNumber, string lineText)
    {
        var screenplayLineTypes = new Dictionary<int, ScreenplayElementType>();
        if (_lastParsed != null && _lastParsed.Elements != null)
        {
            foreach (var element in _lastParsed.Elements)
            {
                for (int l = element.StartLine; l <= element.EndLine; l++)
                {
                    screenplayLineTypes[l] = element.Type;
                }
            }
        }

        if (screenplayLineTypes.TryGetValue(lineNumber, out var mappedType))
        {
            return mappedType;
        }

        var trimmed = (lineText ?? string.Empty).Trim();
        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        if (Passage.Core.TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()))
        {
            return ScreenplayElementType.SceneHeading;
        }

        if (trimmed.StartsWith("(") && trimmed.EndsWith(")", StringComparison.Ordinal))
        {
            return ScreenplayElementType.Parenthetical;
        }

        return Passage.Core.TextAnalysis.IsLiveCharacterCueCandidate(lineText.AsSpan(), 45, 6)
            ? ScreenplayElementType.Character
            : ScreenplayElementType.Action;
    }

    private void UpdateUniqueScreenplayElements(IReadOnlyList<ScreenplayElement> elements)
    {
        _uniqueSceneHeadings.Clear();
        _uniqueCharacterNames.Clear();

        foreach (var element in elements)
        {
            switch (element.Type)
            {
                case ScreenplayElementType.SceneHeading:
                    var heading = element.Text.Trim();
                    if (!string.IsNullOrWhiteSpace(heading))
                    {
                        _uniqueSceneHeadings.Add(heading.ToUpperInvariant());
                    }
                    break;
                case ScreenplayElementType.Character:
                    if (element is CharacterElement character)
                    {
                        var name = character.CharacterName.Trim();
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            _uniqueCharacterNames.Add(name.ToUpperInvariant());
                        }
                    }
                    break;
                case ScreenplayElementType.Dialogue:
                    if (element is DialogueElement dialogue)
                    {
                        var name = dialogue.CharacterName.Trim();
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            _uniqueCharacterNames.Add(name.ToUpperInvariant());
                        }
                    }
                    break;
            }
        }
    }

    public void UpdateSuggestions(string prefix, string elementTypeName)
    {
        AutoCompleteSuggestions.Clear();
        if (string.IsNullOrWhiteSpace(prefix))
        {
            IsAutoCompleteOpen = false;
            return;
        }

        var normalizedPrefix = prefix.Trim().ToUpperInvariant();
        IEnumerable<string> source = Array.Empty<string>();

        if (elementTypeName == "SceneHeading")
        {
            source = _uniqueSceneHeadings;
        }
        else if (elementTypeName == "Character")
        {
            source = _uniqueCharacterNames;
        }

        var matches = source
            .Where(s => s.StartsWith(normalizedPrefix) && !string.Equals(s, normalizedPrefix, StringComparison.OrdinalIgnoreCase))
            .OrderBy(s => s)
            .Take(10)
            .ToList();

        if (matches.Count > 0)
        {
            foreach (var match in matches)
            {
                AutoCompleteSuggestions.Add(match);
            }
            SelectedSuggestionIndex = 0;
            IsAutoCompleteOpen = true;
        }
        else
        {
            IsAutoCompleteOpen = false;
        }
    }

    public void StartRecoveryAutosave()
    {
        if (!IsDirty)
        {
            StopRecoveryAutosave();
            return;
        }

        if (!_recoveryTimer.IsEnabled)
        {
            _recoveryTimer.Start();
        }
    }

    public void StopRecoveryAutosave()
    {
        _recoveryTimer.Stop();
    }

    public void SaveRecoverySnapshot()
    {
        if (!IsDirty)
        {
            StopRecoveryAutosave();
            RecoveryStorage.ClearRecoveryFile();
            return;
        }

        RecoveryStorage.SaveRecoveryFile(new RecoveryDocument
        {
            Text = EditorContent ?? string.Empty,
            FilePath = _currentFilePath,
            SavedAtUtc = DateTimeOffset.UtcNow,
            GoalConfiguration = GoalConfiguration,
            SessionGoalConfiguration = SessionGoalConfiguration,
            EditorZoomPercent = EditorZoomScale * 100.0
        });
    }

    public void LoadRecoveryDocument(RecoveryDocument doc)
    {
        _suppressDirtyTracking = true;
        try
        {
            _currentFilePath = doc.FilePath ?? string.Empty;
            EditorContent = doc.Text;
            EditorZoomScale = doc.EditorZoomPercent / 100.0;

            SelectedGoalType = doc.GoalConfiguration.SelectedGoalType;
            _wordCountGoalTargetValue = doc.GoalConfiguration.WordCountTargetValue;
            _pageCountGoalTargetValue = doc.GoalConfiguration.PageCountTargetValue;
            _timerGoalTargetMinutes = doc.GoalConfiguration.TimerTargetMinutes;

            SessionSelectedGoalType = doc.SessionGoalConfiguration.SelectedGoalType;
            _sessionWordCountTargetValue = doc.SessionGoalConfiguration.WordCountTargetValue;
            _sessionPageCountTargetValue = doc.SessionGoalConfiguration.PageCountTargetValue;

            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }
    }

    public void LoadSessionState(SessionState state)
    {
        if (state.RecentFiles != null)
        {
            RecentFiles.Clear();
            foreach (var path in state.RecentFiles.Take(MaxRecentFiles))
            {
                RecentFiles.Add(path);
            }
            OnPropertyChanged(nameof(HasRecentFiles));
        }

        if (state.Documents == null || state.Documents.Count == 0) return;
        var doc = state.Documents[0];

        _suppressDirtyTracking = true;
        try
        {
            _currentFilePath = doc.FilePath ?? string.Empty;

            if (!doc.IsDirty && !string.IsNullOrWhiteSpace(doc.FilePath) && File.Exists(doc.FilePath))
            {
                EditorContent = File.ReadAllText(doc.FilePath);
            }
            else
            {
                EditorContent = doc.Text;
            }

            // Restore the write mode from the file extension, matching Open().
            if (!string.IsNullOrEmpty(_currentFilePath))
            {
                CurrentMode = string.Equals(Path.GetExtension(_currentFilePath), ".md", StringComparison.OrdinalIgnoreCase)
                    ? WriteMode.Markdown
                    : WriteMode.Screenplay;
            }

            EditorZoomScale = doc.EditorZoomPercent / 100.0;

            SelectedGoalType = doc.GoalConfiguration.SelectedGoalType;
            _wordCountGoalTargetValue = doc.GoalConfiguration.WordCountTargetValue;
            _pageCountGoalTargetValue = doc.GoalConfiguration.PageCountTargetValue;
            _timerGoalTargetMinutes = doc.GoalConfiguration.TimerTargetMinutes;

            SessionSelectedGoalType = doc.SessionGoalConfiguration.SelectedGoalType;
            _sessionWordCountTargetValue = doc.SessionGoalConfiguration.WordCountTargetValue;
            _sessionPageCountTargetValue = doc.SessionGoalConfiguration.PageCountTargetValue;

            IsDirty = doc.IsDirty;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }
    }

    public void SaveSessionNow()
    {
        var docState = new SessionDocumentState
        {
            FilePath = _currentFilePath,
            Text = EditorContent ?? string.Empty,
            IsDirty = IsDirty,
            GoalConfiguration = GoalConfiguration,
            SessionGoalConfiguration = SessionGoalConfiguration,
            EditorZoomPercent = EditorZoomScale * 100.0
        };

        double? width = null;
        double? height = null;
        int? x = null;
        int? y = null;
        string? windowState = null;

        if (_window != null)
        {
            width = _window.Width;
            height = _window.Height;
            x = _window.Position.X;
            y = _window.Position.Y;
            windowState = _window.WindowState.ToString();
        }

        var state = new SessionState(new[] { docState }, 0, width, height, x, y, windowState, RecentFiles.ToList());
        SessionStorage.SaveSession(state);
    }

    public (int startIdx, int endIdx) GetOutlineNodeLineRange(OutlineNodeViewModel node)
    {
        if (node.Kind == OutlineNodeKind.Note)
        {
            if (_lastParsed != null)
            {
                var element = _lastParsed.Elements.FirstOrDefault(e => e.StartLine == node.LineNumber && e.Type == ScreenplayElementType.Note);
                if (element != null)
                {
                    return (element.LineIndex, element.EndLineIndex);
                }
            }
            return (node.LineNumber - 1, node.LineNumber - 1);
        }

        int startIdx = node.LineNumber - 1;
        int endIdx = -1;

        int index = -1;
        for (int i = 0; i < OutlineNodes.Count; i++)
        {
            if (OutlineNodes[i].LineNumber == node.LineNumber && OutlineNodes[i].Title == node.Title)
            {
                index = i;
                break;
            }
        }

        if (index != -1)
        {
            for (int i = index + 1; i < OutlineNodes.Count; i++)
            {
                var nextNode = OutlineNodes[i];
                if (nextNode.Level <= node.Level)
                {
                    endIdx = nextNode.LineNumber - 2;
                    break;
                }
            }
        }

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n');
        if (endIdx == -1 || endIdx >= lines.Length)
        {
            endIdx = lines.Length - 1;
        }

        return (startIdx, endIdx);
    }

    public void MoveOutlineNodeText(OutlineNodeViewModel sourceNode, OutlineNodeViewModel targetNode, WorkspaceDropPosition position)
    {
        var (sourceStart, sourceEnd) = GetOutlineNodeLineRange(sourceNode);
        var (targetStart, targetEnd) = GetOutlineNodeLineRange(targetNode);

        if (sourceStart == -1 || targetStart == -1) return;
        if (sourceStart <= targetStart && sourceEnd >= targetEnd)
        {
            // Cannot drop a node inside itself or its children
            return;
        }

        int targetInsert = position == WorkspaceDropPosition.Above ? targetStart : (targetEnd + 1);

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();
        int sourceCount = sourceEnd - sourceStart + 1;
        if (sourceStart < 0 || sourceStart >= lines.Count || sourceCount <= 0 || sourceStart + sourceCount > lines.Count)
        {
            return;
        }

        var movedLines = lines.GetRange(sourceStart, sourceCount);
        lines.RemoveRange(sourceStart, sourceCount);

        int adjustedTarget = targetInsert;
        if (adjustedTarget > sourceStart)
        {
            adjustedTarget = Math.Max(0, adjustedTarget - sourceCount);
        }

        if (adjustedTarget > lines.Count)
        {
            adjustedTarget = lines.Count;
        }

        lines.InsertRange(adjustedTarget, movedLines);

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Join("\n", lines);
            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        RefreshParsedDocument();
    }

    public (int startIdx, int endIdx) GetBeatBoardCardLineRange(BeatBoardCardViewModel card)
        => GetCardLineRange(card, includeNestedBlock: true);

    // The card's own lines only: the heading/note element plus its trailing
    // synopsis lines — never the content nested beneath a section.
    private (int startIdx, int endIdx) GetCardOwnLineRange(BeatBoardCardViewModel card)
        => GetCardLineRange(card, includeNestedBlock: false);

    private (int startIdx, int endIdx) GetCardLineRange(BeatBoardCardViewModel card, bool includeNestedBlock)
    {
        if (_lastParsed == null) return (-1, -1);
        var element = _lastParsed.Elements.FirstOrDefault(e => e.Id == card.Id);
        if (element == null) return (-1, -1);

        int startLineIdx = element.LineIndex;
        int endLineIdx = element.EndLineIndex;

        int elementIndex = -1;
        for (int k = 0; k < _lastParsed.Elements.Count; k++)
        {
            if (_lastParsed.Elements[k] == element)
            {
                elementIndex = k;
                break;
            }
        }

        if (elementIndex != -1)
        {
            for (int i = elementIndex + 1; i < _lastParsed.Elements.Count; i++)
            {
                var nextEl = _lastParsed.Elements[i];
                if (nextEl.IsSuppressed && nextEl.Type == ScreenplayElementType.Synopsis)
                {
                    endLineIdx = nextEl.EndLineIndex;
                }
                else
                {
                    break;
                }
            }
        }

        // An Act/Sequence card represents its whole block on the board, so dragging
        // it moves everything nested beneath it: extend the range to just before the
        // next section of the same or higher level (or the end of the document).
        if (includeNestedBlock && element is SectionElement section && elementIndex != -1)
        {
            var blockEndLineIdx = CountEditorLines() - 1;
            for (int i = elementIndex + 1; i < _lastParsed.Elements.Count; i++)
            {
                if (_lastParsed.Elements[i] is SectionElement nextSection && nextSection.SectionDepth <= section.SectionDepth)
                {
                    blockEndLineIdx = nextSection.LineIndex - 1;
                    break;
                }
            }

            endLineIdx = Math.Max(endLineIdx, blockEndLineIdx);
        }

        // A Scene's block runs until the next scene heading or section starts.
        if (includeNestedBlock && element is SceneHeadingElement && elementIndex != -1)
        {
            var blockEndLineIdx = CountEditorLines() - 1;
            for (int i = elementIndex + 1; i < _lastParsed.Elements.Count; i++)
            {
                var nextEl = _lastParsed.Elements[i];
                if (nextEl.Type is ScreenplayElementType.SceneHeading or ScreenplayElementType.Section)
                {
                    blockEndLineIdx = nextEl.LineIndex - 1;
                    break;
                }
            }

            endLineIdx = Math.Max(endLineIdx, blockEndLineIdx);
        }

        return (startLineIdx, endLineIdx);
    }

    private int CountEditorLines()
    {
        var text = (EditorContent ?? string.Empty).Replace("\r\n", "\n");
        var count = 1;
        foreach (var ch in text)
        {
            if (ch == '\n') count++;
        }
        return count;
    }

    public void MoveBeatBoardCardText(BeatBoardCardViewModel sourceCard, BeatBoardCardViewModel targetCard, bool insertAfter)
    {
        var (sourceStart, sourceEnd) = GetBeatBoardCardLineRange(sourceCard);
        var (targetStart, targetEnd) = GetBeatBoardCardLineRange(targetCard);

        if (sourceStart == -1 || targetStart == -1) return;
        if (sourceStart == targetStart) return;
        if (targetStart >= sourceStart && targetEnd <= sourceEnd)
        {
            // Cannot drop an Act/Sequence onto a card nested inside its own block.
            return;
        }

        int targetInsert = insertAfter ? (targetEnd + 1) : targetStart;

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();
        int sourceCount = sourceEnd - sourceStart + 1;
        if (sourceStart < 0 || sourceStart >= lines.Count || sourceCount <= 0 || sourceStart + sourceCount > lines.Count)
        {
            return;
        }

        var movedLines = lines.GetRange(sourceStart, sourceCount);
        lines.RemoveRange(sourceStart, sourceCount);

        int adjustedTarget = targetInsert;
        if (adjustedTarget > sourceStart)
        {
            adjustedTarget = Math.Max(0, adjustedTarget - sourceCount);
        }

        if (adjustedTarget > lines.Count)
        {
            adjustedTarget = lines.Count;
        }

        lines.InsertRange(adjustedTarget, movedLines);

        _suppressDirtyTracking = true;
        try
        {
            EditorContent = string.Join("\n", lines);
            IsDirty = true;
        }
        finally
        {
            _suppressDirtyTracking = false;
        }

        RefreshParsedDocument();
    }
}

/// <summary>
/// One horizontal band of the Beat Board: an Act and everything nested under it.
/// A lane without an <see cref="ActCard"/> holds cards that appear before any Act.
/// </summary>
public partial class BeatBoardLaneViewModel : ObservableObject
{
    [ObservableProperty]
    private bool _isExpanded = true;

    public BeatBoardCardViewModel? ActCard { get; init; }
    public bool HasActCard => ActCard != null;
    public ObservableCollection<BeatBoardGroupViewModel> Groups { get; } = new();
}

/// <summary>
/// A Sequence cluster inside an act lane. A group without a
/// <see cref="SequenceCard"/> holds scenes sitting directly under the act.
/// </summary>
public partial class BeatBoardGroupViewModel : ObservableObject
{
    [ObservableProperty]
    private bool _isExpanded = true;

    public BeatBoardCardViewModel? SequenceCard { get; init; }
    public bool HasSequenceCard => SequenceCard != null;
    public ObservableCollection<BeatBoardCardViewModel> Cards { get; } = new();
}

public partial class BeatBoardCardViewModel : ObservableObject
{
    [ObservableProperty]
    private Guid _id = Guid.NewGuid();

    [ObservableProperty]
    private string _heading = string.Empty;

    [ObservableProperty]
    private string _description = string.Empty;

    [ObservableProperty]
    private int _lineNumber;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsActHeader))]
    [NotifyPropertyChangedFor(nameof(IsSequenceHeader))]
    [NotifyPropertyChangedFor(nameof(IsLeaf))]
    private string _type = "Scene";

    public bool IsActHeader => Type == "Act";
    public bool IsSequenceHeader => Type == "Sequence";
    public bool IsLeaf => !IsActHeader && !IsSequenceHeader;

    [ObservableProperty]
    private int _level = 2;

    [ObservableProperty]
    private bool _isEditing;

    [ObservableProperty]
    private string _editingHeading = string.Empty;

    [ObservableProperty]
    private string _editingDescription = string.Empty;

    [ObservableProperty]
    private string _editingType = "Scene";
}
