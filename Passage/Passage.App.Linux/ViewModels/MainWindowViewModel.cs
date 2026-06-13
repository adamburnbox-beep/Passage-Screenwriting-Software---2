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
    private bool _isBoardModeActive;
    private bool _suppressDirtyTracking;
    private bool _isDirty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsScreenplayMode))]
    [NotifyPropertyChangedFor(nameof(ModeStatusText))]
    private WriteMode _currentMode = WriteMode.Screenplay;

    public bool IsScreenplayMode => CurrentMode == WriteMode.Screenplay;
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
        get => _isBoardModeActive;
        set => SetProperty(ref _isBoardModeActive, value);
    }

    public ObservableCollection<BeatBoardCardViewModel> BeatBoardCards { get; }
    public ObservableCollection<IExporter> AvailableExporters { get; }

    // Sidebar properties
    public bool HasOutlineItems => OutlineRoots.Count > 0;
    public bool HasNoteItems => NotesRoots.Count > 0;
    public string OutlineEmptyMessage => "Sections, synopses, and scene headings will appear here.";
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
    private void New()
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
    }

    [RelayCommand]
    private async Task Open()
    {
        if (_window == null) return;

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
            var file = files[0];
            try
            {
                using var stream = await file.OpenReadAsync();
                using var reader = new StreamReader(stream);
                var text = await reader.ReadToEndAsync();

                _suppressDirtyTracking = true;
                try
                {
                    EditorContent = text;
                    _currentFilePath = file.Path.LocalPath;
                    IsDirty = false;
                    ResetSessionGoal();

                    // Auto-detect mode based on file extension
                    var ext = Path.GetExtension(_currentFilePath);
                    if (string.Equals(ext, ".md", StringComparison.OrdinalIgnoreCase))
                    {
                        CurrentMode = WriteMode.Markdown;
                    }
                    else
                    {
                        CurrentMode = WriteMode.Screenplay;
                    }
                }
                finally
                {
                    _suppressDirtyTracking = false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening file: {ex.Message}");
            }
        }
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
        }
        catch (Exception ex)
        {
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving file: {ex.Message}");
            }
        }
    }

    [RelayCommand]
    private void Close()
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
    }

    [RelayCommand]
    private void Exit()
    {
        _window?.Close();
    }

    [RelayCommand]
    private void Undo()
    {
    }

    [RelayCommand]
    private void Redo()
    {
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
        if (App.Current != null)
        {
            App.Current.RequestedThemeVariant = ThemeVariant.Dark;
            UpdateThemeResources(false);
        }
    }

    [RelayCommand]
    private void SetLightTheme()
    {
        if (App.Current != null)
        {
            App.Current.RequestedThemeVariant = ThemeVariant.Light;
            UpdateThemeResources(true);
        }
    }

    private void UpdateThemeResources(bool isLight)
    {
        if (App.Current?.Resources == null) return;
        var Resources = App.Current.Resources;

        if (isLight)
        {
            Resources["ThemeBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F1F1F1"));
            Resources["WindowBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F6F6F6"));
            Resources["SurfaceBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F6F6F6"));
            Resources["SurfaceMutedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#EAEAEA"));
            Resources["SurfaceRaisedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["SurfaceBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["ControlBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#EEEEEE"));
            Resources["ControlForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["ControlBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["ControlAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
            Resources["ControlPressedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DCDCDB"));
            Resources["ControlPressedForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["HeaderText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["SecondaryText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#5E5E5E"));
            Resources["MutedText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#8E8E8E"));
            Resources["EditorForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["EditorPageBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["WindowForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["CardBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["CardBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#E0E0E0"));
            Resources["CardAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
        }
        else
        {
            Resources["ThemeBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["WindowBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#242424"));
            Resources["SurfaceBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#242424"));
            Resources["SurfaceMutedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["SurfaceRaisedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2D2D2D"));
            Resources["SurfaceBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["ControlBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#303030"));
            Resources["ControlForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["ControlBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["ControlAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
            Resources["ControlPressedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3A3A3A"));
            Resources["ControlPressedForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["HeaderText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["SecondaryText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#B0B0B0"));
            Resources["MutedText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#787878"));
            Resources["EditorForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["EditorPageBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["WindowForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["CardBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#303030"));
            Resources["CardBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3D3D3D"));
            Resources["CardAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
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

    private void UpdateCardInScript(BeatBoardCardViewModel card)
    {
        if (_lastParsed == null) return;
        var element = _lastParsed.Elements.FirstOrDefault(e => e.Id == card.Id);
        if (element == null) return;

        var lines = (EditorContent ?? string.Empty).Replace("\r\n", "\n").Split('\n').ToList();
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
            var textBox = mainWin.FindControl<TextBox>("EditorTextBox");
            if (textBox != null)
            {
                var caretIndex = textBox.CaretIndex;
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

    [RelayCommand]
    private async Task Export(IExporter? exporter)
    {
        if (exporter == null || _window == null) return;

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
            // Parse the Fountain file content
            var parsed = _parser.Parse(snapshotText, _lineTypeOverrides);

            if (version == _outlineRefreshVersion)
            {
                _lastParseDuration = TimeSpan.FromMilliseconds(Math.Max(0, Environment.TickCount64 - parseStartedAt));
                ApplyParsedDocument(parsed);
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
        RefreshParsedDocument();
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
    }

    // Goal operations
    private void RefreshGoalState()
    {
        CaptureSessionGoalBaselineIfNeeded();

        var currentWordCount = _goalProgressCalculator.CalculateWordCount(EditorContent);
        var currentPageCount = _pageEstimator.EstimatePageCount(_lastParsed);

        RefreshOverallGoal(currentWordCount, currentPageCount);
        RefreshSessionGoal(currentWordCount, currentPageCount);

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

        var state = new SessionState(new[] { docState }, 0, width, height, x, y, windowState);
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
        return (startLineIdx, endLineIdx);
    }

    public void MoveBeatBoardCardText(BeatBoardCardViewModel sourceCard, BeatBoardCardViewModel targetCard, bool insertAfter)
    {
        var (sourceStart, sourceEnd) = GetBeatBoardCardLineRange(sourceCard);
        var (targetStart, targetEnd) = GetBeatBoardCardLineRange(targetCard);

        if (sourceStart == -1 || targetStart == -1) return;
        if (sourceStart == targetStart) return;

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
    private string _type = "Scene";

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
