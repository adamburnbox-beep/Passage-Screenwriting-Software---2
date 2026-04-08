using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Threading;
using Passage.App.Services;
using Passage.App.Utilities;
using Passage.Core;
using Passage.Core.Goals;
using Passage.Parser;
using Passage.Parser.Goals;

namespace Passage.App.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    // Beat-inspired outline model: sections define the hierarchy, and synopses can group the scene headings beneath them.
    private static readonly ScreenplayElementType[] TabCycleTypes =
    [
        ScreenplayElementType.Action,
        ScreenplayElementType.SceneHeading,
        ScreenplayElementType.Character,
        ScreenplayElementType.Dialogue,
        ScreenplayElementType.Parenthetical,
        ScreenplayElementType.Transition
    ];

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
    private const double DefaultEditorZoomPercent = 100.0;
    private const double MinimumEditorZoomPercent = 50.0;
    private const double MaximumEditorZoomPercent = 200.0;
    private const double EditorZoomStepPercent = 10.0;
    private const double DefaultBoardZoomPercent = 100.0;
    private const double MinimumBoardZoomPercent = 20.0;
    private const double MaximumBoardZoomPercent = 200.0;
    private const double BoardZoomStepPercent = 10.0;
    private const string DefaultNewBoardCardHeading = "NEW SCENE";
    private const string LegacyDefaultNewBoardCardDescription = "Double-click to edit...";
    private const string DefaultNewBoardCardDescription = "Click the pencil to edit. Double-click to locate in script.";
    private readonly FountainParser _parser = new();
    private readonly GoalProgressCalculator _goalProgressCalculator = new();
    private readonly ScreenplayPageEstimator _pageEstimator = new();

    private string _titlePageText = string.Empty;
    private string _documentText = string.Empty;
    private bool _isInternalSync;
    private string? _documentPath;
    private bool _isDirty;
    private bool _suppressDirtyTracking;
    private int _currentLineNumber = 1;
    private string _currentElementType = "Action";
    private string _currentElementText = "No screenplay element";
    private string _enterContinuationText = "Action";
    private readonly Dictionary<int, ScreenplayElementType> _lineTypeOverrides = new();
    private readonly HashSet<int> _automaticActionLineOverrides = [];
    private int? _selectedOutlineLineNumber;
    private readonly DispatcherTimer _refreshTimer;
    private readonly DispatcherTimer _recoveryTimer;
    private ParsedScreenplay _lastParsed = new(string.Empty, Array.Empty<ScreenplayElement>());
    private int _outlineRefreshVersion;
    private bool _outlineRefreshInProgress;
    private bool _outlineRefreshPending;
    private TimeSpan _lastParseDuration = TimeSpan.Zero;
    private GoalType _selectedGoalType = GoalType.WordCount;
    private int _wordCountGoalTargetValue = 1000;
    private int _pageCountGoalTargetValue = 120;
    private int _timerGoalTargetMinutes = 25;
    private GoalType _sessionSelectedGoalType = GoalType.WordCount;
    private int _sessionWordCountTargetValue = 1000;
    private int _sessionPageCountTargetValue = 120;
    private int _sessionWordCountBaseline;
    private int _sessionPageCountBaseline;
    private bool _sessionGoalBaselineNeedsCapture = true;
    private GoalTimerRuntime? _goalTimerRuntime;
    private double _goalProgressPercent;
    private string _goalCurrentDisplayText = "0 words";
    private string _goalTargetDisplayText = "1000 words";
    private string _goalStateText = "In progress";
    private string _goalTargetUnitLabel = "Words";
    private double _sessionGoalProgressPercent;
    private string _sessionGoalCurrentDisplayText = "0 words";
    private string _sessionGoalTargetDisplayText = "1000 words";
    private string _sessionGoalStateText = "In progress";
    private string _sessionGoalTargetUnitLabel = "Words";
    private string _goalTimerElapsedText = "00:00";
    private string _goalTimerRemainingText = "25:00";
    private string _goalProgressSummaryText = "Words 0";
    private double _editorZoomPercent = DefaultEditorZoomPercent;
    private double _boardZoomPercent = DefaultBoardZoomPercent;
    private bool _isBoardSyncRequired;
    private bool _isBoardModeActive;
    private bool _isBoardDropIndicatorVisible;
    private ScreenplayElement? _boardDropTargetElement;
    private ScreenplayElement? _selectedBoardElement;
    private readonly ICollectionView _scratchpadView;
    private string _scratchpadSearchText = string.Empty;
    private ScreenplayElement? _selectedScratchpadElement;
    private readonly Stack<ScratchpadMoveHistoryEntry> _scratchpadUndoHistory = new();
    private readonly Stack<ScratchpadMoveHistoryEntry> _scratchpadRedoHistory = new();
    private readonly TitlePageViewModel _titlePage = new();
    private bool _suppressTitlePageSync;
    private List<ScreenplayElement>? _originalBoardElements;
    private (ScreenplayElement dragged, ScreenplayElement target, bool insertAfter)? _lastPreviewRequest;
    private (ScreenplayElement dragged, ScreenplayElement target, bool insertAfter)? _pendingPreviewRequest;
    private readonly DispatcherTimer _reorderPreviewTimer;

    public MainWindowViewModel()
    {
        OutlineRoots = new ObservableCollection<OutlineNodeViewModel>();
        NotesRoots = new ObservableCollection<OutlineNodeViewModel>();
        BoardElements = new ObservableCollection<ScreenplayElement>();
        VisibleBoardElements = new ObservableCollection<ScreenplayElement>();
        AttachBoardElementsCollection(BoardElements);
        ScratchpadElements = new ObservableCollection<ScreenplayElement>();
        _scratchpadView = CollectionViewSource.GetDefaultView(ScratchpadElements);
        _scratchpadView.Filter = ShouldIncludeScratchpadElement;
        ScratchpadElements.CollectionChanged += (_, _) =>
        {
            if (_selectedScratchpadElement is not null &&
                !ScratchpadElements.Contains(_selectedScratchpadElement))
            {
                SelectedScratchpadElement = null;
            }

            _scratchpadView.Refresh();
            OnPropertyChanged(nameof(ScratchpadElements));
        };
        MoveToScratchpadCommand = new DelegateCommand<RichTextBox>(
            execute: ExecuteMoveToScratchpad,
            canExecute: CanMoveToScratchpad);
        DeleteScratchpadCardCommand = new DelegateCommand<object>(
            execute: ExecuteDeleteScratchpadCard,
            canExecute: CanDeleteScratchpadCard);
        ToggleBoardElementCollapsedCommand = new DelegateCommand<object>(
            execute: ExecuteToggleBoardElementCollapsed);
        ExpandAllBoardCardsCommand = new DelegateCommand<object>(
            execute: ExecuteExpandAllBoardCards);
        CollapseAllBoardCardsCommand = new DelegateCommand<object>(
            execute: ExecuteCollapseAllBoardCards);
        ExpandAllOutlineNodesCommand = new DelegateCommand<object>(
            execute: ExecuteExpandAllOutlineNodes);
        CollapseAllOutlineNodesCommand = new DelegateCommand<object>(
            execute: ExecuteCollapseAllOutlineNodes);
        CreateNewCardCommand = new DelegateCommand<object>(
            execute: ExecuteCreateNewCard);
        SyncBoardToScriptCommand = new DelegateCommand<object>(
            execute: ExecuteSyncBoardToScript,
            canExecute: CanSyncBoardToScript);
        TitlePageEntries = new ObservableCollection<TitlePageEntry>();
        PreviewElements = new ObservableCollection<PreviewElementItem>();
        _titlePage.PropertyChanged += TitlePage_PropertyChanged;
        _titlePage.CustomEntries.CollectionChanged += TitlePageCustomEntries_CollectionChanged;
        _refreshTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(75)
        };
        _refreshTimer.Tick += (_, _) => RefreshParsedDocument();
        _recoveryTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _recoveryTimer.Tick += (_, _) => SaveRecoverySnapshot();
        _reorderPreviewTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(500)
        };
        _reorderPreviewTimer.Tick += (_, _) =>
        {
            if (_pendingPreviewRequest.HasValue)
            {
                var req = _pendingPreviewRequest.Value;
                ApplyReorderPreview(req.dragged, req.target, req.insertAfter);
            }
            _reorderPreviewTimer.Stop();
        };
        UpdateOutline();
        // Seed the session baseline immediately so a new blank document starts at zero.
        CaptureSessionGoalBaseline();
        RefreshGoalState();
    }

    public ObservableCollection<OutlineNodeViewModel> OutlineRoots { get; private set; }

    public ObservableCollection<OutlineNodeViewModel> NotesRoots { get; private set; }

    public ObservableCollection<ScreenplayElement> BoardElements { get; private set; }

    public ObservableCollection<ScreenplayElement> VisibleBoardElements { get; private set; }

    public ObservableCollection<ScreenplayElement> ScratchpadElements { get; private set; }

    public bool IsBoardSyncRequired
    {
        get => _isBoardSyncRequired;
        private set => SetProperty(ref _isBoardSyncRequired, value);
    }

    public bool IsBoardModeActive
    {
        get => _isBoardModeActive;
        set
        {
            if (SetProperty(ref _isBoardModeActive, value))
            {
                OnPropertyChanged(nameof(ActiveZoomDisplayText));
            }
        }
    }

    public bool HasVisibleBoardItems => VisibleBoardElements.Count > 0;

    public bool IsBoardDropIndicatorVisible
    {
        get => _isBoardDropIndicatorVisible;
        private set => SetProperty(ref _isBoardDropIndicatorVisible, value);
    }


    public ScreenplayElement? BoardDropTargetElement
    {
        get => _boardDropTargetElement;
        private set => SetProperty(ref _boardDropTargetElement, value);
    }

    public ScreenplayElement? SelectedBoardElement
    {
        get => _selectedBoardElement;
        private set
        {
            if (!SetProperty(ref _selectedBoardElement, value))
            {
                return;
            }

            CommandManager.InvalidateRequerySuggested();
        }
    }

    public string ScratchpadSearchText
    {
        get => _scratchpadSearchText;
        set
        {
            if (!SetProperty(ref _scratchpadSearchText, value ?? string.Empty))
            {
                return;
            }

            _scratchpadView.Refresh();
        }
    }

    public ScreenplayElement? SelectedScratchpadElement
    {
        get => _selectedScratchpadElement;
        set
        {
            if (!SetProperty(ref _selectedScratchpadElement, value))
            {
                return;
            }

            CommandManager.InvalidateRequerySuggested();
        }
    }

    public ICommand MoveToScratchpadCommand { get; }

    public ICommand DeleteScratchpadCardCommand { get; }

    public ICommand ToggleBoardElementCollapsedCommand { get; }

    public ICommand ExpandAllBoardCardsCommand { get; }

    public ICommand CollapseAllBoardCardsCommand { get; }

    public ICommand ExpandAllOutlineNodesCommand { get; }

    public ICommand CollapseAllOutlineNodesCommand { get; }

    public ICommand CreateNewCardCommand { get; }

    public ICommand SyncBoardToScriptCommand { get; }

    public ObservableCollection<TitlePageEntry> TitlePageEntries { get; private set; }

    public ObservableCollection<PreviewElementItem> PreviewElements { get; private set; }

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
            var normalizedValue = NormalizeOverallGoalType(value);
            if (!SetProperty(ref _selectedGoalType, normalizedValue))
            {
                return;
            }

            OnPropertyChanged(nameof(GoalTargetValue));
            RefreshGoalState();
        }
    }

    public GoalType SessionSelectedGoalType
    {
        get => _sessionSelectedGoalType;
        set
        {
            var normalizedValue = NormalizeSessionGoalType(value);
            if (!SetProperty(ref _sessionSelectedGoalType, normalizedValue))
            {
                return;
            }

            if (normalizedValue != GoalType.Timer)
            {
                _goalTimerRuntime?.Stop();
            }

            OnPropertyChanged(nameof(SessionGoalTargetValue));
            OnPropertyChanged(nameof(IsTimerGoal));
            RefreshGoalState();
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
                    if (_sessionWordCountTargetValue == clampedValue)
                    {
                        return;
                    }

                    _sessionWordCountTargetValue = clampedValue;
                    break;
                case GoalType.PageCount:
                    if (_sessionPageCountTargetValue == clampedValue)
                    {
                        return;
                    }

                    _sessionPageCountTargetValue = clampedValue;
                    break;
                case GoalType.Timer:
                    if (_timerGoalTargetMinutes == clampedValue)
                    {
                        return;
                    }

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
                    if (_wordCountGoalTargetValue == clampedValue)
                    {
                        return;
                    }

                    _wordCountGoalTargetValue = clampedValue;
                    break;
                case GoalType.PageCount:
                    if (_pageCountGoalTargetValue == clampedValue)
                    {
                        return;
                    }

                    _pageCountGoalTargetValue = clampedValue;
                    break;
            }

            OnPropertyChanged(nameof(GoalTargetValue));
            RefreshGoalState();
        }
    }

    public double GoalProgressPercent
    {
        get => _goalProgressPercent;
        private set => SetProperty(ref _goalProgressPercent, value);
    }

    public double SessionGoalProgressPercent
    {
        get => _sessionGoalProgressPercent;
        private set => SetProperty(ref _sessionGoalProgressPercent, value);
    }

    public string GoalCurrentDisplayText
    {
        get => _goalCurrentDisplayText;
        private set => SetProperty(ref _goalCurrentDisplayText, value);
    }

    public string SessionGoalCurrentDisplayText
    {
        get => _sessionGoalCurrentDisplayText;
        private set => SetProperty(ref _sessionGoalCurrentDisplayText, value);
    }

    public string GoalTargetDisplayText
    {
        get => _goalTargetDisplayText;
        private set => SetProperty(ref _goalTargetDisplayText, value);
    }

    public string SessionGoalTargetDisplayText
    {
        get => _sessionGoalTargetDisplayText;
        private set => SetProperty(ref _sessionGoalTargetDisplayText, value);
    }

    public string GoalStateText
    {
        get => _goalStateText;
        private set => SetProperty(ref _goalStateText, value);
    }

    public string SessionGoalStateText
    {
        get => _sessionGoalStateText;
        private set => SetProperty(ref _sessionGoalStateText, value);
    }

    public string GoalTargetUnitLabel
    {
        get => _goalTargetUnitLabel;
        private set => SetProperty(ref _goalTargetUnitLabel, value);
    }

    public string SessionGoalTargetUnitLabel
    {
        get => _sessionGoalTargetUnitLabel;
        private set => SetProperty(ref _sessionGoalTargetUnitLabel, value);
    }

    public string GoalTimerElapsedText
    {
        get => _goalTimerElapsedText;
        private set => SetProperty(ref _goalTimerElapsedText, value);
    }

    public string GoalTimerRemainingText
    {
        get => _goalTimerRemainingText;
        private set => SetProperty(ref _goalTimerRemainingText, value);
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

    public string GoalProgressSummaryText
    {
        get => _goalProgressSummaryText;
        private set => SetProperty(ref _goalProgressSummaryText, value);
    }

    public int? SelectedOutlineLineNumber
    {
        get => _selectedOutlineLineNumber;
        set => SetProperty(ref _selectedOutlineLineNumber, value);
    }

    public string? DocumentPath
    {
        get => _documentPath;
        private set
        {
            if (!SetProperty(ref _documentPath, value))
            {
                return;
            }

            OnPropertyChanged(nameof(WindowTitle));
            OnPropertyChanged(nameof(DisplayName));
        }
    }

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

            OnPropertyChanged(nameof(WindowTitle));
            OnPropertyChanged(nameof(DisplayName));
        }
    }

    public string WindowTitle
    {
        get
        {
            var baseTitle = string.IsNullOrWhiteSpace(DocumentPath)
                ? "Passage"
                : $"Passage - {Path.GetFileName(DocumentPath)}";

            return IsDirty ? $"{baseTitle} *" : baseTitle;
        }
    }

    public string DisplayName
    {
        get
        {
            var baseName = string.IsNullOrWhiteSpace(DocumentPath)
                ? "Untitled"
                : Path.GetFileName(DocumentPath);

            return IsDirty ? $"{baseName} *" : baseName;
        }
    }

    public TitlePageViewModel TitlePage => _titlePage;

    public string TitlePageText
    {
        get => _titlePageText;
        set
        {
            if (SetProperty(ref _titlePageText, value))
            {
                UpdateViewModelFromTitlePageText();
            }
        }
    }

    public string DocumentText
    {
        get => _documentText;
        set
        {
            if (!SetProperty(ref _documentText, value ?? string.Empty))
            {
                return;
            }

            if (!_suppressDirtyTracking)
            {
                IsDirty = true;
                StartRecoveryAutosave();
            }

            OnDocumentTextChanged();
        }
    }

    private void OnDocumentTextChanged()
    {
        if (_isInternalSync)
        {
            return;
        }

        ScheduleOutlineRefresh();
    }

    public int CurrentLineNumber
    {
        get => _currentLineNumber;
        private set => SetProperty(ref _currentLineNumber, value);
    }

    public string CurrentElementType
    {
        get => _currentElementType;
        private set => SetProperty(ref _currentElementType, value);
    }

    public string CurrentElementText
    {
        get => _currentElementText;
        private set => SetProperty(ref _currentElementText, value);
    }

    public string EnterContinuationText
    {
        get => _enterContinuationText;
        private set => SetProperty(ref _enterContinuationText, value);
    }

    public double EditorZoomPercent
    {
        get => _editorZoomPercent;
        set
        {
            var clampedValue = Math.Clamp(value, MinimumEditorZoomPercent, MaximumEditorZoomPercent);
            if (Math.Abs(_editorZoomPercent - clampedValue) <= 0.01)
            {
                return;
            }

            if (!SetProperty(ref _editorZoomPercent, clampedValue))
            {
                return;
            }

            OnPropertyChanged(nameof(EditorZoomScale));
            OnPropertyChanged(nameof(EditorZoomDisplayText));
            OnPropertyChanged(nameof(ActiveZoomDisplayText));
        }
    }

    public double BoardZoomPercent
    {
        get => _boardZoomPercent;
        set
        {
            var clampedValue = Math.Clamp(value, MinimumBoardZoomPercent, MaximumBoardZoomPercent);
            if (Math.Abs(_boardZoomPercent - clampedValue) <= 0.01)
            {
                return;
            }

            if (!SetProperty(ref _boardZoomPercent, clampedValue))
            {
                return;
            }

            OnPropertyChanged(nameof(BoardZoomScale));
            OnPropertyChanged(nameof(BoardZoomDisplayText));
            OnPropertyChanged(nameof(ActiveZoomDisplayText));
        }
    }

    public double EditorZoomScale => EditorZoomPercent / 100.0;

    public double BoardZoomScale => BoardZoomPercent / 100.0;

    public string EditorZoomDisplayText => $"{(int)Math.Round(EditorZoomPercent)}%";

    public string BoardZoomDisplayText => $"{(int)Math.Round(BoardZoomPercent)}%";

    public string ActiveZoomDisplayText => IsBoardModeActive ? BoardZoomDisplayText : EditorZoomDisplayText;

    public event PropertyChangedEventHandler? PropertyChanged;

    public event EventHandler? OutlineUpdated;

    public ParsedScreenplay CreateParsedSnapshot()
    {
        var parsed = _parser.Parse(GetScriptSourceText(), GetLineTypeOverridesSnapshot());
        return StitchParsedSnapshot(parsed);
    }

    public ParsedScreenplay GetLatestParsedSnapshot()
    {
        var scriptSourceText = GetScriptSourceText();
        if (!string.Equals(_lastParsed.RawText, TitlePageText + (TitlePageText.Length > 0 ? "\n\n" : "") + scriptSourceText, StringComparison.Ordinal))
        {
            ScheduleOutlineRefresh();
        }

        return _lastParsed;
    }

    private ParsedScreenplay StitchParsedSnapshot(ParsedScreenplay bodyParsed)
    {
        var titleEndPrefix = TitlePageText.Length > 0 ? "\n" : "";
        var titlePageData = _parser.Parse(TitlePageText + "\n", null).TitlePage;
        return new ParsedScreenplay(
            TitlePageText + titleEndPrefix + bodyParsed.RawText,
            titlePageData,
            bodyParsed.Elements,
            bodyParsed.LineTypeOverrides);
    }

    public void RefreshParsedSnapshotNow()
    {
        _refreshTimer.Stop();
        _outlineRefreshVersion++;
        _outlineRefreshPending = false;

        var parseStartedAt = Environment.TickCount64;
        var parsed = _parser.Parse(GetScriptSourceText(), GetLineTypeOverridesSnapshot());
        var stitched = StitchParsedSnapshot(parsed);
        _lastParseDuration = TimeSpan.FromMilliseconds(Math.Max(0, Environment.TickCount64 - parseStartedAt));
        ApplyParsedDocument(stitched);
    }

    public void NewDocument()
    {
        ReplaceDocument(string.Empty, null, false);
    }

    public void LoadDocument(string text, string? filePath)
    {
        LoadDocument(text, filePath, false);
    }

    public void LoadDocument(string text, string? filePath, bool isDirty)
    {
        ReplaceDocument(text ?? string.Empty, filePath, isDirty);
    }

    public void ApplyGoalConfiguration(GoalConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        _selectedGoalType = NormalizeOverallGoalType(configuration.SelectedGoalType);
        _wordCountGoalTargetValue = Math.Max(0, configuration.WordCountTargetValue);
        _pageCountGoalTargetValue = Math.Max(0, configuration.PageCountTargetValue);
        _timerGoalTargetMinutes = Math.Max(0, configuration.TimerTargetMinutes);

        OnPropertyChanged(nameof(SelectedGoalType));
        OnPropertyChanged(nameof(GoalTargetValue));
        RefreshGoalState();
    }

    public void ApplySessionGoalConfiguration(SessionGoalConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        _sessionSelectedGoalType = NormalizeSessionGoalType(configuration.SelectedGoalType);
        _sessionWordCountTargetValue = Math.Max(0, configuration.WordCountTargetValue);
        _sessionPageCountTargetValue = Math.Max(0, configuration.PageCountTargetValue);
        _timerGoalTargetMinutes = Math.Max(0, configuration.TimerTargetMinutes);

        OnPropertyChanged(nameof(SessionSelectedGoalType));
        OnPropertyChanged(nameof(SessionGoalTargetValue));
        OnPropertyChanged(nameof(IsTimerGoal));
        RefreshGoalState();
    }

    public void MarkSaved()
    {
        IsDirty = false;
        StopRecoveryAutosave();
        RecoveryStorage.ClearRecoveryFile();
    }

    public void IncreaseEditorZoom()
    {
        SetEditorZoomPercent(EditorZoomPercent + EditorZoomStepPercent);
    }

    public void DecreaseEditorZoom()
    {
        SetEditorZoomPercent(EditorZoomPercent - EditorZoomStepPercent);
    }

    public void ResetEditorZoom()
    {
        SetEditorZoomPercent(DefaultEditorZoomPercent);
    }

    public void IncreaseBoardZoom()
    {
        BoardZoomPercent += BoardZoomStepPercent;
    }

    public void DecreaseBoardZoom()
    {
        BoardZoomPercent -= BoardZoomStepPercent;
    }

    public void ResetBoardZoom()
    {
        BoardZoomPercent = DefaultBoardZoomPercent;
    }

    public void SetFilePath(string? filePath)
    {
        DocumentPath = filePath;
    }

    public bool TryMoveToScratchpad(RichTextBox? richTextBox)
    {
        if (richTextBox is null)
        {
            return false;
        }

        var textRange = new TextRange(richTextBox.Selection.Start, richTextBox.Selection.End);
        if (textRange.Start.CompareTo(textRange.End) == 0)
        {
            return false;
        }

        var selectionLength = RichTextBoxTextUtilities.GetSelectionLength(richTextBox);
        if (selectionLength <= 0)
        {
            return false;
        }

        var selectionStart = RichTextBoxTextUtilities.GetSelectionStart(richTextBox);
        var scriptText = RichTextBoxTextUtilities.GetPlainText(richTextBox);
        if (selectionStart < 0 || selectionStart >= scriptText.Length)
        {
            return false;
        }

        var selectedText = textRange.Text;
        if (string.IsNullOrWhiteSpace(selectedText))
        {
            return false;
        }

        var parsedSelection = _parser.Parse(selectedText);
        var scratchpadItems = BuildScratchpadElementsFromSnippet(selectedText, parsedSelection.Elements).ToArray();

        AdjustLineTypeOverridesForRemovedRange(scriptText, selectionStart, selectionLength);

        try
        {
            richTextBox.BeginChange();
            textRange.Text = string.Empty;
        }
        finally
        {
            richTextBox.EndChange();
        }

        var documentTextAfterMove = RichTextBoxTextUtilities.GetPlainText(richTextBox);
        AddScratchpadCards(scratchpadItems);
        RecordScratchpadMove(scriptText, documentTextAfterMove, scratchpadItems);
        DocumentText = documentTextAfterMove;
        RefreshParsedSnapshotNow();
        return true;
    }

    public void SynchronizeScratchpadWithUndoRedo(string? documentText)
    {
        var currentText = documentText ?? string.Empty;
        if (TryUndoScratchpadMove(currentText))
        {
            return;
        }

        _ = TryRedoScratchpadMove(currentText);
    }

    private ParsedScreenplay GetCurrentParsedSnapshot()
    {
        var scriptSourceText = GetScriptSourceText();
        if (string.Equals(_lastParsed.RawText, scriptSourceText, StringComparison.Ordinal))
        {
            return _lastParsed;
        }

        return _parser.Parse(scriptSourceText, GetLineTypeOverridesSnapshot());
    }

    public void SelectBoardElement(ScreenplayElement? element)
    {
        if (element is null)
        {
            SelectedBoardElement = null;
            return;
        }

        var elementIndex = FindBoardElementIndex(element);
        SelectedBoardElement = elementIndex >= 0 ? BoardElements[elementIndex] : null;
    }

    public ScreenplayElement? TryUpdateBoardElementContent(
        ScreenplayElement? element,
        string? heading,
        string? sceneHeading,
        string? description)
    {
        if (element is null)
        {
            return null;
        }

        var elementIndex = FindBoardElementIndex(element);
        if (elementIndex < 0)
        {
            return null;
        }

        var currentElement = BoardElements[elementIndex];
        var normalizedHeading = NormalizeBoardCardHeading(currentElement, heading);
        var normalizedSceneHeading = currentElement.Type == ScreenplayElementType.SceneHeading
            ? NormalizeBoardCardSceneHeading(currentElement, sceneHeading)
            : string.Empty;
        var normalizedDescription = NormalizeBoardCardDescription(description);
        var currentHeading = GetBoardCardHeading(currentElement);
        var currentSceneHeading = GetBoardCardSceneHeading(currentElement);
        var currentDescription = GetBoardCardDescription(currentElement);

        if (currentElement.IsDraft &&
            string.Equals(currentHeading, normalizedHeading, StringComparison.Ordinal) &&
            string.Equals(currentSceneHeading, normalizedSceneHeading, StringComparison.Ordinal) &&
            string.Equals(currentDescription, normalizedDescription, StringComparison.Ordinal))
        {
            return currentElement;
        }

        var updatedElement = CreateBoardDraftElement(
            currentElement,
            currentElement.Type,
            currentElement.Level,
            normalizedHeading,
            normalizedSceneHeading,
            normalizedDescription);
        return ReplaceBoardElement(currentElement, updatedElement);
    }

    public ScreenplayElement? TrySetBoardElementKind(ScreenplayElement? element, string? kind)
    {
        if (element is null || string.IsNullOrWhiteSpace(kind))
        {
            return null;
        }

        var elementIndex = FindBoardElementIndex(element);
        if (elementIndex < 0)
        {
            return null;
        }

        if (!TryResolveBoardCardKind(kind, out var targetType, out var targetLevel))
        {
            return null;
        }

        var currentElement = BoardElements[elementIndex];
        var updatedElement = CreateBoardDraftElement(
            currentElement,
            targetType,
            targetLevel,
            GetBoardCardHeading(currentElement),
            GetBoardCardSceneHeading(currentElement),
            GetBoardCardDescription(currentElement));
        return ReplaceBoardElement(currentElement, updatedElement);
    }

    public bool TrySetBoardElementCollapsed(ScreenplayElement? element, bool isCollapsed)
    {
        if (element is null)
        {
            return false;
        }

        var elementIndex = FindBoardElementIndex(element);
        if (elementIndex < 0)
        {
            return false;
        }

        var currentElement = BoardElements[elementIndex];
        if (currentElement.IsCollapsed == isCollapsed)
        {
            return false;
        }

        var updatedElement = CloneElement(currentElement);
        updatedElement.IsCollapsed = isCollapsed;
        BoardElements[elementIndex] = updatedElement;
        SelectedBoardElement = updatedElement;
        ClearBoardDropIndicator();
        return true;
    }

    public bool TrySetAllBoardElementsCollapsed(bool isCollapsed)
    {
        if (BoardElements.Count == 0)
        {
            return false;
        }

        var changed = false;
        var selectedElementId = _selectedBoardElement?.Id;
        var updatedElements = new List<ScreenplayElement>(BoardElements.Count);

        foreach (var element in BoardElements)
        {
            if (element.IsCollapsed == isCollapsed)
            {
                updatedElements.Add(element);
                continue;
            }

            var updatedElement = CloneElement(element);
            updatedElement.IsCollapsed = isCollapsed;
            updatedElements.Add(updatedElement);
            changed = true;
        }

        if (!changed)
        {
            return false;
        }

        BoardElements = new ObservableCollection<ScreenplayElement>(updatedElements);
        AttachBoardElementsCollection(BoardElements);
        SelectedBoardElement = selectedElementId is null
            ? null
            : VisibleBoardElements.FirstOrDefault(candidate => candidate.Id == selectedElementId);
        ClearBoardDropIndicator();
        OnPropertyChanged(nameof(BoardElements));
        UpdateBoardSyncRequiredState();
        return true;
    }

    public bool TrySetAllOutlineNodesExpanded(bool isExpanded)
    {
        return SetOutlineExpansionState(OutlineRoots, isExpanded);
    }

    public bool TrySetBoardElementLevel(ScreenplayElement? element, int level)
    {
        if (element is null || level < 0)
        {
            return false;
        }

        var elementIndex = FindBoardElementIndex(element);
        if (elementIndex < 0)
        {
            return false;
        }

        var currentElement = BoardElements[elementIndex];
        if (currentElement.Level == level)
        {
            return false;
        }

        var updatedElement = CloneElement(currentElement);
        updatedElement.Level = level;
        BoardElements[elementIndex] = updatedElement;
        SelectedBoardElement = updatedElement;
        UpdateBoardSyncRequiredState();
        return true;
    }

    public void SetBoardDropIndicator(ScreenplayElement? targetElement)
    {
        BoardDropTargetElement = targetElement;
        IsBoardDropIndicatorVisible = BoardDropTargetElement != null;
    }

    public void ClearBoardDropIndicator()
    {
        IsBoardDropIndicatorVisible = false;
        BoardDropTargetElement = null;
    }

    public void PerformReorderPreview(ScreenplayElement dragged, ScreenplayElement target, bool insertAfter)
    {
        if (target == null || dragged == null || ReferenceEquals(dragged, target))
        {
            return;
        }

        // Capture original state immediately so Cancel works even without a shuffle
        if (_originalBoardElements == null)
        {
            _originalBoardElements = BoardElements.ToList();
        }

        // Avoid redundant moves if the target hasn't changed
        if (_lastPreviewRequest.HasValue &&
            ReferenceEquals(_lastPreviewRequest.Value.dragged, dragged) &&
            ReferenceEquals(_lastPreviewRequest.Value.target, target) &&
            _lastPreviewRequest.Value.insertAfter == insertAfter)
        {
            return;
        }

        // Check if this matches the currently pending or active preview
        if (_pendingPreviewRequest.HasValue &&
            ReferenceEquals(_pendingPreviewRequest.Value.dragged, dragged) &&
            ReferenceEquals(_pendingPreviewRequest.Value.target, target) &&
            _pendingPreviewRequest.Value.insertAfter == insertAfter)
        {
            return;
        }

        // Delay the actual shuffle to avoid "blinking" while moving across cards
        _reorderPreviewTimer.Stop();
        _pendingPreviewRequest = (dragged, target, insertAfter);
        _reorderPreviewTimer.Start();
    }

    private void ApplyReorderPreview(ScreenplayElement dragged, ScreenplayElement target, bool insertAfter)
    {
        _lastPreviewRequest = (dragged, target, insertAfter);
        _pendingPreviewRequest = null;

        // Resolve indices and ranges
        if (!TryResolveVisibleBoardDropIndex(target, insertAfter, out var targetIndex))
        {
            return;
        }

        StoryBlockRange movingRange;
        try
        {
            movingRange = StoryHierarchyHelper.GetBlockRange(BoardElements, dragged);
        }
        catch (ArgumentException)
        {
            return;
        }

        // Perform the move in the collection
        var movingElements = BoardElements.Skip(movingRange.Index).Take(movingRange.Count).ToList();
        
        for (int i = 0; i < movingRange.Count; i++)
        {
            BoardElements.RemoveAt(movingRange.Index);
        }

        if (targetIndex > movingRange.Index)
        {
            targetIndex -= movingRange.Count;
        }

        targetIndex = Math.Clamp(targetIndex, 0, BoardElements.Count);

        for (int i = 0; i < movingElements.Count; i++)
        {
            BoardElements.Insert(targetIndex + i, movingElements[i]);
        }
        
        UpdateBoardSyncRequiredState();
    }

    public void CancelReorderPreview()
    {
        _reorderPreviewTimer.Stop();
        _pendingPreviewRequest = null;

        if (_originalBoardElements == null)
        {
            _lastPreviewRequest = null;
            return;
        }

        BoardElements.Clear();
        foreach (var element in _originalBoardElements)
        {
            BoardElements.Add(element);
        }

        _originalBoardElements = null;
        _lastPreviewRequest = null;
        UpdateBoardSyncRequiredState();
    }

    public void FinalizeReorderPreview()
    {
        // If a preview is pending but hasn't fired yet, fire it now
        if (_reorderPreviewTimer.IsEnabled && _pendingPreviewRequest.HasValue)
        {
            _reorderPreviewTimer.Stop();
            var req = _pendingPreviewRequest.Value;
            ApplyReorderPreview(req.dragged, req.target, req.insertAfter);
        }

        _originalBoardElements = null;
        _lastPreviewRequest = null;
        _pendingPreviewRequest = null;
        UpdateBoardSyncRequiredState();
        ExecuteSyncBoardToScript(null);
    }

    public bool TryMoveBoardBlock(ScreenplayElement? element, ScreenplayElement? visibleTargetElement, bool insertAfter)
    {
        if (!TryResolveVisibleBoardDropIndex(visibleTargetElement, insertAfter, out var targetIndex))
        {
            return false;
        }

        return TryMoveBoardBlock(element, targetIndex);
    }

    public bool TryMoveBoardBlockIntoParent(ScreenplayElement? element, ScreenplayElement? parentElement)
    {
        if (element is null ||
            parentElement is null ||
            BoardElements.Count == 0 ||
            !TryResolveBoardChildLevel(parentElement, element, out var targetLevel))
        {
            return false;
        }

        StoryBlockRange movingRange;
        StoryBlockRange parentRange;
        try
        {
            movingRange = StoryHierarchyHelper.GetBlockRange(BoardElements, element);
            parentRange = StoryHierarchyHelper.GetBlockRange(BoardElements, parentElement);
        }
        catch (ArgumentException)
        {
            return false;
        }

        if (parentRange.Index >= movingRange.Index &&
            parentRange.Index < movingRange.Index + movingRange.Count)
        {
            return false;
        }

        var movingElements = BoardElements
            .Skip(movingRange.Index)
            .Take(movingRange.Count)
            .ToArray();

        if (movingElements.Length == 0)
        {
            return false;
        }

        var levelDelta = targetLevel - movingElements[0].Level;
        if (levelDelta != 0)
        {
            foreach (var movingElement in movingElements)
            {
                movingElement.Level = Math.Max(0, movingElement.Level + levelDelta);
            }
        }

        var targetIndex = parentRange.Index + parentRange.Count;
        for (var index = 0; index < movingRange.Count; index++)
        {
            BoardElements.RemoveAt(movingRange.Index);
        }

        if (targetIndex > movingRange.Index)
        {
            targetIndex -= movingRange.Count;
        }

        targetIndex = Math.Clamp(targetIndex, 0, BoardElements.Count);
        for (var index = 0; index < movingElements.Length; index++)
        {
            BoardElements.Insert(targetIndex + index, movingElements[index]);
        }

        SelectedBoardElement = movingElements[0];
        UpdateBoardSyncRequiredState();
        return true;
    }

    private void ExecuteSyncBoardToScript(object? _)
    {
        _isInternalSync = true;
        try
        {
            var parsed = GetCurrentParsedSnapshot();
            var syncPlan = BuildBoardSyncPlan(parsed);
            if (syncPlan is null)
            {
                UpdateBoardSyncRequiredState(parsed.Elements);
                return;
            }

            // Apply state changes
            _lineTypeOverrides.Clear();
            foreach (var entry in syncPlan.LineTypeOverrides)
            {
                _lineTypeOverrides[entry.Key] = entry.Value;
            }

            _automaticActionLineOverrides.Clear();
            foreach (var lineNumber in syncPlan.AutomaticActionOverrides)
            {
                _automaticActionLineOverrides.Add(lineNumber);
            }

            // Update the document text. This will trigger a re-parse internally.
            DocumentText = ScriptSanitizer.CollapseTripleNewlines(syncPlan.DocumentText);
            IsBoardSyncRequired = false;
            
            // Ensure the UI is updated immediately after the sync.
            CommandManager.InvalidateRequerySuggested();
        }
        finally
        {
            _isInternalSync = false;
            RefreshParsedSnapshotNow();
        }
    }

    private bool CanSyncBoardToScript(object? _)
    {
        return IsBoardSyncRequired;
    }

    private void ExecuteToggleBoardElementCollapsed(object? parameter)
    {
        if (parameter is not ScreenplayElement element)
        {
            return;
        }

        _ = TrySetBoardElementCollapsed(element, !element.IsCollapsed);
    }

    private void ExecuteExpandAllBoardCards(object? _)
    {
        _ = TrySetAllBoardElementsCollapsed(isCollapsed: false);
    }

    private void ExecuteCollapseAllBoardCards(object? _)
    {
        _ = TrySetAllBoardElementsCollapsed(isCollapsed: true);
    }

    private void ExecuteExpandAllOutlineNodes(object? _)
    {
        _ = TrySetAllOutlineNodesExpanded(isExpanded: true);
    }

    private void ExecuteCollapseAllOutlineNodes(object? _)
    {
        _ = TrySetAllOutlineNodesExpanded(isExpanded: false);
    }

    private void ExecuteCreateNewCard(object? _)
    {
        var newCard = CreateBoardDraftElement(
            type: ScreenplayElementType.SceneHeading,
            level: 2,
            heading: DefaultNewBoardCardHeading,
            sceneHeading: DefaultNewBoardCardHeading,
            description: DefaultNewBoardCardDescription);
        BoardElements.Insert(0, newCard);
        SelectedBoardElement = newCard;
        UpdateBoardSyncRequiredState();
    }

    public bool TryDeleteBoardBlock(ScreenplayElement? element)
    {
        if (element is null || BoardElements.Count == 0)
        {
            return false;
        }

        StoryBlockRange blockRange;
        try
        {
            blockRange = StoryHierarchyHelper.GetBlockRange(BoardElements, element);
        }
        catch (ArgumentException)
        {
            return false;
        }

        for (var index = 0; index < blockRange.Count; index++)
        {
            BoardElements.RemoveAt(blockRange.Index);
        }

        if (BoardElements.Count == 0)
        {
            SelectedBoardElement = null;
        }
        else
        {
            var selectionIndex = Math.Clamp(blockRange.Index, 0, BoardElements.Count - 1);
            SelectedBoardElement = BoardElements[selectionIndex];
        }

        ClearBoardDropIndicator();
        ExecuteSyncBoardToScript(null);
        return true;
    }

    public bool TryMoveBoardBlock(ScreenplayElement? element, int targetIndex)
    {
        if (element is null || BoardElements.Count == 0)
        {
            return false;
        }

        StoryBlockRange blockRange;
        try
        {
            blockRange = StoryHierarchyHelper.GetBlockRange(BoardElements, element);
        }
        catch (ArgumentException)
        {
            return false;
        }

        targetIndex = Math.Clamp(targetIndex, 0, BoardElements.Count);
        if (targetIndex >= blockRange.Index && targetIndex <= blockRange.Index + blockRange.Count)
        {
            return false;
        }

        var movingElements = BoardElements
            .Skip(blockRange.Index)
            .Take(blockRange.Count)
            .ToArray();

        for (var index = 0; index < blockRange.Count; index++)
        {
            BoardElements.RemoveAt(blockRange.Index);
        }

        if (targetIndex > blockRange.Index)
        {
            targetIndex -= blockRange.Count;
        }

        targetIndex = Math.Clamp(targetIndex, 0, BoardElements.Count);
        for (var index = 0; index < movingElements.Length; index++)
        {
            BoardElements.Insert(targetIndex + index, movingElements[index]);
        }

        UpdateBoardSyncRequiredState();
        return true;
    }

    public void UpdateCaretContext(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        CurrentLineNumber = lineNumber;

        var elementType = GetEffectiveLineType(lineNumber);
        CurrentElementType = elementType.ToString();
        CurrentElementText = GetCurrentElementDescription(elementType, isTabOverride: false);
    }

    public OutlineNodeViewModel? FindActiveOutlineNode(int lineNumber)
    {
        return FindActiveOutlineNodeInCollection(OutlineRoots, lineNumber);
    }

    private static OutlineNodeViewModel? FindActiveOutlineNodeInCollection(IEnumerable<OutlineNodeViewModel> nodes, int lineNumber)
    {
        OutlineNodeViewModel? bestMatch = null;

        foreach (var node in nodes)
        {
            if (node.LineNumber <= lineNumber)
            {
                // If it's a closer match than the current best, take it
                if (bestMatch == null || node.LineNumber >= bestMatch.LineNumber)
                {
                    bestMatch = node;
                }

                // Check children for even better match
                var childMatch = FindActiveOutlineNodeInCollection(node.Children, lineNumber);
                if (childMatch != null)
                {
                    bestMatch = childMatch;
                }
            }
        }

        return bestMatch;
    }

    public void UpdateEnterContinuation(int lineNumber, string currentLineText)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        var currentElementType = GetLatestEffectiveLineType(lineNumber);
        var continuation = DetermineEnterContinuation(currentElementType, currentLineText ?? string.Empty);
        EnterContinuationText = continuation.ToString();
    }

    public void CycleCurrentLineType(bool forward)
    {
        var lineNumber = CurrentLineNumber < 1 ? 1 : CurrentLineNumber;
        var currentType = GetEffectiveLineType(lineNumber);
        var cycledType = CycleTabType(currentType, forward);

        SetLineTypeOverrideCore(lineNumber, cycledType);

        CurrentLineNumber = lineNumber;
        CurrentElementType = cycledType.ToString();
        CurrentElementText = GetCurrentElementDescription(cycledType, isTabOverride: true);
        UpdateEnterContinuation(lineNumber, GetLineText(lineNumber));
        ScheduleOutlineRefresh();
    }

    private void UpdateOutline()
    {
        RefreshParsedDocument();
    }

    private void ScheduleOutlineRefresh()
    {
        _outlineRefreshVersion++;
        _outlineRefreshPending = true;
        _refreshTimer.Stop();
        _refreshTimer.Interval = GetRefreshDelay();
        _refreshTimer.Start();
    }

    private async void RefreshParsedDocument()
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
        var snapshotText = GetScriptSourceText();
        var snapshotLineTypeOverrides = GetLineTypeOverridesSnapshot();
        var parseStartedAt = Environment.TickCount64;

        try
        {
            var parsed = await Task.Run(() => _parser.Parse(snapshotText, snapshotLineTypeOverrides));

            if (version == _outlineRefreshVersion)
            {
                _lastParseDuration = TimeSpan.FromMilliseconds(Math.Max(0, Environment.TickCount64 - parseStartedAt));
                ApplyParsedDocument(parsed);
            }
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

    private void GoalTimerRuntime_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!ReferenceEquals(sender, _goalTimerRuntime))
        {
            return;
        }

        if (e.PropertyName is not nameof(GoalTimerRuntime.Snapshot)
            and not nameof(GoalTimerRuntime.State)
            and not nameof(GoalTimerRuntime.IsCompleted))
        {
            return;
        }

        RefreshGoalState();
    }

    private void ApplyParsedDocument(ParsedScreenplay parsed)
    {
        _lastParsed = parsed;

        var expandedOutline = CaptureExpandedIdentifiers(OutlineRoots);
        var expandedNotes = CaptureExpandedIdentifiers(NotesRoots);

        PopulateBodyText(parsed.Elements);

        var outlineRoots = BuildOutlineTree(parsed.Elements, expandedOutline);
        var noteRoots = BuildNotesTree(parsed.Elements, expandedNotes);
        var boardElements = BuildBoardElements(parsed.Elements, BoardElements, IsBoardSyncRequired);
        var titlePageEntries = parsed.TitlePage.Entries.ToArray();
        var previewElements = parsed.Elements
            .Select(element => new PreviewElementItem(
                element.Type,
                element.Text,
                element.RawText,
                element.StartLine,
                element.EndLine))
            .ToArray();
        var changed = false;

        if (!OutlineTreesEqual(OutlineRoots, outlineRoots))
        {
            OutlineRoots = new ObservableCollection<OutlineNodeViewModel>(outlineRoots);
            OnPropertyChanged(nameof(OutlineRoots));
            changed = true;
        }

        if (!OutlineTreesEqual(NotesRoots, noteRoots))
        {
            NotesRoots = new ObservableCollection<OutlineNodeViewModel>(noteRoots);
            OnPropertyChanged(nameof(NotesRoots));
            changed = true;
        }

        if (!ScreenplayElementsSequenceEqual(BoardElements, boardElements))
        {
            BoardElements = new ObservableCollection<ScreenplayElement>(boardElements);
            AttachBoardElementsCollection(BoardElements);
            if (_selectedBoardElement is not null)
            {
                SelectedBoardElement = BoardElements.FirstOrDefault(candidate => candidate.Id == _selectedBoardElement.Id);
            }
            OnPropertyChanged(nameof(BoardElements));
            changed = true;
        }

        if (!TitlePageEntriesEqual(TitlePageEntries, titlePageEntries))
        {
            TitlePageEntries = new ObservableCollection<TitlePageEntry>(titlePageEntries);
            OnPropertyChanged(nameof(TitlePageEntries));
            changed = true;
        }

        if (!PreviewElementsSequenceEqual(PreviewElements, previewElements))
        {
            PreviewElements = new ObservableCollection<PreviewElementItem>(previewElements);
            OnPropertyChanged(nameof(PreviewElements));
            changed = true;
        }

        UpdateBoardSyncRequiredState(parsed.Elements);

        if (changed)
        {
            OutlineUpdated?.Invoke(this, EventArgs.Empty);
        }

        RefreshGoalState();
    }

    private TimeSpan GetRefreshDelay()
    {
        var length = _documentText.Length;

        if (length >= 100_000 || _lastParseDuration >= TimeSpan.FromMilliseconds(40))
        {
            return TimeSpan.FromMilliseconds(350);
        }

        if (length >= 25_000)
        {
            return TimeSpan.FromMilliseconds(250);
        }

        return TimeSpan.FromMilliseconds(180);
    }

    private void RefreshGoalState()
    {
        CaptureSessionGoalBaselineIfNeeded();

        var currentWordCount = _goalProgressCalculator.CalculateWordCount(DocumentText);
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
        GoalTimerElapsedText = "00:00";
        GoalTimerRemainingText = "00:00";
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
        GoalTimerElapsedText = "00:00";
        GoalTimerRemainingText = "00:00";
    }

    private void RefreshTimerGoal()
    {
        EnsureGoalTimerRuntime();

        if (_goalTimerRuntime is null)
        {
            GoalTargetUnitLabel = FormatCountLabel(_timerGoalTargetMinutes, "Minute", "Minutes");
            GoalCurrentDisplayText = "00:00";
            GoalTargetDisplayText = FormatDuration(TimeSpan.FromMinutes(_timerGoalTargetMinutes));
            GoalStateText = "Idle";
            GoalProgressPercent = 0;
            GoalTimerElapsedText = "00:00";
            GoalTimerRemainingText = FormatDuration(TimeSpan.FromMinutes(_timerGoalTargetMinutes));
            return;
        }

        var targetDuration = _goalTimerRuntime.TargetDuration;
        var elapsed = _goalTimerRuntime.ElapsedTime;
        var remaining = _goalTimerRuntime.RemainingTime;
        var completed = _goalTimerRuntime.IsCompleted;

        GoalTargetUnitLabel = FormatCountLabel(_timerGoalTargetMinutes, "Minute", "Minutes");
        GoalCurrentDisplayText = FormatDuration(elapsed);
        GoalTargetDisplayText = FormatDuration(targetDuration);
        GoalStateText = _goalTimerRuntime.State.ToString();
        GoalProgressPercent = CalculateProgressPercent(elapsed.TotalSeconds, targetDuration.TotalSeconds);
        GoalTimerElapsedText = FormatDuration(elapsed);
        GoalTimerRemainingText = FormatDuration(remaining);

        if (completed)
        {
            GoalStateText = "Completed";
        }
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
        if (!_sessionGoalBaselineNeedsCapture)
        {
            return;
        }

        CaptureSessionGoalBaseline();
    }

    private void CaptureSessionGoalBaseline()
    {
        var parsed = _parser.Parse(DocumentText, GetLineTypeOverridesSnapshot());
        _lastParsed = parsed;
        _sessionWordCountBaseline = _goalProgressCalculator.CalculateWordCount(DocumentText);
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

        if (_goalTimerRuntime is not null && _goalTimerRuntime.TargetDuration == targetDuration)
        {
            return;
        }

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
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

        EnsureGoalTimerRuntime();
        RefreshGoalState();
    }

    public void StartGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

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
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

        _goalTimerRuntime?.Pause();
        RefreshGoalState();
    }

    public void ResumeGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

        EnsureGoalTimerRuntime();
        _goalTimerRuntime?.Resume();
        RefreshGoalState();
    }

    public void StopGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

        _goalTimerRuntime?.Stop();
        RefreshGoalState();
    }

    public void ResetGoalTimer()
    {
        if (SessionSelectedGoalType != GoalType.Timer)
        {
            return;
        }

        EnsureGoalTimerRuntime();
        _goalTimerRuntime?.Reset();
        RefreshGoalState();
    }

    private static double CalculateProgressPercent(double currentValue, double targetValue)
    {
        if (targetValue <= 0)
        {
            return 100;
        }

        return Math.Min(100, Math.Max(0, (currentValue / targetValue) * 100));
    }

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration < TimeSpan.Zero)
        {
            duration = TimeSpan.Zero;
        }

        if (duration.TotalHours >= 1)
        {
            return duration.ToString(@"h\:mm\:ss");
        }

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

    private static void AddOutlineExpansionKeys(
        ISet<string> expanded,
        OutlineNodeKind kind,
        int lineNumber,
        int? sectionLevel,
        string text)
    {
        // Prefer source position so expansion survives inline text edits, but keep a text key so
        // expanded nodes still survive when lines above them move.
        expanded.Add(BuildOutlineExpansionKey(kind, lineNumber, sectionLevel));
        expanded.Add(BuildOutlineLegacyExpansionKey(kind, text));
    }

    private static bool HasExpandedOutlineKey(
        ISet<string>? expandedKeys,
        OutlineNodeKind kind,
        int lineNumber,
        int? sectionLevel,
        string text)
    {
        if (expandedKeys is null)
        {
            return false;
        }

        return expandedKeys.Contains(BuildOutlineExpansionKey(kind, lineNumber, sectionLevel)) ||
            expandedKeys.Contains(BuildOutlineLegacyExpansionKey(kind, text));
    }

    private static string BuildOutlineExpansionKey(
        OutlineNodeKind kind,
        int lineNumber,
        int? sectionLevel)
    {
        return $"{kind}_{lineNumber}_{sectionLevel ?? -1}";
    }

    private static string BuildOutlineLegacyExpansionKey(OutlineNodeKind kind, string text)
    {
        return $"{kind}_{text}";
    }

    private static IReadOnlyList<OutlineNodeViewModel> BuildOutlineTree(
        IReadOnlyList<ScreenplayElement> elements,
        ISet<string>? expandedKeys = null)
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
                        section.BodyText);

                    if (HasExpandedOutlineKey(
                            expandedKeys,
                            OutlineNodeKind.Section,
                            section.StartLine,
                            section.SectionDepth,
                            section.Text))
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
                        sceneHeading.BodyText);

                    if (HasExpandedOutlineKey(
                            expandedKeys,
                            OutlineNodeKind.SceneHeading,
                            sceneHeading.StartLine,
                            sectionLevel: null,
                            sceneHeading.Text))
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
        ISet<string>? expandedKeys = null)
    {
        var roots = new List<OutlineNodeViewModel>();

        foreach (var element in elements)
        {
            if (element is not NoteElement note)
            {
                continue;
            }

            var noteNode = new OutlineNodeViewModel(
                OutlineNodeKind.Note,
                note.Text,
                note.StartLine,
                sectionLevel: null,
                note.BodyText);

            if (HasExpandedOutlineKey(
                    expandedKeys,
                    OutlineNodeKind.Note,
                    note.StartLine,
                    sectionLevel: null,
                    note.Text))
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
            
            // Only sections and headings gather a "Body" (which we treat as their description)
            if (element.Type is ScreenplayElementType.Section or ScreenplayElementType.SceneHeading or ScreenplayElementType.Note)
            {
                var bodyLines = new List<string>();
                for (int j = i + 1; j < elementList.Count; j++)
                {
                    var next = elementList[j];
                    
                    // Gathers ONLY SYNOPSIS lines into the parent card's description body.
                    // This creates an "Owned Synopsis Block".
                    if (next.Type is ScreenplayElementType.Synopsis)
                    {
                        var lineText = StripGuidComment(next.Text);
                        bodyLines.Add(lineText);
                        next.IsSuppressed = true;
                    }
                    else
                    {
                        // Stop at any other element type (Actions, Lyrics, etc. break the owned block)
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
        
        // Matches [[id:guid]] pattern
        var idMatch = System.Text.RegularExpressions.Regex.Match(text, @"\s*\[\[id:[a-f\d\-]+\]\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (idMatch.Success)
        {
            return text.Replace(idMatch.Value, "").Trim();
        }
        return text;
    }

    private static bool OutlineTreesEqual(
        IReadOnlyList<OutlineNodeViewModel> current,
        IReadOnlyList<OutlineNodeViewModel> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            if (!OutlineNodesEqual(current[i], next[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool OutlineNodesEqual(OutlineNodeViewModel left, OutlineNodeViewModel right)
    {
        if (left.Kind != right.Kind ||
            left.Text != right.Text ||
            left.BodyText != right.BodyText ||
            left.LineNumber != right.LineNumber ||
            left.SectionLevel != right.SectionLevel ||
            left.Children.Count != right.Children.Count)
        {
            return false;
        }

        for (var i = 0; i < left.Children.Count; i++)
        {
            if (!OutlineNodesEqual(left.Children[i], right.Children[i]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool SetOutlineExpansionState(
        IEnumerable<OutlineNodeViewModel> nodes,
        bool isExpanded)
    {
        var changed = false;

        foreach (var node in nodes)
        {
            if (node.IsExpanded != isExpanded)
            {
                node.IsExpanded = isExpanded;
                changed = true;
            }

            if (node.Children.Count > 0 &&
                SetOutlineExpansionState(node.Children, isExpanded))
            {
                changed = true;
            }
        }

        return changed;
    }

    private static bool PreviewElementsSequenceEqual(
        IReadOnlyList<PreviewElementItem> current,
        IReadOnlyList<PreviewElementItem> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            var left = current[i];
            var right = next[i];

            if (left.ElementType != right.ElementType ||
                left.Text != right.Text ||
                left.RawText != right.RawText ||
                left.StartLine != right.StartLine ||
                left.EndLine != right.EndLine)
            {
                return false;
            }
        }

        return true;
    }

    private static IReadOnlyList<ScreenplayElement> BuildBoardElements(
        IReadOnlyList<ScreenplayElement> parsedElements,
        IReadOnlyList<ScreenplayElement> currentBoardElements,
        bool preserveCustomLayout)
    {
        // Use persistent ID matching to survive index shifts in the script
        var planningElements = parsedElements
            .Where(ShouldIncludeBoardElement)
            .ToArray();

        // Build a lookup of existing cards by ID (Guid)
        var existingById = currentBoardElements
            .GroupBy(e => e.Id)
            .ToDictionary(g => g.Key, g => new Queue<ScreenplayElement>(g));

        var boardElements = new List<ScreenplayElement>(planningElements.Length);
        var matchedExisting = new HashSet<ScreenplayElement>();

        foreach (var parsedElement in planningElements)
        {
            // Match via persistent Guid
            if (existingById.TryGetValue(parsedElement.Id, out var arrivals) && arrivals.Count > 0)
            {
                var existing = arrivals.Dequeue();
                boardElements.Add(MergeParsedBoardElement(parsedElement, existing));
                matchedExisting.Add(existing);
                continue;
            }

            // Fallback for elements without IDs in the script (e.g. manually added Fountain lines)
            boardElements.Add(CloneElement(parsedElement));
        }

        // If we want to preserve cards that were "deleted" from the script but exist in UI as drafts
        if (preserveCustomLayout)
        {
            foreach (var existing in currentBoardElements)
            {
                if (existing.IsDraft && !matchedExisting.Contains(existing))
                {
                    boardElements.Add(CloneElement(existing));
                }
            }
        }

        return boardElements;
    }

    private BoardSyncPlan? BuildBoardSyncPlan(ParsedScreenplay parsed)
    {
        // ── Non-Destructive ID-Anchored Merge ───────────────────────────────────────
        // Strategy:
        //   1. Build a Dictionary<Guid, int> mapping each [[id:GUID]] in the document to its line index.
        //   2. Build a Dictionary<Guid, List<string>> capturing the "block" of content
        //      (dialogue, action, parentheticals) that follows each card's anchor line.
        //   3. Reconstruct the document in board order:
        //      leading lines + for each BoardCard: [serialized headline] + [block]
        //
        // Preservation guarantee: blocks always travel with their owning card. Moving
        // a card on the beat board moves all its nested content in the script.
        // ───────────────────────────────────────────────────────────────────────────

        var boardElementsFromScript = parsed.Elements
            .Where(ShouldIncludeBoardElement)
            .ToArray();

        if (BoardElements.Count == 0 && boardElementsFromScript.Length == 0)
        {
            return null;
        }

        var normalizedSourceText = (DocumentText ?? string.Empty).ReplaceLineEndings("\n");
        var sourceLines = normalizedSourceText.Split('\n', StringSplitOptions.None);
        if (sourceLines.Length == 0)
        {
            sourceLines = [string.Empty];
        }

        // Step 1: Build the GUID → line-index map directly from raw document text.
        // This is independent of the parser and is always fresh for every sync.
        var idToLineIndex = BuildIdToLineMap(sourceLines);

        // Build a set of GUIDs that are referenced by the current board.
        var activeIds = new HashSet<Guid>(BoardElements.Select(e => e.Id));

        // Signature fallback: for elements that don't yet have an ID in the script 
        // (e.g. first sync), map by signature so we can capture their block.
        var signatureToId = new Dictionary<BoardElementSignature, Guid>();
        foreach (var card in BoardElements)
        {
            signatureToId[CreateBoardElementSignature(card)] = card.Id;
        }

        // Step 2: Capture blocks. Walk the parsed board elements in source order.
        // For each card that is active on the board, capture the trailing lines
        // (from after its headline to the start of the next card's headline).
        var segmentLookup = new Dictionary<Guid, List<string>>();
        var segmentLineInfos = new Dictionary<Guid, (int Start, int End)>();
        var propagationBuffer = new List<string>();

        // Lines that appear before the first board element are always preserved verbatim.
        var initialLineCount = boardElementsFromScript.Length > 0
            ? boardElementsFromScript[0].LineIndex
            : sourceLines.Length;
        var synchronizedLines = new List<string>(sourceLines.Length + 16);
        synchronizedLines.AddRange(sourceLines.Take(initialLineCount));

        // Pre-map suppressed elements to their owners to identify owned synopsis blocks
        var suppressedBlockEndLines = new Dictionary<Guid, int>();
        var elements = parsed.Elements;
        for (int i = 0; i < elements.Count; i++)
        {
            var element = elements[i];
            if (element.Type is ScreenplayElementType.Section or ScreenplayElementType.SceneHeading or ScreenplayElementType.Note)
            {
                int blockEndLine = element.EndLineIndex;
                for (int j = i + 1; j < elements.Count; j++)
                {
                    if (elements[j].IsSuppressed && elements[j].Type == ScreenplayElementType.Synopsis)
                    {
                        blockEndLine = elements[j].EndLineIndex;
                    }
                    else
                    {
                        break;
                    }
                }
                suppressedBlockEndLines[element.Id] = blockEndLine;
            }
        }

        for (var index = 0; index < boardElementsFromScript.Length; index++)
        {
            var element = boardElementsFromScript[index];
            var nextElementLineIndex = index + 1 < boardElementsFromScript.Length
                ? boardElementsFromScript[index + 1].LineIndex
                : sourceLines.Length;

            // Use the pre-calculated consolidated block end to ensure we skip the ENTIRE owned synopsis block
            var blockEndLine = suppressedBlockEndLines.TryGetValue(element.Id, out var end) ? end : element.EndLineIndex;
            var trailingStart = Math.Clamp(blockEndLine + 1, 0, sourceLines.Length);
            var trailingEnd = Math.Clamp(nextElementLineIndex, trailingStart, sourceLines.Length);

            var originalTrailing = sourceLines[trailingStart..trailingEnd];

            // Clean trailing content: avoid preserving more than one trailing blank line at the end 
            // of a block, which prevents compounding whitespace expansion.
            var cleanedTrailing = new List<string>(originalTrailing);
            while (cleanedTrailing.Count > 1 && 
                   string.IsNullOrWhiteSpace(cleanedTrailing[^1]) && 
                   string.IsNullOrWhiteSpace(cleanedTrailing[^2]))
            {
                cleanedTrailing.RemoveAt(cleanedTrailing.Count - 1);
            }

            // Prefer the parsed element's embedded GUID; fall back to the direct id-map scan;
            // then fall back to signature matching for elements that lack IDs.
            var targetId = element.Id;
            if (!activeIds.Contains(targetId))
            {
                if (idToLineIndex.ContainsKey(targetId))
                {
                    // ID is in the document but doesn't match any board card — skip this element.
                }
                else if (signatureToId.TryGetValue(CreateBoardElementSignature(element), out var matchedId))
                {
                    targetId = matchedId;
                }
            }

            if (activeIds.Contains(targetId))
            {
                // Card exists on the board. Carry any propagated orphan lines into its block.
                var content = new List<string>(propagationBuffer);
                content.AddRange(cleanedTrailing);
                segmentLookup[targetId] = content;
                segmentLineInfos[targetId] = (element.LineIndex, nextElementLineIndex);
                propagationBuffer.Clear();
            }
            else
            {
                // Card was deleted from the board or is unmatched — preserve its trailing
                // content in the propagation buffer so it can attach to the next active card.
                propagationBuffer.AddRange(cleanedTrailing);
            }
        }

        // Attach any trailing orphan lines to the last active card.
        if (propagationBuffer.Count > 0 && BoardElements.Count > 0)
        {
            var lastActiveId = BoardElements.Last().Id;
            if (segmentLookup.ContainsKey(lastActiveId))
            {
                segmentLookup[lastActiveId].AddRange(propagationBuffer);
            }
            else
            {
                segmentLookup[lastActiveId] = [.. propagationBuffer];
            }
            propagationBuffer.Clear();
        }

        // Step 3: Reconstruct in board order.
        var synchronizedLineTypeOverrides = new Dictionary<int, ScreenplayElementType>();
        var synchronizedAutomaticActionOverrides = new HashSet<int>();
        CopyLineTypeOverrides(1, initialLineCount, 1, synchronizedLineTypeOverrides, synchronizedAutomaticActionOverrides);

        var currentOutputLineIndex = initialLineCount;

        foreach (var boardElement in BoardElements)
        {
            // Ensure exactly one blank line precedes every card block (Act, Sequence, Scene, Note).
            if (currentOutputLineIndex > 0)
            {
                if (synchronizedLines.Count == 0 || !string.IsNullOrWhiteSpace(synchronizedLines[^1]))
                {
                    synchronizedLines.Add(string.Empty);
                    currentOutputLineIndex++;
                }
            }

            var outputLines = BuildBoardElementSourceLines(boardElement);
            synchronizedLines.AddRange(outputLines);

            if (segmentLookup.TryGetValue(boardElement.Id, out var trailing))
            {
                synchronizedLines.AddRange(trailing);

                if (segmentLineInfos.TryGetValue(boardElement.Id, out var info))
                {
                    CopyLineTypeOverrides(
                        originalStartLineNumber: info.Start + 1,
                        originalEndLineNumber: info.End,
                        newStartLineNumber: currentOutputLineIndex + 1,
                        synchronizedLineTypeOverrides,
                        synchronizedAutomaticActionOverrides);
                }

                currentOutputLineIndex += outputLines.Length + trailing.Count;
            }
            else
            {
                currentOutputLineIndex += outputLines.Length;
            }
        }

        // Any remaining orphan lines (e.g. everything was deleted) are appended at the end.
        if (propagationBuffer.Count > 0)
        {
            synchronizedLines.AddRange(propagationBuffer);
        }

        var lineEnding = DetectPreferredLineEnding(DocumentText);
        return new BoardSyncPlan(
            string.Join(lineEnding, synchronizedLines),
            synchronizedLineTypeOverrides,
            synchronizedAutomaticActionOverrides);
    }

    /// <summary>
    /// Builds a map from each GUID found in the document lines (via [[id:GUID]] tags)
    /// to the line index where it appears. This is a fresh scan on every sync call,
    /// so it is always consistent with the current document state.
    /// </summary>
    private static readonly System.Text.RegularExpressions.Regex IdLineRegex =
        new(@"\[\[id:([a-f\d\-]+)\]\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    private static Dictionary<Guid, int> BuildIdToLineMap(string[] lines)
    {
        var map = new Dictionary<Guid, int>(lines.Length / 4 + 1);
        for (var i = 0; i < lines.Length; i++)
        {
            var match = IdLineRegex.Match(lines[i]);
            if (match.Success && Guid.TryParse(match.Groups[1].Value, out var id))
            {
                // First occurrence wins — handles duplicate IDs gracefully.
                map.TryAdd(id, i);
            }
        }
        return map;
    }

    private void CopyLineTypeOverrides(
        int originalStartLineNumber,
        int originalEndLineNumber,
        int newStartLineNumber,
        IDictionary<int, ScreenplayElementType> synchronizedLineTypeOverrides,
        ISet<int> synchronizedAutomaticActionOverrides)
    {
        if (originalEndLineNumber < originalStartLineNumber)
        {
            return;
        }

        foreach (var entry in _lineTypeOverrides)
        {
            if (entry.Key < originalStartLineNumber || entry.Key > originalEndLineNumber)
            {
                continue;
            }

            var offset = entry.Key - originalStartLineNumber;
            synchronizedLineTypeOverrides[newStartLineNumber + offset] = entry.Value;
        }

        foreach (var lineNumber in _automaticActionLineOverrides)
        {
            if (lineNumber < originalStartLineNumber || lineNumber > originalEndLineNumber)
            {
                continue;
            }

            var offset = lineNumber - originalStartLineNumber;
            synchronizedAutomaticActionOverrides.Add(newStartLineNumber + offset);
        }
    }

    private static string[] BuildBoardElementSourceLines(ScreenplayElement boardElement)
    {
        return SerializeBoardElement(boardElement)
            .ReplaceLineEndings("\n")
            .Split('\n', StringSplitOptions.None);
    }

    private static string SerializeBoardElement(ScreenplayElement boardElement)
    {
        return boardElement switch
        {
            ScratchpadCardElement draftElement when draftElement.IsDraft
                => SerializeDraftBoardElement(
                    draftElement.Id,
                    draftElement.Type,
                    draftElement.Level,
                    draftElement.Heading,
                    draftElement.ScriptHeading,
                    draftElement.Description),
            SectionElement section  => SerializeSectionElement(section, boardElement.Level),
            SceneHeadingElement sceneHeading => SerializeSceneHeadingElement(sceneHeading),
            _                       => AppendIdComment(boardElement.RawText, boardElement.Id)
        };
    }

    private static string SerializeSceneHeadingElement(SceneHeadingElement sceneHeading)
    {
        var lines = new List<string>();
        
        // Step A: Heading + ID Line
        var headingText = (sceneHeading.ScriptHeading ?? sceneHeading.Text ?? "Scene").Trim();
        lines.Add($"{headingText} [[id:{sceneHeading.Id}]]");

        // Step B: Synopsis (Description)
        var description = NormalizeBoardCardDescription(sceneHeading.BodyText);
        if (HasMeaningfulBoardDescription(description))
        {
            foreach (var line in SplitDescriptionLines(description))
            {
                lines.Add($"= {line}");
            }
        }
        return string.Join("\n", lines);
    }

    private static string SerializeSynopsisElement(SynopsisElement synopsis)
    {
        var trimmed = synopsis.Text.Trim();
        var baseLine = trimmed.Length == 0 ? "=" : $"= {trimmed}";
        // Only append ID if not already present (idempotent)
        return baseLine.Contains("[[id:", StringComparison.OrdinalIgnoreCase)
            ? baseLine
            : $"{baseLine} [[id:{synopsis.Id}]]";
    }

    private static string AppendIdComment(string rawText, Guid id)
    {
        if (string.IsNullOrWhiteSpace(rawText)) return rawText;
        if (rawText.Contains("[[id:", StringComparison.OrdinalIgnoreCase)) return rawText;
        
        var lines = rawText.Split('\n');
        lines[0] = lines[0].TrimEnd() + $" [[id:{id}]]";
        return string.Join("\n", lines);
    }

    private static string SerializeSectionElement(SectionElement section, int level)
    {
        var depth = Math.Max(1, level + 1);
        var prefix = new string('#', depth);
        
        // Step A: Heading + ID Line
        // Ensure a blank line precedes every card block for Fountain compliance.
        var lines = new List<string>();
        
        // Use section.Text explicitly for the title.
        var title = (section.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(title)) title = "Section";
        
        var headingLine = $"{prefix} {title} [[id:{section.Id}]]";
        lines.Add(headingLine);

        // Step B: Synopsis (Description)
        var description = NormalizeBoardCardDescription(section.BodyText);
        if (HasMeaningfulBoardDescription(description))
        {
            foreach (var line in SplitDescriptionLines(description))
            {
                lines.Add($"= {line}");
            }
        }
        
        return string.Join("\n", lines);
    }

    private static string SerializeDraftBoardElement(
        Guid id,
        ScreenplayElementType type,
        int level,
        string heading,
        string sceneHeading,
        string description)
    {
        return type switch
        {
            ScreenplayElementType.Section => SerializeDraftSection(id, level, heading, description),
            ScreenplayElementType.SceneHeading => SerializeDraftSceneHeading(id, sceneHeading, description),
            ScreenplayElementType.Note => SerializeDraftNote(id, heading, description),
            _ => SerializeDraftSection(id, level, heading, description)
        };
    }

    private static string SerializeDraftSection(Guid id, int level, string heading, string description)
    {
        var depth = Math.Max(1, level + 1);
        var prefix = new string('#', depth);
        
        // Step A: Title + ID Line
        var lines = new List<string>();
        var title = ResolveBoardScriptText(heading, fallback: "Section");
        lines.Add($"{prefix} {title} [[id:{id}]]");

        // Step B: Synopsis (Description)
        var trimmedDescription = NormalizeBoardCardDescription(description);
        if (HasMeaningfulBoardDescription(trimmedDescription))
        {
            foreach (var line in SplitDescriptionLines(trimmedDescription))
            {
                lines.Add($"= {line}");
            }
        }

        return string.Join("\n", lines);
    }

    private static string SerializeDraftSceneHeading(Guid id, string sceneHeading, string description)
    {
        var lines = new List<string>();
        var trimmedHeading = NormalizeBoardCardSceneHeading(null, sceneHeading);
        var trimmedDescription = NormalizeBoardCardDescription(description);

        // Step A: Heading + ID Line
        string outputHeading;
        if (trimmedHeading.StartsWith(".", StringComparison.Ordinal))
        {
            outputHeading = trimmedHeading;
        }
        else if (TextAnalysis.LooksLikeSceneHeadingStart(trimmedHeading.AsSpan()))
        {
            outputHeading = trimmedHeading;
        }
        else
        {
            outputHeading = $". {trimmedHeading}";
        }
        
        lines.Add($"{outputHeading} [[id:{id}]]");

        // Step B: Synopsis (Description)
        if (HasMeaningfulBoardDescription(trimmedDescription))
        {
            foreach (var line in SplitDescriptionLines(trimmedDescription))
            {
                lines.Add($"= {line}");
            }
        }

        return string.Join("\n", lines);
    }

    private static string SerializeDraftNote(Guid id, string heading, string description)
    {
        var lines = new List<string>();
        var trimmedHeading = ResolveBoardScriptText(heading, fallback: "Note");
        var trimmedDescription = NormalizeBoardCardDescription(description);

        // Step A: Header + ID line
        lines.Add($"[[{trimmedHeading}]] [[id:{id}]]");

        // Step B: Synopsis (Description)
        if (HasMeaningfulBoardDescription(trimmedDescription))
        {
            foreach (var line in SplitDescriptionLines(trimmedDescription))
            {
                lines.Add($"= {line}");
            }
        }

        return string.Join("\n", lines);
    }

    private static string ResolveBoardScriptText(string heading, string fallback)
    {
        var trimmedHeading = (heading ?? string.Empty).Trim();
        return trimmedHeading.Length > 0 ? trimmedHeading : fallback;
    }

    private static IEnumerable<string> SplitDescriptionLines(string description)
    {
        return (description ?? string.Empty)
            .ReplaceLineEndings("\n")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => line.Length > 0);
    }

    private static bool ShouldOmitDraftBoardDescription(string description, string heading)
    {
        return string.IsNullOrWhiteSpace(description) ||
            IsDefaultBoardDescription(description) ||
            string.Equals(description.Trim(), heading?.Trim(), StringComparison.Ordinal);
    }

    private static string DetectPreferredLineEnding(string? text)
    {
        return (text ?? string.Empty).Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : Environment.NewLine;
    }

    private static bool ShouldIncludeBoardElement(ScreenplayElement element)
    {
        return !element.IsSuppressed && 
            (element.Type is ScreenplayElementType.Section
            or ScreenplayElementType.SceneHeading
            or ScreenplayElementType.Note);
    }

    private static ScreenplayElement CloneElement(ScreenplayElement element)
    {
        ScreenplayElement clone = element switch
        {
            SectionElement section => new SectionElement(section.Text, section.RawText, section.LineIndex, section.SectionDepth),
            SynopsisElement synopsis => new SynopsisElement(synopsis.Text, synopsis.RawText, synopsis.LineIndex, synopsis.SectionLevel),
            NoteElement note => new NoteElement(note.Text, note.RawText, note.LineIndex, note.EndLineIndex, note.IsClosed),
            BoneyardElement boneyard => new BoneyardElement(boneyard.Text, boneyard.RawText, boneyard.LineIndex, boneyard.EndLineIndex, boneyard.IsClosed),
            CenteredTextElement centeredText => new CenteredTextElement(centeredText.Text, centeredText.RawText, centeredText.LineIndex),
            LyricsElement lyrics => new LyricsElement(lyrics.Text, lyrics.RawText, lyrics.LineIndex),
            SceneHeadingElement sceneHeading => new SceneHeadingElement(sceneHeading.Text, sceneHeading.RawText, sceneHeading.LineIndex, sceneHeading.IsForced),
            ActionElement action => new ActionElement(action.Text, action.RawText, action.LineIndex, action.EndLineIndex),
            TransitionElement transition => new TransitionElement(transition.Text, transition.RawText, transition.LineIndex, transition.IsForced),
            CharacterElement character => new CharacterElement(character.CharacterName, character.RawText, character.LineIndex, character.IsDualDialogue),
            ParentheticalElement parenthetical => new ParentheticalElement(parenthetical.Text, parenthetical.RawText, parenthetical.LineIndex, parenthetical.CharacterName, parenthetical.IsDualDialogue),
            DialogueElement dialogue => new DialogueElement(
                dialogue.CharacterName,
                dialogue.Parentheticals.ToArray(),
                dialogue.Lines.ToArray(),
                dialogue.Text,
                dialogue.RawText,
                dialogue.LineIndex,
                dialogue.EndLineIndex,
                dialogue.IsDualDialogue),
            ScratchpadCardElement draftCard => new ScratchpadCardElement(
                draftCard.Type,
                draftCard.Heading,
                draftCard.Description,
                draftCard.Text,
                draftCard.RawText,
                draftCard.Level,
                draftCard.SourceLineIndex,
                draftCard.SourceEndLineIndex,
                draftCard.ScriptHeading),
            _ => throw new NotSupportedException($"Unsupported screenplay element type: {element.GetType().Name}.")
        };

        clone.Id = element.Id;
        clone.Level = element.Level;
        clone.IsCollapsed = element.IsCollapsed;
        clone.IsDraft = element.IsDraft;
        clone.BodyText = element.BodyText;
        return clone;
    }

    private static BoardElementSignature CreateBoardElementSignature(ScreenplayElement element)
    {
        return element switch
        {
            SectionElement section => new BoardElementSignature(element.Type, element.RawText, element.LineIndex, element.EndLineIndex, element.Level, section.SectionDepth),
            SynopsisElement synopsis => new BoardElementSignature(element.Type, element.RawText, element.LineIndex, element.EndLineIndex, element.Level, synopsis.SectionLevel ?? -1),
            _ => new BoardElementSignature(element.Type, element.RawText, element.LineIndex, element.EndLineIndex, element.Level, 0)
        };
    }

    private static BoardSourceSignature CreateBoardSourceSignature(ScreenplayElement element)
    {
        return new BoardSourceSignature(
            element.Type,
            element.RawText,
            element.LineIndex,
            element.EndLineIndex);
    }

    private static ScreenplayElement MergeParsedBoardElement(
        ScreenplayElement parsedElement,
        ScreenplayElement existingElement)
    {
        if (existingElement is ScratchpadCardElement draftCard)
        {
            var description = ResolveStoredBoardDescription(existingElement, draftCard.Description);
            var mergedDraft = new ScratchpadCardElement(
                parsedElement.Type,
                draftCard.Heading,
                description,
                BuildBoardDraftText(parsedElement.Type, draftCard.Heading, parsedElement.ScriptHeading, description),
                parsedElement.RawText,
                parsedElement.Level,
                parsedElement.LineIndex,
                parsedElement.EndLineIndex,
                parsedElement.ScriptHeading);
            mergedDraft.Id = existingElement.Id;
            mergedDraft.IsCollapsed = existingElement.IsCollapsed;
            mergedDraft.IsDraft = existingElement.IsDraft;
            return mergedDraft;
        }

        var boardElement = CloneElement(parsedElement);
        boardElement.Id = existingElement.Id;
        boardElement.IsCollapsed = existingElement.IsCollapsed;
        boardElement.IsDraft = existingElement.IsDraft;
        return boardElement;
    }

    private static bool TryGetBoardSourceAnchor(ScreenplayElement element, out BoardSourceAnchor anchor)
    {
        if (element is ScratchpadCardElement draftCard)
        {
            if (draftCard.SourceLineIndex.HasValue && draftCard.SourceEndLineIndex.HasValue)
            {
                anchor = new BoardSourceAnchor(draftCard.SourceLineIndex.Value, draftCard.SourceEndLineIndex.Value);
                return true;
            }

            anchor = default;
            return false;
        }

        anchor = new BoardSourceAnchor(element.LineIndex, element.EndLineIndex);
        return true;
    }

    private static bool TryDequeueBoardMatch(
        Queue<ScreenplayElement> bucket,
        ISet<ScreenplayElement> consumedElements,
        out ScreenplayElement element)
    {
        while (bucket.Count > 0)
        {
            var candidate = bucket.Dequeue();
            if (!consumedElements.Add(candidate))
            {
                continue;
            }

            element = candidate;
            return true;
        }

        element = null!;
        return false;
    }

    private static bool ScreenplayElementsSequenceEqual(
        IReadOnlyList<ScreenplayElement> current,
        IReadOnlyList<ScreenplayElement> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            var left = current[i];
            var right = next[i];

            if (left.Id != right.Id ||
                left.Text != right.Text ||
                left.BodyText != right.BodyText ||
                CreateBoardElementSignature(left) != CreateBoardElementSignature(right) ||
                left.IsCollapsed != right.IsCollapsed ||
                left.IsDraft != right.IsDraft)
            {
                return false;
            }
        }

        return true;
    }

    private static bool BoardStructureElementsSequenceEqual(
        IReadOnlyList<ScreenplayElement> current,
        IReadOnlyList<ScreenplayElement> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            var left = current[i];
            var right = next[i];

            if (CreateBoardElementSignature(left) != CreateBoardElementSignature(right))
            {
                return false;
            }
        }

        return true;
    }

    private bool ShouldIncludeScratchpadElement(object item)
    {
        if (item is not ScreenplayElement element)
        {
            return false;
        }

        var searchText = ScratchpadSearchText.Trim();
        if (searchText.Length == 0)
        {
            return true;
        }

        return element.KindLabel.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
            element.Heading.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
            element.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
            element.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase);
    }

    private void ExecuteMoveToScratchpad(RichTextBox? richTextBox)
    {
        _ = TryMoveToScratchpad(richTextBox);
    }

    private void ExecuteDeleteScratchpadCard(object? _)
    {
        var element = SelectedScratchpadElement;
        if (element is null)
        {
            return;
        }

        if (!ScratchpadElements.Remove(element))
        {
            return;
        }

        RemoveCardFromScratchpadHistory(element);
        SelectedScratchpadElement = null;
    }

    private bool CanMoveToScratchpad(RichTextBox? richTextBox)
    {
        if (richTextBox is null)
        {
            return false;
        }

        var textRange = new TextRange(richTextBox.Selection.Start, richTextBox.Selection.End);
        return textRange.Start.CompareTo(textRange.End) != 0 &&
            RichTextBoxTextUtilities.GetSelectionLength(richTextBox) > 0 &&
            !string.IsNullOrWhiteSpace(textRange.Text);
    }

    private bool CanDeleteScratchpadCard(object? _)
    {
        return SelectedScratchpadElement is not null &&
            ScratchpadElements.Contains(SelectedScratchpadElement);
    }

    private void AddScratchpadCards(IReadOnlyList<ScreenplayElement> cards)
    {
        foreach (var card in cards)
        {
            card.IsCollapsed = true;
            card.IsDraft = true;

            if (!ScratchpadElements.Contains(card))
            {
                ScratchpadElements.Add(card);
            }
        }
    }

    private void RemoveScratchpadCards(IReadOnlyList<ScreenplayElement> cards)
    {
        foreach (var card in cards)
        {
            _ = ScratchpadElements.Remove(card);
        }
    }

    private void RecordScratchpadMove(
        string documentTextBeforeMove,
        string documentTextAfterMove,
        IReadOnlyList<ScreenplayElement> cards)
    {
        if (cards.Count == 0 ||
            string.Equals(documentTextBeforeMove, documentTextAfterMove, StringComparison.Ordinal))
        {
            return;
        }

        _scratchpadRedoHistory.Clear();
        _scratchpadUndoHistory.Push(new ScratchpadMoveHistoryEntry(
            documentTextBeforeMove,
            documentTextAfterMove,
            cards.ToArray()));
    }

    private bool TryUndoScratchpadMove(string currentText)
    {
        if (_scratchpadUndoHistory.Count == 0)
        {
            return false;
        }

        var operation = _scratchpadUndoHistory.Peek();
        if (!string.Equals(currentText, operation.DocumentTextBeforeMove, StringComparison.Ordinal))
        {
            return false;
        }

        _scratchpadUndoHistory.Pop();
        RemoveScratchpadCards(operation.Cards);
        if (operation.Cards.Count > 0)
        {
            _scratchpadRedoHistory.Push(operation);
        }

        return true;
    }

    private bool TryRedoScratchpadMove(string currentText)
    {
        if (_scratchpadRedoHistory.Count == 0)
        {
            return false;
        }

        var operation = _scratchpadRedoHistory.Peek();
        if (!string.Equals(currentText, operation.DocumentTextAfterMove, StringComparison.Ordinal))
        {
            return false;
        }

        _scratchpadRedoHistory.Pop();
        AddScratchpadCards(operation.Cards);
        if (operation.Cards.Count > 0)
        {
            _scratchpadUndoHistory.Push(operation);
        }

        return true;
    }

    private void RemoveCardFromScratchpadHistory(ScreenplayElement element)
    {
        PruneScratchpadHistory(_scratchpadUndoHistory, element);
        PruneScratchpadHistory(_scratchpadRedoHistory, element);
    }

    private static void PruneScratchpadHistory(
        Stack<ScratchpadMoveHistoryEntry> history,
        ScreenplayElement element)
    {
        if (history.Count == 0)
        {
            return;
        }

        var retainedEntries = new List<ScratchpadMoveHistoryEntry>(history.Count);
        foreach (var entry in history.Reverse())
        {
            var retainedCards = entry.Cards
                .Where(card => !ReferenceEquals(card, element))
                .ToArray();

            if (retainedCards.Length == 0)
            {
                continue;
            }

            retainedEntries.Add(entry with
            {
                Cards = retainedCards
            });
        }

        history.Clear();
        foreach (var entry in retainedEntries)
        {
            history.Push(entry);
        }
    }

    private static IReadOnlyList<ScreenplayElement> BuildScratchpadElementsFromSnippet(
        string selectedText,
        IReadOnlyList<ScreenplayElement> parsedElements)
    {
        if (parsedElements.Count == 0)
        {
            return
            [
                new ScratchpadCardElement(
                    ScreenplayElementType.Note,
                    "Note",
                    selectedText,
                    selectedText,
                    selectedText,
                    3)
            ];
        }

        if (parsedElements.Count == 1 && parsedElements[0] is NoteElement)
        {
            return parsedElements.ToArray();
        }

        var sceneHeadings = parsedElements.OfType<SceneHeadingElement>().ToArray();
        if (sceneHeadings.Length == 1 && parsedElements[0] is SceneHeadingElement sceneHeading)
        {
            return
            [
                new ScratchpadCardElement(
                    ScreenplayElementType.SceneHeading,
                    sceneHeading.Text,
                    BuildSceneDescription(selectedText, sceneHeading.Text),
                    selectedText,
                    selectedText,
                    sceneHeading.Level)
            ];
        }

        return parsedElements.ToArray();
    }

    private ScratchpadMoveOperation? CreateScratchpadMoveOperation(
        string scriptText,
        ParsedScreenplay parsed,
        int selectionStart,
        int selectionLength)
    {
        var selectionRange = ExpandSelectionToWholeLines(scriptText, selectionStart, selectionLength);
        if (selectionRange.Length <= 0)
        {
            return null;
        }

        var selectedElements = GetElementsIntersectingSelection(parsed.Elements, scriptText, selectionRange);
        if (selectedElements.Count == 1 && selectedElements[0] is NoteElement note)
        {
            return CreateNoteScratchpadOperation(scriptText, note);
        }

        if (selectedElements.Count > 0 && selectedElements[0] is SceneHeadingElement sceneHeading)
        {
            var blockRange = StoryHierarchyHelper.GetBlockRange(parsed.Elements, sceneHeading);
            var sceneBlock = parsed.Elements
                .Skip(blockRange.Index)
                .Take(blockRange.Count)
                .ToArray();

            if (HaveSameElementSequence(selectedElements, sceneBlock))
            {
                return CreateSceneScratchpadOperation(scriptText, sceneHeading, sceneBlock);
            }
        }

        return CreateFallbackScratchpadOperation(scriptText, selectionRange, selectedElements);
    }

    private ScratchpadMoveOperation CreateNoteScratchpadOperation(string scriptText, NoteElement note)
    {
        var removalRange = GetElementTextRange(scriptText, note);
        var card = new ScratchpadCardElement(
            ScreenplayElementType.Note,
            "Note",
            note.Text,
            note.Text,
            note.RawText,
            note.Level);

        return new ScratchpadMoveOperation(removalRange.Start, removalRange.Length, [card]);
    }

    private ScratchpadMoveOperation CreateSceneScratchpadOperation(
        string scriptText,
        SceneHeadingElement sceneHeading,
        IReadOnlyList<ScreenplayElement> sceneBlock)
    {
        var lastElement = sceneBlock[^1];
        var removalRange = CreateLineAlignedRange(scriptText, sceneHeading.LineIndex, lastElement.EndLineIndex);
        var sceneText = scriptText.Substring(removalRange.Start, removalRange.Length).Trim();
        var description = BuildSceneDescription(sceneText, sceneHeading.Text);
        var card = new ScratchpadCardElement(
            ScreenplayElementType.SceneHeading,
            sceneHeading.Text,
            description,
            sceneText,
            sceneText,
            sceneHeading.Level);

        return new ScratchpadMoveOperation(removalRange.Start, removalRange.Length, [card]);
    }

    private ScratchpadMoveOperation? CreateFallbackScratchpadOperation(
        string scriptText,
        CharacterTextRange selectionRange,
        IReadOnlyList<ScreenplayElement> selectedElements)
    {
        var selectedText = scriptText.Substring(selectionRange.Start, selectionRange.Length).Trim();
        if (selectedText.Length == 0)
        {
            return null;
        }

        var firstSelectedElement = selectedElements.FirstOrDefault();
        var cardType = firstSelectedElement?.Type == ScreenplayElementType.SceneHeading
            ? ScreenplayElementType.SceneHeading
            : ScreenplayElementType.Note;
        var heading = cardType == ScreenplayElementType.SceneHeading
            ? firstSelectedElement?.Heading ?? ExtractFirstNonEmptyLine(selectedText, "Scene")
            : "Note";
        var description = cardType == ScreenplayElementType.SceneHeading
            ? BuildSceneDescription(selectedText, heading)
            : selectedText;
        var level = cardType == ScreenplayElementType.SceneHeading ? 2 : 3;
        var card = new ScratchpadCardElement(
            cardType,
            heading,
            description,
            selectedText,
            selectedText,
            level);

        return new ScratchpadMoveOperation(selectionRange.Start, selectionRange.Length, [card]);
    }

    private static IReadOnlyList<ScreenplayElement> GetElementsIntersectingSelection(
        IReadOnlyList<ScreenplayElement> elements,
        string scriptText,
        CharacterTextRange selectionRange)
    {
        if (selectionRange.Length <= 0)
        {
            return Array.Empty<ScreenplayElement>();
        }

        var selectionStartLineIndex = GetLineIndexFromCharacterIndex(scriptText, selectionRange.Start);
        var selectionEndCharacterIndex = Math.Max(selectionRange.Start, selectionRange.End - 1);
        var selectionEndLineIndex = GetLineIndexFromCharacterIndex(scriptText, selectionEndCharacterIndex);

        return elements
            .Where(element => element.EndLineIndex >= selectionStartLineIndex && element.LineIndex <= selectionEndLineIndex)
            .ToArray();
    }

    private static bool HaveSameElementSequence(
        IReadOnlyList<ScreenplayElement> selectedElements,
        IReadOnlyList<ScreenplayElement> blockElements)
    {
        if (selectedElements.Count != blockElements.Count)
        {
            return false;
        }

        for (var index = 0; index < selectedElements.Count; index++)
        {
            if (!ReferenceEquals(selectedElements[index], blockElements[index]) &&
                selectedElements[index].Id != blockElements[index].Id)
            {
                return false;
            }
        }

        return true;
    }

    private static CharacterTextRange GetElementTextRange(string scriptText, ScreenplayElement element)
    {
        return CreateLineAlignedRange(scriptText, element.LineIndex, element.EndLineIndex);
    }

    private static CharacterTextRange ExpandSelectionToWholeLines(string scriptText, int selectionStart, int selectionLength)
    {
        var selectionEnd = Math.Min(scriptText.Length, selectionStart + selectionLength);
        var startLineIndex = GetLineIndexFromCharacterIndex(scriptText, selectionStart);
        var endCharacterIndex = Math.Max(selectionStart, selectionEnd - 1);
        var endLineIndex = GetLineIndexFromCharacterIndex(scriptText, endCharacterIndex);
        return CreateLineAlignedRange(scriptText, startLineIndex, endLineIndex);
    }

    private static CharacterTextRange CreateLineAlignedRange(string scriptText, int startLineIndex, int endLineIndex)
    {
        var start = GetCharacterIndexFromLineIndex(scriptText, startLineIndex);
        var end = GetCharacterIndexFromLineIndex(scriptText, endLineIndex + 1);
        if (end < start)
        {
            end = start;
        }

        return new CharacterTextRange(start, end - start);
    }

    private static int GetLineIndexFromCharacterIndex(string text, int characterIndex)
    {
        var safeCharacterIndex = Math.Max(0, Math.Min(characterIndex, text.Length));
        var lineIndex = 0;

        for (var index = 0; index < safeCharacterIndex; index++)
        {
            if (text[index] == '\n')
            {
                lineIndex++;
            }
        }

        return lineIndex;
    }

    private static int GetCharacterIndexFromLineIndex(string text, int lineIndex)
    {
        if (lineIndex <= 0)
        {
            return 0;
        }

        var currentLineIndex = 0;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] != '\n')
            {
                continue;
            }

            currentLineIndex++;
            if (currentLineIndex == lineIndex)
            {
                return index + 1;
            }
        }

        return text.Length;
    }

    private static string BuildSceneDescription(string sceneText, string sceneHeading)
    {
        var normalizedText = (sceneText ?? string.Empty).ReplaceLineEndings("\n").Trim();
        if (normalizedText.Length == 0)
        {
            return sceneHeading;
        }

        var newlineIndex = normalizedText.IndexOf('\n');
        if (newlineIndex < 0)
        {
            return sceneHeading;
        }

        var description = normalizedText[(newlineIndex + 1)..].Trim();
        return description.Length > 0 ? description : sceneHeading;
    }

    private static string ExtractFirstNonEmptyLine(string text, string fallback)
    {
        var normalizedText = (text ?? string.Empty).ReplaceLineEndings("\n");
        foreach (var line in normalizedText.Split('\n'))
        {
            var trimmedLine = line.Trim();
            if (trimmedLine.Length > 0)
            {
                return trimmedLine;
            }
        }

        return fallback;
    }

    private static string RemoveTextRange(string text, int start, int length)
    {
        if (length <= 0)
        {
            return text;
        }

        return text.Remove(start, length);
    }

    private static RichTextBox? TryGetOwningRichTextBox(System.Windows.Documents.TextRange selection)
    {
        DependencyObject? current = selection.Start.Parent as DependencyObject;

        while (current is not null)
        {
            if (current is RichTextBox richTextBox)
            {
                return richTextBox;
            }

            current = current switch
            {
                TextElement textElement => textElement.Parent as DependencyObject,
                FrameworkContentElement frameworkContentElement => frameworkContentElement.Parent as DependencyObject,
                FrameworkElement frameworkElement => frameworkElement.Parent as DependencyObject,
                _ => LogicalTreeHelper.GetParent(current)
            };
        }

        return null;
    }

    private void AdjustLineTypeOverridesForRemovedRange(string scriptText, int removalStart, int removalLength)
    {
        if (removalLength <= 0)
        {
            return;
        }

        var startLineNumber = GetLineIndexFromCharacterIndex(scriptText, removalStart) + 1;
        var endCharacterIndex = Math.Max(removalStart, removalStart + removalLength - 1);
        var endLineNumber = GetLineIndexFromCharacterIndex(scriptText, endCharacterIndex) + 1;
        var removedLineBreakCount = 0;

        foreach (var character in scriptText.AsSpan(removalStart, removalLength))
        {
            if (character == '\n')
            {
                removedLineBreakCount++;
            }
        }

        for (var lineNumber = startLineNumber; lineNumber <= endLineNumber; lineNumber++)
        {
            _lineTypeOverrides.Remove(lineNumber);
            _automaticActionLineOverrides.Remove(lineNumber);
        }

        if (removedLineBreakCount > 0)
        {
            ShiftLineTypeOverrides(endLineNumber + 1, -removedLineBreakCount);
        }
    }

    private bool TryNormalizeSelection(
        MoveToScratchpadRequest request,
        int textLength,
        out int selectionStart,
        out int selectionLength)
    {
        selectionStart = Math.Max(0, Math.Min(request.SelectionStart, textLength));
        selectionLength = Math.Max(0, Math.Min(request.SelectionLength, textLength - selectionStart));
        return selectionLength > 0;
    }

    private static bool TitlePageEntriesEqual(
        IReadOnlyList<TitlePageEntry> current,
        IReadOnlyList<TitlePageEntry> next)
    {
        if (current.Count != next.Count)
        {
            return false;
        }

        for (var i = 0; i < current.Count; i++)
        {
            var left = current[i];
            var right = next[i];

            if (left.FieldType != right.FieldType ||
                left.RawText != right.RawText ||
                left.Value != right.Value ||
                left.LineIndex != right.LineIndex)
            {
                return false;
            }
        }

        return true;
    }

    private void ReplaceDocument(string text, string? filePath, bool isDirty)
    {
        try
        {
            _suppressDirtyTracking = true;
            ClearLineTypeOverrides();
            ResetPlanningCollections();
            _sessionGoalBaselineNeedsCapture = true;

            var tempParsed = new FountainParser().Parse(text, null);
            var bodyStart = tempParsed.TitlePage.BodyStartLineIndex;
            var lines = (text ?? string.Empty).Replace("\r\n", "\n").Split('\n');
            
            TitlePageText = string.Join("\n", lines.Take(bodyStart)).TrimEnd();
            UpdateViewModelFromTitlePageText();
            DocumentText = string.Join("\n", lines.Skip(bodyStart));
            
            DocumentPath = filePath;
            IsDirty = isDirty;
            ResetCaretContext();

            if (isDirty)
            {
                StartRecoveryAutosave();
                SaveRecoverySnapshot();
            }
            else
            {
                StopRecoveryAutosave();
                RecoveryStorage.ClearRecoveryFile();
            }

            ScheduleOutlineRefresh();
        }
        finally
        {
            _suppressDirtyTracking = false;
        }
    }

    private void TitlePage_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_suppressTitlePageSync) return;
        UpdateTitlePageTextFromViewModel();
    }

    private void TitlePageCustomEntries_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (_suppressTitlePageSync) return;
        UpdateTitlePageTextFromViewModel();
    }

    private void UpdateTitlePageTextFromViewModel()
    {
        _suppressTitlePageSync = true;
        try
        {
            var sb = new System.Text.StringBuilder();
            if (!TitlePage.ShowTitlePage) sb.AppendLine("Show-Title-Page: false");
            if (TitlePage.ContactAlignRight) sb.AppendLine("Contact-Align: Right");
            
            if (!string.IsNullOrWhiteSpace(TitlePage.Title)) sb.AppendLine($"Title: {TitlePage.Title}");
            if (!string.IsNullOrWhiteSpace(TitlePage.Episode)) sb.AppendLine($"Episode: {TitlePage.Episode}");
            if (!string.IsNullOrWhiteSpace(TitlePage.Author)) sb.AppendLine($"Author: {TitlePage.Author}");
            if (!string.IsNullOrWhiteSpace(TitlePage.Credit)) sb.AppendLine($"Credit: {TitlePage.Credit}");
            if (!string.IsNullOrWhiteSpace(TitlePage.Source)) sb.AppendLine($"Source: {TitlePage.Source}");
            if (!string.IsNullOrWhiteSpace(TitlePage.DraftDate)) sb.AppendLine($"Draft date: {TitlePage.DraftDate}");
            if (!string.IsNullOrWhiteSpace(TitlePage.Revision)) sb.AppendLine($"Revision: {TitlePage.Revision}");
            
            if (!string.IsNullOrWhiteSpace(TitlePage.Notes))
            {
                var valLines = TitlePage.Notes.Replace("\r\n", "\n").Split('\n');
                sb.AppendLine($"Notes: {valLines[0]}");
                for (int i = 1; i < valLines.Length; i++)
                {
                    sb.AppendLine($"    {valLines[i]}");
                }
            }
            
            if (!string.IsNullOrWhiteSpace(TitlePage.Contact))
            {
                var valLines = TitlePage.Contact.Replace("\r\n", "\n").Split('\n');
                sb.AppendLine($"Contact: {valLines[0]}");
                for (int i = 1; i < valLines.Length; i++)
                {
                    sb.AppendLine($"    {valLines[i]}");
                }
            }

            foreach (var entry in TitlePage.CustomEntries)
            {
                if (!string.IsNullOrWhiteSpace(entry.Label) && !string.IsNullOrWhiteSpace(entry.Value))
                {
                    var valLines = entry.Value.Replace("\r\n", "\n").Split('\n');
                    sb.AppendLine($"{entry.Label}: {valLines[0]}");
                    for (int i = 1; i < valLines.Length; i++)
                    {
                        sb.AppendLine($"    {valLines[i]}");
                    }
                }
            }

            TitlePageText = sb.ToString().TrimEnd();
            IsDirty = true;
            ScheduleOutlineRefresh();
        }
        finally
        {
            _suppressTitlePageSync = false;
        }
    }

    private void UpdateViewModelFromTitlePageText()
    {
        if (_suppressTitlePageSync) return;
        _suppressTitlePageSync = true;
        try
        {
            var parsed = _parser.Parse(TitlePageText + "\n\n", null);
            TitlePage.CustomEntries.Clear();
            TitlePage.Title = "";
            TitlePage.Episode = "";
            TitlePage.Author = "";
            TitlePage.Credit = "written by";
            TitlePage.Source = "";
            TitlePage.Contact = "";
            TitlePage.DraftDate = "";
            TitlePage.Revision = "";
            TitlePage.Notes = "";
            TitlePage.ShowTitlePage = true;
            TitlePage.ContactAlignLeft = true;

            foreach (var entry in parsed.TitlePage.Entries)
            {
                var labelLower = entry.Label.ToLowerInvariant();
                switch (labelLower)
                {
                    case "title": TitlePage.Title = entry.Value; break;
                    case "episode": TitlePage.Episode = entry.Value; break;
                    case "author": TitlePage.Author = entry.Value; break;
                    case "credit": TitlePage.Credit = entry.Value; break;
                    case "contact": TitlePage.Contact = entry.Value; break;
                    case "draft date": TitlePage.DraftDate = entry.Value; break;
                    case "source": TitlePage.Source = entry.Value; break;
                    case "revision": TitlePage.Revision = entry.Value; break;
                    case "notes": TitlePage.Notes = entry.Value; break;
                    case "show-title-page": TitlePage.ShowTitlePage = !string.Equals(entry.Value, "false", StringComparison.OrdinalIgnoreCase); break;
                    case "contact-align": TitlePage.ContactAlignLeft = !string.Equals(entry.Value, "right", StringComparison.OrdinalIgnoreCase); break;
                    default:
                        TitlePage.CustomEntries.Add(new TitlePageEntryViewModel { Label = entry.Label, Value = entry.Value });
                        break;
                }
            }
        }
        finally
        {
            _suppressTitlePageSync = false;
        }
    }

    private static ScreenplayElementType DetermineEnterContinuation(
        ScreenplayElementType currentElementType,
        string currentLineText)
    {
        var trimmed = currentLineText.Trim();

        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        return currentElementType is ScreenplayElementType.Character or ScreenplayElementType.Parenthetical
            ? ScreenplayElementType.Dialogue
            : ScreenplayElementType.Action;
    }

    public ScreenplayElementType GetEffectiveLineType(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        if (_lineTypeOverrides.TryGetValue(lineNumber, out var explicitType))
        {
            return explicitType;
        }

        if (_automaticActionLineOverrides.Contains(lineNumber))
        {
            return ScreenplayElementType.Action;
        }

        var element = _lastParsed.Elements.FirstOrDefault(item => lineNumber >= item.StartLine && lineNumber <= item.EndLine);
        return element?.Type ?? ScreenplayElementType.Action;
    }

    public ScreenplayElementType GetLatestEffectiveLineType(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        var parsed = GetLatestParsedSnapshot();
        if (parsed.LineTypeOverrides.TryGetValue(lineNumber, out var explicitType))
        {
            return explicitType;
        }

        if (_automaticActionLineOverrides.Contains(lineNumber))
        {
            return ScreenplayElementType.Action;
        }

        var element = parsed.Elements.FirstOrDefault(item => lineNumber >= item.StartLine && lineNumber <= item.EndLine);
        return element?.Type ?? ScreenplayElementType.Action;
    }

    public void SetLineTypeOverride(int lineNumber, ScreenplayElementType elementType)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        SetLineTypeOverrideCore(lineNumber, elementType);

        CurrentLineNumber = lineNumber;
        CurrentElementType = elementType.ToString();
        CurrentElementText = GetCurrentElementDescription(elementType, isTabOverride: false);
        UpdateEnterContinuation(lineNumber, GetLineText(lineNumber));
        ScheduleOutlineRefresh();
    }

    public void SetAutomaticActionLineOverride(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        _lineTypeOverrides.Remove(lineNumber);
        if (!_automaticActionLineOverrides.Add(lineNumber))
        {
            return;
        }

        CurrentLineNumber = lineNumber;
        CurrentElementType = ScreenplayElementType.Action.ToString();
        CurrentElementText = GetCurrentElementDescription(ScreenplayElementType.Action, isTabOverride: false);
        UpdateEnterContinuation(lineNumber, GetLineText(lineNumber));
        ScheduleOutlineRefresh();
    }

    public bool ReleaseAutomaticActionLineOverrideIfNeeded(int lineNumber, string currentLineText)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        if (!_automaticActionLineOverrides.Contains(lineNumber) ||
            ShouldKeepAutomaticActionLineOverride(currentLineText))
        {
            return false;
        }

        _automaticActionLineOverrides.Remove(lineNumber);

        var elementType = GetLatestEffectiveLineType(lineNumber);
        CurrentLineNumber = lineNumber;
        CurrentElementType = elementType.ToString();
        CurrentElementText = GetCurrentElementDescription(elementType, isTabOverride: false);
        UpdateEnterContinuation(lineNumber, currentLineText);
        ScheduleOutlineRefresh();
        return true;
    }

    public void ClearLineTypeOverride(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        var removedExplicitOverride = _lineTypeOverrides.Remove(lineNumber);
        var removedAutomaticOverride = _automaticActionLineOverrides.Remove(lineNumber);
        if (!removedExplicitOverride && !removedAutomaticOverride)
        {
            return;
        }

        var elementType = GetLatestEffectiveLineType(lineNumber);
        CurrentLineNumber = lineNumber;
        CurrentElementType = elementType.ToString();
        CurrentElementText = GetCurrentElementDescription(elementType, isTabOverride: false);
        UpdateEnterContinuation(lineNumber, GetLineText(lineNumber));
        ScheduleOutlineRefresh();
    }

    public bool HasLineTypeOverride(int lineNumber)
    {
        return HasExplicitLineTypeOverride(lineNumber) || HasAutomaticActionLineOverride(lineNumber);
    }

    public void ShiftLineTypeOverrides(int startingLineNumber, int delta)
    {
        if (delta == 0 || (_lineTypeOverrides.Count == 0 && _automaticActionLineOverrides.Count == 0))
        {
            return;
        }

        if (startingLineNumber < 1)
        {
            startingLineNumber = 1;
        }

        var affectedEntries = _lineTypeOverrides
            .Where(pair => pair.Key >= startingLineNumber)
            .ToArray();
        var affectedAutomaticOverrides = _automaticActionLineOverrides
            .Where(lineNumber => lineNumber >= startingLineNumber)
            .ToArray();

        if (affectedEntries.Length == 0 && affectedAutomaticOverrides.Length == 0)
        {
            return;
        }

        foreach (var entry in affectedEntries)
        {
            _lineTypeOverrides.Remove(entry.Key);
        }

        foreach (var lineNumber in affectedAutomaticOverrides)
        {
            _automaticActionLineOverrides.Remove(lineNumber);
        }

        foreach (var entry in affectedEntries)
        {
            var shiftedLineNumber = Math.Max(1, entry.Key + delta);
            _lineTypeOverrides[shiftedLineNumber] = entry.Value;
        }

        foreach (var lineNumber in affectedAutomaticOverrides)
        {
            var shiftedLineNumber = Math.Max(1, lineNumber + delta);
            _automaticActionLineOverrides.Add(shiftedLineNumber);
        }

        ScheduleOutlineRefresh();
    }

    private static ScreenplayElementType CycleTabType(ScreenplayElementType currentType, bool forward)
    {
        var normalizedType = currentType switch
        {
            ScreenplayElementType.Action => ScreenplayElementType.Action,
            ScreenplayElementType.SceneHeading => ScreenplayElementType.SceneHeading,
            ScreenplayElementType.Character => ScreenplayElementType.Character,
            ScreenplayElementType.Dialogue => ScreenplayElementType.Dialogue,
            ScreenplayElementType.Parenthetical => ScreenplayElementType.Parenthetical,
            ScreenplayElementType.Transition => ScreenplayElementType.Transition,
            _ => ScreenplayElementType.Action
        };

        var cycleIndex = Array.IndexOf(TabCycleTypes, normalizedType);
        if (cycleIndex < 0)
        {
            cycleIndex = 0;
        }

        cycleIndex = forward
            ? (cycleIndex + 1) % TabCycleTypes.Length
            : (cycleIndex - 1 + TabCycleTypes.Length) % TabCycleTypes.Length;

        return TabCycleTypes[cycleIndex];
    }

    private static string GetCurrentElementDescription(ScreenplayElementType elementType, bool isTabOverride)
    {
        var description = elementType switch
        {
            ScreenplayElementType.SceneHeading => "Scene heading",
            ScreenplayElementType.Note => "Note",
            ScreenplayElementType.Boneyard => "Boneyard",
            ScreenplayElementType.CenteredText => "Centered text",
            ScreenplayElementType.Lyrics => "Lyrics",
            ScreenplayElementType.Character => "Character cue",
            ScreenplayElementType.Dialogue => "Dialogue",
            ScreenplayElementType.Parenthetical => "Parenthetical",
            ScreenplayElementType.Transition => "Transition",
            ScreenplayElementType.Section => "Section",
            ScreenplayElementType.Synopsis => "Synopsis",
            _ => isTabOverride ? "Action" : "Action / unclassified line"
        };

        return isTabOverride ? $"Tab cycle: {description}" : description;
    }

    private void SetLineTypeOverrideCore(int lineNumber, ScreenplayElementType elementType)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        _automaticActionLineOverrides.Remove(lineNumber);
        _lineTypeOverrides[lineNumber] = elementType;
    }

    private bool HasExplicitLineTypeOverride(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        return _lineTypeOverrides.ContainsKey(lineNumber);
    }

    private bool HasAutomaticActionLineOverride(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        return _automaticActionLineOverrides.Contains(lineNumber);
    }

    private IReadOnlyDictionary<int, ScreenplayElementType> GetLineTypeOverridesSnapshot()
    {
        var snapshot = new Dictionary<int, ScreenplayElementType>(_lineTypeOverrides);
        foreach (var lineNumber in _automaticActionLineOverrides)
        {
            snapshot[lineNumber] = ScreenplayElementType.Action;
        }

        return snapshot;
    }

    private static GoalType NormalizeSessionGoalType(GoalType goalType)
    {
        return goalType switch
        {
            GoalType.PageCount => GoalType.PageCount,
            GoalType.Timer => GoalType.Timer,
            _ => GoalType.WordCount
        };
    }

    private static GoalType NormalizeOverallGoalType(GoalType goalType)
    {
        return goalType == GoalType.PageCount ? GoalType.PageCount : GoalType.WordCount;
    }

    private string GetLineText(int lineNumber)
    {
        if (lineNumber < 1)
        {
            lineNumber = 1;
        }

        using var reader = new StringReader(DocumentText);

        for (var currentLine = 1; currentLine < lineNumber; currentLine++)
        {
            if (reader.ReadLine() is null)
            {
                return string.Empty;
            }
        }

        return reader.ReadLine() ?? string.Empty;
    }

    private string GetScriptSourceText()
    {
        // Scratchpad cards live outside the source document, so parsing/exporting starts from DocumentText only.
        return DocumentText;
    }

    private void ClearLineTypeOverrides()
    {
        _lineTypeOverrides.Clear();
        _automaticActionLineOverrides.Clear();
    }

    private void ResetPlanningCollections()
    {
        if (BoardElements.Count > 0)
        {
            BoardElements = new ObservableCollection<ScreenplayElement>();
            AttachBoardElementsCollection(BoardElements);
            OnPropertyChanged(nameof(BoardElements));
        }

        if (VisibleBoardElements.Count > 0)
        {
            VisibleBoardElements = new ObservableCollection<ScreenplayElement>();
            OnPropertyChanged(nameof(VisibleBoardElements));
            OnPropertyChanged(nameof(HasVisibleBoardItems));
        }

        SelectedBoardElement = null;
        IsBoardSyncRequired = false;
        ClearBoardDropIndicator();

        if (ScratchpadElements.Count > 0)
        {
            ScratchpadElements.Clear();
        }

        SelectedScratchpadElement = null;
        _scratchpadUndoHistory.Clear();
        _scratchpadRedoHistory.Clear();
        ScratchpadSearchText = string.Empty;
        CommandManager.InvalidateRequerySuggested();
    }

    private static bool ShouldKeepAutomaticActionLineOverride(string currentLineText)
    {
        var trimmed = (currentLineText ?? string.Empty).Trim();
        if (trimmed.Length == 0)
        {
            return true;
        }

        if (TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()) ||
            TextAnalysis.IsLiveCharacterCueCandidate(currentLineText.AsSpan(), 45, 6))
        {
            return false;
        }

        return !(trimmed.StartsWith("(") && trimmed.EndsWith(")", StringComparison.Ordinal));
    }

    private void ResetCaretContext()
    {
        CurrentLineNumber = 1;
        CurrentElementType = "Action";
        CurrentElementText = "No screenplay element";
        EnterContinuationText = "Action";
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

        var titlePrefix = TitlePageText.Length > 0 ? TitlePageText + "\n\n" : "";
        RecoveryStorage.SaveRecoveryFile(new RecoveryDocument
        {
            Text = titlePrefix + DocumentText,
            FilePath = DocumentPath,
            SavedAtUtc = DateTimeOffset.UtcNow,
            GoalConfiguration = GoalConfiguration,
            SessionGoalConfiguration = SessionGoalConfiguration,
            EditorZoomPercent = EditorZoomPercent
        });
    }

    private void SetEditorZoomPercent(double value)
    {
        EditorZoomPercent = value;
    }

    private void AttachBoardElementsCollection(ObservableCollection<ScreenplayElement> boardElements)
    {
        boardElements.CollectionChanged += (_, _) =>
        {
            RefreshVisibleBoardElements();
            if (_selectedBoardElement is not null &&
                FindBoardElementIndex(_selectedBoardElement) < 0)
            {
                SelectedBoardElement = null;
            }

            OnPropertyChanged(nameof(BoardElements));
            CommandManager.InvalidateRequerySuggested();
        };

        RefreshVisibleBoardElements();
    }

    private void RefreshVisibleBoardElements()
    {
        var visibleBoardElements = BuildVisibleBoardElements(BoardElements);
        VisibleBoardElements = new ObservableCollection<ScreenplayElement>(visibleBoardElements);
        if (_selectedBoardElement is not null &&
            !VisibleBoardElements.Any(candidate => candidate.Id == _selectedBoardElement.Id))
        {
            SelectedBoardElement = null;
        }

        OnPropertyChanged(nameof(VisibleBoardElements));
        OnPropertyChanged(nameof(HasVisibleBoardItems));
    }

    private int FindBoardElementIndex(ScreenplayElement element)
    {
        for (var index = 0; index < BoardElements.Count; index++)
        {
            var candidate = BoardElements[index];
            if (ReferenceEquals(candidate, element) || candidate.Id == element.Id)
            {
                return index;
            }
        }

        return -1;
    }

    private bool TryResolveVisibleBoardDropIndex(
        ScreenplayElement? visibleTargetElement,
        bool insertAfter,
        out int targetIndex)
    {
        if (BoardElements.Count == 0)
        {
            targetIndex = 0;
            return true;
        }

        if (visibleTargetElement is null)
        {
            targetIndex = BoardElements.Count;
            return true;
        }

        var elementIndex = FindBoardElementIndex(visibleTargetElement);
        if (elementIndex < 0)
        {
            targetIndex = BoardElements.Count;
            return false;
        }

        if (!insertAfter)
        {
            targetIndex = elementIndex;
            return true;
        }

        try
        {
            var blockRange = StoryHierarchyHelper.GetBlockRange(BoardElements, BoardElements[elementIndex]);
            targetIndex = Math.Clamp(blockRange.Index + blockRange.Count, 0, BoardElements.Count);
            return true;
        }
        catch (ArgumentException)
        {
            targetIndex = BoardElements.Count;
            return false;
        }
    }

    private int GetSelectedBoardInsertIndex()
    {
        if (_selectedBoardElement is null)
        {
            return BoardElements.Count;
        }

        try
        {
            var selectedRange = StoryHierarchyHelper.GetBlockRange(BoardElements, _selectedBoardElement);
            return Math.Clamp(selectedRange.Index + selectedRange.Count, 0, BoardElements.Count);
        }
        catch (ArgumentException)
        {
            return BoardElements.Count;
        }
    }

    private ScreenplayElement? ReplaceBoardElement(ScreenplayElement currentElement, ScreenplayElement updatedElement)
    {
        var elementIndex = FindBoardElementIndex(currentElement);
        if (elementIndex < 0)
        {
            return null;
        }

        updatedElement.Id = currentElement.Id;
        updatedElement.IsCollapsed = currentElement.IsCollapsed;
        updatedElement.IsDraft = true;
        BoardElements[elementIndex] = updatedElement;
        SelectedBoardElement = updatedElement;
        UpdateBoardSyncRequiredState();
        return updatedElement;
    }

    private static IReadOnlyList<ScreenplayElement> BuildVisibleBoardElements(
        IReadOnlyList<ScreenplayElement> boardElements)
    {
        var visibleBoardElements = new List<ScreenplayElement>(boardElements.Count);
        int? collapsedAncestorLevel = null;

        foreach (var element in boardElements)
        {
            if (collapsedAncestorLevel.HasValue &&
                element.Level <= collapsedAncestorLevel.Value)
            {
                collapsedAncestorLevel = null;
            }

            if (!collapsedAncestorLevel.HasValue)
            {
                visibleBoardElements.Add(element);
            }

            if (!collapsedAncestorLevel.HasValue &&
                element.IsCollapsed)
            {
                collapsedAncestorLevel = element.Level;
            }
        }

        return visibleBoardElements;
    }

    private static bool TryResolveBoardCardKind(
        string kind,
        out ScreenplayElementType type,
        out int level)
    {
        switch ((kind ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "act":
                type = ScreenplayElementType.Section;
                level = 0;
                return true;
            case "sequence":
                type = ScreenplayElementType.Section;
                level = 1;
                return true;
            case "scene":
                type = ScreenplayElementType.SceneHeading;
                level = 2;
                return true;
            case "note":
                type = ScreenplayElementType.Note;
                level = 3;
                return true;
            default:
                type = ScreenplayElementType.SceneHeading;
                level = 2;
                return false;
        }
    }

    public static bool TryResolveBoardChildLevel(
        ScreenplayElement? parentElement,
        ScreenplayElement? childElement,
        out int level)
    {
        if (parentElement is null ||
            childElement is null ||
            parentElement.Id == childElement.Id)
        {
            level = 0;
            return false;
        }

        if (parentElement.Type == ScreenplayElementType.Section &&
            parentElement.Level == 0 &&
            childElement.Type == ScreenplayElementType.Section &&
            childElement.Level == 1)
        {
            level = 1;
            return true;
        }

        if (parentElement.Type == ScreenplayElementType.Section &&
            parentElement.Level == 1 &&
            childElement.Type is ScreenplayElementType.SceneHeading or ScreenplayElementType.Note)
        {
            level = 2;
            return true;
        }

        level = 0;
        return false;
    }

    private static string GetBoardCardHeading(ScreenplayElement element)
    {
        if (element is ScratchpadCardElement draftCard)
        {
            return draftCard.Heading;
        }

        return element.Heading;
    }

    private static string GetBoardCardDescription(ScreenplayElement element)
    {
        return NormalizeBoardCardDescription(element.BoardDescription);
    }

    private static string GetBoardCardSceneHeading(ScreenplayElement element)
    {
        return element.Type == ScreenplayElementType.SceneHeading
            ? NormalizeBoardCardSceneHeading(element, element.ScriptHeading)
            : string.Empty;
    }

    private static string GetBoardScriptDescription(ScreenplayElement? element)
    {
        if (element is null)
        {
            return string.Empty;
        }

        return NormalizeBoardCardDescription(element.BoardDescription);
    }

    private static string NormalizeBoardCardHeading(ScreenplayElement? element, string? heading)
    {
        var trimmedHeading = (heading ?? string.Empty).Trim();
        if (trimmedHeading.Length > 0)
        {
            return trimmedHeading;
        }

        if (element is not null)
        {
            return element.Type switch
            {
                ScreenplayElementType.SceneHeading => "NEW SCENE",
                ScreenplayElementType.Note => "Note",
                ScreenplayElementType.Section when element.Level == 0 => "NEW ACT",
                ScreenplayElementType.Section when element.Level == 1 => "NEW SEQUENCE",
                ScreenplayElementType.Section => "NEW BEAT",
                _ => "New Card"
            };
        }

        return "New Card";
    }

    private static string NormalizeBoardCardSceneHeading(ScreenplayElement? element, string? sceneHeading)
    {
        var trimmedSceneHeading = (sceneHeading ?? string.Empty).Trim();
        if (trimmedSceneHeading.Length > 0)
        {
            return trimmedSceneHeading;
        }

        if (element is not null &&
            element.Type == ScreenplayElementType.SceneHeading &&
            !string.IsNullOrWhiteSpace(element.ScriptHeading))
        {
            return element.ScriptHeading.Trim();
        }

        return "NEW SCENE";
    }

    private static string NormalizeBoardCardDescription(string? description)
    {
        var normalizedDescription = (description ?? string.Empty).ReplaceLineEndings("\n").Trim();
        return IsDefaultBoardDescription(normalizedDescription)
            ? string.Empty
            : normalizedDescription;
    }

    public static bool HasMeaningfulBoardDescription(string? description)
    {
        return !string.IsNullOrWhiteSpace(description) &&
            !IsDefaultBoardDescription(description);
    }

    public static bool IsDefaultBoardDescription(string? description)
    {
        var normalizedDescription = (description ?? string.Empty).Trim();
        return string.Equals(normalizedDescription, DefaultNewBoardCardDescription, StringComparison.Ordinal) ||
            string.Equals(normalizedDescription, LegacyDefaultNewBoardCardDescription, StringComparison.Ordinal);
    }

    private static string ResolveStoredBoardDescription(ScreenplayElement? sourceElement, string description)
    {
        if (HasMeaningfulBoardDescription(description))
        {
            return description;
        }

        var sourceDescription = GetBoardScriptDescription(sourceElement);
        return HasMeaningfulBoardDescription(sourceDescription)
            ? sourceDescription.Trim()
            : description;
    }

    private static string BuildBoardDraftText(
        ScreenplayElementType type,
        string heading,
        string sceneHeading,
        string description)
    {
        var parts = new List<string>(3);
        if (!string.IsNullOrWhiteSpace(heading))
        {
            parts.Add(heading);
        }

        if (type == ScreenplayElementType.SceneHeading &&
            !string.IsNullOrWhiteSpace(sceneHeading) &&
            !string.Equals(sceneHeading.Trim(), heading?.Trim(), StringComparison.Ordinal))
        {
            parts.Add(sceneHeading.Trim());
        }

        if (!string.IsNullOrWhiteSpace(description))
        {
            parts.Add(description);
        }

        return parts.Count == 0 ? string.Empty : string.Join("\n", parts);
    }

    private static string BuildBoardDraftRawText(
        Guid id,
        ScreenplayElementType type,
        int level,
        string heading,
        string sceneHeading,
        string description)
    {
        return SerializeDraftBoardElement(id, type, level, heading, sceneHeading, description);
    }

    private static ScratchpadCardElement CreateBoardDraftElement(
        ScreenplayElementType type,
        int level,
        string heading,
        string sceneHeading,
        string description)
    {
        var id = Guid.NewGuid();
        var normalizedHeading = NormalizeBoardCardHeading(null, heading);
        var normalizedSceneHeading = type == ScreenplayElementType.SceneHeading
            ? NormalizeBoardCardSceneHeading(null, sceneHeading)
            : string.Empty;
        var normalizedDescription = ResolveStoredBoardDescription(null, NormalizeBoardCardDescription(description));
        
        return new ScratchpadCardElement(
            type,
            normalizedHeading,
            normalizedDescription,
            BuildBoardDraftText(type, normalizedHeading, normalizedSceneHeading, normalizedDescription),
            BuildBoardDraftRawText(id, type, level, normalizedHeading, normalizedSceneHeading, normalizedDescription),
            level,
            scriptHeading: normalizedSceneHeading)
        {
            Id = id,
            IsCollapsed = false
        };
    }

    private static ScratchpadCardElement CreateBoardDraftElement(
        ScreenplayElement sourceElement,
        ScreenplayElementType type,
        int level,
        string heading,
        string sceneHeading,
        string description)
    {
        var normalizedHeading = NormalizeBoardCardHeading(sourceElement, heading);
        var normalizedSceneHeading = type == ScreenplayElementType.SceneHeading
            ? NormalizeBoardCardSceneHeading(sourceElement, sceneHeading)
            : string.Empty;
        var normalizedDescription = ResolveStoredBoardDescription(sourceElement, NormalizeBoardCardDescription(description));
        GetBoardDraftSourceReference(sourceElement, out var sourceLineIndex, out var sourceEndLineIndex);
        var draftElement = new ScratchpadCardElement(
            type,
            normalizedHeading,
            normalizedDescription,
            BuildBoardDraftText(type, normalizedHeading, normalizedSceneHeading, normalizedDescription),
            BuildBoardDraftRawText(sourceElement.Id, type, level, normalizedHeading, normalizedSceneHeading, normalizedDescription),
            level,
            sourceLineIndex,
            sourceEndLineIndex,
            normalizedSceneHeading);
        draftElement.Id = sourceElement.Id;
        draftElement.IsCollapsed = sourceElement.Type == ScreenplayElementType.Section
            ? sourceElement.IsCollapsed
            : false;
        draftElement.IsDraft = true;
        return draftElement;
    }

    private static void GetBoardDraftSourceReference(
        ScreenplayElement sourceElement,
        out int? sourceLineIndex,
        out int? sourceEndLineIndex)
    {
        if (sourceElement is ScratchpadCardElement draftCard)
        {
            sourceLineIndex = draftCard.SourceLineIndex;
            sourceEndLineIndex = draftCard.SourceEndLineIndex;
            return;
        }

        sourceLineIndex = sourceElement.LineIndex;
        sourceEndLineIndex = sourceElement.EndLineIndex;
    }

    private void UpdateBoardSyncRequiredState()
    {
        UpdateBoardSyncRequiredState(GetCurrentParsedSnapshot().Elements);
    }

    private void UpdateBoardSyncRequiredState(IReadOnlyList<ScreenplayElement> parsedElements)
    {
        var canonicalBoardElements = BuildBoardElements(parsedElements, BoardElements, preserveCustomLayout: false);
        IsBoardSyncRequired = !BoardStructureElementsSequenceEqual(BoardElements, canonicalBoardElements);
        CommandManager.InvalidateRequerySuggested();
    }

    private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(storage, value))
        {
            return false;
        }

        storage = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged(string? propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private readonly record struct BoardElementSignature(
        ScreenplayElementType Type,
        string RawText,
        int LineIndex,
        int EndLineIndex,
        int Level,
        int AuxiliaryLevel);

    private readonly record struct BoardSourceSignature(
        ScreenplayElementType Type,
        string RawText,
        int LineIndex,
        int EndLineIndex);

    private readonly record struct BoardSourceAnchor(
        int LineIndex,
        int EndLineIndex);

    private readonly record struct BoardSyncSegment(
        int OriginalStartLineIndex,
        int OriginalEndLineIndexExclusive,
        IReadOnlyList<string> TrailingLines);

    private sealed record BoardSyncPlan(
        string DocumentText,
        IReadOnlyDictionary<int, ScreenplayElementType> LineTypeOverrides,
        IReadOnlyCollection<int> AutomaticActionOverrides);

    public sealed record MoveToScratchpadRequest(int SelectionStart, int SelectionLength);

    private readonly record struct CharacterTextRange(int Start, int Length)
    {
        public int End => Start + Length;
    }

    private sealed record ScratchpadMoveOperation(
        int RemovalStart,
        int RemovalLength,
        IReadOnlyList<ScreenplayElement> Cards);

    private sealed record ScratchpadMoveHistoryEntry(
        string DocumentTextBeforeMove,
        string DocumentTextAfterMove,
        IReadOnlyList<ScreenplayElement> Cards);
}
