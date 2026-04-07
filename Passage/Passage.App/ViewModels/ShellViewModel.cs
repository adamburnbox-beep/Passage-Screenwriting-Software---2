using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Threading;
using Passage.App.Services;
using Passage.Core.Goals;
using Passage.Export;
using Passage.Parser;

namespace Passage.App.ViewModels;

public sealed class ShellViewModel : INotifyPropertyChanged
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
    private const double DefaultEditorZoomPercent = 100.0;
    private readonly DispatcherTimer _sessionSaveTimer;
    private MainWindowViewModel? _selectedDocument;
    private bool _suppressSessionSave;

    public ShellViewModel(RecoveryDocument? recoveredDocument = null)
    {
        Documents = new ObservableCollection<MainWindowViewModel>();
        Documents.CollectionChanged += Documents_CollectionChanged;

        _sessionSaveTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1.5)
        };
        _sessionSaveTimer.Tick += (_, _) => SaveSessionNow();

        RestoreSession(recoveredDocument);
    }

    public ObservableCollection<MainWindowViewModel> Documents { get; }

    public MainWindowViewModel? SelectedDocument
    {
        get => _selectedDocument;
        set
        {
            if (ReferenceEquals(_selectedDocument, value))
            {
                return;
            }

            if (_selectedDocument is not null)
            {
                _selectedDocument.PropertyChanged -= SelectedDocument_PropertyChanged;
                _selectedDocument.OutlineUpdated -= SelectedDocument_OutlineUpdated;
            }

            _selectedDocument = value;

            if (_selectedDocument is not null)
            {
                _selectedDocument.PropertyChanged += SelectedDocument_PropertyChanged;
                _selectedDocument.OutlineUpdated += SelectedDocument_OutlineUpdated;
            }

            OnPropertyChanged(nameof(SelectedDocument));
            OnPropertyChanged(nameof(DocumentText));
            OnPropertyChanged(nameof(DocumentPath));
            OnPropertyChanged(nameof(IsDirty));
            OnPropertyChanged(nameof(WindowTitle));
            OnPropertyChanged(nameof(OutlineRoots));
            OnPropertyChanged(nameof(NotesRoots));
            OnPropertyChanged(nameof(BoardElements));
            OnPropertyChanged(nameof(HasBoardItems));
            OnPropertyChanged(nameof(IsBoardSyncRequired));
            OnPropertyChanged(nameof(ScratchpadElements));
            OnPropertyChanged(nameof(ScratchpadSearchText));
            OnPropertyChanged(nameof(SelectedScratchpadElement));
            OnPropertyChanged(nameof(HasScratchpadItems));
            OnPropertyChanged(nameof(ScratchpadEmptyMessage));
            OnPropertyChanged(nameof(MoveToScratchpadCommand));
            OnPropertyChanged(nameof(DeleteScratchpadCardCommand));
            OnPropertyChanged(nameof(CreateNewCardCommand));
            OnPropertyChanged(nameof(SyncBoardToScriptCommand));
            OnPropertyChanged(nameof(TitlePage));
            OnPropertyChanged(nameof(TitlePageEntries));
            OnPropertyChanged(nameof(PreviewElements));
            OnPropertyChanged(nameof(SelectedGoalType));
            OnPropertyChanged(nameof(GoalTargetValue));
            OnPropertyChanged(nameof(GoalProgressPercent));
            OnPropertyChanged(nameof(GoalCurrentDisplayText));
            OnPropertyChanged(nameof(GoalTargetDisplayText));
            OnPropertyChanged(nameof(GoalStateText));
            OnPropertyChanged(nameof(GoalTargetUnitLabel));
            OnPropertyChanged(nameof(SessionGoalTypes));
            OnPropertyChanged(nameof(SessionSelectedGoalType));
            OnPropertyChanged(nameof(SessionGoalTargetValue));
            OnPropertyChanged(nameof(SessionGoalProgressPercent));
            OnPropertyChanged(nameof(SessionGoalCurrentDisplayText));
            OnPropertyChanged(nameof(SessionGoalTargetDisplayText));
            OnPropertyChanged(nameof(SessionGoalStateText));
            OnPropertyChanged(nameof(SessionGoalTargetUnitLabel));
            OnPropertyChanged(nameof(GoalTimerElapsedText));
            OnPropertyChanged(nameof(GoalTimerRemainingText));
            OnPropertyChanged(nameof(IsTimerGoal));
            OnPropertyChanged(nameof(GoalTimerState));
            OnPropertyChanged(nameof(GoalTimerPrimaryButtonText));
            OnPropertyChanged(nameof(GoalTimerSecondaryButtonText));
            OnPropertyChanged(nameof(SelectedOutlineLineNumber));
            OnPropertyChanged(nameof(CurrentLineNumber));
            OnPropertyChanged(nameof(CurrentElementType));
            OnPropertyChanged(nameof(CurrentElementText));
            OnPropertyChanged(nameof(EnterContinuationText));
            OnPropertyChanged(nameof(IsBoardModeActive));
            RaiseEditorZoomProperties();

            RefreshStatusState();
            OutlineUpdated?.Invoke(this, EventArgs.Empty);
            ScheduleSessionSave();
        }
    }

    public string DocumentText
    {
        get => SelectedDocument?.DocumentText ?? string.Empty;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.DocumentText = value;
        }
    }

    public string? DocumentPath => SelectedDocument?.DocumentPath;

    public bool IsDirty => SelectedDocument?.IsDirty ?? false;

    public string WindowTitle => SelectedDocument?.WindowTitle ?? "Passage";

    public ObservableCollection<OutlineNodeViewModel> OutlineRoots => SelectedDocument?.OutlineRoots ?? _emptyOutline;

    public ObservableCollection<OutlineNodeViewModel> NotesRoots => SelectedDocument?.NotesRoots ?? _emptyNotes;

    public ObservableCollection<ScreenplayElement> BoardElements => SelectedDocument?.BoardElements ?? _emptyBoard;

    public ObservableCollection<ScreenplayElement> ScratchpadElements => SelectedDocument?.ScratchpadElements ?? _emptyScratchpad;

    public bool HasBoardItems => SelectedDocument?.BoardElements.Count > 0;

    public bool IsBoardSyncRequired => SelectedDocument?.IsBoardSyncRequired ?? false;

    public bool IsBoardModeActive
    {
        get => SelectedDocument?.IsBoardModeActive ?? false;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.IsBoardModeActive = value;
        }
    }

    public string ScratchpadSearchText
    {
        get => SelectedDocument?.ScratchpadSearchText ?? string.Empty;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.ScratchpadSearchText = value;
        }
    }

    public ScreenplayElement? SelectedScratchpadElement
    {
        get => SelectedDocument?.SelectedScratchpadElement;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.SelectedScratchpadElement = value;
        }
    }

    public bool HasScratchpadItems => SelectedDocument?.ScratchpadElements.Count > 0;

    public string ScratchpadEmptyMessage => string.IsNullOrWhiteSpace(ScratchpadSearchText)
        ? "Moved scenes, notes, and loose ideas will appear here."
        : "No scratchpad cards match the current search.";

    public ICommand? MoveToScratchpadCommand => SelectedDocument?.MoveToScratchpadCommand;

    public ICommand? DeleteScratchpadCardCommand => SelectedDocument?.DeleteScratchpadCardCommand;

    public ICommand? CreateNewCardCommand => SelectedDocument?.CreateNewCardCommand;

    public ICommand? SyncBoardToScriptCommand => SelectedDocument?.SyncBoardToScriptCommand;

    public TitlePageViewModel? TitlePage => SelectedDocument?.TitlePage;

    public ObservableCollection<TitlePageEntry> TitlePageEntries => SelectedDocument?.TitlePageEntries ?? _emptyTitlePage;

    public ObservableCollection<PreviewElementItem> PreviewElements => SelectedDocument?.PreviewElements ?? _emptyPreview;

    public IReadOnlyList<IExporter> AvailableExporters { get; } = ExporterCatalog.GetDefaultExporters();

    public IReadOnlyList<GoalType> GoalTypes => GoalTypeOptions;

    public IReadOnlyList<GoalType> SessionGoalTypes => SessionGoalTypeOptions;

    public GoalType SelectedGoalType
    {
        get => SelectedDocument?.SelectedGoalType ?? GoalType.WordCount;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.SelectedGoalType = value;
        }
    }

    public GoalType SessionSelectedGoalType
    {
        get => SelectedDocument?.SessionSelectedGoalType ?? GoalType.WordCount;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.SessionSelectedGoalType = value;
        }
    }

    public int GoalTargetValue
    {
        get => SelectedDocument?.GoalTargetValue ?? 0;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.GoalTargetValue = value;
        }
    }

    public int SessionGoalTargetValue
    {
        get => SelectedDocument?.SessionGoalTargetValue ?? 0;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.SessionGoalTargetValue = value;
        }
    }

    public double GoalProgressPercent => SelectedDocument?.GoalProgressPercent ?? 0;

    public double SessionGoalProgressPercent => SelectedDocument?.SessionGoalProgressPercent ?? 0;

    public string GoalCurrentDisplayText => SelectedDocument?.GoalCurrentDisplayText ?? "0 words";

    public string SessionGoalCurrentDisplayText => SelectedDocument?.SessionGoalCurrentDisplayText ?? "0 words";

    public string GoalTargetDisplayText => SelectedDocument?.GoalTargetDisplayText ?? "1000 words";

    public string SessionGoalTargetDisplayText => SelectedDocument?.SessionGoalTargetDisplayText ?? "1000 words";

    public string GoalStateText => SelectedDocument?.GoalStateText ?? "In progress";

    public string SessionGoalStateText => SelectedDocument?.SessionGoalStateText ?? "In progress";

    public string GoalTargetUnitLabel => SelectedDocument?.GoalTargetUnitLabel ?? "Words";

    public string SessionGoalTargetUnitLabel => SelectedDocument?.SessionGoalTargetUnitLabel ?? "Words";

    public string GoalTimerElapsedText => SelectedDocument?.GoalTimerElapsedText ?? "00:00";

    public string GoalTimerRemainingText => SelectedDocument?.GoalTimerRemainingText ?? "00:00";

    public bool IsTimerGoal => SelectedDocument?.IsTimerGoal ?? false;

    public TimerGoalState GoalTimerState => SelectedDocument?.GoalTimerState ?? TimerGoalState.Idle;

    public string GoalTimerPrimaryButtonText => SelectedDocument?.GoalTimerPrimaryButtonText ?? "Start";

    public string GoalTimerSecondaryButtonText => SelectedDocument?.GoalTimerSecondaryButtonText ?? "Reset";

    public bool HasOutlineItems => SelectedDocument?.OutlineRoots.Count > 0;

    public bool HasNoteItems => SelectedDocument?.NotesRoots.Count > 0;

    public bool HasTitlePageEntries => SelectedDocument?.TitlePageEntries.Count > 0;

    public bool HasPreviewContent => HasTitlePageEntries || SelectedDocument?.PreviewElements.Count > 0;

    public string StatusMessage => SelectedDocument?.GoalProgressSummaryText ?? "Ready";

    public string OutlineEmptyMessage => "Sections, synopses, and scene headings will appear here.";

    public string NotesEmptyMessage => "Notes will appear here.";

    public string ScratchpadLabel => "Scratchpad";

    public string PreviewEmptyMessage => "A live screenplay preview will appear here as you type.";

    public int? SelectedOutlineLineNumber
    {
        get => SelectedDocument?.SelectedOutlineLineNumber;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.SelectedOutlineLineNumber = value;
        }
    }

    public int CurrentLineNumber => SelectedDocument?.CurrentLineNumber ?? 1;

    public string CurrentElementType => SelectedDocument?.CurrentElementType ?? "Action";

    public string CurrentElementText => SelectedDocument?.CurrentElementText ?? "No screenplay element";

    public string EnterContinuationText => SelectedDocument?.EnterContinuationText ?? "Action";

    public double EditorZoomPercent
    {
        get => SelectedDocument?.EditorZoomPercent ?? DefaultEditorZoomPercent;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.EditorZoomPercent = value;
        }
    }

    public double EditorZoomScale => EditorZoomPercent / 100.0;

    public string EditorZoomDisplayText => $"{EditorZoomPercent:0}%";

    public double BoardZoomPercent
    {
        get => SelectedDocument?.BoardZoomPercent ?? DefaultEditorZoomPercent;
        set
        {
            if (SelectedDocument is null)
            {
                return;
            }

            SelectedDocument.BoardZoomPercent = value;
        }
    }

    public double BoardZoomScale => BoardZoomPercent / 100.0;

    public string BoardZoomDisplayText => $"{BoardZoomPercent:0}%";

    public string ActiveZoomDisplayText => IsBoardModeActive ? BoardZoomDisplayText : EditorZoomDisplayText;

    public ScreenplayElementType GetEffectiveLineType(int lineNumber)
    {
        return SelectedDocument?.GetEffectiveLineType(lineNumber) ?? ScreenplayElementType.Action;
    }

    public ScreenplayElementType GetLatestEffectiveLineType(int lineNumber)
    {
        return SelectedDocument?.GetLatestEffectiveLineType(lineNumber) ?? ScreenplayElementType.Action;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public event EventHandler? OutlineUpdated;

    public ParsedScreenplay CreateParsedSnapshot()
    {
        var parsed = SelectedDocument?.CreateParsedSnapshot() ?? new ParsedScreenplay(string.Empty, Array.Empty<ScreenplayElement>());
        if (parsed.TitlePage.Entries.Any(e => e.Label.Equals("Show-Title-Page", StringComparison.OrdinalIgnoreCase) && e.Value.Equals("false", StringComparison.OrdinalIgnoreCase)))
        {
            parsed = new ParsedScreenplay(parsed.RawText, TitlePageData.Empty, parsed.Elements, parsed.LineTypeOverrides);
        }
        return parsed;
    }

    public ParsedScreenplay GetLatestParsedSnapshot()
    {
        return SelectedDocument?.GetLatestParsedSnapshot() ?? new ParsedScreenplay(string.Empty, Array.Empty<ScreenplayElement>());
    }

    public void RefreshParsedSnapshotNow() => SelectedDocument?.RefreshParsedSnapshotNow();

    public void SynchronizeScratchpadWithUndoRedo(string? documentText) => SelectedDocument?.SynchronizeScratchpadWithUndoRedo(documentText);

    public void NewDocument()
    {
        AddDocument(new MainWindowViewModel
        {
            EditorZoomPercent = EditorZoomPercent
        });
        ScheduleSessionSave();
    }

    public void LoadDocument(string text, string? filePath, bool isDirty = false)
    {
        OpenDocument(text, filePath, isDirty);
    }

    public void OpenDocument(string text, string? filePath, bool isDirty = false)
    {
        var document = new MainWindowViewModel
        {
            EditorZoomPercent = EditorZoomPercent
        };
        document.LoadDocument(text, filePath, isDirty);
        AddDocument(document);
        ScheduleSessionSave();
    }

    public bool CloseSelectedDocument()
    {
        return CloseDocument(SelectedDocument);
    }

    public bool CloseDocument(MainWindowViewModel? document)
    {
        if (document is null)
        {
            return false;
        }

        var index = Documents.IndexOf(document);
        if (index < 0)
        {
            return false;
        }

        if (ReferenceEquals(SelectedDocument, document))
        {
            if (Documents.Count > 1)
            {
                var replacementIndex = index < Documents.Count - 1 ? index + 1 : index - 1;
                SelectedDocument = Documents[replacementIndex];
            }
            else
            {
                SelectedDocument = null;
            }
        }

        document.StopRecoveryAutosave();
        Documents.RemoveAt(index);

        if (Documents.Count == 0)
        {
            AddDocument(new MainWindowViewModel());
        }

        return true;
    }

    public void MarkSaved()
    {
        if (SelectedDocument is null)
        {
            return;
        }

        SelectedDocument.MarkSaved();
        ScheduleSessionSave();
    }

    public void SaveCurrentDocument(string path)
    {
        if (SelectedDocument is null)
        {
            return;
        }

        SelectedDocument.SetFilePath(path);
        SelectedDocument.MarkSaved();
        ScheduleSessionSave();
    }

    public void SetFilePath(string? filePath)
    {
        if (SelectedDocument is null)
        {
            return;
        }

        SelectedDocument.SetFilePath(filePath);
        ScheduleSessionSave();
    }

    public void UpdateCaretContext(int lineNumber) => SelectedDocument?.UpdateCaretContext(lineNumber);

    public void UpdateEnterContinuation(int lineNumber, string currentLineText) => SelectedDocument?.UpdateEnterContinuation(lineNumber, currentLineText);

    public void CycleCurrentLineType(bool forward) => SelectedDocument?.CycleCurrentLineType(forward);

    public void SetLineTypeOverride(int lineNumber, ScreenplayElementType elementType) => SelectedDocument?.SetLineTypeOverride(lineNumber, elementType);

    public void SetAutomaticActionLineOverride(int lineNumber) => SelectedDocument?.SetAutomaticActionLineOverride(lineNumber);

    public bool ReleaseAutomaticActionLineOverrideIfNeeded(int lineNumber, string currentLineText)
        => SelectedDocument?.ReleaseAutomaticActionLineOverrideIfNeeded(lineNumber, currentLineText) ?? false;

    public void ClearLineTypeOverride(int lineNumber) => SelectedDocument?.ClearLineTypeOverride(lineNumber);

    public bool HasLineTypeOverride(int lineNumber) => SelectedDocument?.HasLineTypeOverride(lineNumber) ?? false;

    public void ShiftLineTypeOverrides(int startingLineNumber, int delta) => SelectedDocument?.ShiftLineTypeOverrides(startingLineNumber, delta);

    public void StartRecoveryAutosave() => SelectedDocument?.StartRecoveryAutosave();

    public void StopRecoveryAutosave() => SelectedDocument?.StopRecoveryAutosave();

    public void StartGoalTimer() => SelectedDocument?.StartGoalTimer();

    public void PauseGoalTimer() => SelectedDocument?.PauseGoalTimer();

    public void ResumeGoalTimer() => SelectedDocument?.ResumeGoalTimer();

    public void StopGoalTimer() => SelectedDocument?.StopGoalTimer();

    public void ResetGoalTimer() => SelectedDocument?.ResetGoalTimer();

    public void ResetSessionGoal() => SelectedDocument?.ResetSessionGoal();

    public void IncreaseEditorZoom() => SelectedDocument?.IncreaseEditorZoom();

    public void DecreaseEditorZoom() => SelectedDocument?.DecreaseEditorZoom();

    public void ResetEditorZoom() => SelectedDocument?.ResetEditorZoom();

    public void IncreaseBoardZoom() => SelectedDocument?.IncreaseBoardZoom();

    public void DecreaseBoardZoom() => SelectedDocument?.DecreaseBoardZoom();

    public void ResetBoardZoom() => SelectedDocument?.ResetBoardZoom();

    public void SetBoardModeActive(bool isActive)
    {
        if (SelectedDocument is null)
        {
            return;
        }

        SelectedDocument.IsBoardModeActive = isActive;
    }

    public void StopRecoveryAutosaveForAllDocuments()
    {
        foreach (var document in Documents)
        {
            document.StopRecoveryAutosave();
        }
    }

    public void SaveSessionNow()
    {
        _sessionSaveTimer.Stop();

        if (_suppressSessionSave)
        {
            return;
        }

        var documents = Documents
            .Select(document => new SessionDocumentState
            {
                FilePath = document.DocumentPath,
                Text = document.DocumentText,
                IsDirty = document.IsDirty,
                GoalConfiguration = document.GoalConfiguration,
                SessionGoalConfiguration = document.SessionGoalConfiguration,
                EditorZoomPercent = document.EditorZoomPercent
            })
            .ToArray();

        var selectedIndex = SelectedDocument is null ? -1 : Documents.IndexOf(SelectedDocument);
        SessionStorage.SaveSession(new SessionState(documents, selectedIndex));
    }

    public void RestoreSession(RecoveryDocument? recoveredDocument)
    {
        _suppressSessionSave = true;
        try
        {
            MainWindowViewModel? recoveredTarget = null;

            Documents.Clear();

            if (SessionStorage.TryLoadSession(out var state) && state is not null)
            {
                foreach (var documentState in state.Documents)
                {
                    AddRestoredDocument(documentState, select: false);
                }

                if (state.SelectedIndex >= 0 && state.SelectedIndex < Documents.Count)
                {
                    SelectedDocument = Documents[state.SelectedIndex];
                }
            }

            if (recoveredDocument is not null)
            {
                recoveredTarget = TryApplyRecoveredDocument(recoveredDocument);
            }

            if (Documents.Count == 0)
            {
                AddDocument(new MainWindowViewModel());
            }

            if (recoveredTarget is not null)
            {
                SelectedDocument = recoveredTarget;
            }
            else if (SelectedDocument is null && Documents.Count > 0)
            {
                SelectedDocument = Documents[0];
            }
        }
        finally
        {
            _suppressSessionSave = false;
        }

        ScheduleSessionSave();
    }

    public void AddRecoveredDocument(RecoveryDocument recoveryDocument)
    {
        TryApplyRecoveredDocument(recoveryDocument);
        ScheduleSessionSave();
    }

    private static readonly ObservableCollection<OutlineNodeViewModel> _emptyOutline = new();
    private static readonly ObservableCollection<OutlineNodeViewModel> _emptyNotes = new();
    private static readonly ObservableCollection<ScreenplayElement> _emptyBoard = new();
    private static readonly ObservableCollection<ScreenplayElement> _emptyScratchpad = new();
    private static readonly ObservableCollection<TitlePageEntry> _emptyTitlePage = new();
    private static readonly ObservableCollection<PreviewElementItem> _emptyPreview = new();

    private void AddRestoredDocument(SessionDocumentState state, bool select = true)
    {
        var document = new MainWindowViewModel();
        if (!state.IsDirty && !string.IsNullOrWhiteSpace(state.FilePath) && File.Exists(state.FilePath))
        {
            document.LoadDocument(File.ReadAllText(state.FilePath), state.FilePath, false);
        }
        else
        {
            document.LoadDocument(state.Text, state.FilePath, state.IsDirty);
        }

        document.EditorZoomPercent = state.EditorZoomPercent;
        var goalConfiguration = NormalizeOverallGoalConfiguration(state.GoalConfiguration);
        var sessionGoalConfiguration = NormalizeSessionGoalConfiguration(state.GoalConfiguration, state.SessionGoalConfiguration);

        document.ApplyGoalConfiguration(goalConfiguration);
        document.ApplySessionGoalConfiguration(sessionGoalConfiguration);

        AddDocument(document, select);
    }

    private void AddDocument(MainWindowViewModel document, bool select = true)
    {
        Documents.Add(document);

        if (select)
        {
            SelectedDocument = document;
        }
    }

    private MainWindowViewModel? TryApplyRecoveredDocument(RecoveryDocument recoveryDocument)
    {
        var matchingDocument = FindDocumentByPath(recoveryDocument.FilePath);
        if (matchingDocument is not null)
        {
            matchingDocument.LoadDocument(recoveryDocument.Text, recoveryDocument.FilePath, true);
            matchingDocument.EditorZoomPercent = recoveryDocument.EditorZoomPercent;
            var goalConfiguration = NormalizeOverallGoalConfiguration(recoveryDocument.GoalConfiguration);
            var sessionGoalConfiguration = NormalizeSessionGoalConfiguration(recoveryDocument.GoalConfiguration, recoveryDocument.SessionGoalConfiguration);

            matchingDocument.ApplyGoalConfiguration(goalConfiguration);
            matchingDocument.ApplySessionGoalConfiguration(sessionGoalConfiguration);
            return matchingDocument;
        }

        var document = new MainWindowViewModel();
        document.LoadDocument(recoveryDocument.Text, recoveryDocument.FilePath, true);
        document.EditorZoomPercent = recoveryDocument.EditorZoomPercent;
        var normalizedGoalConfiguration = NormalizeOverallGoalConfiguration(recoveryDocument.GoalConfiguration);
        var normalizedSessionGoalConfiguration = NormalizeSessionGoalConfiguration(recoveryDocument.GoalConfiguration, recoveryDocument.SessionGoalConfiguration);

        document.ApplyGoalConfiguration(normalizedGoalConfiguration);
        document.ApplySessionGoalConfiguration(normalizedSessionGoalConfiguration);
        AddDocument(document);
        return document;
    }

    private static GoalConfiguration NormalizeOverallGoalConfiguration(GoalConfiguration configuration)
    {
        return configuration.SelectedGoalType == GoalType.Timer
            ? configuration with { SelectedGoalType = GoalType.WordCount }
            : configuration;
    }

    private static SessionGoalConfiguration NormalizeSessionGoalConfiguration(
        GoalConfiguration overallConfiguration,
        SessionGoalConfiguration sessionConfiguration)
    {
        if (overallConfiguration.SelectedGoalType != GoalType.Timer ||
            sessionConfiguration.SelectedGoalType == GoalType.Timer)
        {
            return sessionConfiguration;
        }

        return sessionConfiguration with
        {
            SelectedGoalType = GoalType.Timer,
            TimerTargetMinutes = overallConfiguration.TimerTargetMinutes
        };
    }

    private MainWindowViewModel? FindDocumentByPath(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return null;
        }

        foreach (var document in Documents)
        {
            if (string.IsNullOrWhiteSpace(document.DocumentPath))
            {
                continue;
            }

            if (string.Equals(document.DocumentPath, filePath, StringComparison.OrdinalIgnoreCase))
            {
                return document;
            }
        }

        return null;
    }

    private void Documents_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        ScheduleSessionSave();
    }

    private void SelectedDocument_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!ReferenceEquals(sender, SelectedDocument))
        {
            return;
        }

        if (e.PropertyName == nameof(MainWindowViewModel.EditorZoomPercent))
        {
            RaiseEditorZoomProperties();
            ScheduleSessionSave();
            return;
        }

        if (e.PropertyName == nameof(MainWindowViewModel.BoardZoomPercent))
        {
            RaiseBoardZoomProperties();
            ScheduleSessionSave();
            return;
        }

        OnPropertyChanged(e.PropertyName);

        if (e.PropertyName is nameof(MainWindowViewModel.GoalProgressSummaryText)
            or nameof(MainWindowViewModel.OutlineRoots)
            or nameof(MainWindowViewModel.NotesRoots)
            or nameof(MainWindowViewModel.BoardElements)
            or nameof(MainWindowViewModel.ScratchpadElements)
            or nameof(MainWindowViewModel.TitlePageEntries)
            or nameof(MainWindowViewModel.PreviewElements))
        {
            RefreshStatusState();
        }

        if (e.PropertyName is nameof(MainWindowViewModel.DocumentText)
            or nameof(MainWindowViewModel.DocumentPath)
            or nameof(MainWindowViewModel.IsDirty)
            or nameof(MainWindowViewModel.SelectedGoalType)
            or nameof(MainWindowViewModel.GoalTargetValue)
            or nameof(MainWindowViewModel.SessionSelectedGoalType)
            or nameof(MainWindowViewModel.SessionGoalTargetValue)
            or nameof(MainWindowViewModel.OutlineRoots)
            or nameof(MainWindowViewModel.BoardElements)
            or nameof(MainWindowViewModel.ScratchpadElements)
            or nameof(MainWindowViewModel.TitlePageEntries)
            or nameof(MainWindowViewModel.PreviewElements))
        {
            ScheduleSessionSave();
        }

        if (e.PropertyName is nameof(MainWindowViewModel.BoardElements)
            or nameof(MainWindowViewModel.IsBoardSyncRequired))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(BoardElements)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasBoardItems)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBoardSyncRequired)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CreateNewCardCommand)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SyncBoardToScriptCommand)));
        }

        if (e.PropertyName == nameof(MainWindowViewModel.IsBoardModeActive))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBoardModeActive)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ActiveZoomDisplayText)));
        }

        if (e.PropertyName is nameof(MainWindowViewModel.ScratchpadElements)
            or nameof(MainWindowViewModel.ScratchpadSearchText)
            or nameof(MainWindowViewModel.SelectedScratchpadElement))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadElements)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadSearchText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedScratchpadElement)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasScratchpadItems)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadEmptyMessage)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DeleteScratchpadCardCommand)));
        }
    }

    private void SelectedDocument_OutlineUpdated(object? sender, EventArgs e)
    {
        if (!ReferenceEquals(sender, SelectedDocument))
        {
            return;
        }

        RefreshStatusState();
        OutlineUpdated?.Invoke(this, EventArgs.Empty);
    }

    private void ScheduleSessionSave()
    {
        if (_suppressSessionSave)
        {
            return;
        }

        _sessionSaveTimer.Stop();
        _sessionSaveTimer.Start();
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        if (string.IsNullOrEmpty(propertyName))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DocumentText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DocumentPath)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsDirty)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WindowTitle)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OutlineRoots)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NotesRoots)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(BoardElements)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasBoardItems)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBoardSyncRequired)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBoardModeActive)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadElements)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadSearchText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedScratchpadElement)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasScratchpadItems)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ScratchpadEmptyMessage)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MoveToScratchpadCommand)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DeleteScratchpadCardCommand)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CreateNewCardCommand)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SyncBoardToScriptCommand)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TitlePageEntries)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PreviewElements)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTypes)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalTypes)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedGoalType)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTargetValue)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalProgressPercent)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalCurrentDisplayText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTargetDisplayText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalStateText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTargetUnitLabel)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionSelectedGoalType)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalTargetValue)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalProgressPercent)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalCurrentDisplayText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalTargetDisplayText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalStateText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SessionGoalTargetUnitLabel)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTimerElapsedText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(GoalTimerRemainingText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsTimerGoal)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SelectedOutlineLineNumber)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentLineNumber)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentElementType)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentElementText)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EnterContinuationText)));
            RaiseEditorZoomProperties();
            RefreshStatusState();
            return;
        }

        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void RefreshStatusState()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasOutlineItems)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasNoteItems)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasBoardItems)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasScratchpadItems)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasTitlePageEntries)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasPreviewContent)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StatusMessage)));
    }

    private void RaiseEditorZoomProperties()
    {
        OnPropertyChanged(nameof(EditorZoomPercent));
        OnPropertyChanged(nameof(EditorZoomScale));
        OnPropertyChanged(nameof(EditorZoomDisplayText));
        OnPropertyChanged(nameof(ActiveZoomDisplayText));
    }

    private void RaiseBoardZoomProperties()
    {
        OnPropertyChanged(nameof(BoardZoomPercent));
        OnPropertyChanged(nameof(BoardZoomScale));
        OnPropertyChanged(nameof(BoardZoomDisplayText));
        OnPropertyChanged(nameof(ActiveZoomDisplayText));
    }
}
