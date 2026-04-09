using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.ComponentModel;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Interop;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using Passage.App.Services;
using Passage.App.Utilities;
using Passage.App.ViewModels;
using Passage.App.Visuals;
using Passage.Core;
using Passage.Export;
using Passage.Parser;

namespace Passage.App;

public partial class MainWindow : Window
{
    private readonly record struct FormattingSpan(int Start, int Length);

    private sealed record ParagraphFormattingResult(
        int LineNumber,
        string Text,
        ScreenplayElementType ScreenplayType,
        IReadOnlyList<FormattingSpan> BoldSpans);

    private sealed record ParagraphFormattingSnapshot(
        int LineNumber,
        string Text,
        TextAlignment TextAlignment,
        double MarginLeft,
        ScreenplayElementType? OverrideType);

    private sealed record FormattingComputationResult(
        string SnapshotText,
        IReadOnlyList<ParagraphFormattingResult> Paragraphs);

    private readonly record struct FormattedRunSegment(string Text, bool IsBold);
    private readonly record struct BeatBoardDropLocation(
        ScreenplayElement? TargetElement,
        bool InsertAfter,
        BeatBoardDropOperation Operation,
        ScreenplayElement? DisplacementTarget);
    private readonly record struct BeatBoardItemLayout(
        ScreenplayElement Element,
        Rect Bounds);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;

        public int Y;
    }

    private enum WriteMode
    {
        Screenplay,
        Markdown
    }

    private enum WorkspaceSurface
    {
        Editor,
        BeatBoard
    }

    private enum BeatBoardViewMode
    {
        Board,
        Outline
    }

    private enum BeatBoardDropOperation
    {
        Insert,
        Nest
    }

    private const double PreferredStartupWidth = 1320.0;
    private const double PreferredStartupHeight = 800.0;
    private const double WorkAreaPadding = 32.0;
    private const double LeftDockExpandedWidth = 245.0;
    private const double SyntaxQuickReferenceMinimumWidth = 280.0;
    private const double SyntaxQuickReferenceDefaultWidth = 320.0;
    private const double ParagraphStyleTolerance = 0.5;
    private const int EditorFormattingDelayMilliseconds = 50;
    private const int EditorDocumentSyncDelayMilliseconds = 180;
    private const double OutlineBeatBoardDragGhostMaxWidth = 420.0;
    private const double OutlineBeatBoardDragGhostMaxHeight = 148.0;
    private const double OutlineBeatBoardDragGhostMaxHotspotX = 160.0;
    private const double BeatBoardVerticalWheelStep = 36.0;
    private const double BeatBoardHorizontalWheelStep = 56.0;
    private const double MinimumEditorPageWidth = 1.0;
    private static readonly Regex MarkdownBoldRegex = new(@"\*\*(.+?)\*\*", RegexOptions.Compiled | RegexOptions.Multiline);
    private static readonly Regex IdCommentRegex = new(@"\s*\[\[id:([a-f\d\-]+)\]\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Thickness ActionParagraphMargin = new(0.0);
    private static readonly Thickness DialogueParagraphMargin = new(96.0, 0.0, 0.0, 0.0);
    private static readonly Thickness ParentheticalParagraphMargin = new(144.0, 0.0, 0.0, 0.0);
    private static readonly Thickness MarkdownDocumentPagePadding = new(50.0);
    private static readonly Thickness ScreenplayDocumentPagePadding = new(0.0);

    private ShellViewModel ViewModel => (ShellViewModel)DataContext;
    private bool _suppressCaretContextUpdates;
    private bool _suppressOutlineSelectionNavigation;
    private bool _editorViewportRefreshPending;
    private bool _ensureCaretVisibleAfterRefresh;
    private bool _editorInteractionRefreshPending;
    private bool _ensureCaretVisibleAfterInteractionRefresh;
    private bool _ensureCaretVisibleAfterCueLayoutRefresh;
    private bool _updateEditorDocumentHeightAfterInteractionRefresh;
    private bool _editorCueInvalidatePending;
    private bool _editorWheelScrollAnimationPending;
    private bool _editorTextChangedProcessingPending;
    private bool _suppressThemeSelectionChanged;
    private bool _isLeftDockCollapsed;
    private bool _isSyntaxQuickReferenceVisible;
    private bool _isFormatting;
    private bool _isSynchronizingEditorDocument;
    private bool _isEditorMouseSelectionActive;
    private bool _sentenceCapitalizationCheckPending;
    private Vector _pendingEditorWheelDelta;
    private const double EditorWheelScrollMultiplier = 0.6;
    private const int WM_MOUSEHWHEEL = 0x020E;
    private bool _hasPendingZoomViewportAnchor;
    private double _pendingZoomViewportAnchorSourceScale;
    private double _pendingZoomViewportAnchorSourceHorizontalOffset;
    private double _pendingZoomViewportAnchorSourceVerticalOffset;
    private double _pendingZoomViewportAnchorSourceViewportWidth;
    private double _pendingZoomViewportAnchorSourceViewportHeight;
    private double _pendingZoomViewportAnchorRatioX = 0.5;
    private double _pendingZoomViewportAnchorRatioY = 0.5;
    private EditorCueAdorner? _editorCueAdorner;
    private HwndSource? _hwndSource;
    private FindReplaceDialog? _findReplaceDialog;
    private GoToLineDialog? _goToLineDialog;
    private GoToSceneDialog? _goToSceneDialog;
    private readonly DispatcherTimer _editorFormattingTimer;
    private readonly DispatcherTimer _editorDocumentSyncTimer;
    private readonly HashSet<int> _pendingPreviewFormattingLineNumbers = [];
    private readonly HashSet<int> _pendingForcedFormattingLineNumbers = [];
    private CancellationTokenSource? _editorPreviewFormattingCancellation;
    private int _editorPreviewFormattingVersion;
    private CancellationTokenSource? _editorFormattingCancellation;
    private int _editorFormattingVersion;
    private bool _editorTextSyncPending;
    private bool _formatEntireDocumentRequested;
    private bool _suppressEditorDocumentReload;
    private string _lastCommittedEditorText = string.Empty;
    private Paragraph? _lastTextChangedParagraph;
    private string _lastTextChangedParagraphText = string.Empty;
    private int? _editorMouseSelectionAnchorIndex;
    private double _syntaxQuickReferenceWidth = SyntaxQuickReferenceDefaultWidth;
    private WriteMode currentMode = WriteMode.Screenplay;
    private WorkspaceSurface _currentWorkspaceSurface = WorkspaceSurface.Editor;
    private BeatBoardViewMode _currentBeatBoardViewMode = BeatBoardViewMode.Board;
    private Point? _beatBoardDragStartPoint;
    private ScreenplayElement? _beatBoardDragSourceElement;
    private FrameworkElement? _beatBoardDragGhostElement;
    private bool _beatBoardDragWasInitiated;
    private Point _beatBoardDragPointerOffset = new(140.0, 80.0);
    private Size _beatBoardDragSourceSize = new(280.0, 140.0);
    private Point _beatBoardDragGhostHotspot = new(140.0, 80.0);
    private readonly DispatcherTimer _beatBoardClickTimer;
    private ScreenplayElement? _pendingBeatBoardClickElement;
    private FrameworkElement? _currentInlineEditHost;
    private ScreenplayElement? _currentInlineEditElement;
    private bool _suppressNextBeatBoardCardClick;
    private bool _suppressNextBeatBoardCardMouseUp;
    private bool _beatBoardRefreshPending;
    private MainWindowViewModel? _observedBeatBoardDocument;
    private bool IsScreenplayMode => currentMode == WriteMode.Screenplay;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out NativePoint point);

    [DllImport("user32.dll")]
    private static extern uint GetDoubleClickTime();

    public MainWindow(RecoveryDocument? recoveredDocument = null)
    {
        var viewModel = new ShellViewModel(recoveredDocument);
        _editorFormattingTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(EditorFormattingDelayMilliseconds)
        };
        _editorFormattingTimer.Tick += EditorFormattingTimer_Tick;
        _editorDocumentSyncTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(EditorDocumentSyncDelayMilliseconds)
        };
        _editorDocumentSyncTimer.Tick += EditorDocumentSyncTimer_Tick;
        _beatBoardClickTimer = new DispatcherTimer(DispatcherPriority.Input, Dispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(GetDoubleClickTime())
        };
        _beatBoardClickTimer.Tick += BeatBoardClickTimer_Tick;
        DataContext = viewModel;
        InitializeComponent();
        UpdateBeatBoardViewMode();
        UpdateWorkspaceSurface();
        RegisterEditorMouseHandlers();
        viewModel.PropertyChanged += ViewModel_PropertyChanged;
        viewModel.OutlineUpdated += ViewModel_OutlineUpdated;
        AttachBeatBoardDocumentHandlers(viewModel.SelectedDocument);
        ThemeManager.ThemeChanged += ThemeManager_ThemeChanged;
        InitializeThemeSelector();
        Loaded += MainWindow_Loaded;
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        _hwndSource = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
        _hwndSource?.AddHook(WindowMessageHook);
        ApplyStartupBounds();
    }

    private void RegisterEditorMouseHandlers()
    {
        if (EditorBox is null)
        {
            return;
        }

        EditorBox.AddHandler(
            UIElement.PreviewMouseLeftButtonDownEvent,
            new MouseButtonEventHandler(EditorBox_PreviewMouseLeftButtonDown),
            handledEventsToo: true);
        EditorBox.AddHandler(
            UIElement.PreviewMouseMoveEvent,
            new MouseEventHandler(EditorBox_PreviewMouseMove),
            handledEventsToo: true);
        EditorBox.AddHandler(
            UIElement.PreviewMouseRightButtonDownEvent,
            new MouseButtonEventHandler(EditorBox_PreviewMouseRightButtonDown),
            handledEventsToo: true);
        EditorBox.AddHandler(
            UIElement.PreviewMouseLeftButtonUpEvent,
            new MouseButtonEventHandler(EditorBox_PreviewMouseLeftButtonUp),
            handledEventsToo: true);
    }

    private void BeatBoard_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _beatBoardDragWasInitiated = false;
        _suppressNextBeatBoardCardClick = false;
        ClearPendingBeatBoardClick();
        _beatBoardDragStartPoint = BeatBoard is null ? null : e.GetPosition(BeatBoard);
        var source = e.OriginalSource as DependencyObject;
        var itemContainer = FindAncestor<ContentPresenter>(source);
        _beatBoardDragSourceElement = itemContainer?.DataContext as ScreenplayElement;

        if (itemContainer is not null &&
            itemContainer.RenderSize.Width > 0.5 &&
            itemContainer.RenderSize.Height > 0.5)
        {
            var pointWithinCard = e.GetPosition(itemContainer);
            _beatBoardDragPointerOffset = new Point(
                Math.Clamp(pointWithinCard.X, 0.0, itemContainer.RenderSize.Width),
                Math.Clamp(pointWithinCard.Y, 0.0, itemContainer.RenderSize.Height));
            _beatBoardDragSourceSize = new Size(
                Math.Max(252.0, itemContainer.RenderSize.Width),
                Math.Max(96.0, itemContainer.RenderSize.Height));
        }
        else
        {
            _beatBoardDragPointerOffset = new Point(140.0, 80.0);
            _beatBoardDragSourceSize = new Size(280.0, 140.0);
        }

        ApplyBeatBoardDragPreviewConstraints();
    }

    private void BeatBoard_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        ClearBeatBoardDragTracking();
    }

    private void BeatBoard_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (BeatBoard is null ||
            _beatBoardDragStartPoint is null ||
            _beatBoardDragSourceElement is null ||
            e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        var currentPosition = e.GetPosition(BeatBoard);
        if (Math.Abs(currentPosition.X - _beatBoardDragStartPoint.Value.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(currentPosition.Y - _beatBoardDragStartPoint.Value.Y) < SystemParameters.MinimumVerticalDragDistance)
        {
            return;
        }

        var draggedElement = _beatBoardDragSourceElement;
        var blockCount = GetBeatBoardBlockCount(draggedElement);
        _beatBoardDragWasInitiated = true;
        ClearBeatBoardDragTracking();
        StartBeatBoardDragGhost(draggedElement, blockCount, _beatBoardDragSourceSize);

        try
        {
            DragDrop.DoDragDrop(BeatBoard, new DataObject(typeof(ScreenplayElement), draggedElement), DragDropEffects.Move);
        }
        finally
        {
            ViewModel.SelectedDocument?.ClearBoardDropIndicator();
            StopBeatBoardDragGhost();
        }
    }

    private void BeatBoard_DragOver(object sender, DragEventArgs e)
    {
        UpdateBeatBoardDragGhostPosition();

        // Handle auto-scroll based on cursor proximity to viewport boundaries
        if (BeatBoardScrollViewer != null)
        {
            var relativePoint = e.GetPosition(BeatBoardScrollViewer);
            const double threshold = 40.0;
            double vDelta = 0;
            double hDelta = 0;

            if (relativePoint.Y < threshold && relativePoint.Y >= 0)
            {
                vDelta = -((threshold - relativePoint.Y) / threshold) * BeatBoardVerticalWheelStep;
            }
            else if (relativePoint.Y > BeatBoardScrollViewer.ActualHeight - threshold && relativePoint.Y <= BeatBoardScrollViewer.ActualHeight)
            {
                vDelta = ((relativePoint.Y - (BeatBoardScrollViewer.ActualHeight - threshold)) / threshold) * BeatBoardVerticalWheelStep;
            }

            if (relativePoint.X < threshold && relativePoint.X >= 0)
            {
                hDelta = -((threshold - relativePoint.X) / threshold) * BeatBoardHorizontalWheelStep;
            }
            else if (relativePoint.X > BeatBoardScrollViewer.ActualWidth - threshold && relativePoint.X <= BeatBoardScrollViewer.ActualWidth)
            {
                hDelta = ((relativePoint.X - (BeatBoardScrollViewer.ActualWidth - threshold)) / threshold) * BeatBoardHorizontalWheelStep;
            }

            if (Math.Abs(vDelta) > 0.5 || Math.Abs(hDelta) > 0.5)
            {
                ScrollBeatBoard(BeatBoardScrollViewer, vDelta, hDelta);
                UpdateBeatBoardDragGhostPosition();
            }
        }

        if (!TryGetBeatBoardDraggedElement(e, out var draggedElement))
        {
            ViewModel.SelectedDocument?.CancelReorderPreview();
            ViewModel.SelectedDocument?.ClearBoardDropIndicator();
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (TryGetBeatBoardDropLocation(e.GetPosition(BeatBoard), draggedElement, out var dropLocation))
        {
            if (dropLocation.Operation == BeatBoardDropOperation.Nest)
            {
                ViewModel.SelectedDocument?.CancelReorderPreview();
                ViewModel.SelectedDocument?.SetBoardDropIndicator(dropLocation.TargetElement);
            }
            else if (dropLocation.TargetElement != null)
            {
                ViewModel.SelectedDocument?.SetBoardDropIndicator(dropLocation.DisplacementTarget);
                ViewModel.SelectedDocument?.PerformReorderPreview(draggedElement, dropLocation.TargetElement, dropLocation.InsertAfter);
            }
            
            e.Effects = DragDropEffects.Move;
        }
        else
        {
            ViewModel.SelectedDocument?.CancelReorderPreview();
            ViewModel.SelectedDocument?.ClearBoardDropIndicator();
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void BeatBoard_DragLeave(object sender, DragEventArgs e)
    {
        ViewModel.SelectedDocument?.CancelReorderPreview();
        ViewModel.SelectedDocument?.ClearBoardDropIndicator();
    }

    private void BeatBoard_Drop(object sender, DragEventArgs e)
    {
        if (!TryGetBeatBoardDraggedElement(e, out var draggedElement))
        {
            ViewModel.SelectedDocument?.CancelReorderPreview();
            ViewModel.SelectedDocument?.ClearBoardDropIndicator();
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (TryGetBeatBoardDropLocation(e.GetPosition(BeatBoard), draggedElement, out var dropLocation))
        {
            if (dropLocation.Operation == BeatBoardDropOperation.Nest)
            {
                ViewModel.SelectedDocument?.CancelReorderPreview();
                var didMove = ViewModel.SelectedDocument?.TryMoveBoardBlockIntoParent(draggedElement, dropLocation.TargetElement) ?? false;
                e.Effects = didMove ? DragDropEffects.Move : DragDropEffects.None;
            }
            else
            {
                ViewModel.SelectedDocument?.FinalizeReorderPreview();
                e.Effects = DragDropEffects.Move;
            }
        }
        else
        {
            ViewModel.SelectedDocument?.CancelReorderPreview();
            e.Effects = DragDropEffects.None;
        }

        ViewModel.SelectedDocument?.ClearBoardDropIndicator();
        e.Handled = true;
    }

    private void BeatBoard_GiveFeedback(object sender, GiveFeedbackEventArgs e)
    {
        UpdateBeatBoardDragGhostPosition();
        e.UseDefaultCursors = true;
        e.Handled = true;
    }

    private void BeatBoard_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
    {
        if (e.EscapePressed)
        {
            ViewModel.SelectedDocument?.CancelReorderPreview();
            ViewModel.SelectedDocument?.ClearBoardDropIndicator();
            SetBeatBoardDeleteButtonState(isActive: false);
            StopBeatBoardDragGhost();
            ClearBeatBoardDragTracking();
        }
    }

    private void BeatBoardDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        _ = DeleteBeatBoardElement(ViewModel.SelectedDocument?.SelectedBoardElement);
    }

    private void BeatBoardDeleteButton_DragOver(object sender, DragEventArgs e)
    {
        UpdateBeatBoardDragGhostPosition();

        var canDelete = TryGetBeatBoardDraggedElement(e, out _);
        ViewModel.SelectedDocument?.ClearBoardDropIndicator();
        SetBeatBoardDeleteButtonState(canDelete);
        e.Effects = canDelete ? DragDropEffects.Move : DragDropEffects.None;
        e.Handled = true;
    }

    private void BeatBoardDeleteButton_DragLeave(object sender, DragEventArgs e)
    {
        SetBeatBoardDeleteButtonState(isActive: false);
    }

    private void BeatBoardDeleteButton_Drop(object sender, DragEventArgs e)
    {
        SetBeatBoardDeleteButtonState(isActive: false);
        ViewModel.SelectedDocument?.ClearBoardDropIndicator();

        var didDelete = TryGetBeatBoardDraggedElement(e, out var draggedElement) &&
            DeleteBeatBoardElement(draggedElement);
        e.Effects = didDelete ? DragDropEffects.Move : DragDropEffects.None;
        e.Handled = true;
    }

    private void BeatBoardTypeSelector_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement selectorHost ||
            selectorHost.ContextMenu is null ||
            !ShouldOpenBeatBoardTypeSelector(selectorHost, e))
        {
            return;
        }

        _suppressNextBeatBoardCardClick = true;
        ClearPendingBeatBoardClick();

        if (selectorHost.DataContext is ScreenplayElement element)
        {
            ViewModel.SelectedDocument?.SelectBoardElement(element);
        }

        selectorHost.ContextMenu.PlacementTarget = selectorHost;
        selectorHost.ContextMenu.IsOpen = true;
        e.Handled = true;
    }

    private static bool ShouldOpenBeatBoardTypeSelector(FrameworkElement selectorHost, MouseButtonEventArgs e)
    {
        if (selectorHost is TextBlock textBlock)
        {
            var arrowHitWidth = Math.Min(16.0, Math.Max(10.0, textBlock.ActualWidth * 0.3));
            return e.GetPosition(textBlock).X <= arrowHitWidth;
        }

        if (selectorHost is System.Windows.Shapes.Path)
        {
            return true;
        }

        if (selectorHost is Panel)
        {
            var source = e.OriginalSource as DependencyObject;
            return source is System.Windows.Shapes.Path || FindAncestor<System.Windows.Shapes.Path>(source) is not null;
        }

        return true;
    }

    private void BeatBoardCollapseButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        ClearPendingBeatBoardClick();
    }

    private void BeatBoardCollapseButton_Click(object sender, RoutedEventArgs e)
    {
        ScheduleBeatBoardRefresh();
    }

    private void BeatBoardActionButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        ClearPendingBeatBoardClick();
    }

    private void BeatBoardBoardViewButton_Click(object sender, RoutedEventArgs e)
    {
        SetBeatBoardViewMode(BeatBoardViewMode.Board);
    }

    private void BeatBoardOutlineViewButton_Click(object sender, RoutedEventArgs e)
    {
        SetBeatBoardViewMode(BeatBoardViewMode.Outline);
    }

    private void BeatBoardTypeMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem menuItem ||
            string.IsNullOrWhiteSpace(menuItem.Tag?.ToString()))
        {
            return;
        }

        var element = ((menuItem.Parent as ContextMenu)?.PlacementTarget as FrameworkElement)?.DataContext as ScreenplayElement;
        var updatedElement = ViewModel.SelectedDocument?.TrySetBoardElementKind(element, menuItem.Tag?.ToString());
        if (updatedElement is not null && _currentInlineEditElement is not null && ReferenceEquals(_currentInlineEditElement, element))
        {
            _currentInlineEditElement = updatedElement;
            if (_currentInlineEditHost is not null)
            {
                PopulateInlineBeatBoardEditorFields(_currentInlineEditHost, updatedElement);
            }
        }
    }

    private void BeatBoardCard_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (FindAncestor<Button>(e.OriginalSource as DependencyObject) is not null)
        {
            return;
        }

        if (sender is not FrameworkElement elementHost ||
            elementHost.DataContext is not ScreenplayElement element)
        {
            return;
        }

        ViewModel.SelectedDocument?.SelectBoardElement(element);

        if (e.ClickCount > 1)
        {
            ClearPendingBeatBoardClick();
            _suppressNextBeatBoardCardMouseUp = true;
            RevealBoardElementInEditor(element);
            e.Handled = true;
        }
    }

    private void BeatBoardCard_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_beatBoardDragWasInitiated)
        {
            _beatBoardDragWasInitiated = false;
            return;
        }

        if (_suppressNextBeatBoardCardMouseUp)
        {
            _suppressNextBeatBoardCardMouseUp = false;
            _beatBoardDragWasInitiated = false;
            e.Handled = true;
            return;
        }

        if (FindAncestor<Button>(e.OriginalSource as DependencyObject) is not null)
        {
            _beatBoardDragWasInitiated = false;
            return;
        }

        if (sender is not FrameworkElement elementHost ||
            elementHost.DataContext is not ScreenplayElement element)
        {
            return;
        }

        if (_suppressNextBeatBoardCardClick)
        {
            _suppressNextBeatBoardCardClick = false;
            _beatBoardDragWasInitiated = false;
            return;
        }

        _beatBoardDragWasInitiated = false;
        ViewModel.SelectedDocument?.SelectBoardElement(element);

        QueueBeatBoardSingleClickAction(element);
        e.Handled = true;
    }

    private void BeatBoardSingleClickTimer_Tick()
    {
        if (_pendingBeatBoardClickElement is not ScreenplayElement element)
        {
            return;
        }

        _pendingBeatBoardClickElement = null;
        ViewModel.SelectedDocument?.SelectBoardElement(element);
        if (ViewModel.SelectedDocument?.TrySetBoardElementCollapsed(element, !element.IsCollapsed) == true)
        {
            ScheduleBeatBoardRefresh();
        }
    }

    private void BeatBoardEditButton_Click(object sender, RoutedEventArgs e)
    {
        ClearPendingBeatBoardClick();

        if (sender is not FrameworkElement buttonHost ||
            buttonHost.DataContext is not ScreenplayElement element)
        {
            return;
        }

        ViewModel.SelectedDocument?.SelectBoardElement(element);

        // Commit any existing edit before starting a new one
        if (_currentInlineEditElement is not null && !ReferenceEquals(_currentInlineEditElement, element))
        {
            CommitAndCloseInlineEditor();
        }

        // Find the current ContentPresenter that hosts this element. Use the ItemContainerGenerator
        // as it is more reliable than FindAncestor for identifying the correct item container.
        var elementHost = BeatBoard?.ItemContainerGenerator.ContainerFromItem(element) as ContentPresenter;

        if (elementHost is null)
        {
            // Fallback to FindAncestor if Generator fails for some reason
            elementHost = FindAncestor<ContentPresenter>(buttonHost);
        }

        if (elementHost is null)
        {
            return;
        }

        if (GetIsEditing(elementHost))
        {
            CommitAndCloseInlineEditor();
        }
        else
        {
            CommitAndCloseInlineEditor();
            _currentInlineEditHost = elementHost;
            _currentInlineEditElement = element;
            SetIsEditing(elementHost, true);
            PopulateInlineBeatBoardEditorFields(elementHost, element);
        }
        e.Handled = true;
    }

    private void EditorWorkspaceButton_Click(object sender, RoutedEventArgs e)
    {
        SetWorkspaceSurface(WorkspaceSurface.Editor);
    }

    private void BeatBoardWorkspaceButton_Click(object sender, RoutedEventArgs e)
    {
        SetWorkspaceSurface(WorkspaceSurface.BeatBoard);
    }

    private void SetBeatBoardViewMode(BeatBoardViewMode viewMode)
    {
        if (_currentBeatBoardViewMode == viewMode)
        {
            UpdateBeatBoardViewMode();
            return;
        }

        _currentBeatBoardViewMode = viewMode;
        UpdateBeatBoardViewMode();
    }

    private void SetWorkspaceSurface(WorkspaceSurface surface)
    {
        if (_currentWorkspaceSurface == surface)
        {
            UpdateWorkspaceSurface();
            return;
        }

        if (surface == WorkspaceSurface.BeatBoard)
        {
            FlushPendingEditorChangeBatch();
            ViewModel.RefreshParsedSnapshotNow();
        }

        _currentWorkspaceSurface = surface;
        UpdateWorkspaceSurface();
    }

    private void UpdateWorkspaceSurface()
    {
        var isEditorSurface = _currentWorkspaceSurface == WorkspaceSurface.Editor;
        ViewModel.SetBoardModeActive(_currentWorkspaceSurface == WorkspaceSurface.BeatBoard);

        if (WorkspaceTitleText is not null)
        {
            WorkspaceTitleText.Text = _currentWorkspaceSurface == WorkspaceSurface.Editor
                ? "Editor"
                : "Beat Board";
        }

        if (EditorWorkspaceButton is not null)
        {
            EditorWorkspaceButton.IsChecked = _currentWorkspaceSurface == WorkspaceSurface.Editor;
        }

        if (BeatBoardWorkspaceButton is not null)
        {
            BeatBoardWorkspaceButton.IsChecked = _currentWorkspaceSurface == WorkspaceSurface.BeatBoard;
        }

        if (EditorScrollHost is not null)
        {
            EditorScrollHost.Visibility = isEditorSurface
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        if (BeatBoardWorkspaceHost is not null)
        {
            BeatBoardWorkspaceHost.Visibility = !isEditorSurface
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        if (_editorCueAdorner is not null)
        {
            _editorCueAdorner.Visibility = isEditorSurface ? Visibility.Visible : Visibility.Collapsed;
        }

        if (!isEditorSurface)
        {
            _ensureCaretVisibleAfterRefresh = false;
            _ensureCaretVisibleAfterInteractionRefresh = false;
            _ensureCaretVisibleAfterCueLayoutRefresh = false;
            return;
        }

        if (IsLoaded)
        {
            _editorCueAdorner?.NotifyParsedLayoutChanged(immediate: true);
            RestoreEditorFocusAndCaret(ensureFocus: false);
            ScheduleEditorViewportRefresh(ensureCaretVisible: EditorBox is not null && EditorBox.IsKeyboardFocusWithin);
        }
    }

    private void UpdateBeatBoardViewMode()
    {
        CommitAndCloseInlineEditor();
        ViewModel.SelectedDocument?.ClearBoardDropIndicator();
        StopBeatBoardDragGhost();
        ClearBeatBoardDragTracking();
        ClearPendingBeatBoardClick();

        if (BeatBoard is not null)
        {
            var isOutlineView = _currentBeatBoardViewMode == BeatBoardViewMode.Outline;
            BeatBoard.ItemTemplateSelector = (DataTemplateSelector)FindResource(
                isOutlineView
                    ? "BeatBoardOutlineCardTemplateSelector"
                    : "BeatBoardCardTemplateSelector");
            BeatBoard.ItemContainerStyle = (Style)FindResource(
                isOutlineView
                    ? "BeatBoardOutlineItemContainerStyle"
                    : "BeatBoardBoardItemContainerStyle");
            BeatBoard.ItemsPanel = (ItemsPanelTemplate)FindResource(
                isOutlineView
                    ? "BeatBoardOutlineItemsPanelTemplate"
                    : "BeatBoardWrapItemsPanelTemplate");
            BeatBoard.Items.Refresh();
            BeatBoard.InvalidateMeasure();
            BeatBoard.InvalidateArrange();
            UpdateBeatBoardSurfaceLayout();
            ScheduleBeatBoardRefresh();
        }

        if (BeatBoardBoardViewButton is not null)
        {
            BeatBoardBoardViewButton.IsChecked = _currentBeatBoardViewMode == BeatBoardViewMode.Board;
        }

        if (BeatBoardOutlineViewButton is not null)
        {
            BeatBoardOutlineViewButton.IsChecked = _currentBeatBoardViewMode == BeatBoardViewMode.Outline;
        }
    }

    private void BeatBoardSurfaceRoot_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateBeatBoardSurfaceLayout();

        if (!e.WidthChanged)
        {
            return;
        }

        BeatBoard?.InvalidateMeasure();
        BeatBoard?.InvalidateArrange();
        BeatBoardContentHost?.InvalidateMeasure();
        BeatBoardContentHost?.InvalidateArrange();
    }

    private void UpdateBeatBoardSurfaceLayout()
    {
        if (BeatBoardSurfaceRoot is null ||
            BeatBoardContentHost is null ||
            BeatBoardScrollViewer is null)
        {
            return;
        }

        var surfaceWidth = Math.Max(300.0, BeatBoardSurfaceRoot.ActualWidth);
        BeatBoardContentHost.MinWidth = surfaceWidth;

        if (_currentBeatBoardViewMode == BeatBoardViewMode.Outline)
        {
            BeatBoardScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            BeatBoardContentHost.Width = surfaceWidth;
            return;
        }

        BeatBoardScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
        BeatBoardContentHost.Width = double.NaN;
    }

    private void BeatBoardScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not ScrollViewer scrollViewer || e.Delta == 0)
        {
            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            e.Handled = true;
            if (e.Delta > 0)
            {
                ViewModel.IncreaseBoardZoom();
            }
            else
            {
                ViewModel.DecreaseBoardZoom();
            }
            return;
        }

        e.Handled = true;

        // Sync sensitivity with the editor for a consistent trackpad experience
        double step = 16.0;
        if (EditorBox is not null)
        {
            var lineSpacing = 1.2;
            try { lineSpacing = EditorBox.FontFamily.LineSpacing; } catch { }
            step = Math.Max(12.0, EditorBox.FontSize * lineSpacing);
        }

        // Adjust for zoom scale to maintain perceived scroll speed
        double scale = ViewModel.BoardZoomScale;
        double delta = -e.Delta / 120.0 * step * EditorWheelScrollMultiplier * Math.Max(0.01, scale);

        if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0 && _currentBeatBoardViewMode == BeatBoardViewMode.Board)
        {
            scrollViewer.ScrollToHorizontalOffset(Math.Clamp(scrollViewer.HorizontalOffset + delta, 0, scrollViewer.ScrollableWidth));
            return;
        }

        scrollViewer.ScrollToVerticalOffset(Math.Clamp(scrollViewer.VerticalOffset + delta, 0, scrollViewer.ScrollableHeight));
    }

    private void ScrollBeatBoard(ScrollViewer scrollViewer, double verticalDelta = 0.0, double horizontalDelta = 0.0)
    {
        if (Math.Abs(horizontalDelta) > 0.01 && scrollViewer.ScrollableWidth > 0.0)
        {
            var targetHorizontalOffset = Math.Clamp(
                scrollViewer.HorizontalOffset + horizontalDelta,
                0.0,
                scrollViewer.ScrollableWidth);
            scrollViewer.ScrollToHorizontalOffset(targetHorizontalOffset);
        }

        if (Math.Abs(verticalDelta) > 0.01 && scrollViewer.ScrollableHeight > 0.0)
        {
            var targetVerticalOffset = Math.Clamp(
                scrollViewer.VerticalOffset + verticalDelta,
                0.0,
                scrollViewer.ScrollableHeight);
            scrollViewer.ScrollToVerticalOffset(targetVerticalOffset);
        }
    }

    private static double NormalizeWheelSteps(int wheelDelta)
    {
        return Math.Clamp(wheelDelta / 120.0, -3.0, 3.0);
    }

    private void ClearBeatBoardDragTracking()
    {
        _beatBoardDragStartPoint = null;
        _beatBoardDragSourceElement = null;
        ViewModel.SelectedDocument?.ClearBoardDropIndicator();
        SetBeatBoardDeleteButtonState(isActive: false);
    }

    private void AttachBeatBoardDocumentHandlers(MainWindowViewModel? document)
    {
        if (ReferenceEquals(_observedBeatBoardDocument, document))
        {
            return;
        }

        if (_observedBeatBoardDocument is not null)
        {
            _observedBeatBoardDocument.PropertyChanged -= ObservedBeatBoardDocument_PropertyChanged;
        }

        _observedBeatBoardDocument = document;
        if (_observedBeatBoardDocument is not null)
        {
            _observedBeatBoardDocument.PropertyChanged += ObservedBeatBoardDocument_PropertyChanged;
        }
    }

    private void ObservedBeatBoardDocument_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MainWindowViewModel.VisibleBoardElements)
            or nameof(MainWindowViewModel.BoardElements)
            or nameof(MainWindowViewModel.HasVisibleBoardItems))
        {
            ScheduleBeatBoardRefresh();
        }
    }

    private void ScheduleBeatBoardRefresh()
    {
        if (_beatBoardRefreshPending)
        {
            return;
        }

        _beatBoardRefreshPending = true;
        Dispatcher.BeginInvoke(new Action(() =>
        {
            _beatBoardRefreshPending = false;
            RefreshBeatBoardLayout();
        }), DispatcherPriority.Render);
    }

    private void RefreshBeatBoardLayout()
    {
        if (BeatBoard is null)
        {
            return;
        }

        BeatBoard.Items.Refresh();
        BeatBoard.InvalidateMeasure();
        BeatBoard.InvalidateArrange();
        BeatBoard.InvalidateVisual();
        BeatBoard.UpdateLayout();

        if (FindDescendant<Panel>(BeatBoard) is Panel itemsPanel)
        {
            itemsPanel.InvalidateMeasure();
            itemsPanel.InvalidateArrange();
            itemsPanel.InvalidateVisual();
            itemsPanel.UpdateLayout();
        }

        if (BeatBoardContentHost is not null)
        {
            BeatBoardContentHost.InvalidateMeasure();
            BeatBoardContentHost.InvalidateArrange();
            BeatBoardContentHost.InvalidateVisual();
        }
    }

    private bool DeleteBeatBoardElement(ScreenplayElement? element)
    {
        if (element is null)
        {
            return false;
        }

        ClearPendingBeatBoardClick();
        SetBeatBoardDeleteButtonState(isActive: false);

        if (_currentInlineEditElement is not null &&
            (_currentInlineEditElement.Id == element.Id || ReferenceEquals(_currentInlineEditElement, element)))
        {
            if (_currentInlineEditHost is not null)
            {
                SetIsEditing(_currentInlineEditHost, false);
            }
            _currentInlineEditHost = null;
            _currentInlineEditElement = null;
        }

        return ViewModel.SelectedDocument?.TryDeleteBoardBlock(element) ?? false;
    }

    private void SetBeatBoardDeleteButtonState(bool isActive)
    {
        if (BeatBoardDeleteButton is null)
        {
            return;
        }

        BeatBoardDeleteButton.Background = isActive
            ? ResolveBrushResource("DestructiveBackground", new SolidColorBrush(Color.FromRgb(74, 29, 29)))
            : ResolveBrushResource("ControlBackground");
        BeatBoardDeleteButton.BorderBrush = isActive
            ? ResolveBrushResource("DestructiveBorder", new SolidColorBrush(Color.FromRgb(214, 106, 106)))
            : ResolveBrushResource("ControlBorder");
        BeatBoardDeleteButton.Foreground = isActive
            ? ResolveBrushResource("DestructiveForeground", Brushes.White)
            : ResolveBrushResource("ControlForeground");
    }

    private Brush ResolveBrushResource(string resourceKey, Brush fallback)
    {
        return TryFindResource(resourceKey) as Brush ?? fallback;
    }

    private Brush ResolveBrushResource(string resourceKey)
    {
        return ResolveBrushResource(resourceKey, Brushes.Transparent);
    }

    private void QueueBeatBoardSingleClickAction(ScreenplayElement element)
    {
        _pendingBeatBoardClickElement = element;
        _beatBoardClickTimer.Stop();
        _beatBoardClickTimer.Start();
    }

    private void ClearPendingBeatBoardClick()
    {
        _beatBoardClickTimer.Stop();
        _pendingBeatBoardClickElement = null;
        _suppressNextBeatBoardCardMouseUp = false;
    }

    private void BeatBoardClickTimer_Tick(object? sender, EventArgs e)
    {
        _beatBoardClickTimer.Stop();
        BeatBoardSingleClickTimer_Tick();
    }

    public static readonly DependencyProperty IsEditingProperty =
        DependencyProperty.RegisterAttached(
            "IsEditing",
            typeof(bool),
            typeof(MainWindow),
            new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.Inherits));

    public static bool GetIsEditing(DependencyObject d) => (bool)d.GetValue(IsEditingProperty);
    public static void SetIsEditing(DependencyObject d, bool value) => d.SetValue(IsEditingProperty, value);

    public static readonly DependencyProperty IsDropTargetProperty =
        DependencyProperty.RegisterAttached(
            "IsDropTarget",
            typeof(bool),
            typeof(MainWindow),
            new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.Inherits));

    public static bool GetIsDropTarget(DependencyObject d) => (bool)d.GetValue(IsDropTargetProperty);
    public static void SetIsDropTarget(DependencyObject d, bool value) => d.SetValue(IsDropTargetProperty, value);

    private void CommitAndCloseInlineEditor()
    {
        if (_currentInlineEditElement is null)
        {
            return;
        }

        // Try to find the current host for the element, as it might have changed during a refresh
        var host = _currentInlineEditHost;
        if (host is null || !ReferenceEquals(host.DataContext, _currentInlineEditElement))
        {
            TryGetBeatBoardElementHost(_currentInlineEditElement, out var currentHost);
            host = currentHost as ContentPresenter;
        }

        if (host is not null)
        {
            CommitInlineBeatBoardEditorChanges(host, _currentInlineEditElement);
            SetIsEditing(host, false);
        }

        _currentInlineEditHost = null;
        _currentInlineEditElement = null;
    }

    private void PopulateInlineBeatBoardEditorFields(FrameworkElement elementHost, ScreenplayElement element)
    {
        var isSceneCard = element.Type == ScreenplayElementType.SceneHeading;

        if (FindDescendant<TextBox>(elementHost, "InlineHeadingTextBox") is TextBox headingBox)
        {
            headingBox.Text = GetEditableBoardHeading(element);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                headingBox.Focus();
                headingBox.SelectAll();
            }), DispatcherPriority.Input);
        }

        if (FindDescendant<TextBox>(elementHost, "InlineSceneHeadingTextBox") is TextBox sceneHeadingBox)
        {
            sceneHeadingBox.Text = isSceneCard ? GetEditableBoardSceneHeading(element) : string.Empty;
        }

        if (FindDescendant<TextBox>(elementHost, "InlineDescriptionTextBox") is TextBox descriptionBox)
        {
            descriptionBox.Text = GetEditableBoardDescription(element);
        }
    }

    private static string GetEditableBoardHeading(ScreenplayElement element)
    {
        return element switch
        {
            ScratchpadCardElement draftCard => draftCard.Heading,
            _ => element.Heading
        };
    }

    private static string GetEditableBoardSceneHeading(ScreenplayElement element)
    {
        if (element.Type != ScreenplayElementType.SceneHeading)
        {
            return string.Empty;
        }

        return element switch
        {
            ScratchpadCardElement draftCard => draftCard.ScriptHeading,
            _ => element.ScriptHeading
        };
    }

    private static string GetEditableBoardDescription(ScreenplayElement element)
    {
        var description = element switch
        {
            ScratchpadCardElement draftCard => draftCard.Description,
            _ => element.BoardDescription
        };

        if (description is null || MainWindowViewModel.IsDefaultBoardDescription(description))
        {
            return string.Empty;
        }

        return description ?? string.Empty;
    }

    private bool CommitInlineBeatBoardEditorChanges(FrameworkElement elementHost, ScreenplayElement element)
    {
        var headingBox = FindDescendant<TextBox>(elementHost, "InlineHeadingTextBox");
        var sceneHeadingBox = FindDescendant<TextBox>(elementHost, "InlineSceneHeadingTextBox");
        var descriptionBox = FindDescendant<TextBox>(elementHost, "InlineDescriptionTextBox");

        var newHeading = headingBox?.Text ?? GetEditableBoardHeading(element);
        var newSceneHeading = element.Type == ScreenplayElementType.SceneHeading 
            ? (sceneHeadingBox?.Text ?? GetEditableBoardSceneHeading(element))
            : string.Empty;
        var newDescription = descriptionBox?.Text ?? GetEditableBoardDescription(element);

        var isDirty = !string.Equals(newHeading, GetEditableBoardHeading(element), StringComparison.Ordinal) ||
                      !string.Equals(newDescription, GetEditableBoardDescription(element), StringComparison.Ordinal) ||
                      (element.Type == ScreenplayElementType.SceneHeading && !string.Equals(newSceneHeading, GetEditableBoardSceneHeading(element), StringComparison.Ordinal));

        if (isDirty)
        {
            var updatedElement = ViewModel.SelectedDocument?.TryUpdateBoardElementContent(element, newHeading, newSceneHeading, newDescription);
            if (updatedElement is not null && ReferenceEquals(_currentInlineEditElement, element))
            {
                _currentInlineEditElement = updatedElement;
            }
            return updatedElement is not null;
        }

        return false;
    }

    private void RevealBoardElementInEditor(ScreenplayElement element)
    {
        if (element is ScratchpadCardElement draftCard &&
            !draftCard.SourceLineIndex.HasValue)
        {
            return;
        }

        if (element.StartLine < 1)
        {
            return;
        }

        CommitAndCloseInlineEditor();

        SetWorkspaceSurface(WorkspaceSurface.Editor);
        Dispatcher.BeginInvoke(
            new Action(() => HighlightBoardElementInEditor(element.StartLine, element.EndLine)),
            DispatcherPriority.ContextIdle);
    }

    private void HighlightBoardElementInEditor(int startLineNumber, int endLineNumber)
    {
        if (!IsLoaded || EditorBox is null || startLineNumber < 1)
        {
            return;
        }

        var startIndex = GetEditorCharacterIndexFromLineIndex(Math.Max(0, startLineNumber - 1));
        if (startIndex < 0)
        {
            return;
        }

        var nextLineStartIndex = GetEditorCharacterIndexFromLineIndex(Math.Max(0, endLineNumber));
        var endIndex = nextLineStartIndex >= 0
            ? nextLineStartIndex
            : GetEditorText().Length;

        SelectEditorRange(startIndex, Math.Max(0, endIndex - startIndex));
        ViewModel.SelectedOutlineLineNumber = startLineNumber;
        RestoreOutlineSelection(startLineNumber);
    }

    private ScreenplayElement? TryGetBeatBoardElement(DependencyObject? source)
    {
        var itemContainer = FindAncestor<ContentPresenter>(source);
        return itemContainer?.DataContext as ScreenplayElement;
    }

    private bool TryGetBeatBoardElementHost(ScreenplayElement element, out FrameworkElement elementHost)
    {
        if (BeatBoard?.ItemContainerGenerator.ContainerFromItem(element) is not ContentPresenter itemContainer ||
            VisualTreeHelper.GetChildrenCount(itemContainer) == 0 ||
            VisualTreeHelper.GetChild(itemContainer, 0) is not FrameworkElement templateRoot)
        {
            elementHost = null!;
            return false;
        }

        elementHost = templateRoot;
        return true;
    }

    private int GetBeatBoardBlockCount(ScreenplayElement element)
    {
        var boardElements = ViewModel.SelectedDocument?.BoardElements;
        if (boardElements is null || boardElements.Count == 0)
        {
            return 1;
        }

        try
        {
            return Math.Max(1, StoryHierarchyHelper.GetBlockRange(boardElements, element).Count);
        }
        catch (ArgumentException)
        {
            return 1;
        }
    }

    private bool TryGetBeatBoardDraggedElement(DragEventArgs e, out ScreenplayElement element)
    {
        if (e.Data.GetDataPresent(typeof(ScreenplayElement)) &&
            e.Data.GetData(typeof(ScreenplayElement)) is ScreenplayElement draggedElement)
        {
            element = draggedElement;
            return true;
        }

        element = null!;
        return false;
    }

    private bool TryGetBeatBoardDropLocation(
        Point boardPosition,
        ScreenplayElement draggedElement,
        out BeatBoardDropLocation dropLocation)
    {
        if (BeatBoard is null)
        {
            dropLocation = default;
            return false;
        }

        var itemLayouts = BeatBoard.Items
            .OfType<ScreenplayElement>()
            .Select(item => new
            {
                Element = item,
                Container = BeatBoard.ItemContainerGenerator.ContainerFromItem(item) as ContentPresenter
            })
            .Where(entry => entry.Container is not null && entry.Container.RenderSize.Width > 0.5 && entry.Container.RenderSize.Height > 0.5)
            .Select(entry => new BeatBoardItemLayout(
                entry.Element,
                entry.Container!.TransformToAncestor(BeatBoard)
                    .TransformBounds(new Rect(new Point(0.0, 0.0), entry.Container.RenderSize))))
            .OrderBy(entry => entry.Bounds.Top)
            .ThenBy(entry => entry.Bounds.Left)
            .ToArray();

        if (itemLayouts.Length == 0)
        {
            dropLocation = new BeatBoardDropLocation(
                null,
                InsertAfter: false,
                BeatBoardDropOperation.Insert,
                null);
            return true;
        }

        if (TryGetBeatBoardNestDropLocation(boardPosition, draggedElement, itemLayouts, out dropLocation))
        {
            return true;
        }

        if (_currentBeatBoardViewMode == BeatBoardViewMode.Outline)
        {
            return TryGetBeatBoardOutlineDropLocation(boardPosition, itemLayouts, out dropLocation);
        }

        return TryGetBeatBoardBoardDropLocation(boardPosition, itemLayouts, out dropLocation);
    }

    private bool TryGetBeatBoardNestDropLocation(
        Point boardPosition,
        ScreenplayElement draggedElement,
        IReadOnlyList<BeatBoardItemLayout> itemLayouts,
        out BeatBoardDropLocation dropLocation)
    {
        var directTarget = itemLayouts.FirstOrDefault(entry => entry.Bounds.Contains(boardPosition));
        if (directTarget.Element is null ||
            !CanNestBeatBoardElement(draggedElement, directTarget.Element))
        {
            dropLocation = default;
            return false;
        }

        var nestZone = CreateBeatBoardNestZone(directTarget.Bounds);
        if (!nestZone.Contains(boardPosition))
        {
            dropLocation = default;
            return false;
        }

        dropLocation = new BeatBoardDropLocation(
            directTarget.Element,
            InsertAfter: false,
            BeatBoardDropOperation.Nest,
            directTarget.Element);
        return true;
    }

    private bool TryGetBeatBoardOutlineDropLocation(
        Point boardPosition,
        IReadOnlyList<BeatBoardItemLayout> itemLayouts,
        out BeatBoardDropLocation dropLocation)
    {
        var firstItem = itemLayouts[0];
        if (boardPosition.Y < firstItem.Bounds.Top + (firstItem.Bounds.Height / 2.0))
        {
            dropLocation = CreateBeatBoardDropLocation(firstItem.Element, firstItem.Bounds, insertAfter: false);
            return true;
        }

        for (var index = 0; index < itemLayouts.Count; index++)
        {
            var current = itemLayouts[index];
            var currentMidpoint = current.Bounds.Top + (current.Bounds.Height / 2.0);
            if (boardPosition.Y < currentMidpoint)
            {
                dropLocation = CreateBeatBoardDropLocation(current.Element, current.Bounds, insertAfter: false);
                return true;
            }

            if (index == itemLayouts.Count - 1)
            {
                dropLocation = CreateBeatBoardDropLocation(current.Element, current.Bounds, insertAfter: true);
                return true;
            }

            var next = itemLayouts[index + 1];
            var boundaryMidpoint = current.Bounds.Bottom + ((next.Bounds.Top - current.Bounds.Bottom) / 2.0);
            if (boardPosition.Y < boundaryMidpoint)
            {
                dropLocation = CreateBeatBoardDropLocation(current.Element, current.Bounds, insertAfter: true);
                return true;
            }
        }

        var lastItem = itemLayouts[^1];
        dropLocation = CreateBeatBoardDropLocation(lastItem.Element, lastItem.Bounds, insertAfter: true);
        return true;
    }

    private bool TryGetBeatBoardBoardDropLocation(
        Point boardPosition,
        IReadOnlyList<BeatBoardItemLayout> itemLayouts,
        out BeatBoardDropLocation dropLocation)
    {
        // Find the closest item by distance to center
        var closestItem = itemLayouts[0];
        var minDistanceSq = double.MaxValue;

        foreach (var layout in itemLayouts)
        {
            var center = new Point(layout.Bounds.Left + (layout.Bounds.Width / 2.0), layout.Bounds.Top + (layout.Bounds.Height / 2.0));
            var dx = boardPosition.X - center.X;
            var dy = boardPosition.Y - center.Y;
            var distSq = (dx * dx) + (dy * dy);

            if (distSq < minDistanceSq)
            {
                minDistanceSq = distSq;
                closestItem = layout;
            }
        }

        var insertAfter = ShouldInsertAfterInBoardView(closestItem.Element, closestItem.Bounds, boardPosition);
        dropLocation = CreateBeatBoardDropLocation(closestItem.Element, closestItem.Bounds, insertAfter);
        return true;
    }

    private BeatBoardDropLocation CreateBeatBoardDropLocation(
        ScreenplayElement targetElement,
        Rect targetBounds,
        bool insertAfter)
    {
        return new BeatBoardDropLocation(
            targetElement,
            insertAfter,
            BeatBoardDropOperation.Insert,
            targetElement);
    }

    private static bool CanNestBeatBoardElement(ScreenplayElement draggedElement, ScreenplayElement targetElement)
    {
        return MainWindowViewModel.TryResolveBoardChildLevel(targetElement, draggedElement, out _);
    }

    private static bool UsesHorizontalBoardChronology(ScreenplayElement targetElement)
    {
        return targetElement.Level >= 2;
    }

    private static bool ShouldInsertAfterInBoardView(
        ScreenplayElement targetElement,
        Rect targetBounds,
        Point boardPosition)
    {
        if (UsesHorizontalBoardChronology(targetElement))
        {
            return boardPosition.X >= targetBounds.Left + (targetBounds.Width / 2.0);
        }

        return boardPosition.Y >= targetBounds.Top + (targetBounds.Height / 2.0);
    }

    private Rect CreateBeatBoardNestZone(Rect targetBounds)
    {
        if (_currentBeatBoardViewMode == BeatBoardViewMode.Outline)
        {
            const double outlineHorizontalInset = 22.0;
            const double outlineVerticalInset = 12.0;
            return new Rect(
                targetBounds.Left + outlineHorizontalInset,
                targetBounds.Top + outlineVerticalInset,
                Math.Max(0.0, targetBounds.Width - (outlineHorizontalInset * 2.0)),
                Math.Max(0.0, targetBounds.Height - (outlineVerticalInset * 2.0)));
        }

        var horizontalInset = Math.Min(36.0, targetBounds.Width * 0.22);
        var verticalInset = Math.Min(18.0, targetBounds.Height * 0.22);
        return new Rect(
            targetBounds.Left + horizontalInset,
            targetBounds.Top + verticalInset,
            Math.Max(0.0, targetBounds.Width - (horizontalInset * 2.0)),
            Math.Max(0.0, targetBounds.Height - (verticalInset * 2.0)));
    }


    private void ApplyBeatBoardDragPreviewConstraints()
    {
        if (_currentBeatBoardViewMode != BeatBoardViewMode.Outline)
        {
            return;
        }

        _beatBoardDragSourceSize = new Size(
            Math.Min(OutlineBeatBoardDragGhostMaxWidth, Math.Max(280.0, _beatBoardDragSourceSize.Width)),
            Math.Min(OutlineBeatBoardDragGhostMaxHeight, Math.Max(96.0, _beatBoardDragSourceSize.Height)));
        _beatBoardDragPointerOffset = new Point(
            Math.Min(_beatBoardDragPointerOffset.X, OutlineBeatBoardDragGhostMaxHotspotX),
            Math.Clamp(_beatBoardDragPointerOffset.Y, 0.0, Math.Max(0.0, _beatBoardDragSourceSize.Height - 1.0)));
    }

    private void StartBeatBoardDragGhost(ScreenplayElement element, int blockCount, Size sourceSize)
    {
        StopBeatBoardDragGhost();

        if (BeatBoardSurfaceRoot is null ||
            BeatBoardDragGhostHost is null)
        {
            return;
        }

        var accentBrush = ResolveBeatBoardAccentBrush(element);
        var ghostWidth = Math.Max(252.0, sourceSize.Width);
        if (BeatBoardSurfaceRoot.ActualWidth > 24.0)
        {
            ghostWidth = Math.Min(ghostWidth, BeatBoardSurfaceRoot.ActualWidth - 24.0);
        }

        var previewRoot = new Grid
        {
            Opacity = 0.56,
            IsHitTestVisible = false
        };

        if (blockCount > 1)
        {
            previewRoot.Children.Add(CreateBeatBoardGhostLayer(accentBrush, ghostWidth, new Thickness(12, 12, 0, 0), 0.08));
            previewRoot.Children.Add(CreateBeatBoardGhostLayer(accentBrush, ghostWidth, new Thickness(6, 6, 0, 0), 0.16));
        }

        previewRoot.Children.Add(CreateBeatBoardGhostCard(element, accentBrush, blockCount, ghostWidth, sourceSize.Height));
        previewRoot.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        var previewSize = previewRoot.DesiredSize;
        _beatBoardDragGhostHotspot = new Point(
            Math.Clamp(_beatBoardDragPointerOffset.X, 0.0, Math.Max(0.0, previewSize.Width - 1.0)),
            Math.Clamp(_beatBoardDragPointerOffset.Y, 0.0, Math.Max(0.0, previewSize.Height - 1.0)));

        _beatBoardDragGhostElement = previewRoot;
        BeatBoardDragGhostHost.Children.Add(previewRoot);
        UpdateBeatBoardDragGhostPosition();
    }

    private void StopBeatBoardDragGhost()
    {
        if (_beatBoardDragGhostElement is null)
        {
            return;
        }

        if (BeatBoardDragGhostHost is not null)
        {
            BeatBoardDragGhostHost.Children.Remove(_beatBoardDragGhostElement);
        }

        _beatBoardDragGhostElement = null;
    }

    private void UpdateBeatBoardDragGhostPosition()
    {
        if (_beatBoardDragGhostElement is null ||
            !TryGetCursorPositionRelativeToBeatBoardSurfaceRoot(out var cursorPoint))
        {
            return;
        }

        Canvas.SetLeft(_beatBoardDragGhostElement, cursorPoint.X - _beatBoardDragGhostHotspot.X);
        Canvas.SetTop(_beatBoardDragGhostElement, cursorPoint.Y - _beatBoardDragGhostHotspot.Y);
    }

    private bool TryGetCursorPositionRelativeToBeatBoardSurfaceRoot(out Point cursorPoint)
    {
        if (BeatBoardSurfaceRoot is null)
        {
            cursorPoint = default;
            return false;
        }

        if (GetCursorPos(out var nativePoint))
        {
            cursorPoint = BeatBoardSurfaceRoot.PointFromScreen(new Point(nativePoint.X, nativePoint.Y));
            return true;
        }

        cursorPoint = Mouse.GetPosition(BeatBoardSurfaceRoot);
        return true;
    }

    private Border CreateBeatBoardGhostLayer(Brush accentBrush, double width, Thickness margin, double opacity)
    {
        return new Border
        {
            Width = width,
            Margin = margin,
            Background = ResolveBrushResource("BeatBoardGhostLayerBackground", new SolidColorBrush(Color.FromRgb(26, 23, 21))),
            BorderBrush = accentBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(14),
            Opacity = opacity
        };
    }

    private Border CreateBeatBoardGhostCard(
        ScreenplayElement element,
        Brush accentBrush,
        int blockCount,
        double width,
        double minHeight)
    {
        var typeLabel = new TextBlock
        {
            Text = element.KindLabel.ToUpperInvariant(),
            Foreground = accentBrush,
            FontSize = 10.5,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 7)
        };

        var titleBlock = new TextBlock
        {
            Text = element.Heading,
            Foreground = ResolveBrushResource("BeatBoardGhostTitleForeground", new SolidColorBrush(Color.FromRgb(248, 241, 232))),
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap
        };

        var descriptionBlock = new TextBlock
        {
            Text = BuildBeatBoardGhostDescription(element, blockCount),
            Foreground = ResolveBrushResource("BeatBoardGhostDescriptionForeground", new SolidColorBrush(Color.FromRgb(213, 195, 177))),
            FontSize = 11.5,
            Margin = new Thickness(0, 10, 0, 0),
            TextWrapping = TextWrapping.Wrap,
            MaxWidth = Math.Max(160.0, width - 32.0)
        };

        var content = new StackPanel();
        content.Children.Add(typeLabel);
        content.Children.Add(titleBlock);
        content.Children.Add(descriptionBlock);

        return new Border
        {
            Width = width,
            MinHeight = minHeight,
            Background = ResolveBrushResource("BeatBoardGhostCardBackground", new SolidColorBrush(Color.FromRgb(36, 31, 30))),
            BorderBrush = accentBrush,
            BorderThickness = new Thickness(1.2),
            CornerRadius = new CornerRadius(14),
            Padding = new Thickness(16, 14, 16, 14),
            Child = content
        };
    }

    private static string BuildBeatBoardGhostDescription(ScreenplayElement element, int blockCount)
    {
        var description = element.BoardDescription;
        if (blockCount <= 1)
        {
            return description;
        }

        return string.IsNullOrWhiteSpace(description)
            ? $"{blockCount} linked cards"
            : $"{description}{Environment.NewLine}{blockCount} linked cards";
    }

    private Brush ResolveBeatBoardAccentBrush(ScreenplayElement element)
    {
        return element.Type switch
        {
            ScreenplayElementType.Note => new SolidColorBrush(Color.FromRgb(240, 169, 59)),
            ScreenplayElementType.Section when element.Level == 0 => new SolidColorBrush(Color.FromRgb(159, 119, 255)),
            ScreenplayElementType.Section when element.Level == 1 => new SolidColorBrush(Color.FromRgb(102, 173, 255)),
            _ => ResolveBrushResource("BeatBoardDefaultCardAccentForeground", new SolidColorBrush(Color.FromRgb(216, 199, 179)))
        };
    }

    private static T? FindAncestor<T>(DependencyObject? source)
        where T : DependencyObject
    {
        var current = source;
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = current switch
            {
                FrameworkContentElement contentElement => contentElement.Parent,
                _ => VisualTreeHelper.GetParent(current)
            };
        }

        return null;
    }

    private static T? FindDescendant<T>(DependencyObject? root)
        where T : DependencyObject
    {
        if (root is null)
        {
            return null;
        }

        var childCount = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < childCount; index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is T match)
            {
                return match;
            }

            var nestedMatch = FindDescendant<T>(child);
            if (nestedMatch is not null)
            {
                return nestedMatch;
            }
        }

        return null;
    }

    private static T? FindDescendant<T>(DependencyObject? root, string name)
        where T : FrameworkElement
    {
        if (root is null)
        {
            return null;
        }

        var childCount = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < childCount; index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is T match && match.Name == name)
            {
                return match;
            }

            var nestedMatch = FindDescendant<T>(child, name);
            if (nestedMatch is not null)
            {
                return nestedMatch;
            }
        }

        return null;
    }

    protected override void OnClosed(EventArgs e)
    {
        ThemeManager.ThemeChanged -= ThemeManager_ThemeChanged;
        AttachBeatBoardDocumentHandlers(null);
        _editorFormattingTimer.Stop();
        _editorDocumentSyncTimer.Stop();
        _beatBoardClickTimer.Stop();
        StopBeatBoardDragGhost();
        ClearBeatBoardDragTracking();
        ClearPendingBeatBoardClick();
        CommitAndCloseInlineEditor();
        CancelPendingEditorPreviewFormatting();
        CancelPendingEditorFormatting();
        StopEditorWheelScrollAnimation();
        if (_hwndSource is not null)
        {
            _hwndSource.RemoveHook(WindowMessageHook);
            _hwndSource = null;
        }
        base.OnClosed(e);
    }

    private IntPtr WindowMessageHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != WM_MOUSEHWHEEL)
        {
            return IntPtr.Zero;
        }

        var wheelDelta = unchecked((short)(((long)wParam >> 16) & 0xFFFF));
        if (wheelDelta == 0)
        {
            return IntPtr.Zero;
        }

        var wheelSteps = NormalizeWheelSteps(wheelDelta);

        if (EditorScrollHost is not null && EditorScrollHost.IsMouseOver)
        {
            handled = true;
            QueueEditorWheelScroll(
                horizontalDelta: wheelSteps * GetEditorWheelScrollStep() * EditorWheelScrollMultiplier * Math.Max(0.01, ViewModel.EditorZoomScale));
            return IntPtr.Zero;
        }

        if (BeatBoardScrollViewer is not null &&
            BeatBoardScrollViewer.IsMouseOver &&
            _currentWorkspaceSurface == WorkspaceSurface.BeatBoard &&
            _currentBeatBoardViewMode == BeatBoardViewMode.Board)
        {
            handled = true;
            ScrollBeatBoard(BeatBoardScrollViewer, horizontalDelta: wheelSteps * BeatBoardHorizontalWheelStep);
        }

        return IntPtr.Zero;
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Loaded -= MainWindow_Loaded;

        EnsureEditorDocumentInitialized();
        LoadEditorFromViewModel(resetSelection: true);
        UpdateUIForMode();
        UpdateWorkspaceSurface();
        SetLeftDockCollapsed(false);
        SetSyntaxQuickReferenceVisible(false);
        UpdateCaretContextFromEditor();

        EnsureEditorCueOverlay();

        RestoreEditorFocusAndCaret(ensureFocus: true);
        ScheduleEditorViewportRefresh(ensureCaretVisible: true);
    }

    // Keep the first launch centered on the primary screen and fully inside the visible work area.
    private void ApplyStartupBounds()
    {
        if (WindowState != WindowState.Normal)
        {
            WindowState = WindowState.Normal;
        }

        var workArea = SystemParameters.WorkArea;
        if (workArea.Width <= 0 || workArea.Height <= 0)
        {
            return;
        }

        var maxWidth = Math.Max(1.0, workArea.Width - WorkAreaPadding);
        var maxHeight = Math.Max(1.0, workArea.Height - WorkAreaPadding);

        if (MinWidth > maxWidth)
        {
            MinWidth = maxWidth;
        }

        if (MinHeight > maxHeight)
        {
            MinHeight = maxHeight;
        }

        Width = Math.Max(MinWidth, Math.Min(PreferredStartupWidth, maxWidth));
        Height = Math.Max(MinHeight, Math.Min(PreferredStartupHeight, maxHeight));

        Left = workArea.Left + Math.Max(0.0, (workArea.Width - Width) / 2.0);
        Top = workArea.Top + Math.Max(0.0, (workArea.Height - Height) / 2.0);
    }

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape)
        {
            if (_isWorkspaceDragging)
            {
                CancelWorkspaceDrag();
                e.Handled = true;
                return;
            }
        }

        if (!IsCtrlShortcut(e))
        {
            return;
        }

        switch (e.Key)
        {
            case System.Windows.Input.Key.N:
                New_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.F:
                ShowFindReplaceDialog();
                e.Handled = true;
                break;
            case System.Windows.Input.Key.G when !HasShiftModifier():
                ShowGoToLineDialog();
                e.Handled = true;
                break;
            case System.Windows.Input.Key.M:
                ToggleWriteMode();
                e.Handled = true;
                break;
            case System.Windows.Input.Key.O when HasShiftModifier():
                ShowGoToSceneDialog();
                e.Handled = true;
                break;
            case System.Windows.Input.Key.O:
                Open_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.Add:
            case System.Windows.Input.Key.OemPlus:
                ZoomIn_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.Subtract:
            case System.Windows.Input.Key.OemMinus:
                ZoomOut_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.D0:
            case System.Windows.Input.Key.NumPad0:
                ResetZoom_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.S when HasShiftModifier():
                SaveAs_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.S:
                Save_Click(sender, new RoutedEventArgs());
                e.Handled = true;
                break;
            case System.Windows.Input.Key.W:
                TryCloseDocument(ViewModel.SelectedDocument);
                e.Handled = true;
                break;
            case System.Windows.Input.Key.Z:
                if (PerformEditorUndoOrRedo(redo: false))
                {
                    e.Handled = true;
                }

                break;
            case System.Windows.Input.Key.Y:
                if (PerformEditorUndoOrRedo(redo: true))
                {
                    e.Handled = true;
                }

                break;
        }
    }

    public void InitializeDocument(string text, string? filePath, bool isDirty)
    {
        ViewModel.LoadDocument(text, filePath, isDirty);
    }

    private void New_Click(object sender, RoutedEventArgs e)
    {
        FlushPendingEditorChangeBatch();
        ViewModel.NewDocument();
    }

    private void Open_Click(object sender, RoutedEventArgs e)
    {
        FlushPendingEditorChangeBatch();

        var dialog = new OpenFileDialog
        {
            Filter = "Fountain files (*.fountain)|*.fountain|Text files (*.txt)|*.txt|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        foreach (var fileName in dialog.FileNames)
        {
            try
            {
                var text = File.ReadAllText(fileName);
                ViewModel.OpenDocument(text, fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    $"Could not open '{Path.GetFileName(fileName)}': {ex.Message}",
                    "Passage",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentDocument();
    }

    private void SaveAs_Click(object sender, RoutedEventArgs e)
    {
        SaveCurrentDocument(forcePromptForPath: true);
    }

    private void CloseCurrent_Click(object sender, RoutedEventArgs e)
    {
        FlushPendingEditorChangeBatch();
        TryCloseDocument(ViewModel.SelectedDocument);
    }

    private void CloseDocument_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not MainWindowViewModel document)
        {
            return;
        }

        FlushPendingEditorChangeBatch();
        TryCloseDocument(document);
    }

    private void DocumentTab_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBoxItem tabItem)
        {
            return;
        }

        if (FindAncestor<Button>(e.OriginalSource as DependencyObject) is not null)
        {
            return;
        }

        tabItem.IsSelected = true;
    }

    private void Find_Click(object sender, RoutedEventArgs e)
    {
        ShowFindReplaceDialog();
    }

    private void Undo_Click(object sender, RoutedEventArgs e)
    {
        PerformEditorUndoOrRedo(redo: false);
    }

    private void Redo_Click(object sender, RoutedEventArgs e)
    {
        PerformEditorUndoOrRedo(redo: true);
    }

    private void GoToLine_Click(object sender, RoutedEventArgs e)
    {
        ShowGoToLineDialog();
    }

    private void GoToScene_Click(object sender, RoutedEventArgs e)
    {
        ShowGoToSceneDialog();
    }

    private bool PerformEditorUndoOrRedo(bool redo)
    {
        if (EditorBox is null)
        {
            return false;
        }

        var initialText = GetEditorText();
        var performed = false;

        for (var attempt = 0; attempt < 12; attempt++)
        {
            var canApply = redo ? EditorBox.CanRedo : EditorBox.CanUndo;
            if (!canApply)
            {
                break;
            }

            if (redo)
            {
                EditorBox.Redo();
            }
            else
            {
                EditorBox.Undo();
            }

            performed = true;

            if (!string.Equals(GetEditorText(), initialText, StringComparison.Ordinal))
            {
                break;
            }
        }

        var currentText = GetEditorText();
        if (performed &&
            !string.Equals(currentText, initialText, StringComparison.Ordinal))
        {
            ViewModel.SynchronizeScratchpadWithUndoRedo(currentText);
        }

        return performed;
    }

    private void ZoomIn_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.IsBoardModeActive)
        {
            ViewModel.IncreaseBoardZoom();
        }
        else
        {
            CaptureEditorViewportAnchorForZoomChange();
            ViewModel.IncreaseEditorZoom();
            RefreshEditorViewportAfterZoomChange();
        }
    }

    private void ZoomOut_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.IsBoardModeActive)
        {
            ViewModel.DecreaseBoardZoom();
        }
        else
        {
            CaptureEditorViewportAnchorForZoomChange();
            ViewModel.DecreaseEditorZoom();
            RefreshEditorViewportAfterZoomChange();
        }
    }

    private void ResetZoom_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.IsBoardModeActive)
        {
            ViewModel.ResetBoardZoom();
        }
        else
        {
            CaptureEditorViewportAnchorForZoomChange();
            ViewModel.ResetEditorZoom();
            RefreshEditorViewportAfterZoomChange();
        }
    }

    private void RefreshEditorViewportAfterZoomChange()
    {
        ScheduleEditorViewportRefresh(ensureCaretVisible: false);
    }

    private void CaptureEditorViewportAnchorForZoomChange()
    {
        if (_hasPendingZoomViewportAnchor || EditorScrollHost is null)
        {
            return;
        }

        _pendingZoomViewportAnchorSourceScale = Math.Max(0.01, ViewModel.EditorZoomScale);
        _pendingZoomViewportAnchorSourceHorizontalOffset = EditorScrollHost.HorizontalOffset;
        _pendingZoomViewportAnchorSourceVerticalOffset = EditorScrollHost.VerticalOffset;
        _pendingZoomViewportAnchorSourceViewportWidth = Math.Max(0.0, EditorScrollHost.ViewportWidth);
        _pendingZoomViewportAnchorSourceViewportHeight = Math.Max(0.0, EditorScrollHost.ViewportHeight);

        var anchorPoint = new Point(
            _pendingZoomViewportAnchorSourceViewportWidth / 2.0,
            _pendingZoomViewportAnchorSourceViewportHeight / 2.0);

        if (IsScreenplayMode &&
            EditorBox is not null &&
            EditorBox.IsKeyboardFocusWithin &&
            _editorCueAdorner is not null &&
            _editorCueAdorner.TryGetCaretRect(out var caretRect) &&
            TryGetCaretRectInScrollHost(caretRect, out var caretBounds))
        {
            anchorPoint = new Point(
                caretBounds.Left + (caretBounds.Width / 2.0),
                caretBounds.Top + (caretBounds.Height / 2.0));
        }

        _pendingZoomViewportAnchorRatioX = _pendingZoomViewportAnchorSourceViewportWidth > 0.0
            ? anchorPoint.X / _pendingZoomViewportAnchorSourceViewportWidth
            : 0.5;
        _pendingZoomViewportAnchorRatioY = _pendingZoomViewportAnchorSourceViewportHeight > 0.0
            ? anchorPoint.Y / _pendingZoomViewportAnchorSourceViewportHeight
            : 0.5;
        _hasPendingZoomViewportAnchor = true;
    }

    private void RestoreEditorViewportAnchorAfterZoomChange()
    {
        if (!_hasPendingZoomViewportAnchor || EditorScrollHost is null)
        {
            return;
        }

        if (EditorScrollHost.ViewportWidth <= 0.0 || EditorScrollHost.ViewportHeight <= 0.0)
        {
            return;
        }

        var sourceScale = Math.Max(0.01, _pendingZoomViewportAnchorSourceScale);
        var targetScale = Math.Max(0.01, ViewModel.EditorZoomScale);

        var sourceAnchorX = _pendingZoomViewportAnchorSourceViewportWidth * _pendingZoomViewportAnchorRatioX;
        var sourceAnchorY = _pendingZoomViewportAnchorSourceViewportHeight * _pendingZoomViewportAnchorRatioY;
        var targetAnchorX = EditorScrollHost.ViewportWidth * _pendingZoomViewportAnchorRatioX;
        var targetAnchorY = EditorScrollHost.ViewportHeight * _pendingZoomViewportAnchorRatioY;

        var sourceContentX = (_pendingZoomViewportAnchorSourceHorizontalOffset + sourceAnchorX) / sourceScale;
        var sourceContentY = (_pendingZoomViewportAnchorSourceVerticalOffset + sourceAnchorY) / sourceScale;

        var targetHorizontalOffset = sourceContentX * targetScale - targetAnchorX;
        var targetVerticalOffset = sourceContentY * targetScale - targetAnchorY;

        _hasPendingZoomViewportAnchor = false;
        ScrollEditorViewportTo(targetHorizontalOffset, targetVerticalOffset);
    }

    private void Exit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void ToggleLeftDock_Click(object sender, RoutedEventArgs e)
    {
        SetLeftDockCollapsed(!_isLeftDockCollapsed);
    }

    private void ToggleSyntaxQuickReference_Click(object sender, RoutedEventArgs e)
    {
        SetSyntaxQuickReferenceVisible(!_isSyntaxQuickReferenceVisible);
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        FlushPendingEditorChangeBatch();
        ViewModel.SaveSessionNow();
        ViewModel.StopRecoveryAutosaveForAllDocuments();
        RecoveryStorage.ClearRecoveryFile();
        base.OnClosing(e);
    }

    private bool SaveCurrentDocument(bool forcePromptForPath = false)
    {
        return SaveDocument(ViewModel.SelectedDocument, forcePromptForPath);
    }

    private string GetDefaultExtension()
    {
        return currentMode switch
        {
            WriteMode.Screenplay => ".fountain",
            WriteMode.Markdown => ".md",
            _ => ".fountain"
        };
    }

    private string GetSaveFileDialogFilter()
    {
        return currentMode switch
        {
            WriteMode.Screenplay => "Fountain files (*.fountain)|*.fountain|Markdown files (*.md)|*.md|Text files (*.txt)|*.txt|All files (*.*)|*.*",
            WriteMode.Markdown => "Markdown files (*.md)|*.md|Fountain files (*.fountain)|*.fountain|Text files (*.txt)|*.txt|All files (*.*)|*.*",
            _ => "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        };
    }

    private string GetSuggestedFileName(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return $"Untitled{GetDefaultExtension()}";
        }

        return Path.ChangeExtension(Path.GetFileName(path), GetDefaultExtension());
    }

    private bool SaveDocument(MainWindowViewModel? document, bool forcePromptForPath = false)
    {
        if (document is null)
        {
            return false;
        }

        FlushPendingEditorChangeBatch();

        var path = document.DocumentPath;

        if (forcePromptForPath || string.IsNullOrWhiteSpace(path))
        {
            var dialog = new SaveFileDialog
            {
                DefaultExt = GetDefaultExtension(),
                AddExtension = true,
                Filter = GetSaveFileDialogFilter(),
                FilterIndex = 1,
                FileName = GetSuggestedFileName(path)
            };

            if (dialog.ShowDialog(this) != true)
            {
                return false;
            }

            path = dialog.FileName;
        }

        try
        {
            var titlePrefix = document.TitlePageText.Length > 0 ? document.TitlePageText + "\n\n" : "";
            File.WriteAllText(path!, titlePrefix + document.DocumentText);
            document.SetFilePath(path);
            document.MarkSaved();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"Could not save '{Path.GetFileName(path!)}': {ex.Message}",
                "Passage",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return false;
        }
    }

    private void ExportItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not IExporter exporter)
        {
            return;
        }

        ExportCurrentDocument(exporter);
    }

    private bool ExportCurrentDocument(IExporter exporter)
    {
        FlushPendingEditorChangeBatch();

        var dialog = new SaveFileDialog
        {
            Filter = $"{exporter.DisplayName} (*{exporter.DefaultExtension})|*{exporter.DefaultExtension}|All files (*.*)|*.*",
            FileName = string.IsNullOrWhiteSpace(ViewModel.DocumentPath)
                ? $"Untitled{exporter.DefaultExtension}"
                : Path.ChangeExtension(Path.GetFileName(ViewModel.DocumentPath), exporter.DefaultExtension)
        };

        if (dialog.ShowDialog(this) != true)
        {
            return false;
        }

        exporter.Export(ViewModel.CreateParsedSnapshot(), dialog.FileName);
        return true;
    }
    private async void PreviewModeToggleButton_Click(object sender, RoutedEventArgs e)
    {
        if (PreviewModeToggleButton.IsChecked == true)
        {
            if (ViewModel.SelectedDocument is null)
            {
                PreviewModeToggleButton.IsChecked = false;
                return;
            }

            var rawText = ViewModel.SelectedDocument.DocumentText;
            var titleText = ViewModel.SelectedDocument.TitlePageText;

            // Enter preview mode
            if (EditorScrollHost is not null) EditorScrollHost.Visibility = Visibility.Collapsed;
            if (BeatBoardWorkspaceHost is not null) BeatBoardWorkspaceHost.Visibility = Visibility.Collapsed;
            if (PreviewWorkspaceHost is not null) PreviewWorkspaceHost.Visibility = Visibility.Visible;

            var previewPages = await System.Threading.Tasks.Task.Run(() =>
            {
                var cleanBodyText = Passage.Parser.ScriptSanitizer.ExtractCleanScript(rawText);
                var fullText = titleText + (string.IsNullOrWhiteSpace(titleText) ? "" : "\n\n") + cleanBodyText;
                
                var cleanScreenplay = new Passage.Parser.FountainParser().Parse(fullText);
                var (titlePages, bodyPages) = Passage.Export.ScreenplayLayoutBuilder.BuildPages(cleanScreenplay);
                
                var pages = new List<PreviewPageViewModel>();
                
                double ComputeWpfX(string text, Passage.Export.LayoutTextStyle style, double rawX)
                {
                    var usableWidth = Passage.Export.ScreenplayLayoutBuilder.PageWidth - Passage.Export.ScreenplayLayoutBuilder.MarginLeft - Passage.Export.ScreenplayLayoutBuilder.MarginRight;
                    var lineWidth = Math.Min(usableWidth, Math.Max(1, text.Length) * Passage.Export.ScreenplayLayoutBuilder.CharWidth);

                    return style switch
                    {
                        Passage.Export.LayoutTextStyle.CenterWithinBody or Passage.Export.LayoutTextStyle.CenterWithinBodyBold => Passage.Export.ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, (usableWidth - lineWidth) / 2),
                        Passage.Export.LayoutTextStyle.RightWithinBody => Passage.Export.ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, usableWidth - lineWidth),
                        _ => rawX
                    };
                }
                
                void AddLayoutPage(Passage.Export.LayoutPage lp)
                {
                    var pageVm = new PreviewPageViewModel();
                    
                    if (!lp.IsTitlePage && lp.PageNumber > 1)
                    {
                        pageVm.PageNumberLabel = $"{lp.PageNumber}.";
                    }
                    
                    var y = Passage.Export.ScreenplayLayoutBuilder.PageHeight - Passage.Export.ScreenplayLayoutBuilder.MarginTop;

                    foreach (var line in lp.Lines)
                    {
                        if (line.IsBlank)
                        {
                            y -= Passage.Export.ScreenplayLayoutBuilder.LineHeight;
                            continue;
                        }

                        var wpfTop = Passage.Export.ScreenplayLayoutBuilder.PageHeight - y - Passage.Export.ScreenplayLayoutBuilder.FontSize;

                        pageVm.Lines.Add(new PreviewLineViewModel
                        {
                            Text = line.Text,
                            X = ComputeWpfX(line.Text, line.Style, line.X),
                            Y = wpfTop,
                            IsBold = line.Style == Passage.Export.LayoutTextStyle.LeftBold || line.Style == Passage.Export.LayoutTextStyle.CenterWithinBodyBold
                        });

                        y -= Passage.Export.ScreenplayLayoutBuilder.LineHeight;
                    }
                    
                    pages.Add(pageVm);
                }

                foreach (var tp in titlePages) AddLayoutPage(tp);
                foreach (var bp in bodyPages) AddLayoutPage(bp);

                return pages;
            });

            PreviewPagesControl.ItemsSource = previewPages;
        }
        else
        {
            // Exit preview mode
            if (PreviewWorkspaceHost is not null) PreviewWorkspaceHost.Visibility = Visibility.Collapsed;
            PreviewPagesControl.ItemsSource = null;
            UpdateWorkspaceSurface(); // This automatically restores the correct surface (Editor or BeatBoard) and manages viewport/focus
        }
    }

    private void PreviewWorkspaceHost_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (PreviewWorkspaceHost is null || e.Delta == 0) return;

        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            // Zoom could be added here later if needed
            return;
        }

        e.Handled = true;

        // Use the same multiplier logic as the editor for consistent trackpad sensitivity
        double step = 16.0;
        if (EditorBox is not null)
        {
            var lineSpacing = 1.2;
            try { lineSpacing = EditorBox.FontFamily.LineSpacing; } catch { }
            step = Math.Max(12.0, EditorBox.FontSize * lineSpacing);
        }

        double delta = -e.Delta / 120.0 * step * EditorWheelScrollMultiplier;
        PreviewWorkspaceHost.ScrollToVerticalOffset(Math.Clamp(PreviewWorkspaceHost.VerticalOffset + delta, 0, PreviewWorkspaceHost.ScrollableHeight));
    }

    private void TitlePage_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedDocument is null) return;

        var currentTitlePage = ViewModel.SelectedDocument.TitlePage;
        var dialogViewModel = new Passage.App.ViewModels.TitlePageViewModel();
        dialogViewModel.CopyFrom(currentTitlePage);

        var dialog = new Passage.App.Views.TitlePageDialog
        {
            DataContext = dialogViewModel,
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            if (dialog.Deleted)
            {
                currentTitlePage.Title = "";
                currentTitlePage.Episode = "";
                currentTitlePage.Author = "";
                currentTitlePage.Credit = "written by";
                currentTitlePage.Source = "";
                currentTitlePage.Contact = "";
                currentTitlePage.DraftDate = "";
                currentTitlePage.Revision = "";
                currentTitlePage.Notes = "";
                currentTitlePage.CustomEntries.Clear();
                currentTitlePage.ShowTitlePage = false;
            }
            else
            {
                currentTitlePage.CopyFrom(dialogViewModel);
            }
        }
    }
    private bool TryCloseDocument(MainWindowViewModel? document)
    {
        if (document is null)
        {
            return false;
        }

        FlushPendingEditorChangeBatch();

        if (document.IsDirty)
        {
            var documentName = string.IsNullOrWhiteSpace(document.DocumentPath)
                ? "Untitled document"
                : Path.GetFileName(document.DocumentPath);

            var result = MessageBox.Show(
                this,
                $"Save changes to '{documentName}' before closing?",
                "Passage",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question,
                MessageBoxResult.Yes);

            switch (result)
            {
                case MessageBoxResult.Yes:
                    if (!SaveDocument(document))
                    {
                        return false;
                    }

                    break;
                case MessageBoxResult.No:
                    document.StopRecoveryAutosave();
                    RecoveryStorage.ClearRecoveryFile();
                    break;
                default:
                    return false;
            }
        }

        return ViewModel.CloseDocument(document);
    }

    private string GetEditorText()
    {
        return EditorBox is null
            ? string.Empty
            : RichTextBoxTextUtilities.GetPlainText(EditorBox);
    }

    private int GetEditorCaretIndex()
    {
        return EditorBox is null
            ? 0
            : RichTextBoxTextUtilities.GetCaretIndex(EditorBox);
    }

    private int GetEditorSelectionStart()
    {
        return EditorBox is null
            ? 0
            : RichTextBoxTextUtilities.GetSelectionStart(EditorBox);
    }

    private int GetEditorSelectionLength()
    {
        return EditorBox is null
            ? 0
            : RichTextBoxTextUtilities.GetSelectionLength(EditorBox);
    }

    private int GetEditorSelectionAnchorIndex()
    {
        return EditorBox is null
            ? 0
            : RichTextBoxTextUtilities.GetSelectionAnchorIndex(EditorBox);
    }

    private int GetEditorLineIndexFromCharacterIndex(int characterIndex)
    {
        return EditorBox is null
            ? 0
            : RichTextBoxTextUtilities.GetLineIndexFromCharacterIndex(EditorBox, characterIndex);
    }

    private int GetEditorCharacterIndexFromLineIndex(int lineIndex)
    {
        return EditorBox is null
            ? -1
            : RichTextBoxTextUtilities.GetCharacterIndexFromLineIndex(EditorBox, lineIndex);
    }

    private string GetEditorLineText(int lineIndex)
    {
        return EditorBox is null
            ? string.Empty
            : RichTextBoxTextUtilities.GetLineText(EditorBox, lineIndex);
    }

    private TextPointer GetEditorTextPointerAtOffset(int offset)
    {
        return EditorBox is null
            ? throw new InvalidOperationException("EditorBox is not initialized.")
            : RichTextBoxTextUtilities.GetTextPointerAtOffset(EditorBox, offset);
    }

    private void SetEditorSelection(int anchorIndex, int activeIndex)
    {
        if (EditorBox is null)
        {
            return;
        }

        RichTextBoxTextUtilities.SetSelection(EditorBox, anchorIndex, activeIndex);
    }

    private void EnsureEditorDocumentInitialized()
    {
        if (EditorBox is null)
        {
            return;
        }

        if (EditorBox.Document is null)
        {
            EditorBox.Document = CreateInitializedEditorDocument();
        }

        if (EditorBox.Document.Blocks.Count == 0)
        {
            EditorBox.Document.Blocks.Add(CreateEditorParagraph(string.Empty));
            return;
        }

        var firstParagraph = EditorBox.Document.Blocks.OfType<Paragraph>().FirstOrDefault();
        if (firstParagraph is null)
        {
            var documentText = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd).Text;
            if (string.IsNullOrWhiteSpace(documentText))
            {
                EditorBox.Document.Blocks.Clear();
                EditorBox.Document.Blocks.Add(CreateEditorParagraph(string.Empty));
            }

            return;
        }

        if (!firstParagraph.Inlines.OfType<Run>().Any())
        {
            firstParagraph.Inlines.Clear();
            firstParagraph.Inlines.Add(new Run(string.Empty));
        }
    }

    private static FlowDocument CreateInitializedEditorDocument()
    {
        var document = new FlowDocument();
        document.PagePadding = ScreenplayDocumentPagePadding;
        document.ColumnWidth = double.PositiveInfinity;
        document.Blocks.Add(CreateEditorParagraph(string.Empty));
        return document;
    }

    private static Paragraph CreateEditorParagraph(string text = "")
    {
        return new Paragraph(new Run(text))
        {
            Margin = new Thickness(0),
            TextAlignment = TextAlignment.Left
        };
    }

    private void SetEditorDocumentText(string text)
    {
        if (EditorBox is null)
        {
            return;
        }

        EnsureEditorDocumentInitialized();
        RichTextBoxTextUtilities.SetPlainText(EditorBox, text ?? string.Empty);
    }

    private void RestoreEditorFocusAndCaret(bool ensureFocus)
    {
        if (EditorBox is null)
        {
            return;
        }

        EnsureEditorDocumentInitialized();
        ApplyEditorCaretBrush();

        if (ensureFocus && !EditorBox.IsKeyboardFocusWithin)
        {
            EditorBox.Focus();
            Keyboard.Focus(EditorBox);
        }
    }

    private void ApplyEditorCaretBrush()
    {
        if (EditorBox is null)
        {
            return;
        }

        if (IsEditorCueOverlayVisible())
        {
            EditorBox.CaretBrush = Brushes.Transparent;
            return;
        }

        if (TryFindResource("EditorCaret") is Brush caretBrush)
        {
            EditorBox.CaretBrush = caretBrush;
            return;
        }

        EditorBox.CaretBrush = string.Equals(ThemeManager.AppliedThemeName, ThemeManager.DarkThemeName, StringComparison.Ordinal)
            ? Brushes.White
            : Brushes.Black;
    }

    private void LoadEditorFromViewModel(bool resetSelection)
    {
        if (EditorBox is null)
        {
            return;
        }

        CancelPendingEditorFormatting();
        EnsureEditorDocumentInitialized();

        var documentText = ViewModel.DocumentText ?? string.Empty;
        var hasEditableParagraphs = EditorBox.Document.Blocks.OfType<Paragraph>().Any();
        var editorText = GetEditorText();

        if (hasEditableParagraphs && (string.Equals(editorText, documentText, StringComparison.Ordinal) ||
            string.Equals(editorText.Replace("\r\n", "\n"), documentText.Replace("\r\n", "\n"), StringComparison.Ordinal)))
        {
            ResetEditorChangeTracking(documentText);
            QueueEditorFormattingForDocument(immediate: true);
            return;
        }

        var anchorIndex = resetSelection ? 0 : GetEditorSelectionAnchorIndex();
        var caretIndex = resetSelection ? 0 : GetEditorCaretIndex();

        _isSynchronizingEditorDocument = true;
        try
        {
            EditorBox.BeginChange();
            SetEditorDocumentText(documentText);
        }
        finally
        {
            EditorBox.EndChange();
            _isSynchronizingEditorDocument = false;
        }

        ResetEditorChangeTracking(documentText);
        RestoreEditorSelection(anchorIndex, caretIndex);
        QueueEditorFormattingForDocument(immediate: true);
    }

    private void RestoreEditorSelection(int anchorIndex, int caretIndex)
    {
        if (EditorBox is null)
        {
            return;
        }

        try
        {
            _suppressCaretContextUpdates = true;
            SetEditorSelection(anchorIndex, caretIndex);
        }
        finally
        {
            _suppressCaretContextUpdates = false;
        }
    }

    private void CancelPendingEditorFormatting()
    {
        var pendingFormatting = Interlocked.Exchange(ref _editorFormattingCancellation, null);
        if (pendingFormatting is null)
        {
            return;
        }

        pendingFormatting.Cancel();
        pendingFormatting.Dispose();
    }

    private void CancelPendingEditorPreviewFormatting()
    {
        var pendingFormatting = Interlocked.Exchange(ref _editorPreviewFormattingCancellation, null);
        if (pendingFormatting is null)
        {
            return;
        }

        pendingFormatting.Cancel();
        pendingFormatting.Dispose();
    }

    private void ResetEditorChangeTracking(string editorText)
    {
        _editorFormattingTimer.Stop();
        _editorDocumentSyncTimer.Stop();
        CancelPendingEditorPreviewFormatting();
        _pendingPreviewFormattingLineNumbers.Clear();
        _pendingForcedFormattingLineNumbers.Clear();
        _editorTextSyncPending = false;
        _formatEntireDocumentRequested = false;
        _lastCommittedEditorText = editorText ?? string.Empty;
        _lastTextChangedParagraph = null;
        _lastTextChangedParagraphText = string.Empty;
    }

    private void RestartEditorFormattingTimer(bool immediate = false)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        if (immediate)
        {
            _editorFormattingTimer.Stop();
            ProcessPendingPreviewFormatting();
            return;
        }

        if (!_editorFormattingTimer.IsEnabled)
        {
            _editorFormattingTimer.Start();
        }
    }

    private void RestartEditorDocumentSyncTimer(bool immediate = false)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        _editorDocumentSyncTimer.Stop();
        if (immediate)
        {
            ProcessPendingEditorChangeBatch();
            return;
        }

        _editorDocumentSyncTimer.Start();
    }

    private void EditorFormattingTimer_Tick(object? sender, EventArgs e)
    {
        _editorFormattingTimer.Stop();
        ProcessPendingPreviewFormatting();
    }

    private void EditorDocumentSyncTimer_Tick(object? sender, EventArgs e)
    {
        _editorDocumentSyncTimer.Stop();
        ProcessPendingEditorChangeBatch();
    }

    private void FlushPendingEditorChangeBatch()
    {
        if (!_editorTextSyncPending &&
            !_formatEntireDocumentRequested &&
            _pendingForcedFormattingLineNumbers.Count == 0)
        {
            return;
        }

        _editorDocumentSyncTimer.Stop();
        ProcessPendingEditorChangeBatch();
    }

    private void QueuePreviewFormattingForActiveParagraph(bool immediate = false)
    {
        if (GetActiveParagraph() is not Paragraph activeParagraph)
        {
            return;
        }

        QueuePreviewFormattingForParagraph(activeParagraph, immediate);
    }

    private void QueuePreviewFormattingForParagraph(Paragraph paragraph, bool immediate = false)
    {
        QueuePreviewFormattingForLines(immediate, GetParagraphLineNumber(paragraph));
    }

    private void QueuePreviewFormattingForLines(bool immediate = false, params int[] lineNumbers)
    {
        foreach (var lineNumber in lineNumbers.Where(lineNumber => lineNumber > 0))
        {
            _pendingPreviewFormattingLineNumbers.Add(lineNumber);
        }

        if (_pendingPreviewFormattingLineNumbers.Count == 0)
        {
            return;
        }

        RestartEditorFormattingTimer(immediate);
    }

    private void ProcessPendingPreviewFormatting()
    {
        if (EditorBox is null || _isSynchronizingEditorDocument || _pendingPreviewFormattingLineNumbers.Count == 0)
        {
            return;
        }

        var targetLineNumbers = _pendingPreviewFormattingLineNumbers
            .OrderBy(lineNumber => lineNumber)
            .ToArray();
        _pendingPreviewFormattingLineNumbers.Clear();

        var paragraphSnapshots = new List<ParagraphFormattingSnapshot>(targetLineNumbers.Length);
        foreach (var lineNumber in targetLineNumbers)
        {
            if (!TryGetParagraphByLineNumber(lineNumber, out var paragraph) ||
                !TryCreateParagraphFormattingSnapshot(paragraph, out var paragraphSnapshot))
            {
                continue;
            }

            paragraphSnapshots.Add(paragraphSnapshot);
        }

        if (paragraphSnapshots.Count == 0)
        {
            return;
        }

        StartEditorPreviewFormatting(paragraphSnapshots);
    }

    private void StartEditorPreviewFormatting(IReadOnlyList<ParagraphFormattingSnapshot> paragraphSnapshots)
    {
        CancelPendingEditorPreviewFormatting();

        var requestVersion = Interlocked.Increment(ref _editorPreviewFormattingVersion);
        var requestMode = currentMode;
        var cancellation = new CancellationTokenSource();
        _editorPreviewFormattingCancellation = cancellation;

        _ = RunEditorPreviewFormattingAsync(
            requestVersion,
            requestMode,
            paragraphSnapshots,
            cancellation.Token);
    }

    private async Task RunEditorPreviewFormattingAsync(
        int requestVersion,
        WriteMode requestMode,
        IReadOnlyList<ParagraphFormattingSnapshot> paragraphSnapshots,
        CancellationToken cancellationToken)
    {
        try
        {
            var formattingResults = await Task.Run(
                () => BuildPreviewFormattingResults(requestMode, paragraphSnapshots, cancellationToken),
                cancellationToken).ConfigureAwait(false);

            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            Dispatcher.Invoke(() =>
            {
                if (cancellationToken.IsCancellationRequested ||
                    requestVersion != _editorPreviewFormattingVersion ||
                    requestMode != currentMode)
                {
                    return;
                }

                ApplyPreviewFormattingResults(formattingResults);
            });
        }
        catch (OperationCanceledException)
        {
        }
    }

    private static IReadOnlyList<ParagraphFormattingResult> BuildPreviewFormattingResults(
        WriteMode requestMode,
        IReadOnlyList<ParagraphFormattingSnapshot> paragraphSnapshots,
        CancellationToken cancellationToken)
    {
        var paragraphResults = new List<ParagraphFormattingResult>(paragraphSnapshots.Count);
        foreach (var paragraphSnapshot in paragraphSnapshots)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var screenplayType = requestMode == WriteMode.Screenplay
                ? ResolveScreenplayParagraphType(paragraphSnapshot)
                : ScreenplayElementType.Action;
            var boldSpans = requestMode == WriteMode.Markdown && paragraphSnapshot.Text.Length > 0
                ? GetMarkdownBoldSpans(paragraphSnapshot.Text)
                : Array.Empty<FormattingSpan>();

            paragraphResults.Add(new ParagraphFormattingResult(
                paragraphSnapshot.LineNumber,
                paragraphSnapshot.Text,
                screenplayType,
                boldSpans));
        }

        return paragraphResults;
    }

    private void ApplyPreviewFormattingResults(IReadOnlyList<ParagraphFormattingResult> paragraphResults)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument || paragraphResults.Count == 0)
        {
            return;
        }

        var validatedResults = new List<ParagraphFormattingResult>(paragraphResults.Count);
        foreach (var paragraphResult in paragraphResults)
        {
            if (!TryGetParagraphByLineNumber(paragraphResult.LineNumber, out var paragraph))
            {
                continue;
            }

            if (!string.Equals(GetParagraphText(paragraph), paragraphResult.Text, StringComparison.Ordinal))
            {
                continue;
            }

            validatedResults.Add(paragraphResult);
        }

        if (validatedResults.Count == 0)
        {
            return;
        }

        ApplyFormattingResults(validatedResults);
    }

    private void ProcessPendingEditorChangeBatch()
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        var snapshotText = GetEditorText();
        var textChanged = !string.Equals(_lastCommittedEditorText, snapshotText, StringComparison.Ordinal);
        var formatEntireDocument = _formatEntireDocumentRequested;
        var targetLineNumbers = _pendingForcedFormattingLineNumbers
            .Where(lineNumber => lineNumber > 0)
            .Distinct()
            .OrderBy(lineNumber => lineNumber)
            .ToArray();

        _lastCommittedEditorText = snapshotText;
        _editorTextSyncPending = false;
        _formatEntireDocumentRequested = false;
        _pendingForcedFormattingLineNumbers.Clear();

        if (textChanged)
        {
            PushEditorTextToViewModel(snapshotText);
        }

        if (!formatEntireDocument && targetLineNumbers.Length == 0)
        {
            return;
        }

        var snapshotLines = SplitEditorLines(snapshotText);
        if (formatEntireDocument)
        {
            targetLineNumbers = Enumerable.Range(1, snapshotLines.Length).ToArray();
        }

        StartEditorFormatting(snapshotText, snapshotLines, targetLineNumbers);
    }

    private void PushEditorTextToViewModel(string editorText)
    {
        if (DataContext is not ShellViewModel viewModel ||
            string.Equals(viewModel.DocumentText, editorText, StringComparison.Ordinal))
        {
            return;
        }

        try
        {
            _suppressEditorDocumentReload = true;
            viewModel.DocumentText = editorText;
        }
        finally
        {
            _suppressEditorDocumentReload = false;
        }
    }

    private void QueueEditorFormattingForDocument(bool immediate = false)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        _formatEntireDocumentRequested = true;
        RestartEditorDocumentSyncTimer(immediate);
    }

    private void QueueEditorFormattingForActiveParagraph(bool force = false, bool immediate = false)
    {
        if (GetActiveParagraph() is not Paragraph activeParagraph)
        {
            return;
        }

        QueueEditorFormattingForParagraph(activeParagraph, force, immediate);
    }

    private void QueueEditorFormattingForParagraph(Paragraph paragraph, bool force = false, bool immediate = false)
    {
        var lineNumber = GetParagraphLineNumber(paragraph);
        QueuePreviewFormattingForParagraph(paragraph, immediate);

        if (!force)
        {
            return;
        }

        _pendingForcedFormattingLineNumbers.Add(lineNumber);
        RestartEditorDocumentSyncTimer(immediate: false);
    }

    private void QueueEditorFormattingForLines(bool immediate = false, params int[] lineNumbers)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        var validLineNumbers = lineNumbers
            .Where(lineNumber => lineNumber > 0)
            .Distinct()
            .ToArray();
        if (validLineNumbers.Length == 0)
        {
            return;
        }

        QueuePreviewFormattingForLines(immediate, validLineNumbers);

        foreach (var lineNumber in validLineNumbers)
        {
            _pendingForcedFormattingLineNumbers.Add(lineNumber);
        }

        RestartEditorDocumentSyncTimer(immediate);
    }

    private void StartEditorFormatting(string snapshotText, string[] snapshotLines, int[] targetLineNumbers)
    {
        if (EditorBox is null || snapshotLines.Length == 0)
        {
            return;
        }

        var validTargetLineNumbers = targetLineNumbers
            .Where(lineNumber => lineNumber >= 1 && lineNumber <= snapshotLines.Length)
            .ToArray();
        if (validTargetLineNumbers.Length == 0)
        {
            return;
        }

        var requestMode = currentMode;
        var lineTypeOverrides = requestMode == WriteMode.Screenplay
            ? CaptureScreenplayLineTypeOverrides(snapshotLines.Length)
            : new Dictionary<int, ScreenplayElementType>();

        CancelPendingEditorFormatting();

        var requestVersion = Interlocked.Increment(ref _editorFormattingVersion);
        var cancellation = new CancellationTokenSource();
        _editorFormattingCancellation = cancellation;

        _ = RunEditorFormattingAsync(
            requestVersion,
            requestMode,
            snapshotText,
            snapshotLines,
            lineTypeOverrides,
            validTargetLineNumbers,
            cancellation.Token);
    }

    private IReadOnlyDictionary<int, ScreenplayElementType> CaptureScreenplayLineTypeOverrides(int lineCount)
    {
        var lineTypeOverrides = new Dictionary<int, ScreenplayElementType>();
        for (var lineNumber = 1; lineNumber <= lineCount; lineNumber++)
        {
            if (!ViewModel.HasLineTypeOverride(lineNumber))
            {
                continue;
            }

            lineTypeOverrides[lineNumber] = ViewModel.GetEffectiveLineType(lineNumber);
        }

        return lineTypeOverrides;
    }

    private async Task RunEditorFormattingAsync(
        int requestVersion,
        WriteMode requestMode,
        string snapshotText,
        string[] snapshotLines,
        IReadOnlyDictionary<int, ScreenplayElementType> lineTypeOverrides,
        int[] targetLineNumbers,
        CancellationToken cancellationToken)
    {
        try
        {
            var formattingResult = await Task.Run(
                () => BuildFormattingComputationResult(
                    requestMode,
                    snapshotText,
                    snapshotLines,
                    lineTypeOverrides,
                    targetLineNumbers,
                    cancellationToken),
                cancellationToken).ConfigureAwait(false);

            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            Dispatcher.Invoke(() =>
            {
                if (cancellationToken.IsCancellationRequested ||
                    requestVersion != _editorFormattingVersion)
                {
                    return;
                }

                ApplyFormattingResult(formattingResult);
            });
        }
        catch (OperationCanceledException)
        {
        }
    }

    private static FormattingComputationResult BuildFormattingComputationResult(
        WriteMode requestMode,
        string snapshotText,
        IReadOnlyList<string> snapshotLines,
        IReadOnlyDictionary<int, ScreenplayElementType> lineTypeOverrides,
        IReadOnlyList<int> targetLineNumbers,
        CancellationToken cancellationToken)
    {
        IReadOnlyDictionary<int, ScreenplayElementType> screenplayLineTypes = requestMode == WriteMode.Screenplay
            ? CreateScreenplayLineTypeLookup(new FountainParser().Parse(snapshotText, lineTypeOverrides))
            : new Dictionary<int, ScreenplayElementType>();

        var paragraphResults = new List<ParagraphFormattingResult>(targetLineNumbers.Count);
        foreach (var lineNumber in targetLineNumbers)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var lineText = lineNumber <= snapshotLines.Count
                ? snapshotLines[lineNumber - 1]
                : string.Empty;

            var screenplayType = requestMode == WriteMode.Screenplay
                ? ResolveScreenplayLineType(lineText, lineNumber, screenplayLineTypes)
                : ScreenplayElementType.Action;
            var boldSpans = requestMode == WriteMode.Markdown && lineText.Length > 0
                ? GetMarkdownBoldSpans(lineText)
                : Array.Empty<FormattingSpan>();

            paragraphResults.Add(new ParagraphFormattingResult(lineNumber, lineText, screenplayType, boldSpans));
        }

        return new FormattingComputationResult(snapshotText, paragraphResults);
    }

    private void ApplyFormattingResult(FormattingComputationResult formattingResult)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        if (!string.Equals(GetEditorText(), formattingResult.SnapshotText, StringComparison.Ordinal))
        {
            return;
        }

        ApplyFormattingResults(formattingResult.Paragraphs);
    }

    private void ApplyFormattingResults(IReadOnlyList<ParagraphFormattingResult> paragraphResults)
    {
        if (EditorBox is null || _isSynchronizingEditorDocument)
        {
            return;
        }

        var paragraphLookup = CreateParagraphLookup(paragraphResults.Select(result => result.LineNumber));

        _isFormatting = true;
        try
        {
            EditorBox.BeginChange();
            using var changeBlock = EditorBox.DeclareChangeBlock();

            foreach (var paragraphFormatting in paragraphResults)
            {
                if (!paragraphLookup.TryGetValue(paragraphFormatting.LineNumber, out var paragraph))
                {
                    continue;
                }

                ApplyParagraphFormatting(paragraph, paragraphFormatting);
            }
        }
        finally
        {
            EditorBox.EndChange();
            _isFormatting = false;
        }

        ScheduleEditorViewportRefresh(ensureCaretVisible: EditorBox.IsKeyboardFocusWithin);
    }

    private void ApplyParagraphFormatting(Paragraph paragraph, ParagraphFormattingResult paragraphFormatting)
    {
        ApplyParagraphLayout(paragraph, paragraphFormatting.ScreenplayType);

        var text = paragraphFormatting.Text ?? string.Empty;
        var match = IdCommentRegex.Match(text);
        
        if (match.Success)
        {
            var before = text.Substring(0, match.Index);
            var idTag = match.Value;
            var after = text.Substring(match.Index + match.Length);

            // Rebuild inlines to ensure the ID tag is in an atomic UI container
            paragraph.Inlines.Clear();
            if (!string.IsNullOrEmpty(before))
            {
                paragraph.Inlines.Add(new Run(before));
            }

            var idBlock = new TextBlock 
            { 
                Text = idTag, 
                FontSize = 0.1, 
                Opacity = 0, 
                Width = 0, 
                Height = 0,
                IsHitTestVisible = false
            };
            paragraph.Inlines.Add(new InlineUIContainer(idBlock));

            if (!string.IsNullOrEmpty(after))
            {
                paragraph.Inlines.Add(new Run(after));
            }
        }
    }

    private Dictionary<int, Paragraph> CreateParagraphLookup(IEnumerable<int> lineNumbers)
    {
        var paragraphLookup = new Dictionary<int, Paragraph>();
        if (EditorBox is null)
        {
            return paragraphLookup;
        }

        var requestedLineNumbers = new HashSet<int>(lineNumbers.Where(lineNumber => lineNumber > 0));
        if (requestedLineNumbers.Count == 0)
        {
            return paragraphLookup;
        }

        var currentLineNumber = 1;
        foreach (var block in EditorBox.Document.Blocks)
        {
            if (block is not Paragraph paragraph)
            {
                continue;
            }

            if (requestedLineNumbers.Contains(currentLineNumber))
            {
                paragraphLookup[currentLineNumber] = paragraph;
                if (paragraphLookup.Count == requestedLineNumbers.Count)
                {
                    break;
                }
            }

            currentLineNumber++;
        }

        return paragraphLookup;
    }

    private static void ApplyParagraphLayout(Paragraph paragraph, ScreenplayElementType screenplayType)
    {
        var margin = ActionParagraphMargin;
        var alignment = TextAlignment.Left;

        switch (screenplayType)
        {
            case ScreenplayElementType.SceneHeading:
            case ScreenplayElementType.Section:
                margin = new Thickness(0, 24, 0, 12);
                alignment = TextAlignment.Left;
                break;
            case ScreenplayElementType.Synopsis:
                margin = new Thickness(0, 12, 0, 12);
                alignment = TextAlignment.Left;
                break;
            case ScreenplayElementType.Character:
            case ScreenplayElementType.CenteredText:
            case ScreenplayElementType.TitlePageCentered:
                alignment = TextAlignment.Center;
                break;
            case ScreenplayElementType.Dialogue:
                margin = DialogueParagraphMargin;
                alignment = TextAlignment.Left;
                break;
            case ScreenplayElementType.Parenthetical:
                margin = ParentheticalParagraphMargin;
                alignment = TextAlignment.Left;
                break;
            case ScreenplayElementType.Transition:
                alignment = TextAlignment.Right;
                break;
            case ScreenplayElementType.TitlePageContact:
                margin = new Thickness(0, 48, 0, 240); // Creates visual page-break spacing in the editor
                break;
        }

        if (!ThicknessEquals(paragraph.Margin, margin))
        {
            paragraph.Margin = margin;
        }

        if (Math.Abs(paragraph.TextIndent) > 0.0)
        {
            paragraph.TextIndent = 0.0;
        }

        if (paragraph.TextAlignment != alignment)
        {
            paragraph.TextAlignment = alignment;
        }
    }

    private static void ApplyScreenplayEnterParagraphLayout(Paragraph paragraph, ScreenplayElementType screenplayType)
    {
        if (screenplayType == ScreenplayElementType.Action)
        {
            paragraph.Margin = ActionParagraphMargin;
            paragraph.TextIndent = 0.0;
            paragraph.TextAlignment = TextAlignment.Left;
            return;
        }

        ApplyParagraphLayout(paragraph, screenplayType);
    }

    private static bool ThicknessEquals(Thickness left, Thickness right)
    {
        return Math.Abs(left.Left - right.Left) <= ParagraphStyleTolerance &&
            Math.Abs(left.Top - right.Top) <= ParagraphStyleTolerance &&
            Math.Abs(left.Right - right.Right) <= ParagraphStyleTolerance &&
            Math.Abs(left.Bottom - right.Bottom) <= ParagraphStyleTolerance;
    }

    private static void ApplyParagraphRunFormatting(
        Paragraph paragraph,
        string paragraphText,
        bool boldEntireParagraph,
        IReadOnlyList<FormattingSpan> boldSpans)
    {
        var desiredSegments = CreateRunSegments(paragraphText, boldEntireParagraph, boldSpans);
        if (ParagraphRunsMatch(paragraph, desiredSegments))
        {
            return;
        }

        paragraph.Inlines.Clear();
        foreach (var segment in desiredSegments)
        {
            paragraph.Inlines.Add(new Run(segment.Text)
            {
                FontWeight = segment.IsBold ? FontWeights.Bold : FontWeights.Normal
            });
        }

        if (paragraph.Inlines.Count == 0)
        {
            paragraph.Inlines.Add(new Run(string.Empty));
        }
    }

    private static IReadOnlyList<FormattedRunSegment> CreateRunSegments(
        string paragraphText,
        bool boldEntireParagraph,
        IReadOnlyList<FormattingSpan> boldSpans)
    {
        var text = paragraphText ?? string.Empty;
        if (boldEntireParagraph)
        {
            return [new FormattedRunSegment(text, true)];
        }

        if (boldSpans.Count == 0)
        {
            return [new FormattedRunSegment(text, false)];
        }

        var segments = new List<FormattedRunSegment>();
        var nextIndex = 0;

        foreach (var boldSpan in boldSpans.OrderBy(span => span.Start))
        {
            var start = Math.Clamp(boldSpan.Start, 0, text.Length);
            var end = Math.Clamp(boldSpan.Start + boldSpan.Length, start, text.Length);
            if (start > nextIndex)
            {
                AppendRunSegment(segments, text.Substring(nextIndex, start - nextIndex), isBold: false);
            }

            if (end > start)
            {
                AppendRunSegment(segments, text.Substring(start, end - start), isBold: true);
            }

            nextIndex = Math.Max(nextIndex, end);
        }

        if (nextIndex < text.Length)
        {
            AppendRunSegment(segments, text[nextIndex..], isBold: false);
        }

        if (segments.Count == 0)
        {
            segments.Add(new FormattedRunSegment(text, false));
        }

        return segments;
    }

    private static void AppendRunSegment(List<FormattedRunSegment> segments, string text, bool isBold)
    {
        if (text.Length == 0)
        {
            return;
        }

        if (segments.Count > 0 && segments[^1].IsBold == isBold)
        {
            var previous = segments[^1];
            segments[^1] = previous with { Text = previous.Text + text };
            return;
        }

        segments.Add(new FormattedRunSegment(text, isBold));
    }

    private static bool ParagraphRunsMatch(Paragraph paragraph, IReadOnlyList<FormattedRunSegment> desiredSegments)
    {
        if (paragraph.Inlines.Count != desiredSegments.Count)
        {
            return false;
        }

        var index = 0;
        foreach (var inline in paragraph.Inlines)
        {
            if (inline is not Run run)
            {
                return false;
            }

            var desiredSegment = desiredSegments[index++];
            if (!string.Equals(run.Text, desiredSegment.Text, StringComparison.Ordinal) ||
                run.FontWeight != (desiredSegment.IsBold ? FontWeights.Bold : FontWeights.Normal))
            {
                return false;
            }
        }

        return true;
    }

    private static IReadOnlyDictionary<int, ScreenplayElementType> CreateScreenplayLineTypeLookup(ParsedScreenplay snapshot)
    {
        var lineTypes = new Dictionary<int, ScreenplayElementType>();
        if (snapshot.TitlePage != null)
        {
            for (var lineNumber = 1; lineNumber <= snapshot.TitlePage.BodyStartLineIndex; lineNumber++)
            {
                lineTypes[lineNumber] = ScreenplayElementType.TitlePageCentered;
            }

            foreach (var entry in snapshot.TitlePage.Entries)
            {
                if (entry.FieldType == TitlePageFieldType.Contact || entry.FieldType == TitlePageFieldType.DraftDate)
                {
                    var lineStart = entry.LineNumber;
                    var lineCount = entry.Value.Split('\n').Length;
                    for (int l = lineStart; l < lineStart + lineCount; l++)
                    {
                        lineTypes[l] = ScreenplayElementType.TitlePageContact;
                    }
                }
            }
        }

        var parentheticalLines = snapshot.Elements
            .OfType<ParentheticalElement>()
            .Select(element => element.LineNumber)
            .ToHashSet();

        foreach (var element in snapshot.Elements)
        {
            if (element is DialogueElement dialogue)
            {
                for (var lineNumber = dialogue.StartLine; lineNumber <= dialogue.EndLine; lineNumber++)
                {
                    if (parentheticalLines.Contains(lineNumber))
                    {
                        continue;
                    }

                    lineTypes[lineNumber] = ScreenplayElementType.Dialogue;
                }

                continue;
            }

            for (var lineNumber = element.StartLine; lineNumber <= element.EndLine; lineNumber++)
            {
                lineTypes[lineNumber] = element.Type;
            }
        }

        foreach (var lineTypeOverride in snapshot.LineTypeOverrides)
        {
            lineTypes[lineTypeOverride.Key] = lineTypeOverride.Value;
        }

        return lineTypes;
    }

    private static ScreenplayElementType ResolveScreenplayLineType(
        string lineText,
        int lineNumber,
        IReadOnlyDictionary<int, ScreenplayElementType> screenplayLineTypes)
    {
        if (screenplayLineTypes.TryGetValue(lineNumber, out var mappedType))
        {
            return mappedType;
        }

        var trimmed = (lineText ?? string.Empty).Trim();
        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        if (TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()))
        {
            return ScreenplayElementType.SceneHeading;
        }

        if (trimmed.StartsWith("(") && trimmed.EndsWith(")", StringComparison.Ordinal))
        {
            return ScreenplayElementType.Parenthetical;
        }

        return TextAnalysis.IsLiveCharacterCueCandidate(lineText.AsSpan(), 45, 6)
            ? ScreenplayElementType.Character
            : ScreenplayElementType.Action;
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

    private bool TryGetParagraphByLineNumber(int lineNumber, out Paragraph paragraph)
    {
        paragraph = null!;
        if (EditorBox is null || lineNumber < 1)
        {
            return false;
        }

        var currentLineNumber = 1;
        foreach (var block in EditorBox.Document.Blocks)
        {
            if (block is not Paragraph candidateParagraph)
            {
                continue;
            }

            if (currentLineNumber == lineNumber)
            {
                paragraph = candidateParagraph;
                return true;
            }

            currentLineNumber++;
        }

        return false;
    }

    private bool TryCreateParagraphFormattingSnapshot(Paragraph paragraph, out ParagraphFormattingSnapshot paragraphSnapshot)
    {
        paragraphSnapshot = null!;
        if (EditorBox is null)
        {
            return false;
        }

        var lineNumber = GetParagraphLineNumber(paragraph);
        ScreenplayElementType? overrideType = currentMode == WriteMode.Screenplay && ViewModel.HasLineTypeOverride(lineNumber)
            ? ViewModel.GetEffectiveLineType(lineNumber)
            : null;

        paragraphSnapshot = new ParagraphFormattingSnapshot(
            lineNumber,
            GetParagraphText(paragraph),
            paragraph.TextAlignment,
            paragraph.Margin.Left,
            overrideType);
        return true;
    }

    private static string[] SplitEditorLines(string text)
    {
        var normalizedText = (text ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
        var lines = normalizedText.Split('\n', StringSplitOptions.None);
        return lines.Length == 0 ? [string.Empty] : lines;
    }

    private ScreenplayElementType ResolveScreenplayParagraphType(
        Paragraph paragraph,
        int lineNumber,
        IReadOnlyDictionary<int, ScreenplayElementType>? screenplayLineTypes = null)
    {
        if (screenplayLineTypes is not null && screenplayLineTypes.TryGetValue(lineNumber, out var mappedType))
        {
            return mappedType;
        }

        if (screenplayLineTypes is null && ViewModel.HasLineTypeOverride(lineNumber))
        {
            return ViewModel.GetEffectiveLineType(lineNumber);
        }

        return ResolveScreenplayParagraphType(new ParagraphFormattingSnapshot(
            lineNumber,
            GetParagraphText(paragraph),
            paragraph.TextAlignment,
            paragraph.Margin.Left,
            null));
    }

    private static ScreenplayElementType ResolveScreenplayParagraphType(ParagraphFormattingSnapshot paragraphSnapshot)
    {
        if (paragraphSnapshot.OverrideType is ScreenplayElementType overrideType)
        {
            return overrideType;
        }

        var trimmed = paragraphSnapshot.Text.Trim();
        if (trimmed.Length == 0)
        {
            if (Math.Abs(paragraphSnapshot.MarginLeft - ParentheticalParagraphMargin.Left) <= ParagraphStyleTolerance)
            {
                return ScreenplayElementType.Parenthetical;
            }

            if (Math.Abs(paragraphSnapshot.MarginLeft - DialogueParagraphMargin.Left) <= ParagraphStyleTolerance)
            {
                return ScreenplayElementType.Dialogue;
            }

            if (paragraphSnapshot.TextAlignment == TextAlignment.Center)
            {
                return ScreenplayElementType.Character;
            }

            if (paragraphSnapshot.TextAlignment == TextAlignment.Right)
            {
                return ScreenplayElementType.Transition;
            }

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

        if (paragraphSnapshot.TextAlignment == TextAlignment.Right)
        {
            return ScreenplayElementType.Transition;
        }

        if (Math.Abs(paragraphSnapshot.MarginLeft - ParentheticalParagraphMargin.Left) <= ParagraphStyleTolerance)
        {
            return ScreenplayElementType.Parenthetical;
        }

        if (Math.Abs(paragraphSnapshot.MarginLeft - DialogueParagraphMargin.Left) <= ParagraphStyleTolerance)
        {
            return ScreenplayElementType.Dialogue;
        }

        if (TextAnalysis.IsLiveCharacterCueCandidate(paragraphSnapshot.Text.AsSpan(), 45, 6) ||
            paragraphSnapshot.TextAlignment == TextAlignment.Center)
        {
            return ScreenplayElementType.Character;
        }

        return ScreenplayElementType.Action;
    }

    private ScreenplayElementType ResolveScreenplayEnterSourceType(Paragraph paragraph, int lineNumber)
    {
        if (ViewModel.HasLineTypeOverride(lineNumber))
        {
            return ViewModel.GetEffectiveLineType(lineNumber);
        }

        var lineText = GetParagraphText(paragraph);
        var trimmed = lineText.Trim();
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

        if (paragraph.TextAlignment == TextAlignment.Right)
        {
            return ScreenplayElementType.Transition;
        }

        if (TextAnalysis.IsLiveCharacterCueCandidate(lineText.AsSpan(), 45, 6))
        {
            return ScreenplayElementType.Character;
        }

        return ScreenplayElementType.Action;
    }

    private Paragraph? GetActiveParagraph()
    {
        if (EditorBox is null)
        {
            return null;
        }

        return EditorBox.CaretPosition.Paragraph
            ?? EditorBox.Selection.Start.Paragraph
            ?? EditorBox.Selection.End.Paragraph
            ?? EditorBox.Document.Blocks.OfType<Paragraph>().FirstOrDefault();
    }

    private string GetParagraphText(Paragraph paragraph)
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

    private int GetParagraphLineNumber(Paragraph paragraph)
    {
        if (EditorBox is null)
        {
            return 1;
        }

        var lineNumber = 1;
        foreach (var block in EditorBox.Document.Blocks)
        {
            if (ReferenceEquals(block, paragraph))
            {
                return lineNumber;
            }

            if (block is Paragraph)
            {
                lineNumber++;
            }
        }

        return Math.Max(1, lineNumber - 1);
    }

    private void EditorBox_SelectionChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressCaretContextUpdates)
        {
            return;
        }

        UpdateCaretContextFromEditor();

        if (EditorBox is null)
        {
            return;
        }

        var hasSelection = GetEditorSelectionLength() > 0;
        ScheduleEditorInteractionRefresh(
            ensureCaretVisible: !hasSelection && EditorBox.IsKeyboardFocusWithin,
            updateDocumentHeight: false);
    }

    private void EditorBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_isSynchronizingEditorDocument ||
            _isFormatting ||
            EditorBox is null)
        {
            return;
        }

        ScheduleSentenceCapitalizationCheck();

        if (_editorTextChangedProcessingPending)
        {
            return;
        }

        _editorTextChangedProcessingPending = true;
        Dispatcher.BeginInvoke(
            new Action(ProcessEditorTextChangedAtIdle),
            DispatcherPriority.ApplicationIdle);
    }

    private void EditorBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (EditorBox is null || string.IsNullOrEmpty(e.Text))
        {
            return;
        }

        if (e.Text.Any(IsWordCompletionCharacter))
        {
            ScheduleSentenceCapitalizationCheck();
        }
    }

    private void EditorBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (EditorBox is null)
        {
            return;
        }

        if (currentMode == WriteMode.Screenplay &&
            e.Key == System.Windows.Input.Key.Enter &&
            !HasCommandModifier())
        {
            HandleScreenplayEnterKey(
                continueDialogueBlock: Keyboard.Modifiers.HasFlag(ModifierKeys.Shift));
            e.Handled = true;
            return;
        }

        if (currentMode == WriteMode.Screenplay && e.Key == System.Windows.Input.Key.Tab)
        {
            var tabLineIndex = GetEditorLineIndexFromCharacterIndex(GetEditorCaretIndex());
            if (tabLineIndex < 0)
            {
                tabLineIndex = 0;
            }

            ViewModel.UpdateCaretContext(tabLineIndex + 1);
            ViewModel.CycleCurrentLineType(!System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Shift));
            QueueEditorFormattingForActiveParagraph(force: true);
            ScheduleEditorViewportRefresh(ensureCaretVisible: true);
            e.Handled = true;
            return;
        }
    }

    private void ScheduleSentenceCapitalizationCheck()
    {
        if (EditorBox is null ||
            _isSynchronizingEditorDocument ||
            _isFormatting ||
            _sentenceCapitalizationCheckPending)
        {
            return;
        }

        _sentenceCapitalizationCheckPending = true;
        Dispatcher.BeginInvoke(
            new Action(ProcessSentenceCapitalizationCheck),
            DispatcherPriority.Input);
    }

    private void ProcessSentenceCapitalizationCheck()
    {
        _sentenceCapitalizationCheckPending = false;
        if (EditorBox is null ||
            _isSynchronizingEditorDocument ||
            _isFormatting)
        {
            return;
        }

        TryAutoCapitalizeCompletedWord();
    }

    private static bool IsWordCharacter(char ch)
    {
        return char.IsLetterOrDigit(ch) || ch is '\'' or '-';
    }

    private static bool IsWordCompletionCharacter(char ch)
    {
        return char.IsWhiteSpace(ch) ||
            ch is '.' or '!' or '?' or ',' or ';' or ':' or ')' or ']' or '}' or '"' or '”' or '’';
    }

    private static bool IsSentenceTerminator(char ch)
    {
        return ch is '.' or '!' or '?';
    }

    private static int FindLastWordCompletionCharacter(ReadOnlySpan<char> text)
    {
        for (var index = text.Length - 1; index >= 0; index--)
        {
            if (IsWordCompletionCharacter(text[index]))
            {
                return index;
            }
        }

        return -1;
    }

    private void TryAutoCapitalizeCompletedWord()
    {
        if (EditorBox is null)
        {
            return;
        }

        var editorText = GetEditorText();
        var caretIndex = Math.Clamp(GetEditorCaretIndex(), 0, editorText.Length);
        if (caretIndex <= 0)
        {
            return;
        }

        var completionIndex = FindLastWordCompletionCharacter(editorText.AsSpan(0, caretIndex));
        if (completionIndex < 0)
        {
            return;
        }

        TryAutoCapitalizeCompletedWordBeforeOffset(completionIndex + 1);
    }

    private bool TryAutoCapitalizeCompletedWordBeforeOffset(int boundaryOffset)
    {
        if (EditorBox is null)
        {
            return false;
        }

        var editorText = GetEditorText();
        if (boundaryOffset <= 0 || boundaryOffset > editorText.Length)
        {
            return false;
        }

        var scanIndex = boundaryOffset - 1;
        while (scanIndex >= 0 && IsWordCompletionCharacter(editorText[scanIndex]))
        {
            scanIndex--;
        }

        if (scanIndex < 0)
        {
            return false;
        }

        var wordEndIndex = scanIndex;
        while (wordEndIndex > 0 && IsWordCharacter(editorText[wordEndIndex - 1]))
        {
            wordEndIndex--;
        }

        var wordStartIndex = wordEndIndex;
        if (wordStartIndex < 0 || wordStartIndex >= editorText.Length || !char.IsLower(editorText[wordStartIndex]))
        {
            return false;
        }

        var boundaryIndex = wordStartIndex - 1;
        while (boundaryIndex >= 0)
        {
            var boundaryCharacter = editorText[boundaryIndex];
            if (IsSentenceTerminator(boundaryCharacter))
            {
                CapitalizeEditorCharacter(wordStartIndex, editorText[wordStartIndex]);
                return true;
            }

            if (char.IsLetterOrDigit(boundaryCharacter))
            {
                return false;
            }

            boundaryIndex--;
        }

        return false;
    }

    private void CapitalizeEditorCharacter(int characterIndex, char character)
    {
        if (EditorBox is null)
        {
            return;
        }

        try
        {
            _isFormatting = true;
            EditorBox.BeginChange();
            using var changeBlock = EditorBox.DeclareChangeBlock();

            var start = GetEditorTextPointerAtOffset(characterIndex);
            var end = GetEditorTextPointerAtOffset(characterIndex + 1);
            new TextRange(start, end).Text = char.ToUpper(character).ToString();
        }
        finally
        {
            EditorBox.EndChange();
            _isFormatting = false;
        }
    }

    private void EditorBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (EditorBox is null || !IsScreenplayMode)
        {
            ClearEditorMouseSelectionState(releaseCapture: false);
            return;
        }

        ApplyEditorDocumentLayout();
        EditorBox.UpdateLayout();

        if (!TryGetEditorCharacterIndexFromPoint(e.GetPosition(EditorBox), out var characterIndex))
        {
            ClearEditorMouseSelectionState(releaseCapture: true);
            return;
        }

        EditorBox.Focus();
        Keyboard.Focus(EditorBox);

        if (e.ClickCount == 2)
        {
            SelectWordAtCharacterIndex(characterIndex);
            ClearEditorMouseSelectionState(releaseCapture: true);
            InvalidateEditorCueOverlay();
            e.Handled = true;
            return;
        }

        if (e.ClickCount != 1)
        {
            ClearEditorMouseSelectionState(releaseCapture: true);
            return;
        }

        var anchorIndex = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)
            ? GetEditorSelectionAnchorIndex()
            : characterIndex;

        SetEditorSelection(anchorIndex, characterIndex);
        InvalidateEditorCueOverlay();
        _editorMouseSelectionAnchorIndex = anchorIndex;
        _isEditorMouseSelectionActive = true;
        EditorBox.CaptureMouse();
        e.Handled = true;
    }

    private void EditorBox_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (EditorBox is null || !IsScreenplayMode || !_isEditorMouseSelectionActive || _editorMouseSelectionAnchorIndex is null)
        {
            return;
        }

        if (e.LeftButton != MouseButtonState.Pressed)
        {
            ClearEditorMouseSelectionState(releaseCapture: true);
            return;
        }

        if (!TryGetEditorCharacterIndexFromPoint(e.GetPosition(EditorBox), out var characterIndex))
        {
            return;
        }

        SetEditorSelection(_editorMouseSelectionAnchorIndex.Value, characterIndex);
        InvalidateEditorCueOverlay();
        e.Handled = true;
    }

    private void EditorBox_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (EditorBox is null || !IsScreenplayMode || !_isEditorMouseSelectionActive || _editorMouseSelectionAnchorIndex is null)
        {
            ClearEditorMouseSelectionState(releaseCapture: false);
            return;
        }

        if (TryGetEditorCharacterIndexFromPoint(e.GetPosition(EditorBox), out var characterIndex))
        {
            SetEditorSelection(_editorMouseSelectionAnchorIndex.Value, characterIndex);
            InvalidateEditorCueOverlay();
            e.Handled = true;
        }

        ClearEditorMouseSelectionState(releaseCapture: true);
    }

    private void EditorBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (EditorBox is null)
        {
            return;
        }

        if (!TryGetEditorCharacterIndexFromPoint(e.GetPosition(EditorBox), out var characterIndex))
        {
            return;
        }

        var selectionStart = GetEditorSelectionStart();
        var selectionLength = GetEditorSelectionLength();
        if (selectionLength <= 0)
        {
            return;
        }

        var selectionEnd = selectionStart + selectionLength;
        if (characterIndex < selectionStart || characterIndex > selectionEnd)
        {
            return;
        }

        e.Handled = true;

        Dispatcher.BeginInvoke(
            new Action(() =>
            {
                if (EditorBox?.ContextMenu is not ContextMenu contextMenu)
                {
                    return;
                }

                contextMenu.PlacementTarget = EditorBox;
                contextMenu.IsOpen = true;
            }),
            DispatcherPriority.Input);
    }

    private void EditorBox_Loaded(object sender, RoutedEventArgs e)
    {
        ApplyEditorDocumentLayout();
        Dispatcher.BeginInvoke(
            new Action(() => RestoreEditorFocusAndCaret(ensureFocus: true)),
            System.Windows.Threading.DispatcherPriority.Loaded);
    }

    private void EditorBox_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyEditorDocumentLayout();
        ScheduleEditorViewportRefresh(ensureCaretVisible: false);
    }

    private void EditorBox_GotFocus(object sender, RoutedEventArgs e)
    {
        RestoreEditorFocusAndCaret(ensureFocus: true);
        InvalidateEditorCueOverlay();
        ScheduleEditorInteractionRefresh(ensureCaretVisible: true);
    }

    private void EditorBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        ClearEditorMouseSelectionState(releaseCapture: true);
        FlushPendingEditorChangeBatch();
    }

    private void ProcessEditorTextChangedAtIdle()
    {
        _editorTextChangedProcessingPending = false;
        if (EditorBox is null)
        {
            return;
        }

        if (_isSynchronizingEditorDocument || _isFormatting)
        {
            return;
        }

        var currentParagraph = EditorBox.CaretPosition.Paragraph;

        if (currentParagraph is null)
        {
            _editorTextSyncPending = true;
            RestartEditorDocumentSyncTimer();
            return;
        }

        var paragraphText = GetParagraphText(currentParagraph);
        var lineNumber = GetParagraphLineNumber(currentParagraph);
        var paragraphTextChanged =
            !ReferenceEquals(_lastTextChangedParagraph, currentParagraph) ||
            !string.Equals(_lastTextChangedParagraphText, paragraphText, StringComparison.Ordinal);

        _lastTextChangedParagraph = currentParagraph;
        _lastTextChangedParagraphText = paragraphText;
        var releasedAutomaticActionOverride = ViewModel.ReleaseAutomaticActionLineOverrideIfNeeded(lineNumber, paragraphText);
        _editorTextSyncPending = true;
        RestartEditorDocumentSyncTimer();

        if (releasedAutomaticActionOverride)
        {
            PushEditorTextToViewModel(GetEditorText());
            ViewModel.RefreshParsedSnapshotNow();
        }

        if (!paragraphTextChanged || paragraphText.Length == 0)
        {
            return;
        }

        QueuePreviewFormattingForParagraph(currentParagraph);
    }

    private void HandleScreenplayEnterKey(bool continueDialogueBlock)
    {
        if (EditorBox is null || currentMode != WriteMode.Screenplay)
        {
            return;
        }

        EnsureEditorDocumentInitialized();

        var currentParagraph = GetActiveParagraph()
            ?? EditorBox.Document.Blocks.OfType<Paragraph>().FirstOrDefault();
        if (currentParagraph is null)
        {
            return;
        }

        var currentLineNumber = GetParagraphLineNumber(currentParagraph);
        var lineText = GetParagraphText(currentParagraph);
        var currentElementType = ResolveScreenplayParagraphType(currentParagraph, currentLineNumber);
        var continuationType = continueDialogueBlock &&
            currentElementType is ScreenplayElementType.Dialogue or ScreenplayElementType.Parenthetical
                ? ScreenplayElementType.Dialogue
                : DetermineScreenplayEnterContinuation(currentElementType, lineText);
        var previousLineCount = RichTextBoxTextUtilities.GetLineCount(EditorBox);
        var previousParagraphBeforeInsert = currentParagraph.PreviousBlock as Paragraph;
        var nextParagraphBeforeInsert = currentParagraph.NextBlock as Paragraph;
        Paragraph? newParagraph = null;

        try
        {
            EditorBox.BeginChange();
            using var changeBlock = EditorBox.DeclareChangeBlock();

            if (!EditorBox.Selection.IsEmpty)
            {
                EditorBox.Selection.Text = string.Empty;
            }

            var insertionPosition = EditorBox.CaretPosition.GetInsertionPosition(LogicalDirection.Forward)
                ?? EditorBox.CaretPosition.GetInsertionPosition(LogicalDirection.Backward)
                ?? EditorBox.Document.ContentEnd;
            var paragraphBreakPosition = insertionPosition.InsertParagraphBreak();
            newParagraph = ResolveInsertedParagraph(
                currentParagraph,
                previousParagraphBeforeInsert,
                nextParagraphBeforeInsert,
                paragraphBreakPosition);

            if (newParagraph is null)
            {
                return;
            }

            ApplyScreenplayEnterParagraphLayout(newParagraph, continuationType);
            EditorBox.CaretPosition = newParagraph.ContentStart;
        }
        finally
        {
            EditorBox.EndChange();
        }

        if (newParagraph is null)
        {
            return;
        }

        ApplyScreenplayContinuationToLine(
            currentElementType,
            continuationType,
            GetParagraphLineNumber(newParagraph),
            previousLineCount);
    }

    private Paragraph? ResolveInsertedParagraph(
        Paragraph currentParagraph,
        Paragraph? previousParagraphBeforeInsert,
        Paragraph? nextParagraphBeforeInsert,
        TextPointer paragraphBreakPosition)
    {
        if (EditorBox is null)
        {
            return null;
        }

        return GetInsertedParagraphCandidate(EditorBox.CaretPosition.Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert)
            ?? GetInsertedParagraphCandidate(currentParagraph.PreviousBlock as Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert)
            ?? GetInsertedParagraphCandidate(currentParagraph.NextBlock as Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert)
            ?? GetInsertedParagraphCandidate(paragraphBreakPosition.GetInsertionPosition(LogicalDirection.Backward)?.Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert)
            ?? GetInsertedParagraphCandidate(paragraphBreakPosition.Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert)
            ?? GetInsertedParagraphCandidate(paragraphBreakPosition.GetInsertionPosition(LogicalDirection.Forward)?.Paragraph, currentParagraph, previousParagraphBeforeInsert, nextParagraphBeforeInsert);
    }

    private static Paragraph? GetInsertedParagraphCandidate(
        Paragraph? candidate,
        Paragraph currentParagraph,
        Paragraph? previousParagraphBeforeInsert,
        Paragraph? nextParagraphBeforeInsert)
    {
        if (candidate is null ||
            ReferenceEquals(candidate, currentParagraph) ||
            ReferenceEquals(candidate, previousParagraphBeforeInsert) ||
            ReferenceEquals(candidate, nextParagraphBeforeInsert))
        {
            return null;
        }

        return candidate;
    }

    private void ToggleWriteMode()
    {
        FlushPendingEditorChangeBatch();

        currentMode = currentMode == WriteMode.Screenplay
            ? WriteMode.Markdown
            : WriteMode.Screenplay;
        UpdateUIForMode();
    }

    private void UpdateUIForMode()
    {
        if (txtModeStatus is not null)
        {
            txtModeStatus.Text = currentMode switch
            {
                WriteMode.Screenplay => "MODE: SCREENPLAY",
                WriteMode.Markdown => "MODE: MARKDOWN",
                _ => "MODE: UNKNOWN"
            };
        }

        if (EditorBox is not null)
        {
            EditorBox.AcceptsTab = currentMode == WriteMode.Markdown;
        }

        ApplyEditorModeVisualState();
        QueueEditorFormattingForDocument(immediate: true);

        if (IsLoaded)
        {
            ScheduleEditorViewportRefresh(ensureCaretVisible: EditorBox is not null && EditorBox.IsKeyboardFocusWithin);
        }
    }

    private void ApplyEditorModeVisualState()
    {
        if (EditorBox is null)
        {
            return;
        }

        EnsureEditorDocumentInitialized();
        ApplyEditorDocumentLayout();
        EditorBox.SetResourceReference(System.Windows.Controls.Control.BackgroundProperty, "WindowBackground");

        EditorBox.Foreground = Brushes.Transparent;
        EditorBox.SelectionBrush = Brushes.Transparent;
        EditorBox.CaretBrush = Brushes.Transparent;
        EnsureEditorCueOverlay();

        if (_editorCueAdorner is not null)
        {
            _editorCueAdorner.SetFormattingEnabled(true);
            _editorCueAdorner.Visibility = Visibility.Visible;
        }

        ApplyEditorCaretBrush();
    }

    private void ApplyEditorDocumentLayout()
    {
        if (EditorBox?.Document is null)
        {
            return;
        }

        var document = EditorBox.Document;
        var pagePadding = currentMode == WriteMode.Screenplay
            ? ScreenplayDocumentPagePadding
            : MarkdownDocumentPagePadding;
        EditorBox.Padding = pagePadding;
        document.PagePadding = pagePadding;
        document.ColumnWidth = double.PositiveInfinity;
        UpdateEditorDocumentPageWidth();

        EditorBox.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
        EditorBox.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
        document.SetResourceReference(FlowDocument.BackgroundProperty, "WindowBackground");
    }

    private void UpdateEditorDocumentPageWidth()
    {
        if (EditorBox?.Document is null)
        {
            return;
        }

        var pageWidth = GetEditorDocumentViewportWidth();

        EditorBox.Document.PageWidth = Math.Max(MinimumEditorPageWidth, pageWidth);
    }

    private bool TryGetEditorCharacterIndexFromPoint(Point point, out int characterIndex)
    {
        characterIndex = 0;

        if (EditorBox is null)
        {
            return false;
        }

        if (IsEditorCueOverlayVisible())
        {
            EnsureEditorCueOverlay();
            if (_editorCueAdorner is not null &&
                _editorCueAdorner.Visibility == Visibility.Visible &&
                _editorCueAdorner.TryGetCharacterIndexFromPoint(point, out characterIndex))
            {
                return true;
            }
        }

        return RichTextBoxTextUtilities.TryGetCharacterIndexFromPoint(EditorBox, point, out characterIndex);
    }

    private void SelectWordAtCharacterIndex(int characterIndex)
    {
        var text = GetEditorText();
        if (text.Length == 0)
        {
            SetEditorSelection(0, 0);
            return;
        }

        var safeIndex = Math.Clamp(characterIndex, 0, text.Length);
        if (safeIndex == text.Length && safeIndex > 0)
        {
            safeIndex--;
        }

        if (safeIndex < 0 || safeIndex >= text.Length)
        {
            SetEditorSelection(0, 0);
            return;
        }

        if (char.IsWhiteSpace(text[safeIndex]))
        {
            SetEditorSelection(safeIndex, safeIndex);
            return;
        }

        var start = safeIndex;
        while (start > 0 && !char.IsWhiteSpace(text[start - 1]))
        {
            start--;
        }

        var end = safeIndex + 1;
        while (end < text.Length && !char.IsWhiteSpace(text[end]))
        {
            end++;
        }

        SetEditorSelection(start, end);
    }

    private double GetEditorDocumentHostWidth()
    {
        if (EditorDocumentHost is not null)
        {
            if (EditorDocumentHost.ActualWidth > MinimumEditorPageWidth)
            {
                return EditorDocumentHost.ActualWidth;
            }

            if (!double.IsNaN(EditorDocumentHost.Width) && !double.IsInfinity(EditorDocumentHost.Width))
            {
                return Math.Max(MinimumEditorPageWidth, EditorDocumentHost.Width);
            }
        }

        return EditorBox is not null && EditorBox.ActualWidth > MinimumEditorPageWidth
            ? EditorBox.ActualWidth
            : MinimumEditorPageWidth;
    }

    private double GetEditorDocumentViewportWidth()
    {
        var hostWidth = GetEditorDocumentHostWidth();
        if (EditorBox is null)
        {
            return hostWidth;
        }

        var horizontalPadding = Math.Max(0.0, EditorBox.Padding.Left) + Math.Max(0.0, EditorBox.Padding.Right);
        return Math.Max(MinimumEditorPageWidth, hostWidth - horizontalPadding);
    }

    private void ClearEditorMouseSelectionState(bool releaseCapture)
    {
        _isEditorMouseSelectionActive = false;
        _editorMouseSelectionAnchorIndex = null;

        if (releaseCapture && EditorBox is not null && EditorBox.IsMouseCaptured)
        {
            EditorBox.ReleaseMouseCapture();
        }
    }

    private void ApplyScreenplayContinuationToLine(
        ScreenplayElementType sourceElementType,
        ScreenplayElementType continuationType,
        int targetLineNumber,
        int previousLineCount)
    {
        if (currentMode != WriteMode.Screenplay || EditorBox is null)
        {
            return;
        }

        var currentLineCount = RichTextBoxTextUtilities.GetLineCount(EditorBox);
        var lineDelta = currentLineCount - previousLineCount;
        if (lineDelta != 0)
        {
            ViewModel.ShiftLineTypeOverrides(targetLineNumber, lineDelta);
        }

        if (continuationType is ScreenplayElementType.Dialogue or ScreenplayElementType.Parenthetical)
        {
            ViewModel.SetLineTypeOverride(targetLineNumber, continuationType);
        }
        else if (continuationType == ScreenplayElementType.Action &&
                 sourceElementType == ScreenplayElementType.Dialogue)
        {
            ViewModel.SetAutomaticActionLineOverride(targetLineNumber);
        }
        else if (lineDelta > 0)
        {
            ViewModel.ClearLineTypeOverride(targetLineNumber);
        }

        PushEditorTextToViewModel(GetEditorText());
        ViewModel.RefreshParsedSnapshotNow();

        if (TryGetParagraphByLineNumber(targetLineNumber, out var targetParagraph))
        {
            ViewModel.UpdateEnterContinuation(targetLineNumber, GetParagraphText(targetParagraph));
        }

        ViewModel.UpdateCaretContext(targetLineNumber);
        InvalidateEditorCueOverlay();
        ScheduleEditorViewportRefresh(ensureCaretVisible: true);
    }

    private static ScreenplayElementType DetermineScreenplayEnterContinuation(
        ScreenplayElementType currentElementType,
        string currentLineText)
    {
        var trimmed = currentLineText.Trim();
        if (trimmed.Length == 0)
        {
            return ScreenplayElementType.Action;
        }

        // Section and Synopsis markers always lead to Action
        if (trimmed.StartsWith('#') || trimmed.StartsWith('='))
        {
            return ScreenplayElementType.Action;
        }

        return currentElementType is ScreenplayElementType.Character or ScreenplayElementType.Parenthetical
            ? ScreenplayElementType.Dialogue
            : ScreenplayElementType.Action;
    }

    private static bool HasCommandModifier()
    {
        return Keyboard.Modifiers.HasFlag(ModifierKeys.Control)
            || Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)
            || Keyboard.Modifiers.HasFlag(ModifierKeys.Windows);
    }

    private void EditorBox_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
    {
        if (!IsLoaded || EditorBox is null)
        {
            return;
        }

        if (!IsEditorCueOverlayVisible())
        {
            return;
        }

        // Suppress WPF's default focus/caret scrolling and route visibility
        // through the page-aware viewport logic instead.
        e.Handled = true;
        ScheduleEditorInteractionRefresh(ensureCaretVisible: true);
    }

    private void OutlineTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (_suppressOutlineSelectionNavigation)
        {
            return;
        }

        if (sender is not TreeView treeView)
        {
            return;
        }

        // Outline refreshes rebuild the tree and can re-raise selection changes for the
        // previously selected node. Ignore those synthetic changes so typing doesn't
        // trigger another navigation back to the selected outline card.
        if (!treeView.IsKeyboardFocusWithin && Mouse.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        if (e.NewValue is not OutlineNodeViewModel node)
        {
            return;
        }

        ViewModel.SelectedOutlineLineNumber = node.LineNumber;
        RestoreOutlineSelection(node.LineNumber);
        NavigateEditorToLine(node.LineNumber);
    }

    private Point _workspaceDragStartPoint;
    private bool _isWorkspaceDragging;
    private OutlineNodeViewModel? _draggedWorkspaceNode;
    private OutlineNodeViewModel? _lastHoveredWorkspaceNode;

    private void WorkspaceNode_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Border border || border.DataContext is not OutlineNodeViewModel node)
        {
            return;
        }

        _workspaceDragStartPoint = e.GetPosition(this);
        _draggedWorkspaceNode = node;

        // Set visual selection in sidebar
        if (node.LineNumber > 0)
        {
            ViewModel.SelectedOutlineLineNumber = node.LineNumber;
            RestoreOutlineSelection(node.LineNumber);
        }

        // Use high-priority dispatcher call to ensure we win the focus battle
        // against the TreeView's selection logic and draw the caret immediately.
        Dispatcher.BeginInvoke(() => {
            NavigateEditorToLine(node.LineNumber);
            EditorBox?.Focus();
        }, System.Windows.Threading.DispatcherPriority.Input);

        // Do not mark handled yet so we can detect MouseMove for drag
        // e.Handled = true; 
    }

    private void WorkspaceNode_MouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _draggedWorkspaceNode == null || _isWorkspaceDragging)
        {
            return;
        }

        var currentPoint = e.GetPosition(this);
        if (Math.Abs(currentPoint.X - _workspaceDragStartPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
            Math.Abs(currentPoint.Y - _workspaceDragStartPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
        {
            _isWorkspaceDragging = true;
            
            // Show ghost before starting drag loop
            UpdateWorkspaceGhostVisibility(true);
            
            DragDrop.DoDragDrop((DependencyObject)sender, _draggedWorkspaceNode, DragDropEffects.Move);
            
            // Cleanup after drag completes (either dropped or cancelled)
            _isWorkspaceDragging = false;
            _draggedWorkspaceNode = null;
            UpdateWorkspaceGhostVisibility(false);
        }
    }

    private void UpdateWorkspaceGhostVisibility(bool visible)
    {
        var visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        var node = visible ? _draggedWorkspaceNode : null;

        if (WorkspaceGhostVisual != null)
        {
            WorkspaceGhostVisual.Visibility = visibility;
            WorkspaceGhostVisual.DataContext = node;
        }
        if (NotesGhostVisual != null)
        {
            NotesGhostVisual.Visibility = visibility;
            NotesGhostVisual.DataContext = node;
        }
    }

    private void UpdateWorkspaceGhostPosition(DragEventArgs e)
    {
        bool isNotes = LeftDockTabs.SelectedIndex == 1;
        var canvas = isNotes ? NotesDragCanvas : WorkspaceDragCanvas;
        var ghost = isNotes ? NotesGhostVisual : WorkspaceGhostVisual;

        if (ghost != null && canvas != null)
        {
            var pos = e.GetPosition(canvas);
            
            // Center the ghost on the cursor using ActualWidth/Height
            double width = ghost.ActualWidth > 0 ? ghost.ActualWidth : 240;
            double height = ghost.ActualHeight > 0 ? ghost.ActualHeight : 60;
            
            Canvas.SetLeft(ghost, pos.X - (width / 2));
            Canvas.SetTop(ghost, pos.Y - (height / 2));
        }
    }

    private void WorkspaceNode_DragOver(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(typeof(OutlineNodeViewModel)) || _draggedWorkspaceNode == null)
        {
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        var draggedNode = (OutlineNodeViewModel)e.Data.GetData(typeof(OutlineNodeViewModel));
        var treeView = sender is DependencyObject d ? FindAncestor<TreeView>(d) : null;
        if (treeView == null) return;

        // Auto-scrolling
        HandleWorkspaceAutoScroll(treeView, e.GetPosition(treeView));

        // Update ghost position
        UpdateWorkspaceGhostPosition(e);

        // Hit-testing
        var hitResult = VisualTreeHelper.HitTest(treeView, e.GetPosition(treeView));
        if (hitResult?.VisualHit == null) return;

        var targetItem = FindAncestor<TreeViewItem>(hitResult.VisualHit);
        if (targetItem == null || targetItem.DataContext is not OutlineNodeViewModel targetNode)
        {
            if (_lastHoveredWorkspaceNode != null)
            {
                _lastHoveredWorkspaceNode.IsDragOver = false;
                _lastHoveredWorkspaceNode = null;
            }
            return;
        }

        if (ReferenceEquals(draggedNode, targetNode))
        {
            if (_lastHoveredWorkspaceNode != null)
            {
                _lastHoveredWorkspaceNode.IsDragOver = false;
                _lastHoveredWorkspaceNode = null;
            }
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (!ReferenceEquals(targetNode, _lastHoveredWorkspaceNode))
        {
            if (_lastHoveredWorkspaceNode != null)
            {
                _lastHoveredWorkspaceNode.IsDragOver = false;
            }
            targetNode.IsDragOver = true;
            _lastHoveredWorkspaceNode = targetNode;
        }

        var position = e.GetPosition(targetItem);
        var height = targetItem.ActualHeight;
        
        // 3-zone hit-testing for reordering vs nesting
        // Top 25%: Above
        // Middle 50%: Onto (Nest)
        // Bottom 25%: Below
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

        ViewModel.SelectedDocument?.PerformWorkspaceReorderPreview(draggedNode, targetNode, dropPos);
        e.Effects = DragDropEffects.Move;
        e.Handled = true;
    }

    private void WorkspaceNode_Drop(object sender, DragEventArgs e)
    {
        if (_lastHoveredWorkspaceNode != null)
        {
            _lastHoveredWorkspaceNode.IsDragOver = false;
            _lastHoveredWorkspaceNode = null;
        }

        if (_isWorkspaceDragging)
        {
            ViewModel.SelectedDocument?.FinalizeWorkspaceReorderPreview();
            _isWorkspaceDragging = false;
            _draggedWorkspaceNode = null;
            UpdateWorkspaceGhostVisibility(false);
        }
        e.Handled = true;
    }

    private void WorkspaceNode_DragLeave(object sender, DragEventArgs e)
    {
        // Highlight cleanup is handled by DragOver (on node change) 
        // and Drop/Cancel (on drag completion). 
        // We leave this empty to prevent flickering when moving over child elements.
    }

    private void CancelWorkspaceDrag()
    {
        if (_lastHoveredWorkspaceNode != null)
        {
            _lastHoveredWorkspaceNode.IsDragOver = false;
            _lastHoveredWorkspaceNode = null;
        }

        UpdateWorkspaceGhostVisibility(false);
        ViewModel.SelectedDocument?.CancelWorkspaceReorderPreview();
        _isWorkspaceDragging = false;
        _draggedWorkspaceNode = null;
    }

    private void HandleWorkspaceAutoScroll(TreeView treeView, Point mousePos)
    {
        var scrollViewer = FindDescendant<ScrollViewer>(treeView);
        if (scrollViewer == null) return;

        double scrollMargin = 40.0;
        double scrollStep = 10.0;

        if (mousePos.Y < scrollMargin)
        {
            scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - scrollStep);
        }
        else if (mousePos.Y > treeView.ActualHeight - scrollMargin)
        {
            scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset + scrollStep);
        }
    }

    private void ViewModel_OutlineUpdated(object? sender, EventArgs e)
    {
        if (IsEditorCueOverlayVisible())
        {
            _editorCueAdorner?.NotifyParsedLayoutChanged();
        }

        if (ViewModel.SelectedOutlineLineNumber is not int selectedLineNumber || selectedLineNumber < 1)
        {
            ScheduleEditorViewportRefresh(ensureCaretVisible: false);
            return;
        }

        if (EditorBox is not null && EditorBox.IsKeyboardFocusWithin)
        {
            // Do not jump if user is typing
            ScheduleEditorViewportRefresh(ensureCaretVisible: false);
            return;
        }

        Dispatcher.BeginInvoke(() => RestoreOutlineSelection(selectedLineNumber), System.Windows.Threading.DispatcherPriority.ContextIdle);
        ScheduleEditorViewportRefresh(ensureCaretVisible: false);
    }

    private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (string.IsNullOrEmpty(e.PropertyName) ||
            e.PropertyName == nameof(ShellViewModel.SelectedDocument))
        {
            AttachBeatBoardDocumentHandlers(ViewModel.SelectedDocument);
            ViewModel.SetBoardModeActive(_currentWorkspaceSurface == WorkspaceSurface.BeatBoard);
            ScheduleBeatBoardRefresh();
        }

        var isEditorDrivenDocumentTextChange =
            _suppressEditorDocumentReload &&
            e.PropertyName == nameof(ShellViewModel.DocumentText);

        if (!isEditorDrivenDocumentTextChange &&
            (e.PropertyName == nameof(ShellViewModel.DocumentText) ||
            e.PropertyName == nameof(ShellViewModel.SelectedDocument)))
        {
            LoadEditorFromViewModel(resetSelection: e.PropertyName != nameof(ShellViewModel.DocumentText));
        }

        if (!isEditorDrivenDocumentTextChange &&
            (string.IsNullOrEmpty(e.PropertyName) ||
            e.PropertyName == nameof(ShellViewModel.DocumentText) ||
            e.PropertyName == nameof(ShellViewModel.SelectedDocument)))
        {
            ScheduleEditorViewportRefresh(ensureCaretVisible: false);
        }

        if (e.PropertyName is nameof(ShellViewModel.EditorZoomPercent)
            or nameof(ShellViewModel.EditorZoomScale)
            or nameof(ShellViewModel.EditorZoomDisplayText))
        {
            ScheduleEditorViewportRefresh(ensureCaretVisible: false);
        }

        if (e.PropertyName == nameof(ShellViewModel.CurrentLineNumber))
        {
            RestoreOutlineSelection(ViewModel.CurrentLineNumber);
        }
    }

    private void ThemeManager_ThemeChanged(object? sender, EventArgs e)
    {
        if (!IsLoaded)
        {
            return;
        }

        ApplyEditorModeVisualState();
        SetBeatBoardDeleteButtonState(isActive: false);

        if (IsEditorCueOverlayVisible())
        {
            InvalidateEditorCueOverlay();
        }
    }

    private void InitializeThemeSelector()
    {
        if (ThemeComboBox is null)
        {
            return;
        }

        _suppressThemeSelectionChanged = true;
        try
        {
            ThemeComboBox.SelectedIndex = GetThemeSelectionIndex(ThemeManager.CurrentThemeName);
        }
        finally
        {
            _suppressThemeSelectionChanged = false;
        }
    }

    private static int GetThemeSelectionIndex(string themeName)
    {
        return themeName switch
        {
            ThemeManager.LightThemeName => 0,
            ThemeManager.DarkThemeName => 1,
            _ => 2
        };
    }

    private void ThemeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_suppressThemeSelectionChanged || sender is not System.Windows.Controls.ComboBox comboBox)
        {
            return;
        }

        switch (comboBox.SelectedIndex)
        {
            case 0:
                ThemeManager.SetTheme(ThemeManager.LightThemeName);
                break;
            case 1:
                ThemeManager.SetTheme(ThemeManager.DarkThemeName);
                break;
            default:
                ThemeManager.SetTheme(ThemeManager.SystemThemeName);
                break;
        }
    }

    private void EditorScrollHost_ScrollChanged(object sender, System.Windows.Controls.ScrollChangedEventArgs e)
    {
        if (!IsEditorCueOverlayVisible())
        {
            return;
        }

        InvalidateEditorCueOverlay();
    }

    private void EditorScrollHost_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (EditorScrollHost is null || e.Delta == 0)
        {
            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            e.Handled = true;

            CaptureEditorViewportAnchorForZoomChange();
            if (e.Delta > 0)
            {
                ViewModel.IncreaseEditorZoom();
            }
            else
            {
                ViewModel.DecreaseEditorZoom();
            }

            RefreshEditorViewportAfterZoomChange();
            return;
        }

        if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
        {
            e.Handled = true;
            QueueEditorWheelScroll(
                horizontalDelta: e.Delta / 120.0 * GetEditorWheelScrollStep() * EditorWheelScrollMultiplier * Math.Max(0.01, ViewModel.EditorZoomScale));
            return;
        }

        e.Handled = true;

        // Keep wheel movement tied to the editor's line height, but much gentler than the default jump.
        QueueEditorWheelScroll(
            verticalDelta: -e.Delta / 120.0 * GetEditorWheelScrollStep() * EditorWheelScrollMultiplier * Math.Max(0.01, ViewModel.EditorZoomScale));
    }

    private void EditorScrollHost_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyEditorDocumentLayout();
        ScheduleEditorViewportRefresh(ensureCaretVisible: false);
    }

    private void UpdateCaretContextFromEditor()
    {
        if (EditorBox is null)
        {
            return;
        }

        UpdateCaretContextFromCharacterIndex(GetEditorCaretIndex());
    }

    private void UpdateCaretContextFromCharacterIndex(int characterIndex)
    {
        if (!IsLoaded || EditorBox is null)
        {
            return;
        }

        var textLength = GetEditorText().Length;
        var safeCharacterIndex = Math.Clamp(characterIndex, 0, textLength);
        var lineIndex = GetEditorLineIndexFromCharacterIndex(safeCharacterIndex);
        if (lineIndex < 0)
        {
            lineIndex = 0;
        }

        ViewModel.UpdateCaretContext(lineIndex + 1);
    }

    private void EnsureEditorCueOverlay()
    {
        if (_editorCueAdorner is not null || EditorBox is null)
        {
            return;
        }

        var layer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(EditorBox);
        if (layer is null)
        {
            return;
        }

        _editorCueAdorner = new EditorCueAdorner(
            EditorBox,
            EditorScrollHost,
            () => ViewModel.GetLatestParsedSnapshot(),
            () => ViewModel.GetLatestParsedSnapshot().LineTypeOverrides,
            () => ViewModel.EditorZoomScale,
            () => IsScreenplayMode);
        _editorCueAdorner.SetFormattingEnabled(IsScreenplayMode);
        _editorCueAdorner.Visibility = Visibility.Visible;
        _editorCueAdorner.LayoutChanged += EditorCueAdorner_LayoutChanged;
        layer.Add(_editorCueAdorner);
    }

    private void EditorCueAdorner_LayoutChanged(object? sender, EventArgs e)
    {
        if (!IsEditorCueOverlayVisible())
        {
            return;
        }

        var ensureCaretVisible = _ensureCaretVisibleAfterCueLayoutRefresh;
        _ensureCaretVisibleAfterCueLayoutRefresh = false;

        ScheduleEditorInteractionRefresh(
            ensureCaretVisible: ensureCaretVisible,
            updateDocumentHeight: true);
    }

    private void InvalidateEditorCueOverlay()
    {
        if (_editorCueAdorner is null)
        {
            EnsureEditorCueOverlay();
        }

        if (!IsEditorCueOverlayVisible())
        {
            return;
        }

        _editorCueAdorner!.InvalidateVisual();
    }

    private void ScheduleEditorCueOverlayInvalidation()
    {
        if (!IsLoaded || !IsEditorCueOverlayVisible())
        {
            return;
        }

        if (_editorCueInvalidatePending)
        {
            return;
        }

        _editorCueInvalidatePending = true;
        Dispatcher.BeginInvoke(
            new Action(() =>
            {
                _editorCueInvalidatePending = false;
                InvalidateEditorCueOverlay();
            }),
            System.Windows.Threading.DispatcherPriority.Render);
    }

    private void QueueEditorWheelScroll(double verticalDelta = 0.0, double horizontalDelta = 0.0)
    {
        if (!IsLoaded || EditorScrollHost is null)
        {
            return;
        }

        var delta = new Vector(horizontalDelta, verticalDelta);
        if (Math.Abs(delta.X) < 0.01 && Math.Abs(delta.Y) < 0.01)
        {
            return;
        }

        _pendingEditorWheelDelta += delta;

        if (_editorWheelScrollAnimationPending)
        {
            return;
        }

        _editorWheelScrollAnimationPending = true;
        CompositionTarget.Rendering += EditorWheelScroll_Rendering;
    }

    private void EditorWheelScroll_Rendering(object? sender, EventArgs e)
    {
        if (!IsLoaded || EditorScrollHost is null)
        {
            _pendingEditorWheelDelta = new Vector();
            StopEditorWheelScrollAnimation();
            return;
        }

        var delta = _pendingEditorWheelDelta;
        if (Math.Abs(delta.X) < 0.01 && Math.Abs(delta.Y) < 0.01)
        {
            StopEditorWheelScrollAnimation();
            return;
        }

        _pendingEditorWheelDelta = new Vector();

        var targetHorizontalOffset = Math.Clamp(
            EditorScrollHost.HorizontalOffset + delta.X,
            0.0,
            EditorScrollHost.ScrollableWidth);

        var targetVerticalOffset = Math.Clamp(
            EditorScrollHost.VerticalOffset + delta.Y,
            0.0,
            EditorScrollHost.ScrollableHeight);

        var moved = false;

        if (Math.Abs(EditorScrollHost.HorizontalOffset - targetHorizontalOffset) > 0.01)
        {
            EditorScrollHost.ScrollToHorizontalOffset(targetHorizontalOffset);
            moved = true;
        }

        if (Math.Abs(EditorScrollHost.VerticalOffset - targetVerticalOffset) > 0.01)
        {
            EditorScrollHost.ScrollToVerticalOffset(targetVerticalOffset);
            moved = true;
        }

        if (moved)
        {
            InvalidateEditorCueOverlay();
        }

        if (Math.Abs(_pendingEditorWheelDelta.X) < 0.01 && Math.Abs(_pendingEditorWheelDelta.Y) < 0.01)
        {
            StopEditorWheelScrollAnimation();
        }
    }

    private void StopEditorWheelScrollAnimation()
    {
        if (!_editorWheelScrollAnimationPending)
        {
            return;
        }

        CompositionTarget.Rendering -= EditorWheelScroll_Rendering;
        _editorWheelScrollAnimationPending = false;
    }

    private double GetEditorWheelScrollStep()
    {
        if (EditorBox is null)
        {
            return 16.0;
        }

        var step = EditorBox.FontSize * EditorBox.FontFamily.LineSpacing;
        if (double.IsNaN(step) || step <= 0.0)
        {
            return 16.0;
        }

        return Math.Max(12.0, step);
    }

    private double GetEditorZoomScale()
    {
        return Math.Max(0.01, ViewModel.EditorZoomScale);
    }

    private void ScheduleEditorInteractionRefresh(bool ensureCaretVisible, bool updateDocumentHeight = false)
    {
        if (!IsLoaded)
        {
            return;
        }

        _ensureCaretVisibleAfterInteractionRefresh |= ensureCaretVisible;
        _updateEditorDocumentHeightAfterInteractionRefresh |= updateDocumentHeight;
        if (_editorInteractionRefreshPending)
        {
            return;
        }

        _editorInteractionRefreshPending = true;
        Dispatcher.BeginInvoke(ApplyEditorInteractionRefresh, DispatcherPriority.Background);
    }

    private void ApplyEditorInteractionRefresh()
    {
        _editorInteractionRefreshPending = false;

        if (!IsLoaded || EditorBox is null || EditorScrollHost is null || EditorDocumentHost is null)
        {
            _ensureCaretVisibleAfterInteractionRefresh = false;
            _updateEditorDocumentHeightAfterInteractionRefresh = false;
            return;
        }

        var ensureCaretVisible = _ensureCaretVisibleAfterInteractionRefresh;
        var updateDocumentHeight = _updateEditorDocumentHeightAfterInteractionRefresh;
        _ensureCaretVisibleAfterInteractionRefresh = false;
        _updateEditorDocumentHeightAfterInteractionRefresh = false;

        if (updateDocumentHeight)
        {
            UpdateEditorDocumentHeight();
        }

        ScheduleEditorCueOverlayInvalidation();

        if (!ensureCaretVisible)
        {
            return;
        }

        if (ShouldDeferCaretVisibilityUntilCueLayoutRefresh())
        {
            _ensureCaretVisibleAfterCueLayoutRefresh = true;
            return;
        }

        if (updateDocumentHeight)
        {
            Dispatcher.BeginInvoke(
                new Action(() => _ = EnsureCaretVisible()),
                DispatcherPriority.ContextIdle);
            return;
        }

        _ = EnsureCaretVisible();
    }

    private void ScheduleEditorViewportRefresh(bool ensureCaretVisible)
    {
        if (!IsLoaded)
        {
            return;
        }

        _ensureCaretVisibleAfterRefresh |= ensureCaretVisible;
        if (_editorViewportRefreshPending)
        {
            return;
        }

        _editorViewportRefreshPending = true;
        Dispatcher.BeginInvoke(ApplyEditorViewportRefresh, System.Windows.Threading.DispatcherPriority.Background);
    }

    private void ApplyEditorViewportRefresh()
    {
        _editorViewportRefreshPending = false;

        if (!IsLoaded || EditorBox is null || EditorScrollHost is null || EditorDocumentHost is null)
        {
            _ensureCaretVisibleAfterRefresh = false;
            return;
        }

        if (IsScreenplayMode)
        {
            EnsureEditorCueOverlay();
        }

        ApplyEditorDocumentLayout();
        UpdateEditorDocumentHeight();
        InvalidateEditorCueOverlay();

        var ensureCaretVisible = _ensureCaretVisibleAfterRefresh;
        _ensureCaretVisibleAfterRefresh = false;

        Dispatcher.BeginInvoke(
            new Action(() => FinalizeEditorViewportRefresh(ensureCaretVisible)),
            System.Windows.Threading.DispatcherPriority.ContextIdle);
    }

    private void FinalizeEditorViewportRefresh(bool ensureCaretVisible)
    {
        if (!IsLoaded || EditorBox is null || EditorScrollHost is null || EditorDocumentHost is null)
        {
            return;
        }

        ApplyEditorDocumentLayout();
        UpdateEditorDocumentHeight();

        RestoreEditorViewportAnchorAfterZoomChange();
        if (ensureCaretVisible)
        {
            if (ShouldDeferCaretVisibilityUntilCueLayoutRefresh())
            {
                _ensureCaretVisibleAfterCueLayoutRefresh = true;
            }
            else
            {
                _ = EnsureCaretVisible();
            }
        }

        InvalidateEditorCueOverlay();
    }

    private bool ShouldDeferCaretVisibilityUntilCueLayoutRefresh()
    {
        return IsEditorCueOverlayVisible() &&
            _editorCueAdorner is not null &&
            _editorCueAdorner.HasPendingLayoutRefresh;
    }

    private void UpdateEditorDocumentHeight()
    {
        if (EditorDocumentHost is null)
        {
            return;
        }

        double desiredHeight;
        if (IsEditorCueOverlayVisible())
        {
            var documentHeight = _editorCueAdorner!.GetDocumentHeight();
            desiredHeight = EditorScrollHost is null
                ? documentHeight
                : Math.Max(documentHeight, EditorScrollHost.ViewportHeight);
        }
        else if (currentMode == WriteMode.Markdown)
        {
            desiredHeight = GetMarkdownDocumentHeight();
        }
        else
        {
            return;
        }

        if (double.IsNaN(EditorDocumentHost.Height) || Math.Abs(EditorDocumentHost.Height - desiredHeight) > 0.5)
        {
            EditorDocumentHost.Height = desiredHeight;
        }
    }

    private double GetMarkdownDocumentHeight()
    {
        if (EditorBox is null)
        {
            return 0.0;
        }

        var lineCount = RichTextBoxTextUtilities.GetLineCount(EditorBox);
        if (lineCount <= 0)
        {
            var text = GetEditorText();
            lineCount = 1;

            for (var index = 0; index < text.Length; index++)
            {
                if (text[index] == '\n')
                {
                    lineCount++;
                }
            }
        }

        var lineHeight = GetEditorWheelScrollStep();
        var pagePadding = EditorBox.Document?.PagePadding ?? new Thickness(0.0);
        var contentHeight =
            EditorBox.Padding.Top +
            EditorBox.Padding.Bottom +
            pagePadding.Top +
            pagePadding.Bottom +
            (lineCount * lineHeight);
        var viewportHeight = EditorScrollHost?.ViewportHeight ?? 0.0;
        return Math.Max(contentHeight, viewportHeight);
    }

    private bool IsEditorCueOverlayVisible()
    {
        return _currentWorkspaceSurface == WorkspaceSurface.Editor &&
            _editorCueAdorner is not null &&
            _editorCueAdorner.Visibility == Visibility.Visible;
    }

    private bool EnsureCaretVisible(double margin = 20.0)
    {
        if (!IsEditorCueOverlayVisible() || _editorCueAdorner is null || EditorScrollHost is null || EditorBox is null)
        {
            return false;
        }

        if (!_editorCueAdorner.TryGetCaretRect(out var caretRect) ||
            !TryGetCaretRectInScrollHost(caretRect, out var caretBounds) ||
            !TryGetEditorViewportBounds(out var viewportBounds))
        {
            return false;
        }

        var targetHorizontalOffset = EditorScrollHost.HorizontalOffset;
        var targetVerticalOffset = EditorScrollHost.VerticalOffset;
        var adjusted = false;

        if (caretBounds.Left < viewportBounds.Left + margin)
        {
            targetHorizontalOffset -= (viewportBounds.Left + margin) - caretBounds.Left;
            adjusted = true;
        }
        else if (caretBounds.Right > viewportBounds.Right - margin)
        {
            targetHorizontalOffset += caretBounds.Right - (viewportBounds.Right - margin);
            adjusted = true;
        }

        if (caretBounds.Top < viewportBounds.Top + margin)
        {
            targetVerticalOffset -= (viewportBounds.Top + margin) - caretBounds.Top;
            adjusted = true;
        }
        else if (caretBounds.Bottom > viewportBounds.Bottom - margin)
        {
            targetVerticalOffset += caretBounds.Bottom - (viewportBounds.Bottom - margin);
            adjusted = true;
        }

        if (!adjusted)
        {
            return false;
        }

        ScrollEditorViewportTo(targetHorizontalOffset, targetVerticalOffset);
        return true;
    }

    private bool TryGetEditorViewportBounds(out Rect viewportBounds)
    {
        viewportBounds = Rect.Empty;
        if (EditorScrollHost is null || EditorScrollHost.ViewportHeight <= 0 || EditorScrollHost.ViewportWidth <= 0)
        {
            return false;
        }

        viewportBounds = new Rect(0.0, 0.0, EditorScrollHost.ViewportWidth, EditorScrollHost.ViewportHeight);

        return true;
    }

    private bool TryGetCaretRectInScrollHost(Rect caretRect, out Rect caretBounds)
    {
        caretBounds = Rect.Empty;

        if (EditorBox is null || EditorScrollHost is null)
        {
            return false;
        }

        try
        {
            var transform = EditorBox.TransformToVisual(EditorScrollHost);
            caretBounds = transform.TransformBounds(caretRect);
            return true;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }

    private void ScrollEditorViewportTo(double horizontalOffset, double verticalOffset)
    {
        if (EditorScrollHost is null)
        {
            return;
        }

        var targetHorizontalOffset = Math.Max(0.0, Math.Min(horizontalOffset, EditorScrollHost.ScrollableWidth));
        var targetVerticalOffset = Math.Max(0.0, Math.Min(verticalOffset, EditorScrollHost.ScrollableHeight));

        if (Math.Abs(EditorScrollHost.HorizontalOffset - targetHorizontalOffset) > 0.5)
        {
            EditorScrollHost.ScrollToHorizontalOffset(targetHorizontalOffset);
        }

        if (Math.Abs(EditorScrollHost.VerticalOffset - targetVerticalOffset) > 0.5)
        {
            EditorScrollHost.ScrollToVerticalOffset(targetVerticalOffset);
        }
    }

    private void SetLeftDockCollapsed(bool collapsed)
    {
        _isLeftDockCollapsed = collapsed;

        if (LeftDockColumn is not null)
        {
            LeftDockColumn.Width = collapsed ? new GridLength(52) : new GridLength(LeftDockExpandedWidth);
        }

        if (LeftSplitter is not null)
        {
            LeftSplitter.Width = collapsed ? 0 : 6;
        }

        if (LeftDockTabs is not null)
        {
            LeftDockTabs.Visibility = collapsed ? Visibility.Collapsed : Visibility.Visible;
        }

        if (LeftDockTitleText is not null)
        {
            LeftDockTitleText.Visibility = collapsed ? Visibility.Collapsed : Visibility.Visible;
        }

        if (LeftDockToggleButton is not null)
        {
            LeftDockToggleButton.Content = collapsed ? ">" : "<";
            LeftDockToggleButton.ToolTip = collapsed ? "Expand left dock" : "Collapse left dock";
        }
    }

    private void SetSyntaxQuickReferenceVisible(bool visible)
    {
        if (!visible && SyntaxQuickReferenceColumn is not null)
        {
            var currentWidth = SyntaxQuickReferenceColumn.ActualWidth;
            if (currentWidth > 0.0)
            {
                _syntaxQuickReferenceWidth = Math.Max(SyntaxQuickReferenceMinimumWidth, currentWidth);
            }
        }

        _isSyntaxQuickReferenceVisible = visible;

        if (SyntaxQuickReferenceColumn is not null)
        {
            SyntaxQuickReferenceColumn.Width = visible
                ? new GridLength(Math.Max(SyntaxQuickReferenceMinimumWidth, _syntaxQuickReferenceWidth))
                : new GridLength(0.0);
        }

        if (SyntaxQuickReferenceSplitterColumn is not null)
        {
            SyntaxQuickReferenceSplitterColumn.Width = visible ? new GridLength(6.0) : new GridLength(0.0);
        }

        if (SyntaxQuickReferenceSplitter is not null)
        {
            SyntaxQuickReferenceSplitter.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        }

        if (SyntaxQuickReferenceBorder is not null)
        {
            SyntaxQuickReferenceBorder.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        }

        if (SyntaxQuickReferenceToggleButton is not null)
        {
            SyntaxQuickReferenceToggleButton.Content = visible ? "Hide Syntax" : "Show Syntax";
            SyntaxQuickReferenceToggleButton.ToolTip = visible
                ? "Hide syntax quick reference"
                : "Show syntax quick reference";
        }
    }

    public void NavigateToLine(int lineNumber)
    {
        NavigateEditorToLine(lineNumber);
    }

    private void NavigateEditorToLine(int lineNumber)
    {
        if (!IsLoaded || EditorBox is null)
        {
            return;
        }

        var targetLineIndex = Math.Max(0, lineNumber - 1);
        var characterIndex = GetEditorCharacterIndexFromLineIndex(targetLineIndex);
        if (characterIndex < 0)
        {
            return;
        }

        try
        {
            _suppressCaretContextUpdates = true;
            EditorBox.Focus();
            EditorBox.SelectionChanged -= EditorBox_SelectionChanged;
            SetEditorSelection(characterIndex, characterIndex);
            ViewModel.UpdateCaretContext(lineNumber);
        }
        finally
        {
            EditorBox.SelectionChanged += EditorBox_SelectionChanged;
            _suppressCaretContextUpdates = false;
        }

        ScheduleEditorViewportRefresh(ensureCaretVisible: true);
    }

    private void ShowGoToLineDialog()
    {
        if (_goToLineDialog is null || !_goToLineDialog.IsLoaded)
        {
            _goToLineDialog = new GoToLineDialog
            {
                Owner = this
            };
            _goToLineDialog.RequestJumpToLine += (_, lineNumber) => NavigateToLine(lineNumber);
            _goToLineDialog.Closed += (_, _) => _goToLineDialog = null;
        }

        _goToLineDialog.Show();
        _goToLineDialog.Activate();
        _goToLineDialog.FocusLineBox();
    }

    private void ShowGoToSceneDialog()
    {
        if (_goToSceneDialog is null || !_goToSceneDialog.IsLoaded)
        {
            _goToSceneDialog = new GoToSceneDialog
            {
                Owner = this
            };
            _goToSceneDialog.RequestJumpToLine += (_, lineNumber) => NavigateToLine(lineNumber);
            _goToSceneDialog.Closed += (_, _) => _goToSceneDialog = null;
        }

        _goToSceneDialog.SetScenes(ViewModel.OutlineRoots);
        _goToSceneDialog.Show();
        _goToSceneDialog.Activate();
    }

    public bool FindNext(string searchText)
    {
        return FindText(searchText, forward: true);
    }

    public bool FindPrevious(string searchText)
    {
        return FindText(searchText, forward: false);
    }

    public bool ReplaceCurrent(string searchText, string replacementText)
    {
        if (EditorBox is null || string.IsNullOrWhiteSpace(searchText))
        {
            return false;
        }

        if (!SelectionMatchesSearch(searchText) && !FindNext(searchText))
        {
            return false;
        }

        var replaceStart = GetEditorSelectionStart();
        var replacement = replacementText ?? string.Empty;

        try
        {
            EditorBox.BeginChange();
            EditorBox.Selection.Text = replacement;
        }
        finally
        {
            EditorBox.EndChange();
        }

        SelectEditorRange(replaceStart + replacement.Length, 0);
        return true;
    }

    public int ReplaceAll(string searchText, string replacementText)
    {
        if (EditorBox is null || string.IsNullOrWhiteSpace(searchText))
        {
            return 0;
        }

        var source = GetEditorText();
        var replacement = replacementText ?? string.Empty;
        var comparison = StringComparison.OrdinalIgnoreCase;
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

            builder.Append(source, index, matchIndex - index);
            builder.Append(replacement);
            replacements++;
            index = matchIndex + searchText.Length;
        }

        if (replacements == 0)
        {
            return 0;
        }

        var mappedCaret = MapCaretAfterReplaceAll(source, searchText, replacement, GetEditorCaretIndex());

        try
        {
            EditorBox.BeginChange();
            EditorBox.SelectAll();
            EditorBox.Selection.Text = builder.ToString();
        }
        finally
        {
            EditorBox.EndChange();
        }

        SelectEditorRange(mappedCaret, 0);
        return replacements;
    }

    private void ShowFindReplaceDialog()
    {
        if (_findReplaceDialog is null || !_findReplaceDialog.IsLoaded)
        {
            _findReplaceDialog = new FindReplaceDialog
            {
                Owner = this
            };
            _findReplaceDialog.Closed += (_, _) => _findReplaceDialog = null;
        }

        _findReplaceDialog.Show();
        _findReplaceDialog.Activate();
        _findReplaceDialog.FocusFindTextBox();
    }

    private bool FindText(string searchText, bool forward)
    {
        if (EditorBox is null || string.IsNullOrWhiteSpace(searchText))
        {
            return false;
        }

        var documentText = GetEditorText();
        var comparison = StringComparison.OrdinalIgnoreCase;

        int matchIndex;
        if (forward)
        {
            var startIndex = GetEditorSelectionStart() + GetEditorSelectionLength();
            if (startIndex > documentText.Length)
            {
                startIndex = documentText.Length;
            }

            matchIndex = documentText.IndexOf(searchText, startIndex, comparison);
            if (matchIndex < 0)
            {
                matchIndex = documentText.IndexOf(searchText, 0, comparison);
            }
        }
        else
        {
            var startIndex = Math.Max(0, GetEditorSelectionStart() - 1);
            matchIndex = documentText.LastIndexOf(searchText, startIndex, comparison);
            if (matchIndex < 0)
            {
                matchIndex = documentText.LastIndexOf(searchText, documentText.Length - 1, comparison);
            }
        }

        if (matchIndex < 0)
        {
            return false;
        }

        SelectEditorRange(matchIndex, searchText.Length);
        return true;
    }

    private void SelectEditorRange(int start, int length)
    {
        if (EditorBox is null)
        {
            return;
        }

        var textLength = GetEditorText().Length;
        var safeStart = Math.Max(0, Math.Min(start, textLength));
        var safeLength = Math.Max(0, Math.Min(length, textLength - safeStart));

        EditorBox.Focus();
        SetEditorSelection(safeStart, safeStart + safeLength);
        ScheduleEditorViewportRefresh(ensureCaretVisible: true);
    }

    private bool SelectionMatchesSearch(string searchText)
    {
        if (EditorBox is null)
        {
            return false;
        }

        return GetEditorSelectionLength() == searchText.Length
            && string.Equals(RichTextBoxTextUtilities.GetSelectedText(EditorBox), searchText, StringComparison.OrdinalIgnoreCase);
    }

    private static int MapCaretAfterReplaceAll(string source, string searchText, string replacementText, int caretIndex)
    {
        var comparison = StringComparison.OrdinalIgnoreCase;
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

    private void RestoreOutlineSelection(int lineNumber)
    {
        _suppressOutlineSelectionNavigation = true;
        try
        {
            ClearTreeSelection(OutlineTree);
            ClearTreeSelection(NotesTree);

            var activeNode = ViewModel.SelectedDocument?.FindActiveOutlineNode(lineNumber);
            if (activeNode != null)
            {
                if (OutlineTree is not null && TrySelectSpecificNode(OutlineTree, activeNode))
                {
                    return;
                }
            }

            if (NotesTree is not null)
            {
                _ = TrySelectOutlineNode(NotesTree, lineNumber);
            }
        }
        finally
        {
            _suppressOutlineSelectionNavigation = false;
        }
    }

    private bool TrySelectSpecificNode(System.Windows.Controls.ItemsControl parent, OutlineNodeViewModel targetNode)
    {
        foreach (var item in parent.Items)
        {
            if (item is not OutlineNodeViewModel node)
            {
                continue;
            }

            var container = parent.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.TreeViewItem;
            if (container is null)
            {
                continue;
            }

            if (ReferenceEquals(node, targetNode))
            {
                if (!container.IsSelected)
                {
                    container.IsSelected = true;
                }
                container.BringIntoView();
                return true;
            }

            if (TrySelectSpecificNode(container, targetNode))
            {
                if (!container.IsExpanded)
                {
                    container.IsExpanded = true;
                }
                return true;
            }
        }

        return false;
    }

    private bool TrySelectOutlineNode(System.Windows.Controls.ItemsControl parent, int lineNumber)
    {
        foreach (var item in parent.Items)
        {
            if (item is not OutlineNodeViewModel node)
            {
                continue;
            }

            var container = parent.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.TreeViewItem;
            if (container is null)
            {
                continue;
            }

            if (node.LineNumber == lineNumber)
            {
                container.IsSelected = true;
                container.BringIntoView();
                return true;
            }

            if (TrySelectOutlineNode(container, lineNumber))
            {
                if (!container.IsExpanded)
                {
                    container.IsExpanded = true;
                }
                return true;
            }
        }

        return false;
    }

    private static void ClearTreeSelection(System.Windows.Controls.ItemsControl? parent)
    {
        if (parent is null)
        {
            return;
        }

        foreach (var item in parent.Items)
        {
            var container = parent.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.TreeViewItem;
            if (container is null)
            {
                continue;
            }

            if (container.IsSelected)
            {
                container.IsSelected = false;
            }

            ClearTreeSelection(container);
        }
    }

    private static bool IsCtrlShortcut(System.Windows.Input.KeyEventArgs e)
    {
        return System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Control)
            && !System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Alt)
            && e.Key is System.Windows.Input.Key.N
                or System.Windows.Input.Key.F
                or System.Windows.Input.Key.G
                or System.Windows.Input.Key.M
                or System.Windows.Input.Key.O
                or System.Windows.Input.Key.Add
                or System.Windows.Input.Key.Subtract
                or System.Windows.Input.Key.OemPlus
                or System.Windows.Input.Key.OemMinus
                or System.Windows.Input.Key.D0
                or System.Windows.Input.Key.NumPad0
                or System.Windows.Input.Key.S
                or System.Windows.Input.Key.W
                or System.Windows.Input.Key.Y
                or System.Windows.Input.Key.Z;
    }

    private static bool HasShiftModifier()
    {
        return System.Windows.Input.Keyboard.Modifiers.HasFlag(System.Windows.Input.ModifierKeys.Shift);
    }
}
