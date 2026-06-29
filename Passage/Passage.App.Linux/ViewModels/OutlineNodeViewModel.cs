using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace Passage.App.ViewModels;

public enum OutlineNodeKind
{
    Section,
    SceneHeading,
    Note
}

public enum WorkspaceDropPosition
{
    Above,
    Below,
    Onto
}

public partial class OutlineNodeViewModel : ViewModelBase
{
    [ObservableProperty]
    private bool _isExpanded;

    [ObservableProperty]
    private bool _isDragOver;

    // True when the editor caret currently sits within this element's range.
    // Drives the workspace "current element" highlight.
    [ObservableProperty]
    private bool _isCurrent;

    [ObservableProperty]
    private int _level;

    public OutlineNodeViewModel(
        OutlineNodeKind kind,
        string text,
        int lineNumber,
        int? sectionLevel = null,
        string? bodyText = null,
        Action<int>? navigateAction = null,
        string? sceneNumber = null,
        int level = 0)
    {
        Kind = kind;
        Text = kind == OutlineNodeKind.SceneHeading ? (text ?? string.Empty).ToUpperInvariant() : text;
        LineNumber = lineNumber;
        SectionLevel = sectionLevel;
        BodyText = (bodyText ?? string.Empty).ReplaceLineEndings("\n").Trim();
        SceneNumber = sceneNumber;
        Level = level;
        Children = new ObservableCollection<OutlineNodeViewModel>();
        NavigateCommand = new RelayCommand(() => navigateAction?.Invoke(LineNumber));
    }

    public OutlineNodeKind Kind { get; }

    public string KindLabel => Kind switch
    {
        OutlineNodeKind.Section => SectionLevel switch
        {
            1 => "Act",
            2 => "Sequence",
            _ => "Beat"
        },
        OutlineNodeKind.SceneHeading => "Scene",
        OutlineNodeKind.Note => "Note",
        _ => Kind.ToString()
    };

    public bool IsActLevel => Kind == OutlineNodeKind.Section && SectionLevel == 1;

    public string Text { get; }

    public string Title => Text;

    public string DisplayText => Text.ReplaceLineEndings(" ").Trim();

    public string BodyText { get; }

    public string Synopsis => BodyText;

    public string? SceneNumber { get; set; }

    public bool HasBodyText => BodyText.Length > 0;

    public ICommand NavigateCommand { get; }

    public string ToolTipText => HasBodyText
        ? $"{Text}{Environment.NewLine}{BodyText}"
        : Text;

    public int LineNumber { get; }

    public int? SectionLevel { get; }

    public ObservableCollection<OutlineNodeViewModel> Children { get; }
}
