using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Passage.App.Utilities;

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

public sealed class OutlineNodeViewModel : INotifyPropertyChanged
{
    private bool _isExpanded;
    private bool _isDragOver;
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
        Text = text;
        LineNumber = lineNumber;
        SectionLevel = sectionLevel;
        BodyText = (bodyText ?? string.Empty).ReplaceLineEndings("\n").Trim();
        SceneNumber = sceneNumber;
        Level = level;
        Children = new ObservableCollection<OutlineNodeViewModel>();
        NavigateCommand = new DelegateCommand<object>(_ => navigateAction?.Invoke(LineNumber));
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

    public bool IsExpanded
    {
        get => _isExpanded;
        set => SetProperty(ref _isExpanded, value);
    }

    public bool IsDragOver
    {
        get => _isDragOver;
        set => SetProperty(ref _isDragOver, value);
    }

    public int Level
    {
        get => _level;
        set => SetProperty(ref _level, value);
    }

    public ObservableCollection<OutlineNodeViewModel> Children { get; }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(storage, value))
        {
            return;
        }

        storage = value;
        OnPropertyChanged(propertyName);
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
