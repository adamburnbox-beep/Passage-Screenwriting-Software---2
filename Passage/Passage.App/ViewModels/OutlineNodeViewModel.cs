using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

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

    public OutlineNodeViewModel(
        OutlineNodeKind kind,
        string text,
        int lineNumber,
        int? sectionLevel = null,
        string? bodyText = null)
    {
        Kind = kind;
        Text = text;
        LineNumber = lineNumber;
        SectionLevel = sectionLevel;
        BodyText = (bodyText ?? string.Empty).ReplaceLineEndings("\n").Trim();
        Children = new ObservableCollection<OutlineNodeViewModel>();
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

    public string DisplayText => Text.ReplaceLineEndings(" ").Trim();

    public string BodyText { get; }

    public bool HasBodyText => BodyText.Length > 0;

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
