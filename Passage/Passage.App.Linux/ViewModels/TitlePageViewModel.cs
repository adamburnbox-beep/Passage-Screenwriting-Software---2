using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace Passage.App.ViewModels;

public partial class TitlePageViewModel : ObservableObject
{
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(UppercaseTitle))]
    private string _title = string.Empty;

    [ObservableProperty]
    private string _episode = string.Empty;

    [ObservableProperty]
    private string _credit = "written by";

    [ObservableProperty]
    private string _author = string.Empty;

    [ObservableProperty]
    private string _source = string.Empty;

    [ObservableProperty]
    private string _contact = string.Empty;

    [ObservableProperty]
    private string _draftDate = string.Empty;

    [ObservableProperty]
    private string _revision = string.Empty;

    [ObservableProperty]
    private string _notes = string.Empty;

    public string UppercaseTitle => Title?.ToUpperInvariant() ?? string.Empty;
}
