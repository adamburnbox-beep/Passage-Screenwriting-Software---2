using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Passage.App.Utilities;

namespace Passage.App.ViewModels;

public class TitlePageEntryViewModel : INotifyPropertyChanged
{
    private string _label = string.Empty;
    private string _value = string.Empty;

    public string Label { get => _label; set { _label = value; OnPropertyChanged(); } }
    public string Value { get => _value; set { _value = value; OnPropertyChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class TitlePageViewModel : INotifyPropertyChanged
{
    private bool _showTitlePage = true;
    private bool _contactAlignLeft = true;
    private string _title = string.Empty;
    private string _episode = string.Empty;
    private string _credit = "written by";
    private string _author = string.Empty;
    private string _source = string.Empty;
    private string _contact = string.Empty;
    private string _draftDate = string.Empty;
    private string _revision = string.Empty;
    private string _notes = string.Empty;

    public bool ShowTitlePage { get => _showTitlePage; set { _showTitlePage = value; OnPropertyChanged(); } }
    public bool ContactAlignLeft { get => _contactAlignLeft; set { _contactAlignLeft = value; OnPropertyChanged(); OnPropertyChanged(nameof(ContactAlignRight)); } }
    public bool ContactAlignRight { get => !_contactAlignLeft; set { _contactAlignLeft = !value; OnPropertyChanged(); OnPropertyChanged(nameof(ContactAlignLeft)); } }

    public string Title { get => _title; set { _title = value; OnPropertyChanged(); } }
    public string Episode { get => _episode; set { _episode = value; OnPropertyChanged(); } }
    public string Credit { get => _credit; set { _credit = value; OnPropertyChanged(); } }
    public string Author { get => _author; set { _author = value; OnPropertyChanged(); } }
    public string Source { get => _source; set { _source = value; OnPropertyChanged(); } }
    public string Contact { get => _contact; set { _contact = value; OnPropertyChanged(); } }
    public string DraftDate { get => _draftDate; set { _draftDate = value; OnPropertyChanged(); } }
    public string Revision { get => _revision; set { _revision = value; OnPropertyChanged(); } }
    public string Notes { get => _notes; set { _notes = value; OnPropertyChanged(); } }

    public ObservableCollection<TitlePageEntryViewModel> CustomEntries { get; } = new ObservableCollection<TitlePageEntryViewModel>();

    public ICommand AddCustomEntryCommand { get; }
    public ICommand RemoveCustomEntryCommand { get; }

    public TitlePageViewModel()
    {
        AddCustomEntryCommand = new DelegateCommand<object>(execute: _ => CustomEntries.Add(new TitlePageEntryViewModel { Label = "New Field", Value = string.Empty }));
        RemoveCustomEntryCommand = new DelegateCommand<TitlePageEntryViewModel>(execute: e => { if (e != null) CustomEntries.Remove(e); });
    }

    public void CopyFrom(TitlePageViewModel other)
    {
        ShowTitlePage = other.ShowTitlePage;
        ContactAlignLeft = other.ContactAlignLeft;
        Title = other.Title;
        Episode = other.Episode;
        Credit = other.Credit;
        Author = other.Author;
        Source = other.Source;
        Contact = other.Contact;
        DraftDate = other.DraftDate;
        Revision = other.Revision;
        Notes = other.Notes;
        
        CustomEntries.Clear();
        foreach (var entry in other.CustomEntries)
        {
            CustomEntries.Add(new TitlePageEntryViewModel { Label = entry.Label, Value = entry.Value });
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
