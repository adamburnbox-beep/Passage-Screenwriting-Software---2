using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public partial class FindReplaceDialog : Window
{
    public string SearchText => FindTextBox.Text ?? string.Empty;
    public string ReplaceText => ReplaceTextBox.Text ?? string.Empty;
    public bool MatchCase => MatchCaseCheckBox.IsChecked ?? false;
    public bool WholeWord => WholeWordCheckBox.IsChecked ?? false;

    public event System.Action? FindNextRequested;
    public event System.Action? ReplaceRequested;
    public event System.Action? ReplaceAllRequested;

    public FindReplaceDialog()
    {
        InitializeComponent();
    }

    private void FindNextClick(object? sender, RoutedEventArgs e)
    {
        FindNextRequested?.Invoke();
    }

    private void ReplaceClick(object? sender, RoutedEventArgs e)
    {
        ReplaceRequested?.Invoke();
    }

    private void ReplaceAllClick(object? sender, RoutedEventArgs e)
    {
        ReplaceAllRequested?.Invoke();
    }

    private void CloseClick(object? sender, RoutedEventArgs e)
    {
        Close();
    }
}
