using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public enum UnsavedChangesChoice
{
    Save,
    Discard,
    Cancel
}

public partial class UnsavedChangesDialog : Window
{
    public UnsavedChangesDialog()
    {
        InitializeComponent();
    }

    private void SaveClick(object? sender, RoutedEventArgs e) => Close(UnsavedChangesChoice.Save);

    private void DiscardClick(object? sender, RoutedEventArgs e) => Close(UnsavedChangesChoice.Discard);

    private void CancelClick(object? sender, RoutedEventArgs e) => Close(UnsavedChangesChoice.Cancel);
}
