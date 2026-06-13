using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public partial class RecoveryPromptDialog : Window
{
    public RecoveryPromptDialog()
    {
        InitializeComponent();
    }

    private void RestoreClick(object? sender, RoutedEventArgs e)
    {
        Close(true);
    }

    private void DiscardClick(object? sender, RoutedEventArgs e)
    {
        Close(false);
    }
}
