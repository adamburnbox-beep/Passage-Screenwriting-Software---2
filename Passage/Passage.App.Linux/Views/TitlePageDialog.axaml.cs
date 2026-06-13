using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public partial class TitlePageDialog : Window
{
    public bool Deleted { get; private set; }

    public TitlePageDialog()
    {
        InitializeComponent();
    }

    private void OkClick(object? sender, RoutedEventArgs e)
    {
        Close(true);
    }

    private void CancelClick(object? sender, RoutedEventArgs e)
    {
        Close(false);
    }

    private void DeleteClick(object? sender, RoutedEventArgs e)
    {
        Deleted = true;
        Close(true);
    }
}
