using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public partial class GoToLineDialog : Window
{
    public int LineNumber { get; private set; }

    public GoToLineDialog()
    {
        InitializeComponent();
    }

    private void GoClick(object? sender, RoutedEventArgs e)
    {
        if (int.TryParse(LineNumberTextBox.Text, out var lineNumber) && lineNumber > 0)
        {
            LineNumber = lineNumber;
            Close(true);
        }
    }

    private void CancelClick(object? sender, RoutedEventArgs e)
    {
        Close(false);
    }
}
