using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Passage.App.Views;

/// <summary>
/// Small reusable OK/Cancel confirmation dialog. Returns true from ShowDialog
/// when the confirm button is pressed.
/// </summary>
public partial class ConfirmDialog : Window
{
    public ConfirmDialog()
    {
        InitializeComponent();
    }

    public ConfirmDialog(string title, string message, string confirmText = "OK") : this()
    {
        Title = title;
        var titleText = this.FindControl<TextBlock>("TitleText");
        var messageText = this.FindControl<TextBlock>("MessageText");
        var confirmButton = this.FindControl<Button>("ConfirmButton");
        if (titleText != null) titleText.Text = title;
        if (messageText != null) messageText.Text = message;
        if (confirmButton != null) confirmButton.Content = confirmText;
    }

    private void ConfirmClick(object? sender, RoutedEventArgs e) => Close(true);

    private void CancelClick(object? sender, RoutedEventArgs e) => Close(false);
}
