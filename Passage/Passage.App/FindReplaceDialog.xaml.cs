using System.Windows;
using System.Windows.Input;

namespace Passage.App;

public partial class FindReplaceDialog : Window
{
    public FindReplaceDialog()
    {
        InitializeComponent();
        Loaded += (_, _) => FocusFindTextBox();
    }

    public void FocusFindTextBox()
    {
        FindTextBox.Focus();
        FindTextBox.SelectAll();
    }

    private MainWindow? OwnerWindow => Owner as MainWindow;

    private void FindTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        StatusTextBlock.Text = "Ready";
    }

    private void FindNext_Click(object sender, RoutedEventArgs e)
    {
        if (OwnerWindow is null)
        {
            return;
        }

        SetStatus(OwnerWindow.FindNext(FindTextBox.Text) ? "Match found." : "No match found.");
    }

    private void FindPrevious_Click(object sender, RoutedEventArgs e)
    {
        if (OwnerWindow is null)
        {
            return;
        }

        SetStatus(OwnerWindow.FindPrevious(FindTextBox.Text) ? "Match found." : "No match found.");
    }

    private void Replace_Click(object sender, RoutedEventArgs e)
    {
        if (OwnerWindow is null)
        {
            return;
        }

        SetStatus(OwnerWindow.ReplaceCurrent(FindTextBox.Text, ReplaceTextBox.Text)
            ? "Replaced current match."
            : "No match to replace.");
    }

    private void ReplaceAll_Click(object sender, RoutedEventArgs e)
    {
        if (OwnerWindow is null)
        {
            return;
        }

        var count = OwnerWindow.ReplaceAll(FindTextBox.Text, ReplaceTextBox.Text);
        SetStatus(count > 0 ? $"Replaced {count} occurrence(s)." : "No matches found.");
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
    }
}
