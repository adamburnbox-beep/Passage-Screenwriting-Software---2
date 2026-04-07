using System.Windows;
using System.Windows.Input;

namespace Passage.App;

public partial class GoToLineDialog : Window
{
    public event EventHandler<int>? RequestJumpToLine;

    public GoToLineDialog()
    {
        InitializeComponent();
        Loaded += (_, _) => FocusLineBox();
    }

    public void FocusLineBox()
    {
        LineTextBox.Focus();
        LineTextBox.SelectAll();
    }

    private void Go_Click(object sender, RoutedEventArgs e)
    {
        if (TryGetLineNumber(out var lineNumber))
        {
            RequestJumpToLine?.Invoke(this, lineNumber);
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void LineTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        e.Handled = true;
        Go_Click(sender, new RoutedEventArgs());
    }

    private bool TryGetLineNumber(out int lineNumber)
    {
        lineNumber = 0;
        if (!int.TryParse(LineTextBox.Text, out lineNumber))
        {
            return false;
        }

        return lineNumber > 0;
    }
}
