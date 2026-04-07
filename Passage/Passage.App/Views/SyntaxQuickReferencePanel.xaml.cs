using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

namespace Passage.App.Views;

public partial class SyntaxQuickReferencePanel : UserControl
{
    public SyntaxQuickReferencePanel()
    {
        InitializeComponent();
    }

    private void CopySyntax_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string syntax } ||
            string.IsNullOrWhiteSpace(syntax))
        {
            return;
        }

        try
        {
            Clipboard.SetText(syntax);
        }
        catch (ExternalException)
        {
            // Ignore clipboard contention so the panel never interrupts writing.
        }
    }
}
