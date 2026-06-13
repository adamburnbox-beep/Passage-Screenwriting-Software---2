using System;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;

namespace Passage.App.Views;

public partial class SyntaxQuickReferencePanel : UserControl
{
    public SyntaxQuickReferencePanel()
    {
        InitializeComponent();
    }

    private async void CopySyntax_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string syntax } ||
            string.IsNullOrWhiteSpace(syntax))
        {
            return;
        }

        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard != null)
            {
                await clipboard.SetValueAsync(DataFormat.Text, syntax);
            }
        }
        catch (Exception)
        {
            // Ignore clipboard contention so the panel never interrupts writing.
        }
    }
}

