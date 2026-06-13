using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Passage.App.ViewModels;

namespace Passage.App.Views;

public partial class GoToSceneDialog : Window
{
    private readonly ObservableCollection<SceneJumpItem> _scenes = new();

    public GoToSceneDialog()
    {
        InitializeComponent();
        SceneListBox.ItemsSource = _scenes;
    }

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        SceneListBox.Focus();
    }

    public void SetScenes(IEnumerable<OutlineNodeViewModel> outlineRoots)
    {
        _scenes.Clear();
        foreach (var scene in FlattenScenes(outlineRoots))
        {
            _scenes.Add(scene);
        }

        if (_scenes.Count > 0 && SceneListBox.SelectedItem is null)
        {
            SceneListBox.SelectedIndex = 0;
        }
    }

    private void JumpClick(object? sender, RoutedEventArgs e)
    {
        JumpToSelection();
    }

    private void CloseClick(object? sender, RoutedEventArgs e)
    {
        Close(null);
    }

    private void SceneListBox_DoubleTapped(object? sender, TappedEventArgs e)
    {
        JumpToSelection();
    }

    private void SceneListBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        e.Handled = true;
        JumpToSelection();
    }

    private void JumpToSelection()
    {
        if (SceneListBox.SelectedItem is not SceneJumpItem selected)
        {
            return;
        }

        Close(selected.LineNumber);
    }

    private static IEnumerable<SceneJumpItem> FlattenScenes(IEnumerable<OutlineNodeViewModel> nodes)
    {
        foreach (var node in nodes)
        {
            if (node.Kind == OutlineNodeKind.SceneHeading)
            {
                yield return new SceneJumpItem(node.LineNumber, node.Text);
            }

            foreach (var child in FlattenScenes(node.Children))
            {
                yield return child;
            }
        }
    }

}

public sealed record SceneJumpItem(int LineNumber, string Text)
{
    public string DisplayText => $"{LineNumber}: {Text}";
}
