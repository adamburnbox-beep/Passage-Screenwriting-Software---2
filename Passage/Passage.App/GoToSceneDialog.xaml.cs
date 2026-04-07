using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Passage.App.ViewModels;

namespace Passage.App;

public partial class GoToSceneDialog : Window
{
    private readonly ObservableCollection<SceneJumpItem> _scenes = new();

    public event EventHandler<int>? RequestJumpToLine;

    public GoToSceneDialog()
    {
        InitializeComponent();
        SceneListBox.ItemsSource = _scenes;
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

    private void Jump_Click(object sender, RoutedEventArgs e)
    {
        JumpToSelection();
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void SceneListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        JumpToSelection();
    }

    private void SceneListBox_PreviewKeyDown(object sender, KeyEventArgs e)
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

        RequestJumpToLine?.Invoke(this, selected.LineNumber);
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

    private sealed record SceneJumpItem(int LineNumber, string Text)
    {
        public string DisplayText => $"{LineNumber}: {Text}";
    }
}
