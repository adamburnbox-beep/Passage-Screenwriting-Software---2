using System.Windows;
using Passage.App.ViewModels;

namespace Passage.App.Views;

public partial class TitlePageDialog : Window
{
    public bool Deleted { get; private set; }

    public TitlePageDialog()
    {
        InitializeComponent();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("Are you sure you want to delete the title page?", "Delete Title Page", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
        {
            Deleted = true;
            DialogResult = true;
            Close();
        }
    }

    private void FieldsScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
    {
        var scrollViewer = sender as System.Windows.Controls.ScrollViewer;
        if (scrollViewer != null)
        {
            // Divide the delta to smooth out trackpad scrolling
            scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta / 3.0);
            e.Handled = true;
        }
    }
}
