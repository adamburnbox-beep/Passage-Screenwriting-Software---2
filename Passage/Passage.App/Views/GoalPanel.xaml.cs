using System;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using Passage.App.ViewModels;

namespace Passage.App.Views;

public partial class GoalPanel : UserControl
{
    public GoalPanel()
    {
        InitializeComponent();
    }

    private void GoalTargetValue_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
    {
        e.Handled = ContainsNonDigits(e.Text);
    }

    private void GoalTargetValue_Pasting(object sender, System.Windows.DataObjectPastingEventArgs e)
    {
        if (!e.SourceDataObject.GetDataPresent(System.Windows.DataFormats.UnicodeText) &&
            !e.SourceDataObject.GetDataPresent(System.Windows.DataFormats.Text))
        {
            e.CancelCommand();
            return;
        }

        var text = e.SourceDataObject.GetData(System.Windows.DataFormats.UnicodeText) as string
            ?? e.SourceDataObject.GetData(System.Windows.DataFormats.Text) as string
            ?? string.Empty;

        if (ContainsNonDigits(text))
        {
            e.CancelCommand();
        }
    }

    private void GoalPrimaryTimerButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is not ShellViewModel viewModel)
        {
            return;
        }

        switch (viewModel.GoalTimerState)
        {
            case Passage.Core.Goals.TimerGoalState.Running:
                viewModel.PauseGoalTimer();
                break;
            case Passage.Core.Goals.TimerGoalState.Paused:
                viewModel.ResumeGoalTimer();
                break;
            case Passage.Core.Goals.TimerGoalState.Completed:
            case Passage.Core.Goals.TimerGoalState.Idle:
            default:
                viewModel.StartGoalTimer();
                break;
        }
    }

    private void GoalSecondaryTimerButton_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is not ShellViewModel viewModel)
        {
            return;
        }

        switch (viewModel.GoalTimerState)
        {
            case Passage.Core.Goals.TimerGoalState.Running:
            case Passage.Core.Goals.TimerGoalState.Paused:
                viewModel.StopGoalTimer();
                break;
            case Passage.Core.Goals.TimerGoalState.Completed:
            case Passage.Core.Goals.TimerGoalState.Idle:
            default:
                viewModel.ResetGoalTimer();
                break;
        }
    }

    private void GoalReset_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is ShellViewModel viewModel)
        {
            viewModel.ResetGoalTimer();
        }
    }

    private void GoalResetSession_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is ShellViewModel viewModel)
        {
            viewModel.ResetSessionGoal();
        }
    }

    private static bool ContainsNonDigits(string text)
    {
        foreach (var ch in text)
        {
            if (!char.IsDigit(ch))
            {
                return true;
            }
        }

        return false;
    }
}

public class ProgressToArcConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        double percent = 0;
        if (value is double d) percent = d;
        else if (value is float f) percent = f;
        else if (value is int i) percent = i;
        else if (value is decimal dec) percent = (double)dec;

        percent = Math.Clamp(percent, 0, 100);
        if (percent >= 100)
        {
            return Geometry.Parse("M 55,5 A 50,50 0 1 1 54.99,5 Z");
        }
        if (percent <= 0)
        {
            return Geometry.Parse("M 55,5");
        }

        double angle = (percent / 100.0) * 360.0;
        double rad = (angle - 90.0) * Math.PI / 180.0;
        double endX = 55.0 + 50.0 * Math.Cos(rad);
        double endY = 55.0 + 50.0 * Math.Sin(rad);

        int isLargeArc = percent > 50 ? 1 : 0;
        string pathData = string.Format(CultureInfo.InvariantCulture, "M 55,5 A 50,50 0 {0} 1 {1:F2},{2:F2}", isLargeArc, endX, endY);
        return Geometry.Parse(pathData);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
