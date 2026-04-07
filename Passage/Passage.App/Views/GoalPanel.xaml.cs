using System.Windows.Controls;
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
