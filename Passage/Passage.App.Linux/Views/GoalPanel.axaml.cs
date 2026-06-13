using System;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Passage.App.ViewModels;

namespace Passage.App.Views;

public partial class GoalPanel : UserControl
{
    public GoalPanel()
    {
        InitializeComponent();
    }

    private void GoalPrimaryTimerButton_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel)
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

    private void GoalSecondaryTimerButton_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel)
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

    private void GoalResetSession_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.ResetSessionGoal();
        }
    }
}
