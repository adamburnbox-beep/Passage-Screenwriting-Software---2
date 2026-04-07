namespace Passage.Core.Goals;

public enum GoalType
{
    WordCount,
    PageCount,
    Timer
}

public enum TimerGoalState
{
    Idle,
    Running,
    Paused,
    Completed
}

public sealed class Goal
{
    public Goal(
        GoalType type,
        int targetValue,
        int currentValue,
        bool isCompleted = false,
        TimerGoalState? timerState = null)
    {
        if (targetValue < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(targetValue));
        }

        if (currentValue < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(currentValue));
        }

        if (type != GoalType.Timer && timerState is not null)
        {
            throw new ArgumentException("Timer state is only valid for timer goals.", nameof(timerState));
        }

        Type = type;
        TargetValue = targetValue;
        CurrentValue = currentValue;
        TimerState = type == GoalType.Timer ? timerState ?? TimerGoalState.Idle : null;
        IsCompleted = isCompleted || TimerState == TimerGoalState.Completed;
    }

    public GoalType Type { get; }

    public int TargetValue { get; }

    public int CurrentValue { get; }

    public bool IsCompleted { get; }

    public TimerGoalState? TimerState { get; }
}
