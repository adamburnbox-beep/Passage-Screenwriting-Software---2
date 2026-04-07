using Passage.Core.Goals;

namespace Passage.App.Services;

public sealed record SessionGoalConfiguration
{
    public SessionGoalConfiguration()
        : this(GoalType.WordCount, 1000, 120, 25)
    {
    }

    public SessionGoalConfiguration(
        GoalType selectedGoalType,
        int wordCountTargetValue,
        int pageCountTargetValue,
        int timerTargetMinutes)
    {
        if (wordCountTargetValue < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(wordCountTargetValue));
        }

        if (pageCountTargetValue < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(pageCountTargetValue));
        }

        if (timerTargetMinutes < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(timerTargetMinutes));
        }

        SelectedGoalType = selectedGoalType;
        WordCountTargetValue = wordCountTargetValue;
        PageCountTargetValue = pageCountTargetValue;
        TimerTargetMinutes = timerTargetMinutes;
    }

    public GoalType SelectedGoalType { get; init; }

    public int WordCountTargetValue { get; init; }

    public int PageCountTargetValue { get; init; }

    public int TimerTargetMinutes { get; init; }
}
