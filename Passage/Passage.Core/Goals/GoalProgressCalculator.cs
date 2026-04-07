using Passage.Core;

namespace Passage.Core.Goals;

public sealed class GoalProgressCalculator
{
    public int CalculateWordCount(string? documentText)
    {
        if (string.IsNullOrWhiteSpace(documentText))
        {
            return 0;
        }

        return TextAnalysis.CountWords(documentText.AsSpan());
    }

    public Goal UpdateWordCountGoal(Goal goal, string? documentText)
    {
        ArgumentNullException.ThrowIfNull(goal);

        if (goal.Type != GoalType.WordCount)
        {
            throw new ArgumentException("Goal must be a word count goal.", nameof(goal));
        }

        var currentValue = CalculateWordCount(documentText);
        return new Goal(
            GoalType.WordCount,
            goal.TargetValue,
            currentValue,
            currentValue >= goal.TargetValue);
    }
}
