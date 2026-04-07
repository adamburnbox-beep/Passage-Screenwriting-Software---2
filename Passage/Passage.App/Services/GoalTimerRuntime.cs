using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Threading;
using Passage.Core.Goals;

namespace Passage.App.Services;

public sealed class GoalTimerRuntime : INotifyPropertyChanged, IDisposable
{
    private const int TickIntervalMilliseconds = 250;

    private readonly DispatcherTimer _timer;
    private readonly EventHandler _tickHandler;
    private readonly Stopwatch _stopwatch = new();
    private TimeSpan _elapsedTime;
    private TimeSpan _remainingTime;
    private TimerGoalState _state;

    public GoalTimerRuntime(TimeSpan targetDuration)
    {
        if (targetDuration < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(targetDuration));
        }

        TargetDuration = targetDuration;
        _remainingTime = targetDuration;
        _state = targetDuration == TimeSpan.Zero ? TimerGoalState.Completed : TimerGoalState.Idle;

        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(TickIntervalMilliseconds)
        };
        _tickHandler = (_, _) => UpdateProgress();
        _timer.Tick += _tickHandler;
    }

    public TimeSpan TargetDuration { get; }

    public TimeSpan ElapsedTime
    {
        get => _elapsedTime;
        private set => SetProperty(ref _elapsedTime, value);
    }

    public TimeSpan RemainingTime
    {
        get => _remainingTime;
        private set => SetProperty(ref _remainingTime, value);
    }

    public TimerGoalState State
    {
        get => _state;
        private set
        {
            if (SetProperty(ref _state, value))
            {
                OnPropertyChanged(nameof(IsCompleted));
                OnPropertyChanged(nameof(Snapshot));
            }
        }
    }

    public bool IsCompleted => State == TimerGoalState.Completed;

    public Goal Snapshot => new(
        GoalType.Timer,
        Math.Max(0, (int)Math.Ceiling(TargetDuration.TotalSeconds)),
        Math.Max(0, (int)Math.Ceiling(ElapsedTime.TotalSeconds)),
        IsCompleted,
        State);

    public event PropertyChangedEventHandler? PropertyChanged;

    public void Start()
    {
        if (State == TimerGoalState.Completed)
        {
            return;
        }

        if (State == TimerGoalState.Running)
        {
            return;
        }

        if (TargetDuration == TimeSpan.Zero)
        {
            CompleteTimer();
            return;
        }

        _stopwatch.Restart();
        _timer.Start();
        State = TimerGoalState.Running;
        OnPropertyChanged(nameof(Snapshot));
    }

    public void Pause()
    {
        if (State != TimerGoalState.Running)
        {
            return;
        }

        CaptureRunningProgress();
        if (State == TimerGoalState.Completed)
        {
            return;
        }

        _stopwatch.Reset();
        _timer.Stop();
        State = TimerGoalState.Paused;
        OnPropertyChanged(nameof(Snapshot));
    }

    public void Resume()
    {
        if (State != TimerGoalState.Paused)
        {
            return;
        }

        if (RemainingTime <= TimeSpan.Zero)
        {
            CompleteTimer();
            return;
        }

        _stopwatch.Restart();
        _timer.Start();
        State = TimerGoalState.Running;
        OnPropertyChanged(nameof(Snapshot));
    }

    public void Stop()
    {
        if (State == TimerGoalState.Running)
        {
            CaptureRunningProgress();
        }

        if (State == TimerGoalState.Completed)
        {
            _timer.Stop();
            _stopwatch.Reset();
            OnPropertyChanged(nameof(Snapshot));
            return;
        }

        _stopwatch.Reset();
        _timer.Stop();
        State = TimerGoalState.Idle;

        OnPropertyChanged(nameof(Snapshot));
    }

    public void Reset()
    {
        _stopwatch.Reset();
        _timer.Stop();
        ElapsedTime = TimeSpan.Zero;
        RemainingTime = TargetDuration;
        State = TargetDuration == TimeSpan.Zero ? TimerGoalState.Completed : TimerGoalState.Idle;
        OnPropertyChanged(nameof(Snapshot));
    }

    public void Dispose()
    {
        _timer.Stop();
        _timer.Tick -= _tickHandler;
    }

    private void UpdateProgress()
    {
        if (State != TimerGoalState.Running)
        {
            return;
        }

        var currentElapsed = ElapsedTime + _stopwatch.Elapsed;
        ApplyProgress(currentElapsed);
    }

    private void CaptureRunningProgress()
    {
        var currentElapsed = ElapsedTime + _stopwatch.Elapsed;
        ApplyProgress(currentElapsed);
    }

    private void ApplyProgress(TimeSpan elapsed)
    {
        if (elapsed >= TargetDuration)
        {
            CompleteTimer();
            return;
        }

        ElapsedTime = elapsed;
        RemainingTime = TargetDuration - elapsed;
        OnPropertyChanged(nameof(Snapshot));
        _stopwatch.Restart();
    }

    private void CompleteTimer()
    {
        _timer.Stop();
        _stopwatch.Reset();
        ElapsedTime = TargetDuration;
        RemainingTime = TimeSpan.Zero;
        State = TimerGoalState.Completed;
        OnPropertyChanged(nameof(Snapshot));
    }

    private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(storage, value))
        {
            return false;
        }

        storage = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
