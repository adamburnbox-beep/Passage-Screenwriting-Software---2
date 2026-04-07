using System.Windows.Input;

namespace Passage.App.Utilities;

internal sealed class DelegateCommand<T> : ICommand
{
    private readonly Action<T?> _execute;
    private readonly Predicate<T?>? _canExecute;

    public DelegateCommand(Action<T?> execute, Predicate<T?>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public bool CanExecute(object? parameter)
    {
        return _canExecute?.Invoke(CastParameter(parameter)) ?? true;
    }

    public void Execute(object? parameter)
    {
        _execute(CastParameter(parameter));
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    private static T? CastParameter(object? parameter)
    {
        return parameter is T typedParameter ? typedParameter : default;
    }
}
