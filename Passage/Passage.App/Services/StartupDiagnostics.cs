using System.IO;
using System.Security;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace Passage.App.Services;

public static class StartupDiagnostics
{
    private static readonly object SyncRoot = new();
    private static bool _handlersRegistered;

    public static string LogFilePath
    {
        get
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Passage");

            return Path.Combine(directory, "startup.log");
        }
    }

    public static void RegisterGlobalHandlers(Application application)
    {
        ArgumentNullException.ThrowIfNull(application);

        if (_handlersRegistered)
        {
            return;
        }

        application.DispatcherUnhandledException += Application_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        _handlersRegistered = true;
        Write("Registered global startup exception handlers.");
    }

    public static void Write(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        AppendLine($"{DateTimeOffset.Now:O} {message}");
    }

    public static void WriteException(string context, Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        var builder = new StringBuilder();
        builder.AppendLine($"{DateTimeOffset.Now:O} {context}");
        builder.AppendLine(exception.ToString());
        AppendLine(builder.ToString().TrimEnd());
    }

    private static void Application_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        WriteException("Dispatcher unhandled exception.", e.Exception);
    }

    private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            WriteException("AppDomain unhandled exception.", exception);
            return;
        }

        Write($"AppDomain unhandled exception: {e.ExceptionObject}");
    }

    private static void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteException("Unobserved task exception.", e.Exception);
    }

    private static void AppendLine(string text)
    {
        try
        {
            var directory = Path.GetDirectoryName(LogFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            lock (SyncRoot)
            {
                File.AppendAllText(LogFilePath, text + Environment.NewLine, Encoding.UTF8);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        catch (SecurityException)
        {
        }
    }
}
