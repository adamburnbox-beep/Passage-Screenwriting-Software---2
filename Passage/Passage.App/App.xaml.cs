using System.Windows;
using System.Windows.Threading;
using Passage.App.Services;

namespace Passage.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        StartupDiagnostics.RegisterGlobalHandlers(this);
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        StartupDiagnostics.Write("Startup requested.");

        Dispatcher.BeginInvoke(new Action(CompleteStartup), DispatcherPriority.Background);
    }

    protected override void OnExit(ExitEventArgs e)
    {
        StartupDiagnostics.Write($"Application exiting with code {e.ApplicationExitCode}.");

        ThemeManager.Shutdown();
        base.OnExit(e);
    }

    private void CompleteStartup()
    {
        StartupDiagnostics.Write("Main startup sequence began.");

        try
        {
            ThemeManager.SetTheme(ThemeManager.DarkThemeName);
            StartupDiagnostics.Write("Theme initialized.");

            RecoveryDocument? recoveredDocument = null;
            if (RecoveryStorage.TryReadRecovery(out var pendingRecoveryDocument))
            {
                StartupDiagnostics.Write("Recovery snapshot found.");

                var result = MessageBox.Show(
                    "A recovered unsaved document was found. Restore it?",
                    "Passage",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question,
                    MessageBoxResult.Yes);

                if (result == MessageBoxResult.Yes)
                {
                    recoveredDocument = pendingRecoveryDocument;
                    StartupDiagnostics.Write("Recovery snapshot accepted.");
                }
                else
                {
                    RecoveryStorage.ClearRecoveryFile();
                    StartupDiagnostics.Write("Recovery snapshot discarded.");
                }
            }

            var mainWindow = new MainWindow(recoveredDocument);
            mainWindow.ContentRendered += MainWindow_ContentRendered;

            MainWindow = mainWindow;
            ShutdownMode = ShutdownMode.OnMainWindowClose;
            mainWindow.Show();
            mainWindow.Activate();
            StartupDiagnostics.Write("Main window show requested.");
        }
        catch (Exception exception)
        {
            StartupDiagnostics.WriteException("Startup failed.", exception);

            MessageBox.Show(
                $"Passage couldn't finish starting.{Environment.NewLine}{Environment.NewLine}Details were written to:{Environment.NewLine}{StartupDiagnostics.LogFilePath}",
                "Passage startup failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            Shutdown(-1);
        }
    }

    private void MainWindow_ContentRendered(object? sender, EventArgs e)
    {
        if (sender is Window window)
        {
            window.ContentRendered -= MainWindow_ContentRendered;
        }

        StartupDiagnostics.Write("Main window content rendered.");
    }
}
