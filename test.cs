using System;
using System.Reflection;
using System.Windows;
using System.Windows.Threading;

namespace CrashTest
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                var asm = Assembly.LoadFrom(@"c:\Users\Adam\Downloads\Passage Screenwriting Software - Copy\artifacts\build-check\Passage.App.dll");
                var appType = asm.GetType("Passage.App.App");
                var app = (Application)Activator.CreateInstance(appType);
                app.InitializeComponent();
                
                app.Startup += (s, e) => {
                    Console.WriteLine("App Started. Scheduling button click...");
                    Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() => {
                        try {
                            var mainWindow = app.MainWindow;
                            var btn = mainWindow.FindName("BeatBoardWorkspaceButton") as System.Windows.Controls.Primitives.ToggleButton;
                            btn.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
                            Console.WriteLine("Button clicked successfully.");
                            app.Shutdown(0);
                        } catch (Exception ex) {
                            Console.WriteLine("CRASH IN CLICK: " + ex.ToString());
                            app.Shutdown(1);
                        }
                    }), DispatcherPriority.ApplicationIdle);
                };
                
                app.DispatcherUnhandledException += (s, e) => {
                    Console.WriteLine("UNHANDLED DISPATCHER EXCEPTION: " + e.Exception.ToString());
                    e.Handled = true;
                    app.Shutdown(1);
                };
                
                AppDomain.CurrentDomain.UnhandledException += (s, e) => {
                    Console.WriteLine("UNHANDLED DOMAIN EXCEPTION: " + e.ExceptionObject.ToString());
                };
                
                app.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("CRASH IN MAIN: " + ex.ToString());
            }
        }
    }
}
