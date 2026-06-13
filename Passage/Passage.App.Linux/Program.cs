using Avalonia;
using Avalonia.Headless;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace Passage.App;

sealed class Program
{
    private static Process? _xwaylandProcess;

    [STAThread]
    public static void Main(string[] args)
    {
        try
        {
            if (args.Contains("--headless"))
            {
                BuildHeadlessApp().StartWithClassicDesktopLifetime(args);
                return;
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                var wayland = Environment.GetEnvironmentVariable("WAYLAND_DISPLAY");
                var display = Environment.GetEnvironmentVariable("DISPLAY");
                bool hasWayland = !string.IsNullOrEmpty(wayland);
                bool hasX11 = !string.IsNullOrEmpty(display);

                bool verbose = args.Contains("--diag");
                if (verbose)
                {
                    Console.Error.WriteLine($"[diag] WAYLAND_DISPLAY='{wayland}' DISPLAY='{display}'");
                    Console.Error.WriteLine($"[diag] XDG_RUNTIME_DIR='{Environment.GetEnvironmentVariable("XDG_RUNTIME_DIR")}'");
                    Console.Error.WriteLine($"[diag] /tmp/.X11-unix: {DescribeX11Sockets()}");
                }

                if (!hasWayland && !hasX11)
                {
                    Console.Error.WriteLine("No display server detected (WAYLAND_DISPLAY and DISPLAY are both unset).");
                    Console.Error.WriteLine("Run with --headless to start without a GUI.");
                    Environment.ExitCode = 1;
                    return;
                }

                // Avalonia 12 only has an X11 backend on Linux. We need a working DISPLAY.
                // Start our own XWayland instance when there is no usable X11 display:
                // either DISPLAY is unset, or it is set but nothing is listening on it.
                if (hasWayland && (!hasX11 || !X11SocketExists(display!)))
                {
                    if (verbose && hasX11)
                        Console.Error.WriteLine($"[diag] DISPLAY={display} set but no live socket — starting our own XWayland.");

                    if (!TryStartXwayland(verbose, out var newDisplay))
                    {
                        Console.Error.WriteLine("Could not start XWayland.");
                        Console.Error.WriteLine("Make sure xwayland is installed: sudo apt install xwayland");
                        Environment.ExitCode = 1;
                        return;
                    }
                    Environment.SetEnvironmentVariable("DISPLAY", newDisplay);
                    if (verbose) Console.Error.WriteLine($"[diag] DISPLAY now set to {newDisplay}");
                }
            }

            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fatal error starting application: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
            Environment.ExitCode = 1;
        }
        finally
        {
            try { _xwaylandProcess?.Kill(); } catch { }
            _xwaylandProcess?.Dispose();
        }
    }

    private static bool X11SocketExists(string display)
    {
        // DISPLAY is like ":0" or ":0.0" — extract the display number.
        var spec = display.TrimStart(':');
        var dot = spec.IndexOf('.');
        if (dot >= 0) spec = spec[..dot];
        return File.Exists($"/tmp/.X11-unix/X{spec}");
    }

    private static string DescribeX11Sockets()
    {
        try
        {
            if (!Directory.Exists("/tmp/.X11-unix")) return "(directory does not exist)";
            var entries = Directory.GetFileSystemEntries("/tmp/.X11-unix");
            return entries.Length == 0 ? "(empty)" : string.Join(", ", Array.ConvertAll(entries, Path.GetFileName));
        }
        catch (Exception ex) { return $"(error: {ex.Message})"; }
    }

    // Starts an XWayland instance on a free display number, waits up to 5 s for
    // its socket to appear, and stores the process handle for cleanup on exit.
    private static bool TryStartXwayland(bool verbose, out string display)
    {
        display = ":1";

        // Common install paths on Debian/Ubuntu-based distros
        string[] binarySearchPaths =
        [
            "/usr/bin/Xwayland",
            "/usr/lib/xorg/Xwayland",
            "/usr/local/bin/Xwayland",
        ];

        var binary = Array.Find(binarySearchPaths, File.Exists);
        if (binary is null)
        {
            Console.Error.WriteLine("Xwayland binary not found. Looked in: " + string.Join(", ", binarySearchPaths));
            return false;
        }
        if (verbose) Console.Error.WriteLine($"[diag] using Xwayland binary: {binary}");

        for (int n = 1; n <= 9; n++)
        {
            var socketPath = $"/tmp/.X11-unix/X{n}";
            if (File.Exists(socketPath)) continue;   // display number already in use

            display = $":{n}";
            // Rootful (NOT -rootless): Xwayland creates a single normal Wayland
            // surface for its X screen, which the compositor composites like any
            // other app. Rootless mode only works when the compositor spawned and
            // is integrated with Xwayland itself (COSMIC does not wire up a foreign
            // rootless Xwayland, so its windows never appear).
            // -decorate: Xwayland draws its own titlebar/window controls — needed
            //            because there is no window manager inside this X server.
            // -noreset:  keep the server alive while our app is the only client.
            // -geometry: initial size of the X screen / Xwayland window.
            var psi = new ProcessStartInfo(binary, $"{display} -decorate -noreset -geometry 1600x1000")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
            };

            if (verbose) Console.Error.WriteLine($"[diag] starting: {binary} {psi.Arguments}");

            Process? proc;
            try { proc = Process.Start(psi); }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[diag] failed to launch Xwayland: {ex.Message}");
                continue;
            }
            if (proc is null) continue;

            // Capture Xwayland's diagnostics so we can report why it exited.
            var stderrBuffer = new System.Text.StringBuilder();
            proc.ErrorDataReceived += (_, e) => { if (e.Data != null) stderrBuffer.AppendLine(e.Data); };
            proc.BeginErrorReadLine();

            // Poll for the socket — XWayland typically appears within ~500 ms
            const int maxWaitMs = 5000;
            const int pollMs = 100;
            int waited = 0;
            while (!File.Exists(socketPath) && waited < maxWaitMs && !proc.HasExited)
            {
                Thread.Sleep(pollMs);
                waited += pollMs;
            }

            if (File.Exists(socketPath) && !proc.HasExited)
            {
                _xwaylandProcess = proc;
                if (verbose) Console.Error.WriteLine($"[diag] Xwayland socket ready at {socketPath} after ~{waited}ms");
                return true;
            }

            // Failed — surface Xwayland's own output so we know why.
            if (proc.HasExited)
                Console.Error.WriteLine($"[diag] Xwayland exited early (code {proc.ExitCode}).");
            else
                Console.Error.WriteLine($"[diag] Xwayland socket {socketPath} never appeared after {maxWaitMs}ms.");

            var err = stderrBuffer.ToString().Trim();
            if (err.Length > 0) Console.Error.WriteLine("[diag] Xwayland output:\n" + err);

            try { if (!proc.HasExited) proc.Kill(); } catch { }
            proc.Dispose();
        }

        return false;
    }

    private static AppBuilder BuildHeadlessApp() =>
        AppBuilder.Configure<App>()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions())
            .WithInterFont()
            .LogToTrace();

    // Also used by the Avalonia visual designer — do not remove.
    public static AppBuilder BuildAvaloniaApp()
    {
        var builder = AppBuilder.Configure<App>().UsePlatformDetect();
#if DEBUG
        builder.WithDeveloperTools();
#endif
        return builder.WithInterFont().LogToTrace();
    }
}
