using System.Diagnostics;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using Avalonia.Threading;
using Passage.App.Views;

namespace Passage.App;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();

            // Detect system theme on Linux
            DetectAndApplySystemTheme();
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void DetectAndApplySystemTheme()
    {
        // Try D-Bus first, then gsettings, then fallback to dark
        var theme = DetectLinuxTheme();
        RequestedThemeVariant = theme;

        // Load the appropriate theme resources
        LoadThemeResources(theme == ThemeVariant.Light);
    }

    private ThemeVariant DetectLinuxTheme()
    {
        // Try D-Bus first (most reliable for GNOME/COSMIC)
        if (TryDbusThemeDetection(out var dbusTheme))
            return dbusTheme;

        // Fallback to gsettings command
        if (TryGsettingsThemeDetection(out var gsettingsTheme))
            return gsettingsTheme;

        // Fallback to environment variable
        if (TryEnvironmentThemeDetection(out var envTheme))
            return envTheme;

        // Default to dark (matches original app)
        return ThemeVariant.Dark;
    }

    private bool TryDbusThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        try
        {
            // Use gsettings via process - D-Bus requires additional libraries
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "gsettings",
                    Arguments = "get org.gnome.desktop.interface color-scheme",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.Start();
            var result = process.StandardOutput.ReadToEnd().Trim().TrimEnd('\''); // Trim trailing quote
            process.WaitForExit();

            // gsettings returns 'prefer-dark', 'prefer-light', or 'default'
            if (result.Contains("prefer-dark"))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (result.Contains("prefer-light"))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }
        catch
        {
            // Ignore errors and try next method
        }
        return false;
    }

    private bool TryGsettingsThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        try
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "gsettings",
                    Arguments = "get org.gnome.desktop.interface gtk-theme",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.Start();
            var result = process.StandardOutput.ReadToEnd().Trim().ToLower();
            process.WaitForExit();

            // Check if theme name contains 'dark'
            if (result.Contains("dark"))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (result.Contains("light"))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }
        catch
        {
            // Ignore errors and try next method
        }
        return false;
    }

    private bool TryEnvironmentThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        // Check common environment variables
        var gtkTheme = Environment.GetEnvironmentVariable("GTK_THEME");
        if (!string.IsNullOrEmpty(gtkTheme))
        {
            if (gtkTheme.Contains("dark", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (gtkTheme.Contains("light", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }

        // Check XDG current theme
        var xdgTheme = Environment.GetEnvironmentVariable("XDG_CURRENT_DESKTOP");
        if (!string.IsNullOrEmpty(xdgTheme))
        {
            // COSMIC/Unity typically uses dark mode by default
            if (xdgTheme.Contains("COSMIC", StringComparison.OrdinalIgnoreCase) ||
                xdgTheme.Contains("unity", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
        }

        return false;
    }

    private static Avalonia.Media.SolidColorBrush Brush(string hex) =>
        new(Avalonia.Media.Color.Parse(hex));

    // The single runtime source of truth for both palettes — a modern e-ink
    // monochrome scheme: paper-white surfaces with true-ink black in light, and
    // its direct inverse in dark. Called at startup (system detection above)
    // and by the View > Theme commands. Keep App.axaml's startup defaults in sync.
    public void LoadThemeResources(bool isLight)
    {
        if (Resources is null) return;

        if (isLight)
        {
            // Light theme: light-gray canvas, borderless pure-white cards that
            // float on soft shadows, solid ink black as the accent. Cards have
            // no visible stroke (CardBorder/EditorPageBorder are transparent) —
            // elevation comes from BoxShadow in the styles.
            Resources["ThemeBackground"] = Brush("#F1F1EF");
            Resources["WindowBackground"] = Brush("#F1F1EF");
            Resources["SurfaceBackground"] = Brush("#F7F7F5");
            Resources["SurfaceMutedBackground"] = Brush("#E9E9E6");
            Resources["SurfaceRaisedBackground"] = Brush("#FFFFFF");
            Resources["SurfaceBorder"] = Brush("#E4E4E0");
            Resources["ControlBackground"] = Brush("#FFFFFF");
            Resources["ControlForeground"] = Brush("#111110");
            Resources["ControlBorder"] = Brush("#E4E4E0");
            Resources["ControlAccent"] = Brush("#111110");
            Resources["ControlPressedBackground"] = Brush("#111110");
            Resources["ControlPressedForeground"] = Brush("#FFFFFF");
            Resources["HeaderText"] = Brush("#0B0B0A");
            Resources["SecondaryText"] = Brush("#5A5A55");
            Resources["MutedText"] = Brush("#98988F");
            Resources["EditorForeground"] = Brush("#171716");
            Resources["EditorPageBorder"] = Brush("#00000000");
            Resources["WindowForeground"] = Brush("#171716");
            Resources["HierarchyIndicatorBrush"] = Brush("#111110");
            Resources["DragOverBackground"] = Brush("#22111111");
            Resources["CardBackground"] = Brush("#FFFFFF");
            Resources["CardBorder"] = Brush("#00000000");
            Resources["CardAccent"] = Brush("#111110");

            // Editor syntax colours, darkened for the paper-white page
            Resources["SyntaxSceneHeading"] = Brush("#2F55B8");
            Resources["SyntaxCharacter"] = Brush("#94306A");
            Resources["SyntaxDialogue"] = Brush("#3C3C39");
            Resources["SyntaxParenthetical"] = Brush("#757570");
            Resources["SyntaxTransition"] = Brush("#5B4392");
            Resources["SyntaxSection"] = Brush("#1D6FA5");
            Resources["SyntaxSynopsis"] = Brush("#9A6A1C");
            Resources["SyntaxNote"] = Brush("#3E7D44");
        }
        else
        {
            // Dark theme (inverse: near-black canvas, slightly raised cards
            // with a hairline stroke since shadows carry nothing on black,
            // paper white as the accent)
            Resources["ThemeBackground"] = Brush("#0D0D0C");
            Resources["WindowBackground"] = Brush("#0D0D0C");
            Resources["SurfaceBackground"] = Brush("#141413");
            Resources["SurfaceMutedBackground"] = Brush("#232321");
            Resources["SurfaceRaisedBackground"] = Brush("#1B1B1A");
            Resources["SurfaceBorder"] = Brush("#262624");
            Resources["ControlBackground"] = Brush("#1B1B1A");
            Resources["ControlForeground"] = Brush("#F2F2F0");
            Resources["ControlBorder"] = Brush("#2E2E2C");
            Resources["ControlAccent"] = Brush("#F4F4F2");
            Resources["ControlPressedBackground"] = Brush("#F4F4F2");
            Resources["ControlPressedForeground"] = Brush("#0D0D0C");
            Resources["HeaderText"] = Brush("#FAFAF8");
            Resources["SecondaryText"] = Brush("#B2B2AD");
            Resources["MutedText"] = Brush("#6E6E69");
            Resources["EditorForeground"] = Brush("#EDEDEB");
            Resources["EditorPageBorder"] = Brush("#262624");
            Resources["WindowForeground"] = Brush("#EDEDEB");
            Resources["HierarchyIndicatorBrush"] = Brush("#F4F4F2");
            Resources["DragOverBackground"] = Brush("#25F4F4F2");
            Resources["CardBackground"] = Brush("#1B1B1A");
            Resources["CardBorder"] = Brush("#2A2A28");
            Resources["CardAccent"] = Brush("#F4F4F2");

            // Editor syntax colours for the ink-black page
            Resources["SyntaxSceneHeading"] = Brush("#6797FF");
            Resources["SyntaxCharacter"] = Brush("#C05587");
            Resources["SyntaxDialogue"] = Brush("#CEB2C9");
            Resources["SyntaxParenthetical"] = Brush("#8A8A85");
            Resources["SyntaxTransition"] = Brush("#8A6FC9");
            Resources["SyntaxSection"] = Brush("#4FC3F7");
            Resources["SyntaxSynopsis"] = Brush("#FFB74D");
            Resources["SyntaxNote"] = Brush("#81C784");
        }
    }
}
