using System.IO;
using System.Security;
using System.Text.Json;
using Passage.Core.Goals;

namespace Passage.App.Services;

public sealed record SessionDocumentState
{
    public string? FilePath { get; init; }

    public string Text { get; init; } = string.Empty;

    public bool IsDirty { get; init; }

    public GoalConfiguration GoalConfiguration { get; init; } = new();

    public SessionGoalConfiguration SessionGoalConfiguration { get; init; } = new();

    public double EditorZoomPercent { get; init; } = 100.0;
}

public sealed record SessionState(IReadOnlyList<SessionDocumentState> Documents, int SelectedIndex);

public static class SessionStorage
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };

    public static string SessionFilePath
    {
        get
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Passage");

            return Path.Combine(directory, "session.json");
        }
    }

    public static bool TryLoadSession(out SessionState? state)
    {
        state = null;

        try
        {
            if (!File.Exists(SessionFilePath))
            {
                return false;
            }

            var json = File.ReadAllText(SessionFilePath);
            state = JsonSerializer.Deserialize<SessionState>(json, SerializerOptions);
            return state is not null;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (SecurityException)
        {
            return false;
        }
    }

    public static void SaveSession(SessionState state)
    {
        ArgumentNullException.ThrowIfNull(state);

        try
        {
            var directory = Path.GetDirectoryName(SessionFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(state, SerializerOptions);
            File.WriteAllText(SessionFilePath, json);
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
