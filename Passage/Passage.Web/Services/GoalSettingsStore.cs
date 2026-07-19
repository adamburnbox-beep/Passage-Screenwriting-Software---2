using System.Text.Json;
using System.Text.Json.Serialization;
using Passage.Core.Goals;

namespace Passage.Web.Services;

/// <summary>
/// Goal configuration persisted on the data volume so targets survive
/// container restarts. Defaults mirror the desktop apps.
/// </summary>
public sealed class GoalSettings
{
    public GoalType DocumentType { get; set; } = GoalType.WordCount;
    public int DocumentWordTarget { get; set; } = 1000;
    public int DocumentPageTarget { get; set; } = 120;
    public GoalType SessionType { get; set; } = GoalType.WordCount;
    public int SessionWordTarget { get; set; } = 1000;
    public int SessionPageTarget { get; set; } = 10;
    public int SessionTimerMinutes { get; set; } = 25;
}

public sealed class GoalSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    private readonly string _path;
    private readonly object _gate = new();

    public GoalSettingsStore(ScriptLibrary library)
    {
        _path = Path.Combine(library.RootPath, ".passage", "goals.json");
    }

    public GoalSettings Load()
    {
        lock (_gate)
        {
            try
            {
                if (File.Exists(_path))
                {
                    return JsonSerializer.Deserialize<GoalSettings>(File.ReadAllText(_path), JsonOptions)
                        ?? new GoalSettings();
                }
            }
            catch (Exception exception) when (exception is IOException or JsonException)
            {
                // A corrupt or unreadable settings file falls back to defaults;
                // the next save rewrites it.
            }

            return new GoalSettings();
        }
    }

    public void Save(GoalSettings settings)
    {
        lock (_gate)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
            File.WriteAllText(_path, JsonSerializer.Serialize(settings, JsonOptions));
        }
    }
}
