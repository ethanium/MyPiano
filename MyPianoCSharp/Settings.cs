using System.Text.Json;

namespace MyPiano;

/// <summary>Persisted user settings (replaces config.dat).</summary>
public sealed class Settings
{
    public int DeviceId { get; set; }
    public int[] Banks { get; set; } = new int[4];
    public int[] Drums { get; set; } = { 49, 57, 51, 45, 47, 48, 46, 42, 44, 37, 36, 40 };

    private static string FilePath => Path.Combine(AppContext.BaseDirectory, "config.json");

    public static Settings Load()
    {
        try
        {
            var s = JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath));
            if (s is { Banks.Length: 4, Drums.Length: 12 }) return s;
        }
        catch { /* missing or corrupt config: fall back to defaults */ }
        return new Settings();
    }

    public void Save()
    {
        try { File.WriteAllText(FilePath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true })); }
        catch { /* read-only location: ignore, same as the original */ }
    }
}
