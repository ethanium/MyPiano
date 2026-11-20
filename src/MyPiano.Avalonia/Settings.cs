using System.Text.Json;

namespace MyPiano;

/// <summary>Persisted user settings (replaces config.dat).</summary>
public sealed class Settings
{
    public int DeviceId { get; set; }
    /// <summary>Path to a General MIDI SoundFont (.sf2) for the built-in synth.</summary>
    public string? SoundFontPath { get; set; }
    /// <summary>True when the built-in SoundFont synth, not a Windows MIDI device, is the output.</summary>
    public bool UseSoundFont { get; set; }
    /// <summary>Reverb and chorus send levels (MIDI controllers 91 and 93), 0-127.</summary>
    public int Reverb { get; set; } = 40;
    public int Chorus { get; set; } = 0;
    /// <summary>Synth-only effects (SoundFont output), 0-127; Tone 64 = flat.</summary>
    public int Delay { get; set; } = 0;
    public int Drive { get; set; } = 0;
    public int Tremolo { get; set; } = 0;
    public int Tone { get; set; } = 64;
    public int[] Banks { get; set; } = new int[4];
    public int[] Drums { get; set; } = { 49, 57, 51, 45, 47, 48, 46, 42, 44, 37, 36, 40 };
    /// <summary>True for the dark colour theme (the default), false for light.</summary>
    public bool Dark { get; set; } = true;

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
