using System.IO;
using System.Text.Json;

namespace SoundSwitcher.Services;

public class WindowSettings
{
    public double Left { get; set; } = double.NaN;
    public double Top { get; set; } = double.NaN;
    public double Width { get; set; } = 600;
    public double Height { get; set; } = 520;

    // Whether the level meters are shown. Default off (they cost CPU during calls).
    public bool MetersVisible { get; set; } = false;

    private static string SettingsPath
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SoundSwitcher");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "window.json");
        }
    }

    public static WindowSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                var s = JsonSerializer.Deserialize<WindowSettings>(json);
                if (s != null) return s;
            }
        }
        catch { }
        return new WindowSettings();
    }

    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsPath, json);
        }
        catch { }
    }
}
