using System;
using System.IO;
using System.Text.Json;

namespace Todo
{
    public enum MoveCompletedBehavior
    {
        MoveToTop = 0,
        DoNotMove = 1
    }

    public class UserSettings
    {
        public MoveCompletedBehavior MoveBehavior { get; set; } = MoveCompletedBehavior.MoveToTop;
    }

    public sealed class SettingsService
    {
        private static readonly Lazy<SettingsService> _lazy = new Lazy<SettingsService>(() => new SettingsService());
        public static SettingsService Instance => _lazy.Value;

        private static readonly string DataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "CHillSW", "TodoBMW");
        private static readonly string SettingsFile = Path.Combine(DataDirectory, "settings.json");

        public UserSettings Settings { get; private set; } = new UserSettings();

        private SettingsService()
        {
            Load();
        }

        public void Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var s = JsonSerializer.Deserialize<UserSettings>(json);
                    if (s != null)
                        Settings = s;
                }
                else
                {
                    Settings = new UserSettings();
                    Save();
                }
            }
            catch
            {
                // fallback to defaults on error
                Settings = new UserSettings();
            }
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(DataDirectory);
                var json = JsonSerializer.Serialize(Settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFile, json);
            }
            catch { }
        }
    }
}
