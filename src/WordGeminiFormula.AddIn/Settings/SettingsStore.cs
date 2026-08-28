using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace WordGeminiFormula.AddIn.Settings
{
    public sealed class SettingsStore
    {
        private readonly string _settingsPath;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();

        public SettingsStore()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordGeminiFormula");
            Directory.CreateDirectory(dir);
            _settingsPath = Path.Combine(dir, "settings.json");
        }

        public AppSettings Load()
        {
            if (!File.Exists(_settingsPath))
                return new AppSettings();

            try
            {
                string json = File.ReadAllText(_settingsPath, Encoding.UTF8);
                return _json.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings, string plainApiKey)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            settings.encryptedApiKey = Protect(plainApiKey ?? string.Empty);
            string json = _json.Serialize(settings);
            File.WriteAllText(_settingsPath, json, new UTF8Encoding(false));
        }

        public string GetApiKey(AppSettings settings)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.encryptedApiKey))
                return string.Empty;

            try
            {
                return Unprotect(settings.encryptedApiKey);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string Protect(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            byte[] protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        private static string Unprotect(string base64)
        {
            byte[] protectedBytes = Convert.FromBase64String(base64);
            byte[] bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }
    }
}
