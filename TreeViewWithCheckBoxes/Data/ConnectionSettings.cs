using System;
using System.IO;
using System.Text.Json;

namespace GenQ.Data
{
    /// <summary>
    /// Manages SQL Server connection settings with persistence
    /// </summary>
    public class ConnectionSettings
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GenQ",
            "connection.json");

        public string Server { get; set; } = ".\\SQLEXPRESS";
        public string Database { get; set; } = "GenQ_BOQ";
        public string Username { get; set; }
        public string Password { get; set; }
        public bool IntegratedSecurity { get; set; } = true;

        /// <summary>
        /// Load settings from file or return defaults
        /// </summary>
        public static ConnectionSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    return JsonSerializer.Deserialize<ConnectionSettings>(json) ?? new ConnectionSettings();
                }
            }
            catch
            {
                // Return defaults on error
            }
            return new ConnectionSettings();
        }

        /// <summary>
        /// Save settings to file
        /// </summary>
        public void Save()
        {
            try
            {
                string directory = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(this, options);
                File.WriteAllText(SettingsPath, json);
            }
            catch
            {
                // Silently fail - settings won't persist
            }
        }

        /// <summary>
        /// Create a connection using these settings
        /// </summary>
        public SqlServerConnection CreateConnection()
        {
            return new SqlServerConnection(Server, Database, Username, Password, IntegratedSecurity);
        }

        /// <summary>
        /// Test the connection with current settings
        /// </summary>
        public bool TestConnection(out string errorMessage)
        {
            using (var conn = CreateConnection())
            {
                return conn.TestConnection(out errorMessage);
            }
        }
    }
}
