using System;
using System.Configuration;

namespace EmailAutomationLegacy
{
    public static class AppSettings
    {
        public static string ClientId => ConfigurationManager.AppSettings["ClientId"];
        public static string TenantId => ConfigurationManager.AppSettings["TenantId"];
        public static string ClientSecret => ConfigurationManager.AppSettings["ClientSecret"];
        public static string TargetEmail => ConfigurationManager.AppSettings["TargetEmail"];
        
        public static int HoursToFetch => GetIntSetting("HoursToFetch", 24);
        public static bool IncludeJunkEmails => GetBoolSetting("IncludeJunkEmails", false);
        public static string DownloadsDirectory => GetSetting("DownloadsDirectory", "./downloads");
        public static string LogsDirectory => GetSetting("LogsDirectory", "./logs");
        public static string TrackingFile => GetSetting("TrackingFile", "./processed_emails.json");
        public static int RetryAttempts => GetIntSetting("RetryAttempts", 3);
        public static int RetryDelayMs => GetIntSetting("RetryDelayMs", 2000);

        private static string GetSetting(string key, string defaultValue = null)
        {
            return ConfigurationManager.AppSettings[key] ?? defaultValue;
        }

        private static int GetIntSetting(string key, int defaultValue)
        {
            var value = ConfigurationManager.AppSettings[key];
            return int.TryParse(value, out int result) ? result : defaultValue;
        }

        private static bool GetBoolSetting(string key, bool defaultValue)
        {
            var value = ConfigurationManager.AppSettings[key];
            return bool.TryParse(value, out bool result) ? result : defaultValue;
        }
    }
}