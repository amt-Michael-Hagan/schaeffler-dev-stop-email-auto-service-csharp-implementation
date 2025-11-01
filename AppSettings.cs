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
        public static string OutputPathGood => GetSetting("OutputPathGood", @"/Users/gentlekiwi/source/repos/schaeffler-dev-stop-email-auto-service-csharp-implementation/ExternalImport");
        public static string LogFile => GetSetting("LogFile", @"/Users/gentlekiwi/source/repos/schaeffler-dev-stop-email-auto-service-csharp-implementation/email_automation_log");
        public static string InboxImportSubDir { get; set; }
        public static string InboxOldSubDir => "ExternalOld";
        public static bool LogAttachments { get; set; } = true;
        public static string LogFileAttachments => GetSetting("LogFileAttachments", @"C:\Users\Hagan\Desktop\Projects\Project Files\schaeffler-dev-stop-email-auto-service-csharp-implementation\attachments.log");
        public static string LogFileBlockedFiles => GetSetting("LogFileBlockedFiles", @"C:\Users\Hagan\Desktop\Projects\Project Files\schaeffler-dev-stop-email-auto-service-csharp-implementation\blocked_files.log");
        public static string BlockedFiles => GetSetting("BlockedFiles", ".exe,.bat,.cmd,.com,.scr,.pif,.vbs,.js,.jar,.zip,.rar,.7z");

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