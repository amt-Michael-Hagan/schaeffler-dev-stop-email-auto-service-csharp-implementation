using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading.Tasks;
using EmailAutomationLegacy.Services;

namespace EmailAutomationLegacy
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        static async Task Main(string[] args)
        {
            try
            {
                // Configure logging
                log4net.Config.XmlConfigurator.Configure();
                
                Console.WriteLine("🤖 C# Email Automation Service (.NET Framework 4.0)");
                Console.WriteLine(new string('=', 60));

                // Validate configuration
                if (!ValidateConfiguration())
                {
                    Console.WriteLine("❌ Configuration validation failed. Exiting.");
                    Environment.Exit(1);
                }

                // Initialize services
                var tokenManager = new TokenManager();
                var emailProcessor = new EmailProcessor(new GraphClient(tokenManager));

                Console.WriteLine("📧 Starting email processing...");

                // Process emails
                var result = await emailProcessor.ProcessEmailsWithGraphAsync(new Dictionary<string, string>(), "");
                
                // Display results
                Console.WriteLine("\n" + new string('=', 60));
                Console.WriteLine("📊 Processing Complete");
                Console.WriteLine(new string('=', 60));
                Console.WriteLine($"Emails processed: {result.EmailsProcessed}");
                Console.WriteLine($"Attachments found: {result.TotalAttachments}");
                Console.WriteLine($"New downloads: {result.NewDownloads}");
                Console.WriteLine($"Already processed: {result.SkippedAttachments}");
                Console.WriteLine($"Downloads directory: {Path.GetFullPath(AppSettings.DownloadsDirectory)}");

                Console.WriteLine("\n✅ Service completed successfully!");
            }
            catch (Exception ex)
            {
                log.Error("Service execution failed", ex);
                Console.WriteLine($"❌ Service failed: {ex.Message}");
                Environment.Exit(1);
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }

        private static bool ValidateConfiguration()
        {
            var required = new[]
            {
                "ClientId",
                "TenantId", 
                "ClientSecret",
                "TargetEmail"
            };

            foreach (var setting in required)
            {
                var value = ConfigurationManager.AppSettings[setting];
                if (string.IsNullOrEmpty(value))
                {
                    Console.WriteLine($"❌ Missing required setting: {setting}");
                    return false;
                }
            }

            // Validate email format
            var targetEmail = ConfigurationManager.AppSettings["TargetEmail"];
            if (targetEmail.Contains("@outlook.com") || targetEmail.Contains("@hotmail.com") || targetEmail.Contains("@live.com"))
            {
                Console.WriteLine("❌ Personal Microsoft accounts are not supported with Application permissions");
                Console.WriteLine("   Please use an organizational account (e.g., user@company.com)");
                return false;
            }

            return true;
        }
    }
}