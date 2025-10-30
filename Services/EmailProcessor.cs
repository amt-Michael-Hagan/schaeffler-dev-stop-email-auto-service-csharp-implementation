using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using EmailAutomationLegacy.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Move;
using AttachmentInfo = EmailAutomationLegacy.Models.AttachmentInfo;

namespace EmailAutomationLegacy.Services
{
    public class EmailProcessor
    {
        private readonly ProcessedEmailAttachmentTracker _trackingData;
        private readonly IGraphServiceClient _graphServiceClient;

        public EmailProcessor(IGraphServiceClient serviceClient)
        {
            _trackingData = LoadTrackingData();
            _graphServiceClient = serviceClient;
        }


        public async Task<ProcessingResult> ProcessEmailsWithGraphAsync(
            Dictionary<string, string> emailList, string logFilePath)
        {
            var result = new ProcessingResult();

            try
            {
                Console.WriteLine("Starting email processing with Microsoft Graph...");

                EnsureDirectoryExists(AppSettings.OutputPathGood, "Output");
                var logDir = EnsureDirectoryExists(Path.GetDirectoryName(AppSettings.LogFile) ?? AppSettings.LogFile, "Log");

                var allowedSenders = new HashSet<string>(emailList.Keys, StringComparer.OrdinalIgnoreCase);
                if (allowedSenders.Count == 0)
                {
                    Console.WriteLine("No allowed senders found — aborting process.");
                    return result;
                }

                Console.WriteLine($"Loaded {allowedSenders.Count} allowed senders/domains.");

                var blockedExtensions = ParseBlockedExtensions();

                // Resolve folders
                var importFolderId = GetFolderIdByDisplayName(AppSettings.InboxImportSubDir)
                                     ?? GetFolderIdByDisplayName("Inbox");

                if (string.IsNullOrEmpty(importFolderId))
                {
                    Console.WriteLine("Could not resolve import folder.");
                    return result;
                }

                // Fetch emails and attachments
                (string oldFolderId, MessageCollectionResponse messagesResponse) = await ReadEmailMessages(importFolderId);

                var messages = messagesResponse?.Value?.ToList() ?? new List<Message>();
                Console.WriteLine($"Found {messages.Count} messages with attachments in import folder.");

                if (messages.Count == 0)
                    return result;

                result.EmailsProcessed = messages.Count;

                var processedIds = new List<string>();

                await ProcessEmailAttachments(
                    result,
                    logDir,
                    allowedSenders,
                    blockedExtensions,
                    messages,
                    processedIds
                );

                // Move processed mails
                //await MoveProcessedMails(oldFolderId, processedIds);


                Console.WriteLine("Email processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Email processing failed: {ex.Message}");
                throw;
            }

            return result;
        }

        private static string EnsureDirectoryExists(string path, string type)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                Console.WriteLine($"Created {type.ToLower()} directory: {path}");
            }
            return path;
        }
        private async Task<(string oldFolderId, MessageCollectionResponse messagesResp)> ReadEmailMessages(string importFolderId)
        {
            var oldFolderId = GetFolderIdByDisplayName(AppSettings.InboxOldSubDir) ?? importFolderId;

            // Build filter for timeframe and attachments
            var hoursAgo = DateTime.UtcNow.AddHours(-AppSettings.HoursToFetch);
            var filter = $"receivedDateTime ge {hoursAgo:yyyy-MM-ddTHH:mm:ss.fffZ} and hasAttachments eq true";

            // Query messages in the import folder via Graph SDK
            var messagesResp = await _graphServiceClient.ReadEmailMessages(oldFolderId, filter);
            return (oldFolderId, messagesResp);
        }

        private ProcessedEmailAttachmentTracker LoadTrackingData()
        {
            try
            {
                if (File.Exists(AppSettings.TrackingFile))
                {
                    var json = File.ReadAllText(AppSettings.TrackingFile);
                    if (string.IsNullOrWhiteSpace(json))
                    {
                        return new ProcessedEmailAttachmentTracker();
                    }

                    var data = JsonConvert.DeserializeObject<ProcessedEmailAttachmentTracker>(json);
                    Console.WriteLine($"Loaded tracking data: {data.Attachments.Count} attachments previously processed");
                    return data;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load tracking file, using empty state: {ex.Message}");
            }

            return new ProcessedEmailAttachmentTracker();
        }

        private async Task ProcessEmailAttachments(
            ProcessingResult result,
            string logDir,
            HashSet<string> allowedSenders,
            HashSet<string> blockedExtensions,
            List<Message> incomingMessages,
            List<string> processedMessageIds)
        {
            Directory.CreateDirectory(AppSettings.OutputPathGood);

            foreach (var message in incomingMessages)
            {
                if (!IsSenderAllowed(message, allowedSenders, out var sender))
                    continue;

                Console.WriteLine(
                    $"\n--- Processing message: {message.Subject} From: {sender} Date: {message.ReceivedDateTime:yyyy-MM-dd HH:mm:ss}");

                var attachments = await GetAttachmentsAsync(message.Id);
                if (attachments.Count == 0)
                {
                    Console.WriteLine("  No attachments found");
                    continue;
                }

                foreach (var attachment in attachments.OfType<FileAttachment>())
                {
                    if (attachment.ContentBytes == null || IsAttachmentProcessed(message.Id, attachment.Id))
                    {
                        LogSkippedAttachment(result, attachment.Name);
                        continue;
                    }

                    HandleAttachmentAsync(result, logDir, sender, message.Id, attachment, blockedExtensions);
                }

                processedMessageIds.Add(message.Id);
            }
        }

        private static bool IsSenderAllowed(Message message, HashSet<string> allowedSenders, out string sender)
        {
            sender = message.From?.EmailAddress?.Address?.ToLowerInvariant();
            if (string.IsNullOrEmpty(sender) || !sender.Contains('@'))
                return false;

            var domain = sender.Substring(sender.LastIndexOf('@'));
            return allowedSenders.Contains(sender) || allowedSenders.Contains(domain);
        }

        private async Task<IList<Attachment>> GetAttachmentsAsync(string messageId)
        {
            return await _graphServiceClient.GetAttachmentsAsync(messageId);
        }

        private static void LogSkippedAttachment(ProcessingResult result, string name)
        {
            Console.WriteLine($"Skipped (file already processed, not a file or no content): {name}");
            result.SkippedAttachments++;
        }

        private void HandleAttachmentAsync(
            ProcessingResult result,
            string logDir,
            string sender,
            string messageId,
            FileAttachment attachment,
            HashSet<string> blockedExtensions)
        {
            var fileName = attachment.Name ?? "attachment";
            var ext = Path.GetExtension(fileName).ToLowerInvariant();

            if (blockedExtensions.Contains(ext))
            {
                LogBlockedAttachmentAsync(result, logDir, sender, fileName, ext);
                return;
            }

            var targetPath = GetUniqueFilePath(AppSettings.OutputPathGood, fileName);
            File.WriteAllBytes(targetPath, attachment.ContentBytes);

            if (AppSettings.LogAttachments)
            {
                var logPath = Path.Combine(logDir, AppSettings.LogFileAttachments);
                File.AppendAllText(logPath,
                    $"{sender} {fileName} {DateTime.Now}{Environment.NewLine}");
            }

            MarkAttachmentProcessed(messageId, attachment.Id, fileName);
            result.TotalAttachments++;
            result.NewDownloads++;
            Console.WriteLine($"Downloaded: {fileName} ({attachment.ContentBytes.Length / 1024.0:F2} KB)");
        }

        private static void LogBlockedAttachmentAsync(
            ProcessingResult result,
            string logDir,
            string sender,
            string fileName,
            string ext)
        {
            // Don’t log common image files as blocked
            if (!ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) &&
                !ext.Equals(".png", StringComparison.OrdinalIgnoreCase))
            {
                var blockedLogPath = Path.Combine(logDir, AppSettings.LogFileBlockedFiles);
                File.AppendAllText(blockedLogPath,
                    $"{sender} {fileName} {DateTime.Now}{Environment.NewLine}");
            }

            result.SkippedAttachments++;
            Console.WriteLine($"Blocked by extension: {fileName}");
        }

        private static string GetUniqueFilePath(string directory, string fileName)
        {
            var targetPath = Path.Combine(directory, fileName);
            if (!File.Exists(targetPath))
                return targetPath;

            var uniquePrefix = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_";
            return Path.Combine(directory, uniquePrefix + fileName);
        }

        private async Task MoveProcessedMails(string oldFolderId, List<string> processedMessageIds)
        {
            foreach (var messageId in processedMessageIds)
            {
                try
                {
                    var moveBody = new MovePostRequestBody { DestinationId = oldFolderId };
                    await _graphServiceClient.MoveProcessedMails(messageId, moveBody);
                    Console.WriteLine($"Moved message {messageId} to folder id {oldFolderId}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to move message {messageId}: {ex.Message}");
                }
            }
        }

        private HashSet<string> LoadAllowedSenders()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Emails.xml");
                if (!File.Exists(path))
                {
                    Console.WriteLine($"Emails.xml not found at {path}");
                    return set;
                }

                var doc = new XmlDocument();
                doc.Load(path);

                foreach (XmlNode node in doc.SelectNodes("//*"))
                {
                    var text = node.InnerText?.Trim();
                    if (string.IsNullOrEmpty(text)) continue;

                    var lower = text.ToLowerInvariant();
                    var parts = lower.Split(new[] { ',', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var p in parts)
                    {
                        var v = p.Trim();
                        if (string.IsNullOrEmpty(v)) continue;
                        if (v.StartsWith("@") || v.Contains("@"))
                            set.Add(v);
                        else
                            set.Add(v);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load Emails.xml: {ex.Message}");
            }

            return set;
        }

        private HashSet<string> ParseBlockedExtensions()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var raw = AppSettings.BlockedFiles ?? string.Empty;
                var parts = raw.Split(new[] { ',', ';', '.' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts)
                {
                    var ext = p.Trim().ToLowerInvariant();
                    if (!ext.StartsWith(".")) ext = "." + ext;
                    set.Add(ext);
                }
            }
            catch
            {
            }

            return set;
        }

        private string GetFolderIdByDisplayName(string displayName)
        {
            return _graphServiceClient.GetFolderIdByDisplayName(displayName);
        }

        //Todo: Save data using file hashes instead.
        private void SaveTrackingData()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_trackingData, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(AppSettings.TrackingFile, json);
                Console.WriteLine("Tracking data saved successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save tracking file: {ex.Message}", ex);
            }
        }

        private void MarkAttachmentProcessed(string emailId, string attachmentId, string fileName)
        {
            var key = $"{emailId}_{attachmentId}";
            _trackingData.Attachments[key] = new AttachmentInfo
            {
                FileName = fileName,
                DownloadedAt = DateTime.UtcNow,
                EmailId = emailId
            };
            SaveTrackingData();
        }

        private bool IsAttachmentProcessed(string emailId, string attachmentId)
        {
            var key = $"{emailId}_{attachmentId}";
            return _trackingData.Attachments.ContainsKey(key);
        }

    }
}