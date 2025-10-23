using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using Newtonsoft.Json;
using EmailAutomationLegacy.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Move;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using AttachmentInfo = EmailAutomationLegacy.Models.AttachmentInfo;

namespace EmailAutomationLegacy.Services
{
    public class EmailProcessor
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(EmailProcessor));

        private readonly TokenManager _tokenManager;
        private readonly GraphApiClient _graphClient;
        private TrackingData _trackingData;
        private readonly GraphServiceClient _graphServiceClient;

        public EmailProcessor(TokenManager tokenManager)
        {
            _tokenManager = tokenManager;
            _graphClient = new GraphApiClient(_tokenManager);
            // _trackingData = LoadTrackingData();

            var handler = new HttpClientHandler();

            var httpClient = new HttpClient(handler, disposeHandler: true)
            {
                Timeout = TimeSpan.FromSeconds(100)
            };

            var authProvider = new SimpleAuthenticationProvider(tokenManager);

            _graphServiceClient = new GraphServiceClient(httpClient, authProvider);
        }

        /*public bool TestConnection()
        {
            try
            {
                log.Info("Testing Microsoft Graph API connection...");

                // Test token acquisition
                var token = _tokenManager.GetAccessToken();
                if (string.IsNullOrEmpty(token))
                {
                    log.Error("Failed to obtain access token");
                    return false;
                }

                // Test API call - get user info
                var userEndpoint = $"users/{AppSettings.TargetEmail}";
                var user = _graphClient.Get<GraphUser>(userEndpoint);

                if (user == null)
                {
                    log.Error("Failed to retrieve user information");
                    return false;
                }

                log.Info($"Successfully connected. User: {user.DisplayName} ({user.Mail ?? user.UserPrincipalName})");

                // Test mailbox access
                var messagesEndpoint = $"users/{AppSettings.TargetEmail}/messages?$top=1";
                var messagesResponse = _graphClient.GetPaged<GraphMessage>(messagesEndpoint);

                log.Info($"Mailbox access confirmed. Found {messagesResponse.Value?.Count ?? 0} messages in test query");
                return true;
            }
            catch (Exception ex)
            {
                log.Error("Connection test failed", ex);
                return false;
            }

        }*/
        /*public async Task<bool> TestConnection()
        {
            try
            {
                log.Info("Testing Microsoft Graph API connection...");

                // Test mailbox access
                var messages = await _graphServiceClient.Users[AppSettings.TargetEmail].Messages.GetAsync();

                log.Info($"Mailbox access confirmed. Found {messages.Value.Count} messages in test query");

                var folders = await _graphServiceClient.Users[AppSettings.TargetEmail].MailFolders
                    .GetAsync();

                log.Info($"Mail folders access confirmed. Found {folders.Value.Count} mail folders in test query");

                foreach (var mailFolder in folders.Value)
                {
                    log.Info($"Mail folder: {mailFolder.DisplayName}");
                }
                return true;
            }
            catch (ServiceException ex) when (ex.ResponseStatusCode == 403)
            {
                log.Error("Permission denied. Please check if the application has the required API permissions in Azure AD.");
                log.Error($"Required permissions: Mail.Read, User.Read (Application permissions)");
                log.Error($"Error details: {ex.Message}");
                IEnumerable<string> requestId = new List<string>();
                if (ex.ResponseHeaders?.TryGetValues("request-id", out requestId) == true)
                {
                    log.Error($"Request ID: {string.Join(", ", requestId)}");
                }
                return false;
            }
            catch (Exception ex)
            {
                log.Error("Connection test failed", ex);
                return false;
            }
        }*/
        /*public ProcessingResult ProcessEmails()
        {
            var result = new ProcessingResult();

            try
            {
                log.Info("Starting email processing...");
                log.Info($"Target mailbox: {AppSettings.TargetEmail}");
                log.Info($"Fetching emails from last {AppSettings.HoursToFetch} hours");

                // Calculate date filter
                var hoursAgo = DateTime.UtcNow.AddHours(-AppSettings.HoursToFetch);
                var filterDate = hoursAgo.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

                // Build filter for emails with attachments
                var filter = $"receivedDateTime ge {filterDate} and hasAttachments eq true";
                var select = "id,subject,from,receivedDateTime,hasAttachments,parentFolderId";
                var orderBy = "receivedDateTime desc";

                var messagesEndpoint = $"users/{HttpUtility.UrlEncode(AppSettings.TargetEmail)}/messages?" +
                    $"$filter={HttpUtility.UrlEncode(filter)}&" +
                    $"$select={select}&" +
                    $"$orderby={orderBy}&" +
                    $"$top=999";

                var messagesResponse = _graphClient.GetPaged<GraphMessage>(messagesEndpoint);
                var emails = messagesResponse.Value ?? new List<GraphMessage>();

                log.Info($"Found {emails.Count} emails with attachments");

                if (emails.Count == 0)
                {
                    log.Info("No emails to process");
                    return result;
                }

                // Get sent items folder to exclude sent emails
                string sentFolderId = null;
                try
                {
                    var sentFolder = _graphClient.Get<GraphFolder>($"users/{HttpUtility.UrlEncode(AppSettings.TargetEmail)}/mailFolders/sentitems");
                    sentFolderId = sentFolder?.Id;
                }
                catch (Exception ex)
                {
                    log.Warn($"Could not get Sent Items folder: {ex.Message}");
                }

                // Filter out sent emails
                var targetEmailLower = AppSettings.TargetEmail.ToLowerInvariant();
                var incomingEmails = emails.Where(email =>
                {
                    // Exclude if in Sent Items folder
                    if (!string.IsNullOrEmpty(sentFolderId) && email.ParentFolderId == sentFolderId)
                        return false;

                    // Exclude if sender is the target mailbox
                    if (email.From?.EmailAddress?.Address?.ToLowerInvariant() == targetEmailLower)
                        return false;

                    return true;
                }).ToList();

                var filteredCount = emails.Count - incomingEmails.Count;
                if (filteredCount > 0)
                {
                    log.Info($"Filtered out {filteredCount} sent emails (processing only incoming emails)");
                }

                result.EmailsProcessed = incomingEmails.Count;

                // Process each email
                for (int i = 0; i < incomingEmails.Count; i++)
                {
                    var email = incomingEmails[i];
                    log.Info($"\n--- Email {i + 1}/{incomingEmails.Count} ---");
                    log.Info($"Subject: {email.Subject}");
                    log.Info($"From: {email.From?.EmailAddress?.Address}");
                    log.Info($"Date: {email.ReceivedDateTime:yyyy-MM-dd HH:mm:ss}");

                    if (email.HasAttachments)
                    {
                        var attachmentResult = ProcessEmailAttachments(email);
                        result.TotalAttachments += attachmentResult.Total;
                        result.NewDownloads += attachmentResult.Downloaded;
                        result.SkippedAttachments += attachmentResult.Skipped;
                    }
                }

                // Save tracking data
                SaveTrackingData();
                log.Info("Processing completed successfully");

            }
            catch (Exception ex)
            {
                log.Error("Email processing failed", ex);
                throw;
            }

            return result;
        }

        private (int Total, int Downloaded, int Skipped) ProcessEmailAttachments(GraphMessage email)
        {
            try
            {
                var attachmentsEndpoint = $"users/{HttpUtility.UrlEncode(AppSettings.TargetEmail)}/messages/{email.Id}/attachments";
                var attachmentsResponse = _graphClient.GetPaged<GraphAttachment>(attachmentsEndpoint);
                var attachments = attachmentsResponse.Value ?? new List<GraphAttachment>();

                if (attachments.Count == 0)
                {
                    log.Info("  No attachments found");
                    return (0, 0, 0);
                }

                log.Info($"  Found {attachments.Count} attachment(s)");

                // Ensure downloads directory exists
                if (!Directory.Exists(AppSettings.DownloadsDirectory))
                {
                    Directory.CreateDirectory(AppSettings.DownloadsDirectory);
                }

                int downloaded = 0;
                int skipped = 0;

                foreach (var attachment in attachments)
                {
                    if (IsAttachmentProcessed(email.Id, attachment.Id))
                    {
                        log.Info($"    ⏭️  Skipped (already downloaded): {attachment.Name}");
                        skipped++;
                        continue;
                    }

                    if (!string.IsNullOrEmpty(attachment.ContentBytes))
                    {
                        var datePrefix = DateTime.Now.ToString("yyyy-MM-dd");
                        var fileName = $"{datePrefix}_{attachment.Name}";
                        var filePath = Path.Combine(AppSettings.DownloadsDirectory, fileName);

                        var bytes = Convert.FromBase64String(attachment.ContentBytes);
                        File.WriteAllBytes(filePath, bytes);

                        MarkAttachmentProcessed(email.Id, attachment.Id, fileName);

                        log.Info($"    ✅ Downloaded: {fileName} ({bytes.Length / 1024.0:F2} KB)");
                        downloaded++;
                    }
                    else
                    {
                        log.Warn($"    ⚠️  Skipped: {attachment.Name} (no content available)");
                    }
                }

                return (attachments.Count, downloaded, skipped);
            }
            catch (Exception ex)
            {
                log.Error($"  Error processing attachments: {ex.Message}", ex);
                return (0, 0, 0);
            }
        }

        private TrackingData LoadTrackingData()
        {
            try
            {
                if (File.Exists(AppSettings.TrackingFile))
                {
                    var json = File.ReadAllText(AppSettings.TrackingFile);
                    if (string.IsNullOrWhiteSpace(json))
                    {
                        return new TrackingData();
                    }

                    var data = JsonConvert.DeserializeObject<TrackingData>(json);
                    log.Info($"Loaded tracking data: {data.Attachments.Count} attachments previously processed");
                    return data;
                }
            }
            catch (Exception ex)
            {
                log.Warn($"Failed to load tracking file, using empty state: {ex.Message}");
            }

            return new TrackingData();
        }

        private void SaveTrackingData()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_trackingData, Formatting.Indented);
                File.WriteAllText(AppSettings.TrackingFile, json);
                log.Info("Tracking data saved successfully");
            }
            catch (Exception ex)
            {
                log.Error($"Failed to save tracking file: {ex.Message}", ex);
            }
        }

        */

        // File: `Services/EmailProcessor.cs`
// Language: csharp
        private bool IsAttachmentProcessed(string emailId, string attachmentId)
        {
            var key = $"{emailId}_{attachmentId}";
            return _trackingData.Attachments.ContainsKey(key);
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
        }

        public async Task<ProcessingResult> ProcessEmailsLegacy()
        {
            var result = new ProcessingResult();
            try
            {
                log.Info("Starting legacy-style email processing (Graph SDK)");

                // Validate paths
                if (!Directory.Exists(AppSettings.OutputPathGood))
                {
                    log.Error($"Output path does not exist: {AppSettings.OutputPathGood}");
                    return result;
                }

                var logDir = Path.GetDirectoryName(AppSettings.LogFile) ?? AppSettings.LogFile;
                if (!Directory.Exists(logDir))
                {
                    log.Error($"Log directory does not exist: {logDir}");
                    return result;
                }

                // Load allowed senders/domains and blocked extensions
                var allowed = LoadAllowedSenders();
                log.Info($"Loaded {allowed.Count} allowed senders/domains");
                var blockedExts = ParseBlockedExtensions();

                // Resolve import and old folder ids
                var importFolderId = GetFolderIdByDisplayName(AppSettings.InboxImportSubDir)
                                     ?? GetFolderIdByDisplayName("Inbox");
                if (string.IsNullOrEmpty(importFolderId))
                {
                    log.Error("Could not resolve import folder");
                    return result;
                }

                var oldFolderId = GetFolderIdByDisplayName(AppSettings.InboxOldSubDir) ?? importFolderId;

                // Build filter for timeframe and attachments
                var hoursAgo = DateTime.UtcNow.AddHours(-AppSettings.HoursToFetch);
                var filter =
                    $"receivedDateTime ge {hoursAgo.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")} and hasAttachments eq true";

                // Query messages in the import folder via Graph SDK
                var messagesResp = await _graphServiceClient
                    .Users[AppSettings.TargetEmail]
                    .MailFolders[importFolderId]
                    .Messages
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Filter = filter;
                        config.QueryParameters.Select = new[]
                            { "id", "subject", "from", "receivedDateTime", "hasAttachments", "parentFolderId" };
                        config.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                        config.QueryParameters.Top = 999;
                    });

                var messages = messagesResp?.Value?.ToList() ?? new List<Message>();
                log.Info($"Found {messages.Count} messages with attachments in import folder");

                if (messages.Count == 0) return result;

                // Legacy: do not filter out sent items or same-mailbox messages (match EmailProcessorLegacy)
                var incoming = messages.ToList();
                log.Info($"Processing {incoming.Count} messages");
                result.EmailsProcessed = incoming.Count;

                var processedMessageIds = new List<string>();

                foreach (var msg in incoming)
                {
                    var sender = msg.From?.EmailAddress?.Address;
                    if (string.IsNullOrEmpty(sender) || !sender.Contains("@"))
                        continue;

                    var senderLower = sender.ToLowerInvariant();
                    var domain = senderLower.Substring(senderLower.LastIndexOf('@')); // includes '@'

                    if (!allowed.Contains(senderLower) && !allowed.Contains(domain))
                        continue;

                    log.Info(
                        $"\n--- Processing message: {msg.Subject} From: {sender} Date: {msg.ReceivedDateTime:yyyy-MM-dd HH:mm:ss}");

                    // Get attachments via Graph SDK
                    var attsResp = await _graphServiceClient
                        .Users[AppSettings.TargetEmail]
                        .Messages[msg.Id]
                        .Attachments
                        .GetAsync();

                    var attachments = attsResp?.Value ?? new List<Attachment>();
                    if (attachments.Count == 0)
                    {
                        log.Info("  No attachments found");
                        continue;
                    }

                    // Ensure output dir exists
                    if (!Directory.Exists(AppSettings.OutputPathGood))
                        Directory.CreateDirectory(AppSettings.OutputPathGood);

                    foreach (var att in attachments)
                    {
                        // Only handle file attachments (skip item/other types)
                        var fileAtt = att as FileAttachment;
                        if (fileAtt == null || fileAtt.ContentBytes == null)
                        {
                            log.Warn($"    ⚠️  Skipped (not a file or no content): {att.Name}");
                            result.SkippedAttachments++;
                            continue;
                        }

                        var fileName = att.Name ?? "attachment";
                        var ext = Path.GetExtension(fileName).ToLowerInvariant();

                        if (!blockedExts.Contains(ext))
                        {
                            // handle filename collisions
                            var targetPath = Path.Combine(AppSettings.OutputPathGood, fileName);
                            if (File.Exists(targetPath))
                            {
                                var prefix = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_";
                                fileName = prefix + fileName;
                                targetPath = Path.Combine(AppSettings.OutputPathGood, fileName);
                            }

                            File.WriteAllBytes(targetPath, fileAtt.ContentBytes);

                            if (AppSettings.LogAttachments)
                            {
                                var attachmentsLogPath = Path.Combine(logDir, AppSettings.LogFileAttachments);
                                File.AppendAllText(attachmentsLogPath,
                                    $"{sender} {fileName} {DateTime.Now}{Environment.NewLine}");
                            }

                            MarkAttachmentProcessed(msg.Id, att.Id, fileName);
                            result.TotalAttachments++;
                            result.NewDownloads++;
                            log.Info($"    ✅ Downloaded: {fileName} ({fileAtt.ContentBytes.Length / 1024.0:F2} KB)");
                        }
                        else
                        {
                            // legacy: do not log jpg/png as blocked
                            if (!ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) &&
                                !ext.Equals(".png", StringComparison.OrdinalIgnoreCase))
                            {
                                var blockedLogPath = Path.Combine(logDir, AppSettings.LogFileBlockedFiles);
                                File.AppendAllText(blockedLogPath,
                                    $"{sender} {fileName} {DateTime.Now}{Environment.NewLine}");
                            }

                            result.SkippedAttachments++;
                            log.Info($"    ⛔ Blocked by extension: {fileName}");
                        }
                    }

                    processedMessageIds.Add(msg.Id);
                }

                // Move processed messages to old folder
                foreach (var messageId in processedMessageIds)
                {
                    try
                    {
                        var moveBody = new MovePostRequestBody { DestinationId = oldFolderId };
                        _graphServiceClient.Users[AppSettings.TargetEmail].Messages[messageId].Move.PostAsync(moveBody)
                            .GetAwaiter().GetResult();
                        log.Info($"Moved message {messageId} to folder id {oldFolderId}");
                    }
                    catch (Exception ex)
                    {
                        log.Warn($"Failed to move message {messageId}: {ex.Message}");
                    }
                }

                log.Info("Legacy-style processing completed");
            }
            catch (Exception ex)
            {
                log.Error("Legacy-style email processing failed", ex);
                throw;
            }

            return result;
        }


        private HashSet<string> LoadAllowedSenders()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Emails.xml");
                if (!File.Exists(path))
                {
                    log.Warn($"Emails.xml not found at {path}");
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
                log.Warn($"Failed to load Emails.xml: {ex.Message}");
            }

            return set;
        }

        private HashSet<string> ParseBlockedExtensions()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var raw = AppSettings.BlockedFiles ?? string.Empty;
                var parts = raw.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
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
            if (string.IsNullOrWhiteSpace(displayName)) return null;
            try
            {
                var foldersResp = _graphServiceClient.Users[AppSettings.TargetEmail].MailFolders.GetAsync().GetAwaiter()
                    .GetResult();
                var folders = foldersResp?.Value ?? new List<MailFolder>();

                var found = folders.FirstOrDefault(f =>
                    string.Equals(f.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));
                if (found != null) return found.Id;

                var inbox = folders.FirstOrDefault(f =>
                    string.Equals(f.DisplayName, "Inbox", StringComparison.OrdinalIgnoreCase));
                if (inbox != null)
                {
                    var childResp = _graphServiceClient.Users[AppSettings.TargetEmail].MailFolders[inbox.Id]
                        .ChildFolders.GetAsync().GetAwaiter().GetResult();
                    var children = childResp?.Value ?? new List<MailFolder>();
                    var sub = children.FirstOrDefault(f =>
                        string.Equals(f.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));
                    if (sub != null) return sub.Id;
                }
            }
            catch (Exception ex)
            {
                log.Warn($"GetFolderIdByDisplayName failed for '{displayName}': {ex.Message}");
            }

            return null;
        }

        public class SimpleAuthenticationProvider : IAuthenticationProvider
        {
            private readonly TokenManager _tokenManager;

            public SimpleAuthenticationProvider(TokenManager tokenManager)
            {
                _tokenManager = tokenManager;
            }

            public Task AuthenticateRequestAsync(RequestInformation request,
                Dictionary<string, object> additionalAuthenticationContext = null,
                CancellationToken cancellationToken = new CancellationToken())
            {
                var token = _tokenManager.GetAccessToken();
                if (!string.IsNullOrEmpty(token))
                {
                    var requestHeaders = new RequestHeaders();
                    requestHeaders.Add("Authorization", $"Bearer {token}");
                    request.AddHeaders(requestHeaders);
                }

                return Task.CompletedTask;
            }
        }
    }
}