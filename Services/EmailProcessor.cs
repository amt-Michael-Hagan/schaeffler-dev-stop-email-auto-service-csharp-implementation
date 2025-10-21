using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using Newtonsoft.Json;
using EmailAutomationLegacy.Models;

namespace EmailAutomationLegacy.Services
{
    public class EmailProcessor
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(EmailProcessor));
        
        private readonly TokenManager _tokenManager;
        private readonly GraphApiClient _graphClient;
        private TrackingData _trackingData;

        public EmailProcessor(TokenManager tokenManager)
        {
            _tokenManager = tokenManager;
            _graphClient = new GraphApiClient(_tokenManager);
            _trackingData = LoadTrackingData();
        }

        public bool TestConnection()
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
        }

        public ProcessingResult ProcessEmails()
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
    }
}