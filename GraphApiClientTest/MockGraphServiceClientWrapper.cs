using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EmailAutomationLegacy.Services;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Move;

namespace GraphApiClientTest
{
    public class MockGraphServiceClientWrapper : IGraphServiceClientWrapper, IGraphServiceClient
    {
        // Test data - can be configured by tests
        public List<MailFolder> MockMailFolders { get; set; } = new List<MailFolder>();
        public Dictionary<string, List<MailFolder>> MockChildFolders { get; set; } = new Dictionary<string, List<MailFolder>>();
        public Dictionary<string, List<Message>> MockMessages { get; set; } = new Dictionary<string, List<Message>>();
        public Dictionary<string, List<Microsoft.Graph.Models.Attachment>> MockAttachments { get; set; } = new Dictionary<string, List<Microsoft.Graph.Models.Attachment>>();
        public List<string> MovedMessages { get; set; } = new List<string>();

        public MockGraphServiceClientWrapper()
        {
            SetupDefaultTestData();
        }

        private void SetupDefaultTestData()
        {
            // Setup default folders
            MockMailFolders.Add(new MailFolder
            {
                Id = "inbox-id",
                DisplayName = "Inbox"
            });

            MockMailFolders.Add(new MailFolder
            {
                Id = "import-id",
                DisplayName = "Import"
            });

            // Setup child folder for Import under Inbox
            MockChildFolders["inbox-id"] = new List<MailFolder>
            {
                new MailFolder { Id = "import-id", DisplayName = "Import" },
                new MailFolder { Id = "old-id", DisplayName = "Old" }
            };

            // Setup default messages
            MockMessages["import-id"] = new List<Message>();
        }

        public async Task<MailFolderCollectionResponse> GetMailFoldersAsync(string userEmail)
        {
            await Task.Delay(1); // Simulate async call
            return new MailFolderCollectionResponse
            {
                Value = MockMailFolders
            };
        }

        public async Task<MailFolderCollectionResponse> GetChildFoldersAsync(string userEmail, string folderId)
        {
            await Task.Delay(1); // Simulate async call
            var childFolders = MockChildFolders.ContainsKey(folderId)
                ? MockChildFolders[folderId]
                : new List<MailFolder>();

            return new MailFolderCollectionResponse
            {
                Value = childFolders
            };
        }

        public async Task<MessageCollectionResponse> GetMessagesAsync(string userEmail, string folderId, string filter, string[] select, string[] orderBy, int top)
        {
            await Task.Delay(1); // Simulate async call
            var messages = MockMessages.ContainsKey(folderId)
                ? MockMessages[folderId]
                : new List<Message>();

            // Apply basic filtering if needed (simplified for testing)
            var filteredMessages = messages.Take(top).ToList();

            return new MessageCollectionResponse
            {
                Value = filteredMessages
            };
        }

        public async Task<AttachmentCollectionResponse> GetAttachmentsAsync(string userEmail, string messageId)
        {
            await Task.Delay(1); // Simulate async call
            var attachments = MockAttachments.ContainsKey(messageId)
                ? MockAttachments[messageId]
                : new List<Microsoft.Graph.Models.Attachment>();

            return new AttachmentCollectionResponse
            {
                Value = attachments
            };
        }

        public async Task MoveMessageAsync(string userEmail, string messageId, string destinationFolderId)
        {
            await Task.Delay(1); // Simulate async call
            MovedMessages.Add($"{messageId}:{destinationFolderId}");
        }

        // Helper methods for test setup

        public void AddTestMessage(string folderId, Message message)
        {
            if (!MockMessages.ContainsKey(folderId))
                MockMessages[folderId] = new List<Message>();

            MockMessages[folderId].Add(message);
        }

        public void AddTestAttachment(string messageId, Microsoft.Graph.Models.Attachment attachment)
        {
            if (!MockAttachments.ContainsKey(messageId))
                MockAttachments[messageId] = new List<Microsoft.Graph.Models.Attachment>();

            MockAttachments[messageId].Add(attachment);
        }

        public void ClearTestData()
        {
            MockMessages.Clear();
            MockAttachments.Clear();
            MovedMessages.Clear();
        }

        public static Message CreateTestMessage(string id, string subject, string senderEmail, bool hasAttachments = true)
        {
            return new Message
            {
                Id = id,
                Subject = subject,
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = senderEmail
                    }
                },
                ReceivedDateTime = DateTimeOffset.UtcNow.AddHours(-1),
                HasAttachments = hasAttachments,
                ParentFolderId = "import-id"
            };
        }

        public static FileAttachment CreateTestFileAttachment(string id, string fileName, byte[] content)
        {
            return new FileAttachment
            {
                Id = id,
                Name = fileName,
                ContentBytes = content,
                Size = content?.Length ?? 0
            };
        }

        // Implementation of IGraphServiceClient methods
        public Task<IList<Attachment>> GetAttachmentsAsync(string messageId)
        {
            var attachments = MockAttachments.ContainsKey(messageId)
                ? MockAttachments[messageId].Cast<Attachment>().ToList()
                : new List<Attachment>();

            return Task.FromResult<IList<Attachment>>(attachments);
        }

        public Task<MessageCollectionResponse> ReadEmailMessages(string folderId, string filter)
        {
            var messages = MockMessages.ContainsKey(folderId)
                ? MockMessages[folderId]
                : new List<Message>();

            // Apply basic filtering if needed (simplified for testing)
            var filteredMessages = messages.Where(m =>
                string.IsNullOrEmpty(filter) ||
                (m.Subject?.Contains(filter) ?? false) ||
                (m.From?.EmailAddress?.Address?.Contains(filter) ?? false)
            ).ToList();

            return Task.FromResult(new MessageCollectionResponse
            {
                Value = filteredMessages
            });
        }

        public string GetFolderIdByDisplayName(string displayName)
        {
            var folder = MockMailFolders.FirstOrDefault(f =>
                string.Equals(f.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));

            return folder?.Id ?? string.Empty;
        }

        public Task MoveProcessedMails(string messageId, MovePostRequestBody moveBody)
        {
            // Just track the move operation
            MovedMessages.Add($"{messageId}:{moveBody?.DestinationId ?? "no-destination"}");
            return Task.CompletedTask;
        }
    }

    public interface IGraphServiceClientWrapper
    {
        Task<MailFolderCollectionResponse> GetMailFoldersAsync(string userEmail);
        Task<MailFolderCollectionResponse> GetChildFoldersAsync(string userEmail, string folderId);
        Task<MessageCollectionResponse> GetMessagesAsync(string userEmail, string folderId, string filter, string[] select, string[] orderBy, int top);
        Task<AttachmentCollectionResponse> GetAttachmentsAsync(string userEmail, string messageId);
        Task MoveMessageAsync(string userEmail, string messageId, string destinationFolderId);
    }

}