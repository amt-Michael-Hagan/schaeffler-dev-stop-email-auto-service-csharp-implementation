using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EmailAutomationLegacy.Services;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Move;

namespace GraphApiClientTest
{
    public class MockGraphServiceClientWrapper : IGraphServiceClient
    {
        // Test data - can be configured by tests
        public List<MailFolder> MockMailFolders { get; set; } = new List<MailFolder>();
        public Dictionary<string, List<Message>> MockMessages { get; set; } = new Dictionary<string, List<Message>>();
        public Dictionary<string, List<Attachment>> MockAttachments { get; set; } = new Dictionary<string, List<Microsoft.Graph.Models.Attachment>>();
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

            // Setup default messages
            MockMessages["import-id"] = new List<Message>();
        }


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
            var messages = MockMessages.TryGetValue(folderId, out var message)
                ? message
                : new List<Message>();

            return Task.FromResult(new MessageCollectionResponse
            {
                Value = messages
            });
        }

        public string GetFolderIdByDisplayName(string displayName)
        {
            var folder = MockMailFolders.FirstOrDefault(f =>
                string.Equals(f.DisplayName, displayName, StringComparison.OrdinalIgnoreCase));

            return folder?.Id;
        }

        public Task MoveProcessedMails(string messageId, MovePostRequestBody moveBody)
        {
            // Just track the move operation
            MovedMessages.Add($"{messageId}:{moveBody?.DestinationId ?? "no-destination"}");
            return Task.CompletedTask;
        }
    }
}