using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EmailAutomationLegacy.Services
{
    public interface IGraphServiceClient
    {
        Task<IList<Attachment>> GetAttachmentsAsync(string messageId);
        Task<MessageCollectionResponse> ReadEmailMessages(string folderId, string filter);
        string GetFolderIdByDisplayName(string displayName);
    }

    public class GraphClient : IGraphServiceClient
    {
        private readonly GraphServiceClient _graphServiceClient;

        public GraphClient(TokenManager tokenManager)
        {
            _graphServiceClient = tokenManager.GetGraphClient();
        }

        public async Task<IList<Attachment>> GetAttachmentsAsync(string messageId)
        {
            var response = await _graphServiceClient
                .Users[AppSettings.TargetEmail]
                .Messages[messageId]
                .Attachments
                .GetAsync();

            return response?.Value;
        }

        public string GetFolderIdByDisplayName(string displayName)
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
                Console.WriteLine($"GetFolderIdByDisplayName failed for '{displayName}': {ex.Message}");
            }

            return null;
        }

        public async Task<MessageCollectionResponse> ReadEmailMessages(string folderId, string filter)
        {
            return await _graphServiceClient
                 .Users[AppSettings.TargetEmail]
                 .MailFolders[folderId]
                 .Messages
                 .GetAsync(config =>
                 {
                     config.QueryParameters.Filter = filter;
                     config.QueryParameters.Select = new[]
                         { "id", "subject", "from", "receivedDateTime", "hasAttachments", "parentFolderId" };
                     config.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                     config.QueryParameters.Top = 999;
                 });
        }
    }
}
