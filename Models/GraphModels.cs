using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace EmailAutomationLegacy.Models
{
    public class GraphResponse<T>
    {
        [JsonProperty("@odata.nextLink")]
        public string NextLink { get; set; }

        [JsonProperty("value")]
        public List<T> Value { get; set; }
    }

    public class GraphUser
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("mail")]
        public string Mail { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }
    }

    public class GraphMessage
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("subject")]
        public string Subject { get; set; }

        [JsonProperty("receivedDateTime")]
        public DateTime ReceivedDateTime { get; set; }

        [JsonProperty("hasAttachments")]
        public bool HasAttachments { get; set; }

        [JsonProperty("parentFolderId")]
        public string ParentFolderId { get; set; }

        [JsonProperty("from")]
        public GraphEmailAddress From { get; set; }

        [JsonProperty("sender")]
        public GraphEmailAddress Sender { get; set; }
    }

    public class GraphEmailAddress
    {
        [JsonProperty("emailAddress")]
        public GraphEmail EmailAddress { get; set; }
    }

    public class GraphEmail
    {
        [JsonProperty("address")]
        public string Address { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
    }

    public class GraphAttachment
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [JsonProperty("size")]
        public int Size { get; set; }

        [JsonProperty("contentBytes")]
        public string ContentBytes { get; set; }

        [JsonProperty("@odata.type")]
        public string ODataType { get; set; }
    }

    public class GraphFolder
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }

    public class ProcessingResult
    {
        public int EmailsProcessed { get; set; }
        public int TotalAttachments { get; set; }
        public int NewDownloads { get; set; }
        public int SkippedAttachments { get; set; }
    }

    public class ProcessedEmailAttachmentTracker
    {
        public Dictionary<string, object> Emails { get; set; }
        public Dictionary<string, AttachmentInfo> Attachments { get; set; }

        public ProcessedEmailAttachmentTracker()
        {
            Emails = new Dictionary<string, object>();
            Attachments = new Dictionary<string, AttachmentInfo>();
        }
    }

    public class AttachmentInfo
    {
        public string FileName { get; set; }
        public DateTime DownloadedAt { get; set; }
        public string EmailId { get; set; }
    }
}