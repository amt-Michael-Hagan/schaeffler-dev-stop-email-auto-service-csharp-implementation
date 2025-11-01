using System;
using System.Collections.Generic;

namespace EmailAutomationLegacy.Models
{

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