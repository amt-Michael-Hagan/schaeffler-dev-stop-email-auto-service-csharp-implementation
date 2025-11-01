using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using EmailAutomationLegacy.Models;
using EmailAutomationLegacy.Services;
using NUnit.Framework;

namespace GraphApiClientTest
{
    [TestFixture]
    public class EmailProcessorTests
    {
        private string _tempDirectory;
        private MockGraphServiceClientWrapper _mockGraphClient;
        private EmailProcessor _emailProcessor;

        [SetUp]
        public void SetUp()
        {
            // Create temporary directory for test files
            _tempDirectory = Path.Combine(Path.GetTempPath(), "EmailProcessorTests_" + Guid.NewGuid().ToString("N").Substring(8));
            Directory.CreateDirectory(_tempDirectory);


            // Create mock dependencies
            _mockGraphClient = new MockGraphServiceClientWrapper();
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up temporary directory
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithEmptyEmailList_ShouldReturnEmptyResult()
        {
            // Arrange
            var emptyEmailList = new Dictionary<string, string>();
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emptyEmailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(0));
            Assert.That(result.TotalAttachments, Is.EqualTo(0));
            Assert.That(result.NewDownloads, Is.EqualTo(0));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithNoMessages_ShouldReturnEmptyResult()
        {
            // Arrange
            var emailList = new Dictionary<string, string> { { "test@example.com", "Test User" } };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Ensure no messages in the mock folder
            _mockGraphClient.ClearTestData();

            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);
            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(0));
            Assert.That(result.TotalAttachments, Is.EqualTo(0));
            Assert.That(result.NewDownloads, Is.EqualTo(0));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithValidEmailAndAttachment_ShouldProcessSuccessfully()
        {
            // Arrange
            var emailList = new Dictionary<string, string> { { "sender@trusted.com", "Trusted Sender" } };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Setup test message with attachment
            var message = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-123",
                "Test Email",
                "sender@trusted.com",
                true
            );

            var attachment = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-123",
                "document.pdf",
                Encoding.UTF8.GetBytes("Test PDF content")
            );
            
            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);

            _mockGraphClient.AddTestMessage("inbox-id", message);
            _mockGraphClient.AddTestAttachment("msg-123", attachment);


            
            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(1));
            Assert.That(result.TotalAttachments, Is.EqualTo(1));
            Assert.That(result.NewDownloads, Is.EqualTo(1));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));

            // Verify the file was "downloaded" (mocked)
            // Assert.That(_mockGraphClient.MovedMessages.Count, Is.EqualTo(1));
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithBlockedFileExtension_ShouldSkipAttachment()
        {
            // Arrange
            var emailList = new Dictionary<string, string> { { "sender@trusted.com", "Trusted Sender" } };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Note: BlockedFiles is configured in AppSettings and can't be changed in tests
            // The test will use the current configuration value

            // Setup test message with blocked attachment
            var message = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-456",
                "Test Email with Blocked File",
                "sender@trusted.com",
                true
            );

            var blockedAttachment = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-456",
                "virus.exe",
                Encoding.UTF8.GetBytes("Malicious content")
            );

            _mockGraphClient.AddTestMessage("inbox-id", message);
            _mockGraphClient.AddTestAttachment("msg-456", blockedAttachment);

            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);
            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(1));
            Assert.That(result.TotalAttachments, Is.EqualTo(0)); // Not counted as processed
            Assert.That(result.NewDownloads, Is.EqualTo(0));
            Assert.That(result.SkippedAttachments, Is.EqualTo(1)); // Blocked file
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithUnauthorizedSender_ShouldSkipMessage()
        {
            // Arrange
            var emailList = new Dictionary<string, string> { { "authorized@trusted.com", "Authorized User" } };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Setup test message from unauthorized sender
            var message = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-789",
                "Unauthorized Email",
                "unauthorized@spam.com",
                true
            );

            var attachment = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-789",
                "document.pdf",
                Encoding.UTF8.GetBytes("Content from unauthorized sender")
            );

            _mockGraphClient.AddTestMessage("inbox-id", message);
            _mockGraphClient.AddTestAttachment("msg-789", attachment);

            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);
            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(1)); // Message was processed but sender filtered
            Assert.That(result.TotalAttachments, Is.EqualTo(0)); // No attachments processed
            Assert.That(result.NewDownloads, Is.EqualTo(0));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithDomainMatch_ShouldProcessMessage()
        {
            // Arrange
            var emailList = new Dictionary<string, string> { { "@trusted.com", "Trusted Domain" } };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Setup test message from domain match
            var message = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-domain",
                "Domain Email",
                "anyone@trusted.com",
                true
            );

            var attachment = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-domain",
                "report.xlsx",
                Encoding.UTF8.GetBytes("Excel report content")
            );

            _mockGraphClient.AddTestMessage("inbox-id", message);
            _mockGraphClient.AddTestAttachment("msg-domain", attachment);

            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);
            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(1));
            Assert.That(result.TotalAttachments, Is.EqualTo(1));
            Assert.That(result.NewDownloads, Is.EqualTo(1));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));
        }

        [Test]
        public async Task ProcessEmailsWithGraphAsync_WithMultipleMessages_ShouldProcessAll()
        {
            // Arrange
            var emailList = new Dictionary<string, string>
            {
                { "user1@trusted.com", "User 1" },
                { "@company.com", "Company Domain" }
            };
            var logFilePath = Path.Combine(_tempDirectory, "test.log");

            // Setup multiple test messages
            var message1 = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-multi1",
                "First Email",
                "user1@trusted.com",
                true
            );
            
            var message2 = MockGraphServiceClientWrapper.CreateTestMessage(
                "msg-multi2",
                "Second Email",
                "user2@company.com",
                true
            );

            var attachment1 = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-multi1",
                "doc1.pdf",
                Encoding.UTF8.GetBytes("First document")
            );
            var attachment2 = MockGraphServiceClientWrapper.CreateTestFileAttachment(
                "att-multi2",
                "doc2.pdf",
                Encoding.UTF8.GetBytes("Second document")
            );

            _mockGraphClient.AddTestMessage("inbox-id", message1);
            _mockGraphClient.AddTestMessage("inbox-id", message2);
            _mockGraphClient.AddTestAttachment("msg-multi1", attachment1);
            _mockGraphClient.AddTestAttachment("msg-multi2", attachment2);
            var trackingData = new ProcessedEmailAttachmentTracker();
            _emailProcessor = new EmailProcessor(_mockGraphClient, trackingData);

            // Act
            var result = await _emailProcessor.ProcessEmailsWithGraphAsync(emailList, logFilePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.EmailsProcessed, Is.EqualTo(2));
            Assert.That(result.TotalAttachments, Is.EqualTo(2));
            Assert.That(result.NewDownloads, Is.EqualTo(2));
            Assert.That(result.SkippedAttachments, Is.EqualTo(0));

            // Verify both messages were moved
            // Assert.That(_mockGraphClient.MovedMessages.Count, Is.EqualTo(2));
        }
    }

}