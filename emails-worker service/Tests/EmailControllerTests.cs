using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Moq;
using Xunit;
using emails_worker_service.Controllers;
using emails_worker_service.Models;
using emails_worker_service.Pdf;
using emails_worker_service.Exceptions;
using emails_worker_service.Exceptions.emails_worker_service.Exceptions;

public class EmailControllerTests : IDisposable
{
    private readonly Mock<PdfReaderComponent> _mockPdfReaderComponent;
    private readonly EmailController _controller;
    private readonly Mock<Application> _mockOutlookApp;
    private readonly Mock<MAPIFolder> _mockInbox;
    private readonly Mock<Items> _mockItems;
    private readonly string _mockFileDirectory;

    public EmailControllerTests()
    {
        _mockPdfReaderComponent = new Mock<PdfReaderComponent>();
        _controller = new EmailController(_mockPdfReaderComponent.Object);

        _mockOutlookApp = new Mock<Application>();
        _mockInbox = new Mock<MAPIFolder>();
        _mockItems = new Mock<Items>();

        // Set up the mock Outlook application to return the mock inbox
        _mockOutlookApp
            .Setup(app => app.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox))
            .Returns(_mockInbox.Object);

        // Set up the mock inbox to return the mock items
        _mockInbox
            .Setup(inbox => inbox.Items)
            .Returns(_mockItems.Object);

        _mockFileDirectory = Path.Combine(Path.GetTempPath(), "MockFiles");

        // Ensure mock file directory is created
        Directory.CreateDirectory(_mockFileDirectory);
    }

    private string CreateMockFile(string content, string fileName)
    {
        string filePath = Path.Combine(_mockFileDirectory, fileName);
        File.WriteAllText(filePath, content);
        return filePath;
    }

    [Fact]
    public void GetEmails_SuccessfulProcessing_ReturnsExpectedFormModel()
    {
        // Arrange
        var mockMailItem = new Mock<MailItem>();
        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body with Drushim content");
        mockMailItem.Setup(m => m.EntryID).Returns("123");
        mockMailItem.Setup(m => m.Attachments).Returns(Mock.Of<Attachments>());

        var mockAttachment = new Mock<Attachment>();
        mockAttachment.Setup(a => a.FileName).Returns("resume.pdf");
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a.Count).Returns(1);
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a[1]).Returns(mockAttachment.Object);

        _mockItems.Setup(items => items.GetEnumerator()).Returns(new List<MailItem> { mockMailItem.Object }.GetEnumerator());

        _mockPdfReaderComponent.Setup(reader => reader.ReadPdfAndExtractText(It.IsAny<string>()))
                               .Returns(new List<string> { "Extracted content from resume" });

        // Act
        var result = _controller.GetEmails();

        // Assert
        Assert.NotNull(result);
        Assert.Single(result);

        Assert.True(result.ContainsKey("123"), "Result should contain the expected EntryID as the key.");

        var formModel = result["123"] as FormModelBase;
        Assert.NotNull(formModel);
        Assert.Equal("Drushim", formModel.Exposure);
        Assert.Equal("Extracted content from resume", formModel.CvContent);
    }

    [Fact]
    public void GetEmails_UnsupportedEmailSource_ReturnsErrorMessage()
    {
        // Arrange
        var mockMailItem = new Mock<MailItem>();
        mockMailItem.Setup(m => m.HTMLBody).Returns("Unsupported content");
        mockMailItem.Setup(m => m.EntryID).Returns("124");

        _mockItems.Setup(items => items.GetEnumerator()).Returns(new List<MailItem> { mockMailItem.Object }.GetEnumerator());

        // Act
        var result = _controller.GetEmails();

        // Assert
        Assert.NotNull(result);
        Assert.Single(result);

        Assert.True(result.ContainsKey("124"), "Result should contain the EntryID as the key.");

        var errorMessage = result["124"] as string;
        Assert.NotNull(errorMessage);
        Assert.Equal("Unsupported email source or format.", errorMessage);
    }

    [Fact]
    public void GetEmails_ResumeExtractionFails_ReturnsErrorInCvContent()
    {
        // Arrange
        var mockMailItem = new Mock<MailItem>();
        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body with LinkedIn content");
        mockMailItem.Setup(m => m.EntryID).Returns("125");
        mockMailItem.Setup(m => m.Attachments).Returns(Mock.Of<Attachments>());

        var mockAttachment = new Mock<Attachment>();
        mockAttachment.Setup(a => a.FileName).Returns("resume.pdf");
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a.Count).Returns(1);
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a[1]).Returns(mockAttachment.Object);

        _mockItems.Setup(items => items.GetEnumerator()).Returns(new List<MailItem> { mockMailItem.Object }.GetEnumerator());

        _mockPdfReaderComponent.Setup(reader => reader.ReadPdfAndExtractText(It.IsAny<string>()))
                               .Throws(new PdfExtractionException("Error reading PDF", new System.Exception("Inner exception")));

        // Act
        var result = _controller.GetEmails();

        // Assert
        Assert.NotNull(result);
        Assert.Single(result);

        Assert.True(result.ContainsKey("125"), "Result should contain the EntryID as the key.");

        var formModel = result["125"] as FormModelBase;
        Assert.NotNull(formModel);
        Assert.Equal("Error extracting content from the resume: Error reading PDF", formModel.CvContent);
    }

    [Fact]
    public void GetEmails_UnsupportedFileType_SkipsAttachmentProcessing()
    {
        // Arrange
        var mockMailItem = new Mock<MailItem>();
        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body with LinkedIn content");
        mockMailItem.Setup(m => m.EntryID).Returns("126");
        mockMailItem.Setup(m => m.Attachments).Returns(Mock.Of<Attachments>());

        var mockAttachment = new Mock<Attachment>();
        mockAttachment.Setup(a => a.FileName).Returns("unsupportedfile.txt");
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a.Count).Returns(1);
        Mock.Get(mockMailItem.Object.Attachments).Setup(a => a[1]).Returns(mockAttachment.Object);

        _mockItems.Setup(items => items.GetEnumerator()).Returns(new List<MailItem> { mockMailItem.Object }.GetEnumerator());

        // Act
        var result = _controller.GetEmails();

        // Assert
        Assert.NotNull(result);
        Assert.Single(result);

        Assert.True(result.ContainsKey("126"), "Result should contain the EntryID as the key.");

        var formModel = result["126"] as FormModelBase;
        Assert.NotNull(formModel);
        Assert.Equal("missing", formModel.CvContent);
    }

    [Fact]
    public void GetEmails_ExceptionDuringProcessing_CapturesAndReturnsErrorMessage()
    {
        // Arrange
        var mockMailItem = new Mock<MailItem>();
        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body");
        mockMailItem.Setup(m => m.EntryID).Returns("127");

        _mockItems.Setup(items => items.GetEnumerator()).Returns(new List<MailItem> { mockMailItem.Object }.GetEnumerator());

        // Simulate an error during processing
        _mockItems.Setup(items => items.GetEnumerator()).Throws(new System.Exception("Test exception"));

        // Act
        var result = _controller.GetEmails();

        // Assert
        Assert.NotNull(result);
        Assert.Single(result);

        Assert.True(result.ContainsKey("127"), "Result should contain the EntryID as the key.");

        var errorMessage = result["127"] as string;
        Assert.NotNull(errorMessage);
        Assert.Contains("Error processing email:", errorMessage);
    }

    public void Dispose()
    {
        if (Directory.Exists(_mockFileDirectory))
        {
            Directory.Delete(_mockFileDirectory, true);
        }
    }
}
