using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Moq;
using Xunit;
using emails_worker_service.document;
using emails_worker_service.Controllers;
using emails_worker_service.Controllers.FormCreator;
using emails_worker_service.Models.FormModel;

public class FormCreatorOutlookTests
{
    private readonly Mock<IEmailService> _mockEmailService;
    private readonly Mock<DocumentReaderComponent> _mockPdfReader;
    private readonly Mock<FormProcessor> _mockFormProcessor;
    private readonly FormCreatorOutlook _formCreator;

    public FormCreatorOutlookTests()
    {
        _mockEmailService = new Mock<IEmailService>();
        _mockPdfReader = new Mock<DocumentReaderComponent>();
        _mockFormProcessor = new Mock<FormProcessor>(_mockPdfReader.Object);

        _formCreator = new FormCreatorOutlook(_mockPdfReader.Object);
    }

    [Fact]
    public void GetForms_ShouldProcessSupportedEmails()
    {
        // Arrange
        var mockInbox = new Mock<MAPIFolder>();
        var mockItems = new Mock<Items>();
        var mockMailItem = new Mock<MailItem>();

        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body with supported content");
        mockMailItem.Setup(m => m.EntryID).Returns("123");

        var mailItemsList = new List<MailItem> { mockMailItem.Object };
        mockItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

        mockInbox.Setup(i => i.Items).Returns(mockItems.Object);
        _mockEmailService.Setup(s => s.GetInbox()).Returns(mockInbox.Object);

        _mockFormProcessor.Setup(p => p.ProcessForm(mockMailItem.Object)).Returns(new FormModelDrushim());

        // Act
        var result = _formCreator.GetForms(mockItems.Object);

        // Assert
        Assert.Single(result);
        Assert.IsType<FormModelDrushim>(result["123"]);
    }

    [Fact]
    public void GetForms_ShouldHandleUnsupportedEmails()
    {
        // Arrange
        var mockInbox = new Mock<MAPIFolder>();
        var mockItems = new Mock<Items>();
        var mockMailItem = new Mock<MailItem>();

        mockMailItem.Setup(m => m.HTMLBody).Returns("Unsupported content");
        mockMailItem.Setup(m => m.EntryID).Returns("456");

        var mailItemsList = new List<MailItem> { mockMailItem.Object };
        mockItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

        mockInbox.Setup(i => i.Items).Returns(mockItems.Object);
        _mockEmailService.Setup(s => s.GetInbox()).Returns(mockInbox.Object);

        // Act
        var result = _formCreator.GetForms(mockItems.Object);

        // Assert
        Assert.Empty(result);  // Unsupported emails should be skipped
    }

    [Fact]
    public void GetForms_ShouldLimitProcessedEmails()
    {
        // Arrange
        var mockInbox = new Mock<MAPIFolder>();
        var mockItems = new Mock<Items>();

        var mockMailItem1 = new Mock<MailItem>();
        var mockMailItem2 = new Mock<MailItem>();

        mockMailItem1.Setup(m => m.HTMLBody).Returns("Sample HTML body with supported content");
        mockMailItem1.Setup(m => m.EntryID).Returns("123");

        mockMailItem2.Setup(m => m.HTMLBody).Returns("Sample HTML body with supported content");
        mockMailItem2.Setup(m => m.EntryID).Returns("456");

        var mailItemsList = new List<MailItem> { mockMailItem1.Object, mockMailItem2.Object };
        mockItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

        mockInbox.Setup(i => i.Items).Returns(mockItems.Object);
        _mockEmailService.Setup(s => s.GetInbox()).Returns(mockInbox.Object);

        _mockFormProcessor.Setup(p => p.ProcessForm(It.IsAny<MailItem>())).Returns(new FormModelDrushim());

        // Act
        var result = _formCreator.GetForms(mockItems.Object);

        // Assert
        Assert.Equal(2, result.Count); // Should process both emails
    }

    [Fact]
    public void GetForms_ShouldHandleProcessingErrors()
    {
        // Arrange
        var mockInbox = new Mock<MAPIFolder>();
        var mockItems = new Mock<Items>();
        var mockMailItem = new Mock<MailItem>();

        mockMailItem.Setup(m => m.HTMLBody).Returns("Sample HTML body with supported content");
        mockMailItem.Setup(m => m.EntryID).Returns("789");

        var mailItemsList = new List<MailItem> { mockMailItem.Object };
        mockItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

        mockInbox.Setup(i => i.Items).Returns(mockItems.Object);
        _mockEmailService.Setup(s => s.GetInbox()).Returns(mockInbox.Object);

        _mockFormProcessor.Setup(p => p.ProcessForm(It.IsAny<MailItem>())).Throws(new System.Exception("Test Exception"));

        // Act
        var result = _formCreator.GetForms(mockItems.Object);

        // Assert
        Assert.Single(result);
        Assert.Equal("Error processing email: Test Exception", result["789"]);
    }
}
