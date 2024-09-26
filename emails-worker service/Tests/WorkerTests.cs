using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;
using emails_worker_service.Controllers;
using emails_worker_service.Models;
using emails_worker_service.Pdf;
using System.IO.Abstractions;
using emails_worker_service.Controllers.FormCreator;
using System.IO.Abstractions.TestingHelpers;
using Microsoft.Office.Interop.Outlook;
using emails_worker_service.Models.FormModel;

namespace emails_worker_service.Tests
{
    public class WorkerTests
    {
        private readonly Mock<ILogger<Worker>> _mockLogger;
        private readonly Mock<IServiceProvider> _mockServiceProvider;
        private readonly Mock<IServiceScope> _mockServiceScope;
        private readonly Mock<IServiceScopeFactory> _mockServiceScopeFactory;
        private readonly Mock<IFormCreator> _mockFormCreatorOutlook;
        private readonly Mock<FormSubmitSalesForce> _mockFormSubmitSalesForce;
        private readonly Mock<IEmailService> _mockEmailService; // Added mock for IEmailService
        private readonly Mock<MAPIFolder> _mockInboxFolder; // Added mock for MAPIFolder
        private readonly Mock<Items> _mockMailItems; // Added mock for Items
        private readonly Worker _worker;

        private readonly MockFileSystem _mockFileSystem; // Mock file system
        private readonly FormModelServiceCsv _formModelServiceCsv; // Actual instance using mock file system

        public WorkerTests()
        {
            _mockLogger = new Mock<ILogger<Worker>>();
            _mockServiceProvider = new Mock<IServiceProvider>();
            _mockServiceScope = new Mock<IServiceScope>();
            _mockServiceScopeFactory = new Mock<IServiceScopeFactory>();

            // Initialize the mock file system
            _mockFileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { "someFilePath.csv", new MockFileData("") } // Creating an empty file in the mock file system
            });

            // Initialize the actual FormModelServiceCsv with the mock file system
            _formModelServiceCsv = new FormModelServiceCsv("someFilePath.csv", _mockFileSystem);

            // Mocking FormCreatorOutlook using the IFormCreator interface
            _mockFormCreatorOutlook = new Mock<IFormCreator>();
            _mockFormSubmitSalesForce = new Mock<FormSubmitSalesForce>(null); // Assuming it takes an IHttpClientFactory

            // Added mock for IEmailService
            _mockEmailService = new Mock<IEmailService>();

            // Added mock for MAPIFolder (inbox)
            _mockInboxFolder = new Mock<MAPIFolder>();

            // Added mock for Items (collection of MailItems)
            _mockMailItems = new Mock<Items>();

            // Setup the service provider to return the mock scope factory
            _mockServiceProvider.Setup(x => x.GetService(typeof(IServiceScopeFactory)))
                .Returns(_mockServiceScopeFactory.Object);

            // Setup the service scope to return the service provider
            _mockServiceScope.Setup(x => x.ServiceProvider).Returns(_mockServiceProvider.Object);
            _mockServiceScopeFactory.Setup(x => x.CreateScope()).Returns(_mockServiceScope.Object);

            // Setup the services that the worker depends on
            _mockServiceProvider.Setup(x => x.GetService(typeof(IFormCreator)))
                .Returns(_mockFormCreatorOutlook.Object);
            _mockServiceProvider.Setup(x => x.GetService(typeof(FormModelServiceCsv)))
                .Returns(_formModelServiceCsv); // Returning the actual instance here
            _mockServiceProvider.Setup(x => x.GetService(typeof(FormSubmitSalesForce)))
                .Returns(_mockFormSubmitSalesForce.Object);
            _mockServiceProvider.Setup(x => x.GetService(typeof(IEmailService)))
                .Returns(_mockEmailService.Object); // Return the IEmailService mock

            // Setup IEmailService to return the mock inbox folder
            _mockEmailService.Setup(x => x.GetInbox()).Returns(_mockInboxFolder.Object);

            // Setup the inbox folder to return the mock Items collection
            _mockInboxFolder.Setup(x => x.Items).Returns(_mockMailItems.Object);

            _worker = new Worker(_mockLogger.Object, _mockServiceProvider.Object, _mockEmailService.Object);
        }

        [Fact]
        public async Task ExecuteAsync_ProcessesFormsSuccessfully()
        {
            // Arrange
            var cancellationTokenSource = new CancellationTokenSource();

            // Create mock MailItems
            var mockMailItem = new Mock<MailItem>();
            mockMailItem.Setup(m => m.Subject).Returns("Test Email");
            var mailItemsList = new List<MailItem> { mockMailItem.Object };

            // Setup the mock Items collection to return the mail items
            _mockMailItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

            var formModels = new Dictionary<string, object>
            {
                { "123", new FormModelDrushim { FirstName = "John", LastName = "Doe" } }
            };
            var formRecords = new Dictionary<string, string>
            {
                { "111", "SomeRecord" }
            };
            var formSubmissionResponse = new FormSubmissionResponse { IsSuccess = true, Message = "Form submitted successfully." };

            _mockFormCreatorOutlook.Setup(x => x.GetForms(_mockMailItems.Object))
                .Returns(formModels);
            _mockFormSubmitSalesForce.Setup(x => x.SubmitForm(It.IsAny<FormModelBase>()))
                .ReturnsAsync(formSubmissionResponse);

            // Act
            await _worker.StartAsync(cancellationTokenSource.Token);
            cancellationTokenSource.Cancel(); // Cancel after the first run

            // Assert
            _mockFormCreatorOutlook.Verify(x => x.GetForms(_mockMailItems.Object), Times.Once);
            _mockFormSubmitSalesForce.Verify(x => x.SubmitForm(It.IsAny<FormModelBase>()), Times.Once);
            _mockLogger.Verify(x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Information),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("Form submitted successfully.")),
                It.IsAny<System.Exception>(),
                It.Is<Func<It.IsAnyType, System.Exception, string>>((v, t) => true)));
        }

        [Fact]
        public async Task ExecuteAsync_HandlesUnsupportedEmailSource()
        {
            // Arrange
            var cancellationTokenSource = new CancellationTokenSource();

            // Create mock MailItems
            var mockMailItem = new Mock<MailItem>();
            mockMailItem.Setup(m => m.Subject).Returns("Test Unsupported Email");
            var mailItemsList = new List<MailItem> { mockMailItem.Object };

            // Setup the mock Items collection to return the mail items
            _mockMailItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

            var formModels = new Dictionary<string, object>
            {
                { "124", "Unsupported email source or format." }
            };

            _mockFormCreatorOutlook.Setup(x => x.GetForms(_mockMailItems.Object))
                .Returns(formModels);

            // Act
            await _worker.StartAsync(cancellationTokenSource.Token);
            cancellationTokenSource.Cancel(); // Cancel after the first run

            // Assert
            _mockFormCreatorOutlook.Verify(x => x.GetForms(_mockMailItems.Object), Times.Once);
            _mockLogger.Verify(x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Error),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("Unsupported email source or format.")),
                It.IsAny<System.Exception>(),
                It.Is<Func<It.IsAnyType, System.Exception, string>>((v, t) => true)));
        }

        [Fact]
        public async Task ExecuteAsync_HandlesPdfExtractionError()
        {
            // Arrange
            var cancellationTokenSource = new CancellationTokenSource();

            // Create mock MailItems
            var mockMailItem = new Mock<MailItem>();
            mockMailItem.Setup(m => m.Subject).Returns("Test Email with PDF Extraction Error");
            var mailItemsList = new List<MailItem> { mockMailItem.Object };

            // Setup the mock Items collection to return the mail items
            _mockMailItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

            var formModels = new Dictionary<string, object>
            {
                { "125", "Error extracting content from the resume." }
            };

            _mockFormCreatorOutlook.Setup(x => x.GetForms(_mockMailItems.Object))
                .Returns(formModels);

            // Act
            await _worker.StartAsync(cancellationTokenSource.Token);
            cancellationTokenSource.Cancel(); // Cancel after the first run

            // Assert
            _mockFormCreatorOutlook.Verify(x => x.GetForms(_mockMailItems.Object), Times.Once);
            _mockLogger.Verify(x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Error),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("Error extracting content from the resume.")),
                It.IsAny<System.Exception>(),
                It.Is<Func<It.IsAnyType, System.Exception, string>>((v, t) => true)));
        }

        [Fact]
        public async Task ExecuteAsync_LogsErrorIfProcessingFails()
        {
            // Arrange
            var cancellationTokenSource = new CancellationTokenSource();

            // Create mock MailItems
            var mockMailItem = new Mock<MailItem>();
            mockMailItem.Setup(m => m.Subject).Returns("Test Email Processing Failure");
            var mailItemsList = new List<MailItem> { mockMailItem.Object };

            // Setup the mock Items collection to return the mail items
            _mockMailItems.Setup(m => m.GetEnumerator()).Returns(mailItemsList.GetEnumerator());

            _mockFormCreatorOutlook.Setup(x => x.GetForms(_mockMailItems.Object))
                .Throws(new System.Exception("Test exception"));

            // Act
            await _worker.StartAsync(cancellationTokenSource.Token);
            cancellationTokenSource.Cancel(); // Cancel after the first run

            // Assert
            _mockLogger.Verify(x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Error),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("An error occurred during processing.")),
                It.IsAny<System.Exception>(),
                It.Is<Func<It.IsAnyType, System.Exception, string>>((v, t) => true)));
        }
    }
}
