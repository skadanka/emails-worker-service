using System.Net;
using Moq;
using Moq.Protected;
using Xunit;
using emails_worker_service.Controllers;
using emails_worker_service.Models.FormModel;
using emails_worker_service.Exceptions;

public class FormSubmitSalesForceTests : IDisposable
{
    private readonly Mock<IHttpClientFactory> _mockHttpClientFactory;
    private readonly HttpClient _httpClient;
    private readonly Mock<HttpMessageHandler> _mockHttpMessageHandler;
    private readonly FormSubmitSalesForce _controller;
    private readonly string _mockFileDirectory;

    public FormSubmitSalesForceTests()
    {
        _mockHttpClientFactory = new Mock<IHttpClientFactory>();
        _mockHttpMessageHandler = new Mock<HttpMessageHandler>(MockBehavior.Strict);

        _httpClient = new HttpClient(_mockHttpMessageHandler.Object);
        _mockHttpClientFactory.Setup(x => x.CreateClient(It.IsAny<string>())).Returns(_httpClient);

        _controller = new FormSubmitSalesForce(_mockHttpClientFactory.Object);
        _mockFileDirectory = Path.Combine(Path.GetTempPath(), "MockFiles");

        // Ensure mock file directory is created
        Directory.CreateDirectory(_mockFileDirectory);

        // Set up Dispose method for HttpMessageHandler
        _mockHttpMessageHandler.Protected().Setup("Dispose", ItExpr.IsAny<bool>());
    }

    private string CreateMockFile(string content, string fileName)
    {
        string filePath = Path.Combine(_mockFileDirectory, fileName);

        File.WriteAllText(filePath, content);

        return filePath;
    }

    [Fact]
    public async Task SubmitForm_NullModel_ThrowsFormModelNullException()
    {
        // Act & Assert
        await Assert.ThrowsAsync<FormModelNullException>(() => _controller.SubmitForm(null));
    }

    [Fact]
    public async Task SubmitForm_MissingRequiredFields_ThrowsMissingRequiredFieldsException()
    {
        // Arrange
        var model = new FormModelDrushim
        {
            TenderType = "Event Planning"
            // Missing other required fields
        };

        // Act & Assert
        await Assert.ThrowsAsync<MissingRequiredFieldsException>(() => _controller.SubmitForm(model));
    }

    [Fact]
    public async Task SubmitForm_MissingResumeFile_ThrowsResumeFileNotFoundException()
    {
        // Arrange
        var model = new FormModelDrushim
        {
            TenderType = "Event Planning",
            JobNumber = "123",
            SubmissionDate = "2021-12-31",
            SubmissionTime = "12:00",
            FirstName = "John",
            LastName = "Doe",
            Phone = "1234567890",
            Email = "john.doe@example.com",
            Exposure = "Internet",
            Resume = Path.Combine(_mockFileDirectory, "nonexistentfile.pdf")
        };

        // Act & Assert
        await Assert.ThrowsAsync<ResumeFileNotFoundException>(() => _controller.SubmitForm(model));
    }

    [Fact]
    public async Task SubmitForm_FailedToRetrieveFormHtml_ThrowsFormHtmlRetrievalException()
    {
        // Arrange
        var model = new FormModelDrushim
        {
            TenderType = "Event Planning",
            JobNumber = "123",
            SubmissionDate = "2021-12-31",
            SubmissionTime = "12:00",
            FirstName = "John",
            LastName = "Doe",
            Phone = "1234567890",
            Email = "john.doe@example.com",
            Exposure = "Internet",
            Resume = CreateMockFile("Mock resume content", "resume.pdf")
        };

        var httpResponse = new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent("Page not found")
        };

        _mockHttpMessageHandler.Protected()
            .Setup<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>()
            )
            .ReturnsAsync(httpResponse);

        // Act & Assert
        await Assert.ThrowsAsync<FormHtmlRetrievalException>(() => _controller.SubmitForm(model));
    }

    [Fact]
    public async Task SubmitForm_FormNotFoundInHtml_ThrowsFormNotFoundException()
    {
        // Arrange
        var model = new FormModelDrushim
        {
            TenderType = "Event Planning",
            JobNumber = "123",
            SubmissionDate = "2021-12-31",
            SubmissionTime = "12:00",
            FirstName = "John",
            LastName = "Doe",
            Phone = "1234567890",
            Email = "john.doe@example.com",
            Exposure = "Internet",
            Resume = CreateMockFile("Mock resume content", "resume.pdf")
        };

        var formHtml = "<html><body><div>No form here</div></body></html>";
        var httpResponse = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(formHtml)
        };

        _mockHttpMessageHandler.Protected()
            .Setup<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>()
            )
            .ReturnsAsync(httpResponse);

        // Act & Assert
        await Assert.ThrowsAsync<FormNotFoundException>(() => _controller.SubmitForm(model));
    }

    [Fact]
    public async Task SubmitForm_SuccessfulSubmission_ReturnsSuccessResponse()
    {
        // Arrange
        var model = new FormModelDrushim
        {
            TenderType = "Event Planning",
            JobNumber = "123",
            SubmissionDate = "2021-12-31",
            SubmissionTime = "12:00",
            FirstName = "John",
            LastName = "Doe",
            Phone = "1234567890",
            Email = "john.doe@example.com",
            Exposure = "Internet",
            Resume = CreateMockFile("Mock resume content", "resume.pdf")
        };

        var formHtml = "<html><body><form id='5116778'><input type='hidden' name='hiddenField1' value='hiddenValue1' /></form></body></html>";
        var httpResponseGet = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(formHtml)
        };
        var httpResponsePost = new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent("Success")
        };

        _mockHttpMessageHandler.Protected()
            .SetupSequence<Task<HttpResponseMessage>>(
                "SendAsync",
                ItExpr.IsAny<HttpRequestMessage>(),
                ItExpr.IsAny<CancellationToken>()
            )
            .ReturnsAsync(httpResponseGet)
            .ReturnsAsync(httpResponsePost);

        // Act
        var result = await _controller.SubmitForm(model);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal("Form submitted successfully.", result.Message);
        Assert.Equal("Success", result.ResponseContent);
    }

    // Cleanup created files and directories
    public void Dispose()
    {
        if (Directory.Exists(_mockFileDirectory))
        {
            Directory.Delete(_mockFileDirectory, true);
        }
    }
}
