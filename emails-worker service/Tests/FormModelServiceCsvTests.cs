using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.IO.Abstractions.TestingHelpers;
using Xunit;
using emails_worker_service.Models;
using emails_worker_service.Models.FormModel;

public class FormModelServiceCsvTests
{
    private readonly MockFileSystem _mockFileSystem;
    private readonly FormModelServiceCsv _service;
    private readonly string _filePath = "formModels.csv";

    public FormModelServiceCsvTests()
    {
        // Initialize the mock file system with an empty file
        _mockFileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
        {
            { _filePath, new MockFileData("") }
        });

        // Create the service with the mock file system
        _service = new FormModelServiceCsv(_filePath, _mockFileSystem);
    }

    [Fact]
    public void SaveMailId_SavesNewMailId()
    {
        // Arrange
        string mailId = "test@example.com";

        // Act
        _service.SaveMailId(mailId);

        // Assert
        var fileContents = _mockFileSystem.File.ReadAllText(_filePath);
        Assert.Contains("test@example.com,Mail ID saved.", fileContents);
    }

    [Fact]
    public void SaveMailId_DoesNotDuplicateMailId()
    {
        // Arrange
        string mailId = "duplicate@example.com";
        _service.SaveMailId(mailId); // Save once

        // Act
        _service.SaveMailId(mailId); // Try to save again

        // Assert
        var fileContents = _mockFileSystem.File.ReadAllLines(_filePath);
        Assert.Single(fileContents); // Only one occurrence of the mail ID
    }

    [Fact]
    public void SaveBatchMailIds_SavesMultipleMailIds()
    {
        // Arrange
        var mailIdsToAdd = new Dictionary<string, object>
        {
            { "batch1@example.com", new FormModelDrushim { FirstName = "John", LastName = "Doe" } },
            { "batch2@example.com", "Some error occurred" }
        };

        // Act
        _service.SaveBatchMailIds(mailIdsToAdd);

        // Assert
        var fileContents = _mockFileSystem.File.ReadAllText(_filePath);
        Assert.Contains("batch1@example.com,Success: Form processed for John Doe", fileContents);
        Assert.Contains("batch2@example.com,Error: Some error occurred", fileContents);
    }

    [Fact]
    public void ExistsMailId_ReturnsTrueIfMailIdExists()
    {
        // Arrange
        var formList = new Dictionary<string, object>
        {
            { "exists@example.com", new FormModelDrushim { FirstName = "John", LastName = "Doe" } }
        };
        _service.SaveBatchMailIds(formList);

        // Act
        var exists = _service.ExistsMailId("exists@example.com");

        // Assert
        Assert.True(exists);
    }

    [Fact]
    public void ExistsMailId_ReturnsFalseIfMailIdDoesNotExist()
    {
        // Act
        var exists = _service.ExistsMailId("nonexistent@example.com");

        // Assert
        Assert.False(exists);
    }

    [Fact]
    public void RemoveMailId_RemovesMailId()
    {
        // Arrange
        var formList = new Dictionary<string, object>
        {
            { "remove@example.com", new FormModelDrushim { FirstName = "John", LastName = "Doe" } }
        };
        _service.SaveBatchMailIds(formList);

        // Act
        _service.RemoveMailId("remove@example.com");

        // Assert
        var fileContents = _mockFileSystem.File.ReadAllText(_filePath);
        Assert.DoesNotContain("remove@example.com", fileContents);
    }
}
