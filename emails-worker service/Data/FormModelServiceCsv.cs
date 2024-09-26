using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Abstractions;
using CsvHelper;
using CsvHelper.Configuration;
using emails_worker_service.Data;
using emails_worker_service.Models.FormModel;

public class FormModelServiceCsv : IFormModelService
{
    private readonly IFileSystem _fileSystem;
    private readonly string _filePath;

    public FormModelServiceCsv(string filePath, IFileSystem fileSystem)
    {
        _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        _filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        EnsureFileExists();
    }

    private void EnsureFileExists()
    {
        if (!_fileSystem.File.Exists(_filePath))
        {
            using (var stream = _fileSystem.File.Create(_filePath))
            {
                // Ensure the file is created and closed properly
            }
        }
    }

    public void SaveMailId(string mailId)
    {
        if (string.IsNullOrWhiteSpace(mailId))
        {
            throw new ArgumentException("MailId cannot be null or empty.", nameof(mailId));
        }

        var mailIds = LoadRecords();
        if (!mailIds.ContainsKey(mailId))
        {
            mailIds[mailId] = "Mail ID saved.";
            WriteRecords(mailIds);
        }
    }

    public void SaveBatchMailIds(Dictionary<string, object> mailIdsToAdd)
    {
        if (mailIdsToAdd == null || (mailIdsToAdd.Count == 0))
        {
            throw new ArgumentException("MailIdsToAdd cannot be null or empty.", nameof(mailIdsToAdd));
        }

        var existingRecords = LoadRecords();
        foreach (var entry in mailIdsToAdd)
        {
            if (!existingRecords.ContainsKey(entry.Key))
            {
                if (entry.Value is FormModelBase formModel)
                {
                    existingRecords[entry.Key] = $"Success: Form processed for {formModel.FirstName} {formModel.LastName}";
                }
                else if (entry.Value is string errorMessage)
                {
                    existingRecords[entry.Key] = $"Error: {errorMessage}";
                }
            }
        }

        WriteRecords(existingRecords);
    }

    public bool ExistsMailId(string mailId)
    {
        if (string.IsNullOrWhiteSpace(mailId))
        {
            throw new ArgumentException("MailId cannot be null or empty.", nameof(mailId));
        }

        var mailIds = LoadRecords();
        return mailIds.ContainsKey(mailId);
    }

    public void RemoveMailId(string mailId)
    {
        if (string.IsNullOrWhiteSpace(mailId))
        {
            throw new ArgumentException("MailId cannot be null or empty.", nameof(mailId));
        }

        var mailIds = LoadRecords();
        if (mailIds.Remove(mailId))
        {
            WriteRecords(mailIds);
        }
    }

    public Dictionary<string, string> LoadRecords()
    {
        if (!_fileSystem.File.Exists(_filePath) || string.IsNullOrWhiteSpace(_fileSystem.File.ReadAllText(_filePath)))
        {
            return new Dictionary<string, string>();
        }

        using (var reader = _fileSystem.File.OpenText(_filePath))
        using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = false // Assuming the file does not have a header
        }))
        {
            var records = new Dictionary<string, string>();
            while (csv.Read())
            {
                var key = csv.GetField<string>(0);
                var value = csv.GetField<string>(1);
                records[key] = value;
            }
            return records;
        }
    }

    private void WriteRecords(Dictionary<string, string> records)
    {
        using (var writer = _fileSystem.File.CreateText(_filePath))
        using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = false // Ensure no header is written
        }))
        {
            foreach (var record in records)
            {
                csv.WriteField(record.Key);
                csv.WriteField(record.Value);
                csv.NextRecord();
            }
        }
    }
}
