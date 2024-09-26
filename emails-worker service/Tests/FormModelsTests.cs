using System.Collections.Generic;
using System.Text.RegularExpressions;
using emails_worker_service.document;
using HtmlAgilityPack;
using iText.Kernel.Pdf;
using Microsoft.Office.Interop.Outlook;
using Moq;
using SharpCompress.Common;
using Xunit;
using emails_worker_service.Models;
using emails_worker_service.Models.FormModel;


namespace emails_worker_service.FormModel.Tests
{
    public class FormModelLinkedInTests
    {
        private readonly DocumentReaderComponent _pdfReader = new DocumentReaderComponent();
        private string LoadHtmlFromFile(string fileName)
        {
            var path = fileName;
            return File.ReadAllText(path);
        }

        [Fact]
        public void FillForm_FillsCorrectly()
        {

            // Arrange
            var mailItemMock = new Mock<MailItem>();
            mailItemMock.Setup(m => m.Subject).Returns("New application: Teaching &amp; Research Laboratory Engineer from John Smith");
            mailItemMock.Setup(m => m.CreationTime).Returns(new System.DateTime(2021, 12, 1, 14, 30, 0));
            mailItemMock.Setup(m => m.EntryID).Returns("entry-id-1");
            mailItemMock.Setup(m => m.HTMLBody).Returns(LoadHtmlFromFile(@"C:\Users\recruitment\source\repos\emails-worker service\emails-worker service\Tests\files\LinkedInHTML.html"));

            var filePath = @"C:\Users\recruitment\source\repos\emails-worker service\emails-worker service\Tests\files\johnsmith.pdf";
            var formModel = new FormModelLinkedIn();

            // Act
            formModel.FillForm(mailItemMock.Object);

            
            try
            {   
                
                List<string> extractedTexts = _pdfReader.ReadPdfAndExtractText(filePath);
                string text = string.Join(Environment.NewLine, extractedTexts);
                FillMissingValuesFromCV(text, formModel);
                formModel.CvContent = text;
            }
            catch (System.Exception ex)
            {
                formModel.CvContent = "Error extracting content from the resume: " + ex.Message;
                Console.WriteLine("Error extracting PDF content: " + ex.Message);
            }
            // Assert
            Assert.Equal("Teaching &amp; Research Laboratory Engineer from John Smith", formModel.TenderType);
            // Missing job number in all LinkedIn templates, add in the Future form HR job number to job description.
            Assert.Equal("0000-2024", formModel.JobNumber);
            Assert.Equal("01/12/2021", formModel.SubmissionDate);
            Assert.Equal("14:30", formModel.SubmissionTime);
            Assert.Equal("John", formModel.FirstName);
            Assert.Equal("Doe", formModel.LastName);
            Assert.Equal("0521234567", formModel.Phone);
            Assert.Equal("johnsmith@gmail.com", formModel.Email);
            Assert.Equal("tfa_7209", formModel.Exposure);
            Assert.Equal("entry-id-1", formModel.MailId);
        }

        private void FillMissingValuesFromCV(string text, FormModelBase formModel)
        {
            // Define regex patterns for email and phone number extraction
            const string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            const string phonePattern = @"0(5[0123456789])[^\D]{7}";

            // Try to extract and assign email if missing
            if ((formModel.Email == "missing" || formModel == null) && Regex.IsMatch(text, emailPattern, RegexOptions.IgnoreCase))
            {
                formModel.Email = Regex.Match(text, emailPattern, RegexOptions.IgnoreCase).Value;
            }

            // Try to extract and assign phone number if missing
            if ((formModel.Phone == "missing" || formModel.Phone == null ) && Regex.IsMatch(text, phonePattern))
            {
                formModel.Phone = Regex.Match(text, phonePattern).Value;
            }
        }
    }

    public class FormModelDrushimTests
    {
        private string LoadHtmlFromFile(string fileName)
        {
            var path = fileName;
            return File.ReadAllText(path);
        }
        [Fact]
        public void FillForm_FillsCorrectly()
        {
            // Arrange
            var mailItemMock = new Mock<MailItem>();
            mailItemMock.Setup(m => m.Subject).Returns("קו\"ח: מנהל/ת פרויקטים- אגף תכנון בינוי ואחזקה - 208-2023 | ג'ון סמית| באר שבע");
            mailItemMock.Setup(m => m.CreationTime).Returns(new System.DateTime(2021, 12, 1, 14, 30, 0));
            mailItemMock.Setup(m => m.EntryID).Returns("entry-id-2");
            mailItemMock.Setup(m => m.HTMLBody).Returns(LoadHtmlFromFile(@"C:\Users\recruitment\source\repos\emails-worker service\emails-worker service\Tests\files\DrushimHTML.html"));

            var formModel = new FormModelDrushim();

            // Act
            formModel.FillForm(mailItemMock.Object);

            // Assert
            Assert.Equal("מנהל/ת פרויקטים- אגף תכנון בינוי ואחזקה", formModel.TenderType);
            Assert.Equal("208-2023", formModel.JobNumber);
            Assert.Equal("01/12/2021", formModel.SubmissionDate);
            Assert.Equal("14:30", formModel.SubmissionTime);
            Assert.Equal("ג'ון", formModel.FirstName);
            Assert.Equal("סמית", formModel.LastName);
            Assert.Equal("0501234567", formModel.Phone);
            Assert.Equal("johnsmith@gmail.com", formModel.Email);
            Assert.Equal("tfa_7211", formModel.Exposure);
            Assert.Equal("entry-id-2", formModel.MailId);
        }
    }
}
