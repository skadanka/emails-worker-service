using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using emails_worker_service.document;
using emails_worker_service.Exceptions;
using System.Text.RegularExpressions;
using emails_worker_service.Models.FormModel;

namespace emails_worker_service.Controllers
{
    public class AttachmentProcessor
    {
        private readonly DocumentReaderComponent _pdfReader;
        private readonly string[] _supportedFileTypes = { ".doc", ".pdf", ".docx" };

        public AttachmentProcessor(DocumentReaderComponent pdfReader)
        {
            _pdfReader = pdfReader;
        }

        /// <summary>
        /// Processes an attachment and extracts content if it's a supported file type.
        /// </summary>
        /// <param name="att">The attachment to process.</param>
        /// <param name="formModel">The form model to update with the extracted content.</param>
        /// <returns>True if the attachment was processed, otherwise false.</returns>
        public bool ProcessAttachment(Attachment att, FormModelBase formModel)
        {
            string ext = Path.GetExtension(att.FileName).ToLower();
            if (Array.Exists(_supportedFileTypes, fileType => fileType.Equals(ext)))
            {
                string filePath = Path.Combine(Path.GetTempPath(), att.FileName);
                att.SaveAsFile(filePath);

                try
                {
                    List<string> extractedTexts = _pdfReader.ReadPdfAndExtractText(filePath);
                    string text = string.Join(Environment.NewLine, extractedTexts);
                    FillMissingValuesFromCV(text, formModel);
                    formModel.CvContent = text;
                }
                catch (System.Exception ex)
                {
                    throw new PdfExtractionException("Error extracting content from the resume.", ex);
                }

                formModel.Resume = filePath;
                return true;
            }
            return false;
        }

        private void FillMissingValuesFromCV(string text, FormModelBase formModel)
        {
            const string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            const string phonePattern = @"0(5[0123456789])[^\D]{7}";

            if (string.IsNullOrEmpty(formModel.Email) && Regex.IsMatch(text, emailPattern))
            {
                formModel.Email = Regex.Match(text, emailPattern).Value;
            }

            if (string.IsNullOrEmpty(formModel.Phone) && Regex.IsMatch(text, phonePattern))
            {
                formModel.Phone = Regex.Match(text, phonePattern).Value;
            }
        }
    }
}
