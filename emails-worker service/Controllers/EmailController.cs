using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using emails_worker_service.Models;
using emails_worker_service.Pdf;
using emails_worker_service.Exceptions;
using emails_worker_service.Exceptions.emails_worker_service.Exceptions;

namespace emails_worker_service.Controllers
{
    public class EmailController
    {
        private readonly PdfReaderComponent _pdfReader;

        public EmailController(PdfReaderComponent pdfReader)
        {
            _pdfReader = pdfReader;
        }

        /// <summary>
        /// Retrieves emails from the inbox and processes them into form models.
        /// </summary>
        /// <returns>A dictionary where the key is the MailId and the value is the FormModel or an error message.</returns>
        public Dictionary<string, object> GetEmails()
        {
            const int Limit = 300;
            Application outlookApp = new Application();
            System.Console.OutputEncoding = new UTF8Encoding();

            // Get the Inbox folder
            MAPIFolder inbox = outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            string[] fileTypes = { ".doc", ".pdf", ".docx" };

            Dictionary<string, object> formModelsDictionary = new Dictionary<string, object>();
            int countForms = 0;

            foreach (MailItem mailItem in inbox.Items)
            {
                try
                {
                    // Filter emails based on specific sender addresses or other criteria
                    if (!Regex.IsMatch(mailItem.HTMLBody, @"cv@drushim.co.il|jobs-listings@linkedin.com"))
                    {
                        continue;
                    }

                    FormModelBase formModel = null;
                    Console.WriteLine("\nProcessing new form from email");

                    // Determine the source of the email and create the appropriate FormModel
                    if (Regex.IsMatch(mailItem.HTMLBody, @"Drushim|drushim", RegexOptions.IgnoreCase))
                    {
                        formModel = FormModelFactory.CreateFormModel("Drushim");
                    }
                    else if (Regex.IsMatch(mailItem.HTMLBody, @"LinkedIn|linkedin", RegexOptions.IgnoreCase))
                    {
                        formModel = FormModelFactory.CreateFormModel("LinkedIn");
                    }

                    // If the formModel is null, throw an exception for unsupported source
                    if (formModel == null)
                    {
                        throw new UnsupportedEmailSourceException("Unsupported email source or format.");
                    }

                    // Fill the form with data extracted from the email
                    formModel.FillForm(mailItem);

                    // Clean up the TenderType field
                    formModel.TenderType = formModel.TenderType.Replace("FW:", string.Empty)
                                                               .Replace("New application:", string.Empty);

                    // Process attachments to extract resume content
                    foreach (Attachment att in mailItem.Attachments)
                    {
                        string ext = Path.GetExtension(att.FileName).ToLower();
                        if (!Array.Exists(fileTypes, fileType => fileType.Equals(ext)))
                            continue;

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
                        break; // Process only the first valid attachment
                    }

                    // Add the processed form model to the dictionary
                    formModelsDictionary[mailItem.EntryID] = formModel;
                    countForms++;
                    Console.WriteLine($"{countForms}. Created Form Model for {formModel.FirstName} {formModel.LastName}");

                    if (countForms >= Limit) break;
                }
                catch (UnsupportedEmailSourceException ex)
                {
                    formModelsDictionary[mailItem.EntryID] = ex.Message;
                    Console.WriteLine($"Unsupported email source: {ex.Message}");
                }
                catch (PdfExtractionException ex)
                {
                    formModelsDictionary[mailItem.EntryID] = ex.Message;
                    Console.WriteLine($"PDF extraction error: {ex.Message}");
                }
                catch (System.Exception ex)
                {
                    formModelsDictionary[mailItem.EntryID] = $"Error processing email: {ex.Message}";
                    Console.WriteLine($"Error processing email {mailItem.Subject}: {ex.Message}");
                }
                finally
                {
                    // Release COM objects to prevent memory leaks
                    Marshal.ReleaseComObject(mailItem);
                }
            }

            // Release COM objects for Outlook application and inbox
            Marshal.ReleaseComObject(inbox);
            Marshal.ReleaseComObject(outlookApp);

            return formModelsDictionary;
        }

        /// <summary>
        /// Fills missing values in the form model from the extracted resume text.
        /// </summary>
        /// <param name="text">The extracted text from the resume.</param>
        /// <param name="formModel">The form model to be updated.</param>
        private void FillMissingValuesFromCV(string text, FormModelBase formModel)
        {
            // Define regex patterns for email and phone number extraction
            const string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            const string phonePattern = @"0(5[0123456789])[^\D]{7}";

            // Try to extract and assign email if missing
            if ((formModel.Email == "missing" || formModel.Email == null) && Regex.IsMatch(text, emailPattern, RegexOptions.IgnoreCase))
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
}
