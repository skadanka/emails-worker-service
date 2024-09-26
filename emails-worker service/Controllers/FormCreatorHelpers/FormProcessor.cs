using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Outlook;
using emails_worker_service.Models;
using emails_worker_service.document;
using emails_worker_service.Exceptions;
using System.Text.RegularExpressions;
using emails_worker_service.Models.FormModel;

namespace emails_worker_service.Controllers
{
    public class FormProcessor
    {
        private readonly AttachmentProcessor _attachmentProcessor;

        public FormProcessor(DocumentReaderComponent pdfReader)
        {
            _attachmentProcessor = new AttachmentProcessor(pdfReader);
        }

        /// <summary>
        /// Processes a form from an email.
        /// </summary>
        /// <param name="mailItem">The email containing the form data.</param>
        /// <returns>Either a FormModelBase instance or an error message.</returns>
        public object ProcessForm(MailItem mailItem)
        {
            FormModelBase formModel = CreateFormModel(mailItem);

            if (formModel == null)
            {
                throw new UnsupportedEmailSourceException("Unsupported email source or format.");
            }

            formModel.FillForm(mailItem);

            // Clean up TenderType
            formModel.TenderType = formModel.TenderType.Replace("FW:", string.Empty)
                                                       .Replace("New application:", string.Empty);

            // Process attachments
            foreach (Attachment att in mailItem.Attachments)
            {
                if (_attachmentProcessor.ProcessAttachment(att, formModel))
                {
                    break; // Process only the first valid attachment
                }
            }

            return formModel;
        }

        private FormModelBase CreateFormModel(MailItem mailItem)
        {
            if (Regex.IsMatch(mailItem.HTMLBody, @"Drushim|drushim", RegexOptions.IgnoreCase))
            {
                return FormModelFactory.CreateFormModel("Drushim");
            }
            else if (Regex.IsMatch(mailItem.HTMLBody, @"LinkedIn|linkedin", RegexOptions.IgnoreCase))
            {
                return FormModelFactory.CreateFormModel("LinkedIn");
            }

            return null;
        }
    }
}
