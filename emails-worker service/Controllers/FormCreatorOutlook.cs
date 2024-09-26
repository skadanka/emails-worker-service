
using emails_worker_service.Controllers.FormCreator;
using emails_worker_service.document;
using Microsoft.Office.Interop.Outlook;


namespace emails_worker_service.Controllers
{
    public class FormCreatorOutlook : IFormCreator
    {
        private readonly DocumentReaderComponent _pdfReader;
        private readonly FormProcessor _formProcessor;


        public FormCreatorOutlook(DocumentReaderComponent pdfReader)
        {
            _pdfReader = pdfReader;
            _formProcessor = new FormProcessor(_pdfReader);
        }

        /// <summary>
        /// Retrieves emails from the inbox and processes them into form models.
        /// </summary>
        /// <returns>A dictionary where the key is the MailId and the value is the FormModel or an error message.</returns>
        public Dictionary<string, object> GetForms(Items inboxItems)
        {
            const int Limit = 50;
        /*    Items inboxItems;
            try
            {
                var inbox = _outlookService.GetInbox();
                inboxItems = inbox.Items;
                inboxItems.Sort("CreationTime", true);

            }
            finally
            {
                _outlookService.Dispose();
            }*/
            Dictionary<string, object> formModelsDictionary = new Dictionary<string, object>();
            int countForms = 0;
            inboxItems.Sort("ReceivedTime");
            foreach (MailItem mailItem in inboxItems)
            {
                try
                {
                    // Filter and process the form
                    if (EmailFilter.IsSupportedEmail(mailItem.HTMLBody))
                    {
                        var result = _formProcessor.ProcessForm(mailItem);
                        formModelsDictionary[mailItem.EntryID] = result;
                        countForms++;
                    }

                    // if (countForms >= Limit) break;
                }
                catch (System.Exception ex)
                {
                    formModelsDictionary[mailItem.EntryID] = $"Error processing email: {ex.Message}";
                    Console.WriteLine($"Error processing email {mailItem.Subject}: {ex.Message}");
                }

            }


            return formModelsDictionary;
        }
    }
}
