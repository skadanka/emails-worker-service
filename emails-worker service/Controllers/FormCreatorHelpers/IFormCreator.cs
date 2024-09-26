using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;


namespace emails_worker_service.Controllers.FormCreator
{
    public interface IFormCreator
    {
        Dictionary<string, object> GetForms(Items inboxItems);
    }
}