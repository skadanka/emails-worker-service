using emails_worker_service.Models.FormModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Controllers
{
    internal interface IFormSubmit
    {
        Task<FormSubmissionResponse> SubmitForm(FormModelBase model);
    }
}
