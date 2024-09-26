using emails_worker_service.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Data
{
    public interface IFormModelService
    {
        void SaveMailId(string mailId);
        void SaveBatchMailIds(Dictionary<string, object> mailIdsToAdd);
        bool ExistsMailId(string mailId);
        void RemoveMailId(string mailId);

    }
}
