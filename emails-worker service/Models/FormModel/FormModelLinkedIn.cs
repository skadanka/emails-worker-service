using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Outlook;

namespace emails_worker_service.Models.FormModel
{
    class FormModelLinkedIn : FormModelBase
    {
        private const string linkedIn_exposure = "tfa_7209";

        public override void FillForm(MailItem mailItem)
        {
            TenderType = ExtractBeforeFrom(Regex.Replace(mailItem.Subject, @"\s\d*-20\d\d\s|Fwd:|New application:", ""));
            // Missing job number in all LinkedIn templates, add in the Future form HR job number to job description.
            JobNumber = "0000-2024";
            SubmissionDate = mailItem.CreationTime.ToShortDateString();
            SubmissionTime = mailItem.CreationTime.ToShortTimeString();
            Exposure = linkedIn_exposure;
            MailId = mailItem.EntryID;

            FirstName = "missing";
            LastName = "missing";
            Phone = "missing";
            Email = "missing";

            TenderType = TenderType.Trim(" ".ToCharArray());
            HtmlDocument _doct = new HtmlDocument();
            _doct.LoadHtml(mailItem.HTMLBody);


            HtmlNodeCollection table_rows = _doct.DocumentNode.SelectNodes("//td//p[@class='MsoNormal']");
            if (table_rows == null) return;
            List<string> values = new List<string>();

            foreach (HtmlNode tr in table_rows)
            {
                string res = Regex.Replace(tr.InnerText, "\r|\n|\t|&nbsp;", string.Empty);
                if (res.Length > 0)
                {
                    values.Add(res.Trim(" ,".ToCharArray()));
                }
            }
            if (!string.IsNullOrEmpty(values[2]))
            {
                string[] FullName = Regex.Replace(values[2], @"\d+(?:st|nd|rd|th)\+?", string.Empty).Split(" ");
                FirstName = FullName[0];
                LastName = FullName[1];

            }

            Email = "missing";
            Phone = "missing";
            /*
                City = values[4];
                experienceCurrent = values[7];
                experienceTotal = values[8];
            */
        }

        public string ExtractBeforeFrom(string input)
        {
            // Find the index of the word "from"
            int fromIndex = input.IndexOf("from", StringComparison.OrdinalIgnoreCase);

            // If "from" is found, return the substring before it
            if (fromIndex >= 0)
            {
                return input.Substring(0, fromIndex).Trim();
            }

            // If "from" is not found, return the original string or handle it as needed
            return input;
        }

    }
}
