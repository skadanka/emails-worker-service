using HtmlAgilityPack;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace emails_worker_service.Models
{
    public class FormModelDrushim : FormModelBase
    {
        private const string drushim_exposure = "tfa_7211";
        public override void FillForm(MailItem mailItem)
        {
            JobNumber = Regex.Match(mailItem.Subject, @"\s\d*-20\d\d\s").Value.Replace(" ", "");

            const string TenderPattern = @"\s\d*-20\d\d\s|Fwd:|FW:|קו""ח"  ;

            TenderType = Regex.Replace(mailItem.Subject.Split("|")[0], TenderPattern, "");
            TenderType = TenderType.Trim("-: ".ToCharArray());
            SubmissionDate = mailItem.CreationTime.ToShortDateString();
            SubmissionTime = mailItem.CreationTime.ToShortTimeString();
            Exposure = drushim_exposure;
            MailId = mailItem.EntryID;
            FirstName = "missing";
            LastName = "missing";
            Phone = "missing";
            Email = "missing";
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
                    values.Add(res.Trim(" :".ToCharArray()));
                }
            }
            
            if (!string.IsNullOrEmpty(values[1]))
            {
                string[] FullName = values[1].Split(" ");
                FirstName = FullName[0];
                LastName = FullName[1];

            }
            const string email_pattern = @"^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$";
            const RegexOptions email_options = RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture;
            const string phone_pattern = @"^0(5[0123456789])[^\D]{7}$";
            if (!string.IsNullOrWhiteSpace(values[5]) && Regex.IsMatch(Regex.Replace(values[3], @"\s+", ""), email_pattern, email_options))
                Email = values[3];
            if (!string.IsNullOrWhiteSpace(values[7]) && Regex.IsMatch(Regex.Replace(values[5], @"\s+", ""), phone_pattern))
                Phone = values[5];
            /*
                City = values[9];
                experienceCurrent = values[11];
                experienceTotal = values[13];
            */

        }
    }
}
