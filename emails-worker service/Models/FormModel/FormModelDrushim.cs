using DocumentFormat.OpenXml.Bibliography;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace emails_worker_service.Models.FormModel
{
    public class FormModelDrushim : FormModelBase
    {
        private const string drushim_exposure = "tfa_7211";
        /*        public override void FillForm(MailItem mailItem)
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
                    *//*
                        City = values[9];
                        experienceCurrent = values[11];
                        experienceTotal = values[13];
                    *//*

                }

        */
        public override void FillForm(MailItem mailItem)
        {
            // Extract job number from subject, e.g., "171-2024"
            JobNumber = Regex.Match(mailItem.Subject, @"\d+-20\d{2}").Value;

            // Extract tender type from subject, removing unwanted prefixes

            const string TenderPattern = @"\s\d*-20\d\d\s|Fwd:|FW:|קו""ח";
            const string email_pattern = @"^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?$";
            const string phone_pattern = @"^0(5[0123456789])[^\D]{7}$";
            const RegexOptions email_options = RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture;
            TenderType = Regex.Replace(mailItem.Subject.Split("|")[0], TenderPattern, "");
            TenderType = TenderType.Trim("-: ".ToCharArray());

            // Set submission date and time
            SubmissionDate = mailItem.CreationTime.ToShortDateString();
            SubmissionTime = mailItem.CreationTime.ToShortTimeString();
            Exposure = drushim_exposure;
            MailId = mailItem.EntryID;

            // Initialize default values
            FirstName = "missing";
            LastName = "missing";
            Phone = "missing";
            Email = "missing";
            //  City = "missing";

            // Load the HTML content from the email body
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(mailItem.HTMLBody);

            // Extract candidate details from the table rows
            var tableRows = doc.DocumentNode.SelectNodes("//table//tr");
            if (tableRows == null) return;

            foreach (var row in tableRows)
            {
                var cells = row.SelectNodes("td");
                if (cells == null || cells.Count < 2) continue;

                var key = CleanUpString(cells[0].InnerText);
                var value = CleanUpString(cells[1].InnerText);

                // Switch case to map the extracted values to the form fields
                switch (key)
                {
                    case "שם המועמד:":
                        var names = value.Split(" ");
                        if (names.Length > 0)
                        {
                            FirstName = names[0]; // The first name is always the first word
                            LastName = string.Join(" ", names.Skip(1)); // The rest of the words are considered the last name
                        }
                        break;
                    case "כתובת אימייל:":
                        if (Regex.IsMatch(Regex.Replace(value, @"\s+", ""), email_pattern, email_options))
                            Email = value;
                        break;
                    case "טלפון:":
                        if (Regex.IsMatch(Regex.Replace(value, @"\s+", ""), phone_pattern))
                            Phone = value;
                        break;
                        /* case "איזור מגורים:":
                             City = value;
                             break;*/
                        /* case "ישוב:":
                             Address = value;
                             break;*/
                        /*case "תפקיד אחרון:":
                            LastPosition = value;
                            break;*/
                        /*case "נסיון רלוונטי:":
                            RelevantExperience = value;
                            break;*/
                        // Add more cases for other fields as needed
                }
            }


        }

        public string CleanUpString(string input)
        {
            // Use regex to replace multiple whitespace characters (including newlines and tabs) with a single space
            string cleanedString = Regex.Replace(input, @"\s+", " ");

            // Trim the leading and trailing spaces
            cleanedString = cleanedString.Trim();

            // Ensure that the colon remains next to the last word by removing the space before it
            cleanedString = cleanedString.Replace(" :", ":");

            return cleanedString;
        }
    }
}

