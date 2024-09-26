using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Candidate_scoring
{
    public class TextPreprocessing
    {
        public static string CleanText(string text)
        {
            text = text.ToLower();
            // text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
            // text = System.Text.RegularExpressions.Regex.Replace(text, @"[^a-zA-Z0-9\s]", "");
            return text;
        }
    }
}
