using System.Text.RegularExpressions;

namespace emails_worker_service.Controllers
{
    public static class EmailFilter
    {
        /// <summary>
        /// Checks if the email body contains a supported email source.
        /// </summary>
        /// <param name="htmlBody">The HTML body of the email.</param>
        /// <returns>True if the email is from a supported source, otherwise false.</returns>
        public static bool IsSupportedEmail(string htmlBody)
        {
            return Regex.IsMatch(htmlBody, @"cv@drushim.co.il|jobs-listings@linkedin.com", RegexOptions.IgnoreCase);
        }
    }
}
