using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Models.FormModel
{
    public class FormModel
    {
        [Key]
        [Required(ErrorMessage = "כל אימייל חייב להיות משויך למזהה ייחודי")]
        [Display(Name = "מזהה אימייל")]
        public string MailId { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "סוג מכרז")]
        public string TenderType { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "מספר משרה/בקשת גיוס")]
        public string JobNumber { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "תאריך הגשה")]
        public string SubmissionDate { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "שעת הגשה")]
        public string SubmissionTime { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "שם פרטי")]
        public string FirstName { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "שם משפחה")]
        public string LastName { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "טלפון נייד")]
        [RegularExpression(@"^0(5[0123456789])[^\D]{7}$", ErrorMessage = "טלפון נייד לא חוקי")]
        public string Phone { get; set; }

        [Required(ErrorMessage = "שדה חובה")]
        [Display(Name = "דואר אלקטרוני")]
        [EmailAddress(ErrorMessage = "דואר אלקטרוני לא חוקי")]
        public string Email { get; set; }


        [Required(ErrorMessage = "חייב לבחור אחת מהאפשרויות")]
        [Display(Name = "כיצד נחשפתי למשרה")]
        public string Exposure { get; set; }

        public string CvContent { get; set; }
    }
}
