using System;

namespace emails_worker_service.Exceptions
{
    public class FormModelNullException : Exception
    {
        public FormModelNullException() : base("The form model cannot be null.") { }
    }

    public class MissingRequiredFieldsException : Exception
    {
        public MissingRequiredFieldsException() : base("One or more required fields are missing in the form model.") { }
    }

    public class ResumeFileNotFoundException : FileNotFoundException
    {
        public ResumeFileNotFoundException(string filePath) : base("The resume file is missing or does not exist.", filePath) { }
    }

    public class FormHtmlRetrievalException : HttpRequestException
    {
        public FormHtmlRetrievalException() : base("Failed to retrieve the form HTML.") { }
    }

    public class FormNotFoundException : InvalidOperationException
    {
        public FormNotFoundException() : base("Form not found in the HTML document.") { }
    }

    public class FormSubmissionException : Exception
    {
        public FormSubmissionException(string responseContent)
            : base("Failed to submit the form.")
        {
            ResponseContent = responseContent;
        }

        public string ResponseContent { get; }
    }

    namespace emails_worker_service.Exceptions
    {
        public class UnsupportedEmailSourceException : Exception
        {
            public UnsupportedEmailSourceException(string message) : base(message) { }
        }

        public class AttachmentProcessingException : Exception
        {
            public AttachmentProcessingException(string message) : base(message) { }
        }

        public class PdfExtractionException : Exception
        {
            public PdfExtractionException(string message, Exception innerException) : base(message, innerException) { }
        }

        public class EmailProcessingException : Exception
        {
            public EmailProcessingException(string message, Exception innerException) : base(message, innerException) { }
        }
    }
}

