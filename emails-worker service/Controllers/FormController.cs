using HtmlAgilityPack;
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using emails_worker_service.Models;
using emails_worker_service.Exceptions;

namespace emails_worker_service.Controllers
{
    public class FormController
    {
        private readonly IHttpClientFactory _httpClientFactory;

        public FormController(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
        }

        public async Task<FormSubmissionResponse> SubmitForm(FormModelBase model)
        {
            if (model == null)
            {
                throw new FormModelNullException();
            }

            if (string.IsNullOrWhiteSpace(model.TenderType) ||
                string.IsNullOrWhiteSpace(model.JobNumber) ||
                string.IsNullOrWhiteSpace(model.SubmissionDate) ||
                string.IsNullOrWhiteSpace(model.SubmissionTime) ||
                string.IsNullOrWhiteSpace(model.FirstName) ||
                string.IsNullOrWhiteSpace(model.LastName) ||
                string.IsNullOrWhiteSpace(model.Phone) ||
                string.IsNullOrWhiteSpace(model.Email) ||
                string.IsNullOrWhiteSpace(model.Exposure))
            {
                throw new MissingRequiredFieldsException();
            }

            if (string.IsNullOrEmpty(model.Resume) || !File.Exists(model.Resume))
            {
                throw new ResumeFileNotFoundException(model.Resume);
            }

            using (var client = _httpClientFactory.CreateClient())
            using (var formData = new MultipartFormDataContent())
            {
                formData.Add(new StringContent(model.TenderType), "tfa_7212");
                formData.Add(new StringContent(model.JobNumber), "tfa_7214");
                formData.Add(new StringContent(model.SubmissionDate), "tfa_7215");
                formData.Add(new StringContent(model.SubmissionTime), "tfa_7217");
                formData.Add(new StringContent(model.FirstName), "tfa_5");
                formData.Add(new StringContent(model.LastName), "tfa_1");
                formData.Add(new StringContent(model.Phone), "tfa_13");
                formData.Add(new StringContent(model.Email), "tfa_15");
                formData.Add(new StringContent(model.Exposure), "tfa_4903");

                using (var stream = File.OpenRead(model.Resume))
                {
                    formData.Add(new StreamContent(stream), "tfa_7218", Path.GetFileName(model.Resume));

                    var url = "https://www.tfaforms.com/5116778";
                    var response_get = await client.GetAsync(url);
                    if (!response_get.IsSuccessStatusCode)
                    {
                        throw new FormHtmlRetrievalException();
                    }

                    var pageContent = await response_get.Content.ReadAsStringAsync();

                    var doc = new HtmlDocument();
                    doc.LoadHtml(pageContent);

                    var form = doc.DocumentNode.SelectSingleNode("//form[@id='5116778']");
                    if (form == null)
                    {
                        throw new FormNotFoundException();
                    }

                    AddHiddenFields(form, formData);
                    var response = await client.PostAsync("https://www.tfaforms.com/api_v2/workflow/processor", formData);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new FormSubmissionException(responseContent);
                    }

                    return new FormSubmissionResponse
                    {
                        IsSuccess = true,
                        Message = "Form submitted successfully.",
                        ResponseContent = responseContent
                    };
                }
            }
        }

        private void AddHiddenFields(HtmlNode form, MultipartFormDataContent content)
        {
            var hiddenFields = form.SelectNodes("//input[@type='hidden']");
            if (hiddenFields != null)
            {
                foreach (var field in hiddenFields)
                {
                    var name = field.GetAttributeValue("name", "");
                    var value = field.GetAttributeValue("value", "");
                    content.Add(new StringContent(value), name);
                }
            }
        }
    }
}
