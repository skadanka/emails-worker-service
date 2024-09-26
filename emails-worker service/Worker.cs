using emails_worker_service.Controllers;
using emails_worker_service.Controllers.FormCreator;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using emails_worker_service.Models.FormModel;

namespace emails_worker_service
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger; // Logger to log information and errors
        private readonly IServiceProvider _serviceProvider; // Service provider to create scopes and resolve dependencies
        private readonly IEmailService _outlookService;

        // Constructor to initialize the worker with the logger and service provider
        public Worker(ILogger<Worker> logger, IServiceProvider serviceProvider, IEmailService outlookService)
        {
            _logger = logger;
            _serviceProvider = serviceProvider;
            _outlookService = outlookService;
        }

        // This method is the entry point for the worker and runs continuously until cancellation is requested
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

                using (var scope = _serviceProvider.CreateScope()) // Create a new scope for resolving scoped services
                {
                    var formCreatorOutlook = scope.ServiceProvider.GetRequiredService<FormCreatorOutlook>(); // Resolve the EmailController
                    var formModelService = scope.ServiceProvider.GetRequiredService<FormModelServiceCsv>(); // Resolve the FormModelServiceCsv
                    var formSubmitSalesForce = scope.ServiceProvider.GetRequiredService<FormSubmitSalesForce>(); // Resolve the FormController

                    MAPIFolder processedFolder = null;
                    MAPIFolder notCompletedFolder = null;
                    Items mailItems;

                    try
                    {
                        // Get the Inbox folder and email items
                        var inbox = _outlookService.GetInbox();
                        mailItems = inbox.Items;

                        // Get the target folders
                        processedFolder = _outlookService.GetSubfolder("Processed");
                        notCompletedFolder = _outlookService.GetSubfolder("Not Completed");

                        if (processedFolder == null || notCompletedFolder == null)
                        {
                            _logger.LogError("One or more target folders do not exist.");
                            break;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        _logger.LogError("Outlook inbox connection failed. " + ex.Message);
                        break;
                    }

                    try
                    {
                        // Step 1: Process emails and get form models
                        Dictionary<string, object> formModels = formCreatorOutlook.GetForms(mailItems);

                        // Step 2: Remove existing emails (already processed) from the form models
                      /*  Dictionary<string, string> records = formModelService.LoadRecords();
                        foreach (var key in records.Keys)
                        {
                            formModels.Remove(key);
                        }*/

                        // Step 3: Save the new form submissions to the CSV
                        try
                        {
                            if (formModels.Count > 0)
                            {
                        formModelService.SaveBatchMailIds(formModels);
                                _logger.LogInformation(formModels.Count + " New records added to the database.");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            _logger.LogError("Failed saving to CSV! " + ex.Message);
                        }

                        // Step 4: Submit each form to an external service
                        foreach (var kvp in formModels)
                        {
                            if (kvp.Value is FormModelBase formModel) // Check if the value is a FormModelBase instance
                            {
                                var result = await formSubmitSalesForce.SubmitForm(formModel); // Submit the form and get the result
                                _logger.LogInformation("Form submitted for {mailId}: {result}", kvp.Key, result.Message);

                                // Move successfully processed email to the "Processed" folder
                                var mailItem = _outlookService.GetMailItemById(kvp.Key);
                                if (mailItem != null)
                                {
                                    _outlookService.MoveItemToFolder(mailItem, processedFolder);
                                }
                            }
                            else if (kvp.Value is string errorMessage) // Check if the value is an error message
                            {
                                _logger.LogError("Error processing {mailId}: {errorMessage}", kvp.Key, errorMessage);

                                // Move failed email to the "Not Fully Completed" folder
                                var mailItem = _outlookService.GetMailItemById(kvp.Key);
                                if (mailItem != null)
                                {
                                    // string coloredErrorMessage = $"<p style='color:red;'>Error encountered during processing: {errorMessage}</p>";
                                    // mailItem.HTMLBody = coloredErrorMessage + mailItem.HTMLBody;
                                    _outlookService.MoveItemToFolder(mailItem, notCompletedFolder);
                                }
                            }
                        }

                        // Log that processing has completed successfully
                        _logger.LogInformation(new DateTime().ToString() + " Processing completed successfully.");
                    }
                    catch (System.Exception ex)
                    {
                        // Log any errors that occurred during processing
                        _logger.LogError(ex, "An error occurred during processing.");
                    }
                }

                // Wait for 24 hours before running the worker again
                await Task.Delay(TimeSpan.FromHours(12), stoppingToken);
            }
        }
    }
}
