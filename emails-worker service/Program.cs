using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using emails_worker_service.Controllers;
using emails_worker_service.Models;
using System.IO.Abstractions;
using emails_worker_service.Controllers;
using emails_worker_service.document;

namespace emails_worker_service
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    // Register IFileSystem
                    services.AddSingleton<IFileSystem, FileSystem>();
                    services.AddSingleton<IEmailService, OutlookService>(); // Email service that interacts with Outlook

                    // Register components and services
                    services.AddTransient<DocumentReaderComponent>(); // Pdf reader component for processing PDFs
                    services.AddTransient<FormCreatorOutlook>(); // Form creator for processing emails
                    services.AddTransient<FormModelServiceCsv>(provider =>
                        new FormModelServiceCsv("formModels.csv", provider.GetRequiredService<IFileSystem>())); // CSV service
                    services.AddTransient<FormSubmitSalesForce>(); // Form submitter service

                    // Register HttpClient factory
                    services.AddHttpClient(); // Register IHttpClientFactory for making HTTP requests

                    // Register the Worker as a hosted service
                    services.AddHostedService<Worker>();

                    // Configure logging
                    services.AddLogging(config =>
                    {
                        config.AddConsole(options =>
                        {
                            options.IncludeScopes = true;
                            options.TimestampFormat = "[HH:mm:ss] "; // Log timestamp format
                        });
                    });
                });
    }
}
