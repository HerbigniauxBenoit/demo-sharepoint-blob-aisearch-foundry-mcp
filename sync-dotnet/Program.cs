using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SharePointSync.Functions.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureAppConfiguration((context, config) =>
    {
        config.AddEnvironmentVariables();
        if (context.HostingEnvironment.IsDevelopment())
        {
            config.AddJsonFile("local.settings.json", optional: true, reloadOnChange: true);
        }
    })
    .ConfigureServices((context, services) =>
    {
        // Application Insights
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Configure logging
        services.AddLogging(logging =>
        {
            logging.AddConsole();
            logging.SetMinimumLevel(LogLevel.Information);
        });

        // Azure SDK clients
        services.AddHttpClient();

        // Configure credential for authentication
        services.AddSingleton<TokenCredential>(sp =>
        {
            var logger = sp.GetRequiredService<ILoggerFactory>().CreateLogger("TokenCredential");
            var configuration = sp.GetRequiredService<IConfiguration>();

            var managedIdentityClientId = configuration["AZURE_CLIENT_ID"];

            if (!string.IsNullOrWhiteSpace(managedIdentityClientId))
            {
                logger.LogInformation("Initializing ManagedIdentityCredential with Client ID: {ClientId}", managedIdentityClientId);
                return new ManagedIdentityCredential(managedIdentityClientId);
            }

            logger.LogInformation("Initializing DefaultAzureCredential for authentication");
            var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
            {
                ExcludeEnvironmentCredential = false,
                ExcludeWorkloadIdentityCredential = false,
                ExcludeManagedIdentityCredential = false,
                ExcludeVisualStudioCredential = false,
                ExcludeVisualStudioCodeCredential = false,
                ExcludeAzureCliCredential = false,
                ExcludeAzurePowerShellCredential = true,
                ExcludeInteractiveBrowserCredential = true
            });

            logger.LogInformation("DefaultAzureCredential initialized successfully");
            return credential;
        });

        // Application services
        services.AddScoped<SharePointGraphClient>();
        services.AddScoped<BlobStorageSyncClient>();
        services.AddScoped<SharePointSyncOrchestrator>();

        // Health checks
        services.AddHealthChecks()
            .AddCheck("self", () => Microsoft.Extensions.Diagnostics.HealthChecks.HealthCheckResult.Healthy());
    })
    .Build();

host.Run();
