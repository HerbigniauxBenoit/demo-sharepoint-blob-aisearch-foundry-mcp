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

            // Pour User Assigned Managed Identity, récupérer le client ID depuis la configuration
            var managedIdentityClientId = configuration["AZURE_CLIENT_ID"];

            if (!string.IsNullOrEmpty(managedIdentityClientId))
            {
                logger.LogInformation("Initializing User Assigned ManagedIdentityCredential with client ID: {ClientId}", managedIdentityClientId);
                return new ManagedIdentityCredential(managedIdentityClientId);
            }
            else
            {
                logger.LogInformation("Initializing System Assigned ManagedIdentityCredential - auto-detecting assigned identity");
                return new ManagedIdentityCredential();
            }
        });

        // Identity logging service
        services.AddSingleton<IdentityService>();

        // Application services
        services.AddScoped<SharePointGraphClient>();
        services.AddScoped<BlobStorageSyncClient>();
        services.AddScoped<SharePointSyncOrchestrator>();

        // Health checks
        services.AddHealthChecks()
            .AddCheck("self", () => Microsoft.Extensions.Diagnostics.HealthChecks.HealthCheckResult.Healthy());
    })
    .Build();

// Log identity information at startup
using (var scope = host.Services.CreateScope())
{
    var identityService = scope.ServiceProvider.GetRequiredService<IdentityService>();
    identityService.LogIdentityDetails();
}

host.Run();


