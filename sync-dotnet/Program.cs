using Azure.Core;
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

        services.AddSingleton<TokenCredentialFactory>();
        services.AddSingleton<TokenCredential>(sp =>
            sp.GetRequiredService<TokenCredentialFactory>().Create());

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
try
{
    using var scope = host.Services.CreateScope();
    var logger = scope.ServiceProvider.GetRequiredService<ILoggerFactory>().CreateLogger("Startup");
    logger.LogInformation("========== APPLICATION STARTING ==========");

    var identityService = scope.ServiceProvider.GetRequiredService<IdentityService>();
    await identityService.LogIdentityDetailsAsync();

    // Validate SharePoint access if configured
    var configuration = scope.ServiceProvider.GetRequiredService<IConfiguration>();
    var sharePointSiteUrl = configuration["SHAREPOINT_SITE_URL"];

    if (!string.IsNullOrEmpty(sharePointSiteUrl))
    {
        var (success, message) = await identityService.ValidateSharePointAccessAsync(sharePointSiteUrl);
        if (!success)
        {
            logger.LogWarning("SharePoint access validation failed: {Message}", message);
            logger.LogWarning("The application will start but may fail when trying to sync.");
        }
    }
    else
    {
        logger.LogWarning("SHAREPOINT_SITE_URL not configured - skipping access validation");
    }

    logger.LogInformation("========== IDENTITY CHECKS COMPLETED ==========");
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR during identity initialization: {ex.Message}");
    Console.WriteLine($"Stack trace: {ex.StackTrace}");
}

host.Run();


