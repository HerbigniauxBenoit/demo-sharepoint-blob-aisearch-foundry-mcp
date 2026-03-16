using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharePointSync.Functions.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        services.AddHttpClient();
        services.AddSingleton<SharePointGraphClient>();
        services.AddSingleton<BlobStorageSyncClient>();
        services.AddSingleton<SharePointSyncOrchestrator>();
    })
    .Build();

host.Run();
