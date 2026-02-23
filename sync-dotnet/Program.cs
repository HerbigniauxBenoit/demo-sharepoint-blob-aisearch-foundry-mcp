using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharePointSync.Functions.Services;

var builder = FunctionsApplication.CreateBuilder(args);

builder.Services.AddHttpClient();
builder.Services.AddSingleton<TokenCredential>(_ => new DefaultAzureCredential());
builder.Services.AddSingleton<SharePointGraphClient>();
builder.Services.AddSingleton<BlobStorageSyncClient>();
builder.Services.AddSingleton<SharePointSyncOrchestrator>();

builder.Build().Run();
