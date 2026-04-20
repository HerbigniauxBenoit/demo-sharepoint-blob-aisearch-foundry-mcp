# SharePoint to Azure Blob Storage Sync with AI Search

This repository syncs documents from SharePoint Online to Azure Blob Storage, then indexes them with Azure AI Search.

## Target architecture

- Hosting: Azure Functions on Azure Container Apps
- Runtime: .NET 8 (isolated worker, Functions v4)
- Execution: TimerTrigger only
- Topology: one Function App (Container Apps hosted) per companion
- Identity: one user-assigned managed identity per companion
- Configuration: environment variables only

## Azure resource-group topology

- shared platform resources live in `companion-shared-<env>`
- all companion-specific resources live in `rg-companion`
- shared RG contains Storage, AI Search, Container Apps Environment, Log Analytics, and Application Insights
- companion RG contains the companion Function App and managed identity
- blob containers and AI Search objects are named by companion to avoid collisions in shared services

## Repository structure

- sync-dotnet/: .NET 8 Azure Functions timer worker that syncs SharePoint to Blob
- ai-search/: Azure AI Search datasource/index/indexer assets and scripts
- .azuredevops/: CI/CD and onboarding pipelines/templates

## Security prerequisites (manual)

The security team must configure Microsoft Graph access manually for each companion identity:

1. Grant Microsoft Graph application permission Sites.Selected to the managed identity service principal.
2. Grant admin consent.
3. Authorize the target SharePoint site for that identity.

These steps are intentionally not automated in pipelines.

## Deployment model

1. Build and push the sync image using .azuredevops/pipeline-ci-cd.yml.
2. Run shared infrastructure pipeline once per environment.
3. Run companion onboarding pipeline once per companion.
4. When needed, run companion decommission pipeline to remove one companion and its dedicated assets.

The onboarding pipeline deploys an Azure Function App on Container Apps hosting, configures timer settings and app settings, and binds the companion managed identity.
The decommission pipeline removes the companion Container App, AI Search assets, Blob container, RBAC assignments, and managed identity.