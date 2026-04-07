# Azure DevOps Pipelines - SharePoint Sync

## Pipelines

- `pipeline-ci-cd.yml`: build and push Docker image for `sync-dotnet/Dockerfile`
- `pipeline-infra-shared.yml`: create shared environment resources once per environment
- `pipeline-onboarding-companion.yml`: deploy one companion independently

## Companion deployment target

The onboarding pipeline deploys:

- one Function App per companion (Azure Functions hosted on Azure Container Apps)
- one user-assigned managed identity per companion
- one blob container per companion
- one AI Search index + datasource + indexer per companion

## Trigger model

The deployed function is timer-only. No HTTP endpoint is deployed or expected.

## Required manual security step

Graph and SharePoint authorization are manual prerequisites managed by the security team:

1. grant Graph `Sites.Selected` to the companion managed identity service principal,
2. grant admin consent,
3. grant site-level access on the target SharePoint site.

## Host storage and identity

The deploy template configures Functions host storage for managed identity with:

- `AzureWebJobsStorage__accountName`
- `AzureWebJobsStorage__credential=managedidentity`
- `AzureWebJobsStorage__clientId`

Storage RBAC template assigns host-compatible roles to each companion identity.
