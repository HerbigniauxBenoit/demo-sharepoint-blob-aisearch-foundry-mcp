# Azure DevOps Pipelines - SharePoint Sync

## Pipelines

- `pipeline-ci-cd.yml`: build and push Docker image for `sync-dotnet/Dockerfile`
- `pipeline-infra-shared.yml`: create the shared platform resources once per environment
- `pipeline-onboarding-companion.yml`: deploy one companion independently in the shared environment resource group

## Resource group model

- all Azure resources are deployed in one shared RG per environment: `rg-companion-<env>`
- storage account, AI Search, Container Apps environment, Log Analytics, App Insights, ACR, managed identities, and Function Apps live in this shared RG
- the onboarding pipeline creates companion-specific resources inside the existing shared RG created by `pipeline-infra-shared.yml`

## Companion deployment target

The onboarding pipeline deploys:

- one Function App per companion (Azure Functions hosted on Azure Container Apps)
- one user-assigned managed identity per companion
- one blob container per companion in shared storage
- one AI Search index + datasource + indexer per companion in shared AI Search

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
