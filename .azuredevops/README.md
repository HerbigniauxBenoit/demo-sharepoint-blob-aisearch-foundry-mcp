# Azure DevOps Pipelines - SharePoint Sync

## Pipelines

- `pipeline-ci-cd.yml`: build and push Docker image for `sync-dotnet/Dockerfile`
- `pipeline-infra-shared.yml`: create the shared platform resources once per environment
- `pipeline-onboarding-companion.yml`: deploy one companion independently in the companion resource group

## Resource group model

- shared resources are deployed in one RG per environment: `companion-shared-<env>`
- companion-dedicated Azure resources are deployed in one RG: `rg-companion`
- the onboarding pipeline creates `rg-companion` automatically if it does not exist
- storage account, AI Search, Container Apps environment, Log Analytics, and Application Insights stay in the shared RG
- managed identity and Function App are created in the companion RG

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
