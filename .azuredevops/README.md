# Azure DevOps Pipelines - SharePoint Sync

## Pipelines

- `pipeline-ci-cd.yml`: build and push Docker image for `sync-dotnet/Dockerfile`
- `pipeline-infra-shared.yml`: create the shared platform resources once per environment
- `pipeline-onboarding-companion.yml`: deploy one companion independently in the shared environment resource group
- `pipeline-decommission-companion.yml`: remove one companion and its companion-specific resources from the shared environment resource group

## Optional network hardening stage (infra pipeline)

`pipeline-infra-shared.yml` now supports an optional final stage for progressive network hardening.

- `allowSecurity` (`true|false`): enables/disables the final security stage
- hardening is always enforced when enabled
- shared VNet/subnet are deployed by the infra pipeline before hardening (`vnet-companion-shared-<env>` + private endpoint subnet)

Scope of this stage:

- hardens Storage, AI Search, and ACR (if Premium)
- keeps Container Apps unchanged (by design)
- keeps monitoring public (Log Analytics and App Insights)

## Resource group architecture

- all Azure resources are deployed in one shared RG per environment: `rg-companion-<env>`
- storage account, AI Search, Container Apps environment, Log Analytics, App Insights, ACR, managed identities, and Function Apps live in this shared RG
- the onboarding pipeline creates companion-specific resources inside the existing shared RG created by `pipeline-infra-shared.yml`

## Companion deployment target

The onboarding pipeline deploys:

- one Function App per companion (Azure Functions hosted on Azure Container Apps)
- one user-assigned managed identity per companion
- one blob container per companion in shared storage
- one AI Search index + datasource + indexer per companion in shared AI Search

## Companion decommission target

The decommission pipeline removes, in a safe order:

- companion Container App (Functions on Container Apps)
- companion AI Search indexer + datasource + index
- companion Blob container
- companion RBAC assignments (Storage, AI Search, ACR)
- companion user-assigned managed identity

The pipeline is idempotent: if a resource is already absent, the stage continues without error.

## Trigger behavior

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
