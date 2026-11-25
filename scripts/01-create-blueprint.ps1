#Requires -Version 7
<#
  Creates an agent identity blueprint, adds a local-dev client secret, exposes an
  access_agent scope, and creates the blueprint principal. Uses the Microsoft Graph
  calls documented at learn.microsoft.com/entra/agent-id/create-blueprint.

  Needs: PowerShell 7, the Microsoft.Graph.Authentication module, and the
  Agent ID Administrator role (required to add a secret or certificate).
  Writes TENANT_ID, BLUEPRINT_APP_ID and BLUEPRINT_SECRET to ..\.env  (git-ignored).
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$BlueprintName = "Portfolio Risk Monitor Blueprint",
    [int]$SecretDays = 30
)
$ErrorActionPreference = "Stop"
$graph = "https://graph.microsoft.com/v1.0"
$odata = @{ "OData-Version" = "4.0" }

Connect-MgGraph -TenantId $TenantId -NoWelcome -Scopes `
    "AgentIdentityBlueprint.Create", "AgentIdentityBlueprint.AddRemoveCreds.All", `
    "AgentIdentityBlueprint.UpdateAuthProperties.All", "AgentIdentityBlueprintPrincipal.Create", "User.Read"

$me = Invoke-MgGraphRequest -Method GET -Uri "$graph/me"

# 1. Blueprint (a sponsor is required, an owner is recommended)
$bp = Invoke-MgGraphRequest -Method POST -Uri "$graph/applications/" -Headers $odata -ContentType "application/json" -Body (@{
    "@odata.type"         = "Microsoft.Graph.AgentIdentityBlueprint"
    displayName           = $BlueprintName
    "sponsors@odata.bind" = @("$graph/users/$($me.id)")
    "owners@odata.bind"   = @("$graph/users/$($me.id)")
} | ConvertTo-Json -Depth 5)
$appId = $bp.appId
Write-Host "Blueprint created. appId = $appId"

# 2. Local-dev client secret (use a managed identity FIC or certificate in production)
$end = (Get-Date).ToUniversalTime().AddDays($SecretDays).ToString("yyyy-MM-ddTHH:mm:ssZ")
$secret = Invoke-MgGraphRequest -Method POST -Uri "$graph/applications(appId='$appId')/addPassword" -ContentType "application/json" -Body (@{
    passwordCredential = @{ displayName = "sidecar-local-dev"; endDateTime = $end }
} | ConvertTo-Json -Depth 5)

# 3. Identifier URI and access_agent scope (lets a front end pass a user token to the agent)
Invoke-MgGraphRequest -Method PATCH -Uri "$graph/applications(appId='$appId')" -Headers $odata -ContentType "application/json" -Body (@{
    identifierUris = @("api://$appId")
    api = @{ oauth2PermissionScopes = @(@{
        adminConsentDescription = "Allow the application to access the agent on behalf of the signed-in user."
        adminConsentDisplayName = "Access agent"
        id        = [guid]::NewGuid().ToString()
        isEnabled = $true
        type      = "User"
        value     = "access_agent"
    }) }
} | ConvertTo-Json -Depth 6) | Out-Null

# 4. Blueprint principal
Invoke-MgGraphRequest -Method POST -Uri "$graph/serviceprincipals/microsoft.graph.agentIdentityBlueprintPrincipal" -Headers $odata -ContentType "application/json" -Body (@{ appId = $appId } | ConvertTo-Json) | Out-Null

# 5. Save for docker compose (git-ignored)
$envFile = Join-Path $PSScriptRoot "..\.env"
@(
    "TENANT_ID=$TenantId"
    "BLUEPRINT_APP_ID=$appId"
    "BLUEPRINT_SECRET=$($secret.secretText)"
    "AGENT_IDENTITY_ID="
    "FRONTEND_CLIENT_ID="
) | Set-Content -Path $envFile -Encoding utf8
Write-Host "Saved $envFile. Next: create the agent identity in the Entra admin center (see README)."
