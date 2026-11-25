# Agent estate audit: list agent identities and flag the ownerless ones
Connect-MgGraph -Scopes "AgentIdentity.Read.All" -NoWelcome

$base = "https://graph.microsoft.com/v1.0"

function Get-GraphPages {
    param([string]$Uri)
    $items = @()
    while ($Uri) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri
        if ($page.value) { $items += $page.value }
        $Uri = $page.'@odata.nextLink'
    }
    $items
}

$agents = Get-GraphPages "$base/servicePrincipals/microsoft.graph.agentIdentity"

function Get-Label($o) {
    if ($o.displayName)           { $o.displayName }
    elseif ($o.userPrincipalName) { $o.userPrincipalName }
    else { "$($o.id) [$(($o.'@odata.type') -replace '#microsoft.graph.','')]" }
}

$report = foreach ($agent in $agents) {
    $root     = "$base/servicePrincipals/$($agent.id)/microsoft.graph.agentIdentity"
    $owners   = @(Get-GraphPages "$root/owners")
    $sponsors = @(Get-GraphPages "$root/sponsors")

    [pscustomobject]@{
        DisplayName = $agent.displayName
        ObjectId    = $agent.id
        Blueprint   = $agent.agentIdentityBlueprintId
        Created     = $agent.createdDateTime
        Enabled     = $agent.accountEnabled
        Owners   = (($owners   | ForEach-Object { Get-Label $_ }) -join "; ")
        Sponsors = (($sponsors | ForEach-Object { Get-Label $_ }) -join "; ")
        Ownerless   = ($owners.Count -eq 0 -and $sponsors.Count -eq 0)
    }
}

$report | Sort-Object Ownerless -Descending |
    Format-Table DisplayName, Owners, Sponsors, Ownerless -AutoSize
$report | Export-Csv ".\agent-estate-audit.csv" -NoTypeInformation

$ownerless = @($report | Where-Object Ownerless).Count
Write-Host "$($report.Count) agent identities found, $ownerless ownerless."
