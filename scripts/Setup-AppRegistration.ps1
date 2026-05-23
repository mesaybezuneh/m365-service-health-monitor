<#
.SYNOPSIS
Creates (or updates) the M365-Service-Health-Monitor Entra app registration, adds the
service-health Graph Application permissions, grants admin consent, and optionally mints
a client secret.

.DESCRIPTION
Idempotent. Re-running:
- Reuses the existing app reg if a single match on display name is found.
- Adds only missing Graph permissions to requiredResourceAccess.
- Adds only missing appRoleAssignments (admin-consent grants).
- Does not create a new secret unless -CreateSecret is passed (so re-runs don't pile up secrets).

The signed-in user must be a Global Administrator or have equivalent rights to manage app
registrations and grant tenant-wide admin consent.

.PARAMETER DisplayName
The app registration's display name. Default: M365-Service-Health-Monitor

.PARAMETER Permissions
Graph Application permissions to grant. Defaults to the two service-health scopes.

.PARAMETER CreateSecret
Switch. When set, creates a new client secret valid for 6 months and writes it to stdout.
Record it in a password manager immediately; it is not retrievable later.

.PARAMETER SecretDescription
Friendly label that appears next to the secret in the Entra portal. Default: PowerAutomate-FlowConnection.

.EXAMPLE
.\Setup-AppRegistration.ps1

.EXAMPLE
.\Setup-AppRegistration.ps1 -CreateSecret
#>
[CmdletBinding()]
param(
    [string]$DisplayName = 'M365-Service-Health-Monitor',
    [string[]]$Permissions = @('ServiceHealth.Read.All', 'ServiceMessage.Read.All'),
    [switch]$CreateSecret,
    [string]$SecretDescription = 'PowerAutomate-FlowConnection'
)

$ErrorActionPreference = 'Stop'
$GraphAppId = '00000003-0000-0000-c000-000000000000'

# --- 0. Modules ----------------------------------------------------------

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    throw "Az.Accounts is not installed. Run: Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts

if (-not (Get-AzContext)) { Connect-AzAccount | Out-Null }

# --- 1. Get Graph token --------------------------------------------------

$tokenResp = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com/'
$token = if ($tokenResp.Token -is [System.Security.SecureString]) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResp.Token)
    try { [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
} else { $tokenResp.Token }

$headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }

function Invoke-Graph($Path, $Method = 'GET', $Body = $null) {
    $uri = "https://graph.microsoft.com/v1.0/$Path"
    if ($Body) {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method -Body ($Body | ConvertTo-Json -Depth 10)
    } else {
        Invoke-RestMethod -Uri $uri -Headers $headers -Method $Method
    }
}

# --- 2. Resolve Microsoft Graph SP + appRoles ----------------------------

Write-Host "Resolving Microsoft Graph service principal..." -ForegroundColor Cyan
$graphSp = (Invoke-Graph "servicePrincipals?`$filter=appId eq '$GraphAppId'&`$select=id,appRoles").value[0]
if (-not $graphSp) { throw "Microsoft Graph service principal not found in tenant" }

$desiredRoles = @()
foreach ($pn in $Permissions) {
    $role = $graphSp.appRoles | Where-Object { $_.value -eq $pn -and $_.allowedMemberTypes -contains 'Application' }
    if (-not $role) { throw "Application permission '$pn' not found on Microsoft Graph SP" }
    $desiredRoles += [PSCustomObject]@{ Name = $pn; Id = $role.id }
    Write-Host "  Permission '$pn' -> $($role.id)" -ForegroundColor Gray
}

# --- 3. Find or create the app registration -----------------------------

Write-Host ""
Write-Host "Looking up application '$DisplayName'..." -ForegroundColor Cyan
$apps = (Invoke-Graph "applications?`$filter=displayName eq '$DisplayName'&`$select=id,appId,displayName,requiredResourceAccess").value

if ($apps.Count -gt 1) {
    throw "Multiple app registrations found with display name '$DisplayName'. Resolve before re-running."
}

if ($apps.Count -eq 1) {
    $app = $apps[0]
    Write-Host "  Found existing app. objectId=$($app.id) appId=$($app.appId)" -ForegroundColor Yellow
} else {
    Write-Host "  Not found. Creating..." -ForegroundColor Green
    $createBody = @{
        displayName        = $DisplayName
        signInAudience     = 'AzureADMyOrg'
        description        = 'Reads Microsoft 365 service health via Microsoft Graph for the m365-service-health-monitor automation.'
    }
    $app = Invoke-Graph 'applications' -Method POST -Body $createBody
    Write-Host "  Created. objectId=$($app.id) appId=$($app.appId)" -ForegroundColor Green
}

# --- 4. Find or create the service principal ----------------------------

Write-Host ""
Write-Host "Looking up service principal for appId $($app.appId)..." -ForegroundColor Cyan
$targetSp = (Invoke-Graph "servicePrincipals?`$filter=appId eq '$($app.appId)'&`$select=id,displayName").value[0]
if ($targetSp) {
    Write-Host "  Found existing SP. objectId=$($targetSp.id)" -ForegroundColor Yellow
} else {
    Write-Host "  Not found. Creating..." -ForegroundColor Green
    $spBody = @{ appId = $app.appId }
    $targetSp = Invoke-Graph 'servicePrincipals' -Method POST -Body $spBody
    Write-Host "  Created. objectId=$($targetSp.id)" -ForegroundColor Green
}

# --- 5. Add missing perms to requiredResourceAccess ---------------------

Write-Host ""
Write-Host "Updating requiredResourceAccess..." -ForegroundColor Cyan
$rraNew = @()
$graphFound = $false
foreach ($entry in @($app.requiredResourceAccess)) {
    $accessList = @()
    foreach ($a in @($entry.resourceAccess)) {
        $accessList += @{ id = $a.id; type = $a.type }
    }
    if ($entry.resourceAppId -eq $GraphAppId) {
        $graphFound = $true
        $existingIds = @($accessList | ForEach-Object { $_.id })
        foreach ($r in $desiredRoles) {
            if ($existingIds -contains $r.Id) {
                Write-Host "  Already in RRA: $($r.Name)" -ForegroundColor DarkGray
            } else {
                $accessList += @{ id = $r.Id; type = 'Role' }
                Write-Host "  Adding to RRA: $($r.Name)" -ForegroundColor Green
            }
        }
    }
    $rraNew += @{ resourceAppId = $entry.resourceAppId; resourceAccess = $accessList }
}
if (-not $graphFound) {
    $accessList = @()
    foreach ($r in $desiredRoles) {
        $accessList += @{ id = $r.Id; type = 'Role' }
        Write-Host "  Adding to RRA (new Graph entry): $($r.Name)" -ForegroundColor Green
    }
    $rraNew += @{ resourceAppId = $GraphAppId; resourceAccess = $accessList }
}

Invoke-Graph "applications/$($app.id)" -Method PATCH -Body @{ requiredResourceAccess = $rraNew } | Out-Null
Write-Host "  PATCH applied." -ForegroundColor Green

# --- 6. Grant admin consent via appRoleAssignment -----------------------

Write-Host ""
Write-Host "Granting admin consent (appRoleAssignment) per permission..." -ForegroundColor Cyan
$existingGrants = (Invoke-Graph "servicePrincipals/$($targetSp.id)/appRoleAssignments").value
foreach ($r in $desiredRoles) {
    $already = $existingGrants | Where-Object { $_.appRoleId -eq $r.Id -and $_.resourceId -eq $graphSp.id }
    if ($already) {
        Write-Host "  Already granted: $($r.Name)" -ForegroundColor DarkGray
        continue
    }
    $body = @{
        principalId = $targetSp.id
        resourceId  = $graphSp.id
        appRoleId   = $r.Id
    }
    try {
        Invoke-Graph "servicePrincipals/$($targetSp.id)/appRoleAssignments" -Method POST -Body $body | Out-Null
        Write-Host "  Granted: $($r.Name)" -ForegroundColor Green
    } catch {
        $status = $_.Exception.Response.StatusCode.value__
        Write-Host "  FAILED to grant $($r.Name): HTTP $status" -ForegroundColor Red
        throw
    }
}

# --- 7. Optional: create a client secret --------------------------------

if ($CreateSecret) {
    Write-Host ""
    Write-Host "Creating client secret '$SecretDescription' (6-month validity)..." -ForegroundColor Cyan
    $endDate = (Get-Date).AddMonths(6).ToString('o')
    $secretBody = @{
        passwordCredential = @{
            displayName = $SecretDescription
            endDateTime = $endDate
        }
    }
    $secret = Invoke-Graph "applications/$($app.id)/addPassword" -Method POST -Body $secretBody
    Write-Host "  Secret created. keyId=$($secret.keyId) expires=$($secret.endDateTime)" -ForegroundColor Green
    Write-Host ""
    Write-Host "==== CLIENT SECRET (record now, not retrievable later) ====" -ForegroundColor Yellow
    Write-Host $secret.secretText -ForegroundColor Yellow
    Write-Host "===========================================================" -ForegroundColor Yellow
}

# --- 8. Summary ---------------------------------------------------------

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "TenantId   : $($tokenResp.TenantId)"
Write-Host "AppId      : $($app.appId)"
Write-Host "App ObjId  : $($app.id)"
Write-Host "SP  ObjId  : $($targetSp.id)"
Write-Host ""
Write-Host "Granted Application permissions:" -ForegroundColor Cyan
foreach ($r in $desiredRoles) { Write-Host "  - $($r.Name)" -ForegroundColor Green }
Write-Host ""
Write-Host "Next: paste appId into config.local.json (appRegistrations.serviceHealthMonitor.clientId)."
if (-not $CreateSecret) {
    Write-Host "      Re-run with -CreateSecret when you are ready to wire the Power Automate connection."
}
