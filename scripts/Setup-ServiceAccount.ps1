<#
.SYNOPSIS
Provisions the svc-servicehealth service account in the dev tenant: creates the user,
assigns a license, and outputs the initial password.

.DESCRIPTION
Idempotent. Re-running:
- Reuses the existing user if present (matched by userPrincipalName).
- Does not reset the password if the user already exists.
- Adds the license only if not already assigned.

The signed-in user must be a User Administrator (or Global Admin) and a License
Administrator on the tenant.

.PARAMETER UserPrincipalName
The UPN to create. Default: svc-servicehealth@cloudopslabs.onmicrosoft.com

.PARAMETER DisplayName
Display name for the user. Default: Service Account - Service Health Monitor

.PARAMETER LicenseSkuPartNumber
SKU part number to assign. Default: SPB (Microsoft 365 Business Premium — current default for
M365 Developer Program tenants as of 2024+; older tenants may have DEVELOPERPACK_E5).

.PARAMETER UsageLocation
ISO 3166-1 alpha-2 country code. Required for license assignment. Default: US.

.EXAMPLE
.\Setup-ServiceAccount.ps1
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName = 'svc-servicehealth@cloudopslabs.onmicrosoft.com',
    [string]$DisplayName       = 'Service Account - Service Health Monitor',
    [string]$LicenseSkuPartNumber = 'SPB',
    [string]$UsageLocation     = 'US'
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable Az.Accounts)) {
    throw "Az.Accounts is not installed. Run: Install-Module Az.Accounts -Scope CurrentUser"
}
Import-Module Az.Accounts
if (-not (Get-AzContext)) { Connect-AzAccount | Out-Null }

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

function New-StrongPassword {
    $upper  = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $lower  = 'abcdefghijkmnopqrstuvwxyz'
    $digit  = '23456789'
    $symbol = '!@#$%^&*-_=+'
    $all = $upper + $lower + $digit + $symbol

    $chars = @(
        $upper[(Get-Random -Maximum $upper.Length)]
        $lower[(Get-Random -Maximum $lower.Length)]
        $digit[(Get-Random -Maximum $digit.Length)]
        $symbol[(Get-Random -Maximum $symbol.Length)]
    )
    for ($i = 0; $i -lt 16; $i++) {
        $chars += $all[(Get-Random -Maximum $all.Length)]
    }
    -join ($chars | Sort-Object { Get-Random })
}

# --- 1. Find or create user ---------------------------------------------

Write-Host "Looking up user '$UserPrincipalName'..." -ForegroundColor Cyan
$user = $null
try {
    $user = Invoke-Graph "users/$UserPrincipalName"
    Write-Host "  Found existing user. id=$($user.id)" -ForegroundColor Yellow
    $newlyCreatedPassword = $null
}
catch {
    if ($_.Exception.Response.StatusCode.value__ -ne 404) { throw }
    Write-Host "  Not found. Creating..." -ForegroundColor Green
    $newlyCreatedPassword = New-StrongPassword
    $mailNickname = ($UserPrincipalName -split '@')[0]
    $createBody = @{
        accountEnabled    = $true
        displayName       = $DisplayName
        mailNickname      = $mailNickname
        userPrincipalName = $UserPrincipalName
        usageLocation     = $UsageLocation
        passwordProfile   = @{
            forceChangePasswordNextSignIn = $false
            password                      = $newlyCreatedPassword
        }
    }
    $user = Invoke-Graph 'users' -Method POST -Body $createBody
    Write-Host "  Created. id=$($user.id)" -ForegroundColor Green

    # Print the password IMMEDIATELY so it's captured even if later steps fail.
    # If you miss it, reset via: PATCH /users/{id} with a new passwordProfile.
    Write-Host ""
    Write-Host "==== INITIAL PASSWORD (record now, will not be shown again) ====" -ForegroundColor Yellow
    Write-Host $newlyCreatedPassword -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host ""
}

# Ensure usageLocation is set even on an existing user (license assignment requires it).
if (-not $user.usageLocation) {
    Write-Host "  Setting usageLocation = $UsageLocation..." -ForegroundColor Cyan
    Invoke-Graph "users/$($user.id)" -Method PATCH -Body @{ usageLocation = $UsageLocation } | Out-Null
}

# --- 2. Resolve license SKU --------------------------------------------

Write-Host ""
Write-Host "Resolving SKU '$LicenseSkuPartNumber'..." -ForegroundColor Cyan
$skus = (Invoke-Graph 'subscribedSkus').value
$sku = $skus | Where-Object { $_.skuPartNumber -eq $LicenseSkuPartNumber }
if (-not $sku) {
    Write-Host "  Available SKUs in this tenant:" -ForegroundColor Yellow
    $skus | ForEach-Object { Write-Host "    - $($_.skuPartNumber) ($($_.skuId))" -ForegroundColor Yellow }
    throw "SKU '$LicenseSkuPartNumber' not found."
}
Write-Host "  SKU id : $($sku.skuId)" -ForegroundColor Gray
Write-Host "  Avail. : $($sku.prepaidUnits.enabled - $sku.consumedUnits) of $($sku.prepaidUnits.enabled)" -ForegroundColor Gray

# --- 3. Assign license if not already assigned -------------------------

Write-Host ""
$alreadyAssigned = $user.assignedLicenses | Where-Object { $_.skuId -eq $sku.skuId }
if ($alreadyAssigned) {
    Write-Host "License already assigned to user." -ForegroundColor DarkGray
} else {
    Write-Host "Assigning license..." -ForegroundColor Cyan
    $assignBody = @{
        addLicenses    = @(@{ skuId = $sku.skuId; disabledPlans = @() })
        removeLicenses = @()
    }
    Invoke-Graph "users/$($user.id)/assignLicense" -Method POST -Body $assignBody | Out-Null
    Write-Host "  Assigned." -ForegroundColor Green
}

# --- 4. Summary --------------------------------------------------------

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "UPN          : $UserPrincipalName"
Write-Host "DisplayName  : $DisplayName"
Write-Host "User id      : $($user.id)"
Write-Host "Tenant id    : $($tokenResp.TenantId)"
Write-Host "License      : $LicenseSkuPartNumber ($($sku.skuId))"
Write-Host "UsageLocation: $UsageLocation"
Write-Host ""
if ($newlyCreatedPassword) {
    Write-Host "Next: sign in once at https://portal.office.com to complete MFA setup" -ForegroundColor Cyan
    Write-Host "      (security defaults enforce MFA on every account in M365 dev tenants)."
    Write-Host "      Password was printed earlier in this run."
} else {
    Write-Host "(User already existed; password not reset by this run.)" -ForegroundColor DarkGray
}
