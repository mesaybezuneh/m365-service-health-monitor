<#
.SYNOPSIS
Creates the ServiceOps Communication Site and the ServiceHealthIncidents list described in
docs/01-sharepoint-schema.md.

.DESCRIPTION
Idempotent. Re-running this script:
- Skips site creation if the site already exists.
- Skips list creation if the list already exists.
- Adds missing columns. Existing columns are left untouched.
- Re-applies indexes and view definitions on every run.

Prerequisites:
- PowerShell 7 (recommended) or Windows PowerShell 5.1.
- PnP.PowerShell module installed:
    Install-Module PnP.PowerShell -Scope CurrentUser
- An Entra ID app registration for PnP PowerShell. The same app you used for the
  m365-lifecycle-automation setup is fine; nothing in this script needs a separate
  identity. If you do not have one yet:
    Register-PnPEntraIDApp -ApplicationName "PnP-Lifecycle-Setup" -Tenant <tenant>.onmicrosoft.com -Interactive
  Record the ClientId it returns and pass it to this script via -ClientId.
- Global Admin or SharePoint Admin on the tenant.

.PARAMETER TenantHostname
The tenant's SharePoint hostname, without protocol or path.
Example: contoso.sharepoint.com

.PARAMETER ClientId
GUID of the Entra ID app registration created for PnP PowerShell. Used by every
internal Connect-PnPOnline call.

.PARAMETER SiteAlias
The URL segment after /sites/ for the new site. Default: ServiceOps

.PARAMETER SiteTitle
The display name of the new site. Default: Service Operations

.PARAMETER ListName
The internal list name. Default: ServiceHealthIncidents

.EXAMPLE
.\Setup-IncidentList.ps1 -TenantHostname contoso.sharepoint.com -ClientId "9ce48a27-ef22-4210-b98c-7bedf274cf09"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantHostname,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$ClientId,

    [string]$SiteAlias = 'ServiceOps',

    [string]$SiteTitle = 'Service Operations',

    [string]$ListName = 'ServiceHealthIncidents'
)

$ErrorActionPreference = 'Stop'

# --- 0. Module check ------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}

Import-Module PnP.PowerShell -DisableNameChecking

# --- 1. Resolve URLs ------------------------------------------------------

$TenantHostname = $TenantHostname.TrimEnd('/').Replace('https://', '').Replace('http://', '')
$AdminUrl = "https://$($TenantHostname.Split('.')[0])-admin.sharepoint.com"
$TenantUrl = "https://$TenantHostname"
$SiteUrl = "$TenantUrl/sites/$SiteAlias"

Write-Host "Tenant URL  : $TenantUrl"
Write-Host "Admin URL   : $AdminUrl"
Write-Host "Site URL    : $SiteUrl"
Write-Host "List name   : $ListName"
Write-Host "Client ID   : $ClientId"
Write-Host ""

# --- 2. Create site if missing -------------------------------------------

Write-Host "Connecting to tenant admin..." -ForegroundColor Cyan
Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId

$existingSite = Get-PnPTenantSite -Identity $SiteUrl -ErrorAction SilentlyContinue

if ($existingSite) {
    Write-Host "Site already exists at $SiteUrl, skipping create." -ForegroundColor Yellow
}
else {
    Write-Host "Creating Communication Site $SiteUrl..." -ForegroundColor Cyan
    New-PnPSite `
        -Type CommunicationSite `
        -Title $SiteTitle `
        -Url $SiteUrl `
        -Description "Hosts service-operations lists. Initial list: ServiceHealthIncidents (Flow-ServiceHealth-Poller)." | Out-Null
    Start-Sleep -Seconds 10
}

# --- 3. Connect to the new site ------------------------------------------

Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# --- 4. Create the list ---------------------------------------------------

$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

if ($list) {
    Write-Host "List $ListName already exists, skipping create." -ForegroundColor Yellow
}
else {
    Write-Host "Creating list $ListName..." -ForegroundColor Cyan
    $list = New-PnPList -Title $ListName -Template GenericList -EnableVersioning -OnQuickLaunch
    Set-PnPList -Identity $ListName -MajorVersions 50
}

# The Title column holds the Graph issue title. Rename its display name for clarity;
# the internal name stays "Title" because SharePoint requires it.
Set-PnPField -List $ListName -Identity 'Title' -Values @{ Title = 'Incident Title' } | Out-Null

# --- 5. Column schema -----------------------------------------------------
# Defined as an ordered array so columns appear in the list in a predictable order.
# Status choices mirror the Microsoft Graph `status` enum exactly; do not localize.

$columns = @(
    @{ InternalName = 'IncidentID';              DisplayName = 'Incident ID';                 Type = 'Text';     Required = $true  }
    @{ InternalName = 'Service';                 DisplayName = 'Service';                     Type = 'Text';     Required = $true  }
    @{ InternalName = 'Feature';                 DisplayName = 'Feature';                     Type = 'Text';     Required = $false }
    @{ InternalName = 'Classification';          DisplayName = 'Classification';              Type = 'Choice';   Required = $true;  Choices = @('Incident', 'Advisory') }
    @{ InternalName = 'Status';                  DisplayName = 'Status';                      Type = 'Choice';   Required = $true;  Choices = @(
        'serviceDegradation',
        'serviceInterruption',
        'restoringService',
        'extendedRecovery',
        'serviceRestored',
        'postIncidentReviewPublished',
        'serviceOperational',
        'falsePositive',
        'investigationSuspended'
    ) }
    @{ InternalName = 'Severity';                DisplayName = 'Severity';                    Type = 'Choice';   Required = $false; Choices = @('High', 'Medium', 'Low', 'Resolved', 'Info') }
    @{ InternalName = 'StartDateTime';           DisplayName = 'Start (UTC)';                 Type = 'DateTime'; Required = $true;  DisplayFormat = 1 }
    @{ InternalName = 'EndDateTime';             DisplayName = 'End (UTC)';                   Type = 'DateTime'; Required = $false; DisplayFormat = 1 }
    @{ InternalName = 'LastModifiedDateTime';    DisplayName = 'Last Modified (UTC)';         Type = 'DateTime'; Required = $true;  DisplayFormat = 1 }
    @{ InternalName = 'ImpactDescription';       DisplayName = 'Impact Description';          Type = 'Note';     Required = $false; PlainText = $true }
    @{ InternalName = 'LatestPostText';          DisplayName = 'Latest Post';                 Type = 'Note';     Required = $false; PlainText = $false }
    @{ InternalName = 'LatestPostType';          DisplayName = 'Latest Post Type';            Type = 'Text';     Required = $false }
    @{ InternalName = 'TeamsPostUrl';            DisplayName = 'Teams Post URL';              Type = 'URL';      Required = $false }
    @{ InternalName = 'PostedToTeams';           DisplayName = 'Posted to Teams';             Type = 'Boolean';  Required = $true }
    @{ InternalName = 'DigestSent';              DisplayName = 'Digest Sent';                 Type = 'Boolean';  Required = $true }
    @{ InternalName = 'FirstSeenDateTime';       DisplayName = 'First Seen (UTC)';            Type = 'DateTime'; Required = $true;  DisplayFormat = 1 }
    @{ InternalName = 'LastSeenDateTime';        DisplayName = 'Last Seen (UTC)';             Type = 'DateTime'; Required = $true;  DisplayFormat = 1 }
)

# --- 6. Add columns -------------------------------------------------------

foreach ($col in $columns) {
    $existing = Get-PnPField -List $ListName -Identity $col.InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [skip] $($col.InternalName) already exists" -ForegroundColor DarkGray
        continue
    }

    Write-Host "  [add ] $($col.InternalName) ($($col.Type))" -ForegroundColor Green

    $params = @{
        List             = $ListName
        InternalName     = $col.InternalName
        DisplayName      = $col.DisplayName
        Type             = $col.Type
        AddToDefaultView = $false
    }
    if ($col.Choices) { $params.Choices = $col.Choices }

    Add-PnPField @params | Out-Null

    # Post-creation property tweaks
    $postValues = @{}
    if ($col.Required) { $postValues.Required = $true }
    if ($null -ne $col.DisplayFormat) { $postValues.DisplayFormat = $col.DisplayFormat }
    if ($col.Type -eq 'Note' -and $null -ne $col.PlainText) {
        $postValues.RichText = (-not $col.PlainText)
        $postValues.AppendOnly = $false
    }
    if ($col.Type -eq 'Boolean') {
        # Default = false. SharePoint stores Boolean defaults as the string '0' or '1'.
        $postValues.DefaultValue = '0'
    }

    if ($postValues.Count -gt 0) {
        Set-PnPField -List $ListName -Identity $col.InternalName -Values $postValues | Out-Null
    }
}

# --- 7. Indexes -----------------------------------------------------------
# Three indexes per docs/01: IncidentID (dedupe), LastModifiedDateTime (digest window),
# Service (digest grouping + ad hoc queries).

$indexedColumns = @('IncidentID', 'LastModifiedDateTime', 'Service')

Write-Host ""
Write-Host "Setting indexes..." -ForegroundColor Cyan
foreach ($name in $indexedColumns) {
    Set-PnPField -List $ListName -Identity $name -Values @{ Indexed = $true } | Out-Null
    Write-Host "  [idx ] $name" -ForegroundColor Green
}

# --- 8. Views -------------------------------------------------------------

Write-Host ""
Write-Host "Creating views..." -ForegroundColor Cyan

# Helper: create a view, or replace it if one with the same name already exists.
function Set-IncidentView {
    param(
        [string]$Title,
        [string[]]$Fields,
        [string]$Query,
        [switch]$SetAsDefault
    )
    $existing = Get-PnPView -List $ListName -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-PnPView -List $ListName -Identity $Title -Force | Out-Null
    }
    Add-PnPView -List $ListName -Title $Title -Fields $Fields -Query $Query -SetAsDefault:$SetAsDefault | Out-Null
    Write-Host "  [view] $Title$(if ($SetAsDefault) { ' (default)' })" -ForegroundColor Green
}

# Sort: most recently modified first
$sortByLastModifiedDesc = '<OrderBy><FieldRef Name="LastModifiedDateTime" Ascending="FALSE" /></OrderBy>'

# Active incidents — anything NOT in a terminal state
$activeFilter = @'
<Where>
  <And>
    <And>
      <Neq><FieldRef Name='Status' /><Value Type='Text'>serviceRestored</Value></Neq>
      <Neq><FieldRef Name='Status' /><Value Type='Text'>postIncidentReviewPublished</Value></Neq>
    </And>
    <And>
      <Neq><FieldRef Name='Status' /><Value Type='Text'>serviceOperational</Value></Neq>
      <Neq><FieldRef Name='Status' /><Value Type='Text'>falsePositive</Value></Neq>
    </And>
  </And>
</Where>
'@

Set-IncidentView -Title 'Active incidents' -SetAsDefault -Fields @(
    'Title','IncidentID','Service','Severity','Status','LastModifiedDateTime','TeamsPostUrl'
) -Query ($activeFilter + $sortByLastModifiedDesc)

# Last 24 hours — SharePoint CAML cannot do hour math, so this is "since yesterday".
# The Flow-B digest uses an ISO datetime filter for true 24-hour precision; this view
# is for the SharePoint UI only.
$last24Filter = @'
<Where>
  <Geq>
    <FieldRef Name='LastModifiedDateTime' />
    <Value Type='DateTime'><Today OffsetDays='-1' /></Value>
  </Geq>
</Where>
'@

Set-IncidentView -Title 'Last 24 hours' -Fields @(
    'Title','IncidentID','Service','Severity','Status','LastModifiedDateTime'
) -Query ($last24Filter + $sortByLastModifiedDesc)

# Pending Teams post — rows where the Teams post failed and needs retry.
$pendingPostFilter = @'
<Where>
  <Eq>
    <FieldRef Name='PostedToTeams' />
    <Value Type='Boolean'>0</Value>
  </Eq>
</Where>
'@
$pendingPostSort = '<OrderBy><FieldRef Name="FirstSeenDateTime" Ascending="TRUE" /></OrderBy>'

Set-IncidentView -Title 'Pending Teams post' -Fields @(
    'Title','IncidentID','Service','Severity','FirstSeenDateTime','PostedToTeams'
) -Query ($pendingPostFilter + $pendingPostSort)

# All by service — grouped, no filter
$groupByService = @"
<GroupBy Collapse='TRUE' GroupLimit='30'>
  <FieldRef Name='Service' />
</GroupBy>
$sortByLastModifiedDesc
"@

Set-IncidentView -Title 'All by service' -Fields @(
    'Title','IncidentID','Service','Severity','Status','LastModifiedDateTime'
) -Query $groupByService

# Detail view: every column. Useful for debugging the poller against a single row.
$allFieldNames = @('Title') + ($columns | ForEach-Object { $_.InternalName })
Set-IncidentView -Title 'Incident Detail' -Fields $allFieldNames -Query $sortByLastModifiedDesc

# --- 9. Summary -----------------------------------------------------------

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green
Write-Host ""
Write-Host "Site URL : $SiteUrl"
Write-Host "List URL : $SiteUrl/Lists/$ListName"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Open the list, confirm 'Active incidents' is the default view and renders empty."
Write-Host "  2. Add a manual test row (e.g., IncidentID = TEST001, Service = Test, Status = serviceDegradation) to verify required fields enforce. Delete after."
Write-Host "  3. Record the site URL and list URL in config.local.json (sharepoint.siteUrl, sharepoint.listName)."
Write-Host "  4. Capture the list GUID for config.local.json:"
Write-Host "       (Get-PnPList -Identity '$ListName').Id"
Write-Host "  5. Move on to: register the M365-Service-Health-Monitor Entra app per docs/03."
