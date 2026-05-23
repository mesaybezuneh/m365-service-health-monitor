<#
.SYNOPSIS
Creates the IT Operations team and the M365 Service Alerts channel via Microsoft Graph,
and ensures the service account is a member.

.DESCRIPTION
Idempotent. Re-running:
- Reuses an existing team if one with the same displayName is found via Get-MgUserJoinedTeam.
- Adds the service account as a member only if not already a member.
- Reuses an existing channel if one with the same displayName is found.

Auth: uses the Microsoft.Graph PowerShell module (NOT Az.Accounts). Az PowerShell's first-party
app token does not carry TeamMember.* scopes, so member-management calls return 403. The
Microsoft.Graph PS module prompts for scope consent explicitly on first run, which works.

The signed-in user must have Teams Administrator or Global Administrator on the tenant and
will be set as the initial team owner. After this script, the user manually configures the
Workflows-based webhook in the Teams client (path B of the phase plan).

.PARAMETER TeamDisplayName
Default: IT Operations

.PARAMETER ChannelDisplayName
Default: M365 Service Alerts

.PARAMETER ServiceAccountUpn
UPN of the service account to add as a member. Default: svc-servicehealth@cloudopslabs.onmicrosoft.com

.EXAMPLE
.\Setup-TeamsChannel.ps1
#>
[CmdletBinding()]
param(
    [string]$TeamDisplayName     = 'IT Operations',
    [string]$ChannelDisplayName  = 'M365 Service Alerts',
    [string]$ServiceAccountUpn   = 'svc-servicehealth@cloudopslabs.onmicrosoft.com'
)

$ErrorActionPreference = 'Stop'

foreach ($m in 'Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Teams','Microsoft.Graph.Groups') {
    if (-not (Get-Module -ListAvailable $m)) {
        throw "$m is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Import-Module $m
}

# --- 1. Connect with the scopes we need --------------------------------

Connect-MgGraph -Scopes @(
    'User.Read.All',
    'Team.Create',
    'Team.ReadBasic.All',
    'TeamMember.ReadWrite.All',
    'Channel.Create',
    'Channel.ReadBasic.All',
    'Group.ReadWrite.All'
) -NoWelcome

# --- 2. Identify the signed-in human admin -----------------------------

Write-Host "Resolving signed-in user (will be team owner)..." -ForegroundColor Cyan
$me = Get-MgContext
$meUser = Get-MgUser -UserId $me.Account
Write-Host "  Owner : $($meUser.UserPrincipalName) (id $($meUser.Id))" -ForegroundColor Gray

Write-Host "Resolving service account..." -ForegroundColor Cyan
$svc = Get-MgUser -UserId $ServiceAccountUpn
Write-Host "  Member: $($svc.UserPrincipalName) (id $($svc.Id))" -ForegroundColor Gray

# --- 3. Find or create the team ----------------------------------------

Write-Host ""
Write-Host "Looking up team '$TeamDisplayName'..." -ForegroundColor Cyan

$existingTeams = Get-MgUserJoinedTeam -UserId $meUser.Id
$team = $existingTeams | Where-Object { $_.DisplayName -eq $TeamDisplayName } | Select-Object -First 1

if ($team) {
    Write-Host "  Found existing team. id=$($team.Id)" -ForegroundColor Yellow
} else {
    Write-Host "  Not found. Creating (async)..." -ForegroundColor Green
    $teamBody = @{
        'template@odata.bind' = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        displayName           = $TeamDisplayName
        description           = 'Operations team for M365 lifecycle and service-health automation. Owner: human admin. Channel posts come from automation flows.'
        members = @(
            @{
                '@odata.type'    = '#microsoft.graph.aadUserConversationMember'
                roles            = @('owner')
                'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$($meUser.Id)')"
            }
        )
    }

    # New-MgTeam returns 202 + Location; use raw request to capture the async-op URL.
    $resp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/teams' `
        -Body $teamBody -OutputType HttpResponseMessage
    $location = $resp.Headers.Location.OriginalString
    Write-Host "  202 Accepted. Polling teamsAsyncOperation..." -ForegroundColor Gray

    $opUri = "https://graph.microsoft.com$location"
    $deadline = (Get-Date).AddMinutes(5)
    do {
        Start-Sleep -Seconds 5
        $op = Invoke-MgGraphRequest -Method GET -Uri $opUri
        Write-Host "  Status: $($op.status)" -ForegroundColor DarkGray
    } until ($op.status -in 'succeeded','failed' -or (Get-Date) -gt $deadline)

    if ($op.status -ne 'succeeded') {
        throw "Team creation did not succeed within 5 min. Final status: $($op.status)."
    }

    $team = Get-MgTeam -TeamId $op.targetResourceId
    Write-Host "  Provisioned. teamId=$($team.Id)" -ForegroundColor Green
}

# --- 4. Ensure service account is a member -----------------------------

Write-Host ""
Write-Host "Checking team membership..." -ForegroundColor Cyan
$members = Get-MgTeamMember -TeamId $team.Id
$svcMember = $members | Where-Object { $_.AdditionalProperties.userId -eq $svc.Id }
if ($svcMember) {
    Write-Host "  Service account already a member." -ForegroundColor DarkGray
} else {
    Write-Host "  Adding service account as member..." -ForegroundColor Green
    $memberParams = @{
        '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
        Roles             = @()
        'User@odata.bind' = "https://graph.microsoft.com/v1.0/users('$($svc.Id)')"
    }
    New-MgTeamMember -TeamId $team.Id -BodyParameter $memberParams | Out-Null
    Write-Host "  Added." -ForegroundColor Green
}

# --- 5. Find or create the channel -------------------------------------

Write-Host ""
Write-Host "Looking up channel '$ChannelDisplayName'..." -ForegroundColor Cyan
$channels = Get-MgTeamChannel -TeamId $team.Id
$channel = $channels | Where-Object { $_.DisplayName -eq $ChannelDisplayName } | Select-Object -First 1
if ($channel) {
    Write-Host "  Found existing channel. id=$($channel.Id)" -ForegroundColor Yellow
} else {
    Write-Host "  Not found. Creating (standard)..." -ForegroundColor Green
    $channelParams = @{
        DisplayName    = $ChannelDisplayName
        Description    = 'Adaptive cards posted by the M365 Service Health poller flow.'
        MembershipType = 'standard'
    }
    $channel = New-MgTeamChannel -TeamId $team.Id -BodyParameter $channelParams
    Write-Host "  Created. id=$($channel.Id)" -ForegroundColor Green
}

# --- 6. Summary --------------------------------------------------------

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "Team        : $TeamDisplayName"
Write-Host "Team id     : $($team.Id)"
Write-Host "Channel     : $ChannelDisplayName"
Write-Host "Channel id  : $($channel.Id)"
Write-Host "Channel web : $($channel.WebUrl)"
Write-Host ""
Write-Host "Next (manual, in Teams desktop or web as YOU, not the service account):" -ForegroundColor Cyan
Write-Host "  1. Open the '$ChannelDisplayName' channel."
Write-Host "  2. Channel ... menu -> Workflows -> search 'Post to a channel when a webhook request is received'."
Write-Host "  3. Pick the team '$TeamDisplayName' and channel '$ChannelDisplayName'. Name the workflow 'ServiceHealth-PosterBot'."
Write-Host "  4. Save. Copy the generated HTTPS URL and paste it back into chat."
Write-Host "  5. In Power Automate (https://make.powerautomate.com), open 'ServiceHealth-PosterBot' -> Edit -> Owners -> Add owner -> svc-servicehealth -> Save."
Write-Host "  6. Optionally remove yourself as owner (only after the service account is confirmed as an owner)."
