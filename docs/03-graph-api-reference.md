# Deliverable 3 — Microsoft Graph API Reference

This document captures the exact endpoints, scopes, request shapes, and response payloads the flow needs. The Graph `serviceAnnouncement` resource is v1.0 (GA) and lives under the `/admin` namespace.

## App registration

- Display name: `M365-Service-Health-Monitor`
- Application (client) ID — fill in `config.local.json` after creation
- Authentication: client secret (single-tenant). Secret stored only in the Power Automate connection.

## Required Application permissions

| Permission | Why |
|---|---|
| `ServiceHealth.Read.All` | Read `/admin/serviceAnnouncement/issues` and `/healthOverviews`. |
| `ServiceMessage.Read.All` | Read `/admin/serviceAnnouncement/messages` (advisories, planned changes). Needed only if Flow A is extended to pull message center items. |

Both are **Application** permissions, not Delegated. Admin consent required. Grant via `Grant-AppGraphPermissions.ps1` (copy from the lifecycle automation repo) or via the Entra portal.

## Endpoints used

### 1. List active service issues

```
GET https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues
    ?$filter=isResolved eq false
    &$top=100
```

Response shape (truncated):

```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#admin/serviceAnnouncement/issues",
  "value": [
    {
      "id": "EX685896",
      "startDateTime": "2026-05-21T10:42:00Z",
      "endDateTime": null,
      "lastModifiedDateTime": "2026-05-21T11:15:00Z",
      "title": "Users may be unable to access their mailboxes in Exchange Online",
      "impactDescription": "Affected users are unable to access their mailboxes via Outlook on the web.",
      "classification": "incident",
      "origin": "microsoft",
      "status": "serviceDegradation",
      "service": "Exchange Online",
      "feature": "E-Mail timely delivery",
      "featureGroup": "E-Mail and calendar access",
      "isResolved": false,
      "highImpact": null,
      "details": [
        { "name": "TenantImpactCount", "value": "1000+" }
      ],
      "posts": [
        {
          "createdDateTime": "2026-05-21T10:45:00Z",
          "postType": "regular",
          "description": {
            "contentType": "html",
            "content": "<p>We're investigating an issue ...</p>"
          }
        }
      ]
    }
  ]
}
```

### 2. List recently resolved issues (for the resolution-sweep phase)

```
GET https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues
    ?$filter=isResolved eq true and lastModifiedDateTime ge <iso8601 minus 1 hour>
    &$top=100
```

Example filter value: `isResolved eq true and lastModifiedDateTime ge 2026-05-21T11:00:00Z`.

### 3. List message center messages (optional, for extended scope)

```
GET https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages
    ?$filter=category eq 'incident' or category eq 'planForChange'
    &$top=100
```

The `messages` resource has a different schema from `issues`. The most useful fields are `id`, `title`, `category`, `severity`, `startDateTime`, `endDateTime`, `lastModifiedDateTime`, `tags`, `services`, `body`.

### 4. Health overviews (optional, for a "current state" dashboard)

```
GET https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews
```

Returns one row per service with the current aggregate status. Useful if a dashboard view is added later. Not required for the v1 alerting flow.

## Authentication in Power Automate HTTP action

Use the built-in `Active Directory OAuth` authentication block:

| Field | Value |
|---|---|
| Authority | `https://login.microsoftonline.com` |
| Tenant | tenant ID (GUID) |
| Audience | `https://graph.microsoft.com` |
| Client ID | app registration client ID |
| Credential Type | Secret |
| Secret | the client secret (stored in the connection, never in the flow JSON in plaintext when exported via Solution; will appear in plaintext when exported via `Get-FlowDefinition.ps1`, hence the `flow-*.json` gitignore rule) |

## Throttling

Service health endpoints are subject to standard Graph throttling:

- Per-app: 10,000 requests per 10 minutes
- Per-tenant: depends on tenant scale; for a dev tenant this is effectively unbounded

Real volume from a 15-minute poller is on the order of 4 requests per hour. Throttling is not a real concern but the HTTP action should still have a retry policy on 429 with `Retry-After` honored. The Power Automate default retry policy already does this.

## Pagination

`@odata.nextLink` is returned when more than `$top` items are available. The `issues` endpoint with `isResolved eq false` will not realistically exceed 50 active items at once for a single tenant, so pagination is **not implemented** in v1. If the recovery sweep is widened or `messages` is added, follow the lifecycle automation's `Until + Compose_NextLink` pattern.

## How to test the endpoint outside Power Automate

Quickest path with PowerShell + an app-only token:

```powershell
$tenantId = "<tenantId>"
$clientId = "<appClientId>"
$clientSecret = "<secret>"

$body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
}
$tokenResp = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Body $body

$headers = @{ Authorization = "Bearer $($tokenResp.access_token)" }
Invoke-RestMethod -Method GET -Headers $headers `
    -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues?`$top=5" |
    ConvertTo-Json -Depth 6
```

Save the response under `scripts/sample-responses/` (gitignored if it contains tenant-identifying details) for use as the `Parse JSON` sample schema in Power Automate.

## Locked decisions

- **App registration:** new app `M365-Service-Health-Monitor`, separate from `M365-Lifecycle-Automation`. Least-privilege boundary. Only the two read-only health scopes. If this secret ever leaks, the blast radius is read-only access to service announcements, not user/group write.
- **Secret rotation:** manual for v1. Key Vault + managed identity is documented as a production-hardening item but not built into the dev-tenant proof-of-concept.
