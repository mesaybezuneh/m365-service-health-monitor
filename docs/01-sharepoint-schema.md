# Deliverable 1 — SharePoint Schema

The poller flow writes one row per service-health incident or advisory and updates that row when Microsoft updates the incident. The schema below supports both the real-time post path and the daily digest query path.

## List

- Site: `https://<yourtenant>.sharepoint.com/sites/ServiceOps` (new site, dedicated to service-operations data — separate from `ITAutomation` which hosts `LifecycleAuditLog`)
- List internal name: `ServiceHealthIncidents`
- List display name: `Service Health Incidents`

### Why a separate site

The lifecycle audit log holds identity events (user creates, terminations, license changes). The service health incident log holds operational telemetry. Different audiences, different retention concerns, different permission boundaries. A dedicated `ServiceOps` site also leaves room to add adjacent ops lists later (capacity planning, license utilization, message-center backlog) without crowding the identity-automation site.

## Columns

| Display name | Internal name | Type | Required | Notes |
|---|---|---|---|---|
| Title | Title | Single line | Yes | Use the incident `title` from Graph. SharePoint requires Title; do not rename. |
| IncidentID | IncidentID | Single line (50) | Yes | Graph `id`. Two-letter service prefix + 6-7 digit serial. Examples seen in the wild: `EX685896` (Exchange), `MO123456` (Microsoft 365 platform), `CW1218323` (Copilot/Copilot Chat), `TM12345` (Teams), `SP67890` (SharePoint), `OD11111` (OneDrive), `MC123456` (Message Center). Microsoft does not document the full prefix set and adds new ones when services launch. Treat as opaque text; the column width of 50 is generous to absorb future format changes. **Dedupe key.** Indexed. |
| Service | Service | Single line (100) | Yes | Graph `service`, e.g., `Exchange Online`. Indexed. |
| Feature | Feature | Single line (200) | No | Graph `feature` (sub-service). |
| Classification | Classification | Choice | Yes | `Incident`, `Advisory`. From Graph `classification`. |
| Status | Status | Choice | Yes | `serviceDegradation`, `serviceInterruption`, `restoringService`, `extendedRecovery`, `serviceRestored`, `postIncidentReviewPublished`, `serviceOperational`, `falsePositive`, `investigationSuspended`. Mirror Graph `status` exactly. |
| Severity | Severity | Choice | No | Derived: map `serviceInterruption` → High, `serviceDegradation` → Medium, advisories → Low. See `docs/02` for the mapping. |
| StartDateTime | StartDateTime | Date and time | Yes | Graph `startDateTime`. UTC. |
| EndDateTime | EndDateTime | Date and time | No | Graph `endDateTime`. UTC. Empty while active. |
| LastModifiedDateTime | LastModifiedDateTime | Date and time | Yes | Graph `lastModifiedDateTime`. **Used to detect updates.** Indexed. |
| ImpactDescription | ImpactDescription | Multiple lines (plain) | No | Graph `impactDescription`. |
| LatestPostText | LatestPostText | Multiple lines (rich) | No | Most recent entry from Graph `posts[]` array. Plain text or basic HTML. |
| LatestPostType | LatestPostType | Single line | No | The `postType` of the latest post (e.g., `regular`, `quickFix`). |
| TeamsPostUrl | TeamsPostUrl | Hyperlink | No | URL of the Adaptive Card posted to Teams. Filled in by the poller flow after the post action. |
| PostedToTeams | PostedToTeams | Yes/No | Yes | Default `No`. Flipped to `Yes` after the Teams post succeeds. Drives retry-on-failure if the post failed previously. |
| DigestSent | DigestSent | Yes/No | Yes | Default `No`. Set to `Yes` by the digest flow when included in a digest email. Avoids double-counting if the digest runs more than once. |
| FirstSeenDateTime | FirstSeenDateTime | Date and time | Yes | Timestamp the poller first wrote this row. Distinct from Graph `startDateTime` (the actual incident start). |
| LastSeenDateTime | LastSeenDateTime | Date and time | Yes | Timestamp of the most recent poller run that touched this row. |

## Indexes

SharePoint allows up to 20 indexed columns. Create these:

1. `IncidentID` — primary dedupe lookup. Every poller run does `Get items` with `IncidentID eq '<id>'`.
2. `LastModifiedDateTime` — digest flow queries by date window.
3. `Service` — used by the digest grouping and by ad-hoc queries.

## Views

| View name | Filter | Sort | Purpose |
|---|---|---|---|
| Active incidents | `Status` is not in (`serviceRestored`, `postIncidentReviewPublished`, `serviceOperational`, `falsePositive`) | `LastModifiedDateTime` descending | Default landing view for IT ops. |
| Last 24 hours | `LastModifiedDateTime` >= `[Today]-1` | `LastModifiedDateTime` descending | Backs the daily digest content. |
| Pending Teams post | `PostedToTeams` = `No` | `FirstSeenDateTime` ascending | Catches incidents the poster failed to post on the first try. |
| All by service | Group by `Service` | `LastModifiedDateTime` descending | Trend review. |

## Dedupe + update logic

The poller flow performs this sequence for each item returned by Graph:

1. `Get items` from `ServiceHealthIncidents` with filter `IncidentID eq '<id>'` and `Top Count = 1`.
2. If zero results: `Create item` with all columns, set `PostedToTeams = No`. Post the Adaptive Card. On success, `Update item` to set `PostedToTeams = Yes` and `TeamsPostUrl`.
3. If one result and `LastModifiedDateTime` from Graph > stored `LastModifiedDateTime`: `Update item`, then post an "update" Adaptive Card.
4. If one result and `LastModifiedDateTime` matches: `Update item` to bump `LastSeenDateTime` only. No Teams post.

This makes the entire pipeline idempotent — the same Graph response can be replayed safely.

## Provisioning

A PowerShell script (PnP) provisions the list with the columns and views above. To be authored as `scripts/Setup-IncidentList.ps1` in the build phase. The lifecycle project's `Setup-SharePointList.ps1` is the template.

## Locked decisions

- **Site:** new site `ServiceOps`, separate from `ITAutomation`.
- **Retention:** deferred to a SharePoint retention policy rather than built into the poller. Resolved incidents stay in the list indefinitely for v1.
