# M365 Service Health Monitoring & Alerting

A Power Automate-based monitor for Microsoft 365 service incidents. Polls the Microsoft Graph service health API on a schedule, deduplicates against a SharePoint incident log, and routes new or updated incidents to a Teams channel in real time and to an email digest once a day.

Sister project to [m365-user-lifecycle-automation](https://github.com/mesaybezuneh/m365-user-lifecycle-automation).

**Author:** Mesay Bezuneh

## Problem statement

The Microsoft 365 admin center shows service health to anyone who logs in. It does not push. If an Exchange incident is opened at 03:00 and IT operations does not log in until 09:00, six hours of user-impact tickets accumulate against an incident Microsoft already knew about. This project closes that loop with two automations:

- A poller flow checks the Graph service health API every 15 minutes, posts new or updated incidents to a dedicated Teams channel, and records them in a SharePoint list.
- A daily digest flow emails IT operations a summary of the prior 24 hours.

## Architecture

```
                       Microsoft Graph
                  serviceAnnouncement API
                            |
            +---------------+---------------+
            |                               |
   /issues (active incidents)      /messages (advisories)
            |                               |
            +---------------+---------------+
                            |
                    Flow-ServiceHealth-Poller
                       (every 15 min)
                            |
            +---------------+---------------+---------------+
            |                               |               |
   SharePoint incident log         Teams Incoming         Failure
   (dedupe + history)              Webhook channel        email
                                  (Adaptive Card)
                            |
                            v
                    Flow-ServiceHealth-DailyDigest
                       (08:00 daily)
                            |
                            v
                    IT-Ops mailing list
                    (HTML summary table)
```

## Tech stack

- Microsoft Graph (Application permissions: `ServiceHealth.Read.All`, `ServiceMessage.Read.All`)
- Power Automate (HTTP + SharePoint + Teams + Office 365 Outlook connectors)
- SharePoint Online (incident log)
- Microsoft Teams (Incoming Webhook)
- PowerShell + PnP (one-time list provisioning, app registration)

## Deliverables

| # | Document | Purpose | Status |
|---|---|---|---|
| 1 | [`docs/01-sharepoint-schema.md`](docs/01-sharepoint-schema.md) | IncidentLog column set, indexes, views | Done |
| 2 | [`docs/02-power-automate-flows.md`](docs/02-power-automate-flows.md) | Poller and digest flow design | Done |
| 3 | [`docs/03-graph-api-reference.md`](docs/03-graph-api-reference.md) | Graph endpoints, scopes, payloads | Done |
| 4 | [`docs/04-teams-webhook-setup.md`](docs/04-teams-webhook-setup.md) | Webhook config + Adaptive Card JSON | Done |

## Build status

| Phase | Description | Status |
|---|---|---|
| 0 | Project scaffold (this README, CLAUDE.md, gitignore) | Done |
| 1 | SharePoint list provisioned | Done (2026-05-22) |
| 2 | App registration + Graph permissions granted + admin consent | Done (2026-05-22) |
| 3 | Workflows-for-Teams webhook captured (legacy Incoming Webhook retired Dec 2025) | Done (2026-05-22) |
| 4 | Poller flow built and verified end to end | Done (2026-05-22) |
| 5 | Dedupe + update-detection verified | Done (2026-05-22) |
| 6 | Daily digest flow built and verified | Done (2026-05-22) |
| 7 | Failure path verified (force a Graph 401 or 429) | Done (2026-05-22) |

## Test log

| # | Date | Scenario | Result |
|---|---|---|---|
| 1 | 2026-05-22 | `Setup-IncidentList.ps1` against `cloudopslabs.sharepoint.com` — site `ServiceOps` created, list `ServiceHealthIncidents` (GUID `ec49d75d-b5a7-4663-bf56-3cd996c7cd73`) provisioned with 17 columns, 3 indexes, 5 views | Pass |
| 2 | 2026-05-22 | `Setup-AppRegistration.ps1 -CreateSecret` — app `M365-Service-Health-Monitor` (`5efeb71c-...`) created, `ServiceHealth.Read.All` + `ServiceMessage.Read.All` granted with admin consent, 6-month secret minted | Pass |
| 3 | 2026-05-22 | Client-credentials smoke test against `GET /v1.0/admin/serviceAnnouncement/issues?$top=3` — token issued, 200 OK, 3 items returned (first ID `CW1218323`) | Pass |
| 4 | 2026-05-22 | `Setup-ServiceAccount.ps1` — `svc-servicehealth@cloudopslabs.onmicrosoft.com` created, SPB license assigned | Pass |
| 5 | 2026-05-22 | `Setup-TeamsChannel.ps1` — `IT Operations` team + `M365 Service Alerts` standard channel created; service account added as team member | Pass |
| 6 | 2026-05-22 | Workflows template `Send webhook alerts to a channel` configured (renamed to `ServiceHealth-PosterBot` in Power Automate) | Pass |
| 7 | 2026-05-22 | `POST` of `sample-adaptive-card.json` to Workflows webhook URL — 202 Accepted, run id `08584221332674337719687992176CU18`, card rendered in channel (see `docs/screenshots/phase-3-smoke-test.png`) | Pass |
| 8 | 2026-05-22 | Phase 4 Test 1 (happy path) — Recurrence triggered manually via REST. After fixing `item/EndDateTime` null-handling (use bare `@...` not `@{...}` to preserve null type), 3 active Exchange issues pulled, 2 new SharePoint rows + 2 Teams cards created, 1 row updated (pre-existing partial). 0 failures. | Pass |
| 9 | 2026-05-22 | Phase 4 Test 2 (dedupe) — re-triggered. First attempt failed semantically (3 updated instead of 3 skipped). Root cause: SharePoint truncates Graph's millisecond-precision `LastModifiedDateTime` to second-precision on storage, so `ticks(graph) > ticks(stored)` returned true by ~99 ms. Fixed by adding 60-second tolerance: `add(ticks(stored), 600000000)`. Retest: 0 new, 0 updated, **3 skipped**. | Pass |
| 10 | 2026-05-22 | Phase 4 Test 3 (update detection) — backdated EX1279815's `LastModifiedDateTime` by 2 hours via PnP, re-triggered. Result: **1 updated** (EX1279815), 2 skipped, 1 "Update" Teams card posted, stored timestamp restored from Graph. | Pass |
| 11 | 2026-05-22 | Adaptive Card visual polish — `color: "Light"` removed (it was meant for dark backgrounds Teams doesn't paint), title text now uses severity-mapped color (`Attention`/`Warning`/`Good`/`Accent`), `bleed: true` extends container tint edge-to-edge, title weight `Bolder`. | Pass |
| 12 | 2026-05-22 | Phase 6 Test 1 (digest happy path) — `Flow-ServiceHealth-DailyDigest` (GUID `9431c274-...`) triggered. `Get_items_LastDay` filter returned 1 of 3 rows (EX1310533, only one with Microsoft-updated `LastModifiedDateTime` < 24h); email sent to `mesay@cloudopstech.net`; `DigestSent` flipped to `True` on that row only. | Pass |
| 13 | 2026-05-22 | Phase 6 Test 2 (empty window) — backdated EX1310533's `LastModifiedDateTime` to 25h ago via PnP, re-triggered. Result: `Compose_RowCount=0`, `Condition_HasAny` else branch, `Send_email_v2_Digest` / `Apply_to_each_DigestRow` both Skipped. All-clear suppression verified. | Pass |
| 14 | 2026-05-22 | Phase 7 (failure path) — corrupted `clientSecret` in `HTTP_GetIssues` and PATCHed. Result: HTTP_GetIssues=Failed, `Append_FailedItems_GetIssues` succeeded, all loop actions Skipped, `Compose_RunSummary` reported `failureCount=1`, `Condition_HadFailures` true branch, `Send_email_FailureSummary` succeeded. Failure email sent to `mesay@cloudopstech.net`. Discovered v2 polish item: `body()` and `outputs()` on a failed action return empty; should use `actions('HTTP_GetIssues')?['outputs']?['statusCode']` and `?['error']?['message']` instead. | Pass |
| 15 | 2026-05-22 | Phase 7 recovery — restored real secret, PATCHed, re-triggered. Result: 0 new, 0 updated, 3 skipped, 0 failures. Flow self-heals on the next poll. | Pass |

## Setup (placeholder — fill in as build progresses)

1. Copy `config.example.json` to `config.local.json` and fill in tenant values.
2. Register the `M365-Service-Health-Monitor` app in Entra. Grant `ServiceHealth.Read.All` and `ServiceMessage.Read.All` Application permissions. Admin-consent.
3. Provision the `ServiceHealthIncidents` SharePoint list per `docs/01-sharepoint-schema.md`.
4. Create the Teams channel and Incoming Webhook. Record the URL in `config.local.json`.
5. Build the two flows per `docs/02-power-automate-flows.md`.

## Disclaimer

This project is a proof-of-concept built in a Microsoft 365 Developer Program tenant and is provided as-is, without warranty, for educational and reference purposes only. It is not production-hardened. Specifically:

- The Microsoft Graph client secret is embedded inline in the Power Automate flow JSON. The flow JSON is gitignored locally and never committed, but the production-correct pattern is to externalize secrets to Azure Key Vault and reference them via a managed identity.
- No automated secret-rotation is wired up. The 6-month expiry is tracked manually.
- The SharePoint and Office 365 Outlook connections authenticate as a user identity rather than a managed identity. If that user leaves the tenant, the flow breaks.
- The failure-summary email path itself has no fallback; if Office 365 Outlook is the failing component, no notification is sent.
- All tenant identifiers (site URLs, list GUIDs, app object IDs, webhook URLs) are tenant-specific. They will not work in another tenant without re-provisioning every artifact via the included setup scripts.

Before adapting any pattern from this repo for a production tenant, perform a security review and replace each of the above with the production-correct alternative.

## License

MIT. See [LICENSE](LICENSE).
