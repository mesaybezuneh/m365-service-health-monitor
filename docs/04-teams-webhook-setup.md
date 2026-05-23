# Deliverable 4 — Teams Webhook + Adaptive Card

## Transport (2026): Workflows for Teams, not legacy Incoming Webhook

Microsoft retired Office 365 Connectors (which provided the classic "Incoming Webhook" Teams channel connector) in December 2025. The supported replacement is **Workflows for Teams** — a Power Automate template embedded inside the Teams client. The runtime characteristics are similar (HTTPS POST a JSON body, get a channel post), but the URL host, template name, and connection model are different.

| | Legacy Incoming Webhook (retired) | Workflows for Teams (current) |
|---|---|---|
| Where you create it | Channel ⋯ → Connectors → Incoming Webhook | Channel ⋯ → Workflows → template "Send webhook alerts to a channel" |
| URL host | `<tenant>.webhook.office.com` | `<env>.<region>.environment.api.powerplatform.com` |
| Auth | Anonymous (URL is the secret) | Anonymous (URL is the secret) — same threat model |
| Identity that posts | "Service Health Bot" (or whatever you named the connector) | The flow owner's identity, surfaced as the workflow name |
| Throttling | Microsoft-managed, lenient | Power Automate per-flow throttling (60 calls / min per flow) |
| Threading replies | No | No (the channel connector posts new top-level messages only) |
| Programmatic creation | Possible via Graph (deprecated) | No documented Graph API — manual setup in the Teams client |

## Channel setup

1. Create or reuse a Teams team. In this project: `IT Operations`.
2. Inside the team, create a standard channel. In this project: `M365 Service Alerts`.
3. The team owner and the service account (`svc-servicehealth@<tenant>.onmicrosoft.com`) must both be members. The service account becomes the flow co-owner so the webhook survives if the human owner leaves the org.

The team and channel are provisioned by `scripts/Setup-TeamsChannel.ps1` (uses `Microsoft.Graph` PowerShell, not `Az.Accounts` — Az's first-party app token does not carry `TeamMember.*` scopes, so `Get-MgTeamMember` / `New-MgTeamMember` fail with 403 under Az auth).

## Workflows-based webhook setup

These steps cannot be scripted as of 2026-05-22 — Microsoft does not expose a Graph API to instantiate the "Send webhook alerts to a channel" Workflows template. They are manual, in the Teams desktop or web client, signed in as a human admin.

1. Open the `M365 Service Alerts` channel.
2. Channel `⋯` menu → `Workflows`.
3. Search: `webhook`. Pick **"Send webhook alerts to a channel"** (Microsoft publisher). Microsoft sometimes shows two identically-named entries; either works.
4. Sign in to the **Microsoft Teams** and **Notifications** connections when prompted.
5. Click `Next`. On the configuration screen:
   - Workflow name: leave the default (the in-Teams rename field is flaky). Rename later in Power Automate.
   - Team: `IT Operations`
   - Channel: `M365 Service Alerts`
6. Click `Add workflow`. Teams shows the generated webhook URL.
7. Copy the URL. Save it to `config.local.json` under `teams.webhookUrl`. Do not commit.
8. Open [make.powerautomate.com](https://make.powerautomate.com) → My flows → rename the flow to `ServiceHealth-PosterBot`.
9. Same flow → Share → add `svc-servicehealth` as co-owner. (Optional: remove yourself after the flow is confirmed working under the service account.)

The generated URL host in this tenant looks like:

```
https://default<env-guid-no-hyphens>.10.environment.api.powerplatform.com:443
    /powerautomate/automations/direct/workflows/<flow-guid>
    /triggers/manual/paths/invoke
    ?api-version=1
    &sp=%2Ftriggers%2Fmanual%2Frun
    &sv=1.0
    &sig=<signature>
```

The `sig` query parameter is the bearer credential. Anyone with the full URL can post to the channel anonymously. Treat as a secret.

### Regenerating the URL (rotation / leak response)

1. Open the flow in Power Automate.
2. Edit → click the trigger ("When a Teams webhook request is received").
3. There is no rotate-button. To invalidate the existing URL: delete the trigger, save, edit again, add a new HTTP webhook trigger. The URL changes.
4. Update `config.local.json` and the poller flow's `Post_to_Teams` HTTP action.

For a more disciplined rotation workflow, recreate the entire flow (~3 minutes via Teams ⋯ → Workflows) and update both consumers.

## Post action (poller flow)

In the poller flow, use the HTTP action:

```
POST <webhook URL>
Content-Type: application/json
```

The body shape mirrors the legacy Incoming Webhook for compatibility — the Workflows template parses the `attachments[].content` Adaptive Card and posts it as a channel message:

```json
{
  "type": "message",
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.card.adaptive",
      "contentUrl": null,
      "content": { /* Adaptive Card body */ }
    }
  ]
}
```

A successful POST returns **HTTP 202 Accepted** with `x-ms-workflow-run-id` and `x-ms-correlation-id` headers. Failure responses are JSON `error.code` + `error.message`.

## Card — New incident

The card is parameterized by Graph response fields. Colors come from the severity mapping in `docs/02`.

```json
{
  "type": "message",
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.card.adaptive",
      "contentUrl": null,
      "content": {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "msteams": { "width": "Full" },
        "body": [
          {
            "type": "Container",
            "style": "attention",
            "items": [
              {
                "type": "TextBlock",
                "size": "Large",
                "weight": "Bolder",
                "color": "Light",
                "text": "New incident - @{items('Apply_to_each_Issue')?['service']}"
              },
              {
                "type": "TextBlock",
                "spacing": "None",
                "color": "Light",
                "text": "@{items('Apply_to_each_Issue')?['title']}",
                "wrap": true
              }
            ]
          },
          {
            "type": "FactSet",
            "facts": [
              { "title": "Incident ID",    "value": "@{items('Apply_to_each_Issue')?['id']}" },
              { "title": "Status",         "value": "@{items('Apply_to_each_Issue')?['status']}" },
              { "title": "Classification", "value": "@{items('Apply_to_each_Issue')?['classification']}" },
              { "title": "Service",        "value": "@{items('Apply_to_each_Issue')?['service']}" },
              { "title": "Feature",        "value": "@{coalesce(items('Apply_to_each_Issue')?['feature'], '-')}" },
              { "title": "Started",        "value": "@{formatDateTime(items('Apply_to_each_Issue')?['startDateTime'], 'yyyy-MM-dd HH:mm') } UTC" }
            ]
          },
          {
            "type": "TextBlock",
            "text": "**Impact**",
            "wrap": true,
            "spacing": "Medium"
          },
          {
            "type": "TextBlock",
            "text": "@{coalesce(items('Apply_to_each_Issue')?['impactDescription'], 'Not yet described by Microsoft.')}",
            "wrap": true
          },
          {
            "type": "TextBlock",
            "text": "**Latest post**",
            "wrap": true,
            "spacing": "Medium"
          },
          {
            "type": "TextBlock",
            "text": "@{last(items('Apply_to_each_Issue')?['posts'])?['description']?['content']}",
            "wrap": true,
            "isSubtle": true
          }
        ],
        "actions": [
          {
            "type": "Action.OpenUrl",
            "title": "Open in M365 admin center",
            "url": "https://admin.microsoft.com/Adminportal/Home#/servicehealth/:/alerts/@{items('Apply_to_each_Issue')?['id']}"
          }
        ]
      }
    }
  ]
}
```

## Card — Update to existing incident

Same body but:

- Header text becomes `Update - @{...['service']}` and container `style` becomes `warning`.
- Insert a fact for `Last update` showing `formatDateTime(item()?['lastModifiedDateTime'], 'yyyy-MM-dd HH:mm') UTC`.
- Below the FactSet, show only the most recent `posts[]` entry (the same `last(...)` expression works because Graph returns posts in chronological order).

## Card — Resolved

- Header text `Resolved - @{...['service']}`. Container `style` becomes `good`.
- Add fact `Ended` from `endDateTime`.
- Add fact `Duration` computed from `endDateTime - startDateTime`. Expression: `div(sub(ticks(item()?['endDateTime']), ticks(item()?['startDateTime'])), 600000000)` for minutes, then format.

## Severity to container style mapping

The Adaptive Card 1.4 `Container.style` accepts: `default`, `emphasis`, `good`, `attention`, `warning`, `accent`. Mapping:

| Severity (from `docs/02`) | Container style |
|---|---|
| High     | `attention` |
| Medium   | `warning`   |
| Low      | `emphasis`  |
| Resolved | `good`      |
| Info     | `accent`    |

## Testing the card

Quickest path before wiring the poller flow:

```powershell
$config = Get-Content -Raw .\config.local.json | ConvertFrom-Json
$webhook = $config.teams.webhookUrl
$cardJson = Get-Content -Raw .\scripts\sample-adaptive-card.json
Invoke-WebRequest -Method POST -Uri $webhook -ContentType 'application/json' -Body $cardJson -UseBasicParsing
```

Expected response: `StatusCode = 202`. Confirm the card appears in the `M365 Service Alerts` channel within 5-30 seconds.

Use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to iterate on the card visually before pasting into the flow JSON.

A reference smoke-test screenshot is at `docs/screenshots/phase-3-smoke-test.png`.

## Open questions for the build phase

- Mention/@-tag the on-call user on `attention`-style cards? Adaptive Cards support `msteams.entities` for mentions but it requires the user's AAD object ID in the card payload. The Workflows template's Adaptive Card renderer respects `msteams.entities` the same way the legacy connector did. Defer to v2 unless the dev tenant has a designated on-call.
- Thread updates as replies to the original post? The Workflows template posts new top-level channel messages — no threading API. If threading is required, switch the post action to the Power Automate `Microsoft Teams` connector's `Reply with a message in a channel` action, which can target a specific parent message by ID.

## Locked decisions

- **Transport:** Workflows for Teams ("Send webhook alerts to a channel" template). Reason: legacy Incoming Webhook was fully retired by Microsoft in December 2025.
- **Owner of the Workflows flow:** the dedicated service account `svc-servicehealth@<tenant>.onmicrosoft.com`. Reason: the URL outlives any human's tenure in the org.
- **URL storage:** `config.local.json` (gitignored). Reason: the URL is a bearer-style credential.
- **Card format:** Adaptive Card 1.4, wrapped in the `{type:"message", attachments:[...]}` envelope. Reason: that envelope is what the template's body parser expects. Verified by the Phase 3 smoke test (HTTP 202, card rendered).
