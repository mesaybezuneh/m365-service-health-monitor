# Deliverable 2 — Power Automate Flows

Two flows. Both are scheduled. Both call the Microsoft Graph service health API via HTTP with Application authentication (client credentials). Detailed Graph payload shapes live in `docs/03-graph-api-reference.md`.

## Flow A — `Flow-ServiceHealth-Poller`

### Trigger
Recurrence — every 15 minutes. Time zone: tenant default. Start time: top of the hour after deployment.

### Variables (Phase A — Initialize)

| Name | Type | Value |
|---|---|---|
| varRunStartUtc | String | `utcNow()` |
| varSubscribedServices | Array | From `config.local.json`, hard-coded as a JSON array. Used to filter Graph results client-side because Graph supports `$filter=service eq 'X'` but not multi-value `in`. |
| varNewCount | Integer | `0` |
| varUpdatedCount | Integer | `0` |
| varSkippedCount | Integer | `0` |
| varFailedItems | Array | `[]` — collects incident IDs whose SharePoint or Teams write failed, surfaced in the failure email at the end. |

### Phase B — Pull current issues

1. `HTTP_GetIssues` — GET `https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues?$filter=isResolved eq false`
   - Authentication: `ActiveDirectoryOAuth`. Tenant + ClientID + Secret from the connection.
   - On HTTP 429: configured retry policy (exponential, 4 retries, honors `Retry-After`).
2. `Parse_JSON_Issues` — schema generated from a sample response. See `docs/03`.

### Phase C — Filter to subscribed services

3. `Filter_array_BySubscribed` — input `body('Parse_JSON_Issues')?['value']`. Condition: `contains(variables('varSubscribedServices'), item()?['service'])`.

### Phase D — Loop and reconcile

4. `Apply_to_each_Issue` over `body('Filter_array_BySubscribed')`. Inside:

   - `Compose_IncidentID` — `item()?['id']`
   - `Get_items_Existing` — SharePoint Get items from `ServiceHealthIncidents`, OData filter `IncidentID eq '@{outputs('Compose_IncidentID')}'`, Top Count 1.
   - `Compose_ExistingCount` — `length(body('Get_items_Existing')?['value'])`
   - `Switch_OnExisting` on `outputs('Compose_ExistingCount')`:

     **Case 0 — new incident**
     - `Compose_LatestPost` — `last(item()?['posts'])` (the most recent post in the Graph response).
     - `Compose_Severity` — see severity mapping below.
     - `SharePoint_CreateItem_Incident` — set every column from `docs/01`. `PostedToTeams = false`.
     - `HTTP_PostToTeamsWebhook` — POST to the Incoming Webhook URL with the Adaptive Card body from `docs/04`. Card type: `new`.
     - `SharePoint_UpdateItem_AfterPost` — set `PostedToTeams = true`, `TeamsPostUrl = body('HTTP_PostToTeamsWebhook')?['<urlField>']` if the webhook returns one; otherwise compose a synthetic deep link.
     - `Increment_NewCount`.

     **Case 1 — possible update**
     - `Compose_StoredLastModified` — `first(body('Get_items_Existing')?['value'])?['LastModifiedDateTime']`.
     - `Compose_GraphLastModified` — `item()?['lastModifiedDateTime']`.
     - `Condition_HasNewerUpdate` — `ticks(graph) > ticks(stored)`.
       - True branch: `SharePoint_UpdateItem_OnChange` (all dynamic columns + bump `LastSeenDateTime`), then `HTTP_PostToTeamsWebhook_Update` with card type `update`. Increment `varUpdatedCount`.
       - False branch: `SharePoint_UpdateItem_Touch` — set `LastSeenDateTime` only. Increment `varSkippedCount`.

   - Each Graph/SharePoint/Teams call has a parallel `Append_to_array_FailedItems` action with `runAfter` set to `[Failed, TimedOut]`. The loop continues. The failure is reported in the run summary email.

### Phase E — Recently resolved sweep

5. `HTTP_GetRecentlyResolved` — GET `/admin/serviceAnnouncement/issues?$filter=isResolved eq true and lastModifiedDateTime ge <now minus 1 hour>`. Catches incidents that flipped to resolved since the last poll.
6. Same loop pattern as Phase D but only updates existing rows. New resolved-only incidents (which were missed entirely) get created and immediately marked resolved.

### Phase F — Finalize

7. `Condition_HadFailures` — `length(variables('varFailedItems')) > 0`.
   - True: `Send_email_FailureSummary` to the operator with `varFailedItems` as a table.
8. `Compose_RunSummary` — counts. (Optional) write to a `PollerRunLog` SharePoint list or just emit via Compose for run-history visibility.

### Severity mapping

| Graph `status` | Severity column | Adaptive Card color |
|---|---|---|
| `serviceInterruption` | High | Attention (red) |
| `serviceDegradation` | Medium | Warning (yellow) |
| `restoringService`, `extendedRecovery` | Medium | Warning |
| `serviceRestored`, `serviceOperational`, `postIncidentReviewPublished` | Resolved | Good (green) |
| `falsePositive`, `investigationSuspended` | Info | Accent (blue) |
| All advisories (`classification: advisory`) | Low | Default |

---

## Flow B — `Flow-ServiceHealth-DailyDigest`

### Trigger
Recurrence — daily at 08:00 tenant time.

### Phase A — Query

1. `Compose_WindowStart` — `addHours(utcNow(), -24)`.
2. `Get_items_LastDay` — SharePoint Get items from `ServiceHealthIncidents`. OData filter: `LastModifiedDateTime ge datetime'@{outputs('Compose_WindowStart')}'`. Order by `Service`, `Severity`, `LastModifiedDateTime desc`.

### Phase B — Format

3. `Select_DigestRows` — Select action projecting to `{Service, IncidentID, Severity, Status, Title, LastModifiedDateTime, LatestPostType, Url}`.
4. `Create_HTML_table` — input is `body('Select_DigestRows')`.

### Phase C — Send

5. `Condition_HasAny` — `length(body('Get_items_LastDay')?['value']) > 0`.
   - True: `Send_an_email_v2_Digest` — To: `digestRecipients.to`. Subject: `M365 Service Health — 24h digest — @{formatDateTime(utcNow(),'yyyy-MM-dd')}`. Body: HTML with table + summary counts + link to the SharePoint "Active incidents" view.
   - False: do nothing. All-clear emails are suppressed (per locked decision below). The digest stays silent on zero-incident days.

### Phase D — Mark sent

6. `Apply_to_each_DigestRow` over `body('Get_items_LastDay')?['value']`. Inside: `SharePoint_UpdateItem_DigestSent` setting `DigestSent = true`.

---

## Cross-flow operational notes

- Both flows run under the same `M365-Service-Health-Monitor` Entra app registration via a single `Premier HTTP with Azure AD` connection so that throttling is shared and the secret is rotated in one place.
- Run history retention is 28 days by default in Power Automate — long enough for debugging. For audit purposes the SharePoint list is the source of truth.
- The poller is intentionally simple: no parallelism across pages. The `/issues` endpoint returns at most ~50 active items in a real-world tenant. Pagination via `@odata.nextLink` is unnecessary for active issues but should be implemented for `/messages` if the message-center backlog is ever pulled.
- The Teams Incoming Webhook is rate-limited at four messages per second per webhook. With 15-minute polling and typical incident volume this is never reached, but if the recovery sweep ever discovers a backlog the poller should `Delay` 0.3 s between posts.

## Locked decisions

- **Endpoint scope:** v1 pulls `/admin/serviceAnnouncement/issues` only. The message center (`/messages`) is a v2 expansion. Keeps the build tight and the schema single-shape.
- **Digest cadence:** daily at 08:00 tenant time, every day. Weekend incidents are not deferred to Monday.
- **All-clear email:** suppressed. Zero incidents in the last 24 hours = no email. `varSendAllClear` defaults to `false`.
- **Sev-1 SMS/phone escalation:** out of scope for v1. Adaptive Card severity color is the only urgency signal.
