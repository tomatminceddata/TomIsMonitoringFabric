# Power BI Tenant Lens — PowerShell Guide

---

## Power BI Tenant Lens — Community Solution

This solution is provided as-is, as a community contribution. It has been developed iteratively against real-world tenants with up to 45,000+ workspaces and battle-tested through hundreds of hours of production runs. Every resilience pattern in the code traces back to an actual failure encountered and fixed.

That said, this is not a Microsoft product and comes with no warranty or support agreement. A few things to keep in mind:

Microsoft APIs change. The Scanner API, Admin APIs, and Fabric APIs evolve continuously. Endpoints may be deprecated, throttling limits adjusted, or response schemas modified without notice. What works today may need adaptation tomorrow.

Data completeness depends on permissions and API behavior. The inventory reflects what the APIs return for your service principal's scope. Gaps in tenant settings, datasource details, or lineage data may exist due to API limitations — some of which are known and documented, others yet to be discovered.

This is not a migration tool. It is an inventory and analysis foundation. The solution includes a semantic model that makes the inventory data immediately explorable in Power BI. Decisions about workspace consolidation, capacity planning, or artifact migration should be validated by your own governance processes and stakeholders. The data supports those decisions — it does not make them for you.

I maintain this solution because I believe in it, and I will continue to fix issues and add features as time permits. But "as time permits" is the operative phrase — this is community work, not a product with an SLA.

If you find issues or have ideas, I welcome the conversation. If you build something great on top of it, even better.

Tom Martens — Data Platform MVP

---

**PowerBITenantLens.ps1**
Wide Scan + Scanner + Refreshables + Dashboards + Apps + Access + Tiles

Author: Tom Martens
Version: 2026-03-06

---

## What This Script Does

The PowerBITenantLens inventory script provides a comprehensive snapshot of a Power BI tenant by combining multiple Microsoft APIs into a single, menu-driven tool. It produces a set of CSV files that can be loaded into a Power BI semantic model for analysis, migration planning, and governance.

The script addresses known gaps in individual Microsoft APIs by layering them together: GetGroupsAsAdmin for complete workspace enumeration, the Scanner API for detailed artifact metadata (including datasource instances from the root level, dataflow-to-dataflow lineage, and dashboards), the Refreshables API for dataset refresh health, GetAppsAsAdmin for tenant-wide app inventory, GetAppUsersAsAdmin for per-app access details, and GetTilesAsAdmin for per-dashboard tile detail.

The script extends previous versions with two additional tables and run modes: 15_AppAccess (per-app access records via GetAppUsersAsAdmin) and 16_DashboardTiles (per-tile detail via GetTilesAsAdmin). Both steps use rate-limited per-item APIs (200 requests/hour) with checkpoint/resume, making them long-running but safe to interrupt and restart. Step 43 provides a fast-run mode that skips these two long-running steps for daily refresh use.

---

## Prerequisites

- **PowerShell 7** or higher (run with `pwsh`)
- **MicrosoftPowerBIMgmt** module (auto-installed if missing)
- **Power BI Service Administrator** or **Fabric Administrator** role
- Network access to `api.powerbi.com` and `api.fabric.microsoft.com`

**Cross-platform:** The script runs on both Windows and macOS. On macOS, install PowerShell via `brew install powershell`, then run `pwsh`. Authentication uses device code flow automatically — open the provided URL, enter the code, and sign in via browser.

---

## Quick Start

```powershell
# Show the interactive menu
pwsh ./PowerBITenantLens.ps1

# Run the full inventory (all steps, ~35h for large tenants)
pwsh ./PowerBITenantLens.ps1 -Step 0

# Fast run — all steps EXCEPT 15 + 16 (~1–2h)
pwsh ./PowerBITenantLens.ps1 -Step 43

# Quick test with first 1000 workspaces
pwsh ./PowerBITenantLens.ps1 -Step 42

# App Access only (~7h, checkpoint every 50 apps)
pwsh ./PowerBITenantLens.ps1 -Step 15

# Dashboard Tiles only (~26h, checkpoint every 50 dashboards)
pwsh ./PowerBITenantLens.ps1 -Step 16
```

---

## Run Modes

| Step | What it runs | Typical duration | Use case |
|------|-------------|-----------------|----------|
| `0` | Everything | ~35h (large tenant) | Full weekend run |
| `43` | Everything except 15 + 16 | ~1–2h | Daily/on-demand refresh |
| `15` | App Access only | ~7h | Overnight, standalone |
| `16` | Dashboard Tiles only | ~26h | Overnight/weekend, standalone |
| `42` | First 1000 workspaces (steps 2–6, 13) | ~5 min | Quick test |

**Recommended workflow:** Run Step 43 daily for fresh workspace/report/dataset/datasource data. Schedule Step 15 and 16 separately overnight or over a weekend. The semantic model picks up whichever CSVs are present — empty tables load gracefully for missing steps.

---

## Menu Options

| Step | Description | API Used | Duration (Large Tenant) |
|------|-------------|----------|------------------------|
| 0 | Run ALL steps | All | ~35h |
| 1 | Tenant Settings | Fabric Admin API | Seconds |
| 2 | Workspaces | GetGroupsAsAdmin + Scanner | ~10–16 min |
| 3 | Workspace Members | Scanner API | Included in Step 2 scan |
| 4 | Reports | Scanner API | Included in Step 2 scan |
| 5 | Semantic Models | Scanner API | Included in Step 2 scan |
| 6 | Data Sources | Scanner API | Included in Step 2 scan |
| 7 | Gateways | Gateways API | Seconds |
| 11 | Refreshables | Admin Refreshables API | ~1–2 min |
| 13 | Dashboards | Scanner API | Included in Step 2 scan |
| 14 | Apps | GetAppsAsAdmin API | Seconds |
| 15 | App Access | GetAppUsersAsAdmin (per-app) | ~7h (200 req/hr) |
| 16 | Dashboard Tiles | GetTilesAsAdmin (per-dashboard) | ~26h (200 req/hr) |
| 42 | First 1000 workspaces | Steps 2–6 + 13 | ~2 min |
| 43 | Fast run (all except 15 + 16) | All except per-item APIs | ~1–2h |

**Step dependencies:** Steps 2–6 and 13 all require the Scanner API scan (GetGroupsAsAdmin → PostWorkspaceInfo → GetScanStatus → GetScanResult). Running any of these steps triggers the full scan pipeline. Steps 1, 7, 11, and 14 are independent and can be run individually without scanning. Step 15 requires Step 14 output (apps list). Step 16 requires Step 13 output (dashboards list). Both 15 and 16 will load from existing CSVs if their parent step wasn't run in the same session.

---

## Output Files

All CSV files are timestamped (e.g., `02_Workspaces_2026-03-04_183052.csv`) and written to the output directory (defaults to current directory).

### 01_TenantSettings

Admin portal tenant settings from the Fabric Admin API.

| Column | Description |
|--------|-------------|
| SettingName | Internal setting identifier |
| Title | Human-readable setting name |
| Enabled | Whether the setting is active |
| CanSpecifySecurityGroups | Whether security groups can be specified |
| EnabledSecurityGroups | Semi-colon separated list of security group IDs |
| TenantSettingGroup | Category grouping |
| Properties | JSON string of additional properties |

### 02_Workspaces

Complete workspace inventory from GetGroupsAsAdmin, enriched with Scanner API coverage information.

| Column | Description |
|--------|-------------|
| WorkspaceId | Unique workspace identifier (GUID) |
| WorkspaceName | Display name |
| Type | Workspace type (Workspace, PersonalGroup, etc.) |
| State | Active, Deleted, Removing, etc. |
| IsReadOnly | Read-only flag |
| IsOnDedicatedCapacity | Whether assigned to Premium/Fabric capacity |
| CapacityId | Assigned capacity GUID |
| Description | Workspace description |
| DefaultDatasetStorageFormat | Small or Large |
| HasWorkspaceLevelSettings | Whether workspace has custom settings |
| CapacityMigrationStatus | Migration status if applicable |
| DataflowStorageId | Custom dataflow storage if configured |
| PipelineId | Deployment pipeline association |
| InScannerAPI | Whether the Scanner API can see this workspace |

**Key insight:** The `InScannerAPI` flag reveals workspaces that GetGroupsAsAdmin returns but the Scanner API misses — typically A-SKU embedded capacity workspaces and certain SPN personal workspaces.

### 03_WorkspaceMembers

All workspace role assignments from the Scanner API.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| WorkspaceType / WorkspaceState | Workspace metadata |
| UserIdentifier | User or service principal identifier |
| UserPrincipalName | UPN (email) for user accounts |
| DisplayName | Friendly display name |
| AccessRight | Admin, Member, Contributor, or Viewer |
| PrincipalType | User, Group, or App |
| GraphId | Entra ID object identifier |

### 04_Reports

All reports across the tenant.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| ReportId | Unique report identifier |
| ReportName | Display name |
| ReportType | PowerBIReport, PaginatedReport, etc. |
| DatasetId | Connected semantic model |
| CreatedDateTime / ModifiedDateTime | Timestamps |
| CreatedBy / ModifiedBy | User identifiers |

### 05_SemanticModels

All datasets (semantic models) across the tenant.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| DatasetId | Unique dataset identifier |
| DatasetName | Display name |
| ConfiguredBy | User who configured the dataset |
| CreatedDate | Creation timestamp |
| IsRefreshable | Whether the dataset supports refresh |
| IsOnPremGatewayRequired | Whether an on-premises gateway is needed |
| TargetStorageMode | Import, DirectQuery, Dual, etc. |
| ContentProviderType | How the dataset was created |

### 06_DataSources

Dataset-to-datasource mappings from the Scanner API. Each row links a dataset to a datasource via `datasourceUsages`, with connection details resolved from the root-level `datasourceInstances` array. Additionally, rows from table 10 (DatasetUpstreamDataflows) are appended with `DatasourceType = "Dataflow"` and `DataSource = "DatasetUpstreamDataflows"`, giving a unified view of all datasource dependencies per dataset.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| DatasetId / DatasetName | Parent dataset |
| DatasourceId | Datasource instance identifier (or DataflowId for dataflow rows) |
| DatasourceType | Sql, SharePointList, Web, File, Oracle, Dataflow, etc. |
| GatewayId | Associated gateway (if any) |
| ConnectionDetails | JSON string of connection properties |
| Server / Database / Url / Path | Extracted connection fields |
| TenantName | Tenant name (from Power Query parameter) |
| DataSource | Origin: "Scanner API" or "DatasetUpstreamDataflows" |

**Key change (2026-03-04):** The `datasourceInstances` array is at the **root level** of the Scanner API scan result, not nested per-workspace. This was confirmed by Microsoft (Meir Shimon). All datasource rows now have fully populated DatasourceType and ConnectionDetails directly from the Scanner API — the previous per-dataset API workaround (table 07_3) is no longer needed.

### 07_1_GatewayClusters

On-premises and VNet gateway cluster definitions.

| Column | Description |
|--------|-------------|
| GatewayId | Gateway cluster identifier |
| GatewayName | Display name |
| GatewayType | Resource, Personal, etc. |
| GatewayAnnotation | JSON configuration metadata |
| PublicKey_Exponent / PublicKey_Modulus | Encryption public key |
| GatewayStatus | Live or Offline |
| GatewayMachineCount | Number of machines in the cluster |

### 07_2_GatewayDatasources

Datasource definitions registered on gateway clusters.

| Column | Description |
|--------|-------------|
| GatewayId / GatewayName | Parent gateway cluster |
| GatewayDatasourceId | Datasource definition identifier |
| DatasourceName | Display name |
| DatasourceType | SQL, Oracle, SAP, etc. |
| ConnectionDetails | JSON connection string |
| CredentialType | OAuth2, Windows, Key, etc. |
| CredentialDetails_UseEndUserOAuth2Credentials | SSO delegation flag |
| GatewayDatasourceStatus | Connection status |


### 08_Dataflows

All dataflows across the tenant.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| DataflowId | Unique dataflow identifier |
| DataflowName | Display name |
| Description | Dataflow description |
| ConfiguredBy / ModifiedBy | User identifiers |
| ModifiedDateTime | Last modification timestamp |
| Endorsement | Promoted, Certified, or empty |
| CertifiedBy | Who certified the dataflow |
| SensitivityLabel | Sensitivity label ID if applied |

### 09_DataflowDatasources

Dataflow-to-datasource mappings, structured identically to table 06.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| DataflowId / DataflowName | Parent dataflow |
| DatasourceId | Datasource instance identifier |
| DatasourceType | Connection type |
| GatewayId | Associated gateway |
| ConnectionDetails | JSON connection string |
| Server / Database / Url / Path | Extracted connection fields |

### 10_DatasetUpstreamDataflows

Lineage links between datasets and their upstream dataflows.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Dataset's workspace |
| DatasetId / DatasetName | The dataset |
| TargetDataflowId | Upstream dataflow identifier |
| DataflowWorkspaceId | Workspace containing the dataflow |

### 11_Refreshables

Dataset refresh health from the Admin Refreshables API (dedicated capacity only).

| Column | Description |
|--------|-------------|
| DatasetId / DatasetName | The dataset |
| Kind | Dataset kind |
| ConfiguredBy | Who configured the dataset |
| WorkspaceId / WorkspaceName | Parent workspace |
| CapacityId / CapacityName / CapacitySku / CapacityState | Capacity details |
| LastRefreshStatus | Completed, Failed, Disabled, Unknown |
| LastRefreshType | Scheduled, OnDemand, etc. |
| LastRefreshStartTime / LastRefreshEndTime | Refresh window |
| LastRefreshDurationSec | Duration in seconds |
| LastRefreshRequestId | Correlation ID for troubleshooting |
| LastRefreshServiceError | JSON error details if failed |
| WindowStartTime / WindowEndTime | Observation window |
| RefreshCount / RefreshFailures | Counts within the window |
| RefreshesPerDay | Average daily refresh frequency |
| AverageDurationSec / MedianDurationSec | Performance statistics |
| ScheduleEnabled | Whether scheduled refresh is on |
| ScheduleDays / ScheduleTimes | Schedule configuration |
| ScheduleTimeZone | Time zone for schedule |
| ScheduleNotifyOption | Notification preference |

### 12_DataflowUpstreamDataflows

Lineage links between dataflows and their upstream dataflows. This is the dataflow-to-dataflow equivalent of table 10 (which covers dataset-to-dataflow). Discovered on 2026-03-04: the Scanner API returns `upstreamDataflows` on dataflow objects (not just datasets), enabling full lineage chain reconstruction.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Dataflow's workspace |
| DataflowId / DataflowName | The consuming dataflow (e.g. _transform) |
| TargetDataflowId | Upstream dataflow identifier (e.g. _extract) |
| TargetDataflowWorkspaceId | Workspace containing the upstream dataflow |

**Example chain:** Dataset `eBFgoesPowerBI emailStorage` → `emailStorage_TreatyUW_transform` (table 10) → `emailStorage_TreatyUW_extract` (table 12) → Oracle ORADLCMP19 (table 09)

### 13_Dashboards

All dashboards across the tenant, extracted from the Scanner API scan result (`dashboards` array per workspace). Tile-level detail is flattened into summary columns. For full per-tile detail, see table 16_DashboardTiles.

| Column | Description |
|--------|-------------|
| WorkspaceId / WorkspaceName | Parent workspace |
| DashboardId | Unique dashboard identifier |
| DashboardName | Display name |
| IsReadOnly | Read-only flag |
| TileCount | Number of tiles on the dashboard |
| TileReportIds | Semicolon-delimited unique report IDs referenced by tiles |
| TileDatasetIds | Semicolon-delimited unique dataset IDs referenced by tiles |
| Endorsement | Promoted, Certified, or empty |
| CertifiedBy | Who certified the dashboard |
| SensitivityLabel | Sensitivity label ID if applied |

**API source:** Dashboards are part of the Scanner API scan result — no additional API calls are needed. They are extracted from the same `$scanResult.workspaces` data that produces reports, datasets, and dataflows.

### 14_Apps

All published apps across the tenant, retrieved via the `GetAppsAsAdmin` endpoint. Each app links to the workspace it was published from.

| Column | Description |
|--------|-------------|
| AppId | Unique app identifier |
| AppName | Display name |
| Description | App description |
| PublishedBy | User who published the app |
| LastUpdate | Last publication timestamp |
| WorkspaceId | Source workspace (joins to 02_Workspaces) |

**API source:** Apps are not part of the Scanner API scan result. They use a separate admin endpoint (`GET /v1.0/myorg/admin/apps`) with `$top=5000` pagination. This is a fast bulk endpoint — even tenants with thousands of apps complete in seconds.

### 15_AppAccess

Per-app access records showing which principals (users, groups, service principals) have access to each app. Retrieved via `GetAppUsersAsAdmin` — a per-app endpoint that requires one API call per app.

| Column | Description |
|--------|-------------|
| AppId | App identifier (joins to 14_Apps) |
| AppName | App display name |
| WorkspaceId | Source workspace |
| Identifier | Principal identifier (user/group/SPN) |
| DisplayName | Principal display name |
| EmailAddress | Email address (if applicable) |
| GraphId | Entra ID object identifier |
| PrincipalType | User, Group, or App |
| AppUserAccessRight | Access level granted |

**Why this matters for migration:** App access does not follow workspace membership. When migrating workspaces, app access must be re-applied manually. This table provides the data needed to reconstruct app permissions after migration.

**Rate limit:** 200 requests/hour. With ~1,400 apps in a large tenant, this step takes approximately 7 hours. A checkpoint file (`15_AppAccess_checkpoint.json`) is saved every 50 apps. On resume, the script reads processed AppIds from the checkpoint and skips them, then merges new results with the checkpoint before writing the final CSV.

**API limitation:** The API returns a flat principal list only — it does not expose app audience membership (which content is grouped per audience, and which users belong to which audience). That data is not available via any public admin API as of 2026-03.

**Standalone execution:** Step 15 can be run independently (`-Step 15`). It loads the app list from the latest `14_Apps_*.csv` in the output directory if Step 14 wasn't run in the same session.

### 16_DashboardTiles

Per-tile detail for all dashboards. One row per tile, retrieved via `GetTilesAsAdmin` — a per-dashboard endpoint that requires one API call per dashboard.

| Column | Description |
|--------|-------------|
| DashboardId | Parent dashboard identifier (joins to 13_Dashboards) |
| DashboardName | Dashboard display name |
| WorkspaceId | Parent workspace |
| WorkspaceName | Workspace display name |
| TileId | Unique tile identifier |
| TileTitle | Tile display title |
| TileType | Tile subtype: VisualReport, WebContent, Card, KPI, Gauge, Qna, etc. |
| ReportId | Linked report ID (null for WebContent tiles) |
| DatasetId | Linked dataset ID (null for WebContent tiles) |
| RowSpan | Tile row span on the dashboard grid |
| ColSpan | Tile column span on the dashboard grid |
| EmbedUrl | Tile embed URL |

**Key use cases:**

- **WebContent tiles (security risk):** Tiles with `TileType = "WebContent"` contain arbitrary HTML/iframe content. They have null `ReportId` and `DatasetId` — they are not connected to Power BI data. Relevant when `WebContentTilesTenant` is enabled tenant-wide.
- **Alertable tiles (migration risk):** Tiles with `TileType` in `Card`, `KPI`, `Gauge` can have user-configured data alerts. These alerts are silently lost during workspace migration — there is no API to export or re-create them.
- **Tile-to-report/dataset lineage:** The `ReportId` and `DatasetId` columns enable answering: "Which reports and datasets are actually pinned to dashboards?"

**Rate limit:** 200 requests/hour. With ~5,100 dashboards in a large tenant, this step takes approximately 26 hours. A checkpoint file (`16_DashboardTiles_checkpoint.json`) is saved every 50 dashboards. On resume, the script reads processed DashboardIds from the checkpoint and skips them.

**Standalone execution:** Step 16 can be run independently (`-Step 16`). It loads the dashboard list from the latest `13_Dashboards_*.csv` in the output directory if Step 13 wasn't run in the same session.

**Console output:** After completion, the script prints a tile type breakdown including WebContent tile count (highlighted in red if > 0), alertable tile count, and dashboard coverage statistics.

---

## Checkpoint / Resume Pattern

Steps 15 and 16 both implement the same checkpoint/resume pattern:

1. **On start:** Check for existing checkpoint JSON file in the output directory
2. **If found:** Load previously processed IDs and accumulated rows, skip already-processed items
3. **Every N items** (default: 50): Save checkpoint with all processed IDs, accumulated rows, and timestamp
4. **On completion:** Write final CSV, then delete the checkpoint file
5. **On interruption:** Checkpoint file remains — next run resumes automatically

Checkpoint files are JSON containing processed item IDs, accumulated row data, save timestamp, and total item count. They are self-contained: all data needed to resume is in the file, including previously collected rows.

**Important:** Do not delete checkpoint files manually while a run is in progress. Do not move or rename CSV files in the output directory between runs of the same step — the script uses the latest matching CSV to resolve parent data (apps for step 15, dashboards for step 16).

---

## Parameters Reference

| Parameter | Default | Description |
|-----------|---------|-------------|
| OutputPath | Current directory | Where CSV files are saved |
| Step | -1 (show menu) | Which step to run |
| BatchSize | 100 | Workspaces per Scanner API batch (max 100) |
| RefreshablePageSize | 500 | Refreshables per API page |
| AppAudienceCheckpointSize | 50 | Apps processed before saving checkpoint |
| DashboardTileCheckpointSize | 50 | Dashboards processed before saving checkpoint |
| MaxRetries | 10 | Retry attempts for failed API calls |
| MaxWaitSeconds | 60 | Max wait between retries |
| PollIntervalSeconds | 5 | Scan status poll interval |
| PollTimeoutSeconds | 1200 | Max wait for scans to complete (20 min) |
| TokenRefreshMinutes | 45 | Proactive token refresh interval |

---

## API Strategy and Known Gaps

### Why Multiple APIs?

Microsoft's Power BI admin APIs are maintained by different product teams, and each has different coverage and limitations:

| API | Strength | Limitation |
|-----|----------|------------|
| GetGroupsAsAdmin | Returns ALL workspaces including A-SKU | No artifact detail |
| Scanner API (GetModifiedWorkspaces) | Fast bulk metadata retrieval | Misses some workspace types |
| Scanner API (datasourceInstances) | Full datasource details at root level | Not documented clearly; requires reading from scan result root, not per-workspace |
| Scanner API (upstreamDataflows) | Dataset→dataflow AND dataflow→dataflow lineage | Exists on both dataset and dataflow objects, but only dataset usage is documented |
| Scanner API (dashboards) | Dashboard inventory with tile details | Tiles reference reports/datasets by ID but no direct model relationship |
| GetAppsAsAdmin | Bulk app inventory with workspace link | Does not include app users or permissions |
| GetAppUsersAsAdmin | Per-app access/permissions | 200 req/hour rate limit; no audience membership data |
| GetTilesAsAdmin | Per-dashboard tile detail | 200 req/hour rate limit; tile subType not documented |
| Gateways API | On-prem gateway details | Does not list cloud gateways |
| Refreshables API | Refresh health data | Dedicated capacity only |

### What's NOT Included (and Why)

| Artifact | API Required | Why Not Included |
|----------|-------------|-----------------|
| Subscriptions | `GetReportSubscriptionsAsAdmin` (per report) | 200 requests/hour rate limit × 75,000 reports = 375 hours. Not feasible. |
| Metrics / Scorecards | Goals REST API (per workspace) | No admin-level bulk endpoint. Requires per-group access. |
| App Audience Membership | No public API | API returns flat principal list only. Audience grouping not exposed. Idea logged with Microsoft. |
| Data Alerts on Tiles | No admin export API | User-configured alerts on Card/KPI/Gauge tiles cannot be exported or re-created programmatically. Silent data loss on migration. |

### The datasourceInstances Discovery

The Scanner API's `PostWorkspaceInfo` with `datasourceDetails=True` populates a `datasourceInstances` array at the **root level** of the scan result — not nested under each workspace. This was not obvious from the documentation and led to weeks of building a per-dataset API workaround (the former Step 12 / table 07_3) that took 80+ hours to run across the full tenant.

The fix was a single change in the PowerShell script: reading `$scanResult.datasourceInstances` instead of looking for it per-workspace. The `datasourceUsages` arrays on datasets and dataflows reference these root-level instances by `datasourceInstanceId`.

**Lesson learned:** Always dump the raw API response and inspect it before assuming what's there. The documentation describes the schema incompletely. A 5-minute JSON dump would have prevented weeks of workaround development.

### The upstreamDataflows Discovery

The Scanner API returns `upstreamDataflows` not only on **dataset** objects (documented — captured in table 10) but also on **dataflow** objects (not clearly documented). This enables full lineage chain reconstruction: Dataset → Dataflow → Dataflow → Datasource. Table 12 captures these dataflow-to-dataflow links.

The pattern is consistent: dataflows that consume external datasources (e.g. _extract) have `datasourceUsages`. Dataflows that consume other dataflows (e.g. _transform consuming _extract) have `upstreamDataflows`. A dataflow has one or the other, never both.

### Cloud Gateways

Power BI auto-creates internal "cloud gateways" for cloud datasources like SharePoint Online, Web sources, and similar. These gateways are not visible through the Gateways API (`GET /myorg/gateways`), which only returns on-premises and VNet gateways. The Scanner API assigns a GatewayId to all datasources including cloud ones — the presence of a GatewayId does not mean the datasource uses an on-premises gateway.

---

## Companion Script: PowerBITenantLens_Diagnostic_SingleWorkspace.ps1

A diagnostic script that scans a single workspace and dumps the complete raw Scanner API JSON response for inspection. Designed to investigate API response structure and discover undocumented properties.

```powershell
# Scan the default workspace
pwsh ./PowerBITenantLens_Diagnostic_SingleWorkspace.ps1

# Scan a specific workspace
pwsh ./PowerBITenantLens_Diagnostic_SingleWorkspace.ps1 -WorkspaceId "your-workspace-id-here"
```

Outputs:
- `ScanResult_FULL.json` — complete GetScanResult response
- `ScanResult_ROOT_KEYS.json` — top-level property names and types
- `ScanResult_datasourceInstances.json` — root-level datasource instances
- `ScanResult_dataflows.json` — all dataflow objects with property inspection
- `ScanResult_dataflow_PROPERTIES.txt` — property names per dataflow (highlights datasource-related)
- `ScanResult_datasets.json` — all dataset objects
- `ScanResult_12_DataflowUpstreamDataflows.csv` — dataflow-to-dataflow lineage

Also prints a full lineage chain reconstruction to the console: Dataset → Dataflow → Dataflow → Datasource.

**Best practice:** Run this script first when integrating any new API to discover the actual shape of the data before writing extraction code.

---

## Companion Script: PowerBITenantLens_Auth_Test.ps1

A lightweight script that verifies your environment is ready to run the inventory. It checks execution policy, installs or updates the `MicrosoftPowerBIMgmt` module, authenticates to the Power BI Service (with MFA via browser), confirms admin access by retrieving a sample of workspaces, and disconnects cleanly.

```powershell
pwsh ./PowerBITenantLens_Auth_Test.ps1
```

No parameters, no elevated rights required. Run this once on a new machine — especially on macOS after `brew install powershell` — to confirm that authentication and admin access work before starting a full inventory run.

---

## Building the Power BI Semantic Model

The CSV files are designed to form a star schema:

- **02_Workspaces** is the central dimension, joined to most other tables via WorkspaceId
- **05_SemanticModels** links to 04_Reports (via DatasetId), 06_DataSources, and 11_Refreshables
- **06_DataSources** links to 07_1/07_2 via GatewayId for gateway classification. DatasourceLookup (Power Query) joins 06 with 07_2, 08, and 10 to classify each datasource by category (On-premises data gateway, Cloud, Dataflow), type, and friendly name
- **08_Dataflows** links to 09_DataflowDatasources and 10_DatasetUpstreamDataflows
- **12_DataflowUpstreamDataflows** completes the lineage chain: Dataset → Dataflow → Dataflow → Datasource
- **13_Dashboards** links to 02_Workspaces via WorkspaceId, and to 16_DashboardTiles via DashboardId
- **14_Apps** links to 02_Workspaces via WorkspaceId, and to 15_AppAccess via AppId
- **15_AppAccess** provides per-app access details, answering: "Who has access to which apps, and with what permissions?"
- **16_DashboardTiles** provides per-tile detail, answering: "What types of tiles exist on each dashboard, and which reports/datasets do they reference?"

Use Power Query parameters for the file path and a name to identify the tenant.

See **Power BI Tenant Lens - the Power BI solution.md** for the complete semantic model documentation including Power Query M code, DAX calculated columns, measures, and the DatasourceLookup table.

---

## Version History

| Date       | Change                 |
| ---------- | ---------------------- |
| 2026-03-10 | Initial public release |
|            |                        |

---

## Troubleshooting

**429 Throttled errors:** The APIs have strict hourly rate limits. The script includes retry logic with exponential backoff. If you see persistent 429s, wait for the hourly window to reset before restarting.

**401 Unauthorized:** The script auto-refreshes tokens every 45 minutes. If you see repeated 401s, verify your admin role assignment.

**Empty datasourceInstances on workspace objects:** The `datasourceInstances` array is at the **root level** of the scan result, not per-workspace. Read it from `$scanResult.datasourceInstances`. If still empty, run the diagnostic script to inspect the raw JSON.

**Workspaces missing from Scanner API:** Check the InScannerAPI flag in table 02. Workspaces not in the Scanner API are typically A-SKU or SPN personal workspaces.

**Missing dataflow-to-dataflow lineage:** Check that the script extracts `upstreamDataflows` from dataflow objects (table 12). Run the diagnostic script against a workspace with chained dataflows to verify.

**Step 15 "No 14_Apps CSV found":** Run Step 14 first (or Step 0 / Step 43) to generate the apps list. Step 15 needs this as input to know which apps to query.

**Step 16 "No 13_Dashboards CSV found":** Run Step 13 first (or Step 0 / Step 43) to generate the dashboards list. Step 16 needs this as input to know which dashboards to query.

**Checkpoint resume shows "0 remaining":** All items were processed in a previous run. The script writes the final CSV from the checkpoint and deletes the checkpoint file. This is normal and means a previous run completed successfully but was interrupted before cleanup.

**MSAL "Request retry failed" warnings on macOS:** These are token cache persistence warnings — the authentication itself succeeds. You can safely ignore these warnings. Verify by checking that `UserName` and `TenantId` appear in the output above the warnings.

**Empty 13_Dashboards:** Some tenants have migrated entirely from dashboards to reports. If the CSV is empty, dashboards simply don't exist in the scanned workspaces.

**Empty 14_Apps:** Verify the admin role. `GetAppsAsAdmin` requires Power BI Service Administrator or Fabric Administrator. If the API returns empty results, check that apps have been published in the tenant.
