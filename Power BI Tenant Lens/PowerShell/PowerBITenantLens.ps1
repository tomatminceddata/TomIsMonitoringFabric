<#
.SYNOPSIS
    PowerBITenantLens.ps1 — Power BI Tenant Lens
    (Wide Scan + Scanner + Refreshables + Dashboards + Apps + Access + Tiles)

.DESCRIPTION
    Extends previous versions with Step 15: App Access and Step 16: Dashboard Tiles.
    
    Step 15 calls GET /admin/apps/{appId}/users for every app collected in
    Step 14, exporting who has access to each app (users, groups, service
    principals) via GetAppUsersAsAdmin. This is migration-critical data:
    app access does not follow workspace membership and must be re-applied
    manually after migration.

    Note: the API returns a flat principal list only — it does not expose
    app audience membership (content grouping within an app). That data is
    not available via any public admin API as of 2026-03.
    
    Rate limit: 200 requests/hour. With ~1,400 apps in a large tenant
    this step takes ~7.2 hours. A checkpoint file is saved every 50 apps so
    the step can be resumed safely after any interruption without losing
    previously collected data.
    
    Checkpoint file: <OutputPath>\15_AppAccess_checkpoint.json
    On resume the script reads processed AppIds from the checkpoint and skips
    them, then merges the new results with the checkpoint before writing the
    final CSV.
    
    All other steps are identical to WCSRDA. See WCSRDA for full API strategy
    documentation.
    
    Exports CSV files:
    01_TenantSettings, 02_Workspaces, 03_WorkspaceMembers, 04_Reports,
    05_SemanticModels, 06_DataSources, 07_1_GatewayClusters,
    07_2_GatewayDatasources, 08_Dataflows, 09_DataflowDatasources,
    10_DatasetUpstreamDataflows, 11_Refreshables,
    12_DataflowUpstreamDataflows, 13_Dashboards, 14_Apps,
    15_AppAccess, 16_DashboardTiles

    Step 16 calls GET /admin/dashboards/{dashboardId}/tiles for every
    dashboard collected in Step 13 (via Scanner API). One row per tile.
    Key use cases:
    - Identify WebContent tiles (type = WebContent, null reportId/datasetId)
      to audit 3rd-party HTML/iframe usage. Relevant because
      WebContentTilesTenant is enabled tenant-wide.
    - Identify tiles eligible for data alerts (type = Card, KPI, Gauge)
      which will lose user-configured alerts silently after migration.

    Rate limit: 200 req/hour. With ~5,100 dashboards in a large tenant this step takes
    ~26 hours. Checkpoint file saved every 50 dashboards.

    Checkpoint file: <OutputPath>\16_DashboardTiles_checkpoint.json

    Requires:
    - PowerShell 7 or higher
    - MicrosoftPowerBIMgmt PowerShell module (for authentication)
    - Power BI Service Administrator or Fabric Administrator role

.PARAMETER OutputPath
    Directory where CSV files will be saved. Defaults to current directory.

.PARAMETER Step
    Which step(s) to run (use menu if not specified).

.PARAMETER BatchSize
    Number of workspaces per Scanner API batch. Default: 100 (max allowed).

.PARAMETER RefreshablePageSize
    Number of refreshables per API page. Default: 500.

.PARAMETER AppAudienceCheckpointSize
    Number of apps processed before saving a checkpoint. Default: 50.

.PARAMETER DashboardTileCheckpointSize
    Number of dashboards processed before saving a checkpoint. Default: 50.

.PARAMETER MaxRetries
    Maximum retry attempts for failed API calls. Default: 10.

.PARAMETER MaxWaitSeconds
    Maximum wait between retries (caps exponential backoff). Default: 60.

.PARAMETER PollIntervalSeconds
    Seconds between scan status polls. Default: 5.

.PARAMETER PollTimeoutSeconds
    Maximum seconds to wait for scans to complete. Default: 1200 (20 minutes).

.PARAMETER TokenRefreshMinutes
    Refresh access token after this many minutes. Default: 45.

.EXAMPLE
    .\PowerBITenantLens.ps1                  # Shows menu
    .\PowerBITenantLens.ps1 -Step 0          # Run all steps
    .\PowerBITenantLens.ps1 -Step 15         # App Access only
    .\PowerBITenantLens.ps1 -Step 16         # Dashboard Tiles only (~26h, checkpoint every 50)
    .\PowerBITenantLens.ps1 -Step 43         # All steps EXCEPT long-running (skips 15 + 16)
    .\PowerBITenantLens.ps1 -Step 42         # First 1000 workspaces (test)

.NOTES
    Author: Tom Martens & Claude
    Date:   2026-03-05
    
    Wide Scan + Scanner + Refreshables + Dashboards + Apps + Access + Tiles Edition
    
    Change history:
    - 2026-03-06: WCSRDAA: Added Step 43 (all steps except 15 + 16, for quick re-runs).
    - 2026-03-05: WCSRDAA: Added Step 16 (Dashboard Tiles) via per-dashboard
                  GetTilesAsAdmin endpoint. One row per tile. Captures tile
                  type (WebContent, VisualReport, Qna etc), reportId, datasetId.
                  Checkpoint/resume every 50 dashboards. ~26h for large tenants.
    - 2026-03-05: WCSRDAA: Added Step 15 (App Access) via per-app
                  GetAppUsersAsAdmin endpoint. Checkpoint/resume every 50 apps.
                  Renamed from WCSRDA to WCSRDAA (2nd A = Access).
    - 2026-03-04: WCSRDA: Added 12_DataflowUpstreamDataflows, 13_Dashboards,
                  14_Apps. Renamed from WCSR to WCSRDA.
    - 2026-03-04: BREAKING FIX: datasourceInstances retrieved from scan result
                  ROOT level, not per-workspace (confirmed by Microsoft/Meir Shimon).
                  Tables 06 and 09 now fully populated. Removed Step 12
                  (per-dataset API) and 07_3_CloudDatasources.
    - 2026-02-23: Added Cloud Datasources (07_3) via per-dataset API.
    - 2026-02-22: Replaced GetModifiedWorkspaces with GetGroupsAsAdmin.
    - 2026-02-22: Added TokenExpired detection + proactive token refresh.
    - 2026-02-20: Added Refreshables API.
    - 2026-02-19: Initial Scanner API version.
#>

param(
    [string]$OutputPath = (Get-Location).Path,
    [int]$Step = -1,
    [int]$BatchSize = 100,
    [int]$RefreshablePageSize = 500,
    [int]$AppAudienceCheckpointSize = 50,
    [int]$DashboardTileCheckpointSize = 50,
    [int]$MaxRetries = 10,
    [int]$MaxWaitSeconds = 60,
    [int]$PollIntervalSeconds = 5,
    [int]$PollTimeoutSeconds = 1200,
    [int]$TokenRefreshMinutes = 45
)

#region Setup & Version Check
$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7 or higher." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "Please run: pwsh .\PowerBITenantLens.ps1" -ForegroundColor Yellow
    exit 1
}

if ($BatchSize -gt 100) { 
    Write-Host "BatchSize capped at 100 (Scanner API maximum)" -ForegroundColor Yellow
    $BatchSize = 100 
}
if ($BatchSize -lt 1) { $BatchSize = 100 }

if ($RefreshablePageSize -lt 1) { $RefreshablePageSize = 500 }
if ($RefreshablePageSize -gt 5000) { 
    Write-Host "RefreshablePageSize capped at 5000 to be safe" -ForegroundColor Yellow
    $RefreshablePageSize = 5000 
}

if ($AppAudienceCheckpointSize -lt 1) { $AppAudienceCheckpointSize = 50 }

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Power BI Tenant Lens" -ForegroundColor Cyan
Write-Host "  Tenant Inventory Script" -ForegroundColor Yellow
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
#endregion

#region Menu
if ($Step -eq -1) {
    Write-Host "Select which step(s) to run:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [0] Run ALL steps (full inventory)" -ForegroundColor White
    Write-Host "  [1] Tenant Settings" -ForegroundColor Gray
    Write-Host "  [2] Workspaces (via GetGroupsAsAdmin + Scanner API)" -ForegroundColor Gray
    Write-Host "  [3] Workspace Members (via Scanner API)" -ForegroundColor Gray
    Write-Host "  [4] Reports (via Scanner API)" -ForegroundColor Gray
    Write-Host "  [5] Semantic Models (via Scanner API)" -ForegroundColor Gray
    Write-Host "  [6] Data Sources (via Scanner API)" -ForegroundColor Gray
    Write-Host "  [7] Gateways (07_1: Clusters, 07_2: Datasources)" -ForegroundColor Gray
    Write-Host "  [11] Refreshables (via Admin Refreshables API)" -ForegroundColor Gray
    Write-Host "  [13] Dashboards (via Scanner API)" -ForegroundColor Gray
    Write-Host "  [14] Apps (via GetAppsAsAdmin API)" -ForegroundColor Gray
    Write-Host "  [15] App Access (via GetAppUsersAsAdmin — ~7h, checkpoint every $AppAudienceCheckpointSize apps)" -ForegroundColor Cyan
    Write-Host "  [16] Dashboard Tiles (via GetTilesAsAdmin — ~26h, checkpoint every $DashboardTileCheckpointSize dashboards)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [43] Fast run — all steps EXCEPT 15 + 16 (skips ~33h of rate-limited calls)" -ForegroundColor Green
    Write-Host "  [42] First 1000 workspaces (steps 2-6, quick test)" -ForegroundColor Green
    Write-Host ""
    $selection = Read-Host "Enter selection (0-7, 11, 13, 14, 15, 16, 42, 43)"
    $Step = [int]$selection
}

$First1000    = ($Step -eq 42)
$FastRun      = ($Step -eq 43)
if ($FastRun) {
    Write-Host ""
    Write-Host "  Running in FAST mode — all steps except Step 15 (App Access) and Step 16 (Dashboard Tiles)..." -ForegroundColor Green
}
if ($First1000) {
    Write-Host ""
    Write-Host "  Running in FIRST 1000 mode for quick testing..." -ForegroundColor Green
}

$needsScan          = ($Step -eq 0) -or ($Step -in 2..6) -or ($Step -eq 13) -or $First1000
$exportWorkspaces   = ($Step -eq 0) -or ($Step -eq 2) -or $First1000
$exportMembers      = ($Step -eq 0) -or ($Step -eq 3) -or $First1000
$exportReports      = ($Step -eq 0) -or ($Step -eq 4) -or $First1000
$exportDatasets     = ($Step -eq 0) -or ($Step -eq 5) -or $First1000
$exportDatasources  = ($Step -eq 0) -or ($Step -eq 6) -or $First1000
$exportGateways     = ($Step -eq 0) -or ($Step -eq 7)
$exportRefreshables = ($Step -eq 0) -or ($Step -eq 11)
$exportDashboards   = ($Step -eq 0) -or ($Step -eq 13) -or $First1000
$exportApps         = ($Step -eq 0) -or ($Step -eq 14)
$exportAppAudiences   = ($Step -eq 0) -or ($Step -eq 15)
$exportDashboardTiles = ($Step -eq 0) -or ($Step -eq 16)

# Step 43: Fast run = all steps except the long-running 15 + 16
if ($FastRun) {
    $needsScan          = $true
    $exportWorkspaces   = $true
    $exportMembers      = $true
    $exportReports      = $true
    $exportDatasets     = $true
    $exportDatasources  = $true
    $exportGateways     = $true
    $exportRefreshables = $true
    $exportDashboards   = $true
    $exportApps         = $true
    $exportAppAudiences   = $false
    $exportDashboardTiles = $false
}

Write-Host ""
if ($needsScan) {
    Write-Host "  Scanner batch size:      $BatchSize workspaces per scan request" -ForegroundColor Gray
}
if ($exportRefreshables) {
    Write-Host "  Refreshable page size:   $RefreshablePageSize per request" -ForegroundColor Gray
}
if ($exportAppAudiences) {
    Write-Host "  App Access checkpoint:     every $AppAudienceCheckpointSize apps" -ForegroundColor Gray
}
if ($exportDashboardTiles) {
    Write-Host "  Dashboard Tile checkpoint: every $DashboardTileCheckpointSize dashboards" -ForegroundColor Gray
}
Write-Host ""
#endregion

#region Module Check
Write-Host "Checking PowerShell module..." -ForegroundColor Gray

$moduleName = "MicrosoftPowerBIMgmt"
$installedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1

if (-not $installedModule) {
    Write-Host "Installing $moduleName module..." -ForegroundColor Yellow
    Install-Module -Name $moduleName -Scope CurrentUser -Force
}
else {
    Write-Host "$moduleName v$($installedModule.Version) found" -ForegroundColor Gray
}

Import-Module $moduleName
Write-Host ""
#endregion

#region Authentication
Write-Host "Connecting to Power BI Service..." -ForegroundColor Yellow
try {
    Connect-PowerBIServiceAccount | Out-Null
    Write-Host "Connected successfully!" -ForegroundColor Green
    
    $script:AccessToken = (Get-PowerBIAccessToken -AsString).Replace("Bearer ", "")
    $script:TokenAcquiredAt = Get-Date
    $script:TokenRefreshMinutes = $TokenRefreshMinutes
    $script:MaxRetries = $MaxRetries
    $script:MaxWaitSeconds = $MaxWaitSeconds
}
catch {
    Write-Host "Failed to connect to Power BI Service: $_" -ForegroundColor Red
    exit 1
}
Write-Host ""
#endregion

#region Helper: Token refresh
function Ensure-FreshToken {
    $tokenAgeMinutes = ((Get-Date) - $script:TokenAcquiredAt).TotalMinutes
    
    if ($tokenAgeMinutes -ge $script:TokenRefreshMinutes) {
        Write-Host "  Token age: $([math]::Round($tokenAgeMinutes, 1)) min - refreshing..." -ForegroundColor Yellow
        try {
            $script:AccessToken = (Get-PowerBIAccessToken -AsString).Replace("Bearer ", "")
            $script:TokenAcquiredAt = Get-Date
            Write-Host "  Token refreshed successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "  Token refresh failed - reconnecting..." -ForegroundColor Yellow
            try {
                Connect-PowerBIServiceAccount | Out-Null
                $script:AccessToken = (Get-PowerBIAccessToken -AsString).Replace("Bearer ", "")
                $script:TokenAcquiredAt = Get-Date
                Write-Host "  Reconnected and token refreshed" -ForegroundColor Green
            }
            catch {
                Write-Host "  FATAL: Could not refresh token: $_" -ForegroundColor Red
                throw
            }
        }
    }
}
#endregion

#region Helper: REST API call with retry
function Invoke-PBIRestMethod {
    param(
        [string]$Url,
        [string]$Method = "Get",
        [string]$Body = $null,
        [int]$MaxRetries = $script:MaxRetries
    )
    
    Ensure-FreshToken
    
    $headers = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Content-Type"  = "application/json"
    }
    
    $retryCount = 0
    while ($true) {
        try {
            $params = @{
                Uri        = $Url
                Headers    = $headers
                Method     = $Method
                TimeoutSec = 120
            }
            if ($Body) { $params.Body = $Body }
            
            $response = Invoke-RestMethod @params
            return $response
        }
        catch {
            $statusCode = $null
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch {}
            $errorMessage = $_.Exception.Message
            
            $isTransient = $false
            $retryReason = ""
            
            if ($statusCode -eq 429) {
                $isTransient = $true
                $retryReason = "429 throttled"
            }
            elseif ($statusCode -eq 401) {
                if ($retryCount -lt $MaxRetries) {
                    $retryCount++
                    Write-Host "    401 unauthorized - forcing token refresh (retry $retryCount/$MaxRetries)..." -ForegroundColor DarkYellow
                    $script:TokenAcquiredAt = [datetime]::MinValue
                    Ensure-FreshToken
                    $headers["Authorization"] = "Bearer $script:AccessToken"
                    continue
                }
            }
            elseif ($statusCode -eq 502 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                $isTransient = $true
                $retryReason = "HTTP $statusCode server error"
            }
            elseif ($statusCode -eq $null) {
                $isTransient = $true
                $retryReason = "Network error"
            }
            
            if ($errorMessage -match 'TokenExpired' -or $errorMessage -match 'token has expired') {
                if ($retryCount -lt $MaxRetries) {
                    $retryCount++
                    Write-Host "    TokenExpired in response body - forcing token refresh (retry $retryCount/$MaxRetries)..." -ForegroundColor DarkYellow
                    $script:TokenAcquiredAt = [datetime]::MinValue
                    Ensure-FreshToken
                    $headers["Authorization"] = "Bearer $script:AccessToken"
                    continue
                }
            }
            
            if ($isTransient -and $retryCount -lt $MaxRetries) {
                $retryCount++
                
                $retryAfter = 0
                if ($statusCode -eq 429) {
                    try {
                        $raHeader = $_.Exception.Response.Headers["Retry-After"]
                        if ($raHeader) {
                            $parsed = 0
                            if ([int]::TryParse($raHeader, [ref]$parsed) -and $parsed -gt 0) {
                                $retryAfter = $parsed
                            }
                        }
                    } catch {}
                }
                
                if ($retryAfter -gt 0) {
                    $waitSeconds = $retryAfter
                }
                else {
                    $waitSeconds = [math]::Min(
                        [math]::Pow(2, $retryCount) + (Get-Random -Minimum 1 -Maximum 5),
                        $script:MaxWaitSeconds
                    )
                }
                
                Write-Host "    $retryReason - waiting ${waitSeconds}s (retry $retryCount/$MaxRetries) [$errorMessage]" -ForegroundColor DarkYellow
                Start-Sleep -Seconds $waitSeconds
            }
            else {
                throw
            }
        }
    }
}
#endregion

#region Step 1: Tenant Settings
if ($Step -eq 0 -or $Step -eq 1 -or $First1000) {
    Write-Host "[1] Exporting Admin Portal Tenant Settings..." -ForegroundColor Yellow
    
    try {
        $tenantSettingsResponse = Invoke-PBIRestMethod -Url "https://api.fabric.microsoft.com/v1/admin/tenantsettings"
        
        $tenantSettings = @()
        
        foreach ($setting in $tenantSettingsResponse.tenantSettings) {
            $settingObj = [PSCustomObject]@{
                SettingName              = $setting.settingName
                Title                    = $setting.title
                Enabled                  = $setting.enabled
                CanSpecifySecurityGroups = $setting.canSpecifySecurityGroups
                EnabledSecurityGroups    = ($setting.enabledSecurityGroups.graphId -join "; ")
                TenantSettingGroup       = $setting.tenantSettingGroup
                Properties               = ($setting.properties | ConvertTo-Json -Compress)
            }
            $tenantSettings += $settingObj
        }
        
        $tenantSettingsFile = Join-Path $OutputPath "01_TenantSettings_$timestamp.csv"
        $tenantSettings | Export-Csv -Path $tenantSettingsFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  Exported $($tenantSettings.Count) tenant settings" -ForegroundColor Green
        Write-Host "  File: $tenantSettingsFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "  Warning: Could not export tenant settings." -ForegroundColor Yellow
        Write-Host "  This requires Fabric Administrator role." -ForegroundColor Yellow
        Write-Host "  Error: $_" -ForegroundColor Gray
        $tenantSettings = @()
    }
    Write-Host ""
}
#endregion

#region Steps 2-6, 8-10, 12-13: GetGroupsAsAdmin + Scanner API
if ($needsScan) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "  Wide Scan: GetGroupsAsAdmin + Scanner API" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    # PHASE 1: GetGroupsAsAdmin
    Write-Host "[Scanner] Phase 1: Enumerating workspaces via GetGroupsAsAdmin..." -ForegroundColor Yellow
    $scanStartTime = Get-Date
    
    $allGroupsData = [System.Collections.Generic.List[object]]::new()
    $groupsBatchSize = 5000
    $groupsSkip = 0
    $groupsBatchNum = 0
    
    do {
        $groupsBatchNum++
        $groupsUrl = "https://api.powerbi.com/v1.0/myorg/admin/groups?`$top=$groupsBatchSize&`$skip=$groupsSkip"
        
        Write-Host "  GetGroupsAsAdmin batch $groupsBatchNum - rows $groupsSkip to $($groupsSkip + $groupsBatchSize - 1)..." -ForegroundColor Gray
        
        try {
            $groupsResponse = Invoke-PBIRestMethod -Url $groupsUrl
            $groupsBatch = $groupsResponse.value
            
            if ($null -eq $groupsBatch -or $groupsBatch.Count -eq 0) {
                Write-Host "  Batch $groupsBatchNum - no more results." -ForegroundColor Gray
                break
            }
            
            $allGroupsData.AddRange($groupsBatch)
            Write-Host "  Batch $groupsBatchNum - $($groupsBatch.Count) workspaces (total: $($allGroupsData.Count))" -ForegroundColor Gray
            
            if ($groupsBatch.Count -lt $groupsBatchSize) { break }
            
            $groupsSkip += $groupsBatchSize
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Host "  Batch $groupsBatchNum FAILED: $_" -ForegroundColor Red
            break
        }
    } while ($true)
    
    Write-Host "  Found $($allGroupsData.Count) workspaces via GetGroupsAsAdmin ($groupsBatchNum batches)" -ForegroundColor Green
    
    if ($First1000) {
        $first1000Items = @($allGroupsData | Select-Object -First 1000)
        $allGroupsData = [System.Collections.Generic.List[object]]::new()
        $allGroupsData.AddRange($first1000Items)
        Write-Host "  Limited to first 1000 workspaces (test mode)" -ForegroundColor Yellow
    }
    
    $workspaceIds = $allGroupsData | ForEach-Object { [PSCustomObject]@{ id = $_.id } }
    Write-Host "  Feeding $($workspaceIds.Count) workspace IDs into Scanner API..." -ForegroundColor Green
    Write-Host ""
    
    # PHASE 2: Submit scan batches
    Write-Host "[Scanner] Phase 2: Submitting scan batches (batch size: $BatchSize)..." -ForegroundColor Yellow
    
    $scanIds = @()
    $totalBatches = [math]::Ceiling($workspaceIds.Count / $BatchSize)
    
    for ($i = 0; $i -lt $workspaceIds.Count; $i += $BatchSize) {
        $batchNum = [math]::Floor($i / $BatchSize) + 1
        $end = [math]::Min($i + $BatchSize - 1, $workspaceIds.Count - 1)
        $batch = $workspaceIds[$i..$end]
        
        $wsIdList = ($batch | ForEach-Object { "`"$($_.id)`"" }) -join ","
        $body = "{ `"workspaces`": [ $wsIdList ] }"
        
        $scanUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?datasourceDetails=True&getArtifactUsers=True&lineage=True"
        
        try {
            $scanResponse = Invoke-PBIRestMethod -Url $scanUrl -Method "Post" -Body $body
            $scanIds += $scanResponse.id
            Write-Host "  Batch $batchNum/$totalBatches submitted (scan ID: $($scanResponse.id))" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Batch $batchNum/$totalBatches FAILED after all retries: $_" -ForegroundColor Red
        }
        
        if ($batchNum -lt $totalBatches) { Start-Sleep -Milliseconds 500 }
    }
    
    Write-Host "  Submitted $($scanIds.Count)/$totalBatches scan batches" -ForegroundColor $(if ($scanIds.Count -lt $totalBatches) { "Yellow" } else { "Green" })
    Write-Host ""
    
    # PHASE 3: Poll for completion
    Write-Host "[Scanner] Phase 3: Waiting for scans to complete..." -ForegroundColor Yellow
    
    $completedScans = @{}
    $failedScans = @()
    $pollStart = Get-Date
    
    while ($completedScans.Count + $failedScans.Count -lt $scanIds.Count) {
        if (((Get-Date) - $pollStart).TotalSeconds -gt $PollTimeoutSeconds) {
            Write-Host "  TIMEOUT: Some scans did not complete within $PollTimeoutSeconds seconds" -ForegroundColor Red
            break
        }
        
        foreach ($scanId in $scanIds) {
            if ($completedScans.ContainsKey($scanId) -or $failedScans -contains $scanId) { continue }
            
            try {
                $statusResponse = Invoke-PBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$scanId"
                
                if ($statusResponse.status -eq "Succeeded") {
                    $completedScans[$scanId] = $true
                    Write-Host "  Scan $scanId completed ($($completedScans.Count)/$($scanIds.Count))" -ForegroundColor Green
                }
                elseif ($statusResponse.status -eq "Failed") {
                    $failedScans += $scanId
                    Write-Host "  Scan $scanId FAILED" -ForegroundColor Red
                }
            }
            catch {}
        }
        
        if ($completedScans.Count + $failedScans.Count -lt $scanIds.Count) {
            $remaining = $scanIds.Count - $completedScans.Count - $failedScans.Count
            Write-Host "  Waiting ${PollIntervalSeconds}s... ($remaining scans still running)" -ForegroundColor DarkGray
            Start-Sleep -Seconds $PollIntervalSeconds
        }
    }
    
    Write-Host "  Completed: $($completedScans.Count), Failed: $($failedScans.Count)" -ForegroundColor $(if ($failedScans.Count -gt 0) { "Yellow" } else { "Green" })
    Write-Host ""
    
    # PHASE 4: Retrieve results
    Write-Host "[Scanner] Phase 4: Retrieving scan results..." -ForegroundColor Yellow
    
    $allWorkspaceData = @()
    $allDatasourceInstances = @()
    $resultCount = 0
    
    foreach ($scanId in $completedScans.Keys) {
        $resultCount++
        try {
            $scanResult = Invoke-PBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$scanId"
            
            $allWorkspaceData += $scanResult.workspaces
            
            if ($scanResult.datasourceInstances) {
                $allDatasourceInstances += $scanResult.datasourceInstances
            }
            
            $dsInstanceCount = if ($scanResult.datasourceInstances) { $scanResult.datasourceInstances.Count } else { 0 }
            Write-Host "  Result $resultCount/$($completedScans.Count): $($scanResult.workspaces.Count) workspaces, $dsInstanceCount datasource instances" -ForegroundColor Gray
        }
        catch {
            Write-Host "  Result $resultCount/$($completedScans.Count) FAILED: $_" -ForegroundColor Red
        }
    }
    
    $scanElapsed = (Get-Date) - $scanStartTime
    
    $dsInstanceLookup = @{}
    foreach ($dsi in $allDatasourceInstances) {
        if ($dsi.datasourceId -and -not $dsInstanceLookup.ContainsKey($dsi.datasourceId)) {
            $dsInstanceLookup[$dsi.datasourceId] = $dsi
        }
    }
    
    Write-Host ""
    Write-Host "  Wide Scan complete:" -ForegroundColor Cyan
    Write-Host "    GetGroupsAsAdmin: $($allGroupsData.Count) workspaces enumerated" -ForegroundColor Cyan
    Write-Host "    Scanner API:      $($allWorkspaceData.Count) workspaces scanned" -ForegroundColor Cyan
    Write-Host "    DatasourceInstances: $($allDatasourceInstances.Count) collected ($($dsInstanceLookup.Count) unique)" -ForegroundColor Cyan
    Write-Host "    Time elapsed:     $($scanElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
    Write-Host ""
    
    # PHASE 5: Flatten and export
    Write-Host "[Scanner] Phase 5: Flattening and exporting data..." -ForegroundColor Yellow
    Write-Host ""
    
    # 02_Workspaces
    if ($exportWorkspaces) {
        Write-Host "  [2] Workspaces..." -ForegroundColor Yellow
        
        $scannerLookup = @{}
        foreach ($ws in $allWorkspaceData) { $scannerLookup[$ws.id] = $true }
        
        $workspaceSummary = foreach ($ws in $allGroupsData) {
            [PSCustomObject]@{
                WorkspaceId                 = $ws.id
                WorkspaceName               = $ws.name
                Type                        = $ws.type
                State                       = $ws.state
                IsReadOnly                  = $ws.isReadOnly
                IsOnDedicatedCapacity       = $ws.isOnDedicatedCapacity
                CapacityId                  = $ws.capacityId
                Description                 = $ws.description
                DefaultDatasetStorageFormat = $ws.defaultDatasetStorageFormat
                HasWorkspaceLevelSettings   = $ws.hasWorkspaceLevelSettings
                CapacityMigrationStatus     = $ws.capacityMigrationStatus
                DataflowStorageId           = $ws.dataflowStorageId
                PipelineId                  = $ws.pipelineId
                InScannerAPI                = [bool]$scannerLookup[$ws.id]
            }
        }
        
        $workspaceSummaryFile = Join-Path $OutputPath "02_Workspaces_$timestamp.csv"
        $workspaceSummary | Export-Csv -Path $workspaceSummaryFile -NoTypeInformation -Encoding UTF8
        
        $inScanner    = ($workspaceSummary | Where-Object { $_.InScannerAPI }).Count
        $notInScanner = $workspaceSummary.Count - $inScanner
        
        Write-Host "  Exported $($workspaceSummary.Count) workspaces" -ForegroundColor Green
        Write-Host "  In Scanner API: $inScanner, Not in Scanner API: $notInScanner" -ForegroundColor $(if ($notInScanner -gt 0) { "Yellow" } else { "Gray" })
        Write-Host "  File: $workspaceSummaryFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # 03_WorkspaceMembers
    if ($exportMembers) {
        Write-Host "  [3] Workspace Members..." -ForegroundColor Yellow
        
        $workspaceMembers = foreach ($ws in $allWorkspaceData) {
            if ($ws.users) {
                foreach ($user in $ws.users) {
                    [PSCustomObject]@{
                        WorkspaceId       = $ws.id
                        WorkspaceName     = $ws.name
                        WorkspaceType     = $ws.type
                        WorkspaceState    = $ws.state
                        UserIdentifier    = $user.identifier
                        UserPrincipalName = $user.userPrincipalName
                        DisplayName       = $user.displayName
                        AccessRight       = $user.groupUserAccessRight
                        PrincipalType     = $user.principalType
                        GraphId           = $user.graphId
                    }
                }
            }
        }
        
        $workspaceMembersFile = Join-Path $OutputPath "03_WorkspaceMembers_$timestamp.csv"
        $workspaceMembers | Export-Csv -Path $workspaceMembersFile -NoTypeInformation -Encoding UTF8
        
        $wsWithMembers    = ($allWorkspaceData | Where-Object { $_.users -and $_.users.Count -gt 0 }).Count
        $wsWithoutMembers = $allWorkspaceData.Count - $wsWithMembers
        
        Write-Host "  Exported $($workspaceMembers.Count) membership records" -ForegroundColor Green
        Write-Host "  Workspaces with members: $wsWithMembers, without: $wsWithoutMembers" -ForegroundColor Gray
        Write-Host "  File: $workspaceMembersFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # 04_Reports
    if ($exportReports) {
        Write-Host "  [4] Reports..." -ForegroundColor Yellow
        
        $allReports = foreach ($ws in $allWorkspaceData) {
            if ($ws.reports) {
                foreach ($report in $ws.reports) {
                    [PSCustomObject]@{
                        WorkspaceId      = $ws.id
                        WorkspaceName    = $ws.name
                        ReportId         = $report.id
                        ReportName       = $report.name
                        ReportType       = $report.reportType
                        DatasetId        = $report.datasetId
                        CreatedDateTime  = if ($report.createdDateTime)  { ([datetime]$report.createdDateTime).ToString("yyyy-MM-ddTHH:mm:ss")  } else { $null }
                        ModifiedDateTime = if ($report.modifiedDateTime) { ([datetime]$report.modifiedDateTime).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                        ModifiedBy       = $report.modifiedBy
                        CreatedBy        = $report.createdBy
                    }
                }
            }
        }
        
        $reportsFile = Join-Path $OutputPath "04_Reports_$timestamp.csv"
        $allReports | Export-Csv -Path $reportsFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  Exported $($allReports.Count) reports" -ForegroundColor Green
        Write-Host "  File: $reportsFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # 05_SemanticModels
    if ($exportDatasets) {
        Write-Host "  [5] Semantic Models..." -ForegroundColor Yellow
        
        $allDatasets = foreach ($ws in $allWorkspaceData) {
            if ($ws.datasets) {
                foreach ($dataset in $ws.datasets) {
                    [PSCustomObject]@{
                        WorkspaceId             = $ws.id
                        WorkspaceName           = $ws.name
                        DatasetId               = $dataset.id
                        DatasetName             = $dataset.name
                        ConfiguredBy            = $dataset.configuredBy
                        CreatedDate             = if ($dataset.createdDate) { ([datetime]$dataset.createdDate).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                        IsRefreshable           = $dataset.isRefreshable
                        IsOnPremGatewayRequired = $dataset.isOnPremGatewayRequired
                        TargetStorageMode       = $dataset.targetStorageMode
                        ContentProviderType     = $dataset.contentProviderType
                    }
                }
            }
        }
        
        $datasetsFile = Join-Path $OutputPath "05_SemanticModels_$timestamp.csv"
        $allDatasets | Export-Csv -Path $datasetsFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  Exported $($allDatasets.Count) semantic models" -ForegroundColor Green
        Write-Host "  File: $datasetsFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # 06_DataSources
    if ($exportDatasources) {
        Write-Host "  [6] Data Sources..." -ForegroundColor Yellow
        
        $allDataSources = foreach ($ws in $allWorkspaceData) {
            if ($ws.datasets) {
                foreach ($dataset in $ws.datasets) {
                    if ($dataset.datasourceUsages) {
                        foreach ($usage in $dataset.datasourceUsages) {
                            $dsInstance = $dsInstanceLookup[$usage.datasourceInstanceId]
                            [PSCustomObject]@{
                                WorkspaceId       = $ws.id
                                WorkspaceName     = $ws.name
                                DatasetId         = $dataset.id
                                DatasetName       = $dataset.name
                                DatasourceId      = $usage.datasourceInstanceId
                                DatasourceType    = $dsInstance.datasourceType
                                GatewayId         = $dsInstance.gatewayId
                                ConnectionDetails = ($dsInstance.connectionDetails | ConvertTo-Json -Compress)
                                Server            = $dsInstance.connectionDetails.server
                                Database          = $dsInstance.connectionDetails.database
                                Url               = $dsInstance.connectionDetails.url
                                Path              = $dsInstance.connectionDetails.path
                            }
                        }
                    }
                }
            }
        }
        
        $datasourcesFile = Join-Path $OutputPath "06_DataSources_$timestamp.csv"
        $allDataSources | Export-Csv -Path $datasourcesFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "  Exported $($allDataSources.Count) data source connections" -ForegroundColor Green
        Write-Host "  File: $datasourcesFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # 08_Dataflows
    Write-Host "  [8] Dataflows..." -ForegroundColor Yellow
    
    $allDataflows = foreach ($ws in $allWorkspaceData) {
        if ($ws.dataflows) {
            foreach ($dataflow in $ws.dataflows) {
                [PSCustomObject]@{
                    WorkspaceId      = $ws.id
                    WorkspaceName    = $ws.name
                    DataflowId       = $dataflow.objectId
                    DataflowName     = $dataflow.name
                    Description      = $dataflow.description
                    ConfiguredBy     = $dataflow.configuredBy
                    ModifiedBy       = $dataflow.modifiedBy
                    ModifiedDateTime = if ($dataflow.modifiedDateTime) { ([datetime]$dataflow.modifiedDateTime).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                    Endorsement      = $dataflow.endorsementDetails.endorsement
                    CertifiedBy      = $dataflow.endorsementDetails.certifiedBy
                    SensitivityLabel = $dataflow.sensitivityLabel.labelId
                }
            }
        }
    }
    
    $dataflowsFile = Join-Path $OutputPath "08_Dataflows_$timestamp.csv"
    $allDataflows | Export-Csv -Path $dataflowsFile -NoTypeInformation -Encoding UTF8
    
    $wsWithDataflows = ($allWorkspaceData | Where-Object { $_.dataflows -and $_.dataflows.Count -gt 0 }).Count
    Write-Host "  Exported $($allDataflows.Count) dataflows from $wsWithDataflows workspaces" -ForegroundColor Green
    Write-Host "  File: $dataflowsFile" -ForegroundColor Gray
    Write-Host ""
    
    # 09_DataflowDatasources
    Write-Host "  [9] Dataflow Data Sources..." -ForegroundColor Yellow
    
    $allDataflowDatasources = foreach ($ws in $allWorkspaceData) {
        if ($ws.dataflows) {
            foreach ($dataflow in $ws.dataflows) {
                if ($dataflow.datasourceUsages) {
                    foreach ($usage in $dataflow.datasourceUsages) {
                        $dsInstance = $dsInstanceLookup[$usage.datasourceInstanceId]
                        [PSCustomObject]@{
                            WorkspaceId       = $ws.id
                            WorkspaceName     = $ws.name
                            DataflowId        = $dataflow.objectId
                            DataflowName      = $dataflow.name
                            DatasourceId      = $usage.datasourceInstanceId
                            DatasourceType    = $dsInstance.datasourceType
                            GatewayId         = $dsInstance.gatewayId
                            ConnectionDetails = ($dsInstance.connectionDetails | ConvertTo-Json -Compress)
                            Server            = $dsInstance.connectionDetails.server
                            Database          = $dsInstance.connectionDetails.database
                            Url               = $dsInstance.connectionDetails.url
                            Path              = $dsInstance.connectionDetails.path
                        }
                    }
                }
            }
        }
    }
    
    $dataflowDsFile = Join-Path $OutputPath "09_DataflowDatasources_$timestamp.csv"
    $allDataflowDatasources | Export-Csv -Path $dataflowDsFile -NoTypeInformation -Encoding UTF8
    
    $dfWithDs = ($allDataflowDatasources | Select-Object -Unique DataflowId).Count
    Write-Host "  Exported $($allDataflowDatasources.Count) dataflow data source connections ($dfWithDs dataflows)" -ForegroundColor Green
    Write-Host "  File: $dataflowDsFile" -ForegroundColor Gray
    Write-Host ""
    
    # 10_DatasetUpstreamDataflows
    Write-Host "  [10] Dataset Upstream Dataflows..." -ForegroundColor Yellow
    
    $allUpstreamDataflows = foreach ($ws in $allWorkspaceData) {
        if ($ws.datasets) {
            foreach ($dataset in $ws.datasets) {
                if ($dataset.upstreamDataflows) {
                    foreach ($upstream in $dataset.upstreamDataflows) {
                        [PSCustomObject]@{
                            WorkspaceId         = $ws.id
                            WorkspaceName       = $ws.name
                            DatasetId           = $dataset.id
                            DatasetName         = $dataset.name
                            TargetDataflowId    = $upstream.targetDataflowId
                            DataflowWorkspaceId = $upstream.groupId
                        }
                    }
                }
            }
        }
    }
    
    $upstreamDfFile = Join-Path $OutputPath "10_DatasetUpstreamDataflows_$timestamp.csv"
    $allUpstreamDataflows | Export-Csv -Path $upstreamDfFile -NoTypeInformation -Encoding UTF8
    
    $dsWithUpstream = ($allUpstreamDataflows | Select-Object -Unique DatasetId).Count
    Write-Host "  Exported $($allUpstreamDataflows.Count) dataset-to-dataflow lineage links ($dsWithUpstream datasets)" -ForegroundColor Green
    Write-Host "  File: $upstreamDfFile" -ForegroundColor Gray
    Write-Host ""
    
    # 12_DataflowUpstreamDataflows
    Write-Host "  [12] Dataflow Upstream Dataflows..." -ForegroundColor Yellow
    
    $allDataflowUpstream = foreach ($ws in $allWorkspaceData) {
        if ($ws.dataflows) {
            foreach ($dataflow in $ws.dataflows) {
                if ($dataflow.upstreamDataflows) {
                    foreach ($upstream in $dataflow.upstreamDataflows) {
                        [PSCustomObject]@{
                            WorkspaceId               = $ws.id
                            WorkspaceName             = $ws.name
                            DataflowId                = $dataflow.objectId
                            DataflowName              = $dataflow.name
                            TargetDataflowId          = $upstream.targetDataflowId
                            TargetDataflowWorkspaceId = $upstream.groupId
                        }
                    }
                }
            }
        }
    }
    
    $dataflowUpstreamFile = Join-Path $OutputPath "12_DataflowUpstreamDataflows_$timestamp.csv"
    $allDataflowUpstream | Export-Csv -Path $dataflowUpstreamFile -NoTypeInformation -Encoding UTF8
    
    $dfWithUpstreamDf = ($allDataflowUpstream | Select-Object -Unique DataflowId).Count
    Write-Host "  Exported $($allDataflowUpstream.Count) dataflow-to-dataflow lineage links ($dfWithUpstreamDf dataflows)" -ForegroundColor Green
    Write-Host "  File: $dataflowUpstreamFile" -ForegroundColor Gray
    Write-Host ""
    
    # 13_Dashboards
    if ($exportDashboards) {
        Write-Host "  [13] Dashboards..." -ForegroundColor Yellow
        
        $allDashboards = foreach ($ws in $allWorkspaceData) {
            if ($ws.dashboards) {
                foreach ($dashboard in $ws.dashboards) {
                    $tileCount     = if ($dashboard.tiles) { $dashboard.tiles.Count } else { 0 }
                    $tileReportIds = if ($dashboard.tiles) {
                        ($dashboard.tiles | Where-Object { $_.reportId } | ForEach-Object { $_.reportId } | Select-Object -Unique) -join "; "
                    } else { $null }
                    $tileDatasetIds = if ($dashboard.tiles) {
                        ($dashboard.tiles | Where-Object { $_.datasetId } | ForEach-Object { $_.datasetId } | Select-Object -Unique) -join "; "
                    } else { $null }
                    
                    [PSCustomObject]@{
                        WorkspaceId      = $ws.id
                        WorkspaceName    = $ws.name
                        DashboardId      = $dashboard.id
                        DashboardName    = $dashboard.displayName
                        IsReadOnly       = $dashboard.isReadOnly
                        TileCount        = $tileCount
                        TileReportIds    = $tileReportIds
                        TileDatasetIds   = $tileDatasetIds
                        Endorsement      = $dashboard.endorsementDetails.endorsement
                        CertifiedBy      = $dashboard.endorsementDetails.certifiedBy
                        SensitivityLabel = if ($dashboard.sensitivityLabel) { $dashboard.sensitivityLabel.labelId } else { $null }
                    }
                }
            }
        }
        
        $dashboardsFile = Join-Path $OutputPath "13_Dashboards_$timestamp.csv"
        $allDashboards | Export-Csv -Path $dashboardsFile -NoTypeInformation -Encoding UTF8
        
        $wsWithDashboards = ($allDashboards | Select-Object -Unique WorkspaceId).Count
        Write-Host "  Exported $($allDashboards.Count) dashboards from $wsWithDashboards workspaces" -ForegroundColor Green
        Write-Host "  File: $dashboardsFile" -ForegroundColor Gray
        Write-Host ""
    }
}
#endregion

#region Step 14: Apps
if ($exportApps) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "  Admin Apps API - GetAppsAsAdmin" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[14] Retrieving Apps via GetAppsAsAdmin..." -ForegroundColor Yellow
    
    $allApps = @()
    $skip    = 0
    $pageNum = 0
    
    do {
        $pageNum++
        $url = "https://api.powerbi.com/v1.0/myorg/admin/apps?`$top=5000&`$skip=$skip"
        
        try {
            $response = Invoke-PBIRestMethod -Url $url
            $items    = $response.value
            
            if ($null -eq $items -or $items.Count -eq 0) {
                Write-Host "  Page $pageNum : empty - pagination complete" -ForegroundColor Gray
                break
            }
            
            $allApps += $items
            Write-Host "  Page $pageNum : $($items.Count) apps (total: $($allApps.Count))" -ForegroundColor Gray
            $skip += 5000
            
            if ($items.Count -lt 5000) { break }
            
            Start-Sleep -Seconds 1
        }
        catch {
            Write-Host "  Page $pageNum FAILED: $_" -ForegroundColor Red
            break
        }
    } while ($true)
    
    $appsFlat = foreach ($app in $allApps) {
        [PSCustomObject]@{
            AppId       = $app.id
            AppName     = $app.name
            Description = $app.description
            PublishedBy = $app.publishedBy
            LastUpdate  = if ($app.lastUpdate) { ([datetime]$app.lastUpdate).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
            WorkspaceId = $app.workspaceId
        }
    }
    
    $appsFile = Join-Path $OutputPath "14_Apps_$timestamp.csv"
    $appsFlat | Export-Csv -Path $appsFile -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Exported $($appsFlat.Count) apps" -ForegroundColor Green
    Write-Host "  File: $appsFile" -ForegroundColor Gray
    Write-Host ""
}
#endregion

#region Step 15: App Audiences (GetAppUsersAsAdmin — per-app, checkpoint every N apps)
if ($exportAppAudiences) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "  Step 15: App Access (GetAppUsersAsAdmin)" -ForegroundColor Cyan
    Write-Host "  Rate limit: 200 req/hour  |  Checkpoint every $AppAudienceCheckpointSize apps" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    # ----------------------------------------------------------------
    # Resolve App list — prefer in-memory $appsFlat from Step 14 if
    # available, otherwise load from the latest 14_Apps CSV on disk.
    # ----------------------------------------------------------------
    $appSource = $null
    
    if ($null -ne $appsFlat -and $appsFlat.Count -gt 0) {
        $appSource = $appsFlat
        Write-Host "  Using $($appSource.Count) apps from Step 14 (this run)" -ForegroundColor Gray
    }
    else {
        Write-Host "  Step 14 not run this session — looking for latest 14_Apps CSV..." -ForegroundColor Yellow
        $appsCSV = Get-ChildItem -Path $OutputPath -Filter "14_Apps_*.csv" -ErrorAction SilentlyContinue |
                   Sort-Object Name -Descending | Select-Object -First 1
        if ($appsCSV) {
            $appSource = Import-Csv -Path $appsCSV.FullName -Encoding UTF8
            Write-Host "  Loaded $($appSource.Count) apps from $($appsCSV.Name)" -ForegroundColor Gray
        }
        else {
            Write-Host "  ERROR: No 14_Apps CSV found in $OutputPath" -ForegroundColor Red
            Write-Host "  Run Step 14 first, or run Step 0 (all steps)." -ForegroundColor Yellow
            Write-Host ""
            # Skip this step gracefully
            $exportAppAudiences = $false
        }
    }
    
    if ($exportAppAudiences) {
        # ----------------------------------------------------------------
        # Checkpoint: load previously processed AppIds + rows if present
        # ----------------------------------------------------------------
        $checkpointFile = Join-Path $OutputPath "15_AppAccess_checkpoint.json"
        $processedAppIds  = [System.Collections.Generic.HashSet[string]]::new()
        $allAudienceRows  = [System.Collections.Generic.List[object]]::new()
        $resuming         = $false
        
        if (Test-Path $checkpointFile) {
            try {
                $checkpoint = Get-Content $checkpointFile -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($id in $checkpoint.ProcessedAppIds) {
                    [void]$processedAppIds.Add($id)
                }
                foreach ($row in $checkpoint.Rows) {
                    $allAudienceRows.Add($row)
                }
                $resuming = $true
                Write-Host "  RESUMING from checkpoint: $($processedAppIds.Count) apps already processed, $($allAudienceRows.Count) rows loaded" -ForegroundColor Yellow
            }
            catch {
                Write-Host "  Warning: Could not read checkpoint file — starting fresh. Error: $_" -ForegroundColor Yellow
                $processedAppIds.Clear()
                $allAudienceRows.Clear()
            }
        }
        
        $pendingApps = @($appSource | Where-Object { -not $processedAppIds.Contains($_.AppId) })
        $totalApps   = $appSource.Count
        $doneCount   = $processedAppIds.Count
        
        Write-Host ""
        Write-Host "  Total apps:     $totalApps" -ForegroundColor Gray
        Write-Host "  Already done:   $doneCount" -ForegroundColor Gray
        Write-Host "  Remaining:      $($pendingApps.Count)" -ForegroundColor Gray
        
        $estimatedHours = [math]::Round($pendingApps.Count / 200, 1)
        Write-Host "  Estimated time: ~$estimatedHours hours at 200 req/hour" -ForegroundColor Gray
        Write-Host ""
        
        if ($pendingApps.Count -eq 0) {
            Write-Host "  All apps already processed. Writing final CSV from checkpoint..." -ForegroundColor Green
        }
        else {
            Write-Host "[15] Fetching app users — one API call per app..." -ForegroundColor Yellow
            Write-Host ""
        }
        
        $audienceStart    = Get-Date
        $appNum           = 0
        $errorCount       = 0
        $sinceLastCheckpt = 0
        
        foreach ($app in $pendingApps) {
            $appNum++
            $sinceLastCheckpt++
            $overallDone = $doneCount + $appNum
            
            $url = "https://api.powerbi.com/v1.0/myorg/admin/apps/$($app.AppId)/users"
            
            try {
                $response = Invoke-PBIRestMethod -Url $url
                $users    = $response.value
                
                if ($users -and $users.Count -gt 0) {
                    foreach ($user in $users) {
                        $row = [PSCustomObject]@{
                            AppId         = $app.AppId
                            AppName       = $app.AppName
                            WorkspaceId   = $app.WorkspaceId
                            Identifier    = $user.identifier
                            DisplayName   = $user.displayName
                            EmailAddress  = $user.emailAddress
                            GraphId       = $user.graphId
                            PrincipalType = $user.principalType
                            AppUserAccessRight = $user.appUserAccessRight
                        }
                        $allAudienceRows.Add($row)
                    }
                    Write-Host "  [$overallDone/$totalApps] $($app.AppName) — $($users.Count) user(s)" -ForegroundColor Gray
                }
                else {
                    Write-Host "  [$overallDone/$totalApps] $($app.AppName) — no users" -ForegroundColor DarkGray
                }
                
                [void]$processedAppIds.Add($app.AppId)
            }
            catch {
                $errorCount++
                Write-Host "  [$overallDone/$totalApps] $($app.AppName) — FAILED: $_" -ForegroundColor Red
                # Still mark as processed to avoid infinite retry loops on
                # permanently failing apps (e.g. deleted apps still in CSV)
                [void]$processedAppIds.Add($app.AppId)
            }
            
            # ----------------------------------------------------------------
            # Checkpoint every N apps
            # ----------------------------------------------------------------
            if ($sinceLastCheckpt -ge $AppAudienceCheckpointSize) {
                $sinceLastCheckpt = 0
                $elapsed          = (Get-Date) - $audienceStart
                $appsRemaining    = $pendingApps.Count - $appNum
                $estimatedLeft    = if ($appNum -gt 0) {
                    [math]::Round(($elapsed.TotalMinutes / $appNum) * $appsRemaining, 0)
                } else { 0 }
                
                Write-Host "" 
                Write-Host "  --- Checkpoint: $overallDone/$totalApps apps done | $($allAudienceRows.Count) rows | ~${estimatedLeft} min remaining ---" -ForegroundColor Cyan
                
                try {
                    $checkpointData = @{
                        ProcessedAppIds = @($processedAppIds)
                        Rows            = @($allAudienceRows)
                        SavedAt         = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
                        TotalApps       = $totalApps
                    }
                    $checkpointData | ConvertTo-Json -Depth 5 -Compress | 
                        Set-Content -Path $checkpointFile -Encoding UTF8
                    Write-Host "  Checkpoint saved: $checkpointFile" -ForegroundColor DarkGray
                }
                catch {
                    Write-Host "  Warning: Could not save checkpoint: $_" -ForegroundColor Yellow
                }
                Write-Host ""
            }
        }
        
        # ----------------------------------------------------------------
        # Write final CSV
        # ----------------------------------------------------------------
        $audienceElapsed = (Get-Date) - $audienceStart
        Write-Host ""
        Write-Host "  Fetch complete: $($processedAppIds.Count)/$totalApps apps processed in $($audienceElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        if ($errorCount -gt 0) {
            Write-Host "  Errors: $errorCount apps failed (marked processed, excluded from CSV)" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "  Writing 15_AppAccess CSV..." -ForegroundColor Yellow
        
        $audiencesFile = Join-Path $OutputPath "15_AppAccess_$timestamp.csv"
        
        if ($allAudienceRows.Count -gt 0) {
            $allAudienceRows | Export-Csv -Path $audiencesFile -NoTypeInformation -Encoding UTF8
            Write-Host "  Exported $($allAudienceRows.Count) app audience records" -ForegroundColor Green
        }
        else {
            # Write an empty CSV with correct headers
            [PSCustomObject]@{
                AppId = $null; AppName = $null; WorkspaceId = $null
                Identifier = $null; DisplayName = $null; EmailAddress = $null
                GraphId = $null; PrincipalType = $null; AppUserAccessRight = $null
            } | Export-Csv -Path $audiencesFile -NoTypeInformation -Encoding UTF8
            Write-Host "  No audience records found — empty CSV written" -ForegroundColor Yellow
        }
        
        Write-Host "  File: $audiencesFile" -ForegroundColor Gray
        
        # ----------------------------------------------------------------
        # Remove checkpoint on successful completion
        # ----------------------------------------------------------------
        if ($processedAppIds.Count -ge $totalApps -and (Test-Path $checkpointFile)) {
            Remove-Item $checkpointFile -Force
            Write-Host "  Checkpoint file removed (step completed successfully)" -ForegroundColor DarkGray
        }
        
        Write-Host ""
        
        # Breakdown by principal type
        if ($allAudienceRows.Count -gt 0) {
            Write-Host "  -- App Access Breakdown by Principal Type --" -ForegroundColor Yellow
            $allAudienceRows | Group-Object PrincipalType | Sort-Object Count -Descending | ForEach-Object {
                Write-Host "     $($_.Name): $($_.Count)" -ForegroundColor Gray
            }
            $appsWithUsers = ($allAudienceRows | Select-Object -Unique AppId).Count
            $appsNoUsers   = $totalApps - $appsWithUsers
            Write-Host ""
            Write-Host "     Apps with audience: $appsWithUsers" -ForegroundColor Gray
            Write-Host "     Apps with no users: $appsNoUsers" -ForegroundColor Gray
            Write-Host ""
        }
    }
}
#endregion

#region Step 7: Gateways
if ($exportGateways) {
    Write-Host "[7] Exporting Gateway Clusters (07_1) and Gateway Datasources (07_2)..." -ForegroundColor Yellow
    
    try {
        $gatewaysResponse = Invoke-PBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/gateways"
        
        $allGateways           = @()
        $allGatewayDatasources = @()
        
        foreach ($gateway in $gatewaysResponse.value) {
            $allGateways += [PSCustomObject]@{
                GatewayId           = $gateway.id
                GatewayName         = $gateway.name
                GatewayType         = $gateway.type
                GatewayAnnotation   = $gateway.gatewayAnnotation
                PublicKey_Exponent  = $gateway.publicKey.exponent
                PublicKey_Modulus   = $gateway.publicKey.modulus
                GatewayStatus       = $gateway.gatewayStatus
                GatewayMachineCount = $gateway.gatewayMachineCount
            }
            
            try {
                $gwDsResponse = Invoke-PBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/gateways/$($gateway.id)/datasources"
                
                foreach ($gds in $gwDsResponse.value) {
                    $allGatewayDatasources += [PSCustomObject]@{
                        GatewayId               = $gateway.id
                        GatewayName             = $gateway.name
                        GatewayDatasourceId     = $gds.id
                        DatasourceName          = $gds.datasourceName
                        DatasourceType          = $gds.datasourceType
                        ConnectionDetails       = ($gds.connectionDetails | ConvertTo-Json -Compress)
                        CredentialType          = $gds.credentialType
                        CredentialDetails_UseEndUserOAuth2Credentials = $gds.credentialDetails.useEndUserOAuth2Credentials
                        GatewayDatasourceStatus = $gds.gatewayDatasourceStatus
                    }
                }
            }
            catch {
                Write-Host "  Warning: Could not get datasources for gateway $($gateway.name)" -ForegroundColor DarkYellow
            }
        }
        
        $gatewaysFile = Join-Path $OutputPath "07_1_GatewayClusters_$timestamp.csv"
        $allGateways | Export-Csv -Path $gatewaysFile -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported $($allGateways.Count) gateway clusters" -ForegroundColor Green
        Write-Host "  File: $gatewaysFile" -ForegroundColor Gray
        
        $gatewayDatasourcesFile = Join-Path $OutputPath "07_2_GatewayDatasources_$timestamp.csv"
        $allGatewayDatasources | Export-Csv -Path $gatewayDatasourcesFile -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported $($allGatewayDatasources.Count) gateway data source definitions" -ForegroundColor Green
        Write-Host "  File: $gatewayDatasourcesFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "  Warning: Could not export gateway details. Error: $_" -ForegroundColor Yellow
    }
    Write-Host ""
}
#endregion

#region Step 11: Refreshables
if ($exportRefreshables) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "  Admin Refreshables API - Dataset Refresh Data" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""
    
    $refreshStart    = Get-Date
    $allRefreshables = @()
    $skip            = 0
    $pageNum         = 0
    
    Write-Host "[11] Retrieving Refreshables (page size: $RefreshablePageSize)..." -ForegroundColor Yellow
    
    do {
        $pageNum++
        Ensure-FreshToken
        
        $url = "https://api.powerbi.com/v1.0/myorg/admin/capacities/refreshables?`$top=$RefreshablePageSize&`$skip=$skip&`$expand=capacity,group"
        
        try {
            $response = Invoke-PBIRestMethod -Url $url
            $items    = $response.value
            
            if ($items -and $items.Count -gt 0) {
                $allRefreshables += $items
                Write-Host "  Page $pageNum : $($items.Count) refreshables (total: $($allRefreshables.Count))" -ForegroundColor Gray
                $skip += $RefreshablePageSize
            }
            else {
                Write-Host "  Page $pageNum : empty - pagination complete" -ForegroundColor Gray
                break
            }
        }
        catch {
            Write-Host "  Page $pageNum FAILED: $_" -ForegroundColor Red
            break
        }
        
        if ($items.Count -lt $RefreshablePageSize) {
            Write-Host "  Last page reached ($($items.Count) < $RefreshablePageSize)" -ForegroundColor Gray
            break
        }
    } while ($true)
    
    $refreshElapsed = (Get-Date) - $refreshStart
    Write-Host ""
    Write-Host "  Retrieved $($allRefreshables.Count) refreshable datasets in $($refreshElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
    Write-Host "  API calls used: $pageNum (of 200/hour limit)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Flattening and exporting..." -ForegroundColor Yellow
    
    $refreshablesFlat = foreach ($r in $allRefreshables) {
        [PSCustomObject]@{
            DatasetId               = $r.id
            DatasetName             = $r.name
            Kind                    = $r.kind
            ConfiguredBy            = ($r.configuredBy -join "; ")
            WorkspaceId             = $r.group.id
            WorkspaceName           = $r.group.name
            CapacityId              = $r.capacity.id
            CapacityName            = $r.capacity.displayName
            CapacitySku             = $r.capacity.sku
            CapacityState           = $r.capacity.state
            LastRefreshStatus       = $r.lastRefresh.status
            LastRefreshType         = $r.lastRefresh.refreshType
            LastRefreshStartTime    = if ($r.lastRefresh.startTime) { ([datetime]$r.lastRefresh.startTime).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
            LastRefreshEndTime      = if ($r.lastRefresh.endTime)   { ([datetime]$r.lastRefresh.endTime).ToString("yyyy-MM-ddTHH:mm:ss")   } else { $null }
            LastRefreshDurationSec  = if ($r.lastRefresh.startTime -and $r.lastRefresh.endTime) { 
                                           ([datetime]$r.lastRefresh.endTime - [datetime]$r.lastRefresh.startTime).TotalSeconds 
                                      } else { $null }
            LastRefreshRequestId    = $r.lastRefresh.requestId
            LastRefreshServiceError = $r.lastRefresh.serviceExceptionJson
            WindowStartTime         = if ($r.startTime) { ([datetime]$r.startTime).ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
            WindowEndTime           = if ($r.endTime)   { ([datetime]$r.endTime).ToString("yyyy-MM-ddTHH:mm:ss")   } else { $null }
            RefreshCount            = $r.refreshCount
            RefreshFailures         = $r.refreshFailures
            RefreshesPerDay         = $r.refreshesPerDay
            AverageDurationSec      = $r.averageDuration
            MedianDurationSec       = $r.medianDuration
            ScheduleEnabled         = $r.refreshSchedule.enabled
            ScheduleDays            = ($r.refreshSchedule.days  -join ", ")
            ScheduleTimes           = ($r.refreshSchedule.times -join ", ")
            ScheduleTimeZone        = $r.refreshSchedule.localTimeZoneId
            ScheduleNotifyOption    = $r.refreshSchedule.notifyOption
        }
    }
    
    $refreshablesFile = Join-Path $OutputPath "11_Refreshables_$timestamp.csv"
    $refreshablesFlat | Export-Csv -Path $refreshablesFile -NoTypeInformation -Encoding UTF8
    
    Write-Host "  Exported $($refreshablesFlat.Count) refreshable datasets" -ForegroundColor Green
    Write-Host "  File: $refreshablesFile" -ForegroundColor Gray
    Write-Host ""
    
    $completed = ($refreshablesFlat | Where-Object { $_.LastRefreshStatus -eq "Completed" }).Count
    $failed    = ($refreshablesFlat | Where-Object { $_.LastRefreshStatus -eq "Failed"    }).Count
    $disabled  = ($refreshablesFlat | Where-Object { $_.LastRefreshStatus -eq "Disabled"  }).Count
    $unknown   = ($refreshablesFlat | Where-Object { $_.LastRefreshStatus -eq "Unknown"   }).Count
    $other     = $refreshablesFlat.Count - $completed - $failed - $disabled - $unknown
    $scheduled   = ($refreshablesFlat | Where-Object { $_.ScheduleEnabled -eq $true }).Count
    $unscheduled = $refreshablesFlat.Count - $scheduled
    $capacities  = ($refreshablesFlat | Select-Object -Unique CapacityId).Count
    $workspaces  = ($refreshablesFlat | Select-Object -Unique WorkspaceId).Count
    
    Write-Host "  -- Refresh Status --" -ForegroundColor Yellow
    Write-Host "     Completed: $completed" -ForegroundColor Green
    Write-Host "     Failed:    $failed"    -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Gray" })
    Write-Host "     Disabled:  $disabled"  -ForegroundColor Gray
    Write-Host "     Unknown:   $unknown"   -ForegroundColor Gray
    if ($other -gt 0) { Write-Host "     Other:     $other" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "  -- Schedule --" -ForegroundColor Yellow
    Write-Host "     Scheduled:   $scheduled"   -ForegroundColor Gray
    Write-Host "     Unscheduled: $unscheduled" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  -- Coverage --" -ForegroundColor Yellow
    Write-Host "     Capacities: $capacities" -ForegroundColor Gray
    Write-Host "     Workspaces: $workspaces"  -ForegroundColor Gray
    Write-Host ""
}
#endregion

#region Step 16: Dashboard Tiles (GetTilesAsAdmin — per-dashboard, checkpoint every N dashboards)
if ($exportDashboardTiles) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "  Step 16: Dashboard Tiles (GetTilesAsAdmin)" -ForegroundColor Cyan
    Write-Host "  Rate limit: 200 req/hour  |  Checkpoint every $DashboardTileCheckpointSize dashboards" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host ""

    # ----------------------------------------------------------------
    # Resolve Dashboard list — prefer in-memory $allDashboards from
    # Step 13 if available, otherwise load from latest 13_Dashboards CSV.
    # ----------------------------------------------------------------
    $dashboardSource = $null

    if ($null -ne $allDashboards -and $allDashboards.Count -gt 0) {
        $dashboardSource = $allDashboards
        Write-Host "  Using $($dashboardSource.Count) dashboards from Step 13 (this run)" -ForegroundColor Gray
    }
    else {
        Write-Host "  Step 13 not run this session — looking for latest 13_Dashboards CSV..." -ForegroundColor Yellow
        $dashboardsCSV = Get-ChildItem -Path $OutputPath -Filter "13_Dashboards_*.csv" -ErrorAction SilentlyContinue |
                         Sort-Object Name -Descending | Select-Object -First 1
        if ($dashboardsCSV) {
            $dashboardSource = Import-Csv -Path $dashboardsCSV.FullName -Encoding UTF8
            Write-Host "  Loaded $($dashboardSource.Count) dashboards from $($dashboardsCSV.Name)" -ForegroundColor Gray
        }
        else {
            Write-Host "  ERROR: No 13_Dashboards CSV found in $OutputPath" -ForegroundColor Red
            Write-Host "  Run Step 13 first, or run Step 0 (all steps)." -ForegroundColor Yellow
            Write-Host ""
            $exportDashboardTiles = $false
        }
    }

    if ($exportDashboardTiles) {
        # ----------------------------------------------------------------
        # Checkpoint: load previously processed DashboardIds + rows
        # ----------------------------------------------------------------
        $checkpointFile       = Join-Path $OutputPath "16_DashboardTiles_checkpoint.json"
        $processedDashIds     = [System.Collections.Generic.HashSet[string]]::new()
        $allTileRows          = [System.Collections.Generic.List[object]]::new()

        if (Test-Path $checkpointFile) {
            try {
                $checkpoint = Get-Content $checkpointFile -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($id in $checkpoint.ProcessedDashboardIds) {
                    [void]$processedDashIds.Add($id)
                }
                foreach ($row in $checkpoint.Rows) {
                    $allTileRows.Add($row)
                }
                Write-Host "  RESUMING from checkpoint: $($processedDashIds.Count) dashboards already processed, $($allTileRows.Count) tile rows loaded" -ForegroundColor Yellow
            }
            catch {
                Write-Host "  Warning: Could not read checkpoint file — starting fresh. Error: $_" -ForegroundColor Yellow
                $processedDashIds.Clear()
                $allTileRows.Clear()
            }
        }

        $pendingDashboards = @($dashboardSource | Where-Object { -not $processedDashIds.Contains($_.DashboardId) })
        $totalDashboards   = $dashboardSource.Count
        $doneCount         = $processedDashIds.Count

        Write-Host ""
        Write-Host "  Total dashboards: $totalDashboards" -ForegroundColor Gray
        Write-Host "  Already done:     $doneCount" -ForegroundColor Gray
        Write-Host "  Remaining:        $($pendingDashboards.Count)" -ForegroundColor Gray

        $estimatedHours = [math]::Round($pendingDashboards.Count / 200, 1)
        Write-Host "  Estimated time:   ~$estimatedHours hours at 200 req/hour" -ForegroundColor Gray
        Write-Host ""

        if ($pendingDashboards.Count -eq 0) {
            Write-Host "  All dashboards already processed. Writing final CSV from checkpoint..." -ForegroundColor Green
        }
        else {
            Write-Host "[16] Fetching tiles — one API call per dashboard..." -ForegroundColor Yellow
            Write-Host ""
        }

        $tilesStart        = Get-Date
        $dashNum           = 0
        $errorCount        = 0
        $sinceLastCheckpt  = 0

        foreach ($dashboard in $pendingDashboards) {
            $dashNum++
            $sinceLastCheckpt++
            $overallDone = $doneCount + $dashNum

            $url = "https://api.powerbi.com/v1.0/myorg/admin/dashboards/$($dashboard.DashboardId)/tiles"

            try {
                $response = Invoke-PBIRestMethod -Url $url
                $tiles    = $response.value

                if ($tiles -and $tiles.Count -gt 0) {
                    foreach ($tile in $tiles) {
                        $row = [PSCustomObject]@{
                            DashboardId   = $dashboard.DashboardId
                            DashboardName = $dashboard.DashboardName
                            WorkspaceId   = $dashboard.WorkspaceId
                            WorkspaceName = $dashboard.WorkspaceName
                            TileId        = $tile.id
                            TileTitle     = $tile.title
                            TileType      = $tile.subType
                            ReportId      = $tile.reportId
                            DatasetId     = $tile.datasetId
                            RowSpan       = $tile.rowSpan
                            ColSpan       = $tile.colSpan
                            EmbedUrl      = $tile.embedUrl
                        }
                        $allTileRows.Add($row)
                    }
                    Write-Host "  [$overallDone/$totalDashboards] $($dashboard.DashboardName) — $($tiles.Count) tile(s)" -ForegroundColor Gray
                }
                else {
                    Write-Host "  [$overallDone/$totalDashboards] $($dashboard.DashboardName) — no tiles" -ForegroundColor DarkGray
                }

                [void]$processedDashIds.Add($dashboard.DashboardId)
            }
            catch {
                $errorCount++
                Write-Host "  [$overallDone/$totalDashboards] $($dashboard.DashboardName) — FAILED: $_" -ForegroundColor Red
                [void]$processedDashIds.Add($dashboard.DashboardId)
            }

            # ----------------------------------------------------------------
            # Checkpoint every N dashboards
            # ----------------------------------------------------------------
            if ($sinceLastCheckpt -ge $DashboardTileCheckpointSize) {
                $sinceLastCheckpt = 0
                $elapsed          = (Get-Date) - $tilesStart
                $dashRemaining    = $pendingDashboards.Count - $dashNum
                $estimatedLeft    = if ($dashNum -gt 0) {
                    [math]::Round(($elapsed.TotalMinutes / $dashNum) * $dashRemaining, 0)
                } else { 0 }

                Write-Host ""
                Write-Host "  --- Checkpoint: $overallDone/$totalDashboards dashboards done | $($allTileRows.Count) tile rows | ~${estimatedLeft} min remaining ---" -ForegroundColor Cyan

                try {
                    $checkpointData = @{
                        ProcessedDashboardIds = @($processedDashIds)
                        Rows                  = @($allTileRows)
                        SavedAt               = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
                        TotalDashboards       = $totalDashboards
                    }
                    $checkpointData | ConvertTo-Json -Depth 5 -Compress |
                        Set-Content -Path $checkpointFile -Encoding UTF8
                    Write-Host "  Checkpoint saved: $checkpointFile" -ForegroundColor DarkGray
                }
                catch {
                    Write-Host "  Warning: Could not save checkpoint: $_" -ForegroundColor Yellow
                }
                Write-Host ""
            }
        }

        # ----------------------------------------------------------------
        # Write final CSV
        # ----------------------------------------------------------------
        $tilesElapsed = (Get-Date) - $tilesStart
        Write-Host ""
        Write-Host "  Fetch complete: $($processedDashIds.Count)/$totalDashboards dashboards processed in $($tilesElapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
        if ($errorCount -gt 0) {
            Write-Host "  Errors: $errorCount dashboards failed (marked processed, excluded from CSV)" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "  Writing 16_DashboardTiles CSV..." -ForegroundColor Yellow

        $tilesFile = Join-Path $OutputPath "16_DashboardTiles_$timestamp.csv"

        if ($allTileRows.Count -gt 0) {
            $allTileRows | Export-Csv -Path $tilesFile -NoTypeInformation -Encoding UTF8
            Write-Host "  Exported $($allTileRows.Count) tile rows" -ForegroundColor Green
        }
        else {
            [PSCustomObject]@{
                DashboardId = $null; DashboardName = $null; WorkspaceId = $null
                WorkspaceName = $null; TileId = $null; TileTitle = $null
                TileType = $null; ReportId = $null; DatasetId = $null
                RowSpan = $null; ColSpan = $null; EmbedUrl = $null
            } | Export-Csv -Path $tilesFile -NoTypeInformation -Encoding UTF8
            Write-Host "  No tile rows found — empty CSV written" -ForegroundColor Yellow
        }

        Write-Host "  File: $tilesFile" -ForegroundColor Gray

        # ----------------------------------------------------------------
        # Remove checkpoint on successful completion
        # ----------------------------------------------------------------
        if ($processedDashIds.Count -ge $totalDashboards -and (Test-Path $checkpointFile)) {
            Remove-Item $checkpointFile -Force
            Write-Host "  Checkpoint file removed (step completed successfully)" -ForegroundColor DarkGray
        }

        Write-Host ""

        # Breakdown by tile type
        if ($allTileRows.Count -gt 0) {
            Write-Host "  -- Tile Type Breakdown --" -ForegroundColor Yellow
            $allTileRows | Group-Object TileType | Sort-Object Count -Descending | ForEach-Object {
                $typeLabel = if ([string]::IsNullOrEmpty($_.Name)) { "(unknown)" } else { $_.Name }
                Write-Host "     $typeLabel`: $($_.Count)" -ForegroundColor Gray
            }
            $webContentCount  = ($allTileRows | Where-Object { $_.TileType -eq "WebContent" }).Count
            $alertableCount   = ($allTileRows | Where-Object { $_.TileType -in @("Card", "KPI", "Gauge") }).Count
            $dashWithTiles    = ($allTileRows | Select-Object -Unique DashboardId).Count
            $dashWithoutTiles = $totalDashboards - $dashWithTiles
            Write-Host ""
            Write-Host "     WebContent tiles (security risk):  $webContentCount" -ForegroundColor $(if ($webContentCount -gt 0) { "Red" } else { "Gray" })
            Write-Host "     Alertable tiles (Card/KPI/Gauge):  $alertableCount" -ForegroundColor Gray
            Write-Host "     Dashboards with tiles:             $dashWithTiles" -ForegroundColor Gray
            Write-Host "     Dashboards with no tiles:          $dashWithoutTiles" -ForegroundColor Gray
            Write-Host ""
        }
    }
}
#endregion

#region Summary
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Inventory Complete!" -ForegroundColor Green
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Generated Files:" -ForegroundColor Yellow
Get-ChildItem -Path $OutputPath -Filter "*_$timestamp.csv" -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
    $sizeKB = [math]::Round($_.Length / 1KB, 1)
    Write-Host "  $($_.Name) ($sizeKB KB)" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Output Directory: $OutputPath" -ForegroundColor Cyan
#endregion

#region Disconnect
Disconnect-PowerBIServiceAccount | Out-Null
Write-Host ""
Write-Host "Disconnected from Power BI Service." -ForegroundColor Gray
Write-Host ""
Write-Host "Thank you for preparing the Power BI Inventory data!" -ForegroundColor Yellow
#endregion
