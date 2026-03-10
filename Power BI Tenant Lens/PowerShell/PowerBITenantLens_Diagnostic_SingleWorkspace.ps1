<#
.SYNOPSIS
    PowerBITenantLens_Diagnostic.ps1 — Dump raw Scanner API JSON for one workspace

.DESCRIPTION
    Scans a single workspace via the Scanner API and saves the COMPLETE raw JSON
    response to disk for inspection. Specifically designed to investigate what
    properties exist on dataflow objects (e.g. upstreamDataflows, datasourceUsages)
    and whether datasourceInstances at the root level cover all scenarios.

    Outputs:
    - ScanResult_FULL.json         — The complete GetScanResult response
    - ScanResult_ROOT_KEYS.json    — Top-level property names of the scan result
    - ScanResult_datasourceInstances.json — Root-level datasourceInstances array
    - ScanResult_workspace.json    — The workspace object (with all nested artifacts)
    - ScanResult_dataflows.json    — All dataflow objects for this workspace
    - ScanResult_dataflow_PROPERTIES.txt — Property names found on each dataflow
    - ScanResult_datasets.json     — All dataset objects for this workspace
    - ScanResult_dataset_PROPERTIES.txt — Property names found on each dataset (for comparison)

.PARAMETER WorkspaceId
    The workspace ID to scan. Required — provide your own workspace ID.

.PARAMETER OutputPath
    Directory where JSON files will be saved. Defaults to current directory.
#>

param(
    [string]$WorkspaceId = "",
    [string]$OutputPath = ".",
    [int]$PollIntervalSeconds = 3,
    [int]$PollTimeoutSeconds = 120,
    [int]$TokenRefreshMinutes = 45
)

$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"

if ([string]::IsNullOrWhiteSpace($WorkspaceId)) {
    Write-Host "ERROR: -WorkspaceId is required. Provide a workspace GUID to scan." -ForegroundColor Red
    Write-Host "  Example: pwsh ./PowerBITenantLens_Diagnostic.ps1 -WorkspaceId 'your-workspace-id-here'" -ForegroundColor Yellow
    exit 1
}

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7 or higher." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Scanner API Diagnostic — Single Workspace" -ForegroundColor Cyan
Write-Host "  Workspace: $WorkspaceId" -ForegroundColor Yellow
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

#region Module & Auth
$moduleName = "MicrosoftPowerBIMgmt"
$installedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
if (-not $installedModule) {
    Write-Host "Installing $moduleName module..." -ForegroundColor Yellow
    Install-Module -Name $moduleName -Scope CurrentUser -Force
}
Import-Module $moduleName

Write-Host "Connecting to Power BI Service..." -ForegroundColor Yellow
try {
    Connect-PowerBIServiceAccount | Out-Null
    Write-Host "Connected successfully!" -ForegroundColor Green
    $script:AccessToken = (Get-PowerBIAccessToken -AsString).Replace("Bearer ", "")
    $script:TokenAcquiredAt = Get-Date
}
catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
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
            Write-Host "  Token refreshed" -ForegroundColor Green
        }
        catch {
            Connect-PowerBIServiceAccount | Out-Null
            $script:AccessToken = (Get-PowerBIAccessToken -AsString).Replace("Bearer ", "")
            $script:TokenAcquiredAt = Get-Date
            Write-Host "  Reconnected and token refreshed" -ForegroundColor Green
        }
    }
}
#endregion

#region Helper: REST call with retry
function Invoke-PBIRestMethod {
    param(
        [string]$Url,
        [string]$Method = "Get",
        [string]$Body = $null,
        [int]$MaxRetries = 3
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
            return (Invoke-RestMethod @params)
        }
        catch {
            $statusCode = $null
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch {}
            
            if ($statusCode -eq 429 -and $retryCount -lt $MaxRetries) {
                $retryCount++
                $wait = 30
                try {
                    $ra = $_.Exception.Response.Headers["Retry-After"]
                    if ($ra) { $wait = [int]$ra }
                } catch {}
                Write-Host "    429 throttled - waiting ${wait}s (retry $retryCount)..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
            }
            elseif ($statusCode -eq 401 -and $retryCount -lt $MaxRetries) {
                $retryCount++
                Write-Host "    401 - refreshing token (retry $retryCount)..." -ForegroundColor DarkYellow
                $script:TokenAcquiredAt = [datetime]::MinValue
                Ensure-FreshToken
                $headers["Authorization"] = "Bearer $script:AccessToken"
            }
            elseif (($statusCode -in 502,503,504 -or $null -eq $statusCode) -and $retryCount -lt $MaxRetries) {
                $retryCount++
                $wait = [math]::Pow(2, $retryCount) + (Get-Random -Minimum 1 -Maximum 5)
                Write-Host "    HTTP $statusCode - waiting ${wait}s (retry $retryCount)..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
            }
            else {
                throw
            }
        }
    }
}
#endregion

# ================================================================
# STEP 1: PostWorkspaceInfo for single workspace
# ================================================================

Write-Host "[1/4] Submitting scan request for workspace $WorkspaceId..." -ForegroundColor Yellow

$scanUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?datasourceDetails=True&getArtifactUsers=True&lineage=True"
$body = "{ `"workspaces`": [ `"$WorkspaceId`" ] }"

$scanResponse = Invoke-PBIRestMethod -Url $scanUrl -Method "Post" -Body $body
$scanId = $scanResponse.id
Write-Host "  Scan submitted. Scan ID: $scanId" -ForegroundColor Green
Write-Host ""

# ================================================================
# STEP 2: Poll for completion
# ================================================================

Write-Host "[2/4] Waiting for scan to complete..." -ForegroundColor Yellow
$pollStart = Get-Date

while ($true) {
    if (((Get-Date) - $pollStart).TotalSeconds -gt $PollTimeoutSeconds) {
        Write-Host "  TIMEOUT after $PollTimeoutSeconds seconds" -ForegroundColor Red
        exit 1
    }
    
    $statusUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$scanId"
    $statusResponse = Invoke-PBIRestMethod -Url $statusUrl
    
    if ($statusResponse.status -eq "Succeeded") {
        Write-Host "  Scan completed!" -ForegroundColor Green
        break
    }
    elseif ($statusResponse.status -eq "Failed") {
        Write-Host "  Scan FAILED!" -ForegroundColor Red
        Write-Host "  Response: $($statusResponse | ConvertTo-Json -Depth 5)" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "  Status: $($statusResponse.status) - waiting ${PollIntervalSeconds}s..." -ForegroundColor Gray
    Start-Sleep -Seconds $PollIntervalSeconds
}
Write-Host ""

# ================================================================
# STEP 3: Get scan result
# ================================================================

Write-Host "[3/4] Retrieving scan result..." -ForegroundColor Yellow

$resultUrl = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$scanId"
$scanResult = Invoke-PBIRestMethod -Url $resultUrl

Write-Host "  Scan result retrieved." -ForegroundColor Green
Write-Host ""

# ================================================================
# STEP 4: Dump everything to JSON files
# ================================================================

Write-Host "[4/4] Dumping JSON files..." -ForegroundColor Yellow

$outPrefix = Join-Path $OutputPath "ScanResult_${timestamp}"

# --- 1. FULL scan result ---
$fullPath = "${outPrefix}_FULL.json"
$scanResult | ConvertTo-Json -Depth 100 | Set-Content -Path $fullPath -Encoding UTF8
$fullSize = [math]::Round((Get-Item $fullPath).Length / 1KB, 1)
Write-Host "  [1] FULL scan result: $fullPath ($fullSize KB)" -ForegroundColor Green

# --- 2. Root-level property names ---
$rootKeys = $scanResult.PSObject.Properties | ForEach-Object {
    [PSCustomObject]@{
        Name     = $_.Name
        TypeName = if ($_.Value -is [array]) { "Array[$($_.Value.Count)]" }
                   elseif ($_.Value -is [PSCustomObject]) { "Object" }
                   elseif ($null -eq $_.Value) { "null" }
                   else { $_.Value.GetType().Name }
    }
}
$rootKeysPath = "${outPrefix}_ROOT_KEYS.json"
$rootKeys | ConvertTo-Json -Depth 5 | Set-Content -Path $rootKeysPath -Encoding UTF8
Write-Host "  [2] Root-level keys: $rootKeysPath" -ForegroundColor Green
Write-Host ""
Write-Host "      Root-level properties:" -ForegroundColor Cyan
foreach ($k in $rootKeys) {
    Write-Host "        $($k.Name) ($($k.TypeName))" -ForegroundColor White
}
Write-Host ""

# --- 3. datasourceInstances (from root) ---
$dsInstancesPath = "${outPrefix}_datasourceInstances.json"
if ($scanResult.datasourceInstances) {
    $scanResult.datasourceInstances | ConvertTo-Json -Depth 50 | Set-Content -Path $dsInstancesPath -Encoding UTF8
    Write-Host "  [3] datasourceInstances (root): $($scanResult.datasourceInstances.Count) entries -> $dsInstancesPath" -ForegroundColor Green
}
else {
    "[]" | Set-Content -Path $dsInstancesPath -Encoding UTF8
    Write-Host "  [3] datasourceInstances (root): EMPTY / NOT PRESENT" -ForegroundColor Yellow
}

# --- 4. Workspace object ---
$ws = $scanResult.workspaces | Where-Object { $_.id -eq $WorkspaceId }
if (-not $ws -and $scanResult.workspaces.Count -gt 0) {
    $ws = $scanResult.workspaces[0]
    Write-Host "  Note: workspace ID not matched, using first workspace: $($ws.id)" -ForegroundColor Yellow
}

$wsPath = "${outPrefix}_workspace.json"
if ($ws) {
    $ws | ConvertTo-Json -Depth 100 | Set-Content -Path $wsPath -Encoding UTF8
    Write-Host "  [4] Workspace object: $wsPath" -ForegroundColor Green
    Write-Host "      Workspace: $($ws.name) ($($ws.id))" -ForegroundColor Cyan
    
    # Summarize what's in the workspace
    $artifactTypes = @("reports", "datasets", "dataflows", "datamarts", "dashboards")
    foreach ($at in $artifactTypes) {
        $count = if ($ws.$at) { $ws.$at.Count } else { 0 }
        if ($count -gt 0) {
            Write-Host "      $at : $count" -ForegroundColor White
        }
    }
    Write-Host ""
}

# --- 5. Dataflows (the main target of this diagnostic) ---
$dataflowsPath = "${outPrefix}_dataflows.json"
$dataflowPropsPath = "${outPrefix}_dataflow_PROPERTIES.txt"

if ($ws.dataflows -and $ws.dataflows.Count -gt 0) {
    $ws.dataflows | ConvertTo-Json -Depth 100 | Set-Content -Path $dataflowsPath -Encoding UTF8
    Write-Host "  [5] Dataflows: $($ws.dataflows.Count) entries -> $dataflowsPath" -ForegroundColor Green
    
    # Dump property names for EVERY dataflow
    $allDfProps = @()
    $propReport = @()
    foreach ($df in $ws.dataflows) {
        $props = $df.PSObject.Properties.Name
        $allDfProps += $props
        
        $propReport += ""
        $propReport += "=== $($df.name) (objectId: $($df.objectId)) ==="
        $propReport += "Properties:"
        foreach ($p in $props) {
            $val = $df.$p
            $typeInfo = if ($val -is [array]) { "Array[$($val.Count)]" }
                        elseif ($val -is [PSCustomObject]) { "Object" }
                        elseif ($null -eq $val) { "null" }
                        else { "$($val.GetType().Name): $($val.ToString().Substring(0, [math]::Min(80, $val.ToString().Length)))" }
            $propReport += "  $p = $typeInfo"
        }
        
        # Specifically highlight datasource-related properties
        $dsProps = $props | Where-Object { $_ -match "datasource|upstream|source|usage|linked|reference" }
        if ($dsProps) {
            $propReport += ""
            $propReport += "DATASOURCE-RELATED PROPERTIES:"
            foreach ($dp in $dsProps) {
                $val = $df.$dp
                if ($val -is [array]) {
                    $propReport += "  $dp = Array[$($val.Count)]"
                    if ($val.Count -gt 0 -and $val.Count -le 5) {
                        foreach ($item in $val) {
                            $propReport += "    -> $($item | ConvertTo-Json -Compress -Depth 10)"
                        }
                    }
                    elseif ($val.Count -gt 5) {
                        for ($i = 0; $i -lt 3; $i++) {
                            $propReport += "    -> $($val[$i] | ConvertTo-Json -Compress -Depth 10)"
                        }
                        $propReport += "    ... and $($val.Count - 3) more"
                    }
                }
                elseif ($val -is [PSCustomObject]) {
                    $propReport += "  $dp = $($val | ConvertTo-Json -Compress -Depth 10)"
                }
                else {
                    $propReport += "  $dp = $val"
                }
            }
        }
        else {
            $propReport += ""
            $propReport += "DATASOURCE-RELATED PROPERTIES: *** NONE FOUND ***"
        }
    }
    
    $propReport | Set-Content -Path $dataflowPropsPath -Encoding UTF8
    Write-Host "  [6] Dataflow properties: $dataflowPropsPath" -ForegroundColor Green
    
    # Summary: unique property names across all dataflows
    $uniqueProps = $allDfProps | Sort-Object -Unique
    Write-Host ""
    Write-Host "      ALL property names found on dataflow objects:" -ForegroundColor Cyan
    foreach ($p in $uniqueProps) {
        $highlight = if ($p -match "datasource|upstream|source|usage|linked|reference") { "Yellow" } else { "White" }
        Write-Host "        $p" -ForegroundColor $highlight
    }
    Write-Host ""
    
    # Specific check for upstreamDataflows
    $hasUpstream = $ws.dataflows | Where-Object { $_.upstreamDataflows -and $_.upstreamDataflows.Count -gt 0 }
    if ($hasUpstream) {
        Write-Host "      *** FOUND upstreamDataflows on $($hasUpstream.Count) dataflow(s)! ***" -ForegroundColor Green
        foreach ($df in $hasUpstream) {
            Write-Host "        $($df.name): $($df.upstreamDataflows.Count) upstream dataflows" -ForegroundColor Green
            foreach ($udf in $df.upstreamDataflows) {
                Write-Host "          -> targetDataflowId: $($udf.targetDataflowId), groupId: $($udf.groupId)" -ForegroundColor Green
            }
        }
    }
    else {
        Write-Host "      *** upstreamDataflows NOT found on any dataflow object ***" -ForegroundColor Yellow
    }
    
    # Also check datasourceUsages
    $hasDsUsages = $ws.dataflows | Where-Object { $_.datasourceUsages -and $_.datasourceUsages.Count -gt 0 }
    Write-Host ""
    Write-Host "      Dataflows with datasourceUsages: $($hasDsUsages.Count) / $($ws.dataflows.Count)" -ForegroundColor Cyan
    $noDsUsages = $ws.dataflows | Where-Object { -not $_.datasourceUsages -or $_.datasourceUsages.Count -eq 0 }
    if ($noDsUsages -and $noDsUsages.Count -gt 0) {
        Write-Host "      Dataflows WITHOUT datasourceUsages ($($noDsUsages.Count)):" -ForegroundColor Yellow
        foreach ($df in $noDsUsages) {
            Write-Host "        $($df.name) ($($df.objectId))" -ForegroundColor Yellow
        }
    }
}
else {
    "[]" | Set-Content -Path $dataflowsPath -Encoding UTF8
    "No dataflows in this workspace" | Set-Content -Path $dataflowPropsPath -Encoding UTF8
    Write-Host "  [5] Dataflows: NONE in this workspace" -ForegroundColor Yellow
}
Write-Host ""

# --- 6. Datasets (for comparison — show upstreamDataflows on datasets) ---
$datasetsPath = "${outPrefix}_datasets.json"
$datasetPropsPath = "${outPrefix}_dataset_PROPERTIES.txt"

if ($ws.datasets -and $ws.datasets.Count -gt 0) {
    $ws.datasets | ConvertTo-Json -Depth 100 | Set-Content -Path $datasetsPath -Encoding UTF8
    Write-Host "  [7] Datasets: $($ws.datasets.Count) entries -> $datasetsPath" -ForegroundColor Green
    
    $dsPropsReport = @()
    foreach ($ds in $ws.datasets) {
        $props = $ds.PSObject.Properties.Name
        $dsPropsReport += ""
        $dsPropsReport += "=== $($ds.name) (id: $($ds.id)) ==="
        $dsPropsReport += "Properties: $($props -join ', ')"
        
        if ($ds.upstreamDataflows -and $ds.upstreamDataflows.Count -gt 0) {
            $dsPropsReport += "upstreamDataflows:"
            foreach ($udf in $ds.upstreamDataflows) {
                $dsPropsReport += "  -> targetDataflowId: $($udf.targetDataflowId), groupId: $($udf.groupId)"
            }
        }
        if ($ds.datasourceUsages -and $ds.datasourceUsages.Count -gt 0) {
            $dsPropsReport += "datasourceUsages: $($ds.datasourceUsages.Count) entries"
        }
    }
    $dsPropsReport | Set-Content -Path $datasetPropsPath -Encoding UTF8
    Write-Host "  [8] Dataset properties: $datasetPropsPath" -ForegroundColor Green
    
    $dsWithUpstream = $ws.datasets | Where-Object { $_.upstreamDataflows -and $_.upstreamDataflows.Count -gt 0 }
    if ($dsWithUpstream) {
        Write-Host "      Datasets with upstreamDataflows: $($dsWithUpstream.Count)" -ForegroundColor Cyan
    }
}
else {
    "[]" | Set-Content -Path $datasetsPath -Encoding UTF8
    "No datasets in this workspace" | Set-Content -Path $datasetPropsPath -Encoding UTF8
    Write-Host "  [7] Datasets: NONE in this workspace" -ForegroundColor Yellow
}

# --- 7. Dataflow Upstream Dataflows (THE MISSING LINK!) ---
$dfUpstreamPath = "${outPrefix}_dataflowUpstreamDataflows.json"

if ($ws.dataflows -and $ws.dataflows.Count -gt 0) {
    $dfWithUpstream = $ws.dataflows | Where-Object { $_.upstreamDataflows -and $_.upstreamDataflows.Count -gt 0 }
    
    if ($dfWithUpstream -and $dfWithUpstream.Count -gt 0) {
        # Build table 12 rows
        $table12Rows = foreach ($df in $dfWithUpstream) {
            foreach ($udf in $df.upstreamDataflows) {
                # Resolve target dataflow name from the same workspace
                $targetName = ($ws.dataflows | Where-Object { $_.objectId -eq $udf.targetDataflowId } | Select-Object -First 1).name
                [PSCustomObject]@{
                    WorkspaceId             = $ws.id
                    WorkspaceName           = $ws.name
                    DataflowId              = $df.objectId
                    DataflowName            = $df.name
                    TargetDataflowId        = $udf.targetDataflowId
                    TargetDataflowName      = $targetName
                    TargetDataflowWorkspaceId = $udf.groupId
                }
            }
        }
        
        $table12Rows | ConvertTo-Json -Depth 10 | Set-Content -Path $dfUpstreamPath -Encoding UTF8
        Write-Host "  [9] Dataflow Upstream Dataflows: $($table12Rows.Count) links -> $dfUpstreamPath" -ForegroundColor Green
        Write-Host ""
        Write-Host "      *** THE MISSING LINK — dataflow-to-dataflow dependencies: ***" -ForegroundColor Green
        foreach ($row in $table12Rows) {
            Write-Host "        $($row.DataflowName) -> $($row.TargetDataflowName)" -ForegroundColor Green
        }
        
        # Also export as CSV for easy use
        $table12CsvPath = "${outPrefix}_12_DataflowUpstreamDataflows.csv"
        $table12Rows | Export-Csv -Path $table12CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "      CSV export: $table12CsvPath" -ForegroundColor Cyan
    }
    else {
        "[]" | Set-Content -Path $dfUpstreamPath -Encoding UTF8
        Write-Host "  [9] Dataflow Upstream Dataflows: no dataflows with upstreamDataflows" -ForegroundColor Yellow
    }
}
else {
    "[]" | Set-Content -Path $dfUpstreamPath -Encoding UTF8
    Write-Host "  [9] Dataflow Upstream Dataflows: no dataflows in workspace" -ForegroundColor Yellow
}

# --- 8. Full lineage chain reconstruction ---
Write-Host ""
Write-Host "  ═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  FULL LINEAGE CHAIN RECONSTRUCTION" -ForegroundColor Cyan
Write-Host "  ═══════════════════════════════════════════════" -ForegroundColor Cyan

# Build lookup: datasourceInstanceId -> datasource details
$dsLookup = @{}
if ($scanResult.datasourceInstances) {
    foreach ($dsi in $scanResult.datasourceInstances) {
        $dsLookup[$dsi.datasourceId] = $dsi
    }
}

# Build lookup: dataflowId -> dataflow object
$dfLookup = @{}
if ($ws.dataflows) {
    foreach ($df in $ws.dataflows) {
        $dfLookup[$df.objectId] = $df
    }
}

# For each dataset with upstreamDataflows, trace the full chain
if ($ws.datasets) {
    foreach ($ds in $ws.datasets) {
        if ($ds.upstreamDataflows -and $ds.upstreamDataflows.Count -gt 0) {
            Write-Host ""
            Write-Host "  Dataset: $($ds.name)" -ForegroundColor White
            
            foreach ($udf in $ds.upstreamDataflows) {
                $df1 = $dfLookup[$udf.targetDataflowId]
                $df1Name = if ($df1) { $df1.name } else { "(unknown: $($udf.targetDataflowId))" }
                
                # Does this dataflow have upstreamDataflows? (chain to another dataflow)
                if ($df1 -and $df1.upstreamDataflows -and $df1.upstreamDataflows.Count -gt 0) {
                    foreach ($udf2 in $df1.upstreamDataflows) {
                        $df2 = $dfLookup[$udf2.targetDataflowId]
                        $df2Name = if ($df2) { $df2.name } else { "(unknown: $($udf2.targetDataflowId))" }
                        
                        # Does this second-level dataflow have datasourceUsages?
                        if ($df2 -and $df2.datasourceUsages -and $df2.datasourceUsages.Count -gt 0) {
                            $sources = foreach ($du in $df2.datasourceUsages) {
                                $dsi = $dsLookup[$du.datasourceInstanceId]
                                if ($dsi) {
                                    "$($dsi.datasourceType): $($dsi.connectionDetails | ConvertTo-Json -Compress)"
                                } else {
                                    "(unknown datasource: $($du.datasourceInstanceId))"
                                }
                            }
                            Write-Host "    -> $df1Name -> $df2Name -> $($sources -join ' | ')" -ForegroundColor Green
                        }
                        else {
                            Write-Host "    -> $df1Name -> $df2Name -> (no datasourceUsages)" -ForegroundColor Yellow
                        }
                    }
                }
                # Does this dataflow have datasourceUsages directly?
                elseif ($df1 -and $df1.datasourceUsages -and $df1.datasourceUsages.Count -gt 0) {
                    $sources = foreach ($du in $df1.datasourceUsages) {
                        $dsi = $dsLookup[$du.datasourceInstanceId]
                        if ($dsi) {
                            "$($dsi.datasourceType): $($dsi.connectionDetails | ConvertTo-Json -Compress)"
                        } else {
                            "(unknown datasource: $($du.datasourceInstanceId))"
                        }
                    }
                    Write-Host "    -> $df1Name -> $($sources -join ' | ')" -ForegroundColor Green
                }
                else {
                    Write-Host "    -> $df1Name -> (no datasourceUsages, no upstreamDataflows)" -ForegroundColor Yellow
                }
            }
        }
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Diagnostic complete!" -ForegroundColor Green
Write-Host "  Output files in: $OutputPath" -ForegroundColor Cyan
Write-Host "  Prefix: ScanResult_${timestamp}_*" -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan
Write-Host "  KEY QUESTION: Does 'upstreamDataflows' appear on" -ForegroundColor Yellow
Write-Host "  any dataflow object? Check:" -ForegroundColor Yellow
Write-Host "    $dataflowPropsPath" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
