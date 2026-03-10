<#
.SYNOPSIS
    Power BI Authentication Test Script
    Tests connection to Power BI Service with MFA-enabled admin account.

.DESCRIPTION
    This script:
    - Sets execution policy for current user (no admin rights needed)
    - Installs/updates MicrosoftPowerBIMgmt module (user scope)
    - Connects to Power BI Service (interactive login with MFA)
    - Verifies connection by retrieving basic info
    - Disconnects cleanly

.NOTES
    Author: Tom Martens & Claude
    Date:   2026-02-19
    
    No elevated rights required.
    MFA prompt will appear in browser.
#>

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Power BI Authentication Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

#region Execution Policy
Write-Host "[1/4] Checking execution policy..." -ForegroundColor Yellow

$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "Undefined") {
    Write-Host "  Setting execution policy to RemoteSigned for CurrentUser..." -ForegroundColor Gray
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "  Done!" -ForegroundColor Green
}
else {
    Write-Host "  Execution policy OK: $currentPolicy" -ForegroundColor Green
}
Write-Host ""
#endregion

#region Module Installation
Write-Host "[2/4] Checking MicrosoftPowerBIMgmt module..." -ForegroundColor Yellow

$moduleName = "MicrosoftPowerBIMgmt"
$installedModule = Get-Module -ListAvailable -Name $moduleName | 
    Sort-Object Version -Descending | 
    Select-Object -First 1

if (-not $installedModule) {
    Write-Host "  Installing $moduleName module (user scope)..." -ForegroundColor Gray
    Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
    Write-Host "  Installed!" -ForegroundColor Green
}
else {
    Write-Host "  Module found: v$($installedModule.Version)" -ForegroundColor Green
    
    # Optional: check for updates
    try {
        $onlineModule = Find-Module -Name $moduleName -ErrorAction Stop
        if ($onlineModule.Version -gt $installedModule.Version) {
            Write-Host "  Update available: v$($onlineModule.Version)" -ForegroundColor Yellow
            Write-Host "  Updating..." -ForegroundColor Gray
            Update-Module -Name $moduleName -Scope CurrentUser -Force
            Write-Host "  Updated!" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Could not check for updates (network?)" -ForegroundColor Gray
    }
}

Import-Module $moduleName
Write-Host ""
#endregion

#region Authentication
Write-Host "[3/4] Connecting to Power BI Service..." -ForegroundColor Yellow
Write-Host "  A browser window will open for sign-in." -ForegroundColor Gray
Write-Host "  Complete MFA when prompted." -ForegroundColor Gray
Write-Host ""

try {
    $connection = Connect-PowerBIServiceAccount
    Write-Host "  Connected successfully!" -ForegroundColor Green
    Write-Host "  User:        $($connection.UserName)" -ForegroundColor Gray
    Write-Host "  Environment: $($connection.Environment)" -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Host "  Connection failed!" -ForegroundColor Red
    Write-Host "  Error: $_" -ForegroundColor Red
    exit 1
}
#endregion

#region Verification
Write-Host "[4/4] Verifying admin access..." -ForegroundColor Yellow

try {
    # Try an admin-only API call
    $workspaces = Get-PowerBIWorkspace -Scope Organization -First 3
    Write-Host "  Admin access confirmed!" -ForegroundColor Green
    Write-Host "  Retrieved $($workspaces.Count) workspaces (sample)" -ForegroundColor Gray
}
catch {
    Write-Host "  Warning: Could not access Organization scope." -ForegroundColor Yellow
    Write-Host "  You may not have Power BI Admin rights." -ForegroundColor Yellow
    Write-Host "  Error: $_" -ForegroundColor Gray
}
Write-Host ""
#endregion

#region Disconnect
Write-Host "Disconnecting..." -ForegroundColor Gray
Disconnect-PowerBIServiceAccount | Out-Null
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Test Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
#endregion
