<#
PREREQUISITES
-------------
- SharePoint Online Management Shell installed
- SharePoint Administrator or Global Administrator permissions
- Ability to authenticate to the SharePoint Online Admin endpoint

Official documentation:
https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
#>

param (
    [switch]$WhatIfOnly
)

# =========================
# CONFIGURATION
# =========================

$AdminUrl          = "https://contoso-admin.sharepoint.com"
$DeleteBeforeDays  = 365
$MajorVersionLimit = 100   # Desired explicit version limit

$Results = @()

# =========================
# CONNECT TO SPO
# =========================

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl

# =========================
# READ CURRENT TENANT STATE
# =========================

Write-Host "Reading current tenant configuration..." -ForegroundColor Cyan
$TenantConfig = Get-SPOTenant

$AutoExpirationEnabled = [bool]$TenantConfig.EnableAutoExpirationVersionTrim

if ($AutoExpirationEnabled) {
    Write-Host "WARNING: Auto-expiration is ENABLED in this tenant." -ForegroundColor Yellow
    Write-Host "Explicit version limits (MajorVersionLimit) cannot be applied while this is enabled." -ForegroundColor Yellow
    Write-Host "The script will NOT attempt to change this setting." -ForegroundColor Yellow
}
else {
    Write-Host "Auto-expiration is disabled. Explicit version limits are under your control." -ForegroundColor Green
}

# =========================
# BUILD TENANT CONFIG (SAFE)
# =========================

$SetTenantParams = @{}
$ConfigDiff = @()

# Only manage MajorVersionLimit if AutoExpiration is disabled
if (-not $AutoExpirationEnabled) {

    if ($TenantConfig.MajorVersionLimit -ne $MajorVersionLimit) {
        $SetTenantParams.MajorVersionLimit = $MajorVersionLimit
        $ConfigDiff += [PSCustomObject]@{
            Setting      = "MajorVersionLimit"
            CurrentValue = $TenantConfig.MajorVersionLimit
            DesiredValue = $MajorVersionLimit
        }
    }

}
else {
    Write-Host "Skipping MajorVersionLimit configuration due to Auto-expiration." -ForegroundColor DarkGray
}

# =========================
# PRESENT DIFF
# =========================

if ($ConfigDiff.Count -eq 0) {
    Write-Host "Tenant versioning configuration already compliant." -ForegroundColor Green
}
else {
    Write-Host "Tenant versioning configuration differs:" -ForegroundColor Yellow
    $ConfigDiff | Format-Table -AutoSize
}

# =========================
# APPLY TENANT CONFIG
# =========================

if ($SetTenantParams.Count -gt 0) {

    if ($WhatIfOnly) {
        Write-Host "WHATIF: Would apply tenant settings:" -ForegroundColor Yellow
        $SetTenantParams | Format-List
    }
    else {
        Write-Host "Applying tenant configuration..." -ForegroundColor Cyan
        try {
            Set-SPOTenant @SetTenantParams
            Write-Host "Tenant configuration updated successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to update tenant configuration: $($_.Exception.Message)"
            throw
        }
    }
}

# =========================
# GET ALL SITES
# =========================

Write-Host "Fetching all SharePoint sites..." -ForegroundColor Cyan

$Sites = Get-SPOSite -Limit All | Where-Object {
    $_.Template -ne "SPSPERS" -and $_.Url -notlike "*-my.sharepoint.com*"
}

Write-Host "Found $($Sites.Count) sites." -ForegroundColor Green

# =========================
# SUBMIT CLEANUP JOBS
# =========================

foreach ($Site in $Sites) {

    Write-Host "Processing site: $($Site.Url)" -ForegroundColor Yellow

    if ($WhatIfOnly) {
        $Results += [PSCustomObject]@{
            SiteUrl   = $Site.Url
            Mode      = "WhatIf"
            Action    = "No cleanup job submitted"
            Timestamp = (Get-Date).ToUniversalTime()
        }
        continue
    }

    try {
        $Job = New-SPOSiteFileVersionBatchDeleteJob `
            -Identity $Site.Url `
            -DeleteBeforeDays $DeleteBeforeDays

        $Results += [PSCustomObject]@{
            SiteUrl    = $Site.Url
            Mode       = "Cleanup"
            Status     = "JobCreated"
            WorkItemId = $Job.WorkItemId
            Timestamp  = (Get-Date).ToUniversalTime()
        }
    }
    catch {
        $Results += [PSCustomObject]@{
            SiteUrl   = $Site.Url
            Mode      = "Cleanup"
            Status    = "Failed"
            Error     = $_.Exception.Message
            Timestamp = (Get-Date).ToUniversalTime()
        }

        Write-Warning "Failed for $($Site.Url): $($_.Exception.Message)"
    }
}

# =========================
# OUTPUT RESULTS
# =========================

$OutFile = if ($WhatIfOnly) {
    ".\SharePoint-Versioning-WhatIf.csv"
}
else {
    ".\SharePoint-Versioning-CleanupJobs.csv"
}

$Results | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8

Write-Host "Results exported to $OutFile" -ForegroundColor Green

if ($WhatIfOnly) {
    Write-Host "WHATIF mode complete. No changes were made." -ForegroundColor Yellow
}
else {
    Write-Host "All cleanup jobs submitted." -ForegroundColor Green
}
