<#
.SYNOPSIS
    Retrieves SharePoint Online file version batch delete job progress
    for all sites in the tenant.

.DESCRIPTION
    Uses Get-SPOSiteFileVersionBatchDeleteJobProgress, which returns
    PowerShell objects (not JSON), scoped per site.

    This script consolidates all job progress into a single CSV.

.PREREQUISITES
    - SharePoint Online Management Shell
    - SharePoint Administrator or Global Administrator
#>

# =========================
# CONFIGURATION
# =========================

$AdminUrl  = "https://contoso-admin.sharepoint.com"
$OutputCsv = ".\SharePoint-VersionCleanup-JobProgress.csv"

$Results = @()

# =========================
# CONNECT TO SPO
# =========================

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl

# =========================
# GET ALL SITES
# =========================

Write-Host "Fetching all SharePoint sites..." -ForegroundColor Cyan

$Sites = Get-SPOSite -Limit All | Where-Object {
    $_.Template -ne "SPSPERS" -and $_.Url -notlike "*-my.sharepoint.com*"
}

Write-Host "Found $($Sites.Count) sites." -ForegroundColor Green

# =========================
# QUERY JOB PROGRESS PER SITE
# =========================

foreach ($Site in $Sites) {

    Write-Host "Checking batch delete jobs for:" $Site.Url -ForegroundColor Yellow

    try {
        $Jobs = Get-SPOSiteFileVersionBatchDeleteJobProgress `
            -Identity $Site.Url

        if (-not $Jobs) {
            continue
        }

        foreach ($Job in $Jobs) {

            $Results += [PSCustomObject]@{
                SiteUrl                   = $Job.Url
                WorkItemId                = $Job.WorkItemId
                Status                    = $Job.Status
                RequestTimeInUTC          = $Job.RequestTimeInUTC
                LastProcessTimeInUTC      = $Job.LastProcessTimeInUTC
                CompleteTimeInUTC         = $Job.CompleteTimeInUTC
                BatchDeleteMode           = $Job.BatchDeleteMode
                DeleteOlderThanInUTC      = $Job.DeleteOlderThanInUTC
                MajorVersionLimit         = $Job.MajorVersionLimit
                VersionsProcessed         = $Job.VersionsProcessed
                VersionsDeleted           = $Job.VersionsDeleted
                VersionsFailed            = $Job.VersionsFailed
                StorageReleasedInBytes    = $Job.StorageReleasedInBytes
                ListsProcessed            = $Job.ListsProcessed
                ListsSynced               = $Job.ListsSynced
                ListSyncFailed            = $Job.ListSyncFailed
                ErrorMessage              = $Job.ErrorMessage
            }
        }
    }
    catch {
        Write-Warning "Failed to query job progress for $($Site.Url): $($_.Exception.Message)"
    }
}

# =========================
# OUTPUT RESULTS
# =========================

if ($Results.Count -eq 0) {
    Write-Host "No batch delete jobs found in the tenant." -ForegroundColor Yellow
    return
}

$Results |
    Sort-Object Status, SiteUrl |
    Export-Csv $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Job progress exported to $OutputCsv" -ForegroundColor Green

# =========================
# OPTIONAL SUMMARY
# =========================


Write-Host ""
Write-Host "Batch delete job status summary:" -ForegroundColor Cyan
$Results |
    Group-Object Status |
    Select-Object Name, Count |
    Format-Table -AutoSize

# =========================
# TOTALS
# =========================

$TotalVersionsDeleted = ($Results |
    Measure-Object -Property VersionsDeleted -Sum).Sum

$TotalBytesReleased = ($Results |
    Measure-Object -Property StorageReleasedInBytes -Sum).Sum

$TotalGBReleased = if ($TotalBytesReleased) {
    [math]::Round($TotalBytesReleased / 1GB, 2)
}
else {
    0
}

Write-Host ""
Write-Host "================ TOTAL IMPACT ================" -ForegroundColor Cyan
Write-Host "Total versions deleted : $TotalVersionsDeleted"
Write-Host "Total storage released : $TotalGBReleased GB"
Write-Host "=============================================" -ForegroundColor Cyan
