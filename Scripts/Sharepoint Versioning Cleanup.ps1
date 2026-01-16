<#
PREREQUISITES
-------------
- SharePoint Online Management Shell installed
- SharePoint Administrator or Global Administrator permissions
- Ability to authenticate to the SharePoint Online Admin endpoint

Official documentation:
https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
#>

# =========================
# CONFIGURATION
# =========================

$AdminUrl            = "https://contoso-admin.sharepoint.com"
$DeleteBeforeDays    = 365
$ExpireAfterDays     = 365
$MajorVersionLimit   = 100

$Results = @()

# =========================
# CONNECT TO SPO
# =========================

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl

# =========================
# SET TENANT DEFAULTS
# =========================

Write-Host "Configuring tenant-wide version settings..." -ForegroundColor Cyan

Set-SPOTenant `
    -EnableAutoExpirationVersionTrim $true `
    -ExpireVersionsAfterDays $ExpireAfterDays `
    -MajorVersionLimit $MajorVersionLimit

Write-Host "Tenant settings applied:" -ForegroundColor Green
Get-SPOTenant | Select-Object `
    EnableAutoExpirationVersionTrim,
    ExpireVersionsAfterDays,
    MajorVersionLimit

# =========================
# GET ALL SITES
# =========================

Write-Host "Fetching all SharePoint sites..." -ForegroundColor Cyan

$Sites = Get-SPOSite -Limit All | Where-Object {
    $_.Template -ne "SPSPERS" -and $_.Url -notlike "*-my.sharepoint.com*"
}

Write-Host "Found $($Sites.Count) sites" -ForegroundColor Green

# =========================
# RUN CLEANUP JOB PER SITE
# =========================

foreach ($Site in $Sites) {
    Write-Host "Starting cleanup job for:" $Site.Url -ForegroundColor Yellow

    try {
        $Job = New-SPOSiteFileVersionBatchDeleteJob `
            -Identity $Site.Url `
            -DeleteBeforeDays $DeleteBeforeDays

        $Results += [PSCustomObject]@{
            SiteUrl    = $Site.Url
            Status     = "JobCreated"
            WorkItemId = $Job.WorkItemId
            Timestamp  = (Get-Date).ToUniversalTime()
        }
    }
    catch {
        $Results += [PSCustomObject]@{
            SiteUrl   = $Site.Url
            Status    = "Failed"
            Error     = $_.Exception.Message
            Timestamp = (Get-Date).ToUniversalTime()
        }

        Write-Warning "Failed for $($Site.Url): $($_.Exception.Message)"
    }
}

Write-Host "All cleanup jobs submitted." -ForegroundColor Green
