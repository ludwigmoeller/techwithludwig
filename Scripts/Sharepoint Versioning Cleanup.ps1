<#
PREREQUISITES
-------------
- SharePoint Online Management Shell installed
- SharePoint Administrator or Global Administrator permissions
- Ability to authenticate to the SharePoint Online Admin endpoint

Official documentation:
https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online

WHATIF MODE (Version Usage Report)
----------------------------------
This script can run in -WhatIfOnly mode to generate version usage reports per site.
Microsoft Learn tutorial:
https://learn.microsoft.com/en-us/sharepoint/tutorial-generate-version-usage-report
#>

param(
    [switch]$WhatIfOnly,

    # Where to write the report INSIDE EACH SITE.
    # Must be a document library location within the same site, and the file must not already exist. :contentReference[oaicite:4]{index=4}
    [string]$ReportLibraryRelativeUrl = "Shared Documents",

    # Optional folder under the library (will NOT be created by this script).
    [string]$ReportFolderRelativeUrl  = "reports/MyReports",

    [string]$ReportFileName           = "VersionReport.csv",

    # Polling settings for report job progress
    [int]$PollIntervalSeconds         = 20,
    [int]$MaxPollMinutes              = 30
)

# =========================
# CONFIGURATION
# =========================

$AdminUrl          = "https://contoso-admin.sharepoint.com"
$DeleteBeforeDays  = 365
$ExpireAfterDays   = 365
$MajorVersionLimit = 100

$Results = @()

# =========================
# CONNECT TO SPO
# =========================

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl

# =========================
# SET TENANT DEFAULTS (CLEANUP MODE ONLY)
# =========================

if ($WhatIfOnly) {
    Write-Host "WHATIF MODE: Tenant settings will NOT be modified, and no cleanup jobs will be created." -ForegroundColor Yellow
    Write-Host "WHATIF MODE: This will generate Version Usage Reports per site." -ForegroundColor Yellow
}
else {
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
}

# =========================
# GET ALL SITES
# =========================

Write-Host "Fetching all SharePoint sites..." -ForegroundColor Cyan

$Sites = Get-SPOSite -Limit All | Where-Object {
    $_.Template -ne "SPSPERS" -and $_.Url -notlike "*-my.sharepoint.com*"
}

Write-Host "Found $($Sites.Count) sites" -ForegroundColor Green

# =========================
# PER-SITE ACTION
# =========================

foreach ($Site in $Sites) {
    Write-Host "Processing site:" $Site.Url -ForegroundColor Yellow

    if ($WhatIfOnly) {
        # Build a ReportUrl that is INSIDE the site and points to a document library path. :contentReference[oaicite:5]{index=5}
        $EncodedLib = $ReportLibraryRelativeUrl -replace " ", "%20"
        $ReportUrl  = "$($Site.Url)/$EncodedLib/$ReportFolderRelativeUrl/$ReportFileName"

        try {
            Write-Host "  Queuing version usage report job..." -ForegroundColor Cyan
            Write-Host "  ReportUrl: $ReportUrl" -ForegroundColor DarkGray

            # Generate site-scoped version usage report :contentReference[oaicite:6]{index=6}
            $null = New-SPOSiteFileVersionExpirationReportJob `
                -Identity $Site.Url `
                -ReportUrl $ReportUrl

            $StartTime = Get-Date
            $Deadline  = $StartTime.AddMinutes($MaxPollMinutes)

            $LastStatus = $null
            $ProgressRaw = $null

            do {
                Start-Sleep -Seconds $PollIntervalSeconds

                # Track report generation progress :contentReference[oaicite:7]{index=7}
                $ProgressRaw = Get-SPOSiteFileVersionExpirationReportJobProgress `
                    -Identity $Site.Url `
                    -ReportUrl $ReportUrl

                # Progress cmdlet returns JSON text; read the "status" value (completed / in_progress / failed / no_report_found). :contentReference[oaicite:8]{index=8}
                $ProgressObj = $ProgressRaw | ConvertFrom-Json
                $LastStatus  = $ProgressObj.status

                Write-Host "  Status: $LastStatus" -ForegroundColor DarkGray

                if ((Get-Date) -ge $Deadline) {
                    throw "Timed out waiting for report generation after $MaxPollMinutes minute(s). Last status: $LastStatus"
                }

            } while ($LastStatus -eq "in_progress")

            if ($LastStatus -eq "completed") {
                $Results += [PSCustomObject]@{
                    SiteUrl    = $Site.Url
                    Mode       = "WhatIf"
                    Status     = "ReportCompleted"
                    ReportUrl  = $ReportUrl
                    Timestamp  = (Get-Date).ToUniversalTime()
                }

                Write-Host "  Report completed." -ForegroundColor Green
            }
            elseif ($LastStatus -eq "failed") {
                $ErrMsg = $ProgressObj.error_message
                $Results += [PSCustomObject]@{
                    SiteUrl    = $Site.Url
                    Mode       = "WhatIf"
                    Status     = "ReportFailed"
                    ReportUrl  = $ReportUrl
                    Error      = $ErrMsg
                    Timestamp  = (Get-Date).ToUniversalTime()
                }

                Write-Warning "  Report failed: $ErrMsg"
            }
            else {
                $Results += [PSCustomObject]@{
                    SiteUrl    = $Site.Url
                    Mode       = "WhatIf"
                    Status     = "ReportNotCompleted"
                    ReportUrl  = $ReportUrl
                    Detail     = $LastStatus
                    Timestamp  = (Get-Date).ToUniversalTime()
                }

                Write-Warning "  Report not completed. Status: $LastStatus"
            }
        }
        catch {
            $Results += [PSCustomObject]@{
                SiteUrl   = $Site.Url
                Mode      = "WhatIf"
                Status    = "Failed"
                ReportUrl = $ReportUrl
                Error     = $_.Exception.Message
                Timestamp = (Get-Date).ToUniversalTime()
            }

            Write-Warning "  WhatIf failed for $($Site.Url): $($_.Exception.Message)"
        }
    }
    else {
        # CLEANUP MODE (original behavior)
        Write-Host "Starting cleanup job for:" $Site.Url -ForegroundColor Yellow

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
}

# =========================
# OUTPUT
# =========================

$OutFile = if ($WhatIfOnly) { ".\SharePoint-Versioning-WhatIf-Reports.csv" } else { ".\SharePoint-Versioning-Cleanup-Jobs.csv" }

$Results | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Results exported to $OutFile" -ForegroundColor Green

if ($WhatIfOnly) {
    Write-Host ""
    Write-Host "Next step:" -ForegroundColor Cyan
    Write-Host "1) Open each ReportUrl in the browser and download the CSV from the document library." -ForegroundColor Gray
    Write-Host "2) Analyze it using Microsoft's Excel template or PowerShell analysis script from the tutorial." -ForegroundColor Gray
    Write-Host "Tutorial: https://learn.microsoft.com/en-us/sharepoint/tutorial-generate-version-usage-report" -ForegroundColor Gray
}
else {
    Write-Host "All cleanup jobs submitted." -ForegroundColor Green
}
