<#
.SYNOPSIS
    Audits Exchange Online mailboxes, reporting detailed storage usage, In-Place Archive status,
    and Auto-Expanding Archive telemetry grouped by mailbox type (User Inbox vs Shared Mailbox).

.DESCRIPTION
    Production-grade automation script for Microsoft 365 Exchange Online administrators.
    Connects to Exchange Online using modern authentication (or reuses active sessions)
    and collects comprehensive storage metrics for user inboxes and shared mailboxes:
    - Primary mailbox total size, byte volume, and item count
    - Primary deleted items (dumpster / recoverable items) volume
    - In-Place Online Archive activation status and state
    - Auto-Expanding Archive enablement (auto-archiwizacja)
    - Live Archive storage consumption, byte volume, and item count
    - Mailbox storage quotas (ProhibitSendReceiveQuota, ArchiveQuota)
    - Combined total footprint (Primary + Archive)
    - Grouped tabular output separating User Inboxes and Shared Mailboxes
    - Executive tenant storage summary with adoption percentages and Top 5 consumers
    - CSV export (-ExportCsv) and pipeline pass-through (-PassThru)
    - Cross-language administrator privilege elevation (SID-based, PL/EN compatible)

.PARAMETER All
    Switch parameter. Audits all eligible mailboxes (both User and Shared mailboxes).

.PARAMETER MailboxType
    Filters mailboxes by type. Valid options: 'All', 'User', 'Shared'.
    Default: 'All'.

.PARAMETER Identity
    The identity (UPN, primary email, or alias) of a single target mailbox to audit.
    Aliases: -UserPrincipalName, -Email, -User, -Mailbox.

.PARAMETER Search
    Optional string filter for matching DisplayName, PrimarySmtpAddress, or UserPrincipalName.
    Alias: -Filter.

.PARAMETER ArchiveOnly
    Switch parameter. Limits the report to mailboxes that currently have an active In-Place Archive.

.PARAMETER AutoExpandingOnly
    Switch parameter. Limits the report to mailboxes that have Auto-Expanding Archive enabled.

.PARAMETER Details
    Switch parameter. Displays extended technical columns including quotas, recoverable items size,
    and retention policies.
    Aliases: -d, -Detailed.

.PARAMETER NoGrouping
    Switch parameter. Outputs a flat unified table instead of grouping into User Inboxes and Shared Mailboxes.

.PARAMETER SortBy
    Specifies the attribute used to sort results.
    Valid options: 'Size', 'ArchiveSize', 'DisplayName', 'ItemCount'.
    Default: 'Size' (descending by consumed bytes).

.PARAMETER Ascending
    Switch parameter. Changes sorting direction from descending to ascending.

.PARAMETER ExportCsv
    Optional file path to export the collected telemetry into a UTF-8 CSV report.
    Aliases: -CsvPath, -Export.

.PARAMETER PassThru
    Switch parameter. Returns custom PowerShell objects to the pipeline for downstream commands or Out-GridView.

.PARAMETER ResultSize
    Maximum number of mailboxes to retrieve from Exchange Online.
    Default: 'Unlimited'.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used for explicit authentication with Exchange Online.
    Alias: -AdminUPN.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -All
    Audits all User and Shared mailboxes, displaying grouped tables and an executive storage summary.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -MailboxType Shared
    Audits all Shared Mailboxes only, reporting storage usage and archive status.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -Identity "user@company.com" -Details
    Inspects storage, quotas, and archive telemetry for a specific mailbox in detailed mode.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -ArchiveOnly
    Lists all mailboxes that have In-Place Archive activated.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -AutoExpandingOnly
    Lists all mailboxes that have Auto-Expanding Archive enabled.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -ExportCsv "C:\Reports\MailboxStorage.csv"
    Audits all mailboxes and exports complete technical metrics to a CSV file.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -NoGrouping -PassThru | Out-GridView
    Emits custom objects to the pipeline for interactive inspection in Out-GridView.

.EXAMPLE
    .\Get-MailboxSizeAndArchiveDetails.ps1 -Help
    Displays comprehensive help, usage syntax, and parameter documentation.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName = 'Default')]
    [switch]$All,

    [Parameter(ParameterSetName = 'Default')]
    [ValidateSet('All', 'User', 'Shared')]
    [string]$MailboxType = 'All',

    [Parameter(ParameterSetName = 'Default', Position = 0)]
    [Alias('UserPrincipalName', 'Email', 'User', 'Mailbox')]
    [string]$Identity,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('Filter')]
    [string]$Search,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$ArchiveOnly,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$AutoExpandingOnly,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('d', 'Detailed')]
    [switch]$Details,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$NoGrouping,

    [Parameter(ParameterSetName = 'Default')]
    [ValidateSet('Size', 'ArchiveSize', 'DisplayName', 'ItemCount')]
    [string]$SortBy = 'Size',

    [Parameter(ParameterSetName = 'Default')]
    [switch]$Ascending,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('CsvPath', 'Export')]
    [string]$ExportCsv,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$PassThru,

    [Parameter(ParameterSetName = 'Default')]
    [string]$ResultSize = 'Unlimited',

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------
# Script Metadata
# ----------------------------------------------------------------------
$ScriptMetadata = @{
    Name        = "Get-MailboxSizeAndArchiveDetails.ps1"
    Version     = "1.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Description = "Exchange Online Mailbox Storage, In-Place Archive & Auto-Expanding Archive Auditor"
}

# ----------------------------------------------------------------------
# Help Function
# ----------------------------------------------------------------------
function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: $($ScriptMetadata.Name)",
        "VERSION: $($ScriptMetadata.Version)",
        "AUTHOR: $($ScriptMetadata.Author)",
        "CONTACT: $($ScriptMetadata.Email) | $($ScriptMetadata.GitHub)",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Audits Exchange Online mailboxes with breakdown by type (User Inbox vs Shared),",
        "    reporting primary storage consumed, In-Place Archive status, Auto-Expanding Archive",
        "    enablement (auto-archiwizacja), live archive space used, and quotas.",
        "    Includes an Executive Storage Summary and optional CSV export.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    - PowerShell 5.1 or PowerShell 7.x+",
        "    - Module: ExchangeOnlineManagement (auto-detected or installed if missing)",
        "    - Permissions: Exchange Administrator or Global Administrator",
        "    - Reuses existing authenticated sessions or prompts for interactive login",
        "",
        "USAGE EXAMPLES:",
        "    # 1. Audit all mailboxes (User Inboxes & Shared Mailboxes):",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -All",
        "",
        "    # 2. Audit only Shared Mailboxes:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -MailboxType Shared",
        "",
        "    # 3. Audit only User Inboxes:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -MailboxType User",
        "",
        "    # 4. Inspect a single target mailbox with extended quota details:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -Identity `"john.doe@contoso.com`" -Details",
        "",
        "    # 5. Search mailboxes matching a keyword (DisplayName / Email / UPN):",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -Search `"sales`"",
        "",
        "    # 6. List only mailboxes with In-Place Archive activated:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -ArchiveOnly",
        "",
        "    # 7. List only mailboxes with Auto-Expanding Archive enabled:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -AutoExpandingOnly",
        "",
        "    # 8. Export full audit report to a UTF-8 CSV file:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -ExportCsv `"C:\Reports\MailboxStorage.csv`"",
        "",
        "    # 9. Flat view without grouping, outputting objects to Out-GridView:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -All -NoGrouping -PassThru | Out-GridView",
        "",
        "    # 10. Display this help screen:",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -Help",
        "    .\Get-MailboxSizeAndArchiveDetails.ps1 -h",
        "",
        "PARAMETERS:",
        "    -All                    Audit all eligible mailboxes (User and Shared).",
        "    -MailboxType            Filter by type: 'All' (default), 'User', or 'Shared'.",
        "    -Identity               Target a single mailbox (UPN, Primary SMTP, or Alias).",
        "    -Search, -Filter        Search filter by DisplayName, Email, or UPN.",
        "    -ArchiveOnly            Limit results to mailboxes with active Online Archive.",
        "    -AutoExpandingOnly      Limit results to mailboxes with Auto-Expanding Archive enabled.",
        "    -Details, -d            Display extended technical columns (quotas, dumpster, etc.).",
        "    -NoGrouping             Render a single flat table instead of grouped sections.",
        "    -SortBy                 Sort by 'Size', 'ArchiveSize', 'DisplayName', or 'ItemCount'.",
        "    -Ascending              Sort in ascending order (default is descending).",
        "    -ExportCsv, -Export     Path to save the audit report as UTF-8 CSV.",
        "    -PassThru               Emit custom PSObjects to the pipeline.",
        "    -ResultSize             Maximum mailboxes to query (default: 'Unlimited').",
        "    -AdminUserPrincipalName Admin identity for Exchange Online connection.",
        "    -Help, -h               Display this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Get-MailboxSizeAndArchiveDetails.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# ----------------------------------------------------------------------
# Invocation without parameters or with -Help
# ----------------------------------------------------------------------
if ($Help.IsPresent -or $PSBoundParameters.Count -eq 0) {
    Show-ScriptHelp
    exit 0
}

# ----------------------------------------------------------------------
# Input Validation & Security Sanitization
# ----------------------------------------------------------------------
$forbiddenCharsRegex = '[\$`;&|><\r\n]'

if (-not [string]::IsNullOrWhiteSpace($Identity)) {
    $Identity = $Identity.Trim()
    if ($Identity -match $forbiddenCharsRegex) {
        Write-Error "Invalid characters detected in -Identity. Disallowed characters: $;&|><`$`r`n"
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    $Search = $Search.Trim()
    if ($Search -match $forbiddenCharsRegex) {
        Write-Error "Invalid characters detected in -Search. Disallowed characters: $;&|><`$`r`n"
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
    $AdminUserPrincipalName = $AdminUserPrincipalName.Trim()
    $emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    if ($AdminUserPrincipalName -notmatch $emailPattern) {
        Write-Error "Invalid format for -AdminUserPrincipalName. Provide a valid email address or UPN."
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($ExportCsv)) {
    $ExportCsv = $ExportCsv.Trim().Trim('"').Trim("'")
    if ($ExportCsv -match $forbiddenCharsRegex) {
        Write-Error "Invalid characters detected in -ExportCsv path."
        exit 1
    }
    if ($ExportCsv -notmatch '\.csv$') {
        Write-Error "-ExportCsv must specify a valid file path ending in .csv."
        exit 1
    }
}

if ($ResultSize -ne 'Unlimited') {
    $parsedResultSize = 0
    if (-not [int]::TryParse($ResultSize, [ref]$parsedResultSize) -or $parsedResultSize -le 0) {
        Write-Error "Invalid -ResultSize. Must be 'Unlimited' or a positive integer."
        exit 1
    }
}

# ----------------------------------------------------------------------
# Language-Independent Administrator Privilege Verification (SID-Based)
# ----------------------------------------------------------------------
$identityPrincipal = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identityPrincipal)
$adminSid = New-Object Security.Principal.SecurityIdentifier([Security.Principal.WellKnownSidType]::BuiltInAdministratorsSid, $null)

$isAdmin = $principal.IsInRole($adminSid)

if (-not $isAdmin) {
    Write-Host "Missing Administrator privileges. Requesting elevation..." -ForegroundColor Yellow
    
    $argumentsList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
    foreach ($param in $PSBoundParameters.GetEnumerator()) {
        if ($param.Value -is [switch]) {
            if ($param.Value.IsPresent) {
                $argumentsList += "-$($param.Key)"
            }
        } else {
            $argumentsList += "-$($param.Key)"
            $argumentsList += "`"$($param.Value)`""
        }
    }

    Start-Process powershell.exe -ArgumentList ($argumentsList -join ' ') -Verb RunAs
    exit 0
}

# ----------------------------------------------------------------------
# Helper Functions: Safe Property Resolution & Byte Size Formatting
# ----------------------------------------------------------------------
function Get-SafePropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = $null
    )
    if ($null -eq $InputObject) { return $DefaultValue }
    $prop = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $prop -and $null -ne $prop.Value) {
        return $prop.Value
    }
    return $DefaultValue
}

function Format-ByteNumber {
    param([int64]$Bytes)
    if ($Bytes -ge 1TB) {
        return "{0:N2} TB" -f ($Bytes / 1TB)
    } elseif ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    } elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    } elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    } else {
        return "{0} B" -f $Bytes
    }
}

function Parse-ByteSize {
    param([object]$SizeValue)

    if ($null -eq $SizeValue) {
        return [PSCustomObject]@{
            Bytes     = [int64]0
            MB        = 0.0
            GB        = 0.0
            Formatted = "0.00 MB"
        }
    }

    $rawString = if ($SizeValue.PSObject.Properties['Value']) {
        $SizeValue.Value.ToString()
    } else {
        $SizeValue.ToString()
    }

    $bytes = [int64]0

    if ($rawString -match '\(([\d,\.\s]+)\s*bytes\)') {
        $cleanBytes = $matches[1] -replace '[\s,\.]', ''
        [int64]::TryParse($cleanBytes, [ref]$bytes) | Out-Null
    } elseif ($SizeValue.PSObject.Methods['ToBytes']) {
        try {
            $bytes = [int64]$SizeValue.ToBytes()
        } catch {
            $bytes = [int64]0
        }
    } elseif ($rawString -match '^\d+$') {
        [int64]::TryParse($rawString, [ref]$bytes) | Out-Null
    }

    $mb = [math]::Round(($bytes / 1MB), 2)
    $gb = [math]::Round(($bytes / 1GB), 2)
    $formatted = Format-ByteNumber $bytes

    return [PSCustomObject]@{
        Bytes     = $bytes
        MB        = $mb
        GB        = $gb
        Formatted = $formatted
    }
}

function Format-QuotaValue {
    param([object]$QuotaVal)
    if ($null -eq $QuotaVal) { return '-' }
    $str = if ($QuotaVal.PSObject.Properties['Value']) { $QuotaVal.Value.ToString() } else { $QuotaVal.ToString() }
    if ($str -eq 'Unlimited') { return 'Unlimited' }
    if ($str -match '\(([\d,\.\s]+)\s*bytes\)') {
        $clean = $matches[1] -replace '[\s,\.]', ''
        $b = [int64]0
        if ([int64]::TryParse($clean, [ref]$b)) {
            return Format-ByteNumber $b
        }
    }
    return $str
}

# ----------------------------------------------------------------------
# Microsoft 365 Exchange Online Authentication
# ----------------------------------------------------------------------
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "[*] ExchangeOnlineManagement module not detected. Installing..." -ForegroundColor Cyan
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] Successfully installed ExchangeOnlineManagement module." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install ExchangeOnlineManagement module: $_`nPlease execute: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
        exit 1
    }
}

try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[+] Active Exchange Online session detected." -ForegroundColor Green
} catch {
    Write-Host "[*] No active Exchange Online session detected. Initiating administrator login..." -ForegroundColor Cyan
    try {
        $connectParams = @{
            ShowBanner  = $false
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
            $connectParams['UserPrincipalName'] = $AdminUserPrincipalName
        }
        Connect-ExchangeOnline @connectParams
        Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

# ----------------------------------------------------------------------
# Target Mailbox Discovery & Query
# ----------------------------------------------------------------------
$mailboxes = @()

if (-not [string]::IsNullOrWhiteSpace($Identity)) {
    Write-Host "`n[*] Querying target mailbox: '$Identity'..." -ForegroundColor Cyan
    try {
        $singleMailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        if ($singleMailbox.RecipientTypeDetails -notin @('UserMailbox', 'SharedMailbox')) {
            Write-Error "Mailbox '$Identity' is of type '$($singleMailbox.RecipientTypeDetails)'. Only UserMailbox and SharedMailbox are supported."
            exit 1
        }
        $mailboxes = @($singleMailbox)
    } catch {
        Write-Error "Unable to locate mailbox '$Identity'. Details: $_"
        exit 1
    }
} else {
    $recipientTypes = switch ($MailboxType) {
        'User'   { @('UserMailbox') }
        'Shared' { @('SharedMailbox') }
        default  { @('UserMailbox', 'SharedMailbox') }
    }

    Write-Host "`n[*] Retrieving mailboxes from Exchange Online (Scope: $($recipientTypes -join ', '))..." -ForegroundColor Cyan
    $queryParam = @{
        RecipientTypeDetails = $recipientTypes
        ErrorAction          = 'Stop'
    }
    if ($ResultSize -eq 'Unlimited') {
        $queryParam['ResultSize'] = 'Unlimited'
    } else {
        $queryParam['ResultSize'] = [int]$ResultSize
    }

    try {
        $mailboxes = @(Get-Mailbox @queryParam)
        Write-Host "    Retrieved $($mailboxes.Count) candidate mailbox(es)." -ForegroundColor Gray
    } catch {
        Write-Error "Failed to retrieve mailboxes: $_"
        exit 1
    }
}

# Apply partial keyword search filter if provided
if (-not [string]::IsNullOrWhiteSpace($Search)) {
    Write-Host "[*] Applying search filter: '$Search'..." -ForegroundColor Cyan
    $mailboxes = @($mailboxes | Where-Object {
        $_.DisplayName -like "*$Search*" -or
        $_.UserPrincipalName -like "*$Search*" -or
        $_.PrimarySmtpAddress -like "*$Search*" -or
        $_.Alias -like "*$Search*"
    })
    Write-Host "    $($mailboxes.Count) mailbox(es) matched search filter." -ForegroundColor Gray
}

if ($mailboxes.Count -eq 0) {
    Write-Host "`n[!] No mailboxes found matching the specified criteria." -ForegroundColor Yellow
    exit 0
}

# ----------------------------------------------------------------------
# Process Mailboxes: Collect Storage & Archive Statistics
# ----------------------------------------------------------------------
$processedList = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalMailboxes = $mailboxes.Count
$currentIndex = 0

Write-Host "`n[*] Collecting storage usage and archive statistics for $totalMailboxes mailbox(es)..." -ForegroundColor Cyan

foreach ($mb in $mailboxes) {
    $currentIndex++
    $percent = [math]::Round(($currentIndex / $totalMailboxes) * 100)
    Write-Progress -Activity "Collecting Mailbox & Archive Statistics" `
                   -Status "Processing ($currentIndex of $totalMailboxes): $($mb.DisplayName)" `
                   -PercentComplete $percent

    # 1. Primary Mailbox Statistics
    $primaryStats = $null
    try {
        $primaryStats = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction Stop
    } catch {
        try {
            $primaryStats = Get-MailboxStatistics -Identity $mb.ExchangeGuid.ToString() -ErrorAction SilentlyContinue
        } catch {
            $primaryStats = $null
        }
    }

    $primarySizeInfo = if ($primaryStats -and (Get-SafePropertyValue $primaryStats 'TotalItemSize')) {
        Parse-ByteSize $primaryStats.TotalItemSize
    } else {
        Parse-ByteSize $null
    }

    $primaryDeletedSizeInfo = if ($primaryStats -and (Get-SafePropertyValue $primaryStats 'TotalDeletedItemSize')) {
        Parse-ByteSize $primaryStats.TotalDeletedItemSize
    } else {
        Parse-ByteSize $null
    }

    $primaryItemCount = if ($primaryStats -and (Get-SafePropertyValue $primaryStats 'ItemCount')) {
        [int64]$primaryStats.ItemCount
    } else {
        [int64]0
    }

    $primaryDeletedItemCount = if ($primaryStats -and (Get-SafePropertyValue $primaryStats 'DeletedItemCount')) {
        [int64]$primaryStats.DeletedItemCount
    } else {
        [int64]0
    }

    $lastLogon = '-'
    if ($primaryStats) {
        $logonDate = Get-SafePropertyValue $primaryStats 'LastLogonTime'
        if ($null -ne $logonDate) {
            $lastLogon = $logonDate.ToString("yyyy-MM-dd HH:mm")
        }
    }

    # 2. Archive & Auto-Expanding Archive Telemetry
    $archiveStatus = if ($mb.ArchiveStatus) { [string]$mb.ArchiveStatus } else { 'None' }
    $isArchiveActive = ($archiveStatus -eq 'Active')

    $autoExpanding = $false
    $autoExpandingProp = Get-SafePropertyValue $mb 'AutoExpandingArchiveEnabled' $null
    if ($null -ne $autoExpandingProp) {
        $autoExpanding = [bool]$autoExpandingProp
    }

    # 3. Archive Statistics (if Active)
    $archiveStats = $null
    if ($isArchiveActive) {
        try {
            $archiveStats = Get-MailboxStatistics -Identity $mb.UserPrincipalName -Archive -ErrorAction Stop
        } catch {
            try {
                $archiveStats = Get-MailboxStatistics -Identity $mb.ExchangeGuid.ToString() -Archive -ErrorAction SilentlyContinue
            } catch {
                $archiveStats = $null
            }
        }
    }

    $archiveSizeInfo = if ($archiveStats -and (Get-SafePropertyValue $archiveStats 'TotalItemSize')) {
        Parse-ByteSize $archiveStats.TotalItemSize
    } elseif ($isArchiveActive) {
        # Archive is active but store is initializing / empty
        Parse-ByteSize $null
    } else {
        $null
    }

    $archiveDeletedSizeInfo = if ($archiveStats -and (Get-SafePropertyValue $archiveStats 'TotalDeletedItemSize')) {
        Parse-ByteSize $archiveStats.TotalDeletedItemSize
    } else {
        Parse-ByteSize $null
    }

    $archiveItemCount = if ($archiveStats -and (Get-SafePropertyValue $archiveStats 'ItemCount')) {
        [int64]$archiveStats.ItemCount
    } elseif ($isArchiveActive) {
        [int64]0
    } else {
        -1
    }

    $archiveDeletedItemCount = if ($archiveStats -and (Get-SafePropertyValue $archiveStats 'DeletedItemCount')) {
        [int64]$archiveStats.DeletedItemCount
    } else {
        0
    }

    # 4. Quotas
    $primaryProhibitSendReceiveQuota = Format-QuotaValue (Get-SafePropertyValue $mb 'ProhibitSendReceiveQuota')
    $primaryIssueWarningQuota        = Format-QuotaValue (Get-SafePropertyValue $mb 'IssueWarningQuota')
    $archiveQuota                    = Format-QuotaValue (Get-SafePropertyValue $mb 'ArchiveQuota')
    $archiveWarningQuota             = Format-QuotaValue (Get-SafePropertyValue $mb 'ArchiveWarningQuota')

    # 5. Combined Primary + Archive Footprint
    $archiveSizeBytes = if ($archiveSizeInfo) { $archiveSizeInfo.Bytes } else { [int64]0 }
    $combinedBytes = $primarySizeInfo.Bytes + $archiveSizeBytes
    $combinedFormatted = Format-ByteNumber $combinedBytes

    # Type friendly display
    $typeFriendly = switch ($mb.RecipientTypeDetails.ToString()) {
        'UserMailbox'   { 'User Inbox' }
        'SharedMailbox' { 'Shared Mailbox' }
        default         { $mb.RecipientTypeDetails.ToString() }
    }

    $retentionPol = Get-SafePropertyValue $mb 'RetentionPolicy' 'None'

    $record = [PSCustomObject]@{
        Nr                           = 0
        DisplayName                  = $mb.DisplayName
        UserPrincipalName            = $mb.UserPrincipalName
        PrimarySmtpAddress           = $mb.PrimarySmtpAddress
        MailboxType                  = $typeFriendly
        RecipientTypeDetails         = $mb.RecipientTypeDetails.ToString()
        PrimaryTotalItemSize         = $primarySizeInfo.Formatted
        PrimaryTotalItemSizeBytes    = $primarySizeInfo.Bytes
        PrimaryTotalItemSizeMB       = $primarySizeInfo.MB
        PrimaryTotalItemSizeGB       = $primarySizeInfo.GB
        PrimaryItemCount             = $primaryItemCount
        PrimaryDeletedItemSize       = $primaryDeletedSizeInfo.Formatted
        PrimaryDeletedItemSizeBytes  = $primaryDeletedSizeInfo.Bytes
        PrimaryDeletedItemCount      = $primaryDeletedItemCount
        PrimaryProhibitSendReceive   = $primaryProhibitSendReceiveQuota
        PrimaryIssueWarning          = $primaryIssueWarningQuota
        ArchiveStatus                = $archiveStatus
        ArchiveEnabled               = $isArchiveActive
        AutoExpandingArchive         = if ($autoExpanding) { 'Enabled' } else { 'Disabled' }
        AutoExpandingArchiveEnabled  = $autoExpanding
        ArchiveTotalItemSize         = if ($archiveSizeInfo) { $archiveSizeInfo.Formatted } else { 'N/A' }
        ArchiveTotalItemSizeBytes    = $archiveSizeBytes
        ArchiveTotalItemSizeMB       = if ($archiveSizeInfo) { $archiveSizeInfo.MB } else { 0.0 }
        ArchiveTotalItemSizeGB       = if ($archiveSizeInfo) { $archiveSizeInfo.GB } else { 0.0 }
        ArchiveItemCount             = if ($archiveItemCount -ge 0) { $archiveItemCount } else { 'N/A' }
        ArchiveDeletedItemSize       = if ($isArchiveActive) { $archiveDeletedSizeInfo.Formatted } else { 'N/A' }
        ArchiveDeletedItemCount      = if ($isArchiveActive) { $archiveDeletedItemCount } else { 'N/A' }
        ArchiveQuota                 = if ($isArchiveActive) { $archiveQuota } else { 'N/A' }
        ArchiveWarningQuota          = if ($isArchiveActive) { $archiveWarningQuota } else { 'N/A' }
        CombinedTotalSize            = $combinedFormatted
        CombinedTotalSizeBytes       = $combinedBytes
        CombinedTotalSizeGB          = [math]::Round(($combinedBytes / 1GB), 2)
        RetentionPolicy              = if ($retentionPol) { $retentionPol } else { 'None' }
        LitigationHoldEnabled        = [bool](Get-SafePropertyValue $mb 'LitigationHoldEnabled' $false)
        LastLogonTime                = $lastLogon
    }

    $processedList.Add($record)
}

Write-Progress -Activity "Collecting Mailbox & Archive Statistics" -Completed

# ----------------------------------------------------------------------
# Post-Processing Filters: -ArchiveOnly and -AutoExpandingOnly
# ----------------------------------------------------------------------
if ($ArchiveOnly.IsPresent) {
    Write-Host "[*] Filtering for mailboxes with Archive Status = Active..." -ForegroundColor Cyan
    $processedList = @($processedList | Where-Object { $_.ArchiveEnabled })
}

if ($AutoExpandingOnly.IsPresent) {
    Write-Host "[*] Filtering for mailboxes with Auto-Expanding Archive Enabled..." -ForegroundColor Cyan
    $processedList = @($processedList | Where-Object { $_.AutoExpandingArchiveEnabled })
}

if ($processedList.Count -eq 0) {
    Write-Host "`n[!] No mailboxes remained after applying archive status filters." -ForegroundColor Yellow
    exit 0
}

# ----------------------------------------------------------------------
# Sorting
# ----------------------------------------------------------------------
$sortProperty = switch ($SortBy) {
    'Size'        { 'PrimaryTotalItemSizeBytes' }
    'ArchiveSize' { 'ArchiveTotalItemSizeBytes' }
    'DisplayName' { 'DisplayName' }
    'ItemCount'   { 'PrimaryItemCount' }
    default       { 'PrimaryTotalItemSizeBytes' }
}

$shouldDescend = if ($SortBy -eq 'DisplayName') {
    $false
} else {
    (-not $Ascending.IsPresent)
}

if ($shouldDescend) {
    $processedList = @($processedList | Sort-Object -Property $sortProperty -Descending)
} else {
    $processedList = @($processedList | Sort-Object -Property $sortProperty)
}

# ----------------------------------------------------------------------
# Grouping & Row Numbering (Nr)
# ----------------------------------------------------------------------
$groups = @($processedList | Group-Object -Property MailboxType | Sort-Object Name)

$globalIndex = 1
foreach ($grp in $groups) {
    foreach ($item in $grp.Group) {
        $item.Nr = $globalIndex
        $globalIndex++
    }
}

Write-Host "`n[+] Successfully analyzed $($processedList.Count) mailbox(es).`n" -ForegroundColor Green

# ----------------------------------------------------------------------
# Console Table Rendering
# ----------------------------------------------------------------------
if ($Details.IsPresent) {
    $tableColumns = @(
        @{ Label = 'Nr';            Expression = { $_.Nr };                         Width = 4 },
        @{ Label = 'Display Name';  Expression = { $_.DisplayName };                Width = 22 },
        @{ Label = 'Email / UPN';   Expression = { $_.UserPrincipalName };          Width = 30 },
        @{ Label = 'Mbx Size';      Expression = { $_.PrimaryTotalItemSize };       Width = 12 },
        @{ Label = 'Mbx Items';     Expression = { $_.PrimaryItemCount };           Width = 9 },
        @{ Label = 'Mbx Quota';     Expression = { $_.PrimaryProhibitSendReceive }; Width = 11 },
        @{ Label = 'Archive';       Expression = { $_.ArchiveStatus };              Width = 8 },
        @{ Label = 'Auto-Expand';   Expression = { $_.AutoExpandingArchive };       Width = 11 },
        @{ Label = 'Arch. Size';    Expression = { $_.ArchiveTotalItemSize };       Width = 12 },
        @{ Label = 'Arch. Items';   Expression = { $_.ArchiveItemCount };           Width = 11 },
        @{ Label = 'Arch. Quota';   Expression = { $_.ArchiveQuota };               Width = 11 },
        @{ Label = 'Total Size';    Expression = { $_.CombinedTotalSize };          Width = 12 },
        @{ Label = 'Last Logon';    Expression = { $_.LastLogonTime };              Width = 16 }
    )
} else {
    $tableColumns = @(
        @{ Label = 'Nr';            Expression = { $_.Nr };                         Width = 4 },
        @{ Label = 'Display Name';  Expression = { $_.DisplayName };                Width = 24 },
        @{ Label = 'Email / UPN';   Expression = { $_.UserPrincipalName };          Width = 32 },
        @{ Label = 'Mailbox Size';  Expression = { $_.PrimaryTotalItemSize };       Width = 14 },
        @{ Label = 'Items';         Expression = { $_.PrimaryItemCount };           Width = 8 },
        @{ Label = 'Archive';       Expression = { $_.ArchiveStatus };              Width = 9 },
        @{ Label = 'Auto-Expand';   Expression = { $_.AutoExpandingArchive };       Width = 12 },
        @{ Label = 'Archive Size';  Expression = { $_.ArchiveTotalItemSize };       Width = 14 },
        @{ Label = 'Arch. Items';   Expression = { $_.ArchiveItemCount };           Width = 12 },
        @{ Label = 'Combined Size'; Expression = { $_.CombinedTotalSize };          Width = 14 }
    )
}

$consoleWidth = 220

if ($NoGrouping.IsPresent) {
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "  ALL MAILBOXES (Unified Flat View - $($processedList.Count) mailboxes)" -ForegroundColor White
    Write-Host "================================================================================" -ForegroundColor Cyan
    $flatItems = @($processedList | Sort-Object Nr)
    $flatItems | Format-Table -Property $tableColumns | Out-String -Width $consoleWidth | Write-Host
} else {
    foreach ($grp in $groups) {
        $groupTitle = if ($grp.Name -eq 'User Inbox') {
            "USER INBOXES (Skrzynki uzytkownikow) - $($grp.Count) mailbox(es)"
        } elseif ($grp.Name -eq 'Shared Mailbox') {
            "SHARED MAILBOXES (Skrzynki wspoldzielone) - $($grp.Count) mailbox(es)"
        } else {
            "$($grp.Name) - $($grp.Count) mailbox(es)"
        }

        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  $groupTitle" -ForegroundColor White
        Write-Host "================================================================================" -ForegroundColor Cyan
        $groupItems = @($grp.Group | Sort-Object Nr)
        $groupItems | Format-Table -Property $tableColumns | Out-String -Width $consoleWidth | Write-Host
    }
}

# ----------------------------------------------------------------------
# Executive Mailbox & Archive Storage Summary
# ----------------------------------------------------------------------
$userMbxList   = @($processedList | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' })
$sharedMbxList = @($processedList | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' })

$totalUserPrimaryBytes   = [int64](($userMbxList | Measure-Object -Property PrimaryTotalItemSizeBytes -Sum).Sum)
$totalUserArchiveBytes   = [int64](($userMbxList | Measure-Object -Property ArchiveTotalItemSizeBytes -Sum).Sum)
$totalSharedPrimaryBytes = [int64](($sharedMbxList | Measure-Object -Property PrimaryTotalItemSizeBytes -Sum).Sum)
$totalSharedArchiveBytes = [int64](($sharedMbxList | Measure-Object -Property ArchiveTotalItemSizeBytes -Sum).Sum)

$totalPrimaryBytes  = [int64](($processedList | Measure-Object -Property PrimaryTotalItemSizeBytes -Sum).Sum)
$totalArchiveBytes  = [int64](($processedList | Measure-Object -Property ArchiveTotalItemSizeBytes -Sum).Sum)
$totalCombinedBytes = $totalPrimaryBytes + $totalArchiveBytes

$totalPrimaryItems  = [int64](($processedList | Measure-Object -Property PrimaryItemCount -Sum).Sum)
$totalArchiveItems  = [int64](($processedList | Where-Object { $_.ArchiveItemCount -is [int64] } | Measure-Object -Property ArchiveItemCount -Sum).Sum)

$archiveActiveCount = @($processedList | Where-Object { $_.ArchiveEnabled }).Count
$autoExpandingCount = @($processedList | Where-Object { $_.AutoExpandingArchiveEnabled }).Count

$archivePct = if ($processedList.Count -gt 0) { [math]::Round(($archiveActiveCount / $processedList.Count) * 100, 1) } else { 0 }
$autoExpandPct = if ($processedList.Count -gt 0) { [math]::Round(($autoExpandingCount / $processedList.Count) * 100, 1) } else { 0 }

Write-Host "`n================================================================================" -ForegroundColor Cyan
Write-Host "EXECUTIVE MAILBOX & STORAGE SUMMARY" -ForegroundColor White
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ("  Total Mailboxes Analyzed     : {0}" -f $processedList.Count) -ForegroundColor Green
Write-Host ("  - User Inboxes               : {0} (Primary: {1} | Archive: {2})" -f $userMbxList.Count, (Format-ByteNumber $totalUserPrimaryBytes), (Format-ByteNumber $totalUserArchiveBytes)) -ForegroundColor White
Write-Host ("  - Shared Mailboxes           : {0} (Primary: {1} | Archive: {2})" -f $sharedMbxList.Count, (Format-ByteNumber $totalSharedPrimaryBytes), (Format-ByteNumber $totalSharedArchiveBytes)) -ForegroundColor White
Write-Host ""
Write-Host ("  Archive Adoption (Active)    : {0} of {1} ({2} %)" -f $archiveActiveCount, $processedList.Count, $archivePct) -ForegroundColor White
Write-Host ("  Auto-Expanding Archive (On)  : {0} of {1} ({2} %)" -f $autoExpandingCount, $processedList.Count, $autoExpandPct) -ForegroundColor White
Write-Host ""
Write-Host ("  Total Primary Storage Used   : {0} ({1:N0} items)" -f (Format-ByteNumber $totalPrimaryBytes), $totalPrimaryItems) -ForegroundColor Yellow
Write-Host ("  Total Archive Storage Used   : {0} ({1:N0} items)" -f (Format-ByteNumber $totalArchiveBytes), $totalArchiveItems) -ForegroundColor Yellow
Write-Host ("  Total Combined Tenant Footprint: {0}" -f (Format-ByteNumber $totalCombinedBytes)) -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Cyan

# Top 5 Largest Mailboxes
$top5Mailboxes = @($processedList | Sort-Object PrimaryTotalItemSizeBytes -Descending | Select-Object -First 5)
if ($top5Mailboxes.Count -gt 0) {
    Write-Host "`n[TOP 5 LARGEST PRIMARY MAILBOXES]" -ForegroundColor White
    $rank = 1
    foreach ($topMbx in $top5Mailboxes) {
        Write-Host ("  {0}. {1,-28} | {2,-32} | {3,12} ({4} items)" -f $rank, $topMbx.DisplayName, $topMbx.UserPrincipalName, $topMbx.PrimaryTotalItemSize, $topMbx.PrimaryItemCount) -ForegroundColor Gray
        $rank++
    }
}

# Top 5 Largest Archives
$top5Archives = @($processedList | Where-Object { $_.ArchiveEnabled -and $_.ArchiveTotalItemSizeBytes -gt 0 } | Sort-Object ArchiveTotalItemSizeBytes -Descending | Select-Object -First 5)
if ($top5Archives.Count -gt 0) {
    Write-Host "`n[TOP 5 LARGEST IN-PLACE ARCHIVES]" -ForegroundColor White
    $rank = 1
    foreach ($topArch in $top5Archives) {
        $autoExpTag = if ($topArch.AutoExpandingArchiveEnabled) { "[Auto-Expand: YES]" } else { "[Auto-Expand: NO]" }
        Write-Host ("  {0}. {1,-28} | {2,-32} | {3,12} | {4}" -f $rank, $topArch.DisplayName, $topArch.UserPrincipalName, $topArch.ArchiveTotalItemSize, $autoExpTag) -ForegroundColor Gray
        $rank++
    }
}
Write-Host ""

# ----------------------------------------------------------------------
# CSV Export
# ----------------------------------------------------------------------
if (-not [string]::IsNullOrWhiteSpace($ExportCsv)) {
    try {
        $exportData = $processedList | Select-Object `
            Nr,
            DisplayName,
            UserPrincipalName,
            PrimarySmtpAddress,
            MailboxType,
            RecipientTypeDetails,
            PrimaryTotalItemSize,
            PrimaryTotalItemSizeBytes,
            PrimaryTotalItemSizeMB,
            PrimaryTotalItemSizeGB,
            PrimaryItemCount,
            PrimaryDeletedItemSize,
            PrimaryDeletedItemSizeBytes,
            PrimaryDeletedItemCount,
            PrimaryProhibitSendReceive,
            PrimaryIssueWarning,
            ArchiveStatus,
            ArchiveEnabled,
            AutoExpandingArchive,
            AutoExpandingArchiveEnabled,
            ArchiveTotalItemSize,
            ArchiveTotalItemSizeBytes,
            ArchiveTotalItemSizeMB,
            ArchiveTotalItemSizeGB,
            ArchiveItemCount,
            ArchiveDeletedItemSize,
            ArchiveDeletedItemCount,
            ArchiveQuota,
            ArchiveWarningQuota,
            CombinedTotalSize,
            CombinedTotalSizeBytes,
            CombinedTotalSizeGB,
            RetentionPolicy,
            LitigationHoldEnabled,
            LastLogonTime

        $parentDir = Split-Path -Path $ExportCsv -Parent
        if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $exportData | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "[+] Detailed report successfully exported to CSV: $ExportCsv" -ForegroundColor Green
    } catch {
        Write-Error "Failed to export report to CSV ($ExportCsv): $_"
    }
}

# ----------------------------------------------------------------------
# Pipeline Output
# ----------------------------------------------------------------------
if ($PassThru.IsPresent) {
    return $processedList
}
