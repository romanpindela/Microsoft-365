<#
.SYNOPSIS
    Enables Exchange Online In-Place Archive or inspects existing mailbox archive configuration.

.DESCRIPTION
    Production-ready automation script designed to audit or provision In-Place Archive
    for Exchange Online mailboxes. In standard mode, it provisions the archive, triggers
    the Managed Folder Assistant (MFA), and reports detailed telemetry. With the -View switch,
    it performs a non-destructive, read-only inspection of the current archive and storage parameters.
    Complies with cross-language OS requirements and includes strict input sanitization.

.PARAMETER UserPrincipalName
    Target mailbox UPN or primary SMTP address. Must be a valid email format.

.PARAMETER View
    Performs a non-destructive read-only inspection of the mailbox and archive configuration
    without enabling the archive or triggering MFA.

.PARAMETER v
    Alias for the -View switch.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used for explicit authentication with Exchange Online.

.PARAMETER Help
    Displays this comprehensive help documentation and usage examples.

.PARAMETER h
    Alias for the -Help switch.

.EXAMPLE
    .\Enable-OnlineArchive.ps1 -UserPrincipalName "user@contoso.com"
    Enables archive, triggers MFA, and outputs diagnostic telemetry.

.EXAMPLE
    .\Enable-OnlineArchive.ps1 -UserPrincipalName "user@contoso.com" -View
    Performs a non-destructive read-only audit of current mailbox and archive settings.

.EXAMPLE
    .\Enable-OnlineArchive.ps1 -UserPrincipalName "user@contoso.com" -v -AdminUserPrincipalName "admin@contoso.onmicrosoft.com"
    Inspects archive configuration using explicit administrator login context.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.3.1
#>

[CmdletBinding(DefaultParameterSetName = 'Execution')]
param(
    [Parameter(ParameterSetName = 'Execution', Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Execution', Mandatory = $false)]
    [Alias('v')]
    [switch]$View,

    [Parameter(ParameterSetName = 'Execution', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$AdminUserPrincipalName,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata
$ScriptMetadata = @{
    Version = "1.3.1"
    Author  = "Roman Pindela"
    Email   = "roman.pindela@gmail.com"
    GitHub  = "https://github.com/romanpindela"
}

function Show-ScriptHelp {
    Write-Host "`n======================================================================" -ForegroundColor Cyan
    Write-Host " Exchange Online Archive Provisioner & Inspector | Version $($ScriptMetadata.Version)" -ForegroundColor Cyan
    Write-Host " Author: $($ScriptMetadata.Author) <$($ScriptMetadata.Email)>" -ForegroundColor Gray
    Write-Host " GitHub: $($ScriptMetadata.GitHub)" -ForegroundColor Gray
    Write-Host "======================================================================" -ForegroundColor Cyan
    Write-Host "`nDESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Manages and inspects In-Place Archive for Exchange Online mailboxes."
    Write-Host "  - Default Mode: Provisions archive if missing, invokes Managed Folder Assistant,"
    Write-Host "    and displays storage quotas and live technical telemetry."
    Write-Host "  - View Mode (-View / -v): Read-only inspection without altering any mailbox state."
    
    Write-Host "`nAUTHENTICATION & PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "  Requires Exchange Online administrative permissions (Exchange Administrator or Global Admin)."
    Write-Host "  If not already authenticated, an interactive login prompt will be initialized."
    Write-Host "  Use -AdminUserPrincipalName to pre-fill your admin login identity."

    Write-Host "`nPARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -UserPrincipalName       Target mailbox UPN / email address."
    Write-Host "  -View, -v                Read-only inspection; audits existing settings without changes."
    Write-Host "  -AdminUserPrincipalName  (Optional) Admin identity for M365 connection."
    Write-Host "  -Help, -h                Displays this help menu."

    Write-Host "`nUSAGE EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\Enable-OnlineArchive.ps1 -UserPrincipalName `"john.doe@domain.com`""
    Write-Host "  .\Enable-OnlineArchive.ps1 -UserPrincipalName `"john.doe@domain.com`" -View"
    Write-Host "  .\Enable-OnlineArchive.ps1 -UserPrincipalName `"john.doe@domain.com`" -v -AdminUserPrincipalName `"admin@domain.onmicrosoft.com`""
    Write-Host "  .\Enable-OnlineArchive.ps1 -h`n"
}

# Auto-show help if requested or invoked without parameters
if ($PSCmdlet.ParameterSetName -eq 'Help' -or [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    Show-ScriptHelp
    exit 0
}

# Language-independent SID verification for Administrator privilege
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
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

# Input validation: defense against malformed or malicious email inputs
$emailPattern = '^[a-zA-Z0-9._\%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
if ($UserPrincipalName -notmatch $emailPattern) {
    Write-Error "Invalid input: '$UserPrincipalName' is not a valid UserPrincipalName / email format."
    exit 1
}

# Ensure ExchangeOnlineManagement module is available
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "[*] ExchangeOnlineManagement module not detected. Installing..." -ForegroundColor Cyan
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
}

# Check or initiate Exchange Online connection
$sessionActive = $false
try {
    Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop | Out-Null
    $sessionActive = $true
} catch {
    $sessionActive = $false
}

if (-not $sessionActive) {
    Write-Host "[*] Initializing Microsoft 365 Exchange Online connection..." -ForegroundColor Cyan
    $connectParams = @{
        ShowProgress = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
        $connectParams['UserPrincipalName'] = $AdminUserPrincipalName
    }
    Connect-ExchangeOnline @connectParams
}

# Fetch primary mailbox configuration
Write-Host "`n[*] Inspecting mailbox: $UserPrincipalName..." -ForegroundColor Cyan
$mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop

if ($View.IsPresent) {
    Write-Host "[i] Running in READ-ONLY mode (-View). No modifications will be made." -ForegroundColor Cyan
} else {
    # Provision archive if not enabled
    if ($mailbox.ArchiveStatus -eq "None") {
        Write-Host "[+] Archive is currently disabled. Provisioning In-Place Archive..." -ForegroundColor Yellow
        Enable-Mailbox -Identity $UserPrincipalName -Archive -ErrorAction Stop
        Write-Host "[*] Waiting for Exchange directory replication (6 seconds)..." -ForegroundColor Gray
        Start-Sleep -Seconds 6
    } else {
        Write-Host "[i] In-Place Archive is already provisioned (Status: $($mailbox.ArchiveStatus))." -ForegroundColor Green
    }

    # Trigger retention processing
    Write-Host "[*] Triggering Managed Folder Assistant (MFA)..." -ForegroundColor Cyan
    try {
        Start-ManagedFolderAssistant -Identity $mailbox.ExchangeGuid.ToString() -ErrorAction Stop
    } catch {
        Write-Warning "Could not trigger MFA immediately. Background processing will run on schedule."
    }
}

# Refresh mailbox properties
$mailboxUpdated = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
$archiveStats = $null

if ($mailboxUpdated.ArchiveStatus -ne "None") {
    try {
        $archiveStats = Get-MailboxStatistics -Identity $UserPrincipalName -Archive -ErrorAction Stop
    } catch {
        Write-Warning "Archive statistics store is initializing. Data might take several minutes to become fully visible."
    }
}

# Technical telemetry output
$modeLabel = if ($View.IsPresent) { "READ-ONLY AUDIT" } else { "PROVISION & INSPECT" }
Write-Host "`n======================================================================" -ForegroundColor Green
Write-Host "      EXCHANGE ONLINE ARCHIVE TELEMETRY [$modeLabel]        " -ForegroundColor Green
Write-Host "======================================================================" -ForegroundColor Green

Write-Host "`n[1] IDENTITY & RECIPIENT INFORMATION" -ForegroundColor Cyan
[PSCustomObject]@{
    DisplayName              = $mailboxUpdated.DisplayName
    UserPrincipalName        = $mailboxUpdated.UserPrincipalName
    PrimarySmtpAddress       = $mailboxUpdated.PrimarySmtpAddress
    ArchiveName              = if ($mailboxUpdated.ArchiveName) { ($mailboxUpdated.ArchiveName -join ", ") } else { "N/A" }
    ArchiveStatus            = $mailboxUpdated.ArchiveStatus
    ArchiveState             = $mailboxUpdated.ArchiveState
    RetentionPolicy          = if ($mailboxUpdated.RetentionPolicy) { $mailboxUpdated.RetentionPolicy } else { "None" }
} | Format-List

Write-Host "[2] TECHNICAL DIRECTORY ATTRIBUTES" -ForegroundColor Cyan
[PSCustomObject]@{
    ExchangeGuid             = $mailboxUpdated.ExchangeGuid
    ArchiveGuid              = $mailboxUpdated.ArchiveGuid
    ArchiveDatabase          = if ($mailboxUpdated.ArchiveDatabase) { $mailboxUpdated.ArchiveDatabase } else { "Not Provisioned" }
    ArchiveDomain            = if ($mailboxUpdated.ArchiveDomain) { $mailboxUpdated.ArchiveDomain } else { "N/A" }
    AutoExpandingArchive     = $mailboxUpdated.AutoExpandingArchiveEnabled
    HiddenFromGAL            = $mailboxUpdated.HiddenFromAddressListsEnabled
    LitigationHoldEnabled    = $mailboxUpdated.LitigationHoldEnabled
} | Format-List

Write-Host "[3] ARCHIVE STORAGE QUOTAS" -ForegroundColor Cyan
[PSCustomObject]@{
    ArchiveWarningQuota      = $mailboxUpdated.ArchiveWarningQuota
    ArchiveQuota             = $mailboxUpdated.ArchiveQuota
    UseDatabaseQuotaDefaults = $mailboxUpdated.UseDatabaseQuotaDefaults
} | Format-List

if ($archiveStats) {
    Write-Host "[4] LIVE STORAGE USAGE METRICS" -ForegroundColor Cyan
    [PSCustomObject]@{
        TotalItemCount       = $archiveStats.ItemCount
        TotalItemSize        = $archiveStats.TotalItemSize.Value
        DeletedItemCount     = $archiveStats.DeletedItemCount
        TotalDeletedItemSize = $archiveStats.TotalDeletedItemSize.Value
        LastLogonTime        = $archiveStats.LastLogonTime
    } | Format-List
} elseif ($mailboxUpdated.ArchiveStatus -eq "None") {
    Write-Host "[!] Archive is not enabled for this mailbox. No storage statistics exist." -ForegroundColor Yellow
} else {
    Write-Host "[!] Live storage metrics pending directory cache generation." -ForegroundColor Gray
}

Write-Host "======================================================================`n" -ForegroundColor Green

