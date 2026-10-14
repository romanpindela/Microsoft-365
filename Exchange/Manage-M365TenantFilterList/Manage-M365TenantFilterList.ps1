<#
.SYNOPSIS
    Manages Microsoft 365 Tenant Allow/Block Lists (whitelist and blacklist for sender domains and email addresses) in Exchange Online Protection.

.DESCRIPTION
    Comprehensive, enterprise-grade PowerShell automation script designed for Microsoft 365 and Exchange Online
    administrators to audit, search, add, and remove sender domains and email addresses from the Tenant Allow/Block List (TABL).

    Key capabilities include:
    1. Automatic Exchange Online administrator session detection and modern interactive login.
    2. Whitelist (Allow List) Management:
       - Add trusted sender domains (e.g. partner.com) or specific email addresses (e.g. alert@partner.com).
       - List all currently allowed domains and addresses with comprehensive details (expiration, status, author, notes).
       - Search for specific domains or addresses within the allow list.
    3. Blacklist (Block List) Management:
       - Add malicious, spam, or phishing domains and email addresses to the tenant-wide block list.
       - List all currently blocked domains and addresses with full administrative details.
       - Search for specific domains or addresses within the block list.
    4. Advanced Auditing & Inspection:
       - Displays entry value, list action (Allow/Block), entry type (Domain vs Email Address), expiration date,
         remaining days active, status (Active/Expired), last modified date, administrator identity, and notes.
       - Detailed card view (-Detailed) for in-depth forensics.
    5. Search Engine:
       - Wildcard and substring search (-Search) across domains, email addresses, administrative notes, and author identities.
    6. Safety & Conflict Protection:
       - Pre-flight format validation (RFC-compliant domain and email syntax validation).
       - Input sanitization guarding against malicious control characters and shell injections.
       - Duplicate and conflict detection: warns if an entry already exists in the opposing list (e.g., trying to allow an already blocked domain).
       - Bilingual interactive confirmation prompts (English Y/N, Polish T/N) with non-interactive override (-Force).
    7. Export & Automation:
       - Export audit records to UTF-8 CSV reports (-ExportCsv).
       - Pipeline support (-PassThru) for downstream PowerShell automation.
       - Interactive console menu wizard (-Interactive / -Menu) for menu-driven administration.

.PARAMETER List
    Switch parameter. Lists entries from the Tenant Allow/Block List based on scope (-Whitelist, -Blacklist, or All).
    Aliases: -ListEntries, -l, -Show.

.PARAMETER Add
    Switch parameter. Adds one or more sender domains or email addresses to the designated list (-Whitelist or -Blacklist).
    Aliases: -AddEntry, -a.

.PARAMETER Remove
    Switch parameter. Removes one or more entries from the designated list (-Whitelist or -Blacklist).
    Aliases: -RemoveEntry, -r.

.PARAMETER Search
    String parameter. Queries the Tenant Allow/Block List for domains, email addresses, or notes matching a pattern.
    Aliases: -Find, -s, -Query.

.PARAMETER Whitelist
    Switch parameter. Targets or filters operations to the Allow List (trusted domains and senders).
    Aliases: -Allow, -AllowList, -w.

.PARAMETER Blacklist
    Switch parameter. Targets or filters operations to the Block List (malicious or unwanted domains and senders).
    Aliases: -Block, -BlockList, -b.

.PARAMETER Entries
    One or more domains (e.g., "partner.com") or email addresses (e.g., "user@badactor.org") to add, remove, or inspect.
    Aliases: -Entry, -Domain, -Domains, -Address, -Addresses, -Identity.

.PARAMETER Notes
    Optional administrative justification or comment associated with added entries (e.g., "Ticket #9842 - Approved Vendor").
    Aliases: -Comment, -Description.

.PARAMETER NoExpiration
    Switch parameter. Specifies that the added entries should never expire (subject to tenant policy and list type).
    Aliases: -NeverExpire.

.PARAMETER ExpirationDays
    Integer parameter (1-730). Sets entry expiration to the specified number of days from today.
    Aliases: -Days.

.PARAMETER ExpirationDate
    DateTime parameter. Sets an explicit UTC expiration date and time for added entries.
    Aliases: -ExpireAt.

.PARAMETER Detailed
    Switch parameter. Displays in-depth configuration cards including GUIDs, timestamps, and full notes.
    Aliases: -d, -Details.

.PARAMETER ExportCsv
    Optional file path to export retrieved allow/block records into a UTF-8 CSV file.
    Aliases: -CsvPath, -Csv, -Export.

.PARAMETER PassThru
    Switch parameter. Outputs custom PowerShell objects to the pipeline for downstream processing.

.PARAMETER Interactive
    Switch parameter. Launches an interactive menu-driven wizard in the console.
    Aliases: -Menu, -i.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used to pre-fill the Microsoft 365 Exchange Online authentication prompt.
    Aliases: -AdminUPN.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N/T/N) for unattended automation.
    Aliases: -f.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1
    Running without parameters displays the built-in help banner, usage examples, author info, and version.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Help
    Displays comprehensive help documentation and parameter reference.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List -Whitelist
    Lists all allowed sender domains and email addresses in a formatted table with detailed status and expiration info.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List -Blacklist -Detailed
    Lists all blocked domains and email addresses with comprehensive detailed cards (including notes and modifier UPN).

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List
    Lists both Whitelist and Blacklist entries across the tenant with summary metrics.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Search "partner.com" -Whitelist
    Searches for "partner.com" specifically within the Whitelist.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Search "spammer"
    Searches for "spammer" across both Whitelist and Blacklist entries.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Add -Whitelist -Entries "trustedpartner.com", "billing@vendor.com" -Notes "Approved external supplier"
    Adds trustedpartner.com and billing@vendor.com to the tenant Whitelist with interactive confirmation.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Add -Blacklist -Entries "phishing-bank.com", "badactor@evil.org" -Notes "Malicious phishing campaign" -Force
    Adds malicious domain and sender address to the tenant Blacklist silently without confirmation prompts.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Remove -Blacklist -Entries "phishing-bank.com"
    Removes phishing-bank.com from the Blacklist after interactive confirmation.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List -Whitelist -ExportCsv "C:\Reports\M365_Whitelist.csv"
    Audits the tenant Whitelist and exports all records to a CSV report.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Interactive
    Launches an interactive, menu-driven console wizard for managing tenant filter lists.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Operations
    [Parameter(ParameterSetName = 'List')]
    [Alias('ListEntries', 'l', 'Show')]
    [switch]$List,

    [Parameter(ParameterSetName = 'Add')]
    [Alias('AddEntry', 'a')]
    [switch]$Add,

    [Parameter(ParameterSetName = 'Remove')]
    [Alias('RemoveEntry', 'r')]
    [switch]$Remove,

    [Parameter(ParameterSetName = 'Search')]
    [Alias('Find', 's', 'Query')]
    [string]$Search,

    [Parameter(ParameterSetName = 'Interactive')]
    [Alias('Menu', 'i')]
    [switch]$Interactive,

    # List Scope (Whitelist vs Blacklist)
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'Add')]
    [Parameter(ParameterSetName = 'Remove')]
    [Parameter(ParameterSetName = 'Search')]
    [Alias('Allow', 'AllowList', 'w')]
    [switch]$Whitelist,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'Add')]
    [Parameter(ParameterSetName = 'Remove')]
    [Parameter(ParameterSetName = 'Search')]
    [Alias('Block', 'BlockList', 'b')]
    [switch]$Blacklist,

    # Target Entries (Domains / Email addresses)
    [Parameter(ParameterSetName = 'Default', Position = 0)]
    [Parameter(ParameterSetName = 'Add', Mandatory = $false, Position = 0)]
    [Parameter(ParameterSetName = 'Remove', Mandatory = $false, Position = 0)]
    [Alias('Entry', 'Domain', 'Domains', 'Address', 'Addresses', 'Identity')]
    [string[]]$Entries,

    # Entry Attributes
    [Parameter(ParameterSetName = 'Add')]
    [Alias('Comment', 'Description')]
    [string]$Notes,

    [Parameter(ParameterSetName = 'Add')]
    [Alias('NeverExpire')]
    [switch]$NoExpiration,

    [Parameter(ParameterSetName = 'Add')]
    [Alias('Days')]
    [ValidateRange(1, 730)]
    [int]$ExpirationDays,

    [Parameter(ParameterSetName = 'Add')]
    [Alias('ExpireAt')]
    [datetime]$ExpirationDate,

    # Output & View
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'Search')]
    [Alias('d', 'Details')]
    [switch]$Detailed,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'Search')]
    [Alias('CsvPath', 'Csv', 'Export')]
    [string]$ExportCsv,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'List')]
    [Parameter(ParameterSetName = 'Search')]
    [switch]$PassThru,

    # Admin Session & Safety
    [Parameter()]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

    [Parameter()]
    [Alias('f')]
    [switch]$Force,

    # Help
    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata declaration
$ScriptMetadata = @{
    Title   = "Manage-M365TenantFilterList.ps1"
    Version = "1.0.0"
    Author  = "Roman Pindela"
    Email   = "roman.pindela@gmail.com"
    GitHub  = "https://github.com/romanpindela"
}

function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: $($ScriptMetadata.Title)",
        "VERSION: $($ScriptMetadata.Version)",
        "AUTHOR: $($ScriptMetadata.Author)",
        "CONTACT: $($ScriptMetadata.Email) | $($ScriptMetadata.GitHub)",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Comprehensive PowerShell automation script for Microsoft 365 Exchange Online",
        "    administrators to audit, search, add, and remove sender domains and email",
        "    addresses in the Tenant Allow/Block List (Whitelist and Blacklist).",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    - Requires PowerShell 5.1 or PowerShell 7+.",
        "    - Requires 'ExchangeOnlineManagement' module (auto-installed if missing).",
        "    - Requires an Exchange Online administrator role (Security Administrator,",
        "      Exchange Administrator, or Global Administrator).",
        "    - Automatically detects active sessions; if not connected, prompts for",
        "      interactive administrator sign-in.",
        "    - Optional -AdminUserPrincipalName can pre-populate the administrator identity.",
        "",
        "USAGE EXAMPLES:",
        "    1. Display Built-In Help & Info (default behavior when run without parameters):",
        "       .\Manage-M365TenantFilterList.ps1",
        "       .\Manage-M365TenantFilterList.ps1 -h",
        "",
        "    2. List Whitelist (Allowed Domains and Email Addresses):",
        "       .\Manage-M365TenantFilterList.ps1 -List -Whitelist",
        "       .\Manage-M365TenantFilterList.ps1 -List -Whitelist -Detailed",
        "",
        "    3. List Blacklist (Blocked Domains and Email Addresses):",
        "       .\Manage-M365TenantFilterList.ps1 -List -Blacklist",
        "       .\Manage-M365TenantFilterList.ps1 -List -Blacklist -Detailed",
        "",
        "    4. List Both Whitelist and Blacklist with Summary:",
        "       .\Manage-M365TenantFilterList.ps1 -List",
        "",
        "    5. Search for a Domain or Email Address in Whitelist:",
        "       .\Manage-M365TenantFilterList.ps1 -Search `"partner.com`" -Whitelist",
        "",
        "    6. Search for a Domain or Email Address in Blacklist:",
        "       .\Manage-M365TenantFilterList.ps1 -Search `"malicious`" -Blacklist",
        "",
        "    7. Search Across Both Lists (Global Search):",
        "       .\Manage-M365TenantFilterList.ps1 -Search `"example.com`"",
        "",
        "    8. Add Domain(s) to Whitelist (Interactive Confirmation):",
        "       .\Manage-M365TenantFilterList.ps1 -Add -Whitelist -Entries `"partner.com`" -Notes `"Trusted supplier`"",
        "       .\Manage-M365TenantFilterList.ps1 -Add -Whitelist -Entries `"domain1.com`", `"domain2.com`"",
        "",
        "    9. Add Domain(s) or Email(s) to Blacklist (Silent Execution via -Force):",
        "       .\Manage-M365TenantFilterList.ps1 -Add -Blacklist -Entries `"spam.com`", `"bot@phish.org`" -Force",
        "",
        "   10. Remove Entry from Whitelist or Blacklist:",
        "       .\Manage-M365TenantFilterList.ps1 -Remove -Whitelist -Entries `"oldpartner.com`"",
        "       .\Manage-M365TenantFilterList.ps1 -Remove -Blacklist -Entries `"spam.com`" -Force",
        "",
        "   11. Export Tenant Filter List to UTF-8 CSV Report:",
        "       .\Manage-M365TenantFilterList.ps1 -List -Whitelist -ExportCsv `"C:\Reports\Whitelist.csv`"",
        "       .\Manage-M365TenantFilterList.ps1 -List -ExportCsv `"C:\Reports\AllFilters.csv`"",
        "",
        "   12. Interactive Console Menu Wizard:",
        "       .\Manage-M365TenantFilterList.ps1 -Interactive",
        "",
        "PARAMETERS:",
        "    -List, -l                  Lists entries from Whitelist, Blacklist, or both.",
        "    -Add, -a                   Adds specified entries to the chosen list.",
        "    -Remove, -r                Removes specified entries from the chosen list.",
        "    -Search, -s                Finds entries matching a domain, email, or notes pattern.",
        "    -Whitelist, -Allow, -w     Specifies the Allow List (trusted domains and senders).",
        "    -Blacklist, -Block, -b     Specifies the Block List (blocked domains and senders).",
        "    -Entries, -Domain, -Entry  One or more domains or email addresses to manage.",
        "    -Notes, -Comment           Administrative justification / comment.",
        "    -NoExpiration              Configures entry never to expire automatically.",
        "    -ExpirationDays            Number of days until expiration (1-730).",
        "    -ExpirationDate            Explicit expiration date and time (UTC).",
        "    -Detailed, -d              Displays comprehensive cards for listed records.",
        "    -ExportCsv, -Export        Exports results to the specified CSV file path.",
        "    -PassThru                  Outputs PowerShell custom objects to pipeline.",
        "    -Interactive, -Menu, -i    Launches guided console menu wizard.",
        "    -AdminUserPrincipalName    (Optional) Administrator UPN for login pre-fill.",
        "    -Force, -f                 Bypasses interactive confirmation prompts (Y/N/T/N).",
        "    -Help, -h                  Displays this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Manage-M365TenantFilterList.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Auto-show help when executed without parameters, or with -Help / -h
$hasExplicitAction = $List -or $Add -or $Remove -or (-not [string]::IsNullOrWhiteSpace($Search)) -or $Interactive -or ($Entries -and $Entries.Count -gt 0)
if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or (-not $hasExplicitAction)) {
    Show-ScriptHelp
    exit 0
}

# Defensive Input Validation Helpers
function Assert-SafeInputString {
    param(
        [string]$InputText,
        [string]$ParamName,
        [int]$MaxLength = 256
    )
    if ([string]::IsNullOrWhiteSpace($InputText)) { return }
    if ($InputText.Length -gt $MaxLength) {
        Write-Error "Parameter -$ParamName exceeds maximum allowed length of $MaxLength characters."
        exit 1
    }
    # Defense against control characters and command injection
    if ($InputText -match '[`$<>|;&\0]') {
        Write-Error "Parameter -$ParamName contains invalid or potentially dangerous characters."
        exit 1
    }
}

function Test-ValidDomainOrEmail {
    param([string]$Entry)

    if ([string]::IsNullOrWhiteSpace($Entry)) {
        return @{ IsValid = $false; EntryType = 'Unknown'; CleanValue = $Entry; Reason = "Value cannot be empty or whitespace." }
    }

    $clean = $Entry.Trim().ToLowerInvariant()

    # Email format validation
    $emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if ($clean -match $emailRegex) {
        return @{ IsValid = $true; EntryType = 'EmailAddress'; CleanValue = $clean; Reason = $null }
    }

    # Domain format validation (supports plain domain like contoso.com, or wildcard subdomain *.contoso.com)
    $domainRegex = '^(\*\.)?([a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$'
    if ($clean -match $domainRegex) {
        return @{ IsValid = $true; EntryType = 'Domain'; CleanValue = $clean; Reason = $null }
    }

    return @{ IsValid = $false; EntryType = 'Invalid'; CleanValue = $clean; Reason = "Entry '$clean' is neither a valid domain name nor a valid email address." }
}

# Validate string parameters
Assert-SafeInputString -InputText $Search -ParamName "Search"
Assert-SafeInputString -InputText $Notes -ParamName "Notes" -MaxLength 500

# Validate scope consistency
if ($Whitelist -and $Blacklist) {
    if ($Add -or $Remove) {
        Write-Error "Conflicting parameters: You cannot specify both -Whitelist and -Blacklist when adding or removing entries."
        exit 1
    }
}

if ($Add -or $Remove) {
    if (-not $Whitelist -and -not $Blacklist) {
        Write-Error "Missing target list: You must specify either -Whitelist or -Blacklist when performing -Add or -Remove operations."
        exit 1
    }
    if (-not $Entries -or $Entries.Count -eq 0) {
        Write-Error "Missing parameter: -Entries is required when adding or removing items. Example: -Entries `"example.com`""
        exit 1
    }

    # Pre-flight format validation before cloud connection
    $invalidInputs = @()
    foreach ($entry in $Entries) {
        $check = Test-ValidDomainOrEmail -Entry $entry
        if (-not $check.IsValid) {
            $invalidInputs += $check
        }
    }
    if ($invalidInputs.Count -gt 0) {
        Write-Host "`n[!] Invalid domain or email address format detected:" -ForegroundColor Red
        foreach ($inv in $invalidInputs) {
            Write-Host "    - $($inv.CleanValue) : $($inv.Reason)" -ForegroundColor Yellow
        }
        Write-Error "Pre-flight validation failed. Please correct domain/email syntax before proceeding."
        exit 1
    }
}

# Validate CSV path if provided
if (-not [string]::IsNullOrWhiteSpace($ExportCsv)) {
    try {
        $csvDirectory = Split-Path -Path $ExportCsv -Parent
        if (-not [string]::IsNullOrWhiteSpace($csvDirectory) -and -not (Test-Path -Path $csvDirectory)) {
            Write-Host "[*] Creating output directory: '$csvDirectory'..." -ForegroundColor Cyan
            New-Item -ItemType Directory -Path $csvDirectory -Force | Out-Null
        }
    } catch {
        Write-Error "Invalid -ExportCsv path specified: $_"
        exit 1
    }
}

# Exchange Online Connection Manager
function Connect-ExchangeEnvironment {
    param([string]$AdminUPN)

    # 1. Check prerequisite module
    if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Write-Host "[!] ExchangeOnlineManagement module is not installed." -ForegroundColor Yellow
        Write-Host "[*] Installing ExchangeOnlineManagement module for CurrentUser from PSGallery..." -ForegroundColor Cyan
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "[+] ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to install ExchangeOnlineManagement module: $_`nPlease install manually via: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
            exit 1
        }
    }

    # 2. Check active session or prompt for administrator login
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        Write-Host "[+] Active Exchange Online session detected." -ForegroundColor Green
    } catch {
        Write-Host "[*] No active Exchange Online session detected. Initiating administrator login..." -ForegroundColor Cyan
        try {
            if (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
                Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
        } catch {
            Write-Error "Failed to connect to Exchange Online: $_"
            exit 1
        }
    }
}

# Data Retrieval Helper
function Get-TenantFilterData {
    param(
        [ValidateSet('All', 'Allow', 'Block')]
        [string]$TargetAction = 'All'
    )

    Write-Host "[*] Querying Microsoft 365 Tenant Allow/Block List (Sender scope)..." -ForegroundColor Cyan
    $rawItems = @()

    try {
        if ($TargetAction -eq 'Allow') {
            $rawItems = Get-TenantAllowBlockListItems -ListType Sender -Allow -ErrorAction Stop
        } elseif ($TargetAction -eq 'Block') {
            $rawItems = Get-TenantAllowBlockListItems -ListType Sender -Block -ErrorAction Stop
        } else {
            $rawItems = Get-TenantAllowBlockListItems -ListType Sender -ErrorAction Stop
        }
    } catch {
        Write-Error "Failed to query Tenant Allow/Block List from Exchange Online: $_"
        exit 1
    }

    $processed = @()
    $nowUtc = (Get-Date).ToUniversalTime()

    foreach ($item in $rawItems) {
        $val = [string]$item.Value
        $entryType = if ($val -match '@') { 'EmailAddress' } else { 'Domain' }
        $action = [string]$item.Action
        if ([string]::IsNullOrWhiteSpace($action)) {
            $action = if ($item.Block) { 'Block' } else { 'Allow' }
        }

        # Calculate expiration metrics
        $expDate = $null
        $remainingDays = "Never"
        $status = "Active"

        if ($item.ExpirationDate) {
            try {
                $expDate = [datetime]$item.ExpirationDate
                if ($expDate.Year -lt 9999) {
                    $diff = $expDate - $nowUtc
                    if ($diff.TotalSeconds -lt 0) {
                        $status = "Expired"
                        $remainingDays = "Expired"
                    } else {
                        $remainingDays = [math]::Round($diff.TotalDays, 1)
                    }
                }
            } catch {
                # Fallback on raw string
            }
        }

        $obj = [PSCustomObject]@{
            Identity         = [string]$item.Identity
            Value            = $val
            Action           = $action
            ListType         = 'Sender'
            EntryType        = $entryType
            Status           = $status
            ExpirationDate   = if ($expDate -and $expDate.Year -lt 9999) { $expDate.ToString("yyyy-MM-dd HH:mm UTC") } else { "Never" }
            RemainingDays    = $remainingDays
            LastModifiedDate = if ($item.LastModifiedDate) { ([datetime]$item.LastModifiedDate).ToString("yyyy-MM-dd HH:mm UTC") } else { "N/A" }
            LastModifiedBy   = [string]$item.LastModifiedBy
            Notes            = [string]$item.Notes
        }

        $processed += $obj
    }

    return $processed
}

# Display Tables and Cards
function Show-TenantFilterSummaryTable {
    param(
        [array]$Items,
        [string]$Title = "Tenant Filter List"
    )

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "`n[i] No entries found for $Title." -ForegroundColor Yellow
        return
    }

    Write-Host "`n================================================================================" -ForegroundColor White
    Write-Host "  $Title (Total: $($Items.Count))" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor White

    # Group metrics
    $allowCount = ($Items | Where-Object { $_.Action -eq 'Allow' }).Count
    $blockCount = ($Items | Where-Object { $_.Action -eq 'Block' }).Count
    $domainCount = ($Items | Where-Object { $_.EntryType -eq 'Domain' }).Count
    $emailCount = ($Items | Where-Object { $_.EntryType -eq 'EmailAddress' }).Count
    $activeCount = ($Items | Where-Object { $_.Status -eq 'Active' }).Count
    $expiredCount = ($Items | Where-Object { $_.Status -eq 'Expired' }).Count

    Write-Host "Summary: Whitelist (Allow): $allowCount | Blacklist (Block): $blockCount | Domains: $domainCount | Emails: $emailCount | Active: $activeCount | Expired: $expiredCount" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

    $tableData = $Items | Select-Object `
        @{Name='Action'; Expression={ $_.Action.ToUpper() }},
        @{Name='Type'; Expression={ $_.EntryType }},
        @{Name='Entry Value'; Expression={ $_.Value }},
        @{Name='Status'; Expression={ $_.Status }},
        @{Name='Expires (UTC)'; Expression={ $_.ExpirationDate }},
        @{Name='Days Left'; Expression={ $_.RemainingDays }},
        @{Name='Notes'; Expression={ if ($_.Notes.Length -gt 30) { $_.Notes.Substring(0, 27) + "..." } else { $_.Notes } }}

    $tableData | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White
}

function Show-TenantFilterDetailedCards {
    param([array]$Items)

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "`n[i] No entries to display in detailed view." -ForegroundColor Yellow
        return
    }

    Write-Host "`n================================================================================" -ForegroundColor White
    Write-Host "  Detailed Inspection Report (Total Entries: $($Items.Count))" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor White

    foreach ($item in $Items) {
        $actionColor = if ($item.Action -eq 'Allow') { 'Green' } else { 'Red' }
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host " Entry: " -NoNewline -ForegroundColor White
        Write-Host "$($item.Value)" -ForegroundColor Cyan
        Write-Host "  List / Action     : " -NoNewline -ForegroundColor Gray
        Write-Host "$($item.Action)" -ForegroundColor $actionColor
        Write-Host "  Entry Type        : $($item.EntryType)" -ForegroundColor Gray
        Write-Host "  Status            : $($item.Status)" -ForegroundColor (if ($item.Status -eq 'Active') { 'Green' } else { 'Yellow' })
        Write-Host "  Expiration Date   : $($item.ExpirationDate) ($($item.RemainingDays) days remaining)" -ForegroundColor Gray
        Write-Host "  Last Modified Date: $($item.LastModifiedDate)" -ForegroundColor Gray
        Write-Host "  Last Modified By  : $($item.LastModifiedBy)" -ForegroundColor Gray
        Write-Host "  Notes             : $(if ($item.Notes) { $item.Notes } else { 'None' })" -ForegroundColor Yellow
        Write-Host "  System ID         : $($item.Identity)" -ForegroundColor DarkGray
    }
    Write-Host "--------------------------------------------------------------------------------`n" -ForegroundColor DarkGray
}

# Search Operation
function Invoke-SearchOperation {
    param(
        [string]$Query,
        [string]$TargetAction,
        [switch]$ShowDetails,
        [string]$CsvPath,
        [switch]$PipelineOut
    )

    $items = Get-TenantFilterData -TargetAction $TargetAction
    $cleanQuery = $Query.Trim()

    Write-Host "[*] Searching for '$cleanQuery' across entries and notes..." -ForegroundColor Cyan

    $matches = $items | Where-Object {
        $_.Value -like "*$cleanQuery*" -or
        $_.Notes -like "*$cleanQuery*" -or
        $_.LastModifiedBy -like "*$cleanQuery*"
    }

    $title = if ($TargetAction -eq 'Allow') { "Whitelist Search Results for '$cleanQuery'" }
             elseif ($TargetAction -eq 'Block') { "Blacklist Search Results for '$cleanQuery'" }
             else { "Tenant Filter Search Results for '$cleanQuery'" }

    if ($ShowDetails) {
        Show-TenantFilterDetailedCards -Items $matches
    } else {
        Show-TenantFilterSummaryTable -Items $matches -Title $title
    }

    if (-not [string]::IsNullOrWhiteSpace($CsvPath) -and $matches.Count -gt 0) {
        Export-TenantFilterReport -Items $matches -Path $CsvPath
    }

    if ($PipelineOut) {
        $matches
    }
}

# Add Operation
function Invoke-AddOperation {
    param(
        [string[]]$RawEntries,
        [ValidateSet('Allow', 'Block')]
        [string]$TargetAction,
        [string]$AdminNotes,
        [switch]$NeverExpire,
        [int]$DaysToExpire,
        [datetime]$ExplicitExpireDate,
        [switch]$BypassPrompt
    )

    $targetListName = if ($TargetAction -eq 'Allow') { "Whitelist (Allow List)" } else { "Blacklist (Block List)" }
    Write-Host "`n[*] Preparing to add entries to $targetListName..." -ForegroundColor Cyan

    # 1. Validate formats
    $validatedEntries = @()
    $invalidEntries = @()

    foreach ($e in $RawEntries) {
        $valResult = Test-ValidDomainOrEmail -Entry $e
        if ($valResult.IsValid) {
            $validatedEntries += $valResult.CleanValue
        } else {
            $invalidEntries += $valResult
        }
    }

    if ($invalidEntries.Count -gt 0) {
        Write-Host "`n[!] The following entries failed domain/email syntax validation and will be skipped:" -ForegroundColor Red
        foreach ($inv in $invalidEntries) {
            Write-Host "    - $($inv.CleanValue) : $($inv.Reason)" -ForegroundColor Yellow
        }
    }

    if ($validatedEntries.Count -eq 0) {
        Write-Error "No valid domains or email addresses provided to add."
        exit 1
    }

    # Remove duplicates from provided array
    $uniqueEntries = $validatedEntries | Select-Object -Unique

    # 2. Pre-flight check against existing entries in Exchange Online
    $existing = Get-TenantFilterData -TargetAction 'All'
    $alreadyInTarget = @()
    $conflictOpposite = @()
    $finalToAdd = @()

    foreach ($entry in $uniqueEntries) {
        $match = $existing | Where-Object { $_.Value -eq $entry }
        if ($match) {
            if ($match.Action -eq $TargetAction) {
                $alreadyInTarget += $entry
            } else {
                $conflictOpposite += [PSCustomObject]@{
                    Entry          = $entry
                    CurrentAction  = $match.Action
                    ProposedAction = $TargetAction
                }
            }
        } else {
            $finalToAdd += $entry
        }
    }

    if ($alreadyInTarget.Count -gt 0) {
        Write-Host "`n[i] The following entries already exist in the target $targetListName (skipping):" -ForegroundColor Yellow
        foreach ($item in $alreadyInTarget) {
            Write-Host "    - $item" -ForegroundColor Gray
        }
    }

    if ($conflictOpposite.Count -gt 0) {
        Write-Host "`n[!] CONFLICT DETECTED: The following entries currently exist in the OPPOSITE list!" -ForegroundColor Red
        foreach ($c in $conflictOpposite) {
            Write-Host "    - $($c.Entry) (Currently: $($c.CurrentAction), Proposed: $($c.ProposedAction))" -ForegroundColor Yellow
        }
        Write-Host "[!] In Microsoft 365, you should remove the item from the conflicting list before adding to the opposing list." -ForegroundColor Yellow
    }

    if ($finalToAdd.Count -eq 0) {
        Write-Host "`n[+] No new entries to add. Everything is up to date." -ForegroundColor Green
        return
    }

    # 3. Determine expiration settings
    $effectiveExpParam = @{}
    if ($NeverExpire) {
        $effectiveExpParam['NoExpiration'] = $true
        Write-Host "    Expiration Policy : No Expiration (Never Expire)" -ForegroundColor Gray
    } elseif ($ExplicitExpireDate -and $ExplicitExpireDate.Year -gt 1) {
        $effectiveExpParam['ExpirationDate'] = $ExplicitExpireDate.ToUniversalTime()
        Write-Host "    Expiration Policy : Explicit Date ($($ExplicitExpireDate.ToString('yyyy-MM-dd HH:mm UTC')))" -ForegroundColor Gray
    } elseif ($DaysToExpire -gt 0) {
        $effectiveExpParam['ExpirationDate'] = (Get-Date).AddDays($DaysToExpire).ToUniversalTime()
        Write-Host "    Expiration Policy : $DaysToExpire days from today" -ForegroundColor Gray
    } else {
        # Default policy
        if ($TargetAction -eq 'Allow') {
            # In Microsoft 365, Allow entries by default have a 30-day maximum lifetime unless submissions are verified
            $effectiveExpParam['ExpirationDate'] = (Get-Date).AddDays(30).ToUniversalTime()
            Write-Host "    Expiration Policy : Default 30-day lifetime (Microsoft 365 recommended for Allow entries)" -ForegroundColor Gray
        } else {
            # Block list items can safely have no expiration
            $effectiveExpParam['NoExpiration'] = $true
            Write-Host "    Expiration Policy : Default No Expiration (Permanent Block)" -ForegroundColor Gray
        }
    }

    # 4. Confirmation Prompt
    Write-Host "`nPlanned Action Summary:" -ForegroundColor White
    Write-Host "  Operation   : ADD" -ForegroundColor White
    Write-Host "  Target List : $targetListName" -ForegroundColor (if ($TargetAction -eq 'Allow') { 'Green' } else { 'Red' })
    Write-Host "  Notes       : $(if ($AdminNotes) { $AdminNotes } else { 'None' })" -ForegroundColor White
    Write-Host "  Entries ($($finalToAdd.Count)): " -ForegroundColor White
    foreach ($entry in $finalToAdd) {
        Write-Host "    + $entry" -ForegroundColor Cyan
    }

    if (-not $BypassPrompt) {
        $promptQuestion = if ($TargetAction -eq 'Allow') {
            "Are you sure you want to ADD these entries to the Whitelist? / Czy na pewno chcesz dodać te wpisy do białej listy? (Y/N/T/N)"
        } else {
            "Are you sure you want to ADD these entries to the Blacklist? / Czy na pewno chcesz dodać te wpisy do czarnej listy? (Y/N/T/N)"
        }

        $confirmation = Read-Host $promptQuestion
        if ($confirmation -notmatch "^(Y|y|T|t|Yes|Tak)$") {
            Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without interactive confirmation." -ForegroundColor Cyan
    }

    # 5. Execute New-TenantAllowBlockListItems
    Write-Host "[*] Committing entries to Microsoft 365 Tenant Allow/Block List..." -ForegroundColor Cyan

    $cmdArgs = @{
        ListType    = 'Sender'
        Entries     = $finalToAdd
        ErrorAction = 'Stop'
    }

    if ($TargetAction -eq 'Allow') {
        $cmdArgs['Allow'] = $true
    } else {
        $cmdArgs['Block'] = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($AdminNotes)) {
        $cmdArgs['Notes'] = $AdminNotes
    }

    if ($effectiveExpParam.ContainsKey('NoExpiration')) {
        $cmdArgs['NoExpiration'] = $true
    } elseif ($effectiveExpParam.ContainsKey('ExpirationDate')) {
        $cmdArgs['ExpirationDate'] = $effectiveExpParam['ExpirationDate']
    }

    try {
        New-TenantAllowBlockListItems @cmdArgs | Out-Null
        Write-Host "[+] Successfully added $($finalToAdd.Count) entries to $targetListName!" -ForegroundColor Green
    } catch {
        # Check if error is related to NoExpiration constraint for Allow entries
        if ($_.ToString() -match "NoExpiration" -and $TargetAction -eq 'Allow') {
            Write-Host "[!] Microsoft 365 policy requires an explicit expiration date for Allow entries (up to 30 or 90 days)." -ForegroundColor Yellow
            Write-Host "[*] Retrying with a 30-day expiration date..." -ForegroundColor Cyan
            $cmdArgs.Remove('NoExpiration')
            $cmdArgs['ExpirationDate'] = (Get-Date).AddDays(30).ToUniversalTime()
            try {
                New-TenantAllowBlockListItems @cmdArgs | Out-Null
                Write-Host "[+] Successfully added $($finalToAdd.Count) entries with 30-day expiration!" -ForegroundColor Green
            } catch {
                Write-Error "Failed to add entries to $($targetListName). Details: $_"
                exit 1
            }
        } else {
            Write-Error "Failed to add entries to $($targetListName). Details: $_"
            exit 1
        }
    }
}

# Remove Operation
function Invoke-RemoveOperation {
    param(
        [string[]]$RawEntries,
        [ValidateSet('Allow', 'Block')]
        [string]$TargetAction,
        [switch]$BypassPrompt
    )

    $targetListName = if ($TargetAction -eq 'Allow') { "Whitelist (Allow List)" } else { "Blacklist (Block List)" }
    Write-Host "`n[*] Preparing to remove entries from $targetListName..." -ForegroundColor Cyan

    $normalized = $RawEntries | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique
    $existing = Get-TenantFilterData -TargetAction $TargetAction

    $foundToRemove = @()
    $notFound = @()

    foreach ($entry in $normalized) {
        $match = $existing | Where-Object { $_.Value -eq $entry }
        if ($match) {
            $foundToRemove += $match
        } else {
            $notFound += $entry
        }
    }

    if ($notFound.Count -gt 0) {
        Write-Host "`n[i] The following entries were not found in $targetListName (skipping):" -ForegroundColor Yellow
        foreach ($nf in $notFound) {
            Write-Host "    - $nf" -ForegroundColor Gray
        }
    }

    if ($foundToRemove.Count -eq 0) {
        Write-Host "[!] None of the specified entries exist in $targetListName. Nothing to remove." -ForegroundColor Yellow
        return
    }

    Write-Host "`nPlanned Removal Summary:" -ForegroundColor White
    Write-Host "  Operation   : REMOVE" -ForegroundColor Red
    Write-Host "  Target List : $targetListName" -ForegroundColor White
    Write-Host "  Entries to Remove ($($foundToRemove.Count)): " -ForegroundColor White
    foreach ($item in $foundToRemove) {
        Write-Host "    - $($item.Value) (ID: $($item.Identity))" -ForegroundColor Red
    }

    if (-not $BypassPrompt) {
        $promptQuestion = "Are you sure you want to REMOVE these entries from the $targetListName? / Czy na pewno usunąć te wpisy? (Y/N/T/N)"
        $confirmation = Read-Host $promptQuestion
        if ($confirmation -notmatch "^(Y|y|T|t|Yes|Tak)$") {
            Write-Host "[!] Removal cancelled by user." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without interactive confirmation." -ForegroundColor Cyan
    }

    Write-Host "[*] Removing entries from Microsoft 365 Tenant Allow/Block List..." -ForegroundColor Cyan
    $removeEntriesList = $foundToRemove | ForEach-Object { $_.Value }

    try {
        Remove-TenantAllowBlockListItems -ListType Sender -Entries $removeEntriesList -ErrorAction Stop
        Write-Host "[+] Successfully removed $($foundToRemove.Count) entries from $targetListName!" -ForegroundColor Green
    } catch {
        # Fallback to removal by Ids if removal by Entries encountered issues
        Write-Host "[!] Note: Direct entry removal returned a notice, attempting fallback by entry IDs..." -ForegroundColor Yellow
        try {
            $ids = $foundToRemove | ForEach-Object { $_.Identity }
            Remove-TenantAllowBlockListItems -ListType Sender -Ids $ids -ErrorAction Stop
            Write-Host "[+] Successfully removed entries by ID!" -ForegroundColor Green
        } catch {
            Write-Error "Failed to remove entries from $($targetListName). Details: $_"
            exit 1
        }
    }
}

# CSV Export Helper
function Export-TenantFilterReport {
    param(
        [array]$Items,
        [string]$Path
    )

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "[!] No data available to export to CSV." -ForegroundColor Yellow
        return
    }

    try {
        Write-Host "[*] Exporting $($Items.Count) records to UTF-8 CSV: '$Path'..." -ForegroundColor Cyan
        $Items | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
        Write-Host "[+] CSV report exported successfully to '$Path'." -ForegroundColor Green
    } catch {
        Write-Error "Failed to export CSV report: $_"
    }
}

# Interactive Console Menu
function Invoke-InteractiveMenu {
    param([string]$AdminUPN)

    Connect-ExchangeEnvironment -AdminUPN $AdminUPN

    $exitMenu = $false
    while (-not $exitMenu) {
        Write-Host "`n================================================================================" -ForegroundColor Cyan
        Write-Host "   Microsoft 365 Tenant Filter List Manager - Console Menu" -ForegroundColor White
        Write-Host "   Version: $($ScriptMetadata.Version) | Author: $($ScriptMetadata.Author)" -ForegroundColor Gray
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  1. List Whitelist (Allowed Domains and Email Addresses)" -ForegroundColor Green
        Write-Host "  2. List Blacklist (Blocked Domains and Email Addresses)" -ForegroundColor Red
        Write-Host "  3. List Both Whitelist and Blacklist (Full Tenant Overview)" -ForegroundColor Yellow
        Write-Host "  4. Search in Filter Lists (Domains, Emails, Notes)" -ForegroundColor Cyan
        Write-Host "  5. Add Domain(s) or Email(s) to Whitelist" -ForegroundColor Green
        Write-Host "  6. Add Domain(s) or Email(s) to Blacklist" -ForegroundColor Red
        Write-Host "  7. Remove Entry from Whitelist or Blacklist" -ForegroundColor Magenta
        Write-Host "  8. Export Tenant Filter List to CSV Report" -ForegroundColor White
        Write-Host "  0. Exit" -ForegroundColor DarkGray
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

        $choice = Read-Host "Select an option [0-8]"

        switch ($choice.Trim()) {
            '1' {
                $items = Get-TenantFilterData -TargetAction 'Allow'
                Show-TenantFilterSummaryTable -Items $items -Title "Tenant Whitelist (Allow List)"
            }
            '2' {
                $items = Get-TenantFilterData -TargetAction 'Block'
                Show-TenantFilterSummaryTable -Items $items -Title "Tenant Blacklist (Block List)"
            }
            '3' {
                $items = Get-TenantFilterData -TargetAction 'All'
                Show-TenantFilterSummaryTable -Items $items -Title "Tenant Filter List (All Records)"
            }
            '4' {
                $searchQuery = Read-Host "Enter domain, email address, or note to search"
                if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {
                    Invoke-SearchOperation -Query $searchQuery -TargetAction 'All'
                }
            }
            '5' {
                $rawInput = Read-Host "Enter domains or email addresses separated by commas (e.g. partner.com, user@partner.com)"
                if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                    $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $noteInput = Read-Host "Enter administrative note / comment (optional)"
                    Invoke-AddOperation -RawEntries $entryList -TargetAction 'Allow' -AdminNotes $noteInput
                }
            }
            '6' {
                $rawInput = Read-Host "Enter domains or email addresses separated by commas (e.g. badspammer.com, evil@hack.org)"
                if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                    $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $noteInput = Read-Host "Enter administrative note / reason for block (optional)"
                    Invoke-AddOperation -RawEntries $entryList -TargetAction 'Block' -AdminNotes $noteInput
                }
            }
            '7' {
                $subChoice = Read-Host "Remove from (1) Whitelist or (2) Blacklist? [1/2]"
                $targetAction = if ($subChoice.Trim() -eq '1') { 'Allow' } elseif ($subChoice.Trim() -eq '2') { 'Block' } else { $null }
                if ($targetAction) {
                    $rawInput = Read-Host "Enter domain or email address to remove"
                    if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                        $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        Invoke-RemoveOperation -RawEntries $entryList -TargetAction $targetAction
                    }
                } else {
                    Write-Host "[!] Invalid selection." -ForegroundColor Yellow
                }
            }
            '8' {
                $csvPath = Read-Host "Enter destination CSV path (e.g. C:\Reports\TenantFilterList.csv)"
                if (-not [string]::IsNullOrWhiteSpace($csvPath)) {
                    $items = Get-TenantFilterData -TargetAction 'All'
                    Export-TenantFilterReport -Items $items -Path $csvPath
                }
            }
            '0' {
                Write-Host "Exiting interactive menu. Goodbye!" -ForegroundColor Cyan
                $exitMenu = $true
            }
            default {
                Write-Host "[!] Invalid selection. Please enter a number between 0 and 8." -ForegroundColor Yellow
            }
        }
    }
}

# Main Execution Controller
if ($Interactive) {
    Invoke-InteractiveMenu -AdminUPN $AdminUserPrincipalName
    exit 0
}

# Connect to Exchange Online
Connect-ExchangeEnvironment -AdminUPN $AdminUserPrincipalName

# Determine active scope
$scopeAction = if ($Whitelist) { 'Allow' } elseif ($Blacklist) { 'Block' } else { 'All' }

# Route to specified action
if ($Add) {
    Invoke-AddOperation `
        -RawEntries $Entries `
        -TargetAction $scopeAction `
        -AdminNotes $Notes `
        -NeverExpire:$NoExpiration `
        -DaysToExpire $ExpirationDays `
        -ExplicitExpireDate $ExpirationDate `
        -BypassPrompt:$Force
    exit 0
}

if ($Remove) {
    Invoke-RemoveOperation `
        -RawEntries $Entries `
        -TargetAction $scopeAction `
        -BypassPrompt:$Force
    exit 0
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    Invoke-SearchOperation `
        -Query $Search `
        -TargetAction $scopeAction `
        -ShowDetails:$Detailed `
        -CsvPath $ExportCsv `
        -PipelineOut:$PassThru
    exit 0
}

if ($List -or ($PSCmdlet.ParameterSetName -eq 'List') -or ($PSBoundParameters.ContainsKey('Whitelist') -or $PSBoundParameters.ContainsKey('Blacklist'))) {
    $items = Get-TenantFilterData -TargetAction $scopeAction
    $listTitle = if ($scopeAction -eq 'Allow') { "Tenant Whitelist (Allow List)" }
                 elseif ($scopeAction -eq 'Block') { "Tenant Blacklist (Block List)" }
                 else { "Tenant Allow/Block List (All Records)" }

    if ($Detailed) {
        Show-TenantFilterDetailedCards -Items $items
    } else {
        Show-TenantFilterSummaryTable -Items $items -Title $listTitle
    }

    if (-not [string]::IsNullOrWhiteSpace($ExportCsv) -and $items.Count -gt 0) {
        Export-TenantFilterReport -Items $items -Path $ExportCsv
    }

    if ($PassThru) {
        $items
    }
    exit 0
}

# Fallback to help
Show-ScriptHelp
exit 0
