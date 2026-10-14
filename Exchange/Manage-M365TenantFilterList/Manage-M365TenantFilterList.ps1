<#
.SYNOPSIS
    Manages Microsoft 365 Allowed and Blocked senders and domains in Exchange Online (Anti-Spam inbound policies and Tenant Allow/Block Lists).

.DESCRIPTION
    Comprehensive, enterprise-grade PowerShell automation script designed for Microsoft 365 and Exchange Online
    administrators to audit, search, add, and remove sender domains and email addresses across:
    1. Anti-Spam Inbound Policies (HostedContentFilterPolicy) - "Allowed and blocked senders and domains":
       - Allowed senders (AllowedSenders - individual email addresses)
       - Allowed domains (AllowedSenderDomains - trusted domain names)
       - Blocked senders (BlockedSenders - blocked email addresses)
       - Blocked domains (BlockedSenderDomains - blocked domain names)
    2. Tenant Allow/Block List (TABL - TenantAllowBlockListItems) for sender domains and email addresses.

    Key capabilities include:
    1. Automatic Exchange Online administrator session detection and modern interactive login.
    2. Whitelist (Allow List) Management:
       - Displays summary and complete listing of Allowed senders and Allowed domains (e.g. 208 domains, 8 senders).
       - Add trusted domains or specific email addresses to Anti-Spam policies and/or TABL.
       - Search specifically within allowed senders and domains.
    3. Blacklist (Block List) Management:
       - Displays summary and complete listing of Blocked senders and Blocked domains.
       - Add malicious, spam, or phishing domains and email addresses tenant-wide.
       - Search specifically within blocked senders and domains.
    4. Advanced Auditing & Forensics:
       - Matches the exact Microsoft Defender layout:
         Allowed senders, Allowed domains, Blocked senders, Blocked domains.
       - Detailed card view (-Detailed) organized by category.
    5. Search Engine:
       - Fast wildcard search (-Search) across all allowed and blocked domains, email addresses, and notes.
    6. Safety & Conflict Protection:
       - Pre-flight RFC format validation for domain names and email addresses.
       - Conflict detection: warns if an entry already exists in the opposing list.
       - Bilingual interactive confirmation prompts (English Y/N, Polish T/N) with silent bypass (-Force).
    7. Export & Automation:
       - Export full records to UTF-8 CSV reports (-ExportCsv).
       - Pipeline support (-PassThru) for PowerShell automation.
       - Interactive console menu wizard (-Interactive / -Menu).

.PARAMETER List
    Switch parameter. Lists entries based on scope (-Whitelist, -Blacklist, or All).
    Aliases: -ListEntries, -l, -Show.

.PARAMETER Scope
    Positional string argument specifying the list scope: 'All' (default), 'Whitelist', 'Blacklist', 'Allow', 'Block'.
    Enables commands like: -List Whitelist, -List Blacklist.
    Aliases: -Target, -FilterScope.

.PARAMETER Whitelist
    Switch parameter. Targets or filters operations to the Allow List (Allowed senders and Allowed domains).
    Aliases: -Allow, -AllowList, -w.

.PARAMETER Blacklist
    Switch parameter. Targets or filters operations to the Block List (Blocked senders and Blocked domains).
    Aliases: -Block, -BlockList, -b.

.PARAMETER Category
    Optional category filter: 'All' (default), 'AllowedSenders', 'AllowedDomains', 'BlockedSenders', 'BlockedDomains', 'Domain', 'Email'.
    Aliases: -Cat, -Type.

.PARAMETER PolicyName
    Specifies the Anti-Spam inbound policy to manage. Default is 'Default'.
    Aliases: -Policy.

.PARAMETER Source
    Specifies the backend repository to query or modify: 'All' (default), 'AntiSpam', 'TABL'.
    Aliases: -Backend.

.PARAMETER Add
    Switch parameter. Adds one or more sender domains or email addresses.
    Aliases: -AddEntry, -a.

.PARAMETER Remove
    Switch parameter. Removes one or more entries.
    Aliases: -RemoveEntry, -r.

.PARAMETER Search
    String parameter. Queries for domains, email addresses, or keywords.
    Aliases: -Find, -s, -Query.

.PARAMETER SearchScope
    Positional parameter for Search: 'All' (default), 'Whitelist', 'Blacklist', 'Allow', 'Block'.

.PARAMETER Entries
    One or more domains (e.g., "partner.com") or email addresses (e.g., "user@badactor.org") to manage.
    Aliases: -Entry, -Domain, -Domains, -Address, -Addresses, -Identity.

.PARAMETER Notes
    Optional administrative justification or comment associated with added entries.
    Aliases: -Comment, -Description.

.PARAMETER SyncToTABL
    Switch parameter. When adding or removing entries on the Anti-Spam policy, also synchronizes them to TABL.

.PARAMETER NoExpiration
    Switch parameter. Configures entries added to TABL never to expire automatically.
    Aliases: -NeverExpire.

.PARAMETER ExpirationDays
    Integer parameter (1-730). Sets entry expiration days from today for TABL items.
    Aliases: -Days.

.PARAMETER ExpirationDate
    DateTime parameter. Sets an explicit UTC expiration date for TABL items.
    Aliases: -ExpireAt.

.PARAMETER Detailed
    Switch parameter. Displays in-depth cards and categorized sections.
    Aliases: -d, -Details.

.PARAMETER ExportCsv
    Optional file path to export retrieved records into a UTF-8 CSV report file.
    Aliases: -CsvPath, -Csv, -Export.

.PARAMETER PassThru
    Switch parameter. Outputs custom PowerShell objects to the pipeline.

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
    .\Manage-M365TenantFilterList.ps1 -List Whitelist -Detailed
    Lists all Allowed senders and Allowed domains in the tenant Anti-Spam policy with categorized breakdowns.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List -Whitelist
    Displays all Allowed senders and Allowed domains in a formatted summary table.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List Blacklist -Detailed
    Lists all Blocked senders and Blocked domains configured in the tenant.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List All
    Displays the complete overview of Allowed and blocked senders and domains (senders, domains, counts).

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Search "partner.com"
    Searches for "partner.com" across Allowed domains, Allowed senders, Blocked domains, and TABL.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Search "partner.com" Whitelist
    Searches for "partner.com" specifically within Allowed senders and Allowed domains.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Add Whitelist "partner.com" -Notes "Approved B2B vendor"
    Adds partner.com to Allowed domains in the Default Anti-Spam policy.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Add Whitelist "alerts@supplier.com"
    Adds alerts@supplier.com to Allowed senders in the Default Anti-Spam policy.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Add Blacklist "spammer.com" -Force
    Adds spammer.com to Blocked domains silently without confirmation prompts.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Remove Whitelist "oldpartner.com"
    Removes oldpartner.com from Allowed domains in the Anti-Spam policy.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -List Whitelist -ExportCsv "C:\Reports\M365_AllowedSendersAndDomains.csv"
    Audits Allowed senders and Allowed domains and exports records to a UTF-8 CSV report.

.EXAMPLE
    .\Manage-M365TenantFilterList.ps1 -Interactive
    Launches an interactive, menu-driven console wizard.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.1.0
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

    [Parameter(ParameterSetName = 'Search', Position = 0)]
    [Alias('Find', 's', 'Query')]
    [string]$Search,

    [Parameter(ParameterSetName = 'Search', Position = 1)]
    [ValidateSet('All', 'Whitelist', 'Blacklist', 'Allow', 'Block')]
    [string]$SearchScope = 'All',

    [Parameter(ParameterSetName = 'Interactive')]
    [Alias('Menu', 'i')]
    [switch]$Interactive,

    # List Scope (Positional argument or switches)
    [Parameter(ParameterSetName = 'List', Position = 0)]
    [Parameter(ParameterSetName = 'Default', Position = 0)]
    [ValidateSet('All', 'Whitelist', 'Blacklist', 'Allow', 'Block')]
    [Alias('Target', 'FilterScope')]
    [string]$Scope,

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

    # Policy and Source selection
    [Parameter()]
    [Alias('Policy')]
    [string]$PolicyName = 'Default',

    [Parameter()]
    [ValidateSet('All', 'AntiSpam', 'TABL')]
    [Alias('Backend')]
    [string]$Source = 'All',

    [Parameter()]
    [ValidateSet('All', 'AllowedSenders', 'AllowedDomains', 'BlockedSenders', 'BlockedDomains', 'Domain', 'Email', 'Senders', 'Domains')]
    [Alias('Cat', 'Type')]
    [string]$Category = 'All',

    # Target Entries (Domains / Email addresses)
    [Parameter(ParameterSetName = 'Add', Mandatory = $false, Position = 0, ValueFromRemainingArguments = $true)]
    [Parameter(ParameterSetName = 'Remove', Mandatory = $false, Position = 0, ValueFromRemainingArguments = $true)]
    [Alias('Entry', 'Domain', 'Domains', 'Address', 'Addresses', 'Identity')]
    [string[]]$Entries,

    # Entry Attributes
    [Parameter(ParameterSetName = 'Add')]
    [Alias('Comment', 'Description')]
    [string]$Notes,

    [Parameter(ParameterSetName = 'Add')]
    [switch]$SyncToTABL,

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
    Version = "1.1.0"
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
        "    addresses across Anti-Spam Inbound Policies and Tenant Allow/Block Lists:",
        "      - Allowed senders (AllowedSenders)",
        "      - Allowed domains (AllowedSenderDomains)",
        "      - Blocked senders (BlockedSenders)",
        "      - Blocked domains (BlockedSenderDomains)",
        "      - Tenant Allow/Block List (TABL)",
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
        "    2. List Whitelist (Allowed senders and Allowed domains):",
        "       .\Manage-M365TenantFilterList.ps1 -List Whitelist -Detailed",
        "       .\Manage-M365TenantFilterList.ps1 -List -Whitelist",
        "       .\Manage-M365TenantFilterList.ps1 -List Whitelist -Category AllowedDomains",
        "",
        "    3. List Blacklist (Blocked senders and Blocked domains):",
        "       .\Manage-M365TenantFilterList.ps1 -List Blacklist -Detailed",
        "       .\Manage-M365TenantFilterList.ps1 -List -Blacklist",
        "",
        "    4. List Full Tenant Overview (All senders, domains, and counts):",
        "       .\Manage-M365TenantFilterList.ps1 -List All",
        "       .\Manage-M365TenantFilterList.ps1 -List",
        "",
        "    5. Search for a Domain or Email Address:",
        "       .\Manage-M365TenantFilterList.ps1 -Search `"partner.com`"",
        "       .\Manage-M365TenantFilterList.ps1 -Search `"partner.com`" Whitelist",
        "",
        "    6. Add Domain to Allowed Domains (Anti-Spam Inbound Policy):",
        "       .\Manage-M365TenantFilterList.ps1 -Add Whitelist `"partner.com`"",
        "       .\Manage-M365TenantFilterList.ps1 -Add -Whitelist -Entries `"partner.com`" -Notes `"Approved vendor`"",
        "",
        "    7. Add Email Address to Allowed Senders:",
        "       .\Manage-M365TenantFilterList.ps1 -Add Whitelist `"alerts@partner.com`"",
        "",
        "    8. Add Domain or Email to Blocked List:",
        "       .\Manage-M365TenantFilterList.ps1 -Add Blacklist `"spam.com`" -Force",
        "       .\Manage-M365TenantFilterList.ps1 -Add Blacklist `"phisher@evil.org`"",
        "",
        "    9. Remove Entry from Whitelist or Blacklist:",
        "       .\Manage-M365TenantFilterList.ps1 -Remove Whitelist `"oldpartner.com`"",
        "       .\Manage-M365TenantFilterList.ps1 -Remove Blacklist `"spam.com`" -Force",
        "",
        "   10. Export Tenant Filter List to UTF-8 CSV Report:",
        "       .\Manage-M365TenantFilterList.ps1 -List Whitelist -ExportCsv `"C:\Reports\AllowedSendersAndDomains.csv`"",
        "       .\Manage-M365TenantFilterList.ps1 -List -ExportCsv `"C:\Reports\AllTenantFilters.csv`"",
        "",
        "   11. Interactive Console Menu Wizard:",
        "       .\Manage-M365TenantFilterList.ps1 -Interactive",
        "",
        "PARAMETERS:",
        "    -List, -l                  Lists entries from Whitelist, Blacklist, or both (e.g. -List Whitelist).",
        "    -Scope                     Scope: 'All' (default), 'Whitelist', 'Blacklist', 'Allow', 'Block'.",
        "    -Category, -Cat, -Type     Category filter: 'All', 'AllowedSenders', 'AllowedDomains',",
        "                               'BlockedSenders', 'BlockedDomains', 'Domain', 'Email'.",
        "    -PolicyName, -Policy       Target Anti-Spam policy name (default: 'Default').",
        "    -Source, -Backend          Data source: 'All' (default), 'AntiSpam', 'TABL'.",
        "    -Add, -a                   Adds specified entries (e.g. -Add Whitelist domain.com).",
        "    -Remove, -r                Removes specified entries from the chosen list.",
        "    -Search, -s                Finds entries matching a domain, email, or notes pattern.",
        "    -Whitelist, -Allow, -w     Specifies the Allow List (Allowed senders and Allowed domains).",
        "    -Blacklist, -Block, -b     Specifies the Block List (Blocked senders and Blocked domains).",
        "    -Entries, -Domain, -Entry  One or more domains or email addresses to manage.",
        "    -Notes, -Comment           Administrative justification / comment.",
        "    -SyncToTABL                Also synchronizes additions/removals with TABL.",
        "    -Detailed, -d              Displays comprehensive cards and categorized sections.",
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
$hasExplicitAction = $List -or $Add -or $Remove -or (-not [string]::IsNullOrWhiteSpace($Search)) -or $Interactive -or ($Entries -and $Entries.Count -gt 0) -or (-not [string]::IsNullOrWhiteSpace($Scope)) -or $Whitelist -or $Blacklist
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

# Handle positional scope argument when passed with -Add or -Remove (e.g. -Add Whitelist domain.com)
if (($Add -or $Remove) -and (-not $Whitelist) -and (-not $Blacklist) -and $Entries -and $Entries.Count -ge 1) {
    if ($Entries[0] -match '^(whitelist|allow)$') {
        $Whitelist = [switch]::new($true)
        $Entries = if ($Entries.Count -gt 1) { $Entries[1..($Entries.Count - 1)] } else { @() }
    } elseif ($Entries[0] -match '^(blacklist|block)$') {
        $Blacklist = [switch]::new($true)
        $Entries = if ($Entries.Count -gt 1) { $Entries[1..($Entries.Count - 1)] } else { @() }
    }
}

# Validate scope consistency
if ($Whitelist -and $Blacklist) {
    if ($Add -or $Remove) {
        Write-Error "Conflicting parameters: You cannot specify both -Whitelist and -Blacklist when adding or removing entries."
        exit 1
    }
}

if ($Add -or $Remove) {
    if (-not $Whitelist -and -not $Blacklist) {
        Write-Error "Missing target list: You must specify either -Whitelist or -Blacklist (e.g. -Add Whitelist domain.com or -Add -Whitelist domain.com)."
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

# Comprehensive Data Retrieval Helper (Anti-Spam Policy + TABL)
function Get-TenantFilterData {
    param(
        [ValidateSet('All', 'Allow', 'Block')]
        [string]$TargetAction = 'All',
        [string]$TargetPolicy = 'Default',
        [ValidateSet('All', 'AntiSpam', 'TABL')]
        [string]$TargetSource = 'All'
    )

    $processed = @()

    # 1. Query Anti-Spam Inbound Policies (HostedContentFilterPolicy)
    if ($TargetSource -in @('All', 'AntiSpam')) {
        Write-Host "[*] Querying Microsoft 365 Anti-Spam inbound policies (Allowed and blocked senders and domains)..." -ForegroundColor Cyan
        try {
            $policies = @()
            if (-not [string]::IsNullOrWhiteSpace($TargetPolicy) -and $TargetPolicy -ne 'All') {
                $p = Get-HostedContentFilterPolicy -Identity $TargetPolicy -ErrorAction SilentlyContinue
                if ($p) { $policies = @($p) }
            }
            if ($policies.Count -eq 0) {
                $policies = Get-HostedContentFilterPolicy -ErrorAction Stop
            }

            foreach ($pol in $policies) {
                $polName = [string]$pol.Name

                # Allowed senders (email addresses)
                if ($TargetAction -ne 'Block' -and $pol.AllowedSenders) {
                    foreach ($s in $pol.AllowedSenders) {
                        if (-not [string]::IsNullOrWhiteSpace($s)) {
                            $cleanS = $s.ToString().Trim().ToLowerInvariant()
                            $processed += [PSCustomObject]@{
                                Identity         = "$polName\AllowedSenders\$cleanS"
                                Value            = $cleanS
                                Action           = 'Allow'
                                Category         = 'Allowed senders'
                                ListType         = 'Sender'
                                EntryType        = 'EmailAddress'
                                Source           = "Anti-Spam Policy ($polName)"
                                PolicyName       = $polName
                                Status           = 'Active'
                                ExpirationDate   = 'Never'
                                RemainingDays    = 'Never'
                                LastModifiedDate = 'N/A'
                                LastModifiedBy   = 'Policy'
                                Notes            = "Anti-spam inbound policy: $polName"
                            }
                        }
                    }
                }

                # Allowed domains
                if ($TargetAction -ne 'Block' -and $pol.AllowedSenderDomains) {
                    foreach ($d in $pol.AllowedSenderDomains) {
                        if (-not [string]::IsNullOrWhiteSpace($d)) {
                            $cleanD = $d.ToString().Trim().ToLowerInvariant()
                            $processed += [PSCustomObject]@{
                                Identity         = "$polName\AllowedSenderDomains\$cleanD"
                                Value            = $cleanD
                                Action           = 'Allow'
                                Category         = 'Allowed domains'
                                ListType         = 'Sender'
                                EntryType        = 'Domain'
                                Source           = "Anti-Spam Policy ($polName)"
                                PolicyName       = $polName
                                Status           = 'Active'
                                ExpirationDate   = 'Never'
                                RemainingDays    = 'Never'
                                LastModifiedDate = 'N/A'
                                LastModifiedBy   = 'Policy'
                                Notes            = "Anti-spam inbound policy: $polName"
                            }
                        }
                    }
                }

                # Blocked senders (email addresses)
                if ($TargetAction -ne 'Allow' -and $pol.BlockedSenders) {
                    foreach ($s in $pol.BlockedSenders) {
                        if (-not [string]::IsNullOrWhiteSpace($s)) {
                            $cleanS = $s.ToString().Trim().ToLowerInvariant()
                            $processed += [PSCustomObject]@{
                                Identity         = "$polName\BlockedSenders\$cleanS"
                                Value            = $cleanS
                                Action           = 'Block'
                                Category         = 'Blocked senders'
                                ListType         = 'Sender'
                                EntryType        = 'EmailAddress'
                                Source           = "Anti-Spam Policy ($polName)"
                                PolicyName       = $polName
                                Status           = 'Active'
                                ExpirationDate   = 'Never'
                                RemainingDays    = 'Never'
                                LastModifiedDate = 'N/A'
                                LastModifiedBy   = 'Policy'
                                Notes            = "Anti-spam inbound policy: $polName"
                            }
                        }
                    }
                }

                # Blocked domains
                if ($TargetAction -ne 'Allow' -and $pol.BlockedSenderDomains) {
                    foreach ($d in $pol.BlockedSenderDomains) {
                        if (-not [string]::IsNullOrWhiteSpace($d)) {
                            $cleanD = $d.ToString().Trim().ToLowerInvariant()
                            $processed += [PSCustomObject]@{
                                Identity         = "$polName\BlockedSenderDomains\$cleanD"
                                Value            = $cleanD
                                Action           = 'Block'
                                Category         = 'Blocked domains'
                                ListType         = 'Sender'
                                EntryType        = 'Domain'
                                Source           = "Anti-Spam Policy ($polName)"
                                PolicyName       = $polName
                                Status           = 'Active'
                                ExpirationDate   = 'Never'
                                RemainingDays    = 'Never'
                                LastModifiedDate = 'N/A'
                                LastModifiedBy   = 'Policy'
                                Notes            = "Anti-spam inbound policy: $polName"
                            }
                        }
                    }
                }
            }
        } catch {
            Write-Warning "Could not query Anti-Spam inbound policies: $_"
        }
    }

    # 2. Query Tenant Allow/Block List (TABL)
    if ($TargetSource -in @('All', 'TABL')) {
        Write-Host "[*] Querying Microsoft 365 Tenant Allow/Block List (TABL - Senders)..." -ForegroundColor Cyan
        try {
            $rawItems = @()
            if ($TargetAction -eq 'Allow') {
                $rawItems = Get-TenantAllowBlockListItems -ListType Sender -Allow -ErrorAction SilentlyContinue
            } elseif ($TargetAction -eq 'Block') {
                $rawItems = Get-TenantAllowBlockListItems -ListType Sender -Block -ErrorAction SilentlyContinue
            } else {
                $rawItems = Get-TenantAllowBlockListItems -ListType Sender -ErrorAction SilentlyContinue
            }

            $nowUtc = (Get-Date).ToUniversalTime()

            foreach ($item in $rawItems) {
                $val = [string]$item.Value
                $entryType = if ($val -match '@') { 'EmailAddress' } else { 'Domain' }
                $action = [string]$item.Action
                if ([string]::IsNullOrWhiteSpace($action)) {
                    $action = if ($item.Block) { 'Block' } else { 'Allow' }
                }

                $cat = if ($action -eq 'Allow') {
                    if ($entryType -eq 'EmailAddress') { 'TABL Allowed sender' } else { 'TABL Allowed domain' }
                } else {
                    if ($entryType -eq 'EmailAddress') { 'TABL Blocked sender' } else { 'TABL Blocked domain' }
                }

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
                        # Fallback
                    }
                }

                $processed += [PSCustomObject]@{
                    Identity         = [string]$item.Identity
                    Value            = $val.Trim().ToLowerInvariant()
                    Action           = $action
                    Category         = $cat
                    ListType         = 'Sender'
                    EntryType        = $entryType
                    Source           = 'Tenant Allow/Block List (TABL)'
                    PolicyName       = 'TABL'
                    Status           = $status
                    ExpirationDate   = if ($expDate -and $expDate.Year -lt 9999) { $expDate.ToString("yyyy-MM-dd HH:mm UTC") } else { "Never" }
                    RemainingDays    = $remainingDays
                    LastModifiedDate = if ($item.LastModifiedDate) { ([datetime]$item.LastModifiedDate).ToString("yyyy-MM-dd HH:mm UTC") } else { "N/A" }
                    LastModifiedBy   = [string]$item.LastModifiedBy
                    Notes            = [string]$item.Notes
                }
            }
        } catch {
            Write-Warning "Could not query Tenant Allow/Block List: $_"
        }
    }

    return $processed
}

# Apply Category Filtering
function Filter-ByCategory {
    param(
        [array]$Items,
        [string]$CategoryFilter
    )

    if ([string]::IsNullOrWhiteSpace($CategoryFilter) -or $CategoryFilter -eq 'All') {
        return $Items
    }

    switch ($CategoryFilter.ToLowerInvariant()) {
        'allowedsenders' { return ($Items | Where-Object { $_.Category -eq 'Allowed senders' }) }
        'alloweddomains' { return ($Items | Where-Object { $_.Category -eq 'Allowed domains' }) }
        'blockedsenders' { return ($Items | Where-Object { $_.Category -eq 'Blocked senders' }) }
        'blockeddomains' { return ($Items | Where-Object { $_.Category -eq 'Blocked domains' }) }
        'domain'         { return ($Items | Where-Object { $_.EntryType -eq 'Domain' }) }
        'domains'        { return ($Items | Where-Object { $_.EntryType -eq 'Domain' }) }
        'email'          { return ($Items | Where-Object { $_.EntryType -eq 'EmailAddress' }) }
        'senders'        { return ($Items | Where-Object { $_.EntryType -eq 'EmailAddress' }) }
        default          { return $Items }
    }
}

# Display Tables and Cards
function Show-TenantFilterSummaryTable {
    param(
        [array]$Items,
        [string]$Title = "Allowed and blocked senders and domains",
        [string]$Policy = "Default"
    )

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "`n[i] No entries found for $Title." -ForegroundColor Yellow
        return
    }

    $allowedSenders = ($Items | Where-Object { $_.Category -eq 'Allowed senders' })
    $allowedDomains = ($Items | Where-Object { $_.Category -eq 'Allowed domains' })
    $blockedSenders = ($Items | Where-Object { $_.Category -eq 'Blocked senders' })
    $blockedDomains = ($Items | Where-Object { $_.Category -eq 'Blocked domains' })
    $tablAllow = ($Items | Where-Object { $_.Source -like '*TABL*' -and $_.Action -eq 'Allow' })
    $tablBlock = ($Items | Where-Object { $_.Source -like '*TABL*' -and $_.Action -eq 'Block' })

    Write-Host "`n================================================================================" -ForegroundColor White
    Write-Host "  Allowed and blocked senders and domains (Policy: $Policy)" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor White
    Write-Host "  Allowed senders : $(if ($allowedSenders.Count -gt 0) { "$($allowedSenders.Count) senders" } else { "-" })" -ForegroundColor Green
    Write-Host "  Allowed domains : $(if ($allowedDomains.Count -gt 0) { "$($allowedDomains.Count) domains" } else { "-" })" -ForegroundColor Green
    Write-Host "  Blocked senders : $(if ($blockedSenders.Count -gt 0) { "$($blockedSenders.Count) senders" } else { "-" })" -ForegroundColor Red
    Write-Host "  Blocked domains : $(if ($blockedDomains.Count -gt 0) { "$($blockedDomains.Count) domains" } else { "-" })" -ForegroundColor Red
    if ($tablAllow.Count -gt 0 -or $tablBlock.Count -gt 0) {
        Write-Host "  Tenant Allow/Block List (TABL) : $($tablAllow.Count) allowed | $($tablBlock.Count) blocked" -ForegroundColor Yellow
    }
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Total entries displayed: $($Items.Count)`n" -ForegroundColor Gray

    $tableData = $Items | Select-Object `
        @{Name='Action'; Expression={ $_.Action.ToUpper() }},
        @{Name='Category'; Expression={ $_.Category }},
        @{Name='Value (Domain/Email)'; Expression={ $_.Value }},
        @{Name='Source'; Expression={ $_.Source }},
        @{Name='Status'; Expression={ $_.Status }},
        @{Name='Expires (UTC)'; Expression={ $_.ExpirationDate }},
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
    Write-Host "  Allowed and Blocked Senders and Domains - Detailed Report ($($Items.Count) entries)" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor White

    $grouped = $Items | Group-Object -Property Category

    foreach ($grp in $grouped) {
        $categoryColor = if ($grp.Name -match 'Allow') { 'Green' } else { 'Red' }
        Write-Host "`n>>> $($grp.Name) ($($grp.Count) entries):" -ForegroundColor $categoryColor
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

        $counter = 1
        foreach ($item in $grp.Group) {
            Write-Host ("{0,4}. " -f $counter) -NoNewline -ForegroundColor Gray
            Write-Host "$($item.Value)" -NoNewline -ForegroundColor White
            Write-Host "  [Source: $($item.Source) | Status: $($item.Status)]" -ForegroundColor Gray
            $counter++
        }
    }
    Write-Host "--------------------------------------------------------------------------------`n" -ForegroundColor DarkGray
}

# Search Operation
function Invoke-SearchOperation {
    param(
        [string]$Query,
        [string]$TargetAction,
        [string]$TargetPolicy,
        [string]$TargetSource,
        [string]$CategoryFilter,
        [switch]$ShowDetails,
        [string]$CsvPath,
        [switch]$PipelineOut
    )

    $items = Get-TenantFilterData -TargetAction $TargetAction -TargetPolicy $TargetPolicy -TargetSource $TargetSource
    $items = Filter-ByCategory -Items $items -CategoryFilter $CategoryFilter

    $cleanQuery = $Query.Trim()
    Write-Host "[*] Searching for '$cleanQuery' across senders, domains, and policies..." -ForegroundColor Cyan

    $matches = $items | Where-Object {
        $_.Value -like "*$cleanQuery*" -or
        $_.Category -like "*$cleanQuery*" -or
        $_.Notes -like "*$cleanQuery*" -or
        $_.Source -like "*$cleanQuery*"
    }

    $title = "Search Results for '$cleanQuery'"

    if ($ShowDetails) {
        Show-TenantFilterDetailedCards -Items $matches
    } else {
        Show-TenantFilterSummaryTable -Items $matches -Title $title -Policy $TargetPolicy
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
        [string]$TargetPolicy,
        [string]$AdminNotes,
        [switch]$SyncTABL,
        [switch]$NeverExpire,
        [int]$DaysToExpire,
        [datetime]$ExplicitExpireDate,
        [switch]$BypassPrompt
    )

    $targetListName = if ($TargetAction -eq 'Allow') { "Allowed senders/domains (Whitelist)" } else { "Blocked senders/domains (Blacklist)" }
    Write-Host "`n[*] Preparing to add entries to $targetListName..." -ForegroundColor Cyan

    # 1. Validate formats
    $validatedEntries = @()
    $invalidEntries = @()

    foreach ($e in $RawEntries) {
        $valResult = Test-ValidDomainOrEmail -Entry $e
        if ($valResult.IsValid) {
            $validatedEntries += $valResult
        } else {
            $invalidEntries += $valResult
        }
    }

    if ($invalidEntries.Count -gt 0) {
        Write-Host "`n[!] Invalid domain or email address format detected:" -ForegroundColor Red
        foreach ($inv in $invalidEntries) {
            Write-Host "    - $($inv.CleanValue) : $($inv.Reason)" -ForegroundColor Yellow
        }
    }

    if ($validatedEntries.Count -eq 0) {
        Write-Error "No valid domains or email addresses provided to add."
        exit 1
    }

    # 2. Check duplicates and conflicts
    $existing = Get-TenantFilterData -TargetAction 'All' -TargetPolicy $TargetPolicy
    $alreadyInTarget = @()
    $conflictOpposite = @()
    $finalToAdd = @()

    foreach ($entryObj in $validatedEntries) {
        $cleanVal = $entryObj.CleanValue
        $match = $existing | Where-Object { $_.Value -eq $cleanVal }

        if ($match) {
            if ($match.Action -eq $TargetAction) {
                $alreadyInTarget += $cleanVal
            } else {
                $conflictOpposite += [PSCustomObject]@{
                    Entry          = $cleanVal
                    CurrentAction  = $match.Action
                    ProposedAction = $TargetAction
                    CurrentCat     = $match.Category
                }
            }
        } else {
            $finalToAdd += $entryObj
        }
    }

    if ($alreadyInTarget.Count -gt 0) {
        Write-Host "`n[i] The following entries already exist in the target list (skipping):" -ForegroundColor Yellow
        foreach ($item in ($alreadyInTarget | Select-Object -Unique)) {
            Write-Host "    - $item" -ForegroundColor Gray
        }
    }

    if ($conflictOpposite.Count -gt 0) {
        Write-Host "`n[!] CONFLICT DETECTED: The following entries currently exist in the OPPOSITE list!" -ForegroundColor Red
        foreach ($c in $conflictOpposite) {
            Write-Host "    - $($c.Entry) (Currently: $($c.CurrentCat) [$($c.CurrentAction)], Proposed: $($c.ProposedAction))" -ForegroundColor Yellow
        }
    }

    if ($finalToAdd.Count -eq 0) {
        Write-Host "`n[+] No new entries to add. Everything is up to date." -ForegroundColor Green
        return
    }

    # 3. Confirmation Prompt
    Write-Host "`nPlanned Action Summary:" -ForegroundColor White
    Write-Host "  Operation     : ADD" -ForegroundColor White
    Write-Host "  Target Policy : $TargetPolicy (Anti-Spam Inbound Policy)" -ForegroundColor White
    Write-Host "  Target Action : $TargetAction" -ForegroundColor (if ($TargetAction -eq 'Allow') { 'Green' } else { 'Red' })
    Write-Host "  Entries ($($finalToAdd.Count)): " -ForegroundColor White
    foreach ($entry in $finalToAdd) {
        $typeLabel = if ($entry.EntryType -eq 'EmailAddress') { 'Sender Email' } else { 'Domain' }
        Write-Host "    + $($entry.CleanValue) ($typeLabel)" -ForegroundColor Cyan
    }

    if (-not $BypassPrompt) {
        $promptQuestion = if ($TargetAction -eq 'Allow') {
            "Are you sure you want to ADD these entries to the Allowed list? / Czy na pewno dodać do dozwolonych? (Y/N/T/N)"
        } else {
            "Are you sure you want to ADD these entries to the Blocked list? / Czy na pewno dodać do zablokowanych? (Y/N/T/N)"
        }

        $confirmation = Read-Host $promptQuestion
        if ($confirmation -notmatch "^(Y|y|T|t|Yes|Tak)$") {
            Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without interactive confirmation." -ForegroundColor Cyan
    }

    # 4. Commit to Anti-Spam Policy (HostedContentFilterPolicy)
    Write-Host "[*] Updating Anti-Spam Policy '$TargetPolicy'..." -ForegroundColor Cyan

    foreach ($entry in $finalToAdd) {
        $val = $entry.CleanValue
        $isEmail = ($entry.EntryType -eq 'EmailAddress')

        try {
            if ($TargetAction -eq 'Allow') {
                if ($isEmail) {
                    Write-Host "    + Adding email to AllowedSenders: $val" -ForegroundColor Green
                    Set-HostedContentFilterPolicy -Identity $TargetPolicy -AllowedSenders @{Add = $val} -ErrorAction Stop
                } else {
                    Write-Host "    + Adding domain to AllowedSenderDomains: $val" -ForegroundColor Green
                    Set-HostedContentFilterPolicy -Identity $TargetPolicy -AllowedSenderDomains @{Add = $val} -ErrorAction Stop
                }
            } else {
                if ($isEmail) {
                    Write-Host "    + Adding email to BlockedSenders: $val" -ForegroundColor Red
                    Set-HostedContentFilterPolicy -Identity $TargetPolicy -BlockedSenders @{Add = $val} -ErrorAction Stop
                } else {
                    Write-Host "    + Adding domain to BlockedSenderDomains: $val" -ForegroundColor Red
                    Set-HostedContentFilterPolicy -Identity $TargetPolicy -BlockedSenderDomains @{Add = $val} -ErrorAction Stop
                }
            }
        } catch {
            Write-Warning "Could not update Anti-Spam policy for '$val': $_"
        }
    }

    Write-Host "[+] Successfully updated Anti-Spam policy '$TargetPolicy'!" -ForegroundColor Green

    # 5. Optional Sync to TABL
    if ($SyncTABL) {
        Write-Host "[*] Synchronizing entries to Tenant Allow/Block List (TABL)..." -ForegroundColor Cyan
        $tablEntries = $finalToAdd | ForEach-Object { $_.CleanValue }
        $cmdArgs = @{
            ListType    = 'Sender'
            Entries     = $tablEntries
            ErrorAction = 'SilentlyContinue'
        }
        if ($TargetAction -eq 'Allow') { $cmdArgs['Allow'] = $true } else { $cmdArgs['Block'] = $true }
        if (-not [string]::IsNullOrWhiteSpace($AdminNotes)) { $cmdArgs['Notes'] = $AdminNotes }

        if ($NeverExpire) {
            $cmdArgs['NoExpiration'] = $true
        } elseif ($ExplicitExpireDate -and $ExplicitExpireDate.Year -gt 1) {
            $cmdArgs['ExpirationDate'] = $ExplicitExpireDate.ToUniversalTime()
        } elseif ($DaysToExpire -gt 0) {
            $cmdArgs['ExpirationDate'] = (Get-Date).AddDays($DaysToExpire).ToUniversalTime()
        } else {
            if ($TargetAction -eq 'Allow') {
                $cmdArgs['ExpirationDate'] = (Get-Date).AddDays(30).ToUniversalTime()
            } else {
                $cmdArgs['NoExpiration'] = $true
            }
        }

        try {
            New-TenantAllowBlockListItems @cmdArgs | Out-Null
            Write-Host "[+] Synchronized $($tablEntries.Count) entries to TABL!" -ForegroundColor Green
        } catch {
            Write-Warning "TABL synchronization notice: $_"
        }
    }
}

# Remove Operation
function Invoke-RemoveOperation {
    param(
        [string[]]$RawEntries,
        [ValidateSet('Allow', 'Block')]
        [string]$TargetAction,
        [string]$TargetPolicy,
        [switch]$BypassPrompt
    )

    $targetListName = if ($TargetAction -eq 'Allow') { "Allowed list (Whitelist)" } else { "Blocked list (Blacklist)" }
    Write-Host "`n[*] Preparing to remove entries from $targetListName..." -ForegroundColor Cyan

    $normalized = $RawEntries | ForEach-Object { $_.Trim().ToLowerInvariant() } | Select-Object -Unique
    $existing = Get-TenantFilterData -TargetAction $TargetAction -TargetPolicy $TargetPolicy

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
        Write-Host "    - $($item.Value) [Category: $($item.Category) | Source: $($item.Source)]" -ForegroundColor Red
    }

    if (-not $BypassPrompt) {
        $promptQuestion = "Are you sure you want to REMOVE these entries? / Czy na pewno usunąć te wpisy? (Y/N/T/N)"
        $confirmation = Read-Host $promptQuestion
        if ($confirmation -notmatch "^(Y|y|T|t|Yes|Tak)$") {
            Write-Host "[!] Removal cancelled by user." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without interactive confirmation." -ForegroundColor Cyan
    }

    # Execute removal from Anti-Spam Policy
    foreach ($item in $foundToRemove) {
        $val = $item.Value
        $isEmail = ($item.EntryType -eq 'EmailAddress')

        if ($item.Source -like '*Anti-Spam*') {
            try {
                if ($item.Action -eq 'Allow') {
                    if ($isEmail) {
                        Write-Host "    - Removing email from AllowedSenders: $val" -ForegroundColor Yellow
                        Set-HostedContentFilterPolicy -Identity $TargetPolicy -AllowedSenders @{Remove = $val} -ErrorAction Stop
                    } else {
                        Write-Host "    - Removing domain from AllowedSenderDomains: $val" -ForegroundColor Yellow
                        Set-HostedContentFilterPolicy -Identity $TargetPolicy -AllowedSenderDomains @{Remove = $val} -ErrorAction Stop
                    }
                } else {
                    if ($isEmail) {
                        Write-Host "    - Removing email from BlockedSenders: $val" -ForegroundColor Yellow
                        Set-HostedContentFilterPolicy -Identity $TargetPolicy -BlockedSenders @{Remove = $val} -ErrorAction Stop
                    } else {
                        Write-Host "    - Removing domain from BlockedSenderDomains: $val" -ForegroundColor Yellow
                        Set-HostedContentFilterPolicy -Identity $TargetPolicy -BlockedSenderDomains @{Remove = $val} -ErrorAction Stop
                    }
                }
            } catch {
                Write-Warning "Could not remove '$val' from Anti-Spam policy: $_"
            }
        } elseif ($item.Source -like '*TABL*') {
            try {
                Remove-TenantAllowBlockListItems -ListType Sender -Entries $val -ErrorAction SilentlyContinue
                Write-Host "    - Removed '$val' from TABL." -ForegroundColor Yellow
            } catch {
                # Fallback
            }
        }
    }

    Write-Host "[+] Removal operation completed." -ForegroundColor Green
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
    param(
        [string]$AdminUPN,
        [string]$TargetPolicy
    )

    Connect-ExchangeEnvironment -AdminUPN $AdminUPN

    $exitMenu = $false
    while (-not $exitMenu) {
        Write-Host "`n================================================================================" -ForegroundColor Cyan
        Write-Host "   Microsoft 365 Allowed and Blocked Senders and Domains Manager" -ForegroundColor White
        Write-Host "   Version: $($ScriptMetadata.Version) | Policy: $TargetPolicy | Author: $($ScriptMetadata.Author)" -ForegroundColor Gray
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  1. Overview: Allowed and blocked senders and domains (Summary Counts)" -ForegroundColor Yellow
        Write-Host "  2. List Whitelist (Allowed senders & Allowed domains)" -ForegroundColor Green
        Write-Host "  3. List Blacklist (Blocked senders & Blocked domains)" -ForegroundColor Red
        Write-Host "  4. List Allowed Domains only (e.g. 208 domains)" -ForegroundColor Green
        Write-Host "  5. List Allowed Senders only (e.g. 8 senders)" -ForegroundColor Green
        Write-Host "  6. Search across all lists (Domains, Emails, Notes)" -ForegroundColor Cyan
        Write-Host "  7. Add Domain(s) or Email(s) to Whitelist" -ForegroundColor Green
        Write-Host "  8. Add Domain(s) or Email(s) to Blacklist" -ForegroundColor Red
        Write-Host "  9. Remove Entry from Whitelist or Blacklist" -ForegroundColor Magenta
        Write-Host " 10. Export All Filter Entries to CSV Report" -ForegroundColor White
        Write-Host "  0. Exit" -ForegroundColor DarkGray
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

        $choice = Read-Host "Select an option [0-10]"

        switch ($choice.Trim()) {
            '1' {
                $items = Get-TenantFilterData -TargetAction 'All' -TargetPolicy $TargetPolicy
                Show-TenantFilterSummaryTable -Items $items -Title "Allowed and blocked senders and domains" -Policy $TargetPolicy
            }
            '2' {
                $items = Get-TenantFilterData -TargetAction 'Allow' -TargetPolicy $TargetPolicy
                Show-TenantFilterDetailedCards -Items $items
            }
            '3' {
                $items = Get-TenantFilterData -TargetAction 'Block' -TargetPolicy $TargetPolicy
                Show-TenantFilterDetailedCards -Items $items
            }
            '4' {
                $items = Get-TenantFilterData -TargetAction 'Allow' -TargetPolicy $TargetPolicy
                $doms = Filter-ByCategory -Items $items -CategoryFilter 'AllowedDomains'
                Show-TenantFilterSummaryTable -Items $doms -Title "Allowed domains" -Policy $TargetPolicy
            }
            '5' {
                $items = Get-TenantFilterData -TargetAction 'Allow' -TargetPolicy $TargetPolicy
                $sends = Filter-ByCategory -Items $items -CategoryFilter 'AllowedSenders'
                Show-TenantFilterSummaryTable -Items $sends -Title "Allowed senders" -Policy $TargetPolicy
            }
            '6' {
                $searchQuery = Read-Host "Enter domain or email address to search"
                if (-not [string]::IsNullOrWhiteSpace($searchQuery)) {
                    Invoke-SearchOperation -Query $searchQuery -TargetAction 'All' -TargetPolicy $TargetPolicy -TargetSource 'All' -CategoryFilter 'All'
                }
            }
            '7' {
                $rawInput = Read-Host "Enter domains or email addresses separated by commas (e.g. partner.com, user@partner.com)"
                if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                    $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $noteInput = Read-Host "Enter administrative note / comment (optional)"
                    Invoke-AddOperation -RawEntries $entryList -TargetAction 'Allow' -TargetPolicy $TargetPolicy -AdminNotes $noteInput
                }
            }
            '8' {
                $rawInput = Read-Host "Enter domains or email addresses separated by commas (e.g. badspammer.com, evil@hack.org)"
                if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                    $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    $noteInput = Read-Host "Enter administrative note / reason for block (optional)"
                    Invoke-AddOperation -RawEntries $entryList -TargetAction 'Block' -TargetPolicy $TargetPolicy -AdminNotes $noteInput
                }
            }
            '9' {
                $subChoice = Read-Host "Remove from (1) Whitelist or (2) Blacklist? [1/2]"
                $targetAction = if ($subChoice.Trim() -eq '1') { 'Allow' } elseif ($subChoice.Trim() -eq '2') { 'Block' } else { $null }
                if ($targetAction) {
                    $rawInput = Read-Host "Enter domain or email address to remove"
                    if (-not [string]::IsNullOrWhiteSpace($rawInput)) {
                        $entryList = $rawInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        Invoke-RemoveOperation -RawEntries $entryList -TargetAction $targetAction -TargetPolicy $TargetPolicy
                    }
                } else {
                    Write-Host "[!] Invalid selection." -ForegroundColor Yellow
                }
            }
            '10' {
                $csvPath = Read-Host "Enter destination CSV path (e.g. C:\Reports\TenantFilterList.csv)"
                if (-not [string]::IsNullOrWhiteSpace($csvPath)) {
                    $items = Get-TenantFilterData -TargetAction 'All' -TargetPolicy $TargetPolicy
                    Export-TenantFilterReport -Items $items -Path $csvPath
                }
            }
            '0' {
                Write-Host "Exiting interactive menu. Goodbye!" -ForegroundColor Cyan
                $exitMenu = $true
            }
            default {
                Write-Host "[!] Invalid selection. Please enter a number between 0 and 10." -ForegroundColor Yellow
            }
        }
    }
}

# Main Execution Controller
if ($Interactive) {
    Invoke-InteractiveMenu -AdminUPN $AdminUserPrincipalName -TargetPolicy $PolicyName
    exit 0
}

# Connect to Exchange Online
Connect-ExchangeEnvironment -AdminUPN $AdminUserPrincipalName

# Determine active scope
$scopeAction = if ($Whitelist -or ($Scope -in @('Whitelist', 'Allow')) -or ($SearchScope -in @('Whitelist', 'Allow'))) { 'Allow' }
               elseif ($Blacklist -or ($Scope -in @('Blacklist', 'Block')) -or ($SearchScope -in @('Blacklist', 'Block'))) { 'Block' }
               else { 'All' }

# Route to specified action
if ($Add) {
    Invoke-AddOperation `
        -RawEntries $Entries `
        -TargetAction $scopeAction `
        -TargetPolicy $PolicyName `
        -AdminNotes $Notes `
        -SyncTABL:$SyncToTABL `
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
        -TargetPolicy $PolicyName `
        -BypassPrompt:$Force
    exit 0
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    Invoke-SearchOperation `
        -Query $Search `
        -TargetAction $scopeAction `
        -TargetPolicy $PolicyName `
        -TargetSource $Source `
        -CategoryFilter $Category `
        -ShowDetails:$Detailed `
        -CsvPath $ExportCsv `
        -PipelineOut:$PassThru
    exit 0
}

if ($List -or ($PSCmdlet.ParameterSetName -eq 'List') -or ($PSBoundParameters.ContainsKey('Whitelist') -or $PSBoundParameters.ContainsKey('Blacklist') -or (-not [string]::IsNullOrWhiteSpace($Scope)))) {
    $items = Get-TenantFilterData -TargetAction $scopeAction -TargetPolicy $PolicyName -TargetSource $Source
    $items = Filter-ByCategory -Items $items -CategoryFilter $Category

    $listTitle = if ($scopeAction -eq 'Allow') { "Allowed senders and domains (Whitelist)" }
                 elseif ($scopeAction -eq 'Block') { "Blocked senders and domains (Blacklist)" }
                 else { "Allowed and blocked senders and domains" }

    if ($Detailed) {
        Show-TenantFilterDetailedCards -Items $items
    } else {
        Show-TenantFilterSummaryTable -Items $items -Title $listTitle -Policy $PolicyName
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
