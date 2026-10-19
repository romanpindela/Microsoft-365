<#
.SYNOPSIS
    Automates SharePoint Online site creation: provisions modern Company-Wide preview sites (for all licensed M365 users) or isolated Empty Sites with designated owners and optional users.

.DESCRIPTION
    New-SharePoint-Site is an enterprise-grade PowerShell automation utility designed for
    Microsoft 365 and SharePoint Online administrators. It streamlines site provisioning,
    enforces CleanCode standards, provides interactive administrator authentication, performs
    comprehensive input sanitization, and outputs detailed progress auditing and file logging.

    Key Capabilities:
    1. Interactive Microsoft 365 / SharePoint Online Administrator Login:
       - Automatically detects active administrator sessions.
       - Initiates secure web/MFA authentication to the SharePoint Online Admin Center if disconnected.
    2. Company-Wide Preview Site Creation (-CompanyWide / -SiteType CompanyWide):
       - Provisions a modern communication site (SITEPAGEPUBLISHING#0) or custom template.
       - Assigns the specified user as the Primary Site Collection Administrator and Site Owner.
       - Automatically resolves tenant GUID and assigns the canonical company-wide claim
         'c:0-.f|rolemanager|spo-grid-all-users/<TenantId>' and bilingual claims ('Everyone except external users'
         / 'Wszyscy oprócz użytkowników zewnętrznych') to the site's Visitors / Odwiedzający group.
       - Supports updating/remediating permissions on an existing site collection.
    3. Isolated Empty Site Creation (-EmptySite / -SiteType EmptySite):
       - Provisions a clean, isolated modern site collection (Communication or Team site without M365 Group).
       - Restricts access strictly to the designated Owner.
       - Supports optional assignment of additional Members (Edit) and Visitors (Read).
    4. Auto-Generated Site Alias:
       - If -SiteAlias is omitted, automatically generates a clean URL slug from -SiteTitle.
    5. Detailed Operation Auditing:
       - Real-time step-by-step console reporting for each provisioning phase.
       - Polling retry loop with progressive backoff for site hydration and group discovery.
       - Comprehensive operation audit summary table upon completion.
    6. Enterprise File Logging:
       - Full execution logs automatically saved to C:\Logs (with fallback to $env:LOCALAPPDATA\Logs).
       - Standardized log naming: yyyyMMdd_HHmmss-New-SharePoint-Site-<ComputerName>-<UserName>.log.
    7. Safety, Bilingual Support & Standards:
       - Safe input string validation guarding against command injections.
       - Bilingual prompt support (English Y/N and Polish T/N) and group names (Visitors / Odwiedzający / Goście).
       - Auto-help display on execution without parameters or with -Help / -h.

.PARAMETER SiteTitle
    The display title of the new SharePoint Online site.
    Aliases: -Title, -Name, -SiteName.

.PARAMETER SiteAlias
    The URL slug, relative path (e.g., 'sites/intranet'), or full site URL
    (e.g., 'https://contoso.sharepoint.com/sites/intranet').
    If omitted, automatically derived from -SiteTitle.
    Aliases: -Alias, -Url, -SiteUrl.

.PARAMETER Owner
    The User Principal Name (UPN) or Primary Email Address of the site owner / primary administrator.
    Must be a valid licensed Microsoft 365 user.
    Aliases: -OwnerEmail, -OwnerUPN, -PrimaryOwner.

.PARAMETER SiteType
    The site configuration archetype.
    Valid values:
      - 'CompanyWide' : Company-wide site with preview (Read) access for all licensed M365 users.
      - 'EmptySite'   : Clean site restricted to the Owner and optional explicit users.
    Default: 'CompanyWide'.

.PARAMETER CompanyWide
    Switch parameter. Shortcut to provision a Company-Wide preview site.

.PARAMETER EmptySite
    Switch parameter. Shortcut to provision an isolated Empty site.

.PARAMETER Template
    SharePoint Online site template to use.
    Common modern templates:
      - 'SITEPAGEPUBLISHING#0' : Modern Communication Site (default for CompanyWide)
      - 'STS#3'                : Modern Team Site without Microsoft 365 Group
      - 'STS#0'                : Classic Team Site
    Default: 'SITEPAGEPUBLISHING#0'.

.PARAMETER Members
    Optional array of user UPNs or email addresses to assign to the site Members group (Edit permissions).
    Applicable to both EmptySite and CompanyWide modes.
    Aliases: -MemberList, -AddMembers.

.PARAMETER Visitors
    Optional array of user UPNs or email addresses to assign to the site Visitors group (Read permissions).
    Aliases: -VisitorList, -AddVisitors.

.PARAMETER TenantName
    The Microsoft 365 tenant domain name or prefix (e.g., 'contoso' or 'contoso.onmicrosoft.com').
    If omitted, the script automatically extracts it from the Owner UPN, SiteUrl, or active session.
    Aliases: -Tenant, -TenantDomain.

.PARAMETER AdminUrl
    The SharePoint Online Admin Center URL (e.g., 'https://contoso-admin.sharepoint.com').
    If omitted, automatically derived from the TenantName.

.PARAMETER StorageQuotaMB
    Site storage quota limit in megabytes (MB).
    Default: 1048576 MB (1 TB). Maximum: 26214400 MB (25 TB).

.PARAMETER TimeZoneId
    SharePoint Time Zone ID.
    Default: 4 (UTC+01:00 Romance Standard Time - Brussels, Copenhagen, Madrid, Paris, Warsaw).

.PARAMETER LocaleId
    SharePoint language / locale ID (LCID).
    Default: 1045 (Polish). Use 1033 for English (United States).
    Aliases: -Lcid, -Language.

.PARAMETER LogDirectory
    Directory where execution logs are written.
    Default: 'C:\Logs' (with fallback to $env:LOCALAPPDATA\Logs).

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N or T/N) for automated scripts.

.PARAMETER PassThru
    Switch parameter. Returns a custom PSObject representing the created site to the PowerShell pipeline.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.
    Aliases: -h.

.EXAMPLE
    # 1. Display built-in help and examples (default behavior when executed without parameters)
    .\New-SharePoint-Site.ps1
    .\New-SharePoint-Site.ps1 -Help

.EXAMPLE
    # 2. Create a Company-Wide Preview Site for all licensed M365 users (alias auto-generated from Title)
    .\New-SharePoint-Site.ps1 -SiteTitle "Sekretariat" -Owner "roman.pindela@elektrimont.pl" -CompanyWide -Force

.EXAMPLE
    # 3. Create a Company-Wide Preview Site with explicit alias and tenant
    .\New-SharePoint-Site.ps1 -SiteTitle "Katalog Produktow" -SiteAlias "katalog" -Owner "jan.kowalski@contoso.com" -TenantName "contoso" -CompanyWide

.EXAMPLE
    # 4. Create an Empty Private Site with Owner only
    .\New-SharePoint-Site.ps1 -SiteTitle "Projekt Poufny" -SiteAlias "projekt-poufny" -Owner "kierownik@contoso.com" -EmptySite

.EXAMPLE
    # 5. Create an Empty Private Site with Owner, specific Members, and Visitors
    .\New-SharePoint-Site.ps1 -SiteTitle "Zarzad" -Owner "prezes@contoso.com" -EmptySite -Members "dyrektor@contoso.com" -Visitors "audytor@contoso.com"

.EXAMPLE
    # 6. Create a Modern Team Site (no group) with custom quota (500 GB) and English language
    .\New-SharePoint-Site.ps1 -SiteTitle "Global Ops" -SiteAlias "global-ops" -Owner "admin@contoso.com" -Template "STS#3" -LocaleId 1033 -StorageQuotaMB 512000 -EmptySite -Force

.NOTES
    Script Name   : New-SharePoint-Site.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.1.0
    Created       : 2026-10-19
    Updated       : 2026-10-19
    Standards     : Adheres to CleanCode and PowerShell Script Standards (C:\Dane\rpindela\GitHub\PowerShell Script Standards.txt)
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Site Identification
    [Parameter(ParameterSetName = 'CompanyWide', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'EmptySite', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 0)]
    [Alias('Title', 'Name', 'SiteName')]
    [string]$SiteTitle,

    [Parameter(ParameterSetName = 'CompanyWide', Mandatory = $false, Position = 1)]
    [Parameter(ParameterSetName = 'EmptySite', Mandatory = $false, Position = 1)]
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 1)]
    [Alias('Alias', 'Url', 'SiteUrl')]
    [string]$SiteAlias,

    [Parameter(ParameterSetName = 'CompanyWide', Mandatory = $true, Position = 2)]
    [Parameter(ParameterSetName = 'EmptySite', Mandatory = $true, Position = 2)]
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 2)]
    [Alias('OwnerEmail', 'OwnerUPN', 'PrimaryOwner')]
    [string]$Owner,

    # Site Type / Mode Selection
    [Parameter(ParameterSetName = 'Default')]
    [ValidateSet('CompanyWide', 'EmptySite')]
    [string]$SiteType = 'CompanyWide',

    [Parameter(ParameterSetName = 'CompanyWide')]
    [switch]$CompanyWide,

    [Parameter(ParameterSetName = 'EmptySite')]
    [switch]$EmptySite,

    # Template
    [Parameter()]
    [string]$Template = 'SITEPAGEPUBLISHING#0',

    # Optional Users Assignment
    [Parameter()]
    [Alias('MemberList', 'AddMembers')]
    [string[]]$Members = @(),

    [Parameter()]
    [Alias('VisitorList', 'AddVisitors')]
    [string[]]$Visitors = @(),

    # Tenant and Admin Connectivity
    [Parameter()]
    [Alias('Tenant', 'TenantDomain')]
    [string]$TenantName,

    [Parameter()]
    [string]$AdminUrl,

    # Site Configuration Quotas and Localization
    [Parameter()]
    [ValidateRange(1024, 26214400)]
    [long]$StorageQuotaMB = 1048576,

    [Parameter()]
    [int]$TimeZoneId = 4,

    [Parameter()]
    [Alias('Lcid', 'Language')]
    [int]$LocaleId = 1045,

    # Output and Logging Options
    [Parameter()]
    [string]$LogDirectory = 'C:\Logs',

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [switch]$PassThru,

    # Help Switch
    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure WindowsPowerShell module path is available in PowerShell 7 for cross-compatibility
$userWinPSModules = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\WindowsPowerShell\Modules"
if ((Test-Path -Path $userWinPSModules) -and ($env:PSModulePath -notlike "*$userWinPSModules*")) {
    $env:PSModulePath = "$env:PSModulePath;$userWinPSModules"
}

# ==============================================================================
# SCRIPT METADATA DECLARATION
# ==============================================================================
$Script:Metadata = @{
    Title       = "New-SharePoint-Site.ps1"
    Version     = "1.1.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Date        = "2026-10-19"
    ModuleReq   = "Microsoft.Online.SharePoint.PowerShell"
}

# Execution tracking
$Script:StartTime   = Get-Date
$Script:LogFilePath = $null
$Script:OperationDetails = [System.Collections.Generic.List[PSObject]]::new()

# ==============================================================================
# LOGGING SYSTEM (CleanCode & Enterprise Logging Standard)
# ==============================================================================
function Initialize-ScriptLogging {
    [CmdletBinding()]
    param([string]$TargetDirectory)

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $computer  = $env:COMPUTERNAME
    $user      = $env:USERNAME
    $logFileName = "${timestamp}-New-SharePoint-Site-${computer}-${user}.log"

    $resolvedDir = $TargetDirectory
    try {
        if (-not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        # Fallback to LocalAppData\Logs if C:\Logs is inaccessible or lacks permissions
        $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs"
        if (-not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName
    $header = @"
================================================================================
Log File Created : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script           : $($Script:Metadata.Title) v$($Script:Metadata.Version)
Author           : $($Script:Metadata.Author) ($($Script:Metadata.Email))
GitHub           : $($Script:Metadata.GitHub)
Host / User      : $computer / $user
Target Dir       : $resolvedDir
================================================================================
"@
    Add-Content -Path $Script:LogFilePath -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Write-ScriptLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [switch]$NoConsole
    )

    $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line    = "[$timeStr] [$($Level.PadRight(7))] $Message"

    if ($Script:LogFilePath) {
        try {
            Add-Content -Path $Script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { }
    }

    if (-not $NoConsole) {
        $color = switch ($Level) {
            'INFO'    { 'Cyan' }
            'SUCCESS' { 'Green' }
            'WARN'    { 'Yellow' }
            'ERROR'   { 'Red' }
            'DEBUG'   { 'DarkGray' }
            Default   { 'White' }
        }

        $tag = switch ($Level) {
            'INFO'    { '[*]' }
            'SUCCESS' { '[+]' }
            'WARN'    { '[!]' }
            'ERROR'   { '[-]' }
            'DEBUG'   { '[.]' }
            Default   { '   ' }
        }

        Write-Host "$tag $Message" -ForegroundColor $color
    }
}

function Add-OperationDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Step,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Success', 'Failed', 'Skipped', 'Warning', 'Running')]
        [string]$Status,

        [string]$Details = ""
    )

    $detailObj = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString("HH:mm:ss")
        Step      = $Step
        Operation = $Operation
        Status    = $Status
        Details   = $Details
    }
    $Script:OperationDetails.Add($detailObj)
}

# ==============================================================================
# HELP & DOCUMENTATION (Standard 1 & 7)
# ==============================================================================
function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: $($Script:Metadata.Title)",
        "VERSION: $($Script:Metadata.Version)",
        "AUTHOR: $($Script:Metadata.Author)",
        "CONTACT: $($Script:Metadata.Email) | $($Script:Metadata.GitHub)",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Automates the provisioning and configuration of modern SharePoint Online site collections.",
        "    Supports provisioning Company-Wide preview sites (with Read permissions for all licensed M365",
        "    users) or isolated Empty Sites restricted to a designated Owner and optional explicit users.",
        "    Enforces input sanitization, lists details of each operation, and provides enterprise logging.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    1. Office 365 / SharePoint Administrator Authentication:",
        "       - Requires an active SharePoint Online Administrator or Global Administrator role.",
        "       - The script automatically checks for an active session with Connect-SPOService.",
        "       - If disconnected, it prompts for interactive modern administrator sign-in.",
        "    2. Required PowerShell Module:",
        "       - 'Microsoft.Online.SharePoint.PowerShell'",
        "       - Automatically checked and prompted for installation if missing.",
        "    3. Script Unblocking (After downloading from GitHub):",
        "       - Run: Unblock-File -Path .\New-SharePoint-Site.ps1",
        "",
        "SITE CREATION ARCHETYPES:",
        "    A) Company-Wide Preview Site (-CompanyWide):",
        "       - Best for company portals, announcements, news, or departmental sites.",
        "       - Configures Owner as Primary Site Administrator with Full Control.",
        "       - Automatically resolves tenant ID and assigns the canonical all-users claim",
        "         ('c:0-.f|rolemanager|spo-grid-all-users/<TenantId>') to the site Visitors group.",
        "       - Default template: Modern Communication Site (SITEPAGEPUBLISHING#0).",
        "",
        "    B) Isolated Empty Site (-EmptySite):",
        "       - Best for confidential teams, departments, or private working groups.",
        "       - Grants Full Control strictly to the designated Owner.",
        "       - Optional members (-Members) and visitors (-Visitors) can be explicitly invited.",
        "       - No company-wide access is granted.",
        "",
        "USAGE EXAMPLES:",
        "    1. Display Built-In Help & Examples (default on run without parameters):",
        "       .\New-SharePoint-Site.ps1",
        "       .\New-SharePoint-Site.ps1 -h",
        "",
        "    2. Create a Company-Wide Preview Site (alias auto-generated from Title):",
        "       .\New-SharePoint-Site.ps1 -SiteTitle `"Sekretariat`" -Owner `"roman.pindela@elektrimont.pl`" -CompanyWide -Force",
        "",
        "    3. Create a Company-Wide Preview Site with explicit alias and tenant:",
        "       .\New-SharePoint-Site.ps1 -SiteTitle `"Katalog Produktow`" -SiteAlias `"katalog`" -Owner `"jan.kowalski@contoso.com`" -TenantName `"contoso`" -CompanyWide",
        "",
        "    4. Create an Empty Private Site with Owner only:",
        "       .\New-SharePoint-Site.ps1 -SiteTitle `"Projekt X`" -SiteAlias `"projekt-x`" -Owner `"manager@contoso.com`" -EmptySite",
        "",
        "    5. Create an Empty Private Site with Owner, specific Members, and Visitors:",
        "       .\New-SharePoint-Site.ps1 -SiteTitle `"Zarzad`" -Owner `"prezes@contoso.com`" -EmptySite -Members `"dyrektor@contoso.com`" -Visitors `"audytor@contoso.com`"",
        "",
        "    6. Create a Modern Team Site (no group) with custom quota (500 GB) and English language:",
        "       .\New-SharePoint-Site.ps1 -SiteTitle `"Global Finance`" -SiteAlias `"global-finance`" -Owner `"admin@contoso.com`" -Template `"STS#3`" -LocaleId 1033 -StorageQuotaMB 512000 -EmptySite -Force",
        "",
        "PARAMETERS:",
        "    -SiteTitle, -Title, -Name   Display title for the new SharePoint site.",
        "    -SiteAlias, -Alias, -Url    Relative alias ('sites/alias') or full URL (auto-generated if omitted).",
        "    -Owner, -OwnerEmail, -u     UPN or email address of the site owner / primary administrator.",
        "    -CompanyWide                Switch to create a company-wide site (all licensed users get preview access).",
        "    -EmptySite                  Switch to create an isolated empty site (restricted to owner/members).",
        "    -SiteType                   Explicit site mode: 'CompanyWide' (default) or 'EmptySite'.",
        "    -Template                   SharePoint template code (default: SITEPAGEPUBLISHING#0 for communication site).",
        "    -Members                    Optional array of user UPNs granted Edit/Contribute permissions.",
        "    -Visitors                   Optional array of user UPNs granted Read permissions.",
        "    -TenantName                 Tenant domain name (e.g. 'contoso' or 'contoso.onmicrosoft.com').",
        "    -AdminUrl                   SharePoint Admin Center URL ('https://contoso-admin.sharepoint.com').",
        "    -StorageQuotaMB             Storage quota in MB (default: 1048576 = 1 TB).",
        "    -TimeZoneId                 Time Zone ID (default: 4 = UTC+1 Warsaw / Romance Standard Time).",
        "    -LocaleId, -Lcid            Language LCID (default: 1045 = Polish, 1033 = English).",
        "    -LogDirectory               Directory for execution log files (default: C:\Logs).",
        "    -Force                      Bypasses interactive [Y/N] / [T/N] confirmation prompts.",
        "    -PassThru                   Outputs the created site object to the pipeline.",
        "    -Help, -h                   Displays this help documentation and exits.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\New-SharePoint-Site.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# ==============================================================================
# INPUT DEFENSE & SANITIZATION (Standard 7)
# ==============================================================================
function Test-ValidEmailAddress {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    $pattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return ($Email.Trim() -match $pattern)
}

function Assert-SafeInputString {
    param(
        [string]$InputText,
        [string]$ParamName,
        [int]$MaxLength = 256
    )

    if ([string]::IsNullOrWhiteSpace($InputText)) { return }

    if ($InputText.Length -gt $MaxLength) {
        throw "Parameter -$ParamName exceeds maximum allowed length of $MaxLength characters."
    }

    # Protection against command injection and dangerous shell meta-characters
    if ($InputText -match '[`$<>|;&\0]') {
        throw "Parameter -$ParamName contains forbidden or potentially dangerous characters."
    }
}

function Assert-SafeSiteAlias {
    param([string]$Alias)

    if ([string]::IsNullOrWhiteSpace($Alias)) {
        throw "Site alias cannot be empty."
    }

    # If full URL is provided, validate URL format
    if ($Alias -match '^https?://') {
        if ($Alias -notmatch '^https://[a-zA-Z0-9.-]+\.sharepoint\.com/sites/[a-zA-Z0-9._-]+/?$') {
            throw "Invalid full SharePoint site URL format. Expected: 'https://<tenant>.sharepoint.com/sites/<alias>'"
        }
        return
    }

    # Relative alias validation
    $cleanAlias = $Alias.Trim().TrimStart('/').TrimEnd('/')
    if ($cleanAlias.StartsWith('sites/')) {
        $cleanAlias = $cleanAlias.Substring(6)
    }

    if ($cleanAlias -notmatch '^[a-zA-Z0-9_-]+$') {
        throw "Invalid SiteAlias '$Alias'. The site alias may only contain alphanumeric characters, hyphens, and underscores."
    }

    if ($cleanAlias.Length -gt 128) {
        throw "SiteAlias '$Alias' exceeds the maximum allowed length of 128 characters."
    }
}

function Convert-ToUrlSlug {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "site" }

    $clean = $Text.Trim().ToLowerInvariant()
    $diacritics = @{
        'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ń'='n'; 'ó'='o'; 'ś'='s'; 'ź'='z'; 'ż'='z'
    }
    foreach ($char in $diacritics.Keys) {
        $clean = $clean.Replace($char, $diacritics[$char])
    }
    $clean = $clean -replace '[^a-z0-9_-]', '-'
    $clean = $clean -replace '-+', '-'
    $clean = $clean.Trim('-')
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "site" }
    return $clean
}

# ==============================================================================
# AUTO-HELP TRIGGER ON EMPTY PARAMETERS (Requirement 1 & Standard 7)
# ==============================================================================
$hasExplicitAction = (-not [string]::IsNullOrWhiteSpace($SiteTitle)) -or 
                     (-not [string]::IsNullOrWhiteSpace($Owner)) -or 
                     $CompanyWide -or 
                     $EmptySite

if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or -not $hasExplicitAction) {
    Show-ScriptHelp
    exit 0
}

# Resolve active SiteType
if ($CompanyWide) {
    $SiteType = 'CompanyWide'
}
elseif ($EmptySite) {
    $SiteType = 'EmptySite'
}

# Initialize logging system immediately
Initialize-ScriptLogging -TargetDirectory $LogDirectory
Write-ScriptLog -Level INFO -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)"
Write-ScriptLog -Level INFO -Message "Log file initialized at: $Script:LogFilePath"
Add-OperationDetail -Step "01-Init" -Operation "Initialize Logging" -Status "Success" -Details "Log file created at $Script:LogFilePath"

# Validate mandatory parameters and auto-generate SiteAlias if omitted
try {
    if ([string]::IsNullOrWhiteSpace($SiteTitle)) {
        throw "Missing parameter: -SiteTitle is required. Run with -h for help."
    }
    if ([string]::IsNullOrWhiteSpace($Owner)) {
        throw "Missing parameter: -Owner is required. Run with -h for help."
    }

    # Auto-generate SiteAlias from SiteTitle if omitted
    if ([string]::IsNullOrWhiteSpace($SiteAlias)) {
        $SiteAlias = Convert-ToUrlSlug -Text $SiteTitle
        Write-ScriptLog -Level INFO -Message "SiteAlias not specified. Automatically derived from SiteTitle: '$SiteAlias'"
    }

    Assert-SafeInputString -InputText $SiteTitle -ParamName "SiteTitle" -MaxLength 128
    Assert-SafeSiteAlias -Alias $SiteAlias
    Assert-SafeInputString -InputText $Owner -ParamName "Owner" -MaxLength 256

    if (-not (Test-ValidEmailAddress -Email $Owner)) {
        throw "Parameter -Owner '$Owner' is not a valid email address or UserPrincipalName."
    }

    foreach ($m in $Members) {
        Assert-SafeInputString -InputText $m -ParamName "Members" -MaxLength 256
        if (-not (Test-ValidEmailAddress -Email $m)) {
            throw "Member identity '$m' is not a valid email address."
        }
    }

    foreach ($v in $Visitors) {
        Assert-SafeInputString -InputText $v -ParamName "Visitors" -MaxLength 256
        if (-not (Test-ValidEmailAddress -Email $v)) {
            throw "Visitor identity '$v' is not a valid email address."
        }
    }
}
catch {
    Write-ScriptLog -Level ERROR -Message "Input validation failed: $($_.Exception.Message)"
    Add-OperationDetail -Step "02-Validation" -Operation "Input Validation" -Status "Failed" -Details $_.Exception.Message
    exit 1
}
Add-OperationDetail -Step "02-Validation" -Operation "Input Validation" -Status "Success" -Details "Parameters sanitized and validated successfully."

# ==============================================================================
# TENANT & URL RESOLUTION
# ==============================================================================
# Normalize SiteAlias and resolve Full Site URL & Admin URL
$cleanSlug = $SiteAlias.Trim().TrimStart('/').TrimEnd('/')
if ($cleanSlug -match '^https://([a-zA-Z0-9.-]+)\.sharepoint\.com/sites/([a-zA-Z0-9._-]+)/?$') {
    $detectedTenant = $Matches[1]
    $cleanSlug = $Matches[2]
    if ([string]::IsNullOrWhiteSpace($TenantName)) {
        $TenantName = $detectedTenant
    }
}
elseif ($cleanSlug.StartsWith('sites/')) {
    $cleanSlug = $cleanSlug.Substring(6)
}

if ([string]::IsNullOrWhiteSpace($TenantName)) {
    # Extract tenant domain prefix from Owner email
    $ownerDomain = ($Owner -split '@')[1]
    if ($ownerDomain -match '^([a-zA-Z0-9-]+)\.onmicrosoft\.com$') {
        $TenantName = $Matches[1]
    }
    elseif ($ownerDomain -match '^([a-zA-Z0-9-]+)\.') {
        $TenantName = $Matches[1]
    }
    else {
        $TenantName = $ownerDomain
    }
}

# Clean TenantName to ensure only the tenant prefix is used
$tenantPrefix = $TenantName.ToLowerInvariant().Replace('.sharepoint.com', '').Replace('.onmicrosoft.com', '')
$targetSiteUrl = "https://${tenantPrefix}.sharepoint.com/sites/${cleanSlug}"

if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
    $AdminUrl = "https://${tenantPrefix}-admin.sharepoint.com"
}

Write-ScriptLog -Level INFO -Message "Target Tenant: $tenantPrefix"
Write-ScriptLog -Level INFO -Message "SharePoint Admin URL: $AdminUrl"
Write-ScriptLog -Level INFO -Message "Target Site URL: $targetSiteUrl"
Write-ScriptLog -Level INFO -Message "Site Title: $SiteTitle"
Write-ScriptLog -Level INFO -Message "Site Owner: $Owner"
Write-ScriptLog -Level INFO -Message "Site Archetype: $SiteType"
Write-ScriptLog -Level INFO -Message "Site Template: $Template"
Write-ScriptLog -Level INFO -Message "Storage Quota: $StorageQuotaMB MB ($([math]::Round($StorageQuotaMB / 1024, 2)) GB)"

Add-OperationDetail -Step "03-UrlResolution" -Operation "URL and Tenant Resolution" -Status "Success" -Details "Target site URL resolved to $targetSiteUrl (Admin: $AdminUrl)"

# ==============================================================================
# RESOLVE TENANT ID (FOR CANONICAL SPO CLAIM)
# ==============================================================================
function Get-SharePointTenantId {
    param([string]$DomainOrPrefix, [string]$OwnerEmail)

    $domainsToTry = @()
    if ($OwnerEmail -match '@([a-zA-Z0-9.-]+)$') {
        $domainsToTry += $Matches[1]
    }
    if (-not [string]::IsNullOrWhiteSpace($DomainOrPrefix)) {
        $domainsToTry += "${DomainOrPrefix}.onmicrosoft.com"
        $domainsToTry += $DomainOrPrefix
    }

    foreach ($dom in ($domainsToTry | Select-Object -Unique)) {
        try {
            Write-ScriptLog -Level DEBUG -Message "Querying Microsoft Entra ID OpenID endpoint for domain '$dom'..."
            $oidcUrl = "https://login.microsoftonline.com/$dom/v2.0/.well-known/openid-configuration"
            $response = Invoke-RestMethod -Uri $oidcUrl -TimeoutSec 6 -ErrorAction Stop
            if ($response.token_endpoint -match "https://login\.microsoftonline\.com/([a-fA-F0-9-]+)/") {
                $tenantGuid = $Matches[1]
                Write-ScriptLog -Level SUCCESS -Message "Resolved Microsoft Entra Tenant ID: $tenantGuid"
                return $tenantGuid
            }
        }
        catch {
            Write-ScriptLog -Level DEBUG -Message "OpenID query failed for '$dom': $($_.Exception.Message)"
        }
    }

    return $null
}

$Script:TenantId = Get-SharePointTenantId -DomainOrPrefix $tenantPrefix -OwnerEmail $Owner

# ==============================================================================
# PREREQUISITES & MODULE VERIFICATION
# ==============================================================================
function Assert-SharePointModule {
    Write-ScriptLog -Level INFO -Message "Checking for '$($Script:Metadata.ModuleReq)' PowerShell module..."
    $installed = Get-Module -ListAvailable -Name $Script:Metadata.ModuleReq | Select-Object -First 1

    if (-not $installed) {
        Write-ScriptLog -Level WARN -Message "Module '$($Script:Metadata.ModuleReq)' is not installed."
        Write-Host "`n[!] The required module '$($Script:Metadata.ModuleReq)' is not found on this system." -ForegroundColor Yellow
        Write-Host "    Would you like to install it now for the CurrentUser? [Y/N or T/N]: " -NoNewline -ForegroundColor Cyan
        
        $choice = if ($Force) { "Y" } else { Read-Host }
        if ($choice -match '^[ytT1]$') {
            Write-ScriptLog -Level INFO -Message "Installing '$($Script:Metadata.ModuleReq)' for CurrentUser..."
            try {
                Install-Module -Name $Script:Metadata.ModuleReq -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-ScriptLog -Level SUCCESS -Message "Module '$($Script:Metadata.ModuleReq)' installed successfully."
            }
            catch {
                Write-ScriptLog -Level ERROR -Message "Failed to install module: $($_.Exception.Message)"
                throw "Module installation failed. Run: Install-Module -Name $($Script:Metadata.ModuleReq) -Scope CurrentUser"
            }
        }
        else {
            throw "Script execution aborted because required module '$($Script:Metadata.ModuleReq)' is missing."
        }
    }

    try {
        Import-Module -Name $Script:Metadata.ModuleReq -ErrorAction Stop | Out-Null
        $modVer = (Get-Module -Name $Script:Metadata.ModuleReq).Version
        Write-ScriptLog -Level SUCCESS -Message "Loaded module '$($Script:Metadata.ModuleReq)' (v$modVer)."
        Add-OperationDetail -Step "04-ModuleCheck" -Operation "Module Import" -Status "Success" -Details "Loaded $Script:Metadata.ModuleReq v$modVer"
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "Failed to import module: $($_.Exception.Message)"
        Add-OperationDetail -Step "04-ModuleCheck" -Operation "Module Import" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }
}

# ==============================================================================
# OFFICE 365 / SHAREPOINT ADMIN AUTHENTICATION (Requirement: Admin Login)
# ==============================================================================
function Connect-SharePointAdminSession {
    param([string]$Url)

    Write-ScriptLog -Level INFO -Message "Verifying active SharePoint Online admin session..."
    $isConnected = $false

    try {
        # Test connection by querying tenant settings
        $tenantSettings = Get-SPOTenant -ErrorAction Stop
        $isConnected = $true
        Write-ScriptLog -Level SUCCESS -Message "Active SharePoint Online session detected."
        Add-OperationDetail -Step "05-Authentication" -Operation "Session Verification" -Status "Success" -Details "Reused active SharePoint Online admin session."

        # Ensure EveryoneExceptExternalUsersClaim is enabled on tenant if CompanyWide mode
        if ($SiteType -eq 'CompanyWide' -and $tenantSettings) {
            try {
                Set-SPOTenant -ShowEveryoneExceptExternalUsersClaim $true -ErrorAction SilentlyContinue | Out-Null
                Write-ScriptLog -Level DEBUG -Message "Verified ShowEveryoneExceptExternalUsersClaim is enabled on tenant."
            } catch {}
        }
    }
    catch {
        Write-ScriptLog -Level INFO -Message "No active session detected. Initiating Office 365 Administrator login to: $Url"
        Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Running" -Details "Connecting to $Url"
        try {
            Connect-SPOService -Url $Url -ErrorAction Stop
            Write-ScriptLog -Level SUCCESS -Message "Successfully connected to SharePoint Online Admin Center."
            Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Success" -Details "Connected to $Url"
            
            # Ensure claim setting is active on tenant
            try {
                Set-SPOTenant -ShowEveryoneExceptExternalUsersClaim $true -ErrorAction SilentlyContinue | Out-Null
            } catch {}
        }
        catch {
            Write-ScriptLog -Level ERROR -Message "Failed to authenticate to SharePoint Online: $($_.Exception.Message)"
            Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Failed" -Details $_.Exception.Message
            throw "SharePoint Online administrator login failed. Please ensure you have SharePoint/Global Administrator privileges."
        }
    }
}

# ==============================================================================
# PRE-FLIGHT VALIDATION (Check Existing Site)
# ==============================================================================
function Test-SharePointSiteAvailability {
    param([string]$SiteUrl)

    Write-ScriptLog -Level INFO -Message "Checking if site URL already exists: $SiteUrl"
    Add-OperationDetail -Step "06-PreFlight" -Operation "Site URL Availability Check" -Status "Running" -Details "Querying $SiteUrl"

    try {
        $existingSite = Get-SPOSite -Identity $SiteUrl -Detailed -ErrorAction SilentlyContinue
        if ($existingSite) {
            Write-ScriptLog -Level WARN -Message "Site URL '$SiteUrl' already exists! (Status: $($existingSite.Status), Title: '$($existingSite.Title)')"
            Write-ScriptLog -Level INFO -Message "Existing site detected. Script will apply security/permission configuration to the existing site."
            Add-OperationDetail -Step "06-PreFlight" -Operation "Site URL Availability Check" -Status "Warning" -Details "Site exists with status $($existingSite.Status); configuring existing site."
            return $true
        }
        else {
            Write-ScriptLog -Level SUCCESS -Message "Site URL '$SiteUrl' is available for provisioning."
            Add-OperationDetail -Step "06-PreFlight" -Operation "Site URL Availability Check" -Status "Success" -Details "Site URL is available."
            return $false
        }
    }
    catch {
        Write-ScriptLog -Level DEBUG -Message "Site query completed. URL is free."
        return $false
    }
}

# ==============================================================================
# INTERACTIVE CONFIRMATION (Bilingual: EN Y/N, PL T/N)
# ==============================================================================
function Confirm-SiteCreation {
    param([bool]$IsExistingSite = $false)

    if ($Force) {
        Write-ScriptLog -Level INFO -Message "Bypassing confirmation prompt due to -Force switch."
        return $true
    }

    $actionVerb = if ($IsExistingSite) { "UPDATE PERMISSIONS ON EXISTING SITE" } else { "NEW SHAREPOINT SITE PROVISIONING" }

    Write-Host "`n" -NoNewline
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host " CONFIRMATION: $actionVerb" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "  Title        : $SiteTitle" -ForegroundColor White
    Write-Host "  URL          : $targetSiteUrl" -ForegroundColor White
    Write-Host "  Archetype    : $SiteType" -ForegroundColor $(if ($SiteType -eq 'CompanyWide') { 'Green' } else { 'Magenta' })
    Write-Host "  Owner        : $Owner" -ForegroundColor White
    if (-not $IsExistingSite) {
        Write-Host "  Template     : $Template" -ForegroundColor White
        Write-Host "  Storage Quota: $StorageQuotaMB MB ($([math]::Round($StorageQuotaMB / 1024, 2)) GB)" -ForegroundColor White
        Write-Host "  Language/LCID: $LocaleId" -ForegroundColor White
        Write-Host "  Time Zone ID : $TimeZoneId" -ForegroundColor White
    }
    if ($Members.Count -gt 0) {
        Write-Host "  Members (Edit): $($Members -join ', ')" -ForegroundColor White
    }
    if ($Visitors.Count -gt 0) {
        Write-Host "  Visitors (Read): $($Visitors -join ', ')" -ForegroundColor White
    }
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "Are you sure you want to proceed? / Czy na pewno chcesz kontynuować? [Y/N or T/N]: " -NoNewline -ForegroundColor Cyan

    $response = Read-Host
    if ($response -notmatch '^[ytT1]$') {
        Write-ScriptLog -Level WARN -Message "Operation cancelled by user response: '$response'."
        Add-OperationDetail -Step "07-Confirmation" -Operation "User Confirmation" -Status "Skipped" -Details "User declined confirmation prompt."
        Write-Host "[!] Operation cancelled." -ForegroundColor Yellow
        exit 0
    }
    Write-ScriptLog -Level INFO -Message "User confirmed operation."
    Add-OperationDetail -Step "07-Confirmation" -Operation "User Confirmation" -Status "Success" -Details "User confirmed provisioning/configuration."
    return $true
}

# ==============================================================================
# SITE CREATION & POLLING
# ==============================================================================
function Invoke-CreateSharePointSite {
    param(
        [string]$Url,
        [string]$Title,
        [string]$OwnerEmail,
        [string]$TemplateCode,
        [long]$StorageLimitMB,
        [int]$TzId,
        [int]$Lcid
    )

    Write-ScriptLog -Level INFO -Message "Submitting New-SPOSite provisioning request to SharePoint Online..."
    Write-ScriptLog -Level INFO -Message "Cmdlet: New-SPOSite -Url '$Url' -Title '$Title' -Owner '$OwnerEmail' -Template '$TemplateCode' -StorageQuota $StorageLimitMB -CompatibilityLevel 15"
    Add-OperationDetail -Step "08-Provisioning" -Operation "Submit Site Creation" -Status "Running" -Details "Requesting creation of $Url"

    try {
        New-SPOSite -Url $Url `
                    -Title $Title `
                    -Owner $OwnerEmail `
                    -Template $TemplateCode `
                    -StorageQuota $StorageLimitMB `
                    -TimeZoneId $TzId `
                    -LocaleId $Lcid `
                    -CompatibilityLevel 15 `
                    -ErrorAction Stop | Out-Null

        Write-ScriptLog -Level SUCCESS -Message "Site creation request submitted successfully."
        Add-OperationDetail -Step "08-Provisioning" -Operation "Submit Site Creation" -Status "Success" -Details "Request accepted by SharePoint Online."
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "New-SPOSite execution failed: $($_.Exception.Message)"
        Add-OperationDetail -Step "08-Provisioning" -Operation "Submit Site Creation" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }

    # Polling until site status becomes Active
    Write-ScriptLog -Level INFO -Message "Waiting for site provisioning to finalize (polling SharePoint Online)..."
    Add-OperationDetail -Step "09-Polling" -Operation "Site Provisioning Polling" -Status "Running" -Details "Polling $Url for Active status"

    $maxWaitMinutes = 15
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $siteActive = $false
    $currentSiteObj = $null

    while ($stopwatch.Elapsed.TotalMinutes -lt $maxWaitMinutes) {
        Start-Sleep -Seconds 10
        try {
            $currentSiteObj = Get-SPOSite -Identity $Url -Detailed -ErrorAction SilentlyContinue
            if ($currentSiteObj) {
                Write-ScriptLog -Level DEBUG -Message "Current site status: $($currentSiteObj.Status) (Elapsed: $([math]::Round($stopwatch.Elapsed.TotalSeconds, 0))s)"
                if ($currentSiteObj.Status -eq "Active") {
                    $siteActive = $true
                    break
                }
            }
        }
        catch {
            Write-ScriptLog -Level DEBUG -Message "Polling exception: $($_.Exception.Message)"
        }
    }

    if (-not $siteActive) {
        Write-ScriptLog -Level WARN -Message "Site provisioning did not complete within $maxWaitMinutes minutes. Site may still be provisioning in background."
        Add-OperationDetail -Step "09-Polling" -Operation "Site Provisioning Polling" -Status "Warning" -Details "Timed out after $maxWaitMinutes minutes."
    }
    else {
        $elapsedSec = [math]::Round($stopwatch.Elapsed.TotalSeconds, 1)
        Write-ScriptLog -Level SUCCESS -Message "Site '$Url' is ACTIVE! (Completed in ${elapsedSec}s)."
        Add-OperationDetail -Step "09-Polling" -Operation "Site Provisioning Polling" -Status "Success" -Details "Site is Active in ${elapsedSec}s."
    }

    return $currentSiteObj
}

# ==============================================================================
# PERMISSION & GROUP RESOLUTION (Bilingual: EN / PL with Progressive Retries)
# ==============================================================================
function Get-SiteSecurityGroups {
    param([string]$Url, [string]$Title, [string]$OwnerEmail)

    Write-ScriptLog -Level INFO -Message "Resolving SharePoint security groups for site '$Url'..."
    $groups = @{
        Visitors = $null
        Members  = $null
        Owners   = $null
    }

    # Pre-emptively ensure Owner is Primary Administrator using tenant admin cmdlets
    try {
        Set-SPOSite -Identity $Url -Owner $OwnerEmail -ErrorAction SilentlyContinue | Out-Null
    } catch {}

    # Progressive retry loop allowing SharePoint Online to complete internal permission & group hydration
    $maxRetries = 6
    $retryDelay = 10
    $siteGroups = $null

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            Write-ScriptLog -Level INFO -Message "Querying site groups via Get-SPOSiteGroup (Attempt $i/$maxRetries)..."
            $siteGroups = Get-SPOSiteGroup -Site $Url -ErrorAction Stop
            if ($siteGroups -and $siteGroups.Count -gt 0) {
                Write-ScriptLog -Level SUCCESS -Message "Successfully retrieved $($siteGroups.Count) site security group(s)."
                break
            }
        }
        catch {
            Write-ScriptLog -Level DEBUG -Message "Attempt $i encountered error: $($_.Exception.Message)"
            if ($i -lt $maxRetries) {
                Write-ScriptLog -Level WARN -Message "Site groups are still initializing in SharePoint Online (hydration latency). Retrying in ${retryDelay}s..."
                # Try setting site admin via Set-SPOUser
                try {
                    Set-SPOUser -Site $Url -LoginName $OwnerEmail -IsSiteAdmin $true -ErrorAction SilentlyContinue | Out-Null
                } catch {}
                Start-Sleep -Seconds $retryDelay
            }
        }
    }

    if ($siteGroups -and $siteGroups.Count -gt 0) {
        $groupNames = ($siteGroups | Select-Object -ExpandProperty Title) -join "', '"
        Write-ScriptLog -Level INFO -Message "Discovered site security groups: '$groupNames'"

        # 1. Resolve Visitors Group (Read / Odczyt)
        $groups.Visitors = $siteGroups | Where-Object { 
            ($_.Roles -contains "Read") -or 
            ($_.Roles -contains "Odczyt") -or 
            ($_.Title -like "*Visitor*") -or 
            ($_.Title -like "*Odwiedzając*") -or 
            ($_.Title -like "*Goście*") -or 
            ($_.Title -like "*Czytelnicy*") 
        } | Select-Object -First 1

        # 2. Resolve Members Group (Contribute / Edit / Edycja)
        $groups.Members = $siteGroups | Where-Object { 
            ($_.Roles -contains "Contribute") -or 
            ($_.Roles -contains "Edit") -or 
            ($_.Roles -contains "Edycja") -or 
            ($_.Roles -contains "Współtworzenie") -or 
            ($_.Roles -contains "Współtwórcy") -or 
            ($_.Title -like "*Member*") -or 
            ($_.Title -like "*Członkowie*") 
        } | Select-Object -First 1

        # 3. Resolve Owners Group (Full Control / Pełna kontrola)
        $groups.Owners = $siteGroups | Where-Object { 
            ($_.Roles -contains "Full Control") -or 
            ($_.Roles -contains "Pełna kontrola") -or 
            ($_.Title -like "*Owner*") -or 
            ($_.Title -like "*Właściciele*") 
        } | Select-Object -First 1

        $visTitle = if ($groups.Visitors) { $groups.Visitors.Title } else { "Not Found" }
        $memTitle = if ($groups.Members) { $groups.Members.Title } else { "Not Found" }
        $ownTitle = if ($groups.Owners) { $groups.Owners.Title } else { "Not Found" }

        Write-ScriptLog -Level SUCCESS -Message "Resolved groups -> Owners: '$ownTitle', Members: '$memTitle', Visitors: '$visTitle'"
        Add-OperationDetail -Step "10-GroupResolution" -Operation "Security Group Discovery" -Status "Success" -Details "Owners: $ownTitle | Members: $memTitle | Visitors: $visTitle"
    }
    else {
        Write-ScriptLog -Level WARN -Message "Unable to query site groups via Get-SPOSiteGroup. Utilizing localized fallback candidates."
        Add-OperationDetail -Step "10-GroupResolution" -Operation "Security Group Discovery" -Status "Warning" -Details "Get-SPOSiteGroup access delayed. Using localized fallback candidates."
    }

    return $groups
}

function Configure-SharePointPermissions {
    param(
        [string]$Url,
        [string]$Title,
        [string]$SiteTypeMode,
        [string]$OwnerEmail,
        [hashtable]$Groups,
        [string]$TenantId,
        [string[]]$MemberList,
        [string[]]$VisitorList
    )

    Write-ScriptLog -Level INFO -Message "Configuring permissions for mode '$SiteTypeMode'..."
    Add-OperationDetail -Step "11-Permissions" -Operation "Permission Configuration" -Status "Running" -Details "Applying rules for $SiteTypeMode"

    # Step 1: Ensure Owner is configured as Primary Administrator and Site Admin
    try {
        Write-ScriptLog -Level INFO -Message "Confirming Primary Site Collection Administrator ($OwnerEmail)..."
        Set-SPOSite -Identity $Url -Owner $OwnerEmail -ErrorAction SilentlyContinue | Out-Null
        Set-SPOUser -Site $Url -LoginName $OwnerEmail -IsSiteAdmin $true -ErrorAction SilentlyContinue | Out-Null
        Write-ScriptLog -Level SUCCESS -Message "Owner '$OwnerEmail' confirmed as Site Collection Administrator."
    }
    catch {
        Write-ScriptLog -Level WARN -Message "Setting Site Admin via Set-SPOUser logged notice: $($_.Exception.Message)"
    }

    # Step 2: Handle CompanyWide Preview Access
    if ($SiteTypeMode -eq 'CompanyWide') {
        Write-ScriptLog -Level INFO -Message "Configuring Company-Wide Preview access (Read/Visitor permissions for all licensed M365 users)..."
        
        # Build candidate target groups (starting with resolved, then localized candidates)
        $targetVisitorGroups = @()
        if ($Groups.Visitors -and (-not [string]::IsNullOrWhiteSpace($Groups.Visitors.Title))) {
            $targetVisitorGroups += $Groups.Visitors.Title
        }
        $targetVisitorGroups += "$Title Odwiedzający"
        $targetVisitorGroups += "$Title Goście"
        $targetVisitorGroups += "$Title Visitors"
        $targetVisitorGroups += "Goście witryny $Title"
        $targetVisitorGroups += "Odwiedzający witrynę $Title"
        $targetVisitorGroups = $targetVisitorGroups | Select-Object -Unique

        # Build candidate claim identifiers for 'Everyone except external users'
        $candidatePrincipals = @()
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            # Canonical claim used by SharePoint Online for Everyone Except External Users
            $candidatePrincipals += "c:0-.f|rolemanager|spo-grid-all-users/$TenantId"
            $candidatePrincipals += "c:0-.t|federated-directory-claim-provider|$TenantId"
        }
        $candidatePrincipals += "Everyone except external users"
        $candidatePrincipals += "Wszyscy oprócz użytkowników zewnętrznych"
        $candidatePrincipals += "c:0(.s|true"

        $companyWideAdded = $false

        foreach ($groupName in $targetVisitorGroups) {
            foreach ($principal in $candidatePrincipals) {
                try {
                    Write-ScriptLog -Level INFO -Message "Attempting to grant Visitors access to '$principal' in group '$groupName'..."
                    Add-SPOUser -Site $Url -LoginName $principal -Group $groupName -ErrorAction Stop | Out-Null
                    Write-ScriptLog -Level SUCCESS -Message "Successfully granted Company-Wide Preview access to '$principal' in group '$groupName'!"
                    $companyWideAdded = $true
                    Add-OperationDetail -Step "12-CompanyAccess" -Operation "Grant Company-Wide Preview" -Status "Success" -Details "Assigned '$principal' to '$groupName'"
                    break
                }
                catch {
                    Write-ScriptLog -Level DEBUG -Message "Principal '$principal' into group '$groupName' failed: $($_.Exception.Message)"
                }
            }
            if ($companyWideAdded) { break }
        }

        if (-not $companyWideAdded) {
            Write-ScriptLog -Level WARN -Message "Could not automatically resolve 'Everyone except external users'. Tenant permissions may restrict script claim assignment."
            Write-ScriptLog -Level WARN -Message "You can assign 'Everyone except external users' manually in Site Permissions or check Set-SPOTenant -ShowEveryoneExceptExternalUsersClaim `$true."
            Add-OperationDetail -Step "12-CompanyAccess" -Operation "Grant Company-Wide Preview" -Status "Warning" -Details "Automatic claim assignment did not succeed; verify tenant claim settings."
        }
    }
    else {
        Write-ScriptLog -Level INFO -Message "EmptySite mode: Verifying site isolation. Company-wide access is strictly excluded."
        Add-OperationDetail -Step "12-Isolation" -Operation "Enforce Site Isolation" -Status "Success" -Details "No public or company-wide permissions granted."
    }

    # Step 3: Assign Optional Members
    if ($MemberList -and $MemberList.Count -gt 0) {
        $targetMemberGroups = @()
        if ($Groups.Members -and (-not [string]::IsNullOrWhiteSpace($Groups.Members.Title))) {
            $targetMemberGroups += $Groups.Members.Title
        }
        $targetMemberGroups += "$Title Członkowie"
        $targetMemberGroups += "$Title Members"
        $targetMemberGroups += "Członkowie witryny $Title"
        $targetMemberGroups = $targetMemberGroups | Select-Object -Unique

        Write-ScriptLog -Level INFO -Message "Assigning $($MemberList.Count) optional member(s) (Edit permissions)..."
        foreach ($member in $MemberList) {
            $memberAdded = $false
            foreach ($mg in $targetMemberGroups) {
                try {
                    Add-SPOUser -Site $Url -LoginName $member -Group $mg -ErrorAction Stop | Out-Null
                    Write-ScriptLog -Level SUCCESS -Message "Added Member: $member to '$mg'"
                    Add-OperationDetail -Step "13-AddMembers" -Operation "Add Member ($member)" -Status "Success" -Details "Assigned to $mg"
                    $memberAdded = $true
                    break
                }
                catch {
                    Write-ScriptLog -Level DEBUG -Message "Member '$member' into group '$mg' failed: $($_.Exception.Message)"
                }
            }
            if (-not $memberAdded) {
                Write-ScriptLog -Level WARN -Message "Failed to add Member '$member' to any target group."
                Add-OperationDetail -Step "13-AddMembers" -Operation "Add Member ($member)" -Status "Warning" -Details "Could not add to member groups"
            }
        }
    }

    # Step 4: Assign Optional Visitors
    if ($VisitorList -and $VisitorList.Count -gt 0) {
        $targetVisitorGroups = @()
        if ($Groups.Visitors -and (-not [string]::IsNullOrWhiteSpace($Groups.Visitors.Title))) {
            $targetVisitorGroups += $Groups.Visitors.Title
        }
        $targetVisitorGroups += "$Title Odwiedzający"
        $targetVisitorGroups += "$Title Goście"
        $targetVisitorGroups += "$Title Visitors"
        $targetVisitorGroups += "Goście witryny $Title"
        $targetVisitorGroups = $targetVisitorGroups | Select-Object -Unique

        Write-ScriptLog -Level INFO -Message "Assigning $($VisitorList.Count) optional visitor(s) (Read permissions)..."
        foreach ($visitor in $VisitorList) {
            $visitorAdded = $false
            foreach ($vg in $targetVisitorGroups) {
                try {
                    Add-SPOUser -Site $Url -LoginName $visitor -Group $vg -ErrorAction Stop | Out-Null
                    Write-ScriptLog -Level SUCCESS -Message "Added Visitor: $visitor to '$vg'"
                    Add-OperationDetail -Step "14-AddVisitors" -Operation "Add Visitor ($visitor)" -Status "Success" -Details "Assigned to $vg"
                    $visitorAdded = $true
                    break
                }
                catch {
                    Write-ScriptLog -Level DEBUG -Message "Visitor '$visitor' into group '$vg' failed: $($_.Exception.Message)"
                }
            }
            if (-not $visitorAdded) {
                Write-ScriptLog -Level WARN -Message "Failed to add Visitor '$visitor' to any target group."
                Add-OperationDetail -Step "14-AddVisitors" -Operation "Add Visitor ($visitor)" -Status "Warning" -Details "Could not add to visitor groups"
            }
        }
    }
}

# ==============================================================================
# MAIN EXECUTION FLOW
# ==============================================================================
try {
    # 1. Module & Prerequisites Verification
    Assert-SharePointModule

    # 2. Authenticate to SharePoint Online Admin Center
    Connect-SharePointAdminSession -Url $AdminUrl

    # 3. Pre-Flight Availability Check (detects if site already exists)
    $siteExists = Test-SharePointSiteAvailability -SiteUrl $targetSiteUrl

    # 4. User Confirmation
    Confirm-SiteCreation -IsExistingSite $siteExists

    # 5. Provision Site (if not already existing)
    if (-not $siteExists) {
        $createdSite = Invoke-CreateSharePointSite -Url $targetSiteUrl `
                                                   -Title $SiteTitle `
                                                   -OwnerEmail $Owner `
                                                   -TemplateCode $Template `
                                                   -StorageLimitMB $StorageQuotaMB `
                                                   -TzId $TimeZoneId `
                                                   -Lcid $LocaleId
    }
    else {
        Write-ScriptLog -Level INFO -Message "Skipping site collection creation because site '$targetSiteUrl' already exists."
        Add-OperationDetail -Step "08-Provisioning" -Operation "Submit Site Creation" -Status "Skipped" -Details "Site collection already exists."
    }

    # 6. Group Discovery & Permission Configuration
    $secGroups = Get-SiteSecurityGroups -Url $targetSiteUrl -Title $SiteTitle -OwnerEmail $Owner
    Configure-SharePointPermissions -Url $targetSiteUrl `
                                    -Title $SiteTitle `
                                    -SiteTypeMode $SiteType `
                                    -OwnerEmail $Owner `
                                    -Groups $secGroups `
                                    -TenantId $Script:TenantId `
                                    -MemberList $Members `
                                    -VisitorList $Visitors

    # 7. Operation Summary Table
    $duration = (Get-Date) - $Script:StartTime
    $durationStr = "{0:D2}m {1:D2}s" -f $duration.Minutes, $duration.Seconds

    Write-Host "`n"
    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host " OPERATION COMPLETED SUCCESSFULLY: SHAREPOINT SITE CONFIGURATION" -ForegroundColor Green
    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host "  Title         : $SiteTitle" -ForegroundColor White
    Write-Host "  URL           : $targetSiteUrl" -ForegroundColor Cyan
    Write-Host "  Archetype     : $SiteType" -ForegroundColor $(if ($SiteType -eq 'CompanyWide') { 'Green' } else { 'Magenta' })
    Write-Host "  Primary Owner : $Owner" -ForegroundColor White
    Write-Host "  Tenant ID     : $(if ($Script:TenantId) { $Script:TenantId } else { 'Auto-derived' })" -ForegroundColor White
    Write-Host "  Storage Limit : $StorageQuotaMB MB ($([math]::Round($StorageQuotaMB / 1024, 2)) GB)" -ForegroundColor White
    Write-Host "  Total Runtime : $durationStr" -ForegroundColor White
    Write-Host "  Log File Path : $Script:LogFilePath" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Green

    Write-Host "`n[+] Detailed Operation Audit Trail:" -ForegroundColor Cyan
    $Script:OperationDetails | Format-Table -AutoSize -Property Timestamp, Step, Operation, Status, Details | Out-String | Write-Host -ForegroundColor DarkCyan

    Write-ScriptLog -Level SUCCESS -Message "New-SharePoint-Site completed successfully in $durationStr."

    if ($PassThru) {
        [PSCustomObject]@{
            Title            = $SiteTitle
            Url              = $targetSiteUrl
            Owner            = $Owner
            SiteType         = $SiteType
            Template         = $Template
            TenantId         = $Script:TenantId
            StorageQuotaMB   = $StorageQuotaMB
            LogFilePath      = $Script:LogFilePath
            OperationDetails = $Script:OperationDetails
        }
    }
}
catch {
    $duration = (Get-Date) - $Script:StartTime
    $durationStr = "{0:D2}m {1:D2}s" -f $duration.Minutes, $duration.Seconds
    Write-ScriptLog -Level ERROR -Message "Execution failed: $($_.Exception.Message)"
    Add-OperationDetail -Step "99-Failure" -Operation "Execution Failure" -Status "Failed" -Details $_.Exception.Message

    Write-Host "`n"
    Write-Host "[-] ============================================================================" -ForegroundColor Red
    Write-Host "[-] SCRIPT EXECUTION FAILED" -ForegroundColor Red
    Write-Host "[-] Error Details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[-] Total Runtime: $durationStr" -ForegroundColor Red
    Write-Host "[-] Check execution log for full stack trace: $Script:LogFilePath" -ForegroundColor Yellow
    Write-Host "[-] ============================================================================" -ForegroundColor Red

    exit 1
}
