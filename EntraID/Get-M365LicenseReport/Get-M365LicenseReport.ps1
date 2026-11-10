<#
.SYNOPSIS
    Audits, categorizes, and reports on Microsoft 365 tenant licenses and assigned users via Microsoft Graph.

.DESCRIPTION
    Get-M365LicenseReport is an enterprise-grade PowerShell tool designed for Microsoft 365 administrators.
    It inspects tenant subscription SKUs, calculates license consumption and remaining quotas, translates
    raw SKU part numbers into human-readable commercial names, and audits user-level license assignments
    including inheritance detection (Direct vs. Group-Based Licensing / Inherited).

    The script provides 3 distinct report types:
    1. License Summary Report (-Summary): Tenant-wide subscription inventory and quotas (Total, Used, Free, % Used).
    2. User-Centric Report (-ByUser [-All | -User <UPN/Name>]): Audits licenses by user account (all users or targeted user).
    3. License-Centric Report (-ByLicense [-SkuPartNumber <SKU>]): Audits users by license type with quota capacity metrics.

    Key Capabilities:
    - Executive Licensing Summary: Active/purchased, consumed, warning, available pools, and utilization %.
    - Commercial Name Translation: Comprehensive built-in catalog translating SKUs (e.g. SPE_E5, ENTERPRISEPACK).
    - User Assignment Forensics: Discovers whether a license is assigned Directly or Inherited via an Entra ID group.
    - Targeted User Lookup: Quickly audit a single user without scanning the entire organization.
    - Multi-Format Export: Generates UTF-8 BOM CSV, formatted TXT, structured JSON, and interactive standalone HTML dashboards.
    - Enterprise Logging: Unified daily log rotation in C:\Logs\Get-M365LicenseReport\ with timestamped entries.

.PARAMETER Summary
    Switch parameter. Report 1: Generates only the tenant subscription summary table and quotas without listing users.
    Aliases: -SummaryOnly, -LicenseSummary, -s.

.PARAMETER ByUser
    Switch parameter. Report 2: Generates a user-centric license report. Combine with -All or -User <UPN/Name>.
    Aliases: -Users, -UserSummary, -u.

.PARAMETER All
    Switch parameter. Used with -ByUser to audit all licensed users across the entire organization.

.PARAMETER User
    Specifies a single user account to audit by UserPrincipalName, DisplayName, or Email address.
    Aliases: -UserPrincipalName, -UPN, -Search.

.PARAMETER ByLicense
    Switch parameter. Report 3: Generates a license-centric report showing quota metrics (Total, Used, Free)
    followed by the itemized roster of users holding each license.
    Aliases: -LicenseReport, -Licenses, -l.

.PARAMETER SkuPartNumber
    Filters the license-centric report (-ByLicense) by a specific SKU Part Number (e.g. "SPE_E5", "ENTERPRISEPACK").
    Aliases: -Sku, -License, -Plan.

.PARAMETER ReportType
    Explicitly selects the report type: 'Summary', 'ByUser', 'ByLicense', '1', '2', or '3'.
    Aliases: -Report, -Type.

.PARAMETER ExportFormat
    Exports the generated report to disk in the specified format: 'CSV', 'TXT', 'HTML', or 'JSON'.

.PARAMETER ExportPath
    Destination directory for exported report files. Defaults to 'C:\Reports\M365Licenses'.

.PARAMETER IncludeDisabledAccounts
    Switch parameter. Includes disabled user accounts in user drilldown reports (by default only active accounts are reported).

.PARAMETER ForceReconnection
    Forces a fresh authentication prompt to Microsoft Graph even if an active session already exists.

.PARAMETER LogDirectory
    Custom directory for persistent execution logs. Defaults to 'C:\Logs\Get-M365LicenseReport'.

.PARAMETER Help
    Displays custom interactive help banner with full usage instructions, parameter details, and examples.
    Alias: -h.

.EXAMPLE
    .\Get-M365LicenseReport.ps1
    Running without parameters displays the built-in interactive help banner and usage examples.

.EXAMPLE
    # Report 1: Tenant license summary table
    .\Get-M365LicenseReport.ps1 -Summary

.EXAMPLE
    # Report 1: Tenant license summary exported to CSV and TXT
    .\Get-M365LicenseReport.ps1 -Summary -ExportFormat CSV
    .\Get-M365LicenseReport.ps1 -Summary -ExportFormat TXT

.EXAMPLE
    # Report 2: User summary for all users in tenant
    .\Get-M365LicenseReport.ps1 -ByUser -All
    # Or simply:
    .\Get-M365LicenseReport.ps1 -All

.EXAMPLE
    # Report 2: User summary for a specific user
    .\Get-M365LicenseReport.ps1 -User "john.doe@company.com"
    .\Get-M365LicenseReport.ps1 -User "john.doe@company.com" -ExportFormat TXT

.EXAMPLE
    # Report 3: License-centric drilldown for all tenant licenses (quotas + user rosters)
    .\Get-M365LicenseReport.ps1 -ByLicense

.EXAMPLE
    # Report 3: License drilldown for a specific plan (e.g. Microsoft 365 E5) with CSV / TXT / HTML export
    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber "SPE_E5" -ExportFormat CSV
    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber "SPE_E5" -ExportFormat TXT
    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber "SPE_E5" -ExportFormat HTML

.NOTES
    Author  : Roman Pindela
    Email   : roman.pindela@gmail.com
    GitHub  : https://github.com/romanpindela
    Version : 2.0.0
#>

[CmdletBinding()]
param(
    # Report 1: License Summary
    [Alias('SummaryOnly', 'LicenseSummary', 's')]
    [switch]$Summary,

    # Report 2: User Summary
    [Alias('Users', 'UserSummary', 'u')]
    [switch]$ByUser,

    [switch]$All,

    [Alias('UserPrincipalName', 'UPN', 'Search')]
    [string]$User,

    # Report 3: By License Drilldown
    [Alias('LicenseReport', 'Licenses', 'l')]
    [switch]$ByLicense,

    [Alias('Sku', 'License', 'Plan')]
    [string]$SkuPartNumber,

    # General & Export Parameters
    [Alias('Report', 'Type')]
    [ValidateSet('Summary', 'ByUser', 'ByLicense', '1', '2', '3')]
    [string]$ReportType,

    [ValidateSet('CSV', 'TXT', 'HTML', 'JSON')]
    [string]$ExportFormat,

    [string]$ExportPath = "C:\Reports\M365Licenses",

    [switch]$IncludeDisabledAccounts,

    [switch]$ForceReconnection,

    [string]$LogDirectory = "C:\Logs\Get-M365LicenseReport",

    [Alias('h')]
    [switch]$Help
)

# Set strict mode and error preferences
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# SCRIPT METADATA & GLOBALS
# ==============================================================================
$Script:Metadata = @{
    Title       = "Get-M365LicenseReport"
    Description = "Microsoft 365 Tenant License Audit and User Drilldown Reporter"
    Version     = "2.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
}

$Script:LogFilePath       = $null
$Script:GroupCache         = @{}
$Script:TenantOrganization = $null

# ==============================================================================
# 0. INTERACTIVE HELP BANNER (-h / -Help or No Parameters)
# ==============================================================================
function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT  : $($Script:Metadata.Title).ps1",
        "VERSION : $($Script:Metadata.Version)",
        "AUTHOR  : $($Script:Metadata.Author) ($($Script:Metadata.Email))",
        "GITHUB  : $($Script:Metadata.GitHub)",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Enterprise-grade PowerShell audit and reporting tool for Microsoft 365",
        "    tenant licenses and assigned users using Microsoft Graph API.",
        "    Translates technical SKU identifiers into friendly commercial names,",
        "    tracks quota consumption, and provides 3 specialized report views:",
        "",
        "3 DEDICATED REPORT TYPES:",
        "    1. LICENSE SUMMARY REPORT (-Summary / -ReportType Summary):",
        "       Tenant-wide licensing inventory and capacity quotas.",
        "       Shows: Friendly Name, SKU, Total purchased, Consumed (Used), Available (Free),",
        "       Warning units, and capacity utilization percentage (% Used).",
        "",
        "    2. USER-CENTRIC REPORT (-ByUser [-All | -User <UPN/Name>]):",
        "       Audits licenses grouped by user account.",
        "       - Use -All to report on all licensed users across the organization.",
        "       - Use -User <UPN/Name> to inspect a specific user's assigned licenses.",
        "       Shows: UPN, Display Name, Department, Job Title, Account Status,",
        "       assigned commercial plans, and assignment forensics (Direct vs. Inherited).",
        "",
        "    3. LICENSE-CENTRIC DRILLDOWN (-ByLicense [-SkuPartNumber <SKU>]):",
        "       Audits users grouped by license type with quota capacity metrics.",
        "       - Audits all tenant licenses or scopes to a single plan via -SkuPartNumber.",
        "       Shows: For each license -> quota metrics (Total, Used, Free, % Used) followed",
        "       by the complete itemized roster of users holding that specific license.",
        "",
        "EXPORT FORMATS (-ExportFormat):",
        "    - CSV  : Saves results to CSV with UTF-8 BOM encoding for Microsoft Excel.",
        "    - TXT  : Exports a clean, monospaced plain text report (.txt).",
        "    - HTML : Compiles a standalone, modern interactive HTML dashboard with KPIs.",
        "    - JSON : Exports complete structured object hierarchy to JSON.",
        "",
        "PARAMETERS:",
        "    -Summary, -s                  Report 1: Outputs tenant licensing inventory and quotas.",
        "    -ByUser, -u                   Report 2: Generates user-centric license report.",
        "    -All                          Used with -ByUser to audit all users across tenant.",
        "    -User <UPN/Name>              Used with -ByUser to audit a specific user account.",
        "    -ByLicense, -l                Report 3: Generates license-centric user drilldown.",
        "    -SkuPartNumber, -Sku, -Plan   Filters -ByLicense by SKU Part Number (e.g. SPE_E5).",
        "    -ReportType                   Explicit report selector: 'Summary', 'ByUser', 'ByLicense'.",
        "    -ExportFormat                 Export format: 'CSV', 'TXT', 'HTML', or 'JSON'.",
        "    -ExportPath                   Destination directory (default: C:\Reports\M365Licenses).",
        "    -IncludeDisabledAccounts      Includes disabled tenant user accounts in the audit.",
        "    -ForceReconnection            Forces a fresh authentication prompt to Microsoft Graph.",
        "    -LogDirectory                 Directory for daily log files (default: C:\Logs\Get-M365LicenseReport).",
        "    -Help, -h                     Displays this interactive help banner and exits.",
        "",
        "USAGE EXAMPLES:",
        "    # 1. Report 1: Executive license summary (console table):",
        "    .\Get-M365LicenseReport.ps1 -Summary",
        "",
        "    # 2. Report 1: Executive license summary exported to CSV (Excel):",
        "    .\Get-M365LicenseReport.ps1 -Summary -ExportFormat CSV",
        "",
        "    # 3. Report 1: Executive license summary exported to plain text (.txt):",
        "    .\Get-M365LicenseReport.ps1 -Summary -ExportFormat TXT",
        "",
        "    # 4. Report 2: User summary for all licensed users in tenant:",
        "    .\Get-M365LicenseReport.ps1 -ByUser -All",
        "    .\Get-M365LicenseReport.ps1 -All",
        "",
        "    # 5. Report 2: User summary for a targeted individual user:",
        "    .\Get-M365LicenseReport.ps1 -User `"john.doe@company.com`"",
        "    .\Get-M365LicenseReport.ps1 -User `"john.doe@company.com`" -ExportFormat TXT",
        "",
        "    # 6. Report 3: License-centric drilldown for all licenses (quotas + user rosters):",
        "    .\Get-M365LicenseReport.ps1 -ByLicense",
        "",
        "    # 7. Report 3: License drilldown for a specific plan (e.g. Microsoft 365 E5):",
        "    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber `"SPE_E5`"",
        "",
        "    # 8. Report 3: Specific license drilldown exported to CSV or TXT:",
        "    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber `"SPE_E5`" -ExportFormat CSV",
        "    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber `"SPE_E5`" -ExportFormat TXT",
        "",
        "    # 9. Report 3: Specific license drilldown with interactive HTML dashboard:",
        "    .\Get-M365LicenseReport.ps1 -ByLicense -SkuPartNumber `"SPE_E5`" -ExportFormat HTML",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# ------------------------------------------------------------------------------
# Resolve Active Report Mode
# ------------------------------------------------------------------------------
$activeReport = $null

if ($Help.IsPresent -or $PSBoundParameters.Count -eq 0) {
    $activeReport = 'Help'
}
elseif ($Summary.IsPresent -or $ReportType -in @('Summary', '1')) {
    $activeReport = '1_Summary'
}
elseif ($ByUser.IsPresent -or (-not [string]::IsNullOrWhiteSpace($User)) -or ($All.IsPresent -and -not $ByLicense.IsPresent -and [string]::IsNullOrWhiteSpace($SkuPartNumber)) -or $ReportType -in @('ByUser', '2')) {
    $activeReport = '2_ByUser'
}
elseif ($ByLicense.IsPresent -or (-not [string]::IsNullOrWhiteSpace($SkuPartNumber)) -or $ReportType -in @('ByLicense', '3')) {
    $activeReport = '3_ByLicense'
}
else {
    # Default fallback to Summary
    $activeReport = '1_Summary'
}

# If Help was triggered or no arguments were supplied, display help banner and exit
if ($activeReport -eq 'Help') {
    Show-ScriptHelp
    return
}

# ==============================================================================
# 1. LOGGING SUBSYSTEM
# ==============================================================================
function Initialize-LoggingDirectory {
    [CmdletBinding()]
    param([string]$TargetDirectory)

    $resolvedDir = $TargetDirectory
    try {
        if (-not (Test-Path -Path $resolvedDir)) {
            $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
        }
        $testFile = Join-Path -Path $resolvedDir -ChildPath ".test_perm_$(Get-Random).tmp"
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
    }
    catch {
        $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Get-M365LicenseReport"
        if (-not (Test-Path -Path $resolvedDir)) {
            $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        }
    }

    $dateStamp = (Get-Date).ToString("yyyy-MM-dd")
    $logFileName = "Get-M365LicenseReport_${dateStamp}.log"
    $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName

    Write-LogEntry -Level 'INFO' -Message "================================================================================" -NoConsole
    Write-LogEntry -Level 'INFO' -Message "$($Script:Metadata.Title) v$($Script:Metadata.Version) - Logging Initialized" -NoConsole
    Write-LogEntry -Level 'INFO' -Message "Host: $env:COMPUTERNAME | User: $env:USERNAME | PowerShell: $($PSVersionTable.PSVersion)" -NoConsole
    Write-LogEntry -Level 'INFO' -Message "Log File: $Script:LogFilePath" -NoConsole
    Write-LogEntry -Level 'INFO' -Message "================================================================================" -NoConsole
}

function Write-LogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $logRecord = "[$timestamp] [$($Level.PadRight(7))] $Message"

    if ($Script:LogFilePath) {
        try {
            [System.IO.File]::AppendAllText($Script:LogFilePath, "$logRecord`r`n", [System.Text.Encoding]::UTF8)
        }
        catch {
            # Silent fallback on log disk error
        }
    }

    if (-not $NoConsole) {
        $consoleColor = switch ($Level) {
            'INFO'    { 'Cyan' }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red' }
            'DEBUG'   { 'DarkGray' }
            Default   { 'White' }
        }

        $symbol = switch ($Level) {
            'INFO'    { '[*]' }
            'WARNING' { '[!]' }
            'ERROR'   { '[-]' }
            'DEBUG'   { '[D]' }
            Default   { '[*]' }
        }

        Write-Host "$symbol $Message" -ForegroundColor $consoleColor
    }
}

# ==============================================================================
# 2. FRIENDLY SKU MAPPING DICTIONARY
# ==============================================================================
$Script:SkuMappingTable = @{
    # Microsoft 365 Enterprise & Frontline
    "SPE_E3"                            = "Microsoft 365 E3"
    "SPE_E5"                            = "Microsoft 365 E5"
    "SPE_F1"                            = "Microsoft 365 F1"
    "SPE_F3"                            = "Microsoft 365 F3"
    "SPE_E5_INFOGOV"                    = "Microsoft 365 E5 Information Protection and Governance"
    "SPE_E5_COMPLIANCE"                 = "Microsoft 365 E5 Compliance"
    "SPE_E5_SECURITY"                   = "Microsoft 365 E5 Security"
    "MICROSOFT_COPILOT_FOR_M365"        = "Microsoft 365 Copilot"
    "COPILOT_STUDIO"                    = "Microsoft Copilot Studio"

    # Microsoft 365 Business Suites
    "O365_BUSINESS_PREMIUM"             = "Microsoft 365 Business Standard"
    "SMB_BUSINESS_PREMIUM"              = "Microsoft 365 Business Premium"
    "SPB"                               = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"          = "Microsoft 365 Business Basic"
    "SMB_BUSINESS"                      = "Microsoft 365 Apps for business"
    "OFFICESUBSCRIPTION"                = "Microsoft 365 Apps for enterprise"

    # Office 365 Enterprise & Frontline
    "STANDARDPACK"                      = "Office 365 E1"
    "ENTERPRISEPACK"                    = "Office 365 E3"
    "ENTERPRISEPREMIUM"                 = "Office 365 E5"
    "ENTERPRISEPACK_WITHOUT_PROPLUS"    = "Office 365 E3 (without Apps)"
    "DESKLESSPACK"                      = "Office 365 F3"
    "DEVELOPERPACK"                     = "Microsoft 365 Developer"

    # Education Suites
    "STANDARDPACK_STUDENT"              = "Office 365 A1 for Students"
    "STANDARDPACK_FACULTY"              = "Office 365 A1 for Faculty"
    "ENTERPRISEPACK_STUDENT"            = "Office 365 A3 for Students"
    "ENTERPRISEPACK_FACULTY"            = "Office 365 A3 for Faculty"
    "ENTERPRISEPREMIUM_STUDENT"         = "Office 365 A5 for Students"
    "ENTERPRISEPREMIUM_FACULTY"         = "Office 365 A5 for Faculty"
    "M365_A3_STUDENT"                   = "Microsoft 365 A3 for Students"
    "M365_A3_FACULTY"                   = "Microsoft 365 A3 for Faculty"
    "M365_A5_STUDENT"                   = "Microsoft 365 A5 for Students"
    "M365_A5_FACULTY"                   = "Microsoft 365 A5 for Faculty"

    # Exchange Online Standalone
    "EXCHANGESTANDARD"                  = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE"                = "Exchange Online (Plan 2)"
    "EXCHANGEDESKLESS"                  = "Exchange Online Kiosk"
    "EXCHANGEARCHIVE"                   = "Exchange Online Archiving"
    "EXCHANGEARCHIVE_ADDON"             = "Exchange Online Archiving Add-on"

    # Security, Compliance & Identity
    "EMS"                               = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                        = "Enterprise Mobility + Security E5"
    "AAD_PREMIUM"                       = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"                    = "Microsoft Entra ID P2"
    "ENTRA_GOVERNANCE"                  = "Microsoft Entra ID Governance"
    "INTUNE_A"                          = "Microsoft Intune Plan 1"
    "INTUNE_P2"                         = "Microsoft Intune Plan 2"
    "INTUNE_SUITE"                      = "Microsoft Intune Suite"
    "ATP_ENTERPRISE"                    = "Microsoft Defender for Office 365 (Plan 1)"
    "THREAT_INTELLIGENCE"               = "Microsoft Defender for Office 365 (Plan 2)"
    "MDM_ENTERPRISE"                    = "Microsoft Defender for Endpoint (Plan 1)"
    "WIN_DEF_ATP"                       = "Microsoft Defender for Endpoint (Plan 2)"
    "DEFENDER_BUSINESS"                 = "Microsoft Defender for Business"
    "CCIBEAGLE"                         = "Microsoft Defender for Cloud Apps"
    "ATA"                               = "Microsoft Defender for Identity"

    # Teams, Voice & Communications
    "TEAMS_EXPLORATORY"                 = "Microsoft Teams Exploratory"
    "TEAMS_ESSENTIALS"                  = "Microsoft Teams Essentials"
    "TEAMS_COMMERCIAL_CLOUD"            = "Microsoft Teams Commercial Cloud"
    "MCOMEETADV"                        = "Teams Audio Conferencing"
    "PHONERF"                           = "Teams Phone Standard"
    "COMMUNICATION_CREDITS"             = "Communication Credits"
    "MCOEV"                             = "Skype for Business / Teams Voice"
    "TEAMS_ROOMS_PRO"                   = "Microsoft Teams Rooms Pro"
    "TEAMS_ROOMS_BASIC"                 = "Microsoft Teams Rooms Basic"

    # Power Platform & Analytics
    "POWER_BI_STANDARD"                 = "Power BI (Free)"
    "POWER_BI_PRO"                      = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER"         = "Power BI Premium Per User"
    "FLOW_FREE"                         = "Power Automate (Free)"
    "POWERAPPS_VIRAL"                   = "Power Apps (Free)"
    "POWERAPPS_PER_USER"                = "Power Apps Premium"
    "FLOW_PER_USER"                     = "Power Automate Premium"

    # Visio & Project
    "VISIOCLIENT"                       = "Visio Plan 2"
    "VISIOONLINE_PLAN1"                 = "Visio Plan 1"
    "PROJECTCLIENT"                     = "Project Plan 3"
    "PROJECTPREMIUM"                    = "Project Plan 5"
    "PROJECTESSENTIALS"                 = "Project Plan 1"

    # Windows OS Subscriptions
    "WIN10_PRO_ENT_SUB"                 = "Windows 10/11 Enterprise E3"
    "WINDOWS_ENTERPRISE_E5"             = "Windows 10/11 Enterprise E5"
    "WINDOWS_STORE"                     = "Windows Store for Business"
    "RIGHTSMANAGEMENT_ADHOC"            = "Azure Rights Management (Free)"
}

function Get-FriendlySkuName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SkuPartNumber
    )

    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) {
        return "Unknown SKU"
    }

    $cleanSku = $SkuPartNumber.Trim()
    if ($Script:SkuMappingTable.ContainsKey($cleanSku)) {
        return $Script:SkuMappingTable[$cleanSku]
    }

    return $cleanSku
}

# ==============================================================================
# 3. AUTHENTICATION & CONNECTION SUBSYSTEM
# ==============================================================================
function Connect-GraphService {
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    $requiredScopes = @("User.Read.All", "Organization.Read.All")
    Write-LogEntry -Level 'INFO' -Message "Checking Microsoft Graph connection state..."

    $currentContext = $null
    try {
        $currentContext = Get-MgContext -ErrorAction SilentlyContinue
    }
    catch {
        Write-LogEntry -Level 'DEBUG' -Message "Get-MgContext threw an exception: $($_.Exception.Message)"
    }

    $needsConnection = $Force -or ($null -eq $currentContext) -or ([string]::IsNullOrWhiteSpace($currentContext.Account))

    if (-not $needsConnection) {
        $grantedScopes = $currentContext.Scopes
        $missingScopes = @()
        foreach ($scope in $requiredScopes) {
            if ($grantedScopes -notcontains $scope -and $grantedScopes -notcontains "Directory.Read.All") {
                $missingScopes += $scope
            }
        }

        if (@($missingScopes).Count -gt 0) {
            Write-LogEntry -Level 'WARNING' -Message "Current session is missing required scopes ($($missingScopes -join ', ')). Re-authenticating..."
            $needsConnection = $true
        }
    }

    if ($needsConnection) {
        Write-LogEntry -Level 'INFO' -Message "Connecting to Microsoft Graph (Scopes: $($requiredScopes -join ', '))..."
        try {
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
            $currentContext = Get-MgContext
            Write-LogEntry -Level 'INFO' -Message "Successfully connected to Microsoft Graph as '$($currentContext.Account)' (Tenant ID: $($currentContext.TenantId))."
        }
        catch {
            if ($_.Exception -is [System.TypeLoadException] -or $_.Exception.Message -match 'RefreshCacheAsync|InMemoryTokenCacheOptions') {
                Write-LogEntry -Level 'ERROR' -Message "MSAL assembly collision detected in Windows PowerShell 5.1! An earlier module in this session (e.g. ExchangeOnlineManagement) locked an older Microsoft.Identity.Client DLL."
                Write-LogEntry -Level 'WARNING' -Message "SOLUTION: Open a NEW, CLEAN PowerShell console window or execute the script in PowerShell 7+ (pwsh)."
            }
            Write-LogEntry -Level 'ERROR' -Message "Failed to authenticate to Microsoft Graph: $($_.Exception.Message)"
            throw
        }
    }
    else {
        Write-LogEntry -Level 'INFO' -Message "Reusing active Microsoft Graph session for '$($currentContext.Account)' (Tenant ID: $($currentContext.TenantId))."
    }

    # Retrieve organization profile
    try {
        $org = $null
        if (Get-Command Get-MgOrganization -ErrorAction SilentlyContinue) {
            $org = Get-MgOrganization -Property Id,DisplayName,VerifiedDomains -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        else {
            $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization?`$select=id,displayName,verifiedDomains" -ErrorAction SilentlyContinue
            if ($response -and $response.value -and @($response.value).Count -gt 0) {
                $org = $response.value[0]
            }
        }

        if ($org) {
            $primaryDomain = ""
            if ($org.VerifiedDomains) {
                $defaultDomainObj = $org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1
                if ($defaultDomainObj) { $primaryDomain = $defaultDomainObj.Name }
                elseif (@($org.VerifiedDomains).Count -gt 0) { $primaryDomain = $org.VerifiedDomains[0].Name }
            }
            $Script:TenantOrganization = @{
                TenantId      = $currentContext.TenantId
                DisplayName   = if ($org.DisplayName) { $org.DisplayName } else { "Microsoft 365 Organization" }
                PrimaryDomain = $primaryDomain
            }
        }
        else {
            $Script:TenantOrganization = @{
                TenantId      = $currentContext.TenantId
                DisplayName   = "Microsoft 365 Organization"
                PrimaryDomain = "onmicrosoft.com"
            }
        }
    }
    catch {
        Write-LogEntry -Level 'DEBUG' -Message "Unable to retrieve organization profile: $($_.Exception.Message)"
        $Script:TenantOrganization = @{
            TenantId      = $currentContext.TenantId
            DisplayName   = "Microsoft 365 Organization"
            PrimaryDomain = "-"
        }
    }
}

# ==============================================================================
# 4. GROUP NAME CACHE RESOLVER (Inherited / Group-Based Licensing)
# ==============================================================================
function Resolve-GroupName {
    [CmdletBinding()]
    param([string]$GroupId)

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        return ""
    }

    if ($Script:GroupCache.ContainsKey($GroupId)) {
        return $Script:GroupCache[$GroupId]
    }

    $resolvedName = ""
    try {
        if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
            $grp = Get-MgGroup -GroupId $GroupId -Property Id,DisplayName -ErrorAction Stop
            if ($grp -and $grp.DisplayName) {
                $resolvedName = $grp.DisplayName
            }
        }
        else {
            $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId?`$select=id,displayName"
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($resp -and $resp.displayName) {
                $resolvedName = $resp.displayName
            }
        }
    }
    catch {
        Write-LogEntry -Level 'DEBUG' -Message "Could not resolve Group ID '$GroupId': $($_.Exception.Message)"
        $resolvedName = "Group: $GroupId"
    }

    if ([string]::IsNullOrWhiteSpace($resolvedName)) {
        $resolvedName = "Group: $GroupId"
    }

    $Script:GroupCache[$GroupId] = $resolvedName
    return $resolvedName
}

# ==============================================================================
# 5. DATA FETCHING: TENANT SUBSCRIBED SKUS (REPORT 1)
# ==============================================================================
function Get-TenantSubscribedSkus {
    [CmdletBinding()]
    param()

    Write-LogEntry -Level 'INFO' -Message "Retrieving subscribed tenant license SKUs from Microsoft Graph..."
    $rawSkus = @()

    try {
        if (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue) {
            $rawSkus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
        }
        else {
            Write-LogEntry -Level 'DEBUG' -Message "Get-MgSubscribedSku not found. Falling back to Invoke-MgGraphRequest..."
            $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response -and $response.value) {
                $rawSkus = @($response.value)
            }
        }
    }
    catch {
        Write-LogEntry -Level 'ERROR' -Message "Failed to query subscribed SKUs: $($_.Exception.Message)"
        throw
    }

    Write-LogEntry -Level 'INFO' -Message "Discovered $(@($rawSkus).Count) subscribed SKU pools in tenant."
    $summaryList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $idx = 1

    foreach ($sku in $rawSkus) {
        $skuPart = $sku.SkuPartNumber
        $skuId   = $sku.SkuId
        $friendlyName = Get-FriendlySkuName -SkuPartNumber $skuPart

        $totalUnits = 0
        $warningUnits = 0
        $suspendedUnits = 0

        if ($sku.PrepaidUnits) {
            $totalUnits     = [int]($sku.PrepaidUnits.Enabled)
            $warningUnits   = [int]($sku.PrepaidUnits.Warning)
            $suspendedUnits = [int]($sku.PrepaidUnits.Suspended)
        }

        $consumedUnits = [int]($sku.ConsumedUnits)
        $availableUnits = [Math]::Max(0, ($totalUnits - $consumedUnits))

        $percentUtilized = 0.0
        if ($totalUnits -gt 0) {
            $percentUtilized = [Math]::Round(([double]$consumedUnits / [double]$totalUnits) * 100, 1)
        }

        $statusTag = "OK"
        if ($availableUnits -eq 0 -and $totalUnits -gt 0) {
            $statusTag = "FULL"
        }
        elseif ($percentUtilized -ge 90.0) {
            $statusTag = "LOW"
        }

        $summaryObject = [PSCustomObject]@{
            Nr              = $idx
            FriendlyName    = $friendlyName
            SkuPartNumber   = $skuPart
            SkuId           = $skuId
            TotalUnits      = $totalUnits
            ConsumedUnits   = $consumedUnits
            AvailableUnits  = $availableUnits
            WarningUnits    = $warningUnits
            PercentUtilized = $percentUtilized
            Status          = $statusTag
        }

        $summaryList.Add($summaryObject)
        $idx++
    }

    return $summaryList
}

# ==============================================================================
# 6. DATA FETCHING: USER LICENSES AUDIT (REPORTS 2 & 3)
# ==============================================================================
function Get-TenantUserLicenseData {
    [CmdletBinding()]
    param(
        [array]$SkuSummaries,
        [string]$TargetUser,
        [string]$FilterSkuPartNumber,
        [switch]$IncludeDisabled
    )

    $skuLookup = @{}
    foreach ($s in $SkuSummaries) {
        $skuInfo = @{
            SkuPartNumber = $s.SkuPartNumber
            FriendlyName  = $s.FriendlyName
        }
        if ($s.SkuId) {
            $cleanSkuId = $s.SkuId.ToString().Trim('{}').Trim()
            $skuLookup[$cleanSkuId] = $skuInfo
        }
        if ($s.SkuPartNumber) {
            $skuLookup[$s.SkuPartNumber.ToString().Trim()] = $skuInfo
        }
    }

    $users = @()
    $userProperties = @("Id", "DisplayName", "UserPrincipalName", "Mail", "AccountEnabled", "AssignedLicenses", "LicenseAssignmentStates", "Department", "JobTitle")

    # Fast targeted single-user lookup if -User was supplied
    if (-not [string]::IsNullOrWhiteSpace($TargetUser)) {
        Write-LogEntry -Level 'INFO' -Message "Performing targeted lookup for user '$TargetUser'..."
        $cleanTarget = $TargetUser.Trim()

        try {
            if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                $u = Get-MgUser -UserId $cleanTarget -Property $userProperties -ErrorAction SilentlyContinue
                if ($u) { $users = @($u) }
                else {
                    # Try filter by UPN, Mail, or DisplayName
                    $filterStr = "userPrincipalName eq '$cleanTarget' or mail eq '$cleanTarget' or displayName eq '$cleanTarget'"
                    $users = @(Get-MgUser -Filter $filterStr -Property $userProperties -ErrorAction SilentlyContinue)
                }
            }
            else {
                $selectStr = ($userProperties -join ',')
                $uri = "https://graph.microsoft.com/v1.0/users/$cleanTarget?`$select=$selectStr"
                $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
                if ($resp) { $users = @($resp) }
            }
        }
        catch {
            Write-LogEntry -Level 'WARNING' -Message "Targeted user query exception: $($_.Exception.Message)"
        }

        if (@($users).Count -eq 0) {
            Write-LogEntry -Level 'WARNING' -Message "User '$cleanTarget' was not found in Microsoft Entra ID."
            return @()
        }
        Write-LogEntry -Level 'INFO' -Message "Found $(@($users).Count) matching user account(s)."
    }
    else {
        # Scan tenant users with optimized paging
        Write-LogEntry -Level 'INFO' -Message "Retrieving user accounts from Microsoft Graph (PageSize 999)..."
        try {
            if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
                $users = @(Get-MgUser -All -PageSize 999 -Property $userProperties -ErrorAction Stop)
            }
            else {
                $selectStr = ($userProperties -join ',')
                $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectStr&`$top=999"
                while (-not [string]::IsNullOrEmpty($uri)) {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    if ($response -and $response.value) {
                        $users += @($response.value)
                    }
                    $uri = $response.'@odata.nextLink'
                }
            }
        }
        catch {
            Write-LogEntry -Level 'ERROR' -Message "Failed to fetch user accounts: $($_.Exception.Message)"
            throw
        }

        Write-LogEntry -Level 'INFO' -Message "Total users retrieved: $(@($users).Count)."

        if (-not $IncludeDisabled) {
            $activeUsers = @($users | Where-Object { $_.AccountEnabled -eq $true })
            Write-LogEntry -Level 'INFO' -Message "Active accounts: $(@($activeUsers).Count) (Excluded $(@($users).Count - @($activeUsers).Count) disabled accounts)."
            $users = $activeUsers
        }
    }

    # Filter only users with assigned licenses
    $licensedUsers = @($users | Where-Object { $_.AssignedLicenses -and @($_.AssignedLicenses).Count -gt 0 })
    Write-LogEntry -Level 'INFO' -Message "Found $(@($licensedUsers).Count) user(s) with assigned licenses."

    $processedUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $totalCount = @($licensedUsers).Count
    $currentIndex = 0

    foreach ($user in $licensedUsers) {
        $currentIndex++
        if ($totalCount -gt 10 -and ($currentIndex % 25 -eq 0 -or $currentIndex -eq $totalCount)) {
            $pct = [int](($currentIndex / [Math]::Max(1, $totalCount)) * 100)
            Write-Progress -Activity "Auditing User Licenses" `
                -Status "Processing $currentIndex of $totalCount ($pct%) - $($user.UserPrincipalName)" `
                -PercentComplete $pct
        }

        $assignmentStateMap = @{}
        if ($user.LicenseAssignmentStates) {
            foreach ($state in $user.LicenseAssignmentStates) {
                if ($state.SkuId) {
                    $assignmentStateMap[$state.SkuId.ToString()] = $state
                }
            }
        }

        $userLicensesList = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($assigned in $user.AssignedLicenses) {
            $rawSkuId = if ($assigned.SkuId) { $assigned.SkuId.ToString() } else { "" }
            $currentSkuId = $rawSkuId.Trim('{}').Trim()
            $skuPart = "UNKNOWN"
            $friendly = "Unknown License ($currentSkuId)"

            if ($skuLookup.ContainsKey($currentSkuId)) {
                $skuPart  = $skuLookup[$currentSkuId].SkuPartNumber
                $friendly = $skuLookup[$currentSkuId].FriendlyName
            }
            elseif ($skuLookup.ContainsKey($rawSkuId)) {
                $skuPart  = $skuLookup[$rawSkuId].SkuPartNumber
                $friendly = $skuLookup[$rawSkuId].FriendlyName
            }

            # Optional SKU filter
            if (-not [string]::IsNullOrWhiteSpace($FilterSkuPartNumber)) {
                if ($skuPart -ne $FilterSkuPartNumber -and $friendly -ne $FilterSkuPartNumber -and $currentSkuId -ne $FilterSkuPartNumber) {
                    continue
                }
            }

            # Determine Assignment Type
            $assignmentType = "Direct"
            if ($assignmentStateMap.ContainsKey($currentSkuId)) {
                $stateObj = $assignmentStateMap[$currentSkuId]
                $assignedByGroup = $stateObj.AssignedByGroup

                if (-not [string]::IsNullOrWhiteSpace($assignedByGroup) -and $assignedByGroup -ne "00000000-0000-0000-0000-000000000000") {
                    $groupName = Resolve-GroupName -GroupId $assignedByGroup
                    $assignmentType = "Inherited ($groupName)"
                }
            }

            $userLicensesList.Add([PSCustomObject]@{
                SkuPartNumber  = $skuPart
                FriendlyName   = $friendly
                SkuId          = $currentSkuId
                AssignmentType = $assignmentType
            })
        }

        if ($userLicensesList.Count -gt 0) {
            $licenseNames = @($userLicensesList | ForEach-Object { "$($_.FriendlyName) [$($_.SkuPartNumber)]" })
            $assignmentDetails = @($userLicensesList | ForEach-Object { "$($_.FriendlyName) [$($_.SkuPartNumber)] [$($_.AssignmentType)]" })

            $processedUsers.Add([PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = if ($user.DisplayName) { $user.DisplayName } else { "-" }
                AccountEnabled    = [bool]$user.AccountEnabled
                Department        = if ($user.Department) { $user.Department } else { "-" }
                JobTitle          = if ($user.JobTitle) { $user.JobTitle } else { "-" }
                AssignedLicenses  = ($licenseNames -join "; ")
                LicenseCount      = $userLicensesList.Count
                AssignmentDetails = ($assignmentDetails -join " | ")
                LicenseItems      = $userLicensesList
            })
        }
    }

    if ($totalCount -gt 10) {
        Write-Progress -Activity "Auditing User Licenses" -Completed
    }

    return $processedUsers
}

# ==============================================================================
# 7. EXPORT SUBSYSTEM (CSV, TXT, HTML, JSON)
# ==============================================================================
function Export-ReportFiles {
    [CmdletBinding()]
    param(
        [string]$ActiveReportMode,
        [array]$SummaryData,
        [array]$UserData,
        [string]$Format,
        [string]$DestinationPath,
        [string]$FilterSku
    )

    Write-LogEntry -Level 'INFO' -Message "Exporting report in '$Format' format to '$DestinationPath'..."

    try {
        if (-not (Test-Path -Path $DestinationPath)) {
            $null = New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction Stop
        }
    }
    catch {
        Write-LogEntry -Level 'WARNING' -Message "Cannot create '$DestinationPath'. Falling back to current directory: $($_.Exception.Message)"
        $DestinationPath = "."
    }

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $fileSuffix = switch ($ActiveReportMode) {
        '1_Summary'   { "Summary" }
        '2_ByUser'    { "ByUser" }
        '3_ByLicense' { if ($FilterSku) { "ByLicense_$FilterSku" } else { "ByLicense" } }
        Default       { "Report" }
    }

    # --------------------------------------------------------------------------
    # 1. CSV EXPORT
    # --------------------------------------------------------------------------
    if ($Format -eq 'CSV') {
        $csvPath = Join-Path -Path $DestinationPath -ChildPath "Get-M365LicenseReport_${fileSuffix}_${timestamp}.csv"
        try {
            $recordsToExport = @()
            if ($ActiveReportMode -eq '1_Summary') {
                $recordsToExport = $SummaryData
            }
            elseif ($ActiveReportMode -eq '2_ByUser') {
                $recordsToExport = $UserData | Select-Object UserPrincipalName, DisplayName, Department, JobTitle, AccountEnabled, AssignedLicenses, LicenseCount, AssignmentDetails
            }
            elseif ($ActiveReportMode -eq '3_ByLicense') {
                # Flatten user licenses
                $flatList = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($u in $UserData) {
                    foreach ($lic in $u.LicenseItems) {
                        $skuInfo = $SummaryData | Where-Object {
                            $_.SkuPartNumber -eq $lic.SkuPartNumber -or
                            ($_.SkuId -and $lic.SkuId -and ($_.SkuId.ToString().Trim('{}') -eq $lic.SkuId.ToString().Trim('{}')))
                        } | Select-Object -First 1
                        $flatList.Add([PSCustomObject]@{
                            FriendlyName      = $lic.FriendlyName
                            SkuPartNumber     = $lic.SkuPartNumber
                            TotalUnits        = if ($skuInfo) { $skuInfo.TotalUnits } else { 0 }
                            ConsumedUnits     = if ($skuInfo) { $skuInfo.ConsumedUnits } else { 0 }
                            AvailableUnits    = if ($skuInfo) { $skuInfo.AvailableUnits } else { 0 }
                            UserPrincipalName = $u.UserPrincipalName
                            DisplayName       = $u.DisplayName
                            Department        = $u.Department
                            JobTitle          = $u.JobTitle
                            AccountEnabled    = $u.AccountEnabled
                            AssignmentType    = $lic.AssignmentType
                        })
                    }
                }
                $recordsToExport = $flatList
            }

            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $recordsToExport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8BOM -Force
            }
            else {
                $recordsToExport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
            }
            Write-LogEntry -Level 'INFO' -Message "Saved CSV Report: $csvPath"
        }
        catch {
            Write-LogEntry -Level 'ERROR' -Message "Failed to export CSV report: $($_.Exception.Message)"
        }
    }

    # --------------------------------------------------------------------------
    # 2. TXT EXPORT (Monospaced formatted text report)
    # --------------------------------------------------------------------------
    if ($Format -eq 'TXT') {
        $txtPath = Join-Path -Path $DestinationPath -ChildPath "Get-M365LicenseReport_${fileSuffix}_${timestamp}.txt"
        try {
            $txtBuilder = [System.Text.StringBuilder]::new()
            [void]$txtBuilder.AppendLine("================================================================================")
            [void]$txtBuilder.AppendLine("MICROSOFT 365 LICENSE REPORT")
            [void]$txtBuilder.AppendLine("Report Type  : $fileSuffix")
            [void]$txtBuilder.AppendLine("Organization : $($Script:TenantOrganization.DisplayName) ($($Script:TenantOrganization.PrimaryDomain))")
            [void]$txtBuilder.AppendLine("Tenant ID    : $($Script:TenantOrganization.TenantId)")
            [void]$txtBuilder.AppendLine("Generated At : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
            [void]$txtBuilder.AppendLine("Generated By : $($Script:Metadata.Title) v$($Script:Metadata.Version)")
            [void]$txtBuilder.AppendLine("================================================================================")
            [void]$txtBuilder.AppendLine("")

            if ($ActiveReportMode -eq '1_Summary') {
                [void]$txtBuilder.AppendLine("TENANT LICENSE INVENTORY & CAPACITY QUOTAS")
                [void]$txtBuilder.AppendLine("--------------------------------------------------------------------------------")
                $tbl = $SummaryData | Format-Table -Property @(
                    @{ Label = "Nr"; Expression = { $_.Nr }; Width = 4 },
                    @{ Label = "Friendly License Name"; Expression = { $_.FriendlyName }; Width = 38 },
                    @{ Label = "SKU Part Number"; Expression = { $_.SkuPartNumber }; Width = 26 },
                    @{ Label = "Total"; Expression = { $_.TotalUnits }; Alignment = "Right"; Width = 7 },
                    @{ Label = "Used"; Expression = { $_.ConsumedUnits }; Alignment = "Right"; Width = 7 },
                    @{ Label = "Free"; Expression = { $_.AvailableUnits }; Alignment = "Right"; Width = 7 },
                    @{ Label = "% Used"; Expression = { "$($_.PercentUtilized)%" }; Alignment = "Right"; Width = 8 },
                    @{ Label = "Status"; Expression = { $_.Status }; Width = 7 }
                ) | Out-String
                [void]$txtBuilder.AppendLine($tbl.TrimEnd())
            }
            elseif ($ActiveReportMode -eq '2_ByUser') {
                [void]$txtBuilder.AppendLine("USER-CENTRIC LICENSE AUDIT ($(@($UserData).Count) users)")
                [void]$txtBuilder.AppendLine("--------------------------------------------------------------------------------")
                $tbl = $UserData | Format-Table -Property @(
                    @{ Label = "UserPrincipalName"; Expression = { $_.UserPrincipalName }; Width = 32 },
                    @{ Label = "DisplayName"; Expression = { $_.DisplayName }; Width = 22 },
                    @{ Label = "Department"; Expression = { $_.Department }; Width = 18 },
                    @{ Label = "Enabled"; Expression = { $_.AccountEnabled }; Width = 8 },
                    @{ Label = "Licenses"; Expression = { $_.AssignedLicenses }; Width = 55 }
                ) | Out-String
                [void]$txtBuilder.AppendLine($tbl.TrimEnd())
            }
            elseif ($ActiveReportMode -eq '3_ByLicense') {
                [void]$txtBuilder.AppendLine("LICENSE-CENTRIC USER DRILLDOWN & QUOTAS")
                [void]$txtBuilder.AppendLine("================================================================================")

                # Group user records by SKU
                $targetSkus = if ($FilterSku) { @($SummaryData | Where-Object { $_.SkuPartNumber -eq $FilterSku -or $_.FriendlyName -eq $FilterSku }) } else { @($SummaryData) }

                foreach ($sku in $targetSkus) {
                    $skuUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
                    foreach ($u in $UserData) {
                        foreach ($lic in $u.LicenseItems) {
                            if ($lic.SkuPartNumber -eq $sku.SkuPartNumber -or ($lic.SkuId -and $sku.SkuId -and ($lic.SkuId.ToString().Trim('{}') -eq $sku.SkuId.ToString().Trim('{}')))) {
                                $skuUsers.Add([PSCustomObject]@{
                                    UserPrincipalName = $u.UserPrincipalName
                                    DisplayName       = $u.DisplayName
                                    Department        = $u.Department
                                    JobTitle          = $u.JobTitle
                                    AccountEnabled    = $u.AccountEnabled
                                    AssignmentType    = $lic.AssignmentType
                                })
                            }
                        }
                    }

                    [void]$txtBuilder.AppendLine("")
                    [void]$txtBuilder.AppendLine("[ PLAN: $($sku.FriendlyName) ($($sku.SkuPartNumber)) ]")
                    [void]$txtBuilder.AppendLine("Capacity Quotas: Total: $($sku.TotalUnits) | Used: $($sku.ConsumedUnits) | Free: $($sku.AvailableUnits) | Utilized: $($sku.PercentUtilized)% | Status: [$($sku.Status)]")
                    [void]$txtBuilder.AppendLine("--------------------------------------------------------------------------------")

                    if (@($skuUsers).Count -gt 0) {
                        $tbl = $skuUsers | Format-Table -Property @(
                            @{ Label = "UserPrincipalName"; Expression = { $_.UserPrincipalName }; Width = 32 },
                            @{ Label = "DisplayName"; Expression = { $_.DisplayName }; Width = 22 },
                            @{ Label = "Department"; Expression = { $_.Department }; Width = 18 },
                            @{ Label = "Assignment"; Expression = { $_.AssignmentType }; Width = 26 }
                        ) | Out-String
                        [void]$txtBuilder.AppendLine($tbl.TrimEnd())
                    }
                    else {
                        [void]$txtBuilder.AppendLine("  (No assigned users found for this license)")
                    }
                }
            }

            [void]$txtBuilder.AppendLine("")
            [void]$txtBuilder.AppendLine("================================================================================")
            [void]$txtBuilder.AppendLine("End of Report - Generated by $($Script:Metadata.Title) v$($Script:Metadata.Version)")
            [void]$txtBuilder.AppendLine("================================================================================")

            [System.IO.File]::WriteAllText($txtPath, $txtBuilder.ToString(), [System.Text.Encoding]::UTF8)
            Write-LogEntry -Level 'INFO' -Message "Saved TXT Report: $txtPath"
        }
        catch {
            Write-LogEntry -Level 'ERROR' -Message "Failed to export TXT report: $($_.Exception.Message)"
        }
    }

    # --------------------------------------------------------------------------
    # 3. JSON EXPORT
    # --------------------------------------------------------------------------
    if ($Format -eq 'JSON') {
        $jsonPath = Join-Path -Path $DestinationPath -ChildPath "Get-M365LicenseReport_${fileSuffix}_${timestamp}.json"
        try {
            $payload = [ordered]@{
                ReportMetadata = @{
                    ReportType    = $fileSuffix
                    GeneratedAt   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    TenantId      = $Script:TenantOrganization.TenantId
                    Organization  = $Script:TenantOrganization.DisplayName
                    PrimaryDomain = $Script:TenantOrganization.PrimaryDomain
                    ScriptVersion = $Script:Metadata.Version
                }
                LicenseSummary = $SummaryData
                UserData       = $UserData
            }
            $jsonString = $payload | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText($jsonPath, $jsonString, [System.Text.Encoding]::UTF8)
            Write-LogEntry -Level 'INFO' -Message "Saved JSON Report: $jsonPath"
        }
        catch {
            Write-LogEntry -Level 'ERROR' -Message "Failed to export JSON report: $($_.Exception.Message)"
        }
    }

    # --------------------------------------------------------------------------
    # 4. HTML EXPORT (Modern standalone dashboard)
    # --------------------------------------------------------------------------
    if ($Format -eq 'HTML') {
        $htmlPath = Join-Path -Path $DestinationPath -ChildPath "Get-M365LicenseReport_${fileSuffix}_${timestamp}.html"
        try {
            $htmlContent = Build-HtmlReport -ActiveReportMode $ActiveReportMode -SummaryData $SummaryData -UserData $UserData -FilterSku $FilterSku
            [System.IO.File]::WriteAllText($htmlPath, $htmlContent, [System.Text.Encoding]::UTF8)
            Write-LogEntry -Level 'INFO' -Message "Saved HTML Dashboard: $htmlPath"
        }
        catch {
            Write-LogEntry -Level 'ERROR' -Message "Failed to export HTML report: $($_.Exception.Message)"
        }
    }
}

# ==============================================================================
# 8. HTML DASHBOARD GENERATOR
# ==============================================================================
function Build-HtmlReport {
    [CmdletBinding()]
    param(
        [string]$ActiveReportMode,
        [array]$SummaryData,
        [array]$UserData,
        [string]$FilterSku
    )

    $generatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $tenantName = [System.Net.WebUtility]::HtmlEncode($Script:TenantOrganization.DisplayName)
    $tenantDomain = [System.Net.WebUtility]::HtmlEncode($Script:TenantOrganization.PrimaryDomain)
    $tenantId = [System.Net.WebUtility]::HtmlEncode($Script:TenantOrganization.TenantId)

    $totalPurchased = 0
    $totalConsumed  = 0
    $totalAvailable = 0
    foreach ($s in $SummaryData) {
        $totalPurchased += $s.TotalUnits
        $totalConsumed  += $s.ConsumedUnits
        $totalAvailable += $s.AvailableUnits
    }

    $overallUtilPct = 0.0
    if ($totalPurchased -gt 0) {
        $overallUtilPct = [Math]::Round(($totalConsumed / $totalPurchased) * 100, 1)
    }

    # Build Summary Table Rows
    $summaryRowsBuilder = [System.Text.StringBuilder]::new()
    foreach ($item in $SummaryData) {
        $pct = $item.PercentUtilized
        $barColor = switch ($item.Status) {
            'FULL' { "#ef4444" }
            'LOW'  { "#f59e0b" }
            Default { "#10b981" }
        }
        $badgeClass = switch ($item.Status) {
            'FULL' { "badge-danger" }
            'LOW'  { "badge-warning" }
            Default { "badge-success" }
        }

        $rowHtml = @"
<tr>
    <td class="text-center text-muted font-mono">$($item.Nr)</td>
    <td>
        <strong>$([System.Net.WebUtility]::HtmlEncode($item.FriendlyName))</strong>
        <div class="text-muted text-xs font-mono">$([System.Net.WebUtility]::HtmlEncode($item.SkuPartNumber))</div>
    </td>
    <td class="text-center font-semibold">$($item.TotalUnits)</td>
    <td class="text-center font-semibold text-primary">$($item.ConsumedUnits)</td>
    <td class="text-center font-semibold text-emerald">$($item.AvailableUnits)</td>
    <td class="text-center text-muted">$($item.WarningUnits)</td>
    <td>
        <div class="progress-container">
            <div class="progress-bar" style="width: $($pct)%; background-color: $barColor;"></div>
        </div>
        <div class="progress-label font-mono text-xs">$pct%</div>
    </td>
    <td class="text-center"><span class="badge $badgeClass">$($item.Status)</span></td>
</tr>
"@
        [void]$summaryRowsBuilder.AppendLine($rowHtml)
    }

    # Build User Table Rows (for Report 2 or Report 3)
    $userRowsBuilder = [System.Text.StringBuilder]::new()
    $userSection = ""

    if ($ActiveReportMode -in @('2_ByUser', '3_ByLicense') -and $UserData -and @($UserData).Count -gt 0) {
        $uIdx = 1

        if ($ActiveReportMode -eq '2_ByUser') {
            foreach ($u in $UserData) {
                $statusBadge = if ($u.AccountEnabled) { "<span class='badge badge-success'>Enabled</span>" } else { "<span class='badge badge-danger'>Disabled</span>" }
                $uRowHtml = @"
<tr>
    <td class="text-center text-muted font-mono">$uIdx</td>
    <td>
        <div class="font-semibold">$([System.Net.WebUtility]::HtmlEncode($u.DisplayName))</div>
        <div class="text-muted text-xs font-mono">$([System.Net.WebUtility]::HtmlEncode($u.UserPrincipalName))</div>
    </td>
    <td>$([System.Net.WebUtility]::HtmlEncode($u.Department))</td>
    <td>$([System.Net.WebUtility]::HtmlEncode($u.JobTitle))</td>
    <td class="text-center font-semibold text-primary">$($u.LicenseCount)</td>
    <td>$([System.Net.WebUtility]::HtmlEncode($u.AssignedLicenses))</td>
    <td class="text-center">$statusBadge</td>
</tr>
"@
                [void]$userRowsBuilder.AppendLine($uRowHtml)
                $uIdx++
            }

            $userSection = @"
        <section class="card mt-6">
            <div class="card-header flex-between">
                <div>
                    <h2>User-Centric License Roster</h2>
                    <p class="subtitle">Individual user accounts and assigned license suites ($(@($UserData).Count) accounts audited)</p>
                </div>
                <div class="search-box">
                    <input type="text" id="userSearch" placeholder="Search user, UPN, department, licenses..." onkeyup="filterUserTable()">
                </div>
            </div>
            <div class="table-responsive">
                <table class="data-table" id="dataTable">
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 50px;">Nr</th>
                            <th onclick="sortTable(1)">User / Principal Name</th>
                            <th onclick="sortTable(2)">Department</th>
                            <th onclick="sortTable(3)">Job Title</th>
                            <th class="text-center" onclick="sortTable(4)">Licenses</th>
                            <th onclick="sortTable(5)">Assigned Plans</th>
                            <th class="text-center" onclick="sortTable(6)">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $($userRowsBuilder.ToString())
                    </tbody>
                </table>
            </div>
        </section>
"@
        }
        elseif ($ActiveReportMode -eq '3_ByLicense') {
            # License-Centric User Rows
            foreach ($u in $UserData) {
                foreach ($lic in $u.LicenseItems) {
                    $statusBadge = if ($u.AccountEnabled) { "<span class='badge badge-success'>Enabled</span>" } else { "<span class='badge badge-danger'>Disabled</span>" }
                    $assignBadge = if ($lic.AssignmentType -like "Direct*") {
                        "<span class='badge badge-direct'>Direct</span>"
                    } else {
                        "<span class='badge badge-inherited' title='$([System.Net.WebUtility]::HtmlEncode($lic.AssignmentType))'>$([System.Net.WebUtility]::HtmlEncode($lic.AssignmentType))</span>"
                    }

                    $uRowHtml = @"
<tr>
    <td class="text-center text-muted font-mono">$uIdx</td>
    <td>
        <span class="font-medium">$([System.Net.WebUtility]::HtmlEncode($lic.FriendlyName))</span>
        <div class="text-muted text-xs font-mono">$([System.Net.WebUtility]::HtmlEncode($lic.SkuPartNumber))</div>
    </td>
    <td>
        <div class="font-semibold">$([System.Net.WebUtility]::HtmlEncode($u.DisplayName))</div>
        <div class="text-muted text-xs font-mono">$([System.Net.WebUtility]::HtmlEncode($u.UserPrincipalName))</div>
    </td>
    <td>$([System.Net.WebUtility]::HtmlEncode($u.Department))</td>
    <td>$([System.Net.WebUtility]::HtmlEncode($u.JobTitle))</td>
    <td class="text-center">$assignBadge</td>
    <td class="text-center">$statusBadge</td>
</tr>
"@
                    [void]$userRowsBuilder.AppendLine($uRowHtml)
                    $uIdx++
                }
            }

            $userSection = @"
        <section class="card mt-6">
            <div class="card-header flex-between">
                <div>
                    <h2>License-Centric User Drilldown</h2>
                    <p class="subtitle">Itemized user roster grouped by license plan with Direct vs. Inherited forensics</p>
                </div>
                <div class="search-box">
                    <input type="text" id="userSearch" placeholder="Search license, user, department, assignment..." onkeyup="filterUserTable()">
                </div>
            </div>
            <div class="table-responsive">
                <table class="data-table" id="dataTable">
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 50px;">Nr</th>
                            <th onclick="sortTable(1)">License Plan (Commercial / SKU)</th>
                            <th onclick="sortTable(2)">User / Principal Name</th>
                            <th onclick="sortTable(3)">Department</th>
                            <th onclick="sortTable(4)">Job Title</th>
                            <th class="text-center" onclick="sortTable(5)">Assignment</th>
                            <th class="text-center" onclick="sortTable(6)">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $($userRowsBuilder.ToString())
                    </tbody>
                </table>
            </div>
        </section>
"@
        }
    }

    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 License Report - $tenantName</title>
    <style>
        :root {
            --bg-color: #0f172a;
            --surface-color: #1e293b;
            --surface-hover: #334155;
            --border-color: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --primary: #38bdf8;
            --primary-hover: #0284c7;
            --success: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --purple: #a855f7;
            --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            --font-mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: var(--bg-color);
            color: var(--text-main);
            font-family: var(--font-sans);
            line-height: 1.5;
            padding: 24px;
        }

        .container { max-width: 1400px; margin: 0 auto; }
        
        header {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px 32px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
        }
        .header-top { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 16px; }
        .title-group h1 { font-size: 26px; font-weight: 700; color: #ffffff; display: flex; align-items: center; gap: 10px; }
        .title-group p { color: var(--text-muted); font-size: 14px; margin-top: 4px; }
        .meta-group { text-align: right; font-size: 13px; color: var(--text-muted); }
        .meta-group strong { color: var(--text-main); }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .kpi-card {
            background-color: var(--surface-color);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            padding: 20px;
            position: relative;
            overflow: hidden;
        }
        .kpi-card::before { content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px; }
        .kpi-card.blue::before { background-color: var(--primary); }
        .kpi-card.green::before { background-color: var(--success); }
        .kpi-card.yellow::before { background-color: var(--warning); }
        .kpi-card.purple::before { background-color: var(--purple); }

        .kpi-title { font-size: 12px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); letter-spacing: 0.05em; }
        .kpi-value { font-size: 30px; font-weight: 700; color: #ffffff; margin-top: 6px; }
        .kpi-sub { font-size: 12px; color: var(--text-muted); margin-top: 4px; }

        .card {
            background-color: var(--surface-color);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            padding: 24px;
            margin-bottom: 24px;
        }
        .card-header { margin-bottom: 20px; }
        .card-header h2 { font-size: 18px; font-weight: 600; }
        .card-header .subtitle { font-size: 13px; color: var(--text-muted); margin-top: 2px; }
        .flex-between { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }

        .search-box input {
            background-color: var(--bg-color);
            border: 1px solid var(--border-color);
            color: var(--text-main);
            padding: 8px 16px;
            border-radius: 6px;
            font-size: 14px;
            width: 340px;
            outline: none;
            transition: border-color 0.2s;
        }
        .search-box input:focus { border-color: var(--primary); }

        .table-responsive { overflow-x: auto; }
        .data-table { width: 100%; border-collapse: collapse; font-size: 13.5px; text-align: left; }
        .data-table th {
            background-color: rgba(15, 23, 42, 0.6);
            color: var(--text-muted);
            font-weight: 600;
            padding: 12px 14px;
            border-bottom: 1px solid var(--border-color);
            cursor: pointer;
            user-select: none;
            white-space: nowrap;
        }
        .data-table th:hover { color: var(--text-main); }
        .data-table td { padding: 12px 14px; border-bottom: 1px solid rgba(51, 65, 85, 0.5); vertical-align: middle; }
        .data-table tr:hover td { background-color: var(--surface-hover); }

        .progress-container {
            background-color: rgba(15, 23, 42, 0.8);
            border-radius: 4px;
            height: 8px;
            width: 100px;
            overflow: hidden;
            display: inline-block;
            vertical-align: middle;
            margin-right: 8px;
        }
        .progress-bar { height: 100%; border-radius: 4px; transition: width 0.3s ease; }
        .progress-label { display: inline-block; width: 45px; text-align: right; }

        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 0.02em;
            text-transform: uppercase;
        }
        .badge-success { background-color: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
        .badge-warning { background-color: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
        .badge-danger  { background-color: rgba(239, 68, 68, 0.15);  color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }
        .badge-direct  { background-color: rgba(56, 189, 248, 0.15); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); }
        .badge-inherited { background-color: rgba(168, 85, 247, 0.15); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.3); text-transform: none; }

        .text-center { text-align: center; }
        .text-muted { color: var(--text-muted); }
        .text-xs { font-size: 11.5px; }
        .text-primary { color: var(--primary); }
        .text-emerald { color: #34d399; }
        .font-mono { font-family: var(--font-mono); }
        .font-semibold { font-weight: 600; }
        .font-medium { font-weight: 500; }
        .mt-6 { margin-top: 24px; }

        footer { text-align: center; font-size: 12px; color: var(--text-muted); margin-top: 32px; padding-bottom: 24px; }
        footer a { color: var(--primary); text-decoration: none; }

        @media print {
            body { background: #fff; color: #000; padding: 0; }
            .card, header { border: 1px solid #ccc; box-shadow: none; background: #fff; color: #000; }
            .search-box { display: none; }
            .data-table th { background: #eee; color: #000; }
            .badge { border: 1px solid #999; }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="header-top">
                <div class="title-group">
                    <h1>
                        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
                            <line x1="8" y1="21" x2="16" y2="21"></line>
                            <line x1="12" y1="17" x2="12" y2="21"></line>
                        </svg>
                        Microsoft 365 License Report
                    </h1>
                    <p>Executive subscription inventory, quota consumption, and user assignment forensics</p>
                </div>
                <div class="meta-group">
                    <div>Organization: <strong>$tenantName</strong></div>
                    <div>Domain: <strong>$tenantDomain</strong></div>
                    <div>Tenant ID: <span class="font-mono">$tenantId</span></div>
                    <div>Generated: <strong>$generatedDate</strong></div>
                </div>
            </div>
        </header>

        <div class="kpi-grid">
            <div class="kpi-card blue">
                <div class="kpi-title">Total Purchased Seats</div>
                <div class="kpi-value">$totalPurchased</div>
                <div class="kpi-sub">Across $(@($SummaryData).Count) active subscriptions</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-title">Assigned / Consumed</div>
                <div class="kpi-value text-primary">$totalConsumed</div>
                <div class="kpi-sub">Active seat utilization: $overallUtilPct%</div>
            </div>
            <div class="kpi-card yellow">
                <div class="kpi-title">Available / Free Seats</div>
                <div class="kpi-value text-emerald">$totalAvailable</div>
                <div class="kpi-sub">Remaining unassigned licenses</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-title">Licensed Accounts Audited</div>
                <div class="kpi-value">$(if ($UserData) { @($UserData).Count } else { "N/A" })</div>
                <div class="kpi-sub">$(if ($UserData) { "User-level audit mode" } else { "Summary mode run" })</div>
            </div>
        </div>

        <section class="card">
            <div class="card-header">
                <h2>Executive License Inventory & Quotas</h2>
                <p class="subtitle">Complete inventory of Microsoft 365 plans, active capacity, and utilization percentage</p>
            </div>
            <div class="table-responsive">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 50px;">Nr</th>
                            <th>License Plan (Commercial Name / SKU)</th>
                            <th class="text-center">Total</th>
                            <th class="text-center">Consumed</th>
                            <th class="text-center">Available</th>
                            <th class="text-center">Warning</th>
                            <th style="min-width: 170px;">Utilization</th>
                            <th class="text-center">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $($summaryRowsBuilder.ToString())
                    </tbody>
                </table>
            </div>
        </section>

        $userSection

        <footer>
            Generated by <strong>$($Script:Metadata.Title)</strong> v$($Script:Metadata.Version) &bull; Author: <a href="$($Script:Metadata.GitHub)" target="_blank">$($Script:Metadata.Author)</a> &bull; $generatedDate
        </footer>
    </div>

    <script>
        function filterUserTable() {
            const input = document.getElementById('userSearch');
            if (!input) return;
            const filter = input.value.toLowerCase();
            const table = document.getElementById('dataTable');
            if (!table) return;
            const tr = table.getElementsByTagName('tr');

            for (let i = 1; i < tr.length; i++) {
                const text = tr[i].textContent || tr[i].innerText;
                if (text.toLowerCase().indexOf(filter) > -1) {
                    tr[i].style.display = "";
                } else {
                    tr[i].style.display = "none";
                }
            }
        }

        function sortTable(n) {
            const table = document.getElementById("dataTable");
            if (!table) return;
            let rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
            switching = true;
            dir = "asc";
            while (switching) {
                switching = false;
                rows = table.rows;
                for (i = 1; i < (rows.length - 1); i++) {
                    shouldSwitch = false;
                    x = rows[i].getElementsByTagName("TD")[n];
                    y = rows[i + 1].getElementsByTagName("TD")[n];
                    let xVal = x.textContent || x.innerText;
                    let yVal = y.textContent || y.innerText;
                    if (dir === "asc") {
                        if (xVal.toLowerCase() > yVal.toLowerCase()) { shouldSwitch = true; break; }
                    } else if (dir === "desc") {
                        if (xVal.toLowerCase() < yVal.toLowerCase()) { shouldSwitch = true; break; }
                    }
                }
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                    switchcount++;
                } else {
                    if (switchcount === 0 && dir === "asc") {
                        dir = "desc";
                        switching = true;
                    }
                }
            }
        }
    </script>
</body>
</html>
"@

    return $htmlTemplate
}

# ==============================================================================
# 9. MAIN CONTROLLER & EXECUTION PIPELINE
# ==============================================================================
try {
    Initialize-LoggingDirectory -TargetDirectory $LogDirectory

    Write-LogEntry -Level 'INFO' -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)..."
    Write-LogEntry -Level 'INFO' -Message "Selected Report Mode: $activeReport"

    # Connect to Microsoft Graph
    Connect-GraphService -Force:($ForceReconnection.IsPresent)

    # --------------------------------------------------------------------------
    # Fetch Tenant Subscribed SKUs (Always needed for quota analysis)
    # --------------------------------------------------------------------------
    $summaryData = @(Get-TenantSubscribedSkus)

    if (@($summaryData).Count -eq 0) {
        Write-LogEntry -Level 'WARNING' -Message "No subscribed license SKUs found in this tenant."
    }

    $userData = $null

    # ==========================================================================
    # REPORT 1: LICENSE SUMMARY ONLY
    # ==========================================================================
    if ($activeReport -eq '1_Summary') {
        Write-Host ""
        Write-Host "=======================================================================================================" -ForegroundColor Cyan
        Write-Host "                               REPORT 1: TENANT LICENSING SUMMARY                                      " -ForegroundColor Cyan
        Write-Host "=======================================================================================================" -ForegroundColor Cyan

        $summaryData | Format-Table -Property @(
            @{ Label = "Nr"; Expression = { $_.Nr }; Width = 4 },
            @{ Label = "Friendly License Name"; Expression = { $_.FriendlyName }; Width = 38 },
            @{ Label = "SKU Part Number"; Expression = { $_.SkuPartNumber }; Width = 26 },
            @{ Label = "Total"; Expression = { $_.TotalUnits }; Alignment = "Right"; Width = 7 },
            @{ Label = "Used"; Expression = { $_.ConsumedUnits }; Alignment = "Right"; Width = 7 },
            @{ Label = "Free"; Expression = { $_.AvailableUnits }; Alignment = "Right"; Width = 7 },
            @{ Label = "% Used"; Expression = { "$($_.PercentUtilized)%" }; Alignment = "Right"; Width = 8 },
            @{ Label = "Status"; Expression = { $_.Status }; Width = 7 }
        ) | Out-Host

        if (-not [string]::IsNullOrWhiteSpace($ExportFormat)) {
            Export-ReportFiles -ActiveReportMode $activeReport -SummaryData $summaryData -UserData $null -Format $ExportFormat -DestinationPath $ExportPath -FilterSku $null
        }

        Write-Output $summaryData
    }

    # ==========================================================================
    # REPORT 2: USER-CENTRIC REPORT (-ByUser [-All | -User <UPN>])
    # ==========================================================================
    elseif ($activeReport -eq '2_ByUser') {
        Write-LogEntry -Level 'INFO' -Message "Generating Report 2 (User-Centric License Summary)..."

        $userData = @(Get-TenantUserLicenseData -SkuSummaries $summaryData `
            -TargetUser $User `
            -FilterSkuPartNumber $null `
            -IncludeDisabled:($IncludeDisabledAccounts.IsPresent))

        Write-Host ""
        Write-Host "=======================================================================================================" -ForegroundColor Cyan
        Write-Host "                               REPORT 2: USER-CENTRIC LICENSE SUMMARY                                  " -ForegroundColor Cyan
        Write-Host "=======================================================================================================" -ForegroundColor Cyan

        if (@($userData).Count -eq 0) {
            Write-Host "No matching licensed users found." -ForegroundColor Yellow
        }
        else {
            $sampleCount = [Math]::Min(25, @($userData).Count)
            $userData | Select-Object -First $sampleCount | Format-Table -Property @(
                @{ Label = "UserPrincipalName"; Expression = { $_.UserPrincipalName }; Width = 32 },
                @{ Label = "DisplayName"; Expression = { $_.DisplayName }; Width = 22 },
                @{ Label = "Department"; Expression = { $_.Department }; Width = 18 },
                @{ Label = "Count"; Expression = { $_.LicenseCount }; Alignment = "Right"; Width = 6 },
                @{ Label = "Assigned Licenses"; Expression = { $_.AssignedLicenses }; Width = 55 }
            ) | Out-Host

            if (@($userData).Count -gt $sampleCount) {
                Write-Host "... and $(@($userData).Count - $sampleCount) more user account(s) emitted to pipeline / export." -ForegroundColor DarkGray
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($ExportFormat)) {
            Export-ReportFiles -ActiveReportMode $activeReport -SummaryData $summaryData -UserData $userData -Format $ExportFormat -DestinationPath $ExportPath -FilterSku $null
        }

        Write-Output $userData
    }

    # ==========================================================================
    # REPORT 3: LICENSE-CENTRIC DRILLDOWN (-ByLicense [-SkuPartNumber <SKU>])
    # ==========================================================================
    elseif ($activeReport -eq '3_ByLicense') {
        Write-LogEntry -Level 'INFO' -Message "Generating Report 3 (License-Centric User Drilldown)..."

        $userData = @(Get-TenantUserLicenseData -SkuSummaries $summaryData `
            -TargetUser $null `
            -FilterSkuPartNumber $SkuPartNumber `
            -IncludeDisabled:($IncludeDisabledAccounts.IsPresent))

        $targetSkus = if ($SkuPartNumber) {
            @($summaryData | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber -or $_.FriendlyName -eq $SkuPartNumber -or $_.SkuId -eq $SkuPartNumber })
        } else {
            @($summaryData)
        }

        Write-Host ""
        Write-Host "=======================================================================================================" -ForegroundColor Cyan
        Write-Host "                               REPORT 3: LICENSE-CENTRIC USER DRILLDOWN                                " -ForegroundColor Cyan
        Write-Host "=======================================================================================================" -ForegroundColor Cyan

        $flattenedOutput = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($sku in $targetSkus) {
            $skuUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($u in $userData) {
                foreach ($lic in $u.LicenseItems) {
                    if ($lic.SkuPartNumber -eq $sku.SkuPartNumber -or ($lic.SkuId -and $sku.SkuId -and ($lic.SkuId.ToString().Trim('{}') -eq $sku.SkuId.ToString().Trim('{}')))) {
                        $itemObj = [PSCustomObject]@{
                            FriendlyName      = $sku.FriendlyName
                            SkuPartNumber     = $sku.SkuPartNumber
                            TotalUnits        = $sku.TotalUnits
                            ConsumedUnits     = $sku.ConsumedUnits
                            AvailableUnits    = $sku.AvailableUnits
                            PercentUtilized   = $sku.PercentUtilized
                            Status            = $sku.Status
                            UserPrincipalName = $u.UserPrincipalName
                            DisplayName       = $u.DisplayName
                            Department        = $u.Department
                            JobTitle          = $u.JobTitle
                            AccountEnabled    = $u.AccountEnabled
                            AssignmentType    = $lic.AssignmentType
                        }
                        $skuUsers.Add($itemObj)
                        $flattenedOutput.Add($itemObj)
                    }
                }
            }

            Write-Host ""
            Write-Host "-------------------------------------------------------------------------------------------------------" -ForegroundColor DarkCyan
            Write-Host " PLAN: $($sku.FriendlyName) [$($sku.SkuPartNumber)]" -ForegroundColor White
            Write-Host " Total: $($sku.TotalUnits) | Consumed (Used): $($sku.ConsumedUnits) | Available (Free): $($sku.AvailableUnits) | Utilized: $($sku.PercentUtilized)% | Status: [$($sku.Status)]" -ForegroundColor Green
            Write-Host "-------------------------------------------------------------------------------------------------------" -ForegroundColor DarkCyan

            if (@($skuUsers).Count -gt 0) {
                $displayLimit = 50
                $skuUsers | Select-Object -First $displayLimit | Format-Table -Property @(
                    @{ Label = "UserPrincipalName"; Expression = { $_.UserPrincipalName }; Width = 34 },
                    @{ Label = "DisplayName"; Expression = { $_.DisplayName }; Width = 22 },
                    @{ Label = "Department"; Expression = { $_.Department }; Width = 18 },
                    @{ Label = "Assignment Type"; Expression = { $_.AssignmentType }; Width = 26 }
                ) | Out-Host

                if (@($skuUsers).Count -gt $displayLimit) {
                    Write-Host "  ... and $(@($skuUsers).Count - $displayLimit) more users holding this plan (all users emitted to pipeline / export)." -ForegroundColor DarkGray
                }
            }
            else {
                Write-Host "  (No assigned users found for this license)" -ForegroundColor DarkGray
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($ExportFormat)) {
            Export-ReportFiles -ActiveReportMode $activeReport -SummaryData $summaryData -UserData $userData -Format $ExportFormat -DestinationPath $ExportPath -FilterSku $SkuPartNumber
        }

        Write-Output $flattenedOutput
    }

    Write-LogEntry -Level 'INFO' -Message "Script execution finished successfully."
}
catch {
    Write-LogEntry -Level 'ERROR' -Message "Execution failed: $($_.Exception.Message)"
    if ($_.ScriptStackTrace) {
        Write-LogEntry -Level 'DEBUG' -Message "Stack Trace: $($_.ScriptStackTrace)"
    }
    throw
}
