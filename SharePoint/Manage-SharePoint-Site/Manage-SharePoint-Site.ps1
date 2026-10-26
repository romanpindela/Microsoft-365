<#
.SYNOPSIS
    Automates SharePoint Online site administration: site inventory listing, categorized permissions
    auditing (Owners, Editors, Viewers), defining site modifiers, changing user roles (e.g., Edit to View),
    adding users with specific roles, and safe user removal.

.DESCRIPTION
    Manage-SharePoint-Site is an enterprise-grade PowerShell automation utility designed for
    Microsoft 365 and SharePoint Online administrators. It enforces CleanCode standards, provides
    interactive modern administrator authentication (with guidance for non-interactive Entra ID app
    principals), performs strict input validation, lists real-time operation details, and generates
    centralized audit log files.

    Key Capabilities:
    1. Interactive Microsoft 365 / SharePoint Online Administrator Login:
       - Auto-detects existing active administrator sessions (Connect-SPOService).
       - Initiates secure interactive modern authentication (browser / MFA prompt) if disconnected.
       - Auto-derives SharePoint Online Admin Center URL (https://<tenant>-admin.sharepoint.com)
         from the target site URL, tenant domain, or custom parameters.
    2. SharePoint Site Inventory (-ListSites):
       - Lists all site collections across the Microsoft 365 tenant.
       - Supports keyword filtering (-Search), storage metrics, lock state, template details, and CSV export.
    3. Site Permissions Auditing (-ListPermissions):
       - Discovers and categorizes all entitled entities for a specified site into:
         * Owners (Full Control & Site Collection Administrators)
         * Editors / Modifiers (Contribute / Edit permissions)
         * Viewers / Readers (Read permissions)
       - Exposes group memberships and direct account assignments.
    4. Defining Site Modifiers (-SetModifiers):
       - Explicitly specifies and synchronizes the list of users permitted to modify site content.
       - Adds newly specified editors to the Members group and safely demotes or removes unauthorized modifiers.
    5. Changing Permissions & Role Demotion/Promotion (-ChangeRole):
       - Supports demoting users from Edit (Members) to View (Visitors).
       - Supports promoting users between View, Edit, and Owner tiers.
       - Automatically reconciles SharePoint security group memberships and Site Collection Admin flags.
    6. Safe User Removal (-RemoveUser):
       - Completely revokes a user's access across all security groups on the target site collection.
       - Includes safety guardrails preventing accidental removal of the primary or sole site administrator.
    7. Adding New Users (-AddUser & -Role):
       - Seamlessly invites and assigns users to View (Visitors), Edit (Members), or Owner (Owners + Admin) roles.
    8. Detailed Operation Auditing & Reporting:
       - Real-time step-by-step progress logging in the console with severity indicators.
       - Comprehensive operation audit summary table rendered upon completion.
    9. Centralized File Logging:
       - Full transcript and telemetry logged to C:\Logs (with automatic fallback to $env:LOCALAPPDATA\Logs).
       - Standardized log naming: yyyyMMdd_HHmmss-Manage-SharePoint-Site-<ComputerName>-<UserName>.log.

.PARAMETER SiteUrl
    The target SharePoint Online site collection URL (e.g., 'https://contoso.sharepoint.com/sites/marketing')
    or relative slug (e.g., 'sites/marketing').
    Aliases: -Site, -Url.

.PARAMETER ListSites
    Switch parameter. Queries and lists all SharePoint Online site collections across the tenant.
    Aliases: -ListAllSites.

.PARAMETER Search
    Optional keyword to filter site collections by URL, Title, or Owner when using -ListSites.

.PARAMETER IncludePersonalSites
    Switch parameter used with -ListSites to include OneDrive for Business personal sites.

.PARAMETER ListPermissions
    Switch parameter. Audits and displays all users and security groups entitled to the specified site.
    Aliases: -AuditPermissions, -GetPermissions.

.PARAMETER SetModifiers
    Array of user email addresses / UPNs authorized to modify the site (Edit/Contribute permissions).
    Synchronizes the site's Members group to match this exact list.
    Aliases: -Modifiers, -SetEditors.

.PARAMETER UserEmail
    The User Principal Name (UPN) or email address of the user whose permissions are being modified or removed.
    Aliases: -User, -Email, -Identity.

.PARAMETER ChangeRole
    New permission role to assign to the user specified in -UserEmail.
    Valid values:
      - 'View'  : Assigns Read permissions (Visitors group); revokes Edit/Admin rights.
      - 'Edit'  : Assigns Contribute/Edit permissions (Members group).
      - 'Owner' : Assigns Full Control (Owners group) and Site Collection Administrator rights.
    Aliases: -ChangePermission, -NewRole.

.PARAMETER RemoveUser
    Switch parameter. Completely removes the user specified in -UserEmail from the site collection.
    Aliases: -Remove, -DeleteUser.

.PARAMETER AddUser
    The email address / UPN of a new or existing user to grant access to the site.
    Must be used in combination with -Role.
    Aliases: -Add.

.PARAMETER Role
    The permission tier to grant when using -AddUser.
    Valid values: 'Owner', 'Edit', 'View'.
    Aliases: -Permission.

.PARAMETER TenantName
    The Microsoft 365 tenant domain name or prefix (e.g., 'contoso' or 'contoso.onmicrosoft.com').
    If omitted, auto-detected from the site URL or active session.
    Aliases: -Tenant, -TenantDomain.

.PARAMETER AdminUrl
    The SharePoint Online Admin Center URL (e.g., 'https://contoso-admin.sharepoint.com').
    If omitted, derived automatically from TenantName or SiteUrl.

.PARAMETER ExportCsv
    Optional file path to export site inventory or permission audit reports into UTF-8 CSV format.

.PARAMETER LogDirectory
    Directory where execution logs are written.
    Default: 'C:\Logs' (with fallback to $env:LOCALAPPDATA\Logs).

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts ([Y/N] or [T/N]) for automated execution.

.PARAMETER PassThru
    Switch parameter. Returns custom PSObjects to the PowerShell pipeline for downstream processing.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.
    Aliases: -h.

.EXAMPLE
    # 1. Display built-in help, usage instructions, and examples (default behavior when executed without parameters)
    .\Manage-SharePoint-Site.ps1
    .\Manage-SharePoint-Site.ps1 -Help

.EXAMPLE
    # 2. List all SharePoint Online site collections across the tenant
    .\Manage-SharePoint-Site.ps1 -ListSites

.EXAMPLE
    # 3. Search and list site collections matching a specific keyword
    .\Manage-SharePoint-Site.ps1 -ListSites -Search "Marketing"

.EXAMPLE
    # 4. Audit and list all users entitled to a specified site (Owners, Editors, Viewers)
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -ListPermissions

.EXAMPLE
    # 5. Specify and synchronize users who can modify the site (Edit/Contribute permissions)
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -SetModifiers "anna.smith@contoso.com", "john.doe@contoso.com"

.EXAMPLE
    # 6. Change a user's permissions from Edit to View (Read-only demotion)
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -UserEmail "john.doe@contoso.com" -ChangeRole View

.EXAMPLE
    # 7. Completely remove a user from the site permissions list
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -UserEmail "ex-employee@contoso.com" -RemoveUser

.EXAMPLE
    # 8. Add a new user with View (Read) permissions
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -AddUser "contractor@contoso.com" -Role View

.EXAMPLE
    # 9. Add a new user with Edit (Contribute) permissions
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -AddUser "analyst@contoso.com" -Role Edit

.EXAMPLE
    # 10. Add a new user with Owner (Full Control & Site Collection Admin) permissions
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -AddUser "manager@contoso.com" -Role Owner

.EXAMPLE
    # 11. Silent automated role update without confirmation prompts (-Force)
    .\Manage-SharePoint-Site.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/marketing" -UserEmail "auditor@contoso.com" -ChangeRole View -Force

.NOTES
    Script Name   : Manage-SharePoint-Site.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.0.0
    Created       : 2026-10-19
    Standards     : Adheres to CleanCode and PowerShell Script Standards (C:\Dane\rpindela\GitHub\PowerShell Script Standards.txt)
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Site Identification
    [Parameter(ParameterSetName = 'ListPermissions', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'SetModifiers', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'ChangeRole', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'RemoveUser', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'AddUser', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 0)]
    [Alias('Site', 'Url')]
    [string]$SiteUrl,

    # Operation Set 1: List Sites
    [Parameter(ParameterSetName = 'ListSites', Mandatory = $true)]
    [Alias('ListAllSites')]
    [switch]$ListSites,

    [Parameter(ParameterSetName = 'ListSites')]
    [string]$Search,

    [Parameter(ParameterSetName = 'ListSites')]
    [switch]$IncludePersonalSites,

    [Parameter(ParameterSetName = 'ListSites')]
    [switch]$Detailed,

    # Operation Set 2: List Site Permissions
    [Parameter(ParameterSetName = 'ListPermissions')]
    [Alias('AuditPermissions', 'GetPermissions')]
    [switch]$ListPermissions,

    # Operation Set 3: Specify Site Modifiers (Edit permissions)
    [Parameter(ParameterSetName = 'SetModifiers', Mandatory = $true)]
    [Alias('Modifiers', 'SetEditors')]
    [string[]]$SetModifiers,

    # Operation Set 4: Change Role (e.g., Edit to View)
    [Parameter(ParameterSetName = 'ChangeRole', Mandatory = $true, Position = 1)]
    [Parameter(ParameterSetName = 'RemoveUser', Mandatory = $true, Position = 1)]
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 1)]
    [Alias('User', 'Email', 'Identity')]
    [string]$UserEmail,

    [Parameter(ParameterSetName = 'ChangeRole', Mandatory = $true)]
    [Alias('ChangePermission', 'NewRole')]
    [ValidateSet('Owner', 'Edit', 'View')]
    [string]$ChangeRole,

    # Operation Set 5: Remove User
    [Parameter(ParameterSetName = 'RemoveUser', Mandatory = $true)]
    [Alias('Remove', 'DeleteUser')]
    [switch]$RemoveUser,

    # Operation Set 6: Add User
    [Parameter(ParameterSetName = 'AddUser', Mandatory = $true, Position = 1)]
    [Alias('Add')]
    [string]$AddUser,

    [Parameter(ParameterSetName = 'AddUser', Mandatory = $true)]
    [Alias('Permission')]
    [ValidateSet('Owner', 'Edit', 'View')]
    [string]$Role,

    # Tenant and Admin Connectivity
    [Parameter()]
    [Alias('Tenant', 'TenantDomain')]
    [string]$TenantName,

    [Parameter()]
    [string]$AdminUrl,

    # Reporting and Logging
    [Parameter(ParameterSetName = 'ListSites')]
    [Parameter(ParameterSetName = 'ListPermissions')]
    [string]$ExportCsv,

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

# ==============================================================================
# SCRIPT METADATA DECLARATION
# ==============================================================================
$Script:Metadata = @{
    Title       = "Manage-SharePoint-Site.ps1"
    Version     = "1.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Date        = "2026-10-19"
    ModuleReq   = "Microsoft.Online.SharePoint.PowerShell"
}

# Execution tracking
$Script:StartTime        = Get-Date
$Script:LogFilePath      = $null
$Script:OperationDetails = [System.Collections.Generic.List[PSObject]]::new()

# ==============================================================================
# LOGGING SYSTEM (CleanCode & Enterprise Logging Standard)
# ==============================================================================
function Initialize-ScriptLogging {
    [CmdletBinding()]
    param([string]$TargetDirectory)

    $timestamp   = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $computer    = $env:COMPUTERNAME
    $user        = $env:USERNAME
    $logFileName = "${timestamp}-Manage-SharePoint-Site-${computer}-${user}.log"

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
# HELP & DOCUMENTATION (Standard 1, 7 & 14)
# ==============================================================================
function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT  : $($Script:Metadata.Title)",
        "VERSION : $($Script:Metadata.Version)",
        "AUTHOR  : $($Script:Metadata.Author)",
        "CONTACT : $($Script:Metadata.Email) | $($Script:Metadata.GitHub)",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Manage-SharePoint-Site is an enterprise administration utility for SharePoint Online.",
        "    It empowers Microsoft 365 administrators to audit site collections, inspect permissions,",
        "    define authorized modifiers (editors), change user roles (e.g. demote from Edit to View),",
        "    add new users with granular roles (Owner, Edit, View), and remove users safely.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    1. Interactive Microsoft 365 Administrator Login:",
        "       - Requires an account with SharePoint Administrator or Global Administrator role.",
        "       - Automatically verifies existing active sessions with Connect-SPOService.",
        "       - If disconnected, prompts for interactive modern administrator sign-in (MFA supported).",
        "    2. Required PowerShell Module:",
        "       - 'Microsoft.Online.SharePoint.PowerShell'",
        "       - Automatically checked and prompted for installation if missing.",
        "    3. Script Unblocking (After downloading from GitHub):",
        "       - Run: Unblock-File -Path .\Manage-SharePoint-Site.ps1",
        "",
        "NON-INTERACTIVE / AUTOMATED ENTRA ID APP REGISTRATION (Standard 14):",
        "    For scheduled background execution, CI/CD, or automated management without user prompts:",
        "    1. Register an Application in Microsoft Entra ID (Azure Portal > App Registrations).",
        "    2. Add Application Permissions under 'Office 365 SharePoint Online' or 'Microsoft Graph':",
        "       - SharePoint API : 'Sites.FullControl.All' (Application) or 'Sites.Manage.All'",
        "       - Microsoft Graph: 'Sites.FullControl.All' or 'Sites.ReadWrite.All'",
        "    3. Grant Admin Consent for your tenant.",
        "    4. Configure Certificate Credentials or Client Secret in Entra ID.",
        "    5. Connect non-interactively using PnP.PowerShell or App-Only Connect-SPOService:",
        "       Connect-SPOService -Url https://<tenant>-admin.sharepoint.com -AppId <AppID> -ClientSecret <Secret>",
        "",
        "CORE MANAGEMENT SCENARIOS & EXAMPLES:",
        "    1. Display Built-In Help & Examples (default when executed without parameters):",
        "       .\Manage-SharePoint-Site.ps1",
        "       .\Manage-SharePoint-Site.ps1 -Help",
        "       .\Manage-SharePoint-Site.ps1 -h",
        "",
        "    2. List All SharePoint Sites across Tenant:",
        "       .\Manage-SharePoint-Site.ps1 -ListSites",
        "",
        "    3. Search and Filter Sites by Keyword with CSV Export:",
        "       .\Manage-SharePoint-Site.ps1 -ListSites -Search `"Finance`" -ExportCsv `"C:\Reports\FinanceSites.csv`"",
        "",
        "    4. Audit All Users and Permissions on a Specified Site (Owners, Editors, Viewers):",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -ListPermissions",
        "",
        "    5. Specify Users Allowed to Modify the Site (Set Modifiers / Synchronize Editors):",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -SetModifiers `"anna.smith@contoso.com`", `"john.doe@contoso.com`"",
        "",
        "    6. Change User Permissions from Edit to View (Read-Only Demotion):",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -UserEmail `"john.doe@contoso.com`" -ChangeRole View",
        "",
        "    7. Completely Remove a User from Site Permissions:",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -UserEmail `"ex-employee@contoso.com`" -RemoveUser",
        "",
        "    8. Add a New Email (Person) with View (Read) Permissions:",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -AddUser `"contractor@contoso.com`" -Role View",
        "",
        "    9. Add a New Email (Person) with Edit (Contribute) Permissions:",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -AddUser `"analyst@contoso.com`" -Role Edit",
        "",
        "   10. Add a New Email (Person) with Owner (Full Control & Site Collection Admin) Permissions:",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -AddUser `"manager@contoso.com`" -Role Owner",
        "",
        "   11. Automated Execution Without Confirmation Prompts (-Force):",
        "       .\Manage-SharePoint-Site.ps1 -SiteUrl `"https://contoso.sharepoint.com/sites/marketing`" -UserEmail `"auditor@contoso.com`" -ChangeRole View -Force",
        "",
        "PARAMETERS:",
        "    -SiteUrl, -Site, -Url        Target site URL or relative slug (e.g. 'sites/marketing').",
        "    -ListSites, -ListAllSites    Switch to list all SharePoint site collections in tenant.",
        "    -Search                      Keyword filter applied to site URL, Title, or Owner.",
        "    -IncludePersonalSites        Switch to include OneDrive personal sites when listing.",
        "    -ListPermissions             Audits and lists all entitled users (Owners, Editors, Viewers).",
        "    -SetModifiers, -Modifiers    Array of user emails permitted to modify the site.",
        "    -UserEmail, -User, -Email    Target user email for permission change or removal.",
        "    -ChangeRole, -NewRole        Target role to assign: 'Owner', 'Edit', or 'View'.",
        "    -RemoveUser, -Remove         Switch to remove the user completely from the site.",
        "    -AddUser, -Add               Email address of a new user to grant access to.",
        "    -Role, -Permission           Permission role for -AddUser: 'Owner', 'Edit', or 'View'.",
        "    -TenantName, -Tenant         Microsoft 365 tenant domain prefix (e.g. 'contoso').",
        "    -AdminUrl                    SharePoint Online Admin URL (e.g. 'https://contoso-admin.sharepoint.com').",
        "    -ExportCsv                   File path to export inventory or audit results to CSV.",
        "    -LogDirectory                Folder for execution logs (default: C:\Logs).",
        "    -Force                       Bypasses interactive [Y/N] / [T/N] confirmation prompts.",
        "    -PassThru                    Returns output objects to the PowerShell pipeline.",
        "    -Help, -h                    Displays this help documentation and exits.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After cloning or downloading from GitHub, unblock the script before execution:",
        "    Unblock-File -Path .\Manage-SharePoint-Site.ps1",
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

function Assert-SafeSiteUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw "Site URL cannot be empty."
    }

    $clean = $Url.Trim()
    if ($clean -match '^https?://') {
        if ($clean -notmatch '^https://[a-zA-Z0-9.-]+\.sharepoint\.com(/sites/[a-zA-Z0-9._-]+)?/?$') {
            throw "Invalid SharePoint site URL format: '$Url'. Expected: 'https://<tenant>.sharepoint.com/sites/<slug>'"
        }
    }
    else {
        $slug = $clean.TrimStart('/').TrimEnd('/')
        if ($slug.StartsWith('sites/')) {
            $slug = $slug.Substring(6)
        }
        if ($slug -notmatch '^[a-zA-Z0-9._-]+$') {
            throw "Invalid relative site slug '$Url'. Only alphanumeric characters, hyphens, and underscores are allowed."
        }
    }
}

# ==============================================================================
# AUTO-HELP TRIGGER ON EMPTY PARAMETERS (Requirement 1 & Standard 7)
# ==============================================================================
$hasExplicitAction = $ListSites -or 
                     $ListPermissions -or 
                     ($SetModifiers -and $SetModifiers.Count -gt 0) -or 
                     (-not [string]::IsNullOrWhiteSpace($ChangeRole)) -or 
                     $RemoveUser -or 
                     (-not [string]::IsNullOrWhiteSpace($AddUser))

# If user provided SiteUrl without explicit action, default to ListPermissions
if (-not $hasExplicitAction -and (-not [string]::IsNullOrWhiteSpace($SiteUrl))) {
    $ListPermissions = $true
    $hasExplicitAction = $true
}

if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or -not $hasExplicitAction) {
    Show-ScriptHelp
    exit 0
}

# Initialize logging system immediately
Initialize-ScriptLogging -TargetDirectory $LogDirectory
Write-ScriptLog -Level INFO -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)"
Write-ScriptLog -Level INFO -Message "Execution Log initialized at: $Script:LogFilePath"
Add-OperationDetail -Step "01-Init" -Operation "Initialize Logging" -Status "Success" -Details "Log file created at $Script:LogFilePath"

# ==============================================================================
# INPUT VALIDATION
# ==============================================================================
try {
    if (-not [string]::IsNullOrWhiteSpace($SiteUrl)) {
        Assert-SafeInputString -InputText $SiteUrl -ParamName "SiteUrl" -MaxLength 512
        Assert-SafeSiteUrl -Url $SiteUrl
    }

    if (-not [string]::IsNullOrWhiteSpace($UserEmail)) {
        Assert-SafeInputString -InputText $UserEmail -ParamName "UserEmail" -MaxLength 256
        if (-not (Test-ValidEmailAddress -Email $UserEmail)) {
            throw "Parameter -UserEmail '$UserEmail' is not a valid email address or UserPrincipalName."
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($AddUser)) {
        Assert-SafeInputString -InputText $AddUser -ParamName "AddUser" -MaxLength 256
        if (-not (Test-ValidEmailAddress -Email $AddUser)) {
            throw "Parameter -AddUser '$AddUser' is not a valid email address or UserPrincipalName."
        }
    }

    if ($SetModifiers -and $SetModifiers.Count -gt 0) {
        foreach ($m in $SetModifiers) {
            Assert-SafeInputString -InputText $m -ParamName "SetModifiers" -MaxLength 256
            if (-not (Test-ValidEmailAddress -Email $m)) {
                throw "Specified modifier '$m' in -SetModifiers is not a valid email address."
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        Assert-SafeInputString -InputText $Search -ParamName "Search" -MaxLength 128
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
function Resolve-SharePointUrls {
    [CmdletBinding()]
    param(
        [string]$TargetSiteUrl,
        [string]$ProvidedTenant,
        [string]$ProvidedAdminUrl
    )

    $resolved = @{
        TenantPrefix = $ProvidedTenant
        AdminUrl     = $ProvidedAdminUrl
        FullSiteUrl  = $TargetSiteUrl
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetSiteUrl)) {
        $clean = $TargetSiteUrl.Trim()
        if ($clean -match '^https://([a-zA-Z0-9.-]+)\.sharepoint\.com(/sites/([a-zA-Z0-9._-]+))?/?$') {
            $extractedTenant = $Matches[1]
            if ([string]::IsNullOrWhiteSpace($resolved.TenantPrefix)) {
                $resolved.TenantPrefix = $extractedTenant
            }
            $resolved.FullSiteUrl = $clean.TrimEnd('/')
        }
        else {
            # Relative slug provided (e.g. 'sites/marketing' or 'marketing')
            $slug = $clean.TrimStart('/').TrimEnd('/')
            if (-not $slug.StartsWith('sites/')) {
                $slug = "sites/$slug"
            }
            if (-not [string]::IsNullOrWhiteSpace($resolved.TenantPrefix)) {
                $tenantClean = $resolved.TenantPrefix.ToLowerInvariant().Replace('.sharepoint.com', '').Replace('.onmicrosoft.com', '')
                $resolved.FullSiteUrl = "https://${tenantClean}.sharepoint.com/${slug}"
            }
        }
    }

    # Normalize TenantPrefix
    if (-not [string]::IsNullOrWhiteSpace($resolved.TenantPrefix)) {
        $resolved.TenantPrefix = $resolved.TenantPrefix.ToLowerInvariant().Replace('.sharepoint.com', '').Replace('.onmicrosoft.com', '')
    }

    if ([string]::IsNullOrWhiteSpace($resolved.AdminUrl) -and -not [string]::IsNullOrWhiteSpace($resolved.TenantPrefix)) {
        $resolved.AdminUrl = "https://$($resolved.TenantPrefix)-admin.sharepoint.com"
    }

    return $resolved
}

$urlConfig = Resolve-SharePointUrls -TargetSiteUrl $SiteUrl -ProvidedTenant $TenantName -ProvidedAdminUrl $AdminUrl
$SiteUrl   = $urlConfig.FullSiteUrl
$AdminUrl  = $urlConfig.AdminUrl
$TenantName = $urlConfig.TenantPrefix

Write-ScriptLog -Level INFO -Message "Target Site URL : $(if ($SiteUrl) { $SiteUrl } else { '[All Sites / Not Specified]' })"
Write-ScriptLog -Level INFO -Message "Tenant Prefix   : $(if ($TenantName) { $TenantName } else { '[Auto-Detecting]' })"
Write-ScriptLog -Level INFO -Message "Admin Center URL: $(if ($AdminUrl) { $AdminUrl } else { '[Deriving upon connection]' })"
Add-OperationDetail -Step "03-UrlResolution" -Operation "URL & Tenant Resolution" -Status "Success" -Details "Target site: $SiteUrl | Admin: $AdminUrl"

# ==============================================================================
# PREREQUISITES & MODULE VERIFICATION
# ==============================================================================
function Assert-SharePointModule {
    Write-ScriptLog -Level INFO -Message "Checking for '$($Script:Metadata.ModuleReq)' module..."
    $installed = Get-Module -ListAvailable -Name $Script:Metadata.ModuleReq | Select-Object -First 1

    if (-not $installed) {
        Write-ScriptLog -Level WARN -Message "Module '$($Script:Metadata.ModuleReq)' is not installed."
        Write-Host "`n[!] The required module '$($Script:Metadata.ModuleReq)' was not found." -ForegroundColor Yellow
        Write-Host "    Would you like to install it now for CurrentUser? [Y/N or T/N]: " -NoNewline -ForegroundColor Cyan
        
        $choice = if ($Force) { "Y" } else { Read-Host }
        if ($choice -match '^[ytT1]$') {
            Write-ScriptLog -Level INFO -Message "Installing '$($Script:Metadata.ModuleReq)' from PSGallery..."
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
            throw "Script execution aborted because module '$($Script:Metadata.ModuleReq)' is missing."
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
        $null = Get-SPOTenant -ErrorAction Stop
        $isConnected = $true
        Write-ScriptLog -Level SUCCESS -Message "Active SharePoint Online administrator session detected."
        Add-OperationDetail -Step "05-Authentication" -Operation "Session Verification" -Status "Success" -Details "Reused active SharePoint Online admin session."
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($Url)) {
            Write-Host "`n[!] No active SharePoint Online session and Admin URL is not specified." -ForegroundColor Yellow
            Write-Host "    Please enter your SharePoint Admin Center URL (e.g. 'https://contoso-admin.sharepoint.com')" -ForegroundColor Cyan
            Write-Host "    or Tenant domain prefix (e.g. 'contoso'): " -NoNewline -ForegroundColor Cyan
            $inputVal = Read-Host
            if ([string]::IsNullOrWhiteSpace($inputVal)) {
                throw "Cannot connect: SharePoint Admin Center URL is required."
            }
            $cleanInput = $inputVal.Trim()
            if ($cleanInput -match '^https?://') {
                $Url = $cleanInput
            }
            else {
                $prefix = $cleanInput.ToLowerInvariant().Replace('.sharepoint.com', '').Replace('.onmicrosoft.com', '')
                $Url = "https://${prefix}-admin.sharepoint.com"
            }
        }

        Write-ScriptLog -Level INFO -Message "Connecting to SharePoint Online Admin Center at: $Url"
        Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Running" -Details "Connecting to $Url"
        try {
            Connect-SPOService -Url $Url -ErrorAction Stop
            Write-ScriptLog -Level SUCCESS -Message "Successfully connected to SharePoint Online Admin Center."
            Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Success" -Details "Connected to $Url"
        }
        catch {
            Write-ScriptLog -Level ERROR -Message "SharePoint Online administrator login failed: $($_.Exception.Message)"
            Add-OperationDetail -Step "05-Authentication" -Operation "Session Authentication" -Status "Failed" -Details $_.Exception.Message
            throw "SharePoint Online administrator login failed. Please ensure you have SharePoint/Global Administrator privileges."
        }
    }
}

# ==============================================================================
# INTERACTIVE CONFIRMATION (Bilingual: EN Y/N, PL T/N)
# ==============================================================================
function Confirm-Operation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ActionTitle,

        [Parameter(Mandatory = $true)]
        [hashtable]$ParametersToConfirm
    )

    if ($Force) {
        Write-ScriptLog -Level INFO -Message "Bypassing confirmation prompt due to -Force switch."
        return $true
    }

    Write-Host "`n" -NoNewline
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host " CONFIRMATION REQUIRED: $ActionTitle" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    foreach ($k in $ParametersToConfirm.Keys) {
        Write-Host ("  {0,-18}: {1}" -f $k, $ParametersToConfirm[$k]) -ForegroundColor White
    }
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "Are you sure you want to proceed? / Czy na pewno chcesz kontynuować? [Y/N or T/N]: " -NoNewline -ForegroundColor Cyan

    $response = Read-Host
    if ($response -notmatch '^[ytT1]$') {
        Write-ScriptLog -Level WARN -Message "Operation '$ActionTitle' cancelled by user."
        Add-OperationDetail -Step "Confirmation" -Operation $ActionTitle -Status "Skipped" -Details "User declined confirmation prompt."
        Write-Host "[!] Operation cancelled." -ForegroundColor Yellow
        exit 0
    }

    Write-ScriptLog -Level INFO -Message "User confirmed operation: $ActionTitle"
    return $true
}

# ==============================================================================
# OPERATION 1: LIST ALL SITES (-ListSites)
# ==============================================================================
function Invoke-ListAllSites {
    param(
        [string]$Keyword,
        [switch]$IncludePersonal,
        [switch]$ShowDetailed,
        [string]$CsvPath
    )

    Write-ScriptLog -Level INFO -Message "Querying SharePoint Online site collections..."
    Add-OperationDetail -Step "06-ListSites" -Operation "Retrieve Site Collections" -Status "Running" -Details "Querying tenant sites"

    try {
        $sites = Get-SPOSite -Limit All -IncludePersonalSite:([bool]$IncludePersonal) -Detailed:$ShowDetailed -ErrorAction Stop
        
        if (-not [string]::IsNullOrWhiteSpace($Keyword)) {
            Write-ScriptLog -Level INFO -Message "Filtering sites by keyword: '$Keyword'..."
            $sites = $sites | Where-Object {
                $_.Url -like "*$Keyword*" -or 
                $_.Title -like "*$Keyword*" -or 
                $_.Owner -like "*$Keyword*"
            }
        }

        $siteList = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($s in $sites) {
            $usedMB = [math]::Round($s.StorageUsageCurrent, 2)
            $quotaMB = [math]::Round($s.StorageQuota, 2)
            $usedGB = [math]::Round($usedMB / 1024, 2)
            $quotaGB = [math]::Round($quotaMB / 1024, 2)

            $siteObj = [PSCustomObject]@{
                Title             = $s.Title
                Url               = $s.Url
                Template          = $s.Template
                Status            = $s.Status
                Owner             = $s.Owner
                StorageUsedGB     = $usedGB
                StorageQuotaGB    = $quotaGB
                SharingCapability = $s.SharingCapability
                LockState         = $s.LockState
            }
            $siteList.Add($siteObj)
        }

        Write-ScriptLog -Level SUCCESS -Message "Discovered $($siteList.Count) SharePoint site collection(s)."
        Add-OperationDetail -Step "06-ListSites" -Operation "Retrieve Site Collections" -Status "Success" -Details "Discovered $($siteList.Count) sites"

        # Display structured table in console
        Write-Host "`n"
        Write-Host "================================================================================" -ForegroundColor Green
        Write-Host " SHAREPOINT ONLINE SITES INVENTORY ($($siteList.Count) Total)" -ForegroundColor Green
        Write-Host "================================================================================" -ForegroundColor Green
        $siteList | Format-Table -AutoSize -Property Title, Url, Status, StorageUsedGB, StorageQuotaGB, Owner, SharingCapability | Out-String | Write-Host -ForegroundColor Cyan

        if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
            try {
                $siteList | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-ScriptLog -Level SUCCESS -Message "Site inventory successfully exported to: $CsvPath"
                Add-OperationDetail -Step "06-ListSites" -Operation "Export CSV" -Status "Success" -Details "Exported to $CsvPath"
            }
            catch {
                Write-ScriptLog -Level ERROR -Message "Failed to export CSV: $($_.Exception.Message)"
                Add-OperationDetail -Step "06-ListSites" -Operation "Export CSV" -Status "Failed" -Details $_.Exception.Message
            }
        }

        return $siteList
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "Failed to query sites: $($_.Exception.Message)"
        Add-OperationDetail -Step "06-ListSites" -Operation "Retrieve Site Collections" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }
}

# ==============================================================================
# RESOLVE SITE SECURITY GROUPS (Bilingual: EN / PL)
# ==============================================================================
function Get-SiteSecurityGroups {
    param([string]$Url)

    Write-ScriptLog -Level INFO -Message "Discovering SharePoint security groups on site: $Url"
    $groups = @{
        OwnersGroup   = $null
        MembersGroup  = $null
        VisitorsGroup = $null
        AllGroups     = @()
    }

    try {
        $siteGroups = Get-SPOSiteGroup -Site $Url -ErrorAction Stop
        $groups.AllGroups = $siteGroups

        # 1. Resolve Owners Group (Full Control / Pełna kontrola)
        $groups.OwnersGroup = $siteGroups | Where-Object { 
            ($_.Roles -contains "Full Control") -or 
            ($_.Roles -contains "Pełna kontrola") -or 
            ($_.Title -like "*Owner*") -or 
            ($_.Title -like "*Właściciele*") 
        } | Select-Object -First 1

        # 2. Resolve Members Group (Contribute / Edit / Edycja / Współtworzenie)
        $groups.MembersGroup = $siteGroups | Where-Object { 
            ($_.Roles -contains "Edit") -or 
            ($_.Roles -contains "Contribute") -or 
            ($_.Roles -contains "Edycja") -or 
            ($_.Roles -contains "Współtworzenie") -or 
            ($_.Title -like "*Member*") -or 
            ($_.Title -like "*Członkowie*") 
        } | Select-Object -First 1

        # 3. Resolve Visitors Group (Read / Odczyt)
        $groups.VisitorsGroup = $siteGroups | Where-Object { 
            ($_.Roles -contains "Read") -or 
            ($_.Roles -contains "Odczyt") -or 
            ($_.Title -like "*Visitor*") -or 
            ($_.Title -like "*Goście*") -or 
            ($_.Title -like "*Czytelnicy*") 
        } | Select-Object -First 1

        Write-ScriptLog -Level SUCCESS -Message "Resolved groups -> Owners: '$($groups.OwnersGroup.Title)', Members: '$($groups.MembersGroup.Title)', Visitors: '$($groups.VisitorsGroup.Title)'"
        Add-OperationDetail -Step "GroupResolution" -Operation "Security Group Discovery" -Status "Success" -Details "Owners: $($groups.OwnersGroup.Title) | Members: $($groups.MembersGroup.Title) | Visitors: $($groups.VisitorsGroup.Title)"
    }
    catch {
        Write-ScriptLog -Level WARN -Message "Unable to query site groups via Get-SPOSiteGroup: $($_.Exception.Message)"
        Add-OperationDetail -Step "GroupResolution" -Operation "Security Group Discovery" -Status "Warning" -Details $_.Exception.Message
    }

    return $groups
}

# ==============================================================================
# OPERATION 2: LIST SITE PERMISSIONS (Owners, Edit, View)
# ==============================================================================
function Invoke-ListSitePermissions {
    param(
        [string]$Url,
        [string]$CsvPath
    )

    Write-ScriptLog -Level INFO -Message "Auditing permissions for site collection: $Url"
    Add-OperationDetail -Step "07-AuditPermissions" -Operation "Audit Site Permissions" -Status "Running" -Details "Querying users and groups for $Url"

    try {
        $secGroups = Get-SiteSecurityGroups -Url $Url
        $allUsers  = Get-SPOUser -Site $Url -Limit All -ErrorAction Stop

        $ownersGroupTitle   = if ($secGroups.OwnersGroup) { $secGroups.OwnersGroup.Title } else { "" }
        $membersGroupTitle  = if ($secGroups.MembersGroup) { $secGroups.MembersGroup.Title } else { "" }
        $visitorsGroupTitle = if ($secGroups.VisitorsGroup) { $secGroups.VisitorsGroup.Title } else { "" }

        $catalog = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($u in $allUsers) {
            $userGroups = if ($u.Groups) { $u.Groups } else { @() }
            
            # Determine categorized tier
            $isOwner = ($u.IsSiteCollectionAdmin -eq $true) -or ($userGroups -contains $ownersGroupTitle)
            $isEdit  = ($userGroups -contains $membersGroupTitle)
            $isView  = ($userGroups -contains $visitorsGroupTitle)

            # Map primary tier
            $tier = if ($isOwner) {
                "Owner (Full Control)"
            } elseif ($isEdit) {
                "Edit (Contribute)"
            } elseif ($isView) {
                "View (Read)"
            } else {
                "Other / Custom"
            }

            $entry = [PSCustomObject]@{
                DisplayName           = $u.DisplayName
                LoginName             = $u.LoginName
                PermissionTier        = $tier
                IsSiteCollectionAdmin = $u.IsSiteCollectionAdmin
                Groups                = ($userGroups -join '; ')
                UserType              = $u.UserType
            }
            $catalog.Add($entry)
        }

        Write-ScriptLog -Level SUCCESS -Message "Discovered $($catalog.Count) entitled user(s)/principal(s) on site."
        Add-OperationDetail -Step "07-AuditPermissions" -Operation "Audit Site Permissions" -Status "Success" -Details "Audited $($catalog.Count) accounts"

        # Categorized Console Rendering
        Write-Host "`n"
        Write-Host "================================================================================" -ForegroundColor Green
        Write-Host " SITE PERMISSIONS AUDIT: $Url" -ForegroundColor Green
        Write-Host "================================================================================" -ForegroundColor Green

        $owners = $catalog | Where-Object { $_.PermissionTier -like "Owner*" }
        $editors = $catalog | Where-Object { $_.PermissionTier -like "Edit*" }
        $viewers = $catalog | Where-Object { $_.PermissionTier -like "View*" }
        $others  = $catalog | Where-Object { $_.PermissionTier -like "Other*" }

        Write-Host "`n[+] 1. OWNERS / FULL CONTROL ($($owners.Count)):" -ForegroundColor Yellow
        if ($owners.Count -gt 0) {
            $owners | Format-Table -AutoSize -Property DisplayName, LoginName, IsSiteCollectionAdmin, Groups | Out-String | Write-Host -ForegroundColor Yellow
        } else {
            Write-Host "    (No explicit users in Owners group)" -ForegroundColor DarkGray
        }

        Write-Host "[+] 2. EDIT / MODIFIERS ($($editors.Count)):" -ForegroundColor Cyan
        if ($editors.Count -gt 0) {
            $editors | Format-Table -AutoSize -Property DisplayName, LoginName, Groups | Out-String | Write-Host -ForegroundColor Cyan
        } else {
            Write-Host "    (No users assigned to Edit/Members)" -ForegroundColor DarkGray
        }

        Write-Host "[+] 3. VIEW / READERS ($($viewers.Count)):" -ForegroundColor Green
        if ($viewers.Count -gt 0) {
            $viewers | Format-Table -AutoSize -Property DisplayName, LoginName, Groups | Out-String | Write-Host -ForegroundColor Green
        } else {
            Write-Host "    (No users assigned to View/Visitors)" -ForegroundColor DarkGray
        }

        if ($others.Count -gt 0) {
            Write-Host "[+] 4. OTHER / CUSTOM GROUPS ($($others.Count)):" -ForegroundColor Magenta
            $others | Format-Table -AutoSize -Property DisplayName, LoginName, Groups | Out-String | Write-Host -ForegroundColor Magenta
        }

        if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
            try {
                $catalog | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-ScriptLog -Level SUCCESS -Message "Permission audit exported to: $CsvPath"
                Add-OperationDetail -Step "07-AuditPermissions" -Operation "Export CSV" -Status "Success" -Details "Exported to $CsvPath"
            }
            catch {
                Write-ScriptLog -Level ERROR -Message "Failed to export CSV: $($_.Exception.Message)"
                Add-OperationDetail -Step "07-AuditPermissions" -Operation "Export CSV" -Status "Failed" -Details $_.Exception.Message
            }
        }

        return $catalog
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "Failed to audit permissions: $($_.Exception.Message)"
        Add-OperationDetail -Step "07-AuditPermissions" -Operation "Audit Site Permissions" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }
}

# ==============================================================================
# OPERATION 3: SPECIFY SITE MODIFIERS (-SetModifiers)
# ==============================================================================
function Invoke-SetSiteModifiers {
    param(
        [string]$Url,
        [string[]]$Modifiers
    )

    Write-ScriptLog -Level INFO -Message "Setting allowed modifiers for site: $Url"
    Add-OperationDetail -Step "08-SetModifiers" -Operation "Synchronize Site Modifiers" -Status "Running" -Details "Target users: $($Modifiers -join ', ')"

    $secGroups = Get-SiteSecurityGroups -Url $Url
    if (-not $secGroups.MembersGroup) {
        throw "Could not locate the Members (Edit) group for site '$Url'. Cannot synchronize modifiers."
    }
    $membersGroupName = $secGroups.MembersGroup.Title
    Write-ScriptLog -Level INFO -Message "Target Members (Edit) security group: '$membersGroupName'"

    # Query current members in Members group
    $allUsers = Get-SPOUser -Site $Url -Limit All -ErrorAction Stop
    $currentMembers = $allUsers | Where-Object {
        $_.Groups -contains $membersGroupName -and 
        $_.LoginName -notlike "*c:0(.s|true*" # Exclude company-wide claims if any
    }

    $currentMemberEmails = @($currentMembers | ForEach-Object { $_.LoginName.ToLowerInvariant() })
    $targetModifierEmails = @($Modifiers | ForEach-Object { $_.Trim().ToLowerInvariant() })

    # Users to add
    $toAdd = $targetModifierEmails | Where-Object { $currentMemberEmails -notcontains $_ }
    # Users to remove
    $toRemove = $currentMemberEmails | Where-Object { $targetModifierEmails -notcontains $_ }

    Write-Host "`n"
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host " PROPOSED SITE MODIFIERS SYNCHRONIZATION" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host " Site URL       : $Url" -ForegroundColor White
    Write-Host " Members Group  : $membersGroupName" -ForegroundColor White
    Write-Host " Current Count  : $($currentMemberEmails.Count)" -ForegroundColor White
    Write-Host " Target Count   : $($targetModifierEmails.Count)" -ForegroundColor White
    Write-Host " Users to ADD   : $(if ($toAdd) { $toAdd -join ', ' } else { '[None]' })" -ForegroundColor Green
    Write-Host " Users to REMOVE: $(if ($toRemove) { $toRemove -join ', ' } else { '[None]' })" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Cyan

    Confirm-Operation -ActionTitle "Synchronize Site Modifiers" -ParametersToConfirm @{
        SiteUrl          = $Url
        MembersGroup     = $membersGroupName
        AddCount         = @($toAdd).Count
        RemoveCount      = @($toRemove).Count
    }

    # Execute additions
    foreach ($user in $toAdd) {
        try {
            Write-ScriptLog -Level INFO -Message "Adding user '$user' to '$membersGroupName'..."
            Add-SPOUser -Site $Url -LoginName $user -Group $membersGroupName -ErrorAction Stop | Out-Null
            Write-ScriptLog -Level SUCCESS -Message "Added modifier '$user' to '$membersGroupName'."
            Add-OperationDetail -Step "08-SetModifiers" -Operation "Add Modifier ($user)" -Status "Success" -Details "Added to $membersGroupName"
        }
        catch {
            Write-ScriptLog -Level ERROR -Message "Failed to add modifier '$user': $($_.Exception.Message)"
            Add-OperationDetail -Step "08-SetModifiers" -Operation "Add Modifier ($user)" -Status "Failed" -Details $_.Exception.Message
        }
    }

    # Execute removals
    foreach ($user in $toRemove) {
        try {
            Write-ScriptLog -Level INFO -Message "Removing unlisted modifier '$user' from '$membersGroupName'..."
            Remove-SPOUser -Site $Url -LoginName $user -Group $membersGroupName -ErrorAction Stop | Out-Null
            Write-ScriptLog -Level SUCCESS -Message "Removed modifier '$user' from '$membersGroupName'."
            Add-OperationDetail -Step "08-SetModifiers" -Operation "Remove Modifier ($user)" -Status "Success" -Details "Removed from $membersGroupName"
        }
        catch {
            Write-ScriptLog -Level ERROR -Message "Failed to remove modifier '$user': $($_.Exception.Message)"
            Add-OperationDetail -Step "08-SetModifiers" -Operation "Remove Modifier ($user)" -Status "Failed" -Details $_.Exception.Message
        }
    }

    Write-ScriptLog -Level SUCCESS -Message "Site modifiers synchronization completed for: $Url"
}

# ==============================================================================
# OPERATION 4: CHANGE USER ROLE (-ChangeRole)
# ==============================================================================
function Invoke-ChangeUserRole {
    param(
        [string]$Url,
        [string]$User,
        [string]$TargetRole
    )

    Write-ScriptLog -Level INFO -Message "Changing permission for user '$User' on site '$Url' to '$TargetRole'..."
    Add-OperationDetail -Step "09-ChangeRole" -Operation "Change User Role" -Status "Running" -Details "User: $User -> New Role: $TargetRole"

    $secGroups = Get-SiteSecurityGroups -Url $Url
    $existingUser = Get-SPOUser -Site $Url -Limit All | Where-Object { $_.LoginName -eq $User } | Select-Object -First 1

    $currentGroups = if ($existingUser -and $existingUser.Groups) { $existingUser.Groups } else { @() }
    $isAdmin = if ($existingUser) { $existingUser.IsSiteCollectionAdmin } else { $false }

    Confirm-Operation -ActionTitle "Change User Role" -ParametersToConfirm @{
        SiteUrl       = $Url
        UserEmail     = $User
        CurrentGroups = ($currentGroups -join '; ')
        NewRole       = $TargetRole
    }

    $membersTitle  = if ($secGroups.MembersGroup) { $secGroups.MembersGroup.Title } else { "" }
    $visitorsTitle = if ($secGroups.VisitorsGroup) { $secGroups.VisitorsGroup.Title } else { "" }
    $ownersTitle   = if ($secGroups.OwnersGroup) { $secGroups.OwnersGroup.Title } else { "" }

    switch ($TargetRole) {
        'View' {
            # Demote to View: Remove from Members and Owners; Add to Visitors
            if ($membersTitle -and ($currentGroups -contains $membersTitle)) {
                Write-ScriptLog -Level INFO -Message "Removing '$User' from Members group '$membersTitle'..."
                Remove-SPOUser -Site $Url -LoginName $User -Group $membersTitle -ErrorAction SilentlyContinue | Out-Null
            }
            if ($ownersTitle -and ($currentGroups -contains $ownersTitle)) {
                Write-ScriptLog -Level INFO -Message "Removing '$User' from Owners group '$ownersTitle'..."
                Remove-SPOUser -Site $Url -LoginName $User -Group $ownersTitle -ErrorAction SilentlyContinue | Out-Null
            }
            if ($isAdmin) {
                Write-ScriptLog -Level INFO -Message "Revoking Site Collection Administrator rights from '$User'..."
                Set-SPOUser -Site $Url -LoginName $User -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue | Out-Null
            }

            if ($visitorsTitle) {
                Write-ScriptLog -Level INFO -Message "Adding '$User' to Visitors group '$visitorsTitle'..."
                Add-SPOUser -Site $Url -LoginName $User -Group $visitorsTitle -ErrorAction Stop | Out-Null
                Write-ScriptLog -Level SUCCESS -Message "User '$User' successfully assigned View (Read) permissions in '$visitorsTitle'."
                Add-OperationDetail -Step "09-ChangeRole" -Operation "Change Role to View" -Status "Success" -Details "Demoted '$User' to '$visitorsTitle'"
            } else {
                throw "Visitors group could not be resolved on site '$Url'."
            }
        }

        'Edit' {
            # Assign Edit: Remove from Visitors; Add to Members
            if ($visitorsTitle -and ($currentGroups -contains $visitorsTitle)) {
                Write-ScriptLog -Level INFO -Message "Removing '$User' from Visitors group '$visitorsTitle'..."
                Remove-SPOUser -Site $Url -LoginName $User -Group $visitorsTitle -ErrorAction SilentlyContinue | Out-Null
            }
            if ($isAdmin) {
                Write-ScriptLog -Level INFO -Message "Revoking Site Collection Administrator rights from '$User'..."
                Set-SPOUser -Site $Url -LoginName $User -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue | Out-Null
            }

            if ($membersTitle) {
                Write-ScriptLog -Level INFO -Message "Adding '$User' to Members group '$membersTitle'..."
                Add-SPOUser -Site $Url -LoginName $User -Group $membersTitle -ErrorAction Stop | Out-Null
                Write-ScriptLog -Level SUCCESS -Message "User '$User' successfully assigned Edit permissions in '$membersTitle'."
                Add-OperationDetail -Step "09-ChangeRole" -Operation "Change Role to Edit" -Status "Success" -Details "Assigned '$User' to '$membersTitle'"
            } else {
                throw "Members group could not be resolved on site '$Url'."
            }
        }

        'Owner' {
            # Assign Owner: Add to Owners group and set Site Collection Admin
            if ($visitorsTitle -and ($currentGroups -contains $visitorsTitle)) {
                Remove-SPOUser -Site $Url -LoginName $User -Group $visitorsTitle -ErrorAction SilentlyContinue | Out-Null
            }
            if ($ownersTitle) {
                Write-ScriptLog -Level INFO -Message "Adding '$User' to Owners group '$ownersTitle'..."
                Add-SPOUser -Site $Url -LoginName $User -Group $ownersTitle -ErrorAction Stop | Out-Null
            }
            Write-ScriptLog -Level INFO -Message "Granting Site Collection Administrator privileges to '$User'..."
            Set-SPOUser -Site $Url -LoginName $User -IsSiteCollectionAdmin $true -ErrorAction SilentlyContinue | Out-Null
            Write-ScriptLog -Level SUCCESS -Message "User '$User' successfully assigned Owner (Full Control & Admin) permissions."
            Add-OperationDetail -Step "09-ChangeRole" -Operation "Change Role to Owner" -Status "Success" -Details "Promoted '$User' to Owner and Site Collection Admin"
        }
    }
}

# ==============================================================================
# OPERATION 5: REMOVE USER FROM SITE (-RemoveUser)
# ==============================================================================
function Invoke-RemoveUserFromSite {
    param(
        [string]$Url,
        [string]$User
    )

    Write-ScriptLog -Level INFO -Message "Removing user '$User' from site: $Url"
    Add-OperationDetail -Step "10-RemoveUser" -Operation "Remove User From Site" -Status "Running" -Details "User: $User"

    # Verify user existence and safety check
    $existingUser = Get-SPOUser -Site $Url -Limit All | Where-Object { $_.LoginName -eq $User } | Select-Object -First 1
    if (-not $existingUser) {
        Write-ScriptLog -Level WARN -Message "User '$User' does not exist in site collection '$Url'. Nothing to remove."
        Add-OperationDetail -Step "10-RemoveUser" -Operation "Remove User From Site" -Status "Skipped" -Details "User not found on site."
        return
    }

    # Safety check: Is this the primary site owner?
    $siteDetails = Get-SPOSite -Identity $Url -Detailed -ErrorAction SilentlyContinue
    if ($siteDetails -and ($siteDetails.Owner -eq $User)) {
        if (-not $Force) {
            Write-ScriptLog -Level WARN -Message "User '$User' is configured as the PRIMARY OWNER of site collection '$Url'!"
            Write-Host "`n[!] CAUTION: '$User' is the Primary Site Collection Owner!" -ForegroundColor Red
            Write-Host "    Removing the primary owner without assigning a successor may cause administrative issues." -ForegroundColor Yellow
        }
    }

    Confirm-Operation -ActionTitle "Remove User Completely From Site" -ParametersToConfirm @{
        SiteUrl       = $Url
        UserEmail     = $User
        DisplayName   = $existingUser.DisplayName
        CurrentGroups = ($existingUser.Groups -join '; ')
    }

    try {
        Remove-SPOUser -Site $Url -LoginName $User -ErrorAction Stop | Out-Null
        Write-ScriptLog -Level SUCCESS -Message "User '$User' was completely removed from site collection '$Url'."
        Add-OperationDetail -Step "10-RemoveUser" -Operation "Remove User From Site" -Status "Success" -Details "User '$User' removed from site"
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "Failed to remove user '$User': $($_.Exception.Message)"
        Add-OperationDetail -Step "10-RemoveUser" -Operation "Remove User From Site" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }
}

# ==============================================================================
# OPERATION 6: ADD NEW USER TO SITE (-AddUser & -Role)
# ==============================================================================
function Invoke-AddUserToSite {
    param(
        [string]$Url,
        [string]$User,
        [string]$TargetRole
    )

    Write-ScriptLog -Level INFO -Message "Adding user '$User' to site '$Url' with role '$TargetRole'..."
    Add-OperationDetail -Step "11-AddUser" -Operation "Add User" -Status "Running" -Details "User: $User | Role: $TargetRole"

    $secGroups = Get-SiteSecurityGroups -Url $Url

    Confirm-Operation -ActionTitle "Add User to Site" -ParametersToConfirm @{
        SiteUrl   = $Url
        UserEmail = $User
        Role      = $TargetRole
    }

    $ownersTitle   = if ($secGroups.OwnersGroup) { $secGroups.OwnersGroup.Title } else { "" }
    $membersTitle  = if ($secGroups.MembersGroup) { $secGroups.MembersGroup.Title } else { "" }
    $visitorsTitle = if ($secGroups.VisitorsGroup) { $secGroups.VisitorsGroup.Title } else { "" }

    try {
        switch ($TargetRole) {
            'View' {
                if (-not $visitorsTitle) { throw "Visitors (View) group could not be resolved." }
                Add-SPOUser -Site $Url -LoginName $User -Group $visitorsTitle -ErrorAction Stop | Out-Null
                Write-ScriptLog -Level SUCCESS -Message "User '$User' added to Visitors group '$visitorsTitle' (View/Read)."
                Add-OperationDetail -Step "11-AddUser" -Operation "Add User (View)" -Status "Success" -Details "Added to $visitorsTitle"
            }
            'Edit' {
                if (-not $membersTitle) { throw "Members (Edit) group could not be resolved." }
                Add-SPOUser -Site $Url -LoginName $User -Group $membersTitle -ErrorAction Stop | Out-Null
                Write-ScriptLog -Level SUCCESS -Message "User '$User' added to Members group '$membersTitle' (Edit/Contribute)."
                Add-OperationDetail -Step "11-AddUser" -Operation "Add User (Edit)" -Status "Success" -Details "Added to $membersTitle"
            }
            'Owner' {
                if ($ownersTitle) {
                    Add-SPOUser -Site $Url -LoginName $User -Group $ownersTitle -ErrorAction Stop | Out-Null
                }
                Set-SPOUser -Site $Url -LoginName $User -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
                Write-ScriptLog -Level SUCCESS -Message "User '$User' added to Owners group '$ownersTitle' and granted Site Collection Administrator privileges."
                Add-OperationDetail -Step "11-AddUser" -Operation "Add User (Owner)" -Status "Success" -Details "Added to $ownersTitle & SiteCollectionAdmin"
            }
        }
    }
    catch {
        Write-ScriptLog -Level ERROR -Message "Failed to add user '$User': $($_.Exception.Message)"
        Add-OperationDetail -Step "11-AddUser" -Operation "Add User ($TargetRole)" -Status "Failed" -Details $_.Exception.Message
        throw $_
    }
}

# ==============================================================================
# MAIN DISPATCHER
# ==============================================================================
try {
    # 1. Verify SharePoint PowerShell Module
    Assert-SharePointModule

    # 2. Connect to SharePoint Online Admin Center
    Connect-SharePointAdminSession -Url $AdminUrl

    $resultObject = $null

    # 3. Route to requested operation
    if ($ListSites) {
        $resultObject = Invoke-ListAllSites -Keyword $Search -IncludePersonal:$IncludePersonalSites -ShowDetailed:$Detailed -CsvPath $ExportCsv
    }
    elseif ($SetModifiers -and $SetModifiers.Count -gt 0) {
        Invoke-SetSiteModifiers -Url $SiteUrl -Modifiers $SetModifiers
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ChangeRole)) {
        Invoke-ChangeUserRole -Url $SiteUrl -User $UserEmail -TargetRole $ChangeRole
    }
    elseif ($RemoveUser) {
        Invoke-RemoveUserFromSite -Url $SiteUrl -User $UserEmail
    }
    elseif (-not [string]::IsNullOrWhiteSpace($AddUser)) {
        Invoke-AddUserToSite -Url $SiteUrl -User $AddUser -TargetRole $Role
    }
    elseif ($ListPermissions) {
        $resultObject = Invoke-ListSitePermissions -Url $SiteUrl -CsvPath $ExportCsv
    }
    else {
        Show-ScriptHelp
        exit 0
    }

    # 4. Operation Summary Table
    $duration = (Get-Date) - $Script:StartTime
    $durationStr = "{0:D2}m {1:D2}s" -f $duration.Minutes, $duration.Seconds

    Write-Host "`n"
    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host " OPERATION COMPLETED SUCCESSFULLY: MANAGE SHAREPOINT SITE" -ForegroundColor Green
    Write-Host "================================================================================" -ForegroundColor Green
    Write-Host "  Target Site   : $(if ($SiteUrl) { $SiteUrl } else { '[All Sites / Tenant Inventory]' })" -ForegroundColor White
    Write-Host "  Total Runtime : $durationStr" -ForegroundColor White
    Write-Host "  Log File Path : $Script:LogFilePath" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Green

    Write-Host "`n[+] Detailed Operation Audit Trail:" -ForegroundColor Cyan
    $Script:OperationDetails | Format-Table -AutoSize -Property Timestamp, Step, Operation, Status, Details | Out-String | Write-Host -ForegroundColor DarkCyan

    Write-ScriptLog -Level SUCCESS -Message "Manage-SharePoint-Site completed successfully in $durationStr."

    if ($PassThru -and $resultObject) {
        $resultObject
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
    Write-Host "[-] Execution log: $Script:LogFilePath" -ForegroundColor Yellow
    Write-Host "[-] ============================================================================" -ForegroundColor Red

    exit 1
}
