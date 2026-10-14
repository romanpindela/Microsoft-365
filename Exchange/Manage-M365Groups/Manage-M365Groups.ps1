<#
.SYNOPSIS
    Manages Microsoft 365 and Exchange Online groups: lists groups with rich details (including Microsoft Teams integration status), adds members/owners, and removes members/owners.

.DESCRIPTION
    Production-ready automation script designed for Microsoft 365 and Exchange Online administrators
    to audit and manage tenant groups with enterprise-grade safety, input validation, and interactive
    or automated execution.

    Key capabilities include:
    1. Automatic Exchange Online administrator session detection and modern interactive login.
    2. Comprehensive group listing across all types:
       - Microsoft 365 Groups (Unified Groups)
       - Microsoft Teams-connected groups (auto-detected via ResourceProvisioningOptions)
       - Distribution Lists (Distribution Groups)
       - Mail-Enabled Security Groups
       - Dynamic Distribution Groups
    3. In-depth group inspection:
       - Teams enablement status (Connected Team vs. Standalone)
       - Group type, privacy/access type (Public vs. Private)
       - Primary SMTP, alias, and proxy email addresses
       - Owners/Managers list and total count
       - Members list, roles, recipient types, and total count
       - SharePoint Site URL and document library links
       - External senders authorization and GAL visibility
       - Dynamic RecipientFilter rules (for Dynamic Distribution Groups)
    4. Member and Owner assignment (-AddMember):
       - Assign users as Members or Owners
       - Pre-flight identity verification (validates both user and group existence)
       - Duplicate membership prevention
    5. Member and Owner removal (-RemoveMember):
       - Remove users from Members, Owners, or All roles
       - Orphan prevention: prevents removing the sole owner of a Microsoft 365 Group
    6. Safety and automation controls:
       - Interactive confirmation prompts (bilingual support: English Y/N, Polish T/N)
       - Non-interactive automation support via -Force
       - CSV reporting support via -ExportCsv

.PARAMETER ListGroups
    Switch parameter. Lists groups across the Microsoft 365 tenant matching specified filters.
    Aliases: -List, -l.

.PARAMETER GroupIdentity
    The identity (DisplayName, Primary SMTP email address, Alias, or GUID) of the target group.
    When specified alone, displays comprehensive details and membership of that group.
    Aliases: -Identity, -Group, -GroupName, -g.

.PARAMETER UserIdentity
    The identity (UPN, Primary SMTP email address, or Alias) of the user to add or remove.
    Required when -AddMember or -RemoveMember is specified.
    Aliases: -User, -Member, -u, -UserPrincipalName.

.PARAMETER AddMember
    Switch parameter. Adds the specified user to the target group.
    Aliases: -Add, -Assign.

.PARAMETER RemoveMember
    Switch parameter. Removes the specified user from the target group.
    Aliases: -Remove, -Revoke.

.PARAMETER Role
    Specifies the role when adding or removing group members.
    Valid values: 'Member' (default), 'Owner', or 'All' (removal only).

.PARAMETER GroupType
    Filters groups by type when listing.
    Valid values: 'All' (default), 'Microsoft365', 'Teams', 'DistributionList', 'MailEnabledSecurity', 'DynamicDistribution'.
    Aliases: -Type.

.PARAMETER Search
    Optional string filter for matching DisplayName, PrimarySmtpAddress, or Alias.
    Aliases: -Filter, -NameFilter.

.PARAMETER Detailed
    Switch parameter. Displays in-depth configuration and membership details for listed groups.
    Aliases: -d, -ShowMembers.

.PARAMETER ExportCsv
    Optional file path to export the retrieved group data into a UTF-8 CSV report.
    Aliases: -CsvPath, -Export.

.PARAMETER PassThru
    Switch parameter. Outputs custom PowerShell objects to the pipeline for downstream processing.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used to pre-fill the Microsoft 365 Exchange Online login prompt.
    Aliases: -AdminUPN.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N) for unattended automation.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Manage-M365Groups.ps1
    Running without parameters displays the built-in help banner, examples, and version info.

.EXAMPLE
    .\Manage-M365Groups.ps1 -Help
    Displays comprehensive help and usage documentation.

.EXAMPLE
    .\Manage-M365Groups.ps1 -ListGroups
    Lists all groups across the tenant in a formatted summary table.

.EXAMPLE
    .\Manage-M365Groups.ps1 -ListGroups -GroupType Teams
    Lists only Microsoft Teams-enabled groups.

.EXAMPLE
    .\Manage-M365Groups.ps1 -ListGroups -Search "Sales" -Detailed
    Lists groups matching "Sales" with in-depth details, members, and SharePoint links.

.EXAMPLE
    .\Manage-M365Groups.ps1 -GroupIdentity "finance@domain.com"
    Inspects the "finance@domain.com" group in detail, showing all members, owners, and Teams status.

.EXAMPLE
    .\Manage-M365Groups.ps1 -GroupIdentity "it-team@domain.com" -UserIdentity "john.doe@domain.com" -AddMember
    Adds john.doe@domain.com as a member to the it-team group with interactive confirmation.

.EXAMPLE
    .\Manage-M365Groups.ps1 -GroupIdentity "it-team@domain.com" -UserIdentity "john.doe@domain.com" -AddMember -Role Owner -Force
    Adds john.doe@domain.com as an Owner silently without confirmation prompts.

.EXAMPLE
    .\Manage-M365Groups.ps1 -GroupIdentity "it-team@domain.com" -UserIdentity "john.doe@domain.com" -RemoveMember
    Removes john.doe@domain.com from the it-team group with interactive confirmation.

.EXAMPLE
    .\Manage-M365Groups.ps1 -GroupIdentity "it-team@domain.com" -UserIdentity "john.doe@domain.com" -RemoveMember -Force
    Removes john.doe@domain.com silently without confirmation prompts.

.EXAMPLE
    .\Manage-M365Groups.ps1 -ListGroups -ExportCsv "C:\Reports\M365_Groups.csv"
    Exports all tenant groups into a UTF-8 CSV report.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Action Switches
    [Parameter(ParameterSetName = 'ListGroups')]
    [Alias('List', 'l')]
    [switch]$ListGroups,

    [Parameter(ParameterSetName = 'AddMember')]
    [Alias('Add', 'Assign')]
    [switch]$AddMember,

    [Parameter(ParameterSetName = 'RemoveMember')]
    [Alias('Remove', 'Revoke')]
    [switch]$RemoveMember,

    # Target Group Identity
    [Parameter(ParameterSetName = 'Default', Position = 0)]
    [Parameter(ParameterSetName = 'ListGroups', Position = 0)]
    [Parameter(ParameterSetName = 'AddMember', Position = 0)]
    [Parameter(ParameterSetName = 'RemoveMember', Position = 0)]
    [Alias('Identity', 'Group', 'GroupName', 'g')]
    [string]$GroupIdentity,

    # Target User Identity
    [Parameter(ParameterSetName = 'AddMember', Position = 1)]
    [Parameter(ParameterSetName = 'RemoveMember', Position = 1)]
    [Alias('User', 'Member', 'u', 'UserPrincipalName')]
    [string]$UserIdentity,

    # Role (Member, Owner, or All)
    [Parameter(ParameterSetName = 'AddMember')]
    [Parameter(ParameterSetName = 'RemoveMember')]
    [ValidateSet('Member', 'Owner', 'All')]
    [string]$Role = 'Member',

    # Filters for Listing
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ListGroups')]
    [ValidateSet('All', 'Microsoft365', 'Teams', 'DistributionList', 'MailEnabledSecurity', 'DynamicDistribution')]
    [Alias('Type')]
    [string]$GroupType = 'All',

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ListGroups')]
    [Alias('Filter', 'NameFilter')]
    [string]$Search,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ListGroups')]
    [Alias('d', 'ShowMembers')]
    [switch]$Detailed,

    # Export
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ListGroups')]
    [Alias('CsvPath', 'Export')]
    [string]$ExportCsv,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ListGroups')]
    [switch]$PassThru,

    # Administrative and Session Options
    [Parameter()]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

    [Parameter()]
    [switch]$Force,

    # Help Switch
    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata declaration
$ScriptMetadata = @{
    Title   = "Manage-M365Groups.ps1"
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
        "    administrators to audit, list, and manage groups and their memberships.",
        "    Supports Microsoft 365 Groups (with Teams detection), Distribution Lists,",
        "    Mail-Enabled Security Groups, and Dynamic Distribution Groups.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    - Requires PowerShell 5.1 or PowerShell 7+.",
        "    - Requires 'ExchangeOnlineManagement' PowerShell module (auto-installed if missing).",
        "    - Requires an Exchange Online administrator role (Exchange Administrator or Global Admin).",
        "    - Automatically detects active sessions; if not connected, prompts for administrator sign-in.",
        "    - Optional -AdminUserPrincipalName can pre-populate the administrator identity.",
        "",
        "USAGE EXAMPLES:",
        "    1. Display Built-In Help & Info (default behavior when run without parameters):",
        "       .\Manage-M365Groups.ps1",
        "       .\Manage-M365Groups.ps1 -h",
        "",
        "    2. List All Groups Across Tenant (Summary Table):",
        "       .\Manage-M365Groups.ps1 -ListGroups",
        "",
        "    3. List Only Microsoft Teams Groups:",
        "       .\Manage-M365Groups.ps1 -ListGroups -GroupType Teams",
        "",
        "    4. List Groups with Wildcard Filter & In-Depth Details (Members & Owners):",
        "       .\Manage-M365Groups.ps1 -ListGroups -Search `"Sales`" -Detailed",
        "",
        "    5. View Comprehensive Details for a Specific Group (Members, Owners, Teams Status):",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"sales@domain.com`"",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"Sales Department`" -Detailed",
        "",
        "    6. Assign User to Group as Member (Interactive Confirmation):",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"sales@domain.com`" -UserIdentity `"user@domain.com`" -AddMember",
        "",
        "    7. Assign User as Owner Silently (-Force):",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"sales@domain.com`" -UserIdentity `"admin@domain.com`" -AddMember -Role Owner -Force",
        "",
        "    8. Remove User from Group (Interactive Confirmation):",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"sales@domain.com`" -UserIdentity `"user@domain.com`" -RemoveMember",
        "",
        "    9. Remove User from Group Silently (-Force):",
        "       .\Manage-M365Groups.ps1 -GroupIdentity `"sales@domain.com`" -UserIdentity `"user@domain.com`" -RemoveMember -Force",
        "",
        "   10. Export Tenant Groups to UTF-8 CSV Report:",
        "       .\Manage-M365Groups.ps1 -ListGroups -ExportCsv `"C:\Reports\M365Groups.csv`"",
        "",
        "PARAMETERS:",
        "    -ListGroups, -List, -l     Lists groups matching specified filters.",
        "    -GroupIdentity, -Group, -g Target group DisplayName, SMTP email address, Alias, or GUID.",
        "    -UserIdentity, -User, -u   Target user UPN, email address, or Alias to add or remove.",
        "    -AddMember, -Add, -Assign  Switch to add the user to the target group.",
        "    -RemoveMember, -Remove     Switch to remove the user from the target group.",
        "    -Role                      Role to manage: 'Member' (default), 'Owner', or 'All'.",
        "    -GroupType, -Type          Filter by type: 'All', 'Microsoft365', 'Teams', 'DistributionList',",
        "                               'MailEnabledSecurity', 'DynamicDistribution'.",
        "    -Search, -Filter           Wildcard pattern to search group names, emails, or aliases.",
        "    -Detailed, -d              Displays in-depth configuration, owners, and members.",
        "    -ExportCsv, -Export        Exports results to the specified CSV file path.",
        "    -PassThru                  Outputs objects to the PowerShell pipeline.",
        "    -AdminUserPrincipalName    (Optional) Administrator UPN for Exchange Online connection.",
        "    -Force                     Bypasses interactive confirmation prompts (Y/N).",
        "    -Help, -h                  Displays this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Manage-M365Groups.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Auto-show help when executed without parameters, or with -Help / -h
$hasExplicitAction = $ListGroups -or $AddMember -or $RemoveMember -or (-not [string]::IsNullOrWhiteSpace($GroupIdentity)) -or (-not [string]::IsNullOrWhiteSpace($Search))
if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or -not $hasExplicitAction) {
    Show-ScriptHelp
    exit 0
}

# Defensive input validation
function Test-ValidEmailAddress {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    $emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return ($Email.Trim() -match $emailPattern)
}

function Assert-SafeInputString {
    param([string]$InputText, [string]$ParamName)
    if ([string]::IsNullOrWhiteSpace($InputText)) { return }
    if ($InputText.Length -gt 256) {
        Write-Error "Parameter -$ParamName exceeds maximum allowed length of 256 characters."
        exit 1
    }
    # Defense against control characters and malicious shell injections
    if ($InputText -match '[`$<>|;&\0]') {
        Write-Error "Parameter -$ParamName contains invalid or potentially dangerous characters."
        exit 1
    }
}

# Validate inputs
Assert-SafeInputString -InputText $GroupIdentity -ParamName "GroupIdentity"
Assert-SafeInputString -InputText $UserIdentity -ParamName "UserIdentity"
Assert-SafeInputString -InputText $Search -ParamName "Search"

# Validate action prerequisites
if ($AddMember -or $RemoveMember) {
    if ([string]::IsNullOrWhiteSpace($GroupIdentity)) {
        Write-Error "Missing parameter: -GroupIdentity is required when adding or removing members. Run with -h for help."
        exit 1
    }
    if ([string]::IsNullOrWhiteSpace($UserIdentity)) {
        Write-Error "Missing parameter: -UserIdentity is required when adding or removing members. Run with -h for help."
        exit 1
    }
    if (-not (Test-ValidEmailAddress -Email $UserIdentity) -and $UserIdentity -notmatch '^[a-zA-Z0-9._-]+$') {
        Write-Error "Invalid format for -UserIdentity: '$UserIdentity'. Provide a valid email address, UPN, or alias."
        exit 1
    }
}

if ($AddMember -and $RemoveMember) {
    Write-Error "Conflicting parameters: You cannot specify both -AddMember and -RemoveMember at the same time."
    exit 1
}

if ($Role -eq 'All' -and $AddMember) {
    Write-Error "Invalid role: 'All' is only valid when removing members with -RemoveMember."
    exit 1
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
        Write-Host "[*] Installing ExchangeOnlineManagement module for CurrentUser..." -ForegroundColor Cyan
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

# Group Resolver Helper
function Get-M365GroupEntity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Host "`n[*] Resolving group identity: '$Identity'..." -ForegroundColor Cyan

    $recipient = $null
    try {
        $recipient = Get-Recipient -Identity $Identity -ErrorAction SilentlyContinue
    } catch {
        # Proceed with fallback searches
    }

    $groupType = $null
    $rawGroup = $null
    $isTeams = $false

    if ($recipient) {
        switch ($recipient.RecipientTypeDetails) {
            'GroupMailbox' {
                $groupType = 'Microsoft365'
                try {
                    $rawGroup = Get-UnifiedGroup -Identity $recipient.Identity -ErrorAction Stop
                } catch {
                    $rawGroup = Get-UnifiedGroup -Identity $recipient.PrimarySmtpAddress -ErrorAction Stop
                }
                $isTeams = ($rawGroup.ResourceProvisioningOptions -contains 'Team')
            }
            { $_ -in @('MailUniversalDistributionGroup', 'RoomList') } {
                $groupType = 'DistributionList'
                $rawGroup = Get-DistributionGroup -Identity $recipient.Identity -ErrorAction Stop
                $isTeams = $false
            }
            'MailUniversalSecurityGroup' {
                $groupType = 'MailEnabledSecurity'
                $rawGroup = Get-DistributionGroup -Identity $recipient.Identity -ErrorAction Stop
                $isTeams = $false
            }
            'DynamicDistributionGroup' {
                $groupType = 'DynamicDistribution'
                $rawGroup = Get-DynamicDistributionGroup -Identity $recipient.Identity -ErrorAction Stop
                $isTeams = $false
            }
            default {
                Write-Error "Recipient '$Identity' found, but it is of type '$($recipient.RecipientTypeDetails)' and not a recognized group."
                return $null
            }
        }
    } else {
        # Fallback 1: Unified Group
        try {
            $rawGroup = Get-UnifiedGroup -Identity $Identity -ErrorAction SilentlyContinue
            if ($rawGroup) {
                $groupType = 'Microsoft365'
                $isTeams = ($rawGroup.ResourceProvisioningOptions -contains 'Team')
            }
        } catch {}

        # Fallback 2: Distribution / Security Group
        if (-not $rawGroup) {
            try {
                $rawGroup = Get-DistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
                if ($rawGroup) {
                    if ($rawGroup.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
                        $groupType = 'MailEnabledSecurity'
                    } else {
                        $groupType = 'DistributionList'
                    }
                    $isTeams = $false
                }
            } catch {}
        }

        # Fallback 3: Dynamic Distribution Group
        if (-not $rawGroup) {
            try {
                $rawGroup = Get-DynamicDistributionGroup -Identity $Identity -ErrorAction SilentlyContinue
                if ($rawGroup) {
                    $groupType = 'DynamicDistribution'
                    $isTeams = $false
                }
            } catch {}
        }
    }

    if (-not $rawGroup) {
        Write-Error "Group '$Identity' was not found in Exchange Online. Verify the name, email, or alias."
        return $null
    }

    return [PSCustomObject]@{
        RawGroup             = $rawGroup
        GroupType            = $groupType
        IsTeamsEnabled       = $isTeams
        DisplayName          = $rawGroup.DisplayName
        PrimarySmtpAddress   = $rawGroup.PrimarySmtpAddress
        Alias                = $rawGroup.Alias
        Identity             = $rawGroup.Identity
        RecipientTypeDetails = if ($recipient) { $recipient.RecipientTypeDetails } else { $groupType }
    }
}

# Group Members and Owners Retriever
function Get-GroupMembersAndOwners {
    param([PSCustomObject]$GroupEntity)

    $raw = $GroupEntity.RawGroup
    $owners = @()
    $members = @()

    switch ($GroupEntity.GroupType) {
        'Microsoft365' {
            try {
                $owners = @(Get-UnifiedGroupLinks -Identity $raw.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                Write-Warning "Could not retrieve owners for Unified Group '$($raw.DisplayName)': $_"
            }
            try {
                $members = @(Get-UnifiedGroupLinks -Identity $raw.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                Write-Warning "Could not retrieve members for Unified Group '$($raw.DisplayName)': $_"
            }
        }
        { $_ -in @('DistributionList', 'MailEnabledSecurity') } {
            # Owners / Managers in Distribution Groups are stored in ManagedBy
            if ($raw.ManagedBy) {
                foreach ($mgr in $raw.ManagedBy) {
                    try {
                        $rec = Get-Recipient -Identity $mgr -ErrorAction SilentlyContinue
                        if ($rec) {
                            $owners += $rec
                        } else {
                            $owners += [PSCustomObject]@{
                                DisplayName        = $mgr
                                PrimarySmtpAddress = $mgr
                                UserPrincipalName  = $mgr
                                RecipientType      = 'Manager'
                                Identity           = $mgr
                            }
                        }
                    } catch {
                        $owners += [PSCustomObject]@{
                            DisplayName        = $mgr
                            PrimarySmtpAddress = $mgr
                            UserPrincipalName  = $mgr
                            RecipientType      = 'Manager'
                            Identity           = $mgr
                        }
                    }
                }
            }
            try {
                $members = @(Get-DistributionGroupMember -Identity $raw.Identity -ResultSize Unlimited -ErrorAction Stop)
            } catch {
                Write-Warning "Could not retrieve members for Distribution Group '$($raw.DisplayName)': $_"
            }
        }
        'DynamicDistribution' {
            if ($raw.ManagedBy) {
                foreach ($mgr in $raw.ManagedBy) {
                    $owners += [PSCustomObject]@{
                        DisplayName        = $mgr
                        PrimarySmtpAddress = $mgr
                        UserPrincipalName  = $mgr
                        RecipientType      = 'Manager'
                        Identity           = $mgr
                    }
                }
            }
        }
    }

    return @{
        Owners  = $owners
        Members = $members
    }
}

# Detailed Group Display Formatter
function Show-GroupDetails {
    param(
        [PSCustomObject]$GroupEntity,
        [switch]$ShowFullMembership
    )

    $raw = $GroupEntity.RawGroup
    $membership = Get-GroupMembersAndOwners -GroupEntity $GroupEntity
    $owners = $membership.Owners
    $members = $membership.Members

    $teamsStatus = if ($GroupEntity.IsTeamsEnabled) {
        "[YES] Microsoft Teams Connected"
    } elseif ($GroupEntity.GroupType -eq 'Microsoft365') {
        "[NO] Standard M365 Group (No Team connected)"
    } else {
        "[NO] Not applicable ($($GroupEntity.GroupType))"
    }

    $groupTypeFriendly = switch ($GroupEntity.GroupType) {
        'Microsoft365'         { if ($GroupEntity.IsTeamsEnabled) { "Microsoft 365 Group (Teams-Enabled)" } else { "Microsoft 365 Group (Unified)" } }
        'DistributionList'     { "Distribution List / Group" }
        'MailEnabledSecurity'  { "Mail-Enabled Security Group" }
        'DynamicDistribution' { "Dynamic Distribution Group" }
        default                { $GroupEntity.GroupType }
    }

    $accessType = if ($GroupEntity.GroupType -eq 'Microsoft365') {
        $raw.AccessType
    } else {
        "N/A"
    }

    $sharePointUrl = if ($GroupEntity.GroupType -eq 'Microsoft365' -and $raw.SharePointSiteUrl) {
        $raw.SharePointSiteUrl
    } else {
        "N/A"
    }

    $externalSenders = if ($raw.RequireSenderAuthenticationEnabled -eq $false) {
        "Allowed (External & Internal)"
    } else {
        "Blocked (Internal Organization Only)"
    }

    $hiddenGAL = if ($raw.HiddenFromAddressListsEnabled) { "Yes (Hidden)" } else { "No (Visible in GAL)" }

    Write-Host "`n================================================================================" -ForegroundColor Cyan
    Write-Host "GROUP DETAILS: $($GroupEntity.DisplayName)" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "  Display Name         : $($raw.DisplayName)" -ForegroundColor White
    Write-Host "  Primary Email (SMTP) : $($raw.PrimarySmtpAddress)" -ForegroundColor White
    Write-Host "  Alias                : $($raw.Alias)" -ForegroundColor White
    Write-Host "  Group Type           : $groupTypeFriendly" -ForegroundColor White
    Write-Host "  Microsoft Teams      : $teamsStatus" -ForegroundColor $(if ($GroupEntity.IsTeamsEnabled) { "Green" } else { "Gray" })
    Write-Host "  Privacy / AccessType : $accessType" -ForegroundColor White
    Write-Host "  Hidden from GAL      : $hiddenGAL" -ForegroundColor White
    Write-Host "  External Senders     : $externalSenders" -ForegroundColor White

    if ($GroupEntity.GroupType -eq 'Microsoft365') {
        Write-Host "  SharePoint Site URL  : $sharePointUrl" -ForegroundColor Gray
        if ($raw.Notes) {
            Write-Host "  Description/Notes    : $($raw.Notes)" -ForegroundColor Gray
        }
    } elseif ($GroupEntity.GroupType -eq 'DynamicDistribution') {
        Write-Host "  Recipient Filter     : $($raw.RecipientFilter)" -ForegroundColor Yellow
        Write-Host "  Recipient Container  : $($raw.RecipientContainer)" -ForegroundColor Gray
    }

    if ($raw.WhenCreated) {
        Write-Host "  Created Date         : $($raw.WhenCreated)" -ForegroundColor Gray
    }

    Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "GROUP OWNERS / MANAGERS ($($owners.Count)):" -ForegroundColor Yellow
    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Cyan
    if ($owners.Count -gt 0) {
        $nr = 1
        $ownerRows = foreach ($o in $owners) {
            $email = if ($o.PrimarySmtpAddress) { $o.PrimarySmtpAddress } elseif ($o.UserPrincipalName) { $o.UserPrincipalName } else { $o.Identity }
            [PSCustomObject]@{
                Nr            = $nr++
                DisplayName   = $o.DisplayName
                EmailOrUPN    = $email
                RecipientType = $o.RecipientType
            }
        }
        $ownerRows | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Yellow
    } else {
        Write-Host "  No explicit owners found." -ForegroundColor Gray
    }

    Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor Cyan
    if ($GroupEntity.GroupType -eq 'DynamicDistribution') {
        Write-Host "GROUP MEMBERS (Dynamic Query-Based):" -ForegroundColor Yellow
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "  Members are evaluated dynamically at email delivery time." -ForegroundColor Gray
        Write-Host "  RecipientFilter: $($raw.RecipientFilter)" -ForegroundColor Gray
    } else {
        Write-Host "GROUP MEMBERS ($($members.Count)):" -ForegroundColor Yellow
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Cyan
        if ($members.Count -gt 0) {
            $nr = 1
            $memberRows = foreach ($m in $members) {
                $email = if ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress } elseif ($m.UserPrincipalName) { $m.UserPrincipalName } else { $m.Identity }
                [PSCustomObject]@{
                    Nr            = $nr++
                    DisplayName   = $m.DisplayName
                    EmailOrUPN    = $email
                    RecipientType = $m.RecipientType
                }
            }
            $memberRows | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White
        } else {
            Write-Host "  No members currently in this group." -ForegroundColor Gray
        }
    }
    Write-Host "================================================================================`n" -ForegroundColor Cyan
}

# Tenant Group Listing Function
function List-AllTenantGroups {
    param(
        [string]$GroupTypeFilter,
        [string]$SearchPattern,
        [switch]$DetailedView,
        [string]$CsvPath,
        [switch]$PassThruObjects
    )

    Write-Host "`n[*] Retrieving groups from Exchange Online (Filter: '$GroupTypeFilter')..." -ForegroundColor Cyan

    $allGroups = [System.Collections.Generic.List[PSCustomObject]]::new()

    # 1. Fetch Microsoft 365 Groups
    if ($GroupTypeFilter -in @('All', 'Microsoft365', 'Teams')) {
        Write-Host "    Querying Microsoft 365 Groups..." -ForegroundColor Gray
        try {
            $unified = @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop)
            foreach ($ug in $unified) {
                $isTeams = ($ug.ResourceProvisioningOptions -contains 'Team')
                if ($GroupTypeFilter -eq 'Teams' -and -not $isTeams) {
                    continue
                }

                $allGroups.Add([PSCustomObject]@{
                    DisplayName          = $ug.DisplayName
                    PrimarySmtpAddress   = $ug.PrimarySmtpAddress
                    Alias                = $ug.Alias
                    GroupType            = if ($isTeams) { 'M365 (Teams)' } else { 'M365 Group' }
                    RawType              = 'Microsoft365'
                    TeamsEnabled         = if ($isTeams) { 'Yes' } else { 'No' }
                    AccessType           = $ug.AccessType
                    HiddenFromGAL        = if ($ug.HiddenFromAddressListsEnabled) { 'True' } else { 'False' }
                    ExternalSenders      = if ($ug.RequireSenderAuthenticationEnabled -eq $false) { 'Allowed' } else { 'Blocked' }
                    SharePointSiteUrl    = $ug.SharePointSiteUrl
                    WhenCreated          = $ug.WhenCreated
                    Identity             = $ug.Identity
                    RawObject            = $ug
                })
            }
        } catch {
            Write-Warning "Failed to query Microsoft 365 Groups: $_"
        }
    }

    # 2. Fetch Distribution Groups and Security Groups
    if ($GroupTypeFilter -in @('All', 'DistributionList', 'MailEnabledSecurity')) {
        Write-Host "    Querying Distribution & Mail-Enabled Security Groups..." -ForegroundColor Gray
        try {
            $dlList = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
            foreach ($dl in $dlList) {
                $isSecurity = ($dl.RecipientTypeDetails -eq 'MailUniversalSecurityGroup' -or $dl.GroupType -like '*SecurityEnabled*')
                if ($GroupTypeFilter -eq 'DistributionList' -and $isSecurity) { continue }
                if ($GroupTypeFilter -eq 'MailEnabledSecurity' -and -not $isSecurity) { continue }

                $allGroups.Add([PSCustomObject]@{
                    DisplayName          = $dl.DisplayName
                    PrimarySmtpAddress   = $dl.PrimarySmtpAddress
                    Alias                = $dl.Alias
                    GroupType            = if ($isSecurity) { 'Mail-Security' } else { 'DistributionList' }
                    RawType              = if ($isSecurity) { 'MailEnabledSecurity' } else { 'DistributionList' }
                    TeamsEnabled         = 'No'
                    AccessType           = 'N/A'
                    HiddenFromGAL        = if ($dl.HiddenFromAddressListsEnabled) { 'True' } else { 'False' }
                    ExternalSenders      = if ($dl.RequireSenderAuthenticationEnabled -eq $false) { 'Allowed' } else { 'Blocked' }
                    SharePointSiteUrl    = 'N/A'
                    WhenCreated          = $dl.WhenCreated
                    Identity             = $dl.Identity
                    RawObject            = $dl
                })
            }
        } catch {
            Write-Warning "Failed to query Distribution Groups: $_"
        }
    }

    # 3. Fetch Dynamic Distribution Groups
    if ($GroupTypeFilter -in @('All', 'DynamicDistribution')) {
        Write-Host "    Querying Dynamic Distribution Groups..." -ForegroundColor Gray
        try {
            $dynList = @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop)
            foreach ($ddg in $dynList) {
                $allGroups.Add([PSCustomObject]@{
                    DisplayName          = $ddg.DisplayName
                    PrimarySmtpAddress   = $ddg.PrimarySmtpAddress
                    Alias                = $ddg.Alias
                    GroupType            = 'DynamicDL'
                    RawType              = 'DynamicDistribution'
                    TeamsEnabled         = 'No'
                    AccessType           = 'N/A'
                    HiddenFromGAL        = if ($ddg.HiddenFromAddressListsEnabled) { 'True' } else { 'False' }
                    ExternalSenders      = if ($ddg.RequireSenderAuthenticationEnabled -eq $false) { 'Allowed' } else { 'Blocked' }
                    SharePointSiteUrl    = 'N/A'
                    WhenCreated          = $ddg.WhenCreated
                    Identity             = $ddg.Identity
                    RawObject            = $ddg
                })
            }
        } catch {
            Write-Warning "Failed to query Dynamic Distribution Groups: $_"
        }
    }

    # Wildcard filter
    $filteredGroups = $allGroups
    if (-not [string]::IsNullOrWhiteSpace($SearchPattern)) {
        $pattern = "*$SearchPattern*"
        $filteredGroups = @($allGroups | Where-Object {
            $_.DisplayName -like $pattern -or
            $_.PrimarySmtpAddress -like $pattern -or
            $_.Alias -like $pattern
        })
    }

    if ($filteredGroups.Count -eq 0) {
        Write-Host "`n[!] No groups matching the specified criteria were found." -ForegroundColor Yellow
        return
    }

    Write-Host "`n[+] Found $($filteredGroups.Count) group(s):" -ForegroundColor Green

    # Output formatted summary table
    $index = 1
    $displayRows = foreach ($g in $filteredGroups) {
        [PSCustomObject]@{
            Nr            = $index++
            DisplayName   = $g.DisplayName
            GroupType     = $g.GroupType
            Teams         = $g.TeamsEnabled
            PrimaryEmail  = $g.PrimarySmtpAddress
            AccessType    = $g.AccessType
            HiddenFromGAL = $g.HiddenFromGAL
        }
    }
    $displayRows | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White

    # Detailed breakdown if requested
    if ($DetailedView) {
        Write-Host "`n[*] Generating detailed breakdown for listed groups..." -ForegroundColor Cyan
        foreach ($item in $filteredGroups) {
            $groupEnt = [PSCustomObject]@{
                RawGroup             = $item.RawObject
                GroupType            = $item.RawType
                IsTeamsEnabled       = ($item.TeamsEnabled -eq 'Yes')
                DisplayName          = $item.DisplayName
                PrimarySmtpAddress   = $item.PrimarySmtpAddress
                Alias                = $item.Alias
                Identity             = $item.Identity
                RecipientTypeDetails = $item.GroupType
            }
            Show-GroupDetails -GroupEntity $groupEnt -ShowFullMembership
        }
    }

    # CSV Export
    if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
        try {
            $exportData = $filteredGroups | Select-Object DisplayName, GroupType, TeamsEnabled, PrimarySmtpAddress, Alias, AccessType, HiddenFromGAL, ExternalSenders, SharePointSiteUrl, WhenCreated
            $exportData | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            Write-Host "[+] Successfully exported $($filteredGroups.Count) group(s) to CSV: '$CsvPath'" -ForegroundColor Green
        } catch {
            Write-Error "Failed to export groups to CSV '$CsvPath': $_"
        }
    }

    if ($PassThruObjects) {
        return $filteredGroups
    }
}

# Assign User to Group
function Assign-UserToGroup {
    param(
        [PSCustomObject]$GroupEntity,
        [string]$TargetUserIdentity,
        [string]$AssignedRole,
        [switch]$BypassConfirmation
    )

    # 1. Prevent assignment to dynamic groups
    if ($GroupEntity.GroupType -eq 'DynamicDistribution') {
        Write-Error "Cannot manually assign members to a Dynamic Distribution Group. Membership is rule-based and managed via RecipientFilter."
        exit 1
    }

    # 2. Validate target user in tenant
    Write-Host "`n[*] Validating user identity: '$TargetUserIdentity'..." -ForegroundColor Cyan
    $recipientUser = $null
    try {
        $recipientUser = Get-Recipient -Identity $TargetUserIdentity -ErrorAction Stop
        Write-Host "    Found recipient: $($recipientUser.DisplayName) ($($recipientUser.RecipientTypeDetails))" -ForegroundColor Gray
    } catch {
        Write-Error "User '$TargetUserIdentity' was not found in Exchange Online. Ensure the identity (email, UPN, or alias) is correct."
        exit 1
    }

    # 3. Check existing membership
    $membership = Get-GroupMembersAndOwners -GroupEntity $GroupEntity
    $isAlreadyMember = $false
    $isAlreadyOwner = $false

    foreach ($m in $membership.Members) {
        $mEmail = if ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress } elseif ($m.UserPrincipalName) { $m.UserPrincipalName } else { "" }
        if ($m.Identity -eq $recipientUser.Identity -or
            ($mEmail -and $mEmail -eq $recipientUser.PrimarySmtpAddress) -or
            $m.DisplayName -eq $recipientUser.DisplayName) {
            $isAlreadyMember = $true
            break
        }
    }

    foreach ($o in $membership.Owners) {
        $oEmail = if ($o.PrimarySmtpAddress) { $o.PrimarySmtpAddress } elseif ($o.UserPrincipalName) { $o.UserPrincipalName } else { "" }
        if ($o.Identity -eq $recipientUser.Identity -or
            ($oEmail -and $oEmail -eq $recipientUser.PrimarySmtpAddress) -or
            $o.DisplayName -eq $recipientUser.DisplayName) {
            $isAlreadyOwner = $true
            break
        }
    }

    if ($AssignedRole -eq 'Member' -and $isAlreadyMember) {
        Write-Host "`n[!] User '$($recipientUser.DisplayName)' ($TargetUserIdentity) is ALREADY a member of group '$($GroupEntity.DisplayName)'." -ForegroundColor Yellow
        Show-GroupDetails -GroupEntity $GroupEntity
        return
    }

    if ($AssignedRole -eq 'Owner' -and $isAlreadyOwner) {
        Write-Host "`n[!] User '$($recipientUser.DisplayName)' ($TargetUserIdentity) is ALREADY an owner of group '$($GroupEntity.DisplayName)'." -ForegroundColor Yellow
        Show-GroupDetails -GroupEntity $GroupEntity
        return
    }

    # 4. Display action summary
    $teamsBadge = if ($GroupEntity.IsTeamsEnabled) { " [Microsoft Teams]" } else { "" }
    Write-Host "`nAction Summary:" -ForegroundColor White
    Write-Host "  Action       : ASSIGN (Add to Group)" -ForegroundColor White
    Write-Host "  Target Group : $($GroupEntity.DisplayName) ($($GroupEntity.PrimarySmtpAddress))$teamsBadge" -ForegroundColor White
    Write-Host "  Group Type   : $($GroupEntity.GroupType)" -ForegroundColor White
    Write-Host "  Target User  : $($recipientUser.DisplayName) ($TargetUserIdentity)" -ForegroundColor White
    Write-Host "  Role         : $AssignedRole" -ForegroundColor White
    Write-Host ""

    # 5. Interactive confirmation (supports EN: Y/N and PL: T/N)
    if (-not $BypassConfirmation) {
        $confirmation = Read-Host "Are you sure you want to ADD '$($recipientUser.DisplayName)' to '$($GroupEntity.DisplayName)' as $AssignedRole? (Y/N)"
        if ($confirmation -notmatch '^(Y|y|T|t)$') {
            Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without confirmation prompt." -ForegroundColor Cyan
    }

    # 6. Perform addition
    Write-Host "`n[*] Adding '$($recipientUser.DisplayName)' to '$($GroupEntity.DisplayName)' as $AssignedRole..." -ForegroundColor Cyan
    try {
        if ($GroupEntity.GroupType -eq 'Microsoft365') {
            if ($AssignedRole -eq 'Owner') {
                Add-UnifiedGroupLinks -Identity $GroupEntity.Identity -LinkType Owners -Links $recipientUser.Identity -ErrorAction Stop
                # Ensure owner is also in Members list for full Teams/SharePoint access
                Add-UnifiedGroupLinks -Identity $GroupEntity.Identity -LinkType Members -Links $recipientUser.Identity -ErrorAction SilentlyContinue
            } else {
                Add-UnifiedGroupLinks -Identity $GroupEntity.Identity -LinkType Members -Links $recipientUser.Identity -ErrorAction Stop
            }
        } else {
            # DistributionList or MailEnabledSecurity
            if ($AssignedRole -eq 'Owner') {
                Set-DistributionGroup -Identity $GroupEntity.Identity -ManagedBy @{Add = $recipientUser.Identity} -BypassSecurityGroupManagerCheck -ErrorAction Stop
            } else {
                Add-DistributionGroupMember -Identity $GroupEntity.Identity -Member $recipientUser.Identity -BypassSecurityGroupManagerCheck -ErrorAction Stop
            }
        }

        Write-Host "`n[SUCCESS] Successfully added '$($recipientUser.DisplayName)' ($TargetUserIdentity) to group '$($GroupEntity.DisplayName)' as $AssignedRole!" -ForegroundColor Green
    } catch {
        Write-Error "Failed to add user to group: $_"
        exit 1
    }

    # 7. Display updated group details
    Show-GroupDetails -GroupEntity $GroupEntity
}

# Remove User from Group
function Remove-UserFromGroup {
    param(
        [PSCustomObject]$GroupEntity,
        [string]$TargetUserIdentity,
        [string]$TargetRole,
        [switch]$BypassConfirmation
    )

    # 1. Prevent removal from dynamic groups
    if ($GroupEntity.GroupType -eq 'DynamicDistribution') {
        Write-Error "Cannot manually remove members from a Dynamic Distribution Group. Membership is rule-based and managed via RecipientFilter."
        exit 1
    }

    # 2. Check current membership
    $membership = Get-GroupMembersAndOwners -GroupEntity $GroupEntity
    $matchedMember = $null
    $matchedOwner = $null

    foreach ($m in $membership.Members) {
        $mEmail = if ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress } elseif ($m.UserPrincipalName) { $m.UserPrincipalName } else { "" }
        if ($m.Identity -like "*$TargetUserIdentity*" -or
            ($mEmail -and $mEmail -eq $TargetUserIdentity) -or
            $m.DisplayName -eq $TargetUserIdentity) {
            $matchedMember = $m
            break
        }
    }

    foreach ($o in $membership.Owners) {
        $oEmail = if ($o.PrimarySmtpAddress) { $o.PrimarySmtpAddress } elseif ($o.UserPrincipalName) { $o.UserPrincipalName } else { "" }
        if ($o.Identity -like "*$TargetUserIdentity*" -or
            ($oEmail -and $oEmail -eq $TargetUserIdentity) -or
            $o.DisplayName -eq $TargetUserIdentity) {
            $matchedOwner = $o
            break
        }
    }

    if (-not $matchedMember -and -not $matchedOwner) {
        Write-Host "`n[!] User '$TargetUserIdentity' is NOT a member or owner of group '$($GroupEntity.DisplayName)'. Nothing to remove." -ForegroundColor Yellow
        Show-GroupDetails -GroupEntity $GroupEntity
        return
    }

    # 3. Sole Owner safety check for M365 Groups
    if ($GroupEntity.GroupType -eq 'Microsoft365' -and ($TargetRole -in @('Owner', 'All') -or (-not $matchedMember -and $matchedOwner))) {
        if ($membership.Owners.Count -le 1 -and $matchedOwner) {
            Write-Error "Cannot remove user '$TargetUserIdentity' because they are the SOLE OWNER of group '$($GroupEntity.DisplayName)'. Microsoft 365 Groups require at least one active owner. Please assign another owner first."
            exit 1
        }
    }

    # 4. Display action summary
    $teamsBadge = if ($GroupEntity.IsTeamsEnabled) { " [Microsoft Teams]" } else { "" }
    Write-Host "`nAction Summary:" -ForegroundColor White
    Write-Host "  Action       : REVOKE (Remove from Group)" -ForegroundColor White
    Write-Host "  Target Group : $($GroupEntity.DisplayName) ($($GroupEntity.PrimarySmtpAddress))$teamsBadge" -ForegroundColor White
    Write-Host "  Group Type   : $($GroupEntity.GroupType)" -ForegroundColor White
    Write-Host "  Target User  : $TargetUserIdentity" -ForegroundColor White
    Write-Host "  Role         : $TargetRole" -ForegroundColor White
    Write-Host ""

    # 5. Interactive confirmation (supports EN: Y/N and PL: T/N)
    if (-not $BypassConfirmation) {
        $confirmation = Read-Host "Are you sure you want to REMOVE '$TargetUserIdentity' from group '$($GroupEntity.DisplayName)'? (Y/N)"
        if ($confirmation -notmatch '^(Y|y|T|t)$') {
            Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without confirmation prompt." -ForegroundColor Cyan
    }

    # 6. Perform removal
    Write-Host "`n[*] Removing '$TargetUserIdentity' from '$($GroupEntity.DisplayName)' ($TargetRole)..." -ForegroundColor Cyan
    try {
        if ($GroupEntity.GroupType -eq 'Microsoft365') {
            if ($TargetRole -in @('Member', 'All') -and $matchedMember) {
                Remove-UnifiedGroupLinks -Identity $GroupEntity.Identity -LinkType Members -Links $TargetUserIdentity -Confirm:$false -ErrorAction Stop
            }
            if ($TargetRole -in @('Owner', 'All') -and $matchedOwner) {
                Remove-UnifiedGroupLinks -Identity $GroupEntity.Identity -LinkType Owners -Links $TargetUserIdentity -Confirm:$false -ErrorAction Stop
            }
        } else {
            # DistributionList or MailEnabledSecurity
            if ($TargetRole -in @('Member', 'All') -and $matchedMember) {
                Remove-DistributionGroupMember -Identity $GroupEntity.Identity -Member $TargetUserIdentity -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
            }
            if ($TargetRole -in @('Owner', 'All') -and $matchedOwner) {
                Set-DistributionGroup -Identity $GroupEntity.Identity -ManagedBy @{Remove = $TargetUserIdentity} -BypassSecurityGroupManagerCheck -ErrorAction Stop
            }
        }

        Write-Host "`n[SUCCESS] Successfully removed '$TargetUserIdentity' from group '$($GroupEntity.DisplayName)'!" -ForegroundColor Green
    } catch {
        Write-Error "Failed to remove user from group: $_"
        exit 1
    }

    # 7. Display updated group details
    Show-GroupDetails -GroupEntity $GroupEntity
}

# ==============================================================================
# MAIN EXECUTION DISPATCHER
# ==============================================================================

# Ensure connection to Exchange Online
Connect-ExchangeEnvironment -AdminUPN $AdminUserPrincipalName

if ($AddMember) {
    $groupEntity = Get-M365GroupEntity -Identity $GroupIdentity
    if ($groupEntity) {
        Assign-UserToGroup -GroupEntity $groupEntity -TargetUserIdentity $UserIdentity -AssignedRole $Role -BypassConfirmation:$Force
    }
} elseif ($RemoveMember) {
    $groupEntity = Get-M365GroupEntity -Identity $GroupIdentity
    if ($groupEntity) {
        Remove-UserFromGroup -GroupEntity $groupEntity -TargetUserIdentity $UserIdentity -TargetRole $Role -BypassConfirmation:$Force
    }
} elseif (-not [string]::IsNullOrWhiteSpace($GroupIdentity)) {
    # Inspect a specific group
    $groupEntity = Get-M365GroupEntity -Identity $GroupIdentity
    if ($groupEntity) {
        Show-GroupDetails -GroupEntity $groupEntity -ShowFullMembership
    }
} else {
    # List groups across tenant
    List-AllTenantGroups -GroupTypeFilter $GroupType -SearchPattern $Search -DetailedView:$Detailed -CsvPath $ExportCsv -PassThruObjects:$PassThru
}
