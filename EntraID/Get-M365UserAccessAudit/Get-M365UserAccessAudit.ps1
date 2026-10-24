<#
.SYNOPSIS
    Performs a comprehensive, enterprise-grade audit of user access, group memberships, directory roles,
    Teams, SharePoint Online sites, and Exchange Online mailbox permissions in Microsoft 365 / Entra ID.

.DESCRIPTION
    Get-M365UserAccessAudit.ps1 audits and maps all permissions, direct and transitive group memberships,
    administrative directory roles (including PIM eligible assignments), Teams and channels, associated
    SharePoint Online site collections, and Exchange Online delegated mailbox permissions for one or more
    specified Microsoft 365 / Entra ID users.

    Key Features:
    - Entra ID Groups: Deep audit of direct, transitive (nested), and owned security groups, M365 groups,
      and distribution lists with static vs dynamic rule classification and role-assignable detection.
    - Entra ID Directory Roles: Discovers directly assigned directory roles, roles inherited via role-assignable
      security groups, and optional Privileged Identity Management (PIM) eligible schedule instances.
    - Microsoft Teams: Maps associated Teams, team roles (Owner/Member), and private/shared channel memberships.
    - SharePoint Online: Discovers associated SharePoint site collections, site URLs, and optional site permissions.
    - Exchange Online: Maps distribution lists, dynamic distribution groups, and optional mailbox delegation
      rights (FullAccess, SendAs, SendOnBehalf) assigned directly or via mail-enabled security groups.
    - High Performance & Throttling Protection: Leverages Graph API pagination (-All) and synchronized dictionary
      caching (Hashtables) to minimize API roundtrips and avoid HTTP 429 throttling.
    - Enterprise Reporting: Automatically exports comprehensive UTF-8 BOM CSV, JSON, and interactive,
      modern HTML executive dashboards with real-time client-side search, KPI summary cards, and workload filters.
    - Standardized Logging: Centralized dual logging to console and C:\Logs\Get-M365UserAccessAudit\ with level
      tagging ([INFO], [WARN], [ERROR], [SUCCESS], [DEBUG]).

.PARAMETER UserPrincipalName
    One or more User Principal Names (UPN) or ObjectIds of the users to audit.
    Accepts pipeline input by value and by property name.
    Aliases: -UPN, -User, -Identity, -UserId.

.PARAMETER IncludeExchangePermissions
    Switch parameter. Enables scanning of Exchange Online mailbox permissions (FullAccess, SendAs,
    SendOnBehalf) across mailboxes. Requires active Exchange Online PowerShell connection.

.PARAMETER IncludeSharePointSiteAudit
    Switch parameter. Performs deep permission inspection on associated SharePoint Online site collections
    to verify explicit and group-inherited site roles.

.PARAMETER IncludePIM
    Switch parameter. Audits Privileged Identity Management (PIM) eligible role assignment schedules
    in Entra ID (requires RoleEligibilitySchedule.Read.Directory permission and Entra ID P2/Governance license).

.PARAMETER ExportFormat
    Specifies the output format(s) for the audit report. Supported formats: 'CSV', 'HTML', 'JSON'.
    Default: @('CSV', 'HTML').

.PARAMETER OutputPath
    Target directory where generated reports (CSV, HTML, JSON) are stored.
    Default: 'C:\Logs\Get-M365UserAccessAudit\Reports'.

.PARAMETER LogPath
    Directory where timestamped execution log files (.log) are saved.
    Default: 'C:\Logs\Get-M365UserAccessAudit'.

.PARAMETER TenantId
    Optional Azure Active Directory / Entra ID Tenant ID to target during authentication.

.PARAMETER PassThru
    Switch parameter. Emits custom PowerShell objects ([PSCustomObject]) to the pipeline for downstream
    cmdlets, scripting, or Out-GridView.

.PARAMETER Help
    Displays built-in help, usage examples, parameter descriptions, and exits.
    Alias: -h.

.EXAMPLE
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "jan.kowalski@contoso.com"
    Audits Jan Kowalski across Entra ID groups, directory roles, Teams, and SharePoint sites,
    generating CSV and HTML executive reports in C:\Logs\Get-M365UserAccessAudit\Reports\.

.EXAMPLE
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "anna.nowak@contoso.com" -IncludeExchangePermissions -IncludePIM
    Performs a deep audit including Exchange Online mailbox delegations and Entra ID PIM eligible roles.

.EXAMPLE
    "user1@contoso.com", "user2@contoso.com" | .\Get-M365UserAccessAudit.ps1 -ExportFormat CSV,HTML,JSON
    Audits multiple users via pipeline input and generates CSV, HTML, and JSON reports.

.EXAMPLE
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "admin@contoso.com" -PassThru | Out-GridView
    Audits the administrator and displays results in an interactive Out-GridView window.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
    Standards: Adheres to CleanCode and Microsoft 365 PowerShell Standards
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0, ParameterSetName = 'Default')]
    [ValidateNotNullOrEmpty()]
    [Alias('UPN', 'User', 'Identity', 'UserId')]
    [string[]]$UserPrincipalName,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$IncludeExchangePermissions,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$IncludeSharePointSiteAudit,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$IncludePIM,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [ValidateSet('CSV', 'HTML', 'JSON')]
    [string[]]$ExportFormat = @('CSV', 'HTML'),

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [string]$OutputPath = 'C:\Logs\Get-M365UserAccessAudit\Reports',

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [string]$LogPath = 'C:\Logs\Get-M365UserAccessAudit',

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [string]$TenantId,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$PassThru,

    [Parameter(Mandatory = $false, ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # ==============================================================================
    # SCRIPT METADATA & GLOBAL AUDIT STATE
    # ==============================================================================
    $Script:Metadata = @{
        Title        = "Get-M365UserAccessAudit.ps1"
        Version      = "1.0.0"
        Author       = "Roman Pindela"
        Email        = "roman.pindela@gmail.com"
        GitHub       = "https://github.com/romanpindela"
        Specification= "SPECIFICATION-Get-M365UserAccessAudit.md"
        LastModified = "2026-10-24"
    }

    $Script:StartTime        = Get-Date
    $Script:LogFilePath      = $null
    $Script:ActiveTenantId   = $null
    $Script:ActiveTenantName = $null
    $Script:GeneratedHtmlPath = $null
    $Script:GeneratedCsvPath  = $null
    $Script:GeneratedJsonPath = $null

    # Synchronized caches to eliminate duplicate Graph API calls and prevent HTTP 429 throttling
    $Script:GroupCache               = [System.Collections.Hashtable]::Synchronized(@{})
    $Script:RoleDefinitionCache      = [System.Collections.Hashtable]::Synchronized(@{})
    $Script:GroupRoleAssignmentCache = [System.Collections.Hashtable]::Synchronized(@{})
    $Script:TeamCache                = [System.Collections.Hashtable]::Synchronized(@{})
    $Script:SiteCache                = [System.Collections.Hashtable]::Synchronized(@{})
    $Script:AllAuditRecords          = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ==============================================================================
    # BUILT-IN HELP & BANNER
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        $banner = @"
================================================================================
SCRIPT  : $($Script:Metadata.Title) v$($Script:Metadata.Version)
AUTHOR  : $($Script:Metadata.Author) <$($Script:Metadata.Email)>
GITHUB  : $($Script:Metadata.GitHub)
================================================================================

DESCRIPTION:
    Comprehensive, enterprise-ready user access audit for Microsoft 365 / Entra ID.
    Extracts direct & transitive group memberships, administrative directory roles,
    PIM eligible assignments, Teams & channels, SharePoint Online sites, and
    Exchange Online mailbox delegations into standardized CSV, HTML, and JSON reports.

PARAMETERS:
    -UserPrincipalName, -UPN        One or more UPNs or ObjectIds of users to audit.
    -IncludeExchangePermissions     Audit Exchange Online mailbox delegation rights.
    -IncludeSharePointSiteAudit     Deep audit of SharePoint site collection permissions.
    -IncludePIM                     Audit Entra ID Privileged Identity Management schedules.
    -ExportFormat                   Report formats: 'CSV', 'HTML', 'JSON' (Default: CSV, HTML).
    -OutputPath                     Directory for generated audit reports.
    -LogPath                        Directory for execution log files.
    -TenantId                       Optional Entra ID Tenant ID to force sign-in context.
    -PassThru                       Output custom objects to pipeline for Out-GridView.
    -Help, -h                       Display this help banner and exit.

USAGE EXAMPLES:
    # 1. Standard comprehensive audit for single user:
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "jan.kowalski@contoso.com"

    # 2. Deep audit including Exchange Online mailbox permissions & PIM:
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "anna.nowak@contoso.com" -IncludeExchangePermissions -IncludePIM

    # 3. Pipeline audit for multiple users with JSON & HTML export:
    "user1@contoso.com", "user2@contoso.com" | .\Get-M365UserAccessAudit.ps1 -ExportFormat CSV,HTML,JSON

    # 4. Interactive pipeline analysis with Out-GridView:
    .\Get-M365UserAccessAudit.ps1 -UserPrincipalName "admin@contoso.com" -PassThru | Out-GridView

SECURITY & UNBLOCKING:
    Unblock-File -Path .\$($Script:Metadata.Title)
================================================================================
"@
        Write-Host $banner -ForegroundColor Cyan
    }

    # ==============================================================================
    # LOGGING SYSTEM (CleanCode & Enterprise Standard)
    # ==============================================================================
    function Initialize-AuditLogging {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$TargetDirectory = 'C:\Logs\Get-M365UserAccessAudit'
        )

        $resolvedDir = $TargetDirectory
        try {
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Get-M365UserAccessAudit"
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
        }

        $timestamp   = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $logFileName = "Get-M365UserAccessAudit_${timestamp}.log"
        $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName

        $initHeader = @"
================================================================================
Audit Log Initialized : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script                : $($Script:Metadata.Title) v$($Script:Metadata.Version)
Author                : $($Script:Metadata.Author) ($($Script:Metadata.Email))
GitHub                : $($Script:Metadata.GitHub)
Host Machine          : $env:COMPUTERNAME
User Context          : $env:USERNAME
PowerShell Version    : $($PSVersionTable.PSVersion)
Log File Path         : $Script:LogFilePath
================================================================================
"@
        Add-Content -Path $Script:LogFilePath -Value $initHeader -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    function Write-Log {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
            [string]$Level = 'INFO',

            [Parameter(Mandatory = $false)]
            [string]$Component = 'GENERAL',

            [Parameter(Mandatory = $false)]
            [switch]$NoConsole
        )

        $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $logLine = "[$timeStr] [$($Level.PadRight(7))] [$($Component.PadRight(10))] $Message"

        if ($Script:LogFilePath) {
            try {
                Add-Content -Path $Script:LogFilePath -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose -Message "Unable to write log entry: $($_.Exception.Message)"
            }
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

            $symbol = switch ($Level) {
                'INFO'    { '[*]' }
                'SUCCESS' { '[+]' }
                'WARN'    { '[!]' }
                'ERROR'   { '[-]' }
                'DEBUG'   { '[D]' }
                Default   { '[*]' }
            }

            Write-Host "$symbol [$Component] $Message" -ForegroundColor $color
        }
    }

    # ==============================================================================
    # PREREQUISITES & MODULE VERIFICATION
    # ==============================================================================
    function Test-AuditPrerequisite {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [switch]$RequireExchange
        )

        Write-Log -Message "Verifying required PowerShell modules and runtime environment..." -Level 'INFO' -Component 'INIT'

        $coreGraphModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Groups',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Teams',
            'Microsoft.Graph.Sites'
        )

        $missingModules = [System.Collections.Generic.List[string]]::new()
        foreach ($mod in $coreGraphModules) {
            $found = Get-Module -Name $mod -ListAvailable -ErrorAction SilentlyContinue
            if (-not $found) {
                $missingModules.Add($mod)
            }
        }

        if ($missingModules.Count -gt 0) {
            $msg = "Missing Microsoft Graph module(s): $($missingModules -join ', '). Install via: Install-Module -Name $($missingModules[0]) -Scope CurrentUser"
            Write-Log -Message $msg -Level 'WARN' -Component 'INIT'
        }

        if ($RequireExchange) {
            $exoFound = Get-Module -Name 'ExchangeOnlineManagement' -ListAvailable -ErrorAction SilentlyContinue
            if (-not $exoFound) {
                $msg = "ExchangeOnlineManagement module is missing. Install via: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
                Write-Log -Message $msg -Level 'ERROR' -Component 'INIT'
                throw $msg
            }
        }

        Write-Log -Message "Module inspection completed successfully." -Level 'SUCCESS' -Component 'INIT'
    }

    # ==============================================================================
    # AUTHENTICATION & SERVICE CONNECTIONS
    # ==============================================================================
    function Connect-AuditService {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$TargetTenantId,

            [Parameter(Mandatory = $false)]
            [switch]$AuditExchange,

            [Parameter(Mandatory = $false)]
            [switch]$AuditPIM
        )

        Write-Log -Message "Checking Microsoft Graph authentication context..." -Level 'INFO' -Component 'AUTH'

        $graphScopes = @(
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All',
            'RoleManagement.Read.Directory',
            'Team.ReadBasic.All',
            'Channel.ReadBasic.All',
            'Sites.Read.All'
        )

        if ($AuditPIM) {
            $graphScopes += 'RoleEligibilitySchedule.Read.Directory'
        }

        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        $needGraphConnect = $true

        if ($mgContext) {
            Write-Log -Message "Existing Microsoft Graph session detected. Account: $($mgContext.Account), Tenant: $($mgContext.TenantId)" -Level 'INFO' -Component 'AUTH'
            $currentScopes = $mgContext.Scopes
            $missingScopes = @()
            foreach ($sc in $graphScopes) {
                if ($currentScopes -notcontains $sc) {
                    $missingScopes += $sc
                }
            }

            if ($missingScopes.Count -eq 0) {
                $needGraphConnect = $false
                $Script:ActiveTenantId = $mgContext.TenantId
                Write-Log -Message "Existing Graph session contains all necessary scopes." -Level 'SUCCESS' -Component 'AUTH'
            }
            else {
                Write-Log -Message "Existing Graph session lacks required scopes: $($missingScopes -join ', '). Reconnecting..." -Level 'WARN' -Component 'AUTH'
            }
        }

        if ($needGraphConnect) {
            Write-Log -Message "Initiating connection to Microsoft Graph..." -Level 'INFO' -Component 'AUTH'
            $connectParams = @{
                Scopes      = $graphScopes
                NoWelcome   = $true
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($TargetTenantId)) {
                $connectParams['TenantId'] = $TargetTenantId
            }

            try {
                Connect-MgGraph @connectParams
                $newContext = Get-MgContext
                $Script:ActiveTenantId = $newContext.TenantId
                Write-Log -Message "Connected to Microsoft Graph. Tenant ID: $($Script:ActiveTenantId)" -Level 'SUCCESS' -Component 'AUTH'
            }
            catch {
                $err = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
                Write-Log -Message $err -Level 'ERROR' -Component 'AUTH'
                throw $err
            }
        }

        # Retrieve Tenant Organization Display Name
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction SilentlyContinue
            if ($org -and $org.value -and $org.value[0].displayName) {
                $Script:ActiveTenantName = $org.value[0].displayName
                Write-Log -Message "Tenant Organization: '$($Script:ActiveTenantName)'" -Level 'INFO' -Component 'AUTH'
            }
        }
        catch {
            Write-Log -Message "Could not resolve tenant organization display name." -Level 'DEBUG' -Component 'AUTH'
        }

        # Exchange Online Connection
        if ($AuditExchange) {
            Write-Log -Message "Checking Exchange Online session..." -Level 'INFO' -Component 'EXO'
            $exoConnected = $false
            try {
                $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
                if ($connInfo -and $connInfo.Count -gt 0) {
                    $exoConnected = $true
                    Write-Log -Message "Active Exchange Online session detected." -Level 'SUCCESS' -Component 'EXO'
                }
            }
            catch {
                $exoConnected = $false
            }

            if (-not $exoConnected) {
                Write-Log -Message "Connecting to Exchange Online PowerShell..." -Level 'INFO' -Component 'EXO'
                $exoParams = @{
                    ShowBanner  = $false
                    ErrorAction = 'Stop'
                }
                if (-not [string]::IsNullOrWhiteSpace($TargetTenantId)) {
                    $exoParams['Organization'] = $TargetTenantId
                }

                try {
                    Connect-ExchangeOnline @exoParams
                    Write-Log -Message "Successfully connected to Exchange Online." -Level 'SUCCESS' -Component 'EXO'
                }
                catch {
                    $err = "Failed to connect to Exchange Online: $($_.Exception.Message)"
                    Write-Log -Message $err -Level 'ERROR' -Component 'EXO'
                    throw $err
                }
            }
        }
    }

    # ==============================================================================
    # CACHING REPOSITORIES & GRAPH HELPERS
    # ==============================================================================
    function Get-CachedGroup {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$GroupId
        )

        if ($Script:GroupCache.ContainsKey($GroupId)) {
            return $Script:GroupCache[$GroupId]
        }

        try {
            $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=id,displayName,mail,securityEnabled,mailEnabled,groupTypes,isAssignableToRole,onPremisesSyncEnabled,resourceProvisioningOptions"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response) {
                $Script:GroupCache[$GroupId] = $response
                return $response
            }
        }
        catch {
            Write-Log -Message "Failed to fetch metadata for group '$GroupId': $($_.Exception.Message)" -Level 'DEBUG' -Component 'ENTRA_ID'
        }

        return $null
    }

    function Get-CachedRoleDefinition {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$RoleDefinitionId
        )

        if ($Script:RoleDefinitionCache.ContainsKey($RoleDefinitionId)) {
            return $Script:RoleDefinitionCache[$RoleDefinitionId]
        }

        try {
            $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions/$RoleDefinitionId`?`$select=id,displayName,description,isBuiltIn"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response) {
                $Script:RoleDefinitionCache[$RoleDefinitionId] = $response
                return $response
            }
        }
        catch {
            Write-Log -Message "Failed to fetch role definition '$RoleDefinitionId': $($_.Exception.Message)" -Level 'DEBUG' -Component 'ROLES'
        }

        return $null
    }

    function Get-CachedSiteForGroup {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$GroupId
        )

        if ($Script:SiteCache.ContainsKey($GroupId)) {
            return $Script:SiteCache[$GroupId]
        }

        try {
            $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/sites/root`?`$select=id,displayName,webUrl,name"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response) {
                $Script:SiteCache[$GroupId] = $response
                return $response
            }
        }
        catch {
            Write-Log -Message "No root site provisioned for M365 Group '$GroupId'." -Level 'DEBUG' -Component 'SPO'
        }

        return $null
    }

    # ==============================================================================
    # AUDIT ENGINE: AUDIT SINGLE USER
    # ==============================================================================
    function Invoke-UserAudit {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetUser,

            [Parameter(Mandatory = $false)]
            [switch]$IncludeExchange,

            [Parameter(Mandatory = $false)]
            [switch]$IncludeSharePoint,

            [Parameter(Mandatory = $false)]
            [switch]$IncludePIMSchedule
        )

        $nowTimestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Write-Log -Message "================================================================================" -Level 'INFO' -Component 'AUDIT'
        Write-Log -Message "Starting access audit for user: '$TargetUser'" -Level 'INFO' -Component 'AUDIT'

        # --------------------------------------------------------------------------
        # Phase 1/6: Resolving User Identity
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 1/6: Resolving User Identity..." -PercentComplete 10
        Write-Log -Message "Resolving Entra ID user profile..." -Level 'INFO' -Component 'ENTRA_ID'

        $userProfile = $null
        try {
            $escapedUser = [System.Uri]::EscapeDataString($TargetUser)
            $userUri = "https://graph.microsoft.com/v1.0/users/$escapedUser`?`$select=id,displayName,userPrincipalName,accountEnabled,department,jobTitle,onPremisesSyncEnabled,mail"
            $userProfile = Invoke-MgGraphRequest -Method GET -Uri $userUri -ErrorAction Stop
        }
        catch {
            Write-Log -Message "Unable to locate user '$TargetUser' in Entra ID: $($_.Exception.Message)" -Level 'ERROR' -Component 'ENTRA_ID'
            return
        }

        if (-not $userProfile -or -not $userProfile.id) {
            Write-Log -Message "User profile returned empty for '$TargetUser'." -Level 'ERROR' -Component 'ENTRA_ID'
            return
        }

        $userId             = $userProfile.id
        $userUpn            = $userProfile.userPrincipalName
        $userDisplayName    = if ($userProfile.displayName) { $userProfile.displayName } else { $userUpn }
        $accountEnabled     = if ($null -ne $userProfile.accountEnabled) { [bool]$userProfile.accountEnabled } else { $true }
        $userDepartment     = if ($userProfile.department) { $userProfile.department } else { "N/A" }
        $userJobTitle       = if ($userProfile.jobTitle) { $userProfile.jobTitle } else { "N/A" }
        $userSyncSource     = if ($userProfile.onPremisesSyncEnabled -eq $true) { "On-Premises AD" } else { "Cloud" }

        Write-Log -Message "Identified: $userDisplayName ($userUpn) | Title: $userJobTitle | Dept: $userDepartment | Enabled: $accountEnabled | Sync: $userSyncSource" -Level 'SUCCESS' -Component 'ENTRA_ID'

        # --------------------------------------------------------------------------
        # Phase 2/6: Auditing Entra ID Groups (Direct, Transitive, Owned)
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 2/6: Auditing Entra ID Groups (Direct & Transitive)..." -PercentComplete 25
        Write-Log -Message "Querying direct, transitive, and owned group memberships..." -Level 'INFO' -Component 'ENTRA_ID'

        $directGroupIds     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $ownedGroupIds      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $allGroupIds        = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        # 1. Direct Memberships
        try {
            $directUri = "https://graph.microsoft.com/v1.0/users/$userId/memberOf`?`$top=999`&`$select=id,displayName"
            $directResp = Invoke-MgGraphRequest -Method GET -Uri $directUri -ErrorAction Stop
            if ($directResp -and $directResp.value) {
                foreach ($item in $directResp.value) {
                    $itemType = $item.'@odata.type'
                    if ($null -eq $itemType -or $itemType -eq '#microsoft.graph.group') {
                        $null = $directGroupIds.Add($item.id)
                        $null = $allGroupIds.Add($item.id)
                    }
                }
            }
        }
        catch {
            Write-Log -Message "Error querying direct group memberships: $($_.Exception.Message)" -Level 'WARN' -Component 'ENTRA_ID'
        }

        # 2. Transitive (Nested) Memberships
        try {
            $transUri = "https://graph.microsoft.com/v1.0/users/$userId/transitiveMemberOf`?`$top=999`&`$select=id,displayName"
            $transResp = Invoke-MgGraphRequest -Method GET -Uri $transUri -ErrorAction Stop
            if ($transResp -and $transResp.value) {
                foreach ($item in $transResp.value) {
                    $itemType = $item.'@odata.type'
                    if ($null -eq $itemType -or $itemType -eq '#microsoft.graph.group') {
                        $null = $allGroupIds.Add($item.id)
                    }
                }
            }
        }
        catch {
            Write-Log -Message "Error querying transitive group memberships: $($_.Exception.Message)" -Level 'WARN' -Component 'ENTRA_ID'
        }

        # 3. Owned Groups
        try {
            $ownedUri = "https://graph.microsoft.com/v1.0/users/$userId/ownedObjects`?`$top=999`&`$select=id,displayName"
            $ownedResp = Invoke-MgGraphRequest -Method GET -Uri $ownedUri -ErrorAction Stop
            if ($ownedResp -and $ownedResp.value) {
                foreach ($item in $ownedResp.value) {
                    $itemType = $item.'@odata.type'
                    if ($null -eq $itemType -or $itemType -eq '#microsoft.graph.group') {
                        $null = $ownedGroupIds.Add($item.id)
                        $null = $allGroupIds.Add($item.id)
                    }
                }
            }
        }
        catch {
            Write-Log -Message "Error querying owned groups: $($_.Exception.Message)" -Level 'WARN' -Component 'ENTRA_ID'
        }

        Write-Log -Message "Total unique groups discovered: $($allGroupIds.Count) (Direct: $($directGroupIds.Count), Owned: $($ownedGroupIds.Count))" -Level 'INFO' -Component 'ENTRA_ID'

        # Process and classify each discovered group
        $userM365Groups = [System.Collections.Generic.List[PSCustomObject]]::new()
        $userSecurityGroups = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($gId in $allGroupIds) {
            $groupObj = Get-CachedGroup -GroupId $gId
            if (-not $groupObj) {
                continue
            }

            $groupName       = if ($groupObj.displayName) { $groupObj.displayName } else { $gId }
            $secEnabled      = [bool]$groupObj.securityEnabled
            $mailEnabled     = [bool]$groupObj.mailEnabled
            $groupTypes      = if ($groupObj.groupTypes) { @($groupObj.groupTypes) } else { @() }
            $isRoleAssign    = if ($null -ne $groupObj.isAssignableToRole) { [bool]$groupObj.isAssignableToRole } else { $false }
            $isDynamic       = ($groupTypes -contains 'DynamicMembership')
            $syncSrc         = if ($groupObj.onPremisesSyncEnabled -eq $true) { "On-Premises AD" } else { "Cloud" }

            # Group Classification
            $resourceType = "Security Group"
            $workload     = "Entra ID"

            if ($groupTypes -contains 'Unified') {
                $resourceType = "M365 Group"
                $workload     = "Entra ID"
                $userM365Groups.Add($groupObj)
            }
            elseif ($secEnabled -and -not $mailEnabled) {
                $resourceType = "Security Group"
                $workload     = "Entra ID"
                $userSecurityGroups.Add($groupObj)
            }
            elseif ($secEnabled -and $mailEnabled) {
                $resourceType = "Mail-Enabled Security Group"
                $workload     = "Exchange Online"
                $userSecurityGroups.Add($groupObj)
            }
            elseif (-not $secEnabled -and $mailEnabled) {
                $resourceType = "Distribution Group"
                $workload     = "Exchange Online"
            }

            # Membership Type
            $membershipType = if ($isDynamic) { "Dynamic" } else { "Assigned (Static)" }

            # Assignment Method & Access Level
            $isOwner  = $ownedGroupIds.Contains($gId)
            $isDirect = $directGroupIds.Contains($gId)

            if ($isOwner) {
                $assignMethod = "Direct"
                $grantedVia   = "Direct"
                $accessLvl    = "Owner"

                $record = [PSCustomObject]@{
                    AuditTimestamp      = $nowTimestamp
                    UserPrincipalName   = $userUpn
                    UserDisplayName     = $userDisplayName
                    AccountEnabled      = $accountEnabled
                    Workload            = $workload
                    ResourceType        = $resourceType
                    ResourceName        = $groupName
                    ResourceId          = $gId
                    AccessLevel         = $accessLvl
                    AssignmentMethod    = $assignMethod
                    GrantedViaGroup     = $grantedVia
                    GroupMembershipType = $membershipType
                    IsRoleAssignable    = $isRoleAssign
                    SyncSource          = $syncSrc
                    ResourceUrl         = ""
                }
                $Script:AllAuditRecords.Add($record)
            }

            if ($isDirect -or (-not $isOwner)) {
                $assignMethod = if ($isDynamic) {
                    "Dynamic Rule"
                }
                elseif ($isDirect) {
                    "Direct"
                }
                else {
                    "Nested / Group Inherited"
                }

                $grantedVia = if ($isDirect) { "Direct" } else { "Nested / Transitive" }
                $accessLvl  = "Member"

                $record = [PSCustomObject]@{
                    AuditTimestamp      = $nowTimestamp
                    UserPrincipalName   = $userUpn
                    UserDisplayName     = $userDisplayName
                    AccountEnabled      = $accountEnabled
                    Workload            = $workload
                    ResourceType        = $resourceType
                    ResourceName        = $groupName
                    ResourceId          = $gId
                    AccessLevel         = $accessLvl
                    AssignmentMethod    = $assignMethod
                    GrantedViaGroup     = $grantedVia
                    GroupMembershipType = $membershipType
                    IsRoleAssignable    = $isRoleAssign
                    SyncSource          = $syncSrc
                    ResourceUrl         = ""
                }
                $Script:AllAuditRecords.Add($record)
            }
        }

        # --------------------------------------------------------------------------
        # Phase 3/6: Auditing Entra ID Administrative Roles & PIM
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 3/6: Auditing Directory Roles & PIM..." -PercentComplete 40
        Write-Log -Message "Auditing Entra ID Directory Roles (Direct & Group Inherited)..." -Level 'INFO' -Component 'ROLES'

        # 1. Direct Directory Role Assignments
        try {
            $roleAssignUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments`?`$filter=principalId eq '$userId'`&`$expand=roleDefinition"
            $roleAssignResp = Invoke-MgGraphRequest -Method GET -Uri $roleAssignUri -ErrorAction Stop

            if ($roleAssignResp -and $roleAssignResp.value) {
                foreach ($ra in $roleAssignResp.value) {
                    $roleDefId = $ra.roleDefinitionId
                    $roleDefObj = if ($ra.roleDefinition) { $ra.roleDefinition } else { Get-CachedRoleDefinition -RoleDefinitionId $roleDefId }
                    $roleName   = if ($roleDefObj -and $roleDefObj.displayName) { $roleDefObj.displayName } else { $roleDefId }

                    Write-Log -Message "Direct Directory Role Found: '$roleName'" -Level 'WARN' -Component 'ROLES'

                    $record = [PSCustomObject]@{
                        AuditTimestamp      = $nowTimestamp
                        UserPrincipalName   = $userUpn
                        UserDisplayName     = $userDisplayName
                        AccountEnabled      = $accountEnabled
                        Workload            = "Entra ID"
                        ResourceType        = "Directory Role"
                        ResourceName        = $roleName
                        ResourceId          = $roleDefId
                        AccessLevel         = "Directory Role (Active)"
                        AssignmentMethod    = "Direct"
                        GrantedViaGroup     = "Direct"
                        GroupMembershipType = "N/A"
                        IsRoleAssignable    = $false
                        SyncSource          = "Cloud"
                        ResourceUrl         = ""
                    }
                    $Script:AllAuditRecords.Add($record)
                }
            }
        }
        catch {
            Write-Log -Message "Failed to query direct directory role assignments: $($_.Exception.Message)" -Level 'WARN' -Component 'ROLES'
        }

        # 2. Inherited Directory Role Assignments via Groups
        foreach ($grp in $userSecurityGroups) {
            if ($grp.isAssignableToRole -eq $true) {
                $gId = $grp.id
                $gName = $grp.displayName

                try {
                    $grpRoleUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments`?`$filter=principalId eq '$gId'`&`$expand=roleDefinition"
                    $grpRoleResp = Invoke-MgGraphRequest -Method GET -Uri $grpRoleUri -ErrorAction Stop

                    if ($grpRoleResp -and $grpRoleResp.value) {
                        foreach ($ra in $grpRoleResp.value) {
                            $roleDefId = $ra.roleDefinitionId
                            $roleDefObj = if ($ra.roleDefinition) { $ra.roleDefinition } else { Get-CachedRoleDefinition -RoleDefinitionId $roleDefId }
                            $roleName   = if ($roleDefObj -and $roleDefObj.displayName) { $roleDefObj.displayName } else { $roleDefId }

                            Write-Log -Message "Inherited Directory Role Found: '$roleName' via Group '$gName'" -Level 'WARN' -Component 'ROLES'

                            $record = [PSCustomObject]@{
                                AuditTimestamp      = $nowTimestamp
                                UserPrincipalName   = $userUpn
                                UserDisplayName     = $userDisplayName
                                AccountEnabled      = $accountEnabled
                                Workload            = "Entra ID"
                                ResourceType        = "Directory Role"
                                ResourceName        = $roleName
                                ResourceId          = $roleDefId
                                AccessLevel         = "Directory Role (Inherited)"
                                AssignmentMethod    = "Nested / Group Inherited"
                                GrantedViaGroup     = $gName
                                GroupMembershipType = "N/A"
                                IsRoleAssignable    = $true
                                SyncSource          = "Cloud"
                                ResourceUrl         = ""
                            }
                            $Script:AllAuditRecords.Add($record)
                        }
                    }
                }
                catch {
                    Write-Log -Message "Error inspecting role assignments for group '$gName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'ROLES'
                }
            }
        }

        # 3. Privileged Identity Management (PIM) Eligible Roles
        if ($IncludePIMSchedule) {
            Write-Log -Message "Auditing Entra ID PIM Eligible Role Assignments..." -Level 'INFO' -Component 'PIM'
            try {
                $pimUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances`?`$filter=principalId eq '$userId'`&`$expand=roleDefinition"
                $pimResp = Invoke-MgGraphRequest -Method GET -Uri $pimUri -ErrorAction Stop

                if ($pimResp -and $pimResp.value) {
                    foreach ($pimItem in $pimResp.value) {
                        $roleDefId = $pimItem.roleDefinitionId
                        $roleDefObj = if ($pimItem.roleDefinition) { $pimItem.roleDefinition } else { Get-CachedRoleDefinition -RoleDefinitionId $roleDefId }
                        $roleName   = if ($roleDefObj -and $roleDefObj.displayName) { $roleDefObj.displayName } else { $roleDefId }

                        Write-Log -Message "PIM Eligible Role Found: '$roleName'" -Level 'WARN' -Component 'PIM'

                        $record = [PSCustomObject]@{
                            AuditTimestamp      = $nowTimestamp
                            UserPrincipalName   = $userUpn
                            UserDisplayName     = $userDisplayName
                            AccountEnabled      = $accountEnabled
                            Workload            = "Entra ID"
                            ResourceType        = "Directory Role"
                            ResourceName        = $roleName
                            ResourceId          = $roleDefId
                            AccessLevel         = "PIM Eligible"
                            AssignmentMethod    = "PIM Eligible"
                            GrantedViaGroup     = "Direct"
                            GroupMembershipType = "N/A"
                            IsRoleAssignable    = $false
                            SyncSource          = "Cloud"
                            ResourceUrl         = ""
                        }
                        $Script:AllAuditRecords.Add($record)
                    }
                }
            }
            catch {
                Write-Log -Message "Unable to query PIM role eligibility schedule instances: $($_.Exception.Message)" -Level 'WARN' -Component 'PIM'
            }
        }

        # --------------------------------------------------------------------------
        # Phase 4/6: Auditing Microsoft Teams & Channels
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 4/6: Auditing Microsoft Teams & Channels..." -PercentComplete 60
        Write-Log -Message "Auditing Microsoft Teams and private/shared channel memberships..." -Level 'INFO' -Component 'TEAMS'

        foreach ($m365Grp in $userM365Groups) {
            $gId = $m365Grp.id
            $gName = $m365Grp.displayName
            $provOptions = if ($m365Grp.resourceProvisioningOptions) { @($m365Grp.resourceProvisioningOptions) } else { @() }

            if ($provOptions -contains 'Team') {
                $isOwner = $ownedGroupIds.Contains($gId)
                $teamAccessLvl = if ($isOwner) { "Owner" } else { "Member" }
                $isDirect = $directGroupIds.Contains($gId)

                Write-Log -Message "Team Member: '$gName' ($teamAccessLvl)" -Level 'INFO' -Component 'TEAMS'

                $record = [PSCustomObject]@{
                    AuditTimestamp      = $nowTimestamp
                    UserPrincipalName   = $userUpn
                    UserDisplayName     = $userDisplayName
                    AccountEnabled      = $accountEnabled
                    Workload            = "Microsoft Teams"
                    ResourceType        = "Team"
                    ResourceName        = $gName
                    ResourceId          = $gId
                    AccessLevel         = $teamAccessLvl
                    AssignmentMethod    = if ($isDirect) { "Direct" } else { "Nested / Group Inherited" }
                    GrantedViaGroup     = if ($isDirect) { "Direct" } else { "Nested" }
                    GroupMembershipType = "Assigned (Static)"
                    IsRoleAssignable    = $false
                    SyncSource          = "Cloud"
                    ResourceUrl         = ""
                }
                $Script:AllAuditRecords.Add($record)

                # Query Private & Shared Channels
                try {
                    $channelsUri = "https://graph.microsoft.com/v1.0/teams/$gId/channels`?`$filter=membershipType eq 'private' or membershipType eq 'shared'`&`$select=id,displayName,membershipType"
                    $chResp = Invoke-MgGraphRequest -Method GET -Uri $channelsUri -ErrorAction SilentlyContinue

                    if ($chResp -and $chResp.value) {
                        foreach ($channel in $chResp.value) {
                            $chId = $channel.id
                            $chName = $channel.displayName
                            $chType = $channel.membershipType

                            try {
                                $memberUri = "https://graph.microsoft.com/v1.0/teams/$gId/channels/$chId/members`?`$top=999"
                                $memberResp = Invoke-MgGraphRequest -Method GET -Uri $memberUri -ErrorAction Stop

                                if ($memberResp -and $memberResp.value) {
                                    $matchMember = $memberResp.value | Where-Object {
                                        $_.userId -eq $userId -or
                                        $_.email -eq $userUpn -or
                                        $_.displayName -eq $userDisplayName
                                    }

                                    if ($matchMember) {
                                        $chRole = if ($matchMember.roles -contains 'owner') { "Owner" } else { "Member" }
                                        $resourceChannelType = if ($chType -eq 'private') { "Private Channel" } else { "Shared Channel" }

                                        Write-Log -Message "Channel Access Found: '$gName > $chName' ($resourceChannelType - $chRole)" -Level 'INFO' -Component 'TEAMS'

                                        $chRecord = [PSCustomObject]@{
                                            AuditTimestamp      = $nowTimestamp
                                            UserPrincipalName   = $userUpn
                                            UserDisplayName     = $userDisplayName
                                            AccountEnabled      = $accountEnabled
                                            Workload            = "Microsoft Teams"
                                            ResourceType        = $resourceChannelType
                                            ResourceName        = "$gName > $chName"
                                            ResourceId          = $chId
                                            AccessLevel         = $chRole
                                            AssignmentMethod    = "Direct"
                                            GrantedViaGroup     = "Direct"
                                            GroupMembershipType = "Assigned (Static)"
                                            IsRoleAssignable    = $false
                                            SyncSource          = "Cloud"
                                            ResourceUrl         = ""
                                        }
                                        $Script:AllAuditRecords.Add($chRecord)
                                    }
                                }
                            }
                            catch {
                                Write-Log -Message "Could not inspect members of channel '$chName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'TEAMS'
                            }
                        }
                    }
                }
                catch {
                    Write-Log -Message "Could not retrieve channels for team '$gName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'TEAMS'
                }
            }
        }

        # --------------------------------------------------------------------------
        # Phase 5/6: Auditing SharePoint Online Sites & Permissions
        # --------------------------------------------------------------------------
        Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 5/6: Auditing SharePoint Online Sites & Permissions..." -PercentComplete 75
        Write-Log -Message "Auditing SharePoint Online site collections associated with user groups..." -Level 'INFO' -Component 'SPO'

        foreach ($m365Grp in $userM365Groups) {
            $gId = $m365Grp.id
            $gName = $m365Grp.displayName
            $siteObj = Get-CachedSiteForGroup -GroupId $gId

            if ($siteObj) {
                $siteUrl   = if ($siteObj.webUrl) { $siteObj.webUrl } else { "" }
                $siteTitle = if ($siteObj.displayName) { $siteObj.displayName } else { $gName }
                $siteId    = if ($siteObj.id) { $siteObj.id } else { $gId }
                $isOwner   = $ownedGroupIds.Contains($gId)
                $spoAccess = if ($isOwner) { "Site Owner" } else { "Site Member" }
                $isDirect  = $directGroupIds.Contains($gId)

                Write-Log -Message "SPO Site Associated: '$siteTitle' ($spoAccess) -> $siteUrl" -Level 'INFO' -Component 'SPO'

                $record = [PSCustomObject]@{
                    AuditTimestamp      = $nowTimestamp
                    UserPrincipalName   = $userUpn
                    UserDisplayName     = $userDisplayName
                    AccountEnabled      = $accountEnabled
                    Workload            = "SharePoint Online"
                    ResourceType        = "SPO Site"
                    ResourceName        = $siteTitle
                    ResourceId          = $siteId
                    AccessLevel         = $spoAccess
                    AssignmentMethod    = if ($isDirect) { "Direct" } else { "Nested / Group Inherited" }
                    GrantedViaGroup     = $gName
                    GroupMembershipType = "Assigned (Static)"
                    IsRoleAssignable    = $false
                    SyncSource          = "Cloud"
                    ResourceUrl         = $siteUrl
                }
                $Script:AllAuditRecords.Add($record)

                # Deep SharePoint Permission Audit
                if ($IncludeSharePoint -and -not [string]::IsNullOrWhiteSpace($siteId)) {
                    try {
                        $sitePermUri = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions`?`$top=999"
                        $sitePermResp = Invoke-MgGraphRequest -Method GET -Uri $sitePermUri -ErrorAction Stop

                        if ($sitePermResp -and $sitePermResp.value) {
                            foreach ($p in $sitePermResp.value) {
                                $grantedRoles = if ($p.roles) { ($p.roles -join ', ') } else { "Custom" }
                                $granteeId = if ($p.grantedToV2 -and $p.grantedToV2.user) { $p.grantedToV2.user.id }
                                             elseif ($p.grantedToV2 -and $p.grantedToV2.group) { $p.grantedToV2.group.id }
                                             else { $null }

                                if ($granteeId -eq $userId) {
                                    $record = [PSCustomObject]@{
                                        AuditTimestamp      = $nowTimestamp
                                        UserPrincipalName   = $userUpn
                                        UserDisplayName     = $userDisplayName
                                        AccountEnabled      = $accountEnabled
                                        Workload            = "SharePoint Online"
                                        ResourceType        = "SPO Site Permission"
                                        ResourceName        = "$siteTitle (Explicit Grant)"
                                        ResourceId          = $siteId
                                        AccessLevel         = $grantedRoles
                                        AssignmentMethod    = "Direct"
                                        GrantedViaGroup     = "Direct"
                                        GroupMembershipType = "N/A"
                                        IsRoleAssignable    = $false
                                        SyncSource          = "Cloud"
                                        ResourceUrl         = $siteUrl
                                    }
                                    $Script:AllAuditRecords.Add($record)
                                }
                                elseif ($allGroupIds.Contains($granteeId)) {
                                    $matchedGrp = Get-CachedGroup -GroupId $granteeId
                                    $matchedGrpName = if ($matchedGrp) { $matchedGrp.displayName } else { $granteeId }

                                    $record = [PSCustomObject]@{
                                        AuditTimestamp      = $nowTimestamp
                                        UserPrincipalName   = $userUpn
                                        UserDisplayName     = $userDisplayName
                                        AccountEnabled      = $accountEnabled
                                        Workload            = "SharePoint Online"
                                        ResourceType        = "SPO Site Permission"
                                        ResourceName        = "$siteTitle via $matchedGrpName"
                                        ResourceId          = $siteId
                                        AccessLevel         = $grantedRoles
                                        AssignmentMethod    = "Nested / Group Inherited"
                                        GrantedViaGroup     = $matchedGrpName
                                        GroupMembershipType = "N/A"
                                        IsRoleAssignable    = $false
                                        SyncSource          = "Cloud"
                                        ResourceUrl         = $siteUrl
                                    }
                                    $Script:AllAuditRecords.Add($record)
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log -Message "Could not audit permissions for site '$siteTitle': $($_.Exception.Message)" -Level 'DEBUG' -Component 'SPO'
                    }
                }
            }
        }

        # --------------------------------------------------------------------------
        # Phase 6/6: Auditing Exchange Online Mailbox Permissions
        # --------------------------------------------------------------------------
        if ($IncludeExchange) {
            Write-Progress -Activity "Auditing User: $TargetUser" -Status "Phase 6/6: Auditing Exchange Online Permissions..." -PercentComplete 90
            Write-Log -Message "Auditing Exchange Online mailbox permissions (FullAccess, SendAs, SendOnBehalf)..." -Level 'INFO' -Component 'EXO'

            try {
                $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox,UserMailbox -ErrorAction Stop

                Write-Log -Message "Scanning delegations across $($mailboxes.Count) mailboxes..." -Level 'INFO' -Component 'EXO'

                foreach ($mbx in $mailboxes) {
                    $mbxIdentity = $mbx.Identity.ToString()
                    $mbxName     = $mbx.DisplayName
                    $mbxEmail    = $mbx.PrimarySmtpAddress

                    # 1. FullAccess Permissions
                    try {
                        $fullAccessPerms = Get-MailboxPermission -Identity $mbxIdentity -User $userUpn -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }

                        if ($fullAccessPerms) {
                            Write-Log -Message "FullAccess found on mailbox '$mbxName' ($mbxEmail)" -Level 'WARN' -Component 'EXO'
                            $record = [PSCustomObject]@{
                                AuditTimestamp      = $nowTimestamp
                                UserPrincipalName   = $userUpn
                                UserDisplayName     = $userDisplayName
                                AccountEnabled      = $accountEnabled
                                Workload            = "Exchange Online"
                                ResourceType        = "Mailbox"
                                ResourceName        = $mbxName
                                ResourceId          = $mbxIdentity
                                AccessLevel         = "FullAccess"
                                AssignmentMethod    = "Direct"
                                GrantedViaGroup     = "Direct"
                                GroupMembershipType = "N/A"
                                IsRoleAssignable    = $false
                                SyncSource          = "Cloud"
                                ResourceUrl         = ""
                            }
                            $Script:AllAuditRecords.Add($record)
                        }
                    }
                    catch {
                        Write-Log -Message "Error checking FullAccess on '$mbxName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'EXO'
                    }

                    # 2. SendAs Permissions
                    try {
                        $sendAsPerms = Get-RecipientPermission -Identity $mbxIdentity -Trustee $userUpn -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'SendAs' }

                        if ($sendAsPerms) {
                            Write-Log -Message "SendAs found on mailbox '$mbxName' ($mbxEmail)" -Level 'WARN' -Component 'EXO'
                            $record = [PSCustomObject]@{
                                AuditTimestamp      = $nowTimestamp
                                UserPrincipalName   = $userUpn
                                UserDisplayName     = $userDisplayName
                                AccountEnabled      = $accountEnabled
                                Workload            = "Exchange Online"
                                ResourceType        = "Mailbox"
                                ResourceName        = $mbxName
                                ResourceId          = $mbxIdentity
                                AccessLevel         = "SendAs"
                                AssignmentMethod    = "Direct"
                                GrantedViaGroup     = "Direct"
                                GroupMembershipType = "N/A"
                                IsRoleAssignable    = $false
                                SyncSource          = "Cloud"
                                ResourceUrl         = ""
                            }
                            $Script:AllAuditRecords.Add($record)
                        }
                    }
                    catch {
                        Write-Log -Message "Error checking SendAs on '$mbxName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'EXO'
                    }

                    # 3. SendOnBehalf Permissions
                    if ($mbx.GrantSendOnBehalfTo) {
                        $grantList = @($mbx.GrantSendOnBehalfTo)
                        if ($grantList -contains $userUpn -or $grantList -contains $userDisplayName) {
                            Write-Log -Message "SendOnBehalf found on mailbox '$mbxName' ($mbxEmail)" -Level 'WARN' -Component 'EXO'
                            $record = [PSCustomObject]@{
                                AuditTimestamp      = $nowTimestamp
                                UserPrincipalName   = $userUpn
                                UserDisplayName     = $userDisplayName
                                AccountEnabled      = $accountEnabled
                                Workload            = "Exchange Online"
                                ResourceType        = "Mailbox"
                                ResourceName        = $mbxName
                                ResourceId          = $mbxIdentity
                                AccessLevel         = "SendOnBehalf"
                                AssignmentMethod    = "Direct"
                                GrantedViaGroup     = "Direct"
                                GroupMembershipType = "N/A"
                                IsRoleAssignable    = $false
                                SyncSource          = "Cloud"
                                ResourceUrl         = ""
                            }
                            $Script:AllAuditRecords.Add($record)
                        }
                    }

                    # 4. Check Group-Inherited Mailbox Permissions for Mail-Enabled Security Groups
                    foreach ($secGrp in $userSecurityGroups) {
                        if ($secGrp.mailEnabled -eq $true) {
                            $secGrpName = $secGrp.displayName
                            try {
                                $grpFullAccess = Get-MailboxPermission -Identity $mbxIdentity -User $secGrpName -ErrorAction SilentlyContinue |
                                    Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }

                                if ($grpFullAccess) {
                                    Write-Log -Message "FullAccess via Group '$secGrpName' on mailbox '$mbxName'" -Level 'WARN' -Component 'EXO'
                                    $record = [PSCustomObject]@{
                                        AuditTimestamp      = $nowTimestamp
                                        UserPrincipalName   = $userUpn
                                        UserDisplayName     = $userDisplayName
                                        AccountEnabled      = $accountEnabled
                                        Workload            = "Exchange Online"
                                        ResourceType        = "Mailbox"
                                        ResourceName        = $mbxName
                                        ResourceId          = $mbxIdentity
                                        AccessLevel         = "FullAccess (Group Inherited)"
                                        AssignmentMethod    = "Nested / Group Inherited"
                                        GrantedViaGroup     = $secGrpName
                                        GroupMembershipType = "N/A"
                                        IsRoleAssignable    = $false
                                        SyncSource          = "Cloud"
                                        ResourceUrl         = ""
                                    }
                                    $Script:AllAuditRecords.Add($record)
                                }
                            }
                            catch {
                                Write-Log -Message "Could not query mailbox permission for group '$secGrpName': $($_.Exception.Message)" -Level 'DEBUG' -Component 'EXO'
                            }
                        }
                    }
                }
            }
            catch {
                Write-Log -Message "Exchange Online mailbox audit error: $($_.Exception.Message)" -Level 'ERROR' -Component 'EXO'
            }
        }

        Write-Progress -Activity "Auditing User: $TargetUser" -Completed
        Write-Log -Message "Completed audit for user '$userUpn'." -Level 'SUCCESS' -Component 'AUDIT'
    }

    # ==============================================================================
    # REPORT EXPORT GENERATOR (HTML, CSV, JSON)
    # ==============================================================================
    function Export-AuditReport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string[]]$Formats,

            [Parameter(Mandatory = $true)]
            [string]$DestinationPath
        )

        Write-Log -Message "================================================================================" -Level 'INFO' -Component 'REPORT'
        Write-Log -Message "Aggregating and exporting audit results ($($Records.Count) total records)..." -Level 'INFO' -Component 'REPORT'

        try {
            if (-not (Test-Path -Path $DestinationPath)) {
                $null = New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            Write-Log -Message "Unable to create output directory '$DestinationPath'. Falling back to current folder." -Level 'WARN' -Component 'REPORT'
            $DestinationPath = "."
        }

        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $baseFileName = "Get-M365UserAccessAudit_${timestamp}"

        # --------------------------------------------------------------------------
        # 1. CSV EXPORT (UTF-8 with BOM for MS Excel compatibility)
        # --------------------------------------------------------------------------
        if ($Formats -contains 'CSV') {
            $csvPath = Join-Path -Path $DestinationPath -ChildPath "${baseFileName}.csv"
            try {
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    $Records | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8BOM -Force
                }
                else {
                    $Records | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                }
                $Script:GeneratedCsvPath = $csvPath
                Write-Log -Message "CSV Report generated: $csvPath" -Level 'SUCCESS' -Component 'REPORT'
            }
            catch {
                Write-Log -Message "Failed to export CSV report: $($_.Exception.Message)" -Level 'ERROR' -Component 'REPORT'
            }
        }

        # --------------------------------------------------------------------------
        # 2. JSON EXPORT
        # --------------------------------------------------------------------------
        if ($Formats -contains 'JSON') {
            $jsonPath = Join-Path -Path $DestinationPath -ChildPath "${baseFileName}.json"
            try {
                $jsonContent = $Records | ConvertTo-Json -Depth 5
                [System.IO.File]::WriteAllText($jsonPath, $jsonContent, [System.Text.Encoding]::UTF8)
                $Script:GeneratedJsonPath = $jsonPath
                Write-Log -Message "JSON Report generated: $jsonPath" -Level 'SUCCESS' -Component 'REPORT'
            }
            catch {
                Write-Log -Message "Failed to export JSON report: $($_.Exception.Message)" -Level 'ERROR' -Component 'REPORT'
            }
        }

        # --------------------------------------------------------------------------
        # 3. INTERACTIVE HTML REPORT
        # --------------------------------------------------------------------------
        if ($Formats -contains 'HTML') {
            $htmlPath = Join-Path -Path $DestinationPath -ChildPath "${baseFileName}.html"
            try {
                $totalCount      = $Records.Count
                $securityGroups  = @($Records | Where-Object { $_.ResourceType -in @('Security Group', 'Mail-Enabled Security Group') }).Count
                $m365Groups      = @($Records | Where-Object { $_.ResourceType -eq 'M365 Group' }).Count
                $directoryRoles  = @($Records | Where-Object { $_.ResourceType -eq 'Directory Role' }).Count
                $teamsChannels   = @($Records | Where-Object { $_.Workload -eq 'Microsoft Teams' }).Count
                $spoCount        = @($Records | Where-Object { $_.Workload -eq 'SharePoint Online' }).Count
                $exoCount        = @($Records | Where-Object { $_.Workload -eq 'Exchange Online' -and $_.ResourceType -eq 'Mailbox' }).Count

                $uniqueUsersList = @($Records | Select-Object -ExpandProperty UserPrincipalName -Unique)
                $uniqueUsers     = $uniqueUsersList.Count

                $userInfoBlock = ""
                if ($uniqueUsers -eq 1 -and $Records.Count -gt 0) {
                    $uSample = $Records[0]
                    $statusBadge = if ($uSample.AccountEnabled) { "<span class='badge badge-member'>Enabled</span>" } else { "<span class='badge badge-danger'>Disabled</span>" }
                    $avatarInitial = if (-not [string]::IsNullOrWhiteSpace($uSample.UserDisplayName)) { $uSample.UserDisplayName.Substring(0,1).ToUpper() } else { "U" }
                    $userInfoBlock = @"
        <section class="user-hero-card">
            <div class="user-hero-avatar">$avatarInitial</div>
            <div class="user-hero-info">
                <h2>$([System.Net.WebUtility]::HtmlEncode($uSample.UserDisplayName))</h2>
                <div class="user-hero-tags">
                    <span>UPN: <strong>$([System.Net.WebUtility]::HtmlEncode($uSample.UserPrincipalName))</strong></span>
                    <span>Account: $statusBadge</span>
                    <span>Authority: <strong>$($uSample.SyncSource)</strong></span>
                </div>
            </div>
        </section>
"@
                }

                $tableRowsBuilder = [System.Text.StringBuilder]::new()
                $rowNum = 1
                foreach ($r in $Records) {
                    $workloadClass = switch ($r.Workload) {
                        'Entra ID'          { 'badge-entra' }
                        'Microsoft Teams'   { 'badge-teams' }
                        'SharePoint Online' { 'badge-spo' }
                        'Exchange Online'   { 'badge-exo' }
                        Default             { 'badge-neutral' }
                    }

                    $accessBadgeClass = switch -Regex ($r.AccessLevel) {
                        'Owner'          { 'badge-owner' }
                        'Role|Admin'     { 'badge-admin' }
                        'FullAccess'     { 'badge-danger' }
                        'SendAs|SendOn'  { 'badge-warn' }
                        'Eligible'       { 'badge-pim' }
                        Default          { 'badge-member' }
                    }

                    $assignBadgeClass = switch ($r.AssignmentMethod) {
                        'Direct'                    { 'badge-direct' }
                        'Nested / Group Inherited'  { 'badge-nested' }
                        'Dynamic Rule'              { 'badge-dynamic' }
                        'PIM Eligible'              { 'badge-pim' }
                        Default                     { 'badge-neutral' }
                    }

                    $urlCell = if (-not [string]::IsNullOrWhiteSpace($r.ResourceUrl)) {
                        "<a href='$($r.ResourceUrl)' target='_blank' rel='noopener noreferrer' class='site-link'>Open Link &nearr;</a>"
                    }
                    else {
                        "<span class='text-muted'>&mdash;</span>"
                    }

                    $syncBadge = if ($r.SyncSource -eq 'On-Premises AD') {
                        "<span class='badge badge-sync-onprem'>On-Prem AD</span>"
                    }
                    else {
                        "<span class='badge badge-sync-cloud'>Cloud</span>"
                    }

                    $roleAssignBadge = if ($r.IsRoleAssignable) {
                        "<span class='badge badge-roleassign-true'>Yes</span>"
                    }
                    else {
                        "<span class='badge badge-roleassign-false'>No</span>"
                    }

                    $rowHtml = @"
<tr>
    <td class='text-center text-muted mono'>$rowNum</td>
    <td><strong>$([System.Net.WebUtility]::HtmlEncode($r.UserDisplayName))</strong><br><small class='text-muted'>$([System.Net.WebUtility]::HtmlEncode($r.UserPrincipalName))</small></td>
    <td><span class='badge $workloadClass'>$($r.Workload)</span></td>
    <td>$([System.Net.WebUtility]::HtmlEncode($r.ResourceType))</td>
    <td><strong>$([System.Net.WebUtility]::HtmlEncode($r.ResourceName))</strong><br><small class='text-muted mono'>$($r.ResourceId)</small></td>
    <td><span class='badge $accessBadgeClass'>$($r.AccessLevel)</span></td>
    <td><span class='badge $assignBadgeClass'>$($r.AssignmentMethod)</span></td>
    <td>$([System.Net.WebUtility]::HtmlEncode($r.GrantedViaGroup))</td>
    <td>$($r.GroupMembershipType)</td>
    <td class='text-center'>$roleAssignBadge</td>
    <td class='text-center'>$syncBadge</td>
    <td>$urlCell</td>
</tr>
"@
                    $null = $tableRowsBuilder.AppendLine($rowHtml)
                    $rowNum++
                }

                $Script:GeneratedHtmlPath = $htmlPath

                $htmlDocument = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 User Access Audit Report</title>
    <style>
        :root {
            --bg-color: #0f172a;
            --surface-color: #1e293b;
            --surface-border: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --primary: #38bdf8;
            --primary-hover: #0284c7;
            --accent-purple: #a855f7;
            --accent-emerald: #10b981;
            --accent-rose: #f43f5e;
            --accent-amber: #f59e0b;
            --accent-indigo: #6366f1;
            --accent-cyan: #06b6d4;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-main);
            padding: 24px;
            line-height: 1.5;
        }

        .container {
            max-width: 1600px;
            margin: 0 auto;
        }

        .header {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            border: 1px solid var(--surface-border);
            border-radius: 16px;
            padding: 28px 32px;
            margin-bottom: 24px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .header-title h1 {
            font-size: 1.85rem;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.025em;
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .header-title p {
            color: var(--text-muted);
            margin-top: 6px;
            font-size: 0.95rem;
        }

        .header-meta {
            text-align: right;
            font-size: 0.85rem;
            color: var(--text-muted);
        }

        .header-meta strong {
            color: var(--primary);
        }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .kpi-card {
            background-color: var(--surface-color);
            border: 1px solid var(--surface-border);
            border-radius: 12px;
            padding: 20px;
            position: relative;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px -2px rgba(0, 0, 0, 0.3);
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
        }

        .kpi-card.blue::before { background-color: var(--primary); }
        .kpi-card.purple::before { background-color: var(--accent-purple); }
        .kpi-card.emerald::before { background-color: var(--accent-emerald); }
        .kpi-card.rose::before { background-color: var(--accent-rose); }
        .kpi-card.indigo::before { background-color: var(--accent-indigo); }
        .kpi-card.cyan::before { background-color: var(--accent-cyan); }
        .kpi-card.amber::before { background-color: var(--accent-amber); }

        .kpi-title {
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-muted);
            margin-bottom: 8px;
            font-weight: 600;
        }

        .kpi-value {
            font-size: 2rem;
            font-weight: 800;
            color: #ffffff;
        }

        .controls-card {
            background-color: var(--surface-color);
            border: 1px solid var(--surface-border);
            border-radius: 12px;
            padding: 16px 20px;
            margin-bottom: 20px;
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            align-items: center;
            justify-content: space-between;
        }

        .search-box {
            position: relative;
            flex: 1;
            min-width: 280px;
        }

        .search-input {
            width: 100%;
            background-color: #0f172a;
            border: 1px solid var(--surface-border);
            border-radius: 8px;
            padding: 10px 16px;
            color: #fff;
            font-size: 0.95rem;
            outline: none;
            transition: border-color 0.2s;
        }

        .search-input:focus {
            border-color: var(--primary);
        }

        .filter-group {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }

        .filter-btn {
            background-color: #0f172a;
            border: 1px solid var(--surface-border);
            color: var(--text-muted);
            padding: 8px 14px;
            border-radius: 6px;
            font-size: 0.85rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }

        .filter-btn:hover {
            color: #fff;
            border-color: var(--primary);
        }

        .filter-btn.active {
            background-color: var(--primary);
            color: #0f172a;
            border-color: var(--primary);
        }

        .record-counter {
            font-size: 0.85rem;
            color: var(--text-muted);
            font-weight: 600;
        }

        .table-container {
            background-color: var(--surface-color);
            border: 1px solid var(--surface-border);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.875rem;
            text-align: left;
            min-width: 1200px;
        }

        th {
            background-color: #172033;
            color: var(--text-muted);
            font-weight: 600;
            padding: 14px 16px;
            border-bottom: 1px solid var(--surface-border);
            white-space: nowrap;
            cursor: pointer;
            user-select: none;
        }

        th:hover {
            color: #fff;
        }

        td {
            padding: 12px 16px;
            border-bottom: 1px solid rgba(51, 65, 85, 0.4);
            vertical-align: middle;
        }

        tr:nth-child(even) {
            background-color: rgba(15, 23, 42, 0.25);
        }

        tr:hover {
            background-color: rgba(56, 189, 248, 0.05);
        }

        .badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 600;
            white-space: nowrap;
        }

        .badge-entra   { background-color: rgba(56, 189, 248, 0.15); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); }
        .badge-teams   { background-color: rgba(99, 102, 241, 0.15); color: #818cf8; border: 1px solid rgba(99, 102, 241, 0.3); }
        .badge-spo     { background-color: rgba(6, 182, 212, 0.15);  color: #22d3ee; border: 1px solid rgba(6, 182, 212, 0.3); }
        .badge-exo     { background-color: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }
        .badge-neutral { background-color: rgba(148, 163, 184, 0.15); color: #94a3b8; border: 1px solid rgba(148, 163, 184, 0.3); }

        .badge-owner   { background-color: rgba(168, 85, 247, 0.2); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.4); }
        .badge-admin   { background-color: rgba(244, 63, 94, 0.2);  color: #fb7185; border: 1px solid rgba(244, 63, 94, 0.4); }
        .badge-danger  { background-color: rgba(239, 68, 68, 0.2);  color: #f87171; border: 1px solid rgba(239, 68, 68, 0.4); }
        .badge-warn    { background-color: rgba(245, 158, 11, 0.2); color: #fcd34d; border: 1px solid rgba(245, 158, 11, 0.4); }
        .badge-member  { background-color: rgba(16, 185, 129, 0.15); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.3); }
        .badge-pim     { background-color: rgba(236, 72, 153, 0.2); color: #f472b6; border: 1px solid rgba(236, 72, 153, 0.4); }

        .badge-direct  { background-color: rgba(16, 185, 129, 0.15); color: #34d399; }
        .badge-nested  { background-color: rgba(245, 158, 11, 0.15); color: #fbbf24; }
        .badge-dynamic { background-color: rgba(6, 182, 212, 0.15);  color: #22d3ee; }

        .badge-sync-cloud  { background-color: rgba(56, 189, 248, 0.1); color: #38bdf8; }
        .badge-sync-onprem { background-color: rgba(168, 85, 247, 0.1); color: #c084fc; }

        .badge-roleassign-true  { background-color: rgba(244, 63, 94, 0.15); color: #fb7185; }
        .badge-roleassign-false { background-color: rgba(148, 163, 184, 0.1); color: #64748b; }

        .site-link {
            color: var(--primary);
            text-decoration: none;
            font-weight: 600;
        }

        .site-link:hover {
            text-decoration: underline;
        }

        .mono {
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            font-size: 0.75rem;
        }

        .text-muted {
            color: var(--text-muted);
        }

        .text-center {
            text-align: center;
        }

        .footer {
            margin-top: 30px;
            text-align: center;
            font-size: 0.8rem;
            color: var(--text-muted);
            border-top: 1px solid var(--surface-border);
            padding-top: 20px;
        }

        .footer a {
            color: var(--primary);
            text-decoration: none;
        }

        .user-hero-card {
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0.9) 100%);
            border: 1px solid var(--surface-border);
            border-left: 4px solid var(--primary);
            border-radius: 12px;
            padding: 16px 24px;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            gap: 20px;
        }

        .user-hero-avatar {
            width: 52px;
            height: 52px;
            border-radius: 50%;
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent-indigo) 100%);
            color: #0f172a;
            font-size: 1.5rem;
            font-weight: 800;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 10px rgba(56, 189, 248, 0.3);
            flex-shrink: 0;
        }

        .user-hero-info h2 {
            font-size: 1.35rem;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 4px;
        }

        .user-hero-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            font-size: 0.875rem;
            color: var(--text-muted);
        }

        .user-hero-tags span {
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .user-hero-tags strong {
            color: var(--text-main);
        }
    </style>
</head>
<body>
    <div class="container">
        <header class="header">
            <div class="header-title">
                <h1>
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                    </svg>
                    Microsoft 365 User Access Audit Report
                </h1>
                <p>Enterprise Permission & Membership Matrix for Entra ID, Teams, SharePoint & Exchange Online</p>
            </div>
            <div class="header-meta">
                <div>Tenant ID: <strong>$($Script:ActiveTenantId)</strong></div>
                <div>Tenant Name: <strong>$($Script:ActiveTenantName)</strong></div>
                <div>Generated: <strong>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong></div>
                <div>Audited Users: <strong>$uniqueUsers</strong></div>
            </div>
        </header>

$userInfoBlock

        <section class="kpi-grid">
            <div class="kpi-card blue">
                <div class="kpi-title">Total Audit Entries</div>
                <div class="kpi-value">$totalCount</div>
            </div>
            <div class="kpi-card emerald">
                <div class="kpi-title">Security Groups</div>
                <div class="kpi-value">$securityGroups</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-title">M365 Unified Groups</div>
                <div class="kpi-value">$m365Groups</div>
            </div>
            <div class="kpi-card rose">
                <div class="kpi-title">Directory Roles & PIM</div>
                <div class="kpi-value">$directoryRoles</div>
            </div>
            <div class="kpi-card indigo">
                <div class="kpi-title">Teams & Channels</div>
                <div class="kpi-value">$teamsChannels</div>
            </div>
            <div class="kpi-card cyan">
                <div class="kpi-title">SharePoint Sites</div>
                <div class="kpi-value">$spoCount</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-title">Exchange Mailboxes</div>
                <div class="kpi-value">$exoCount</div>
            </div>
        </section>

        <section class="controls-card">
            <div class="search-box">
                <input type="text" id="searchInput" class="search-input" placeholder="Live search across all columns (user, group, role, site, access level)...">
            </div>
            <div class="filter-group">
                <button class="filter-btn active" data-filter="all">All Workloads</button>
                <button class="filter-btn" data-filter="Entra ID">Entra ID</button>
                <button class="filter-btn" data-filter="Microsoft Teams">Teams</button>
                <button class="filter-btn" data-filter="SharePoint Online">SharePoint</button>
                <button class="filter-btn" data-filter="Exchange Online">Exchange</button>
            </div>
            <div class="record-counter" id="recordCounter">Showing $totalCount of $totalCount records</div>
        </section>

        <main class="table-container">
            <table id="auditTable">
                <thead>
                    <tr>
                        <th onclick="sortTable(0)" class="text-center" style="width: 50px;">Nr &#x21C5;</th>
                        <th onclick="sortTable(1)">User &#x21C5;</th>
                        <th onclick="sortTable(2)">Workload &#x21C5;</th>
                        <th onclick="sortTable(3)">Resource Type &#x21C5;</th>
                        <th onclick="sortTable(4)">Resource Name & ID &#x21C5;</th>
                        <th onclick="sortTable(5)">Access Level &#x21C5;</th>
                        <th onclick="sortTable(6)">Assignment &#x21C5;</th>
                        <th onclick="sortTable(7)">Granted Via &#x21C5;</th>
                        <th onclick="sortTable(8)">Group Type &#x21C5;</th>
                        <th onclick="sortTable(9)" class="text-center">Role Assignable &#x21C5;</th>
                        <th onclick="sortTable(10)" class="text-center">Sync Source &#x21C5;</th>
                        <th>Resource URL</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
$($tableRowsBuilder.ToString())
                </tbody>
            </table>
        </main>

        <footer class="footer">
            <p>Generated by <strong>$($Script:Metadata.Title)</strong> v$($Script:Metadata.Version) | Author: <strong>$($Script:Metadata.Author)</strong> (<a href="mailto:$($Script:Metadata.Email)">$($Script:Metadata.Email)</a>) | <a href="$($Script:Metadata.GitHub)" target="_blank">GitHub Repository</a></p>
            <p style="margin-top: 4px;">Log file: <span class="mono">$Script:LogFilePath</span></p>
        </footer>
    </div>

    <script>
        const searchInput = document.getElementById('searchInput');
        const tableBody = document.getElementById('tableBody');
        const rows = Array.from(tableBody.getElementsByTagName('tr'));
        const counter = document.getElementById('recordCounter');
        const filterBtns = document.querySelectorAll('.filter-btn');

        let activeFilter = 'all';

        function updateFilter() {
            const query = searchInput.value.toLowerCase().trim();
            let visibleCount = 0;

            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                const workloadCell = row.cells[2] ? row.cells[2].textContent.trim() : '';

                const matchesSearch = query === '' || text.includes(query);
                const matchesFilter = (activeFilter === 'all') || (workloadCell === activeFilter);

                if (matchesSearch && matchesFilter) {
                    row.style.display = '';
                    visibleCount++;
                } else {
                    row.style.display = 'none';
                }
            });

            counter.textContent = `Showing ` + visibleCount + ` of ` + rows.length + ` records`;
        }

        searchInput.addEventListener('input', updateFilter);

        filterBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                filterBtns.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                activeFilter = btn.getAttribute('data-filter');
                updateFilter();
            });
        });

        function sortTable(n) {
            const table = document.getElementById("auditTable");
            let rowsArray = Array.from(table.rows).slice(1);
            let dir = table.getAttribute("data-sort-dir-" + n) === "asc" ? "desc" : "asc";
            table.setAttribute("data-sort-dir-" + n, dir);

            rowsArray.sort((a, b) => {
                let x = a.cells[n].textContent.trim().toLowerCase();
                let y = b.cells[n].textContent.trim().toLowerCase();
                if (dir === "asc") {
                    return x.localeCompare(y);
                } else {
                    return y.localeCompare(x);
                }
            });

            rowsArray.forEach(row => tableBody.appendChild(row));
        }
    </script>
</body>
</html>
"@
                [System.IO.File]::WriteAllText($htmlPath, $htmlDocument, [System.Text.Encoding]::UTF8)
                Write-Log -Message "Interactive HTML Report generated: $htmlPath" -Level 'SUCCESS' -Component 'REPORT'
            }
            catch {
                Write-Log -Message "Failed to export HTML report: $($_.Exception.Message)" -Level 'ERROR' -Component 'REPORT'
            }
        }

        Write-Log -Message "Reporting complete. All selected reports saved to: '$DestinationPath'" -Level 'SUCCESS' -Component 'REPORT'
    }

    # ==============================================================================
    # SUMMARY DISPLAY (Console Human-Readable Categorized Report & Metrics)
    # ==============================================================================
    function Show-AuditSummary {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $true)]
            [string]$ReportsDir,

            [Parameter(Mandatory = $true)]
            [string]$LogFile,

            [Parameter(Mandatory = $false)]
            [string]$HtmlFilePath,

            [Parameter(Mandatory = $false)]
            [string]$CsvFilePath,

            [Parameter(Mandatory = $false)]
            [switch]$ExchangeAudited
        )

        Write-Host ""
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " MICROSOFT 365 USER ACCESS AUDIT - EXECUTIVE & CONSOLE REPORT" -ForegroundColor Yellow
        Write-Host "================================================================================" -ForegroundColor Cyan

        $uniqueUsers = @($Records | Select-Object -ExpandProperty UserPrincipalName -Unique)
        if ($uniqueUsers.Count -eq 1) {
            $sample = $Records[0]
            $statusStr = if ($sample.AccountEnabled) { "Active / Enabled" } else { "Disabled" }
            Write-Host "Audited User    : " -NoNewline -ForegroundColor White
            Write-Host "$($sample.UserDisplayName) ($($sample.UserPrincipalName))" -ForegroundColor Cyan
            Write-Host "Account Status  : " -NoNewline -ForegroundColor White
            Write-Host "$statusStr" -ForegroundColor (if ($sample.AccountEnabled) { "Green" } else { "Red" }) -NoNewline
            Write-Host " | Sync Authority: " -NoNewline -ForegroundColor White
            Write-Host "$($sample.SyncSource)" -ForegroundColor Yellow
        }
        else {
            Write-Host "Audited Users   : " -NoNewline -ForegroundColor White
            Write-Host "$($uniqueUsers.Count) user(s): $($uniqueUsers -join ', ')" -ForegroundColor Cyan
        }

        if (-not [string]::IsNullOrWhiteSpace($Script:ActiveTenantName)) {
            Write-Host "Tenant Name     : " -NoNewline -ForegroundColor White
            Write-Host "$($Script:ActiveTenantName)" -ForegroundColor Magenta -NoNewline
            Write-Host " ($($Script:ActiveTenantId))" -ForegroundColor DarkGray
        }
        Write-Host "Total Records   : " -NoNewline -ForegroundColor White
        Write-Host "$($Records.Count)" -ForegroundColor Green
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray

        # --------------------------------------------------------------------------
        # [1] ENTRA ID DIRECTORY ROLES & PRIVILEGED ACCESS
        # --------------------------------------------------------------------------
        $roles = @($Records | Where-Object { $_.ResourceType -in @('Directory Role', 'PIM Eligible Role') })
        Write-Host "`n[1] ENTRA ID DIRECTORY ROLES & PRIVILEGED ACCESS ($($roles.Count))" -ForegroundColor (if ($roles.Count -gt 0) { "Red" } else { "Green" })
        if ($roles.Count -gt 0) {
            $rIdx = 0
            $roleRows = foreach ($item in $roles) {
                $rIdx++
                [PSCustomObject]@{
                    '#'           = $rIdx
                    'Role Name'   = $item.ResourceName
                    'Access Level'= $item.AccessLevel
                    'Assignment'  = $item.AssignmentMethod
                    'Role Assign' = if ($item.IsRoleAssignable) { 'Yes' } else { 'No' }
                    'Granted Via' = $item.GrantedViaGroup
                }
            }
            $roleRows | Format-Table -AutoSize | Out-Host
        }
        else {
            Write-Host "    [i] No Entra ID Directory Roles or PIM privileged assignments (Standard User)." -ForegroundColor DarkGray
        }

        # --------------------------------------------------------------------------
        # [2] MICROSOFT TEAMS & CHANNELS
        # --------------------------------------------------------------------------
        $teams = @($Records | Where-Object { $_.Workload -eq 'Microsoft Teams' })
        Write-Host "`n[2] MICROSOFT TEAMS & CHANNELS ($($teams.Count))" -ForegroundColor Magenta
        if ($teams.Count -gt 0) {
            $tIdx = 0
            $teamRows = foreach ($item in $teams) {
                $tIdx++
                [PSCustomObject]@{
                    '#'             = $tIdx
                    'Team / Channel'= $item.ResourceName
                    'Type'          = $item.ResourceType
                    'Role / Access' = $item.AccessLevel
                    'Assignment'    = $item.AssignmentMethod
                }
            }
            $teamRows | Format-Table -AutoSize | Out-Host
        }
        else {
            Write-Host "    [i] No Microsoft Teams or Channel memberships found." -ForegroundColor DarkGray
        }

        # --------------------------------------------------------------------------
        # [3] SHAREPOINT ONLINE SITE COLLECTIONS & PERMISSIONS
        # --------------------------------------------------------------------------
        $spo = @($Records | Where-Object { $_.Workload -eq 'SharePoint Online' })
        Write-Host "`n[3] SHAREPOINT ONLINE SITE COLLECTIONS & PERMISSIONS ($($spo.Count))" -ForegroundColor Cyan
        if ($spo.Count -gt 0) {
            $sIdx = 0
            $spoRows = foreach ($item in $spo) {
                $sIdx++
                [PSCustomObject]@{
                    '#'           = $sIdx
                    'Site Name'   = $item.ResourceName
                    'Access Level'= $item.AccessLevel
                    'Assignment'  = $item.AssignmentMethod
                    'Granted Via' = $item.GrantedViaGroup
                    'Site URL'    = $item.ResourceUrl
                }
            }
            $spoRows | Format-Table -AutoSize | Out-Host
        }
        else {
            Write-Host "    [i] No SharePoint Online site associations found." -ForegroundColor DarkGray
        }

        # --------------------------------------------------------------------------
        # [4] ENTRA ID & M365 GROUPS (SECURITY, UNIFIED, DISTRIBUTION)
        # --------------------------------------------------------------------------
        $groups = @($Records | Where-Object {
            $_.ResourceType -in @('Security Group', 'M365 Group', 'Mail-Enabled Security Group', 'Distribution Group', 'Dynamic Security Group')
        })
        Write-Host "`n[4] ENTRA ID & M365 GROUPS ($($groups.Count))" -ForegroundColor Yellow
        if ($groups.Count -gt 0) {
            $gIdx = 0
            $groupRows = foreach ($item in $groups) {
                $gIdx++
                [PSCustomObject]@{
                    '#'           = $gIdx
                    'Group Name'  = $item.ResourceName
                    'Group Type'  = $item.ResourceType
                    'Role / Access'= $item.AccessLevel
                    'Assignment'  = $item.AssignmentMethod
                    'Sync Source' = $item.SyncSource
                    'Role Assign' = if ($item.IsRoleAssignable) { 'Yes' } else { 'No' }
                }
            }
            $groupRows | Format-Table -AutoSize | Out-Host
        }
        else {
            Write-Host "    [i] No group memberships found." -ForegroundColor DarkGray
        }

        # --------------------------------------------------------------------------
        # [5] EXCHANGE ONLINE MAILBOX PERMISSIONS
        # --------------------------------------------------------------------------
        if ($ExchangeAudited) {
            $exo = @($Records | Where-Object { $_.Workload -eq 'Exchange Online' -and $_.ResourceType -eq 'Mailbox' })
            Write-Host "`n[5] EXCHANGE ONLINE MAILBOX PERMISSIONS ($($exo.Count))" -ForegroundColor DarkYellow
            if ($exo.Count -gt 0) {
                $eIdx = 0
                $exoRows = foreach ($item in $exo) {
                    $eIdx++
                    [PSCustomObject]@{
                        '#'           = $eIdx
                        'Mailbox'     = $item.ResourceName
                        'Permission'  = $item.AccessLevel
                        'Assignment'  = $item.AssignmentMethod
                        'Granted Via' = $item.GrantedViaGroup
                    }
                }
                $exoRows | Format-Table -AutoSize | Out-Host
            }
            else {
                Write-Host "    [i] No delegated mailbox permissions (FullAccess, SendAs, SendOnBehalf) found." -ForegroundColor Green
            }
        }
        else {
            Write-Host "`n[5] EXCHANGE ONLINE MAILBOX PERMISSIONS" -ForegroundColor DarkGray
            Write-Host "    [i] Exchange Online audit skipped (use -IncludeExchangePermissions to audit mailbox delegations)." -ForegroundColor DarkGray
        }

        # --------------------------------------------------------------------------
        # [6] EXECUTIVE METRICS & ARTIFACT DIRECTORY
        # --------------------------------------------------------------------------
        Write-Host "`n================================================================================" -ForegroundColor Cyan
        Write-Host " EXECUTIVE METRICS & ARTIFACT SUMMARY" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan

        $summary = $Records | Group-Object Workload, ResourceType | Select-Object @{
            Name = 'Workload'
            Expression = { $_.Values[0] }
        }, @{
            Name = 'Resource Type'
            Expression = { $_.Values[1] }
        }, @{
            Name = 'Total Records'
            Expression = { $_.Count }
        }

        $summary | Format-Table -AutoSize | Out-Host

        Write-Host "Total Records Audited : $($Records.Count)" -ForegroundColor Green
        if (-not [string]::IsNullOrWhiteSpace($HtmlFilePath)) {
            Write-Host "Interactive HTML      : $HtmlFilePath" -ForegroundColor Cyan
        }
        if (-not [string]::IsNullOrWhiteSpace($CsvFilePath)) {
            Write-Host "CSV Export            : $CsvFilePath" -ForegroundColor Cyan
        }
        Write-Host "Reports Directory     : $ReportsDir" -ForegroundColor Green
        Write-Host "Execution Log Path    : $LogFile" -ForegroundColor Green
        Write-Host "================================================================================`n" -ForegroundColor Cyan
    }

    # --------------------------------------------------------------------------
    # Check for Help invocation, no-parameter execution, and target identity
    # --------------------------------------------------------------------------
    $hasPipelineInput = $MyInvocation.ExpectingInput
    $hasUserParam     = $PSBoundParameters.ContainsKey('UserPrincipalName') -and ($null -ne $UserPrincipalName) -and ($UserPrincipalName.Count -gt 0)

    # 1. Running without parameters or with -Help displays help banner and exits cleanly
    if ($Help -or ($PSBoundParameters.Count -eq 0 -and -not $hasPipelineInput)) {
        Show-ScriptHelp
        exit 0
    }

    # 2. Running with options but omitting -UserPrincipalName and without pipeline input
    if (-not $hasUserParam -and -not $hasPipelineInput) {
        Write-Host "[-] Error: Parameter -UserPrincipalName is required. Specify at least one user to audit (e.g. -UserPrincipalName `"user@domain.com`") or use pipeline input." -ForegroundColor Red
        Write-Host "[*] Run without parameters or with -Help to display the full usage guide.`n" -ForegroundColor Yellow
        exit 1
    }

    Initialize-AuditLogging -TargetDirectory $LogPath
    Write-Log -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)" -Level 'INFO' -Component 'INIT'

    Test-AuditPrerequisite -RequireExchange:$IncludeExchangePermissions

    Connect-AuditService -TargetTenantId $TenantId `
                         -AuditExchange:$IncludeExchangePermissions `
                         -AuditPIM:$IncludePIM

    $Script:AllAuditRecords.Clear()
}

process {
    foreach ($user in $UserPrincipalName) {
        if (-not [string]::IsNullOrWhiteSpace($user)) {
            Invoke-UserAudit -TargetUser $user.Trim() `
                             -IncludeExchange:$IncludeExchangePermissions `
                             -IncludeSharePoint:$IncludeSharePointSiteAudit `
                             -IncludePIMSchedule:$IncludePIM
        }
    }
}

end {
    Write-Log -Message "All user audits completed. Total collected records: $($Script:AllAuditRecords.Count)" -Level 'INFO' -Component 'AUDIT'

    if ($Script:AllAuditRecords.Count -gt 0) {
        Export-AuditReport -Records $Script:AllAuditRecords `
                           -Formats $ExportFormat `
                           -DestinationPath $OutputPath

        Show-AuditSummary -Records $Script:AllAuditRecords `
                          -ReportsDir $OutputPath `
                          -LogFile $Script:LogFilePath `
                          -HtmlFilePath $Script:GeneratedHtmlPath `
                          -CsvFilePath $Script:GeneratedCsvPath `
                          -ExchangeAudited:$IncludeExchangePermissions
    }
    else {
        Write-Log -Message "No audit records were produced during this execution." -Level 'WARN' -Component 'AUDIT'
    }

    if ($PassThru) {
        $Script:AllAuditRecords
    }

    Write-Log -Message "Audit run finished successfully. Elapsed time: $(((Get-Date) - $Script:StartTime).ToString('hh\:mm\:ss'))" -Level 'SUCCESS' -Component 'INIT'
}
