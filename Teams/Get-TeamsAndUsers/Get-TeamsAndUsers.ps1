<#
.SYNOPSIS
    Comprehensive audit and inventory of Microsoft Teams and associated underlying SharePoint Sites.

.DESCRIPTION
    Get-TeamsAndUsers.ps1 is an enterprise-grade PowerShell automation script designed for Microsoft 365,
    Teams, and SharePoint Online administrators. It performs deep tenant-wide or site-specific audits of
    Microsoft Teams, identifying connected SharePoint Online sites, group privacy settings, and user
    permission mappings.

    Key Capabilities:
    1. Audits all Microsoft Teams across the tenant (-All) or targets a specific Team/Site (-Site).
    2. Maps user roles explicitly distinguishing between Owner, Member, and Guest accounts.
    3. Captures User Display Name, UserPrincipalName (UPN), Role, Account Status (Enabled/Disabled), and User Type.
    4. Discovers underlying SharePoint Online site URL, Group ID, and Team privacy settings (Public/Private).
    5. Outputs structured [PSCustomObject] records to the pipeline and renders clean console tables.
    6. Optional CSV reporting (-ExportCsv) to a customizable directory (-ExportPath).
    7. Automated Microsoft Graph session validation and reconnection with required permission scopes.
    8. Robust execution logging to C:\Logs with timestamped entries, severity levels, and automated directory creation.
    9. Dual compatibility across Windows PowerShell 5.1 and PowerShell 7+.
    10. Adheres strictly to Clean Code standards with zero PSScriptAnalyzer warnings.

.PARAMETER All
    Switch parameter. Audits all Microsoft Teams across the tenant.

.PARAMETER Site
    String parameter. Audits a specific Team or Site by Team Display Name, MailNickName,
    Group ID (GUID), or SharePoint site URL.
    Aliases: -Team, -TeamName, -Identity.

.PARAMETER ExportCsv
    Switch parameter. Exports audit results to a CSV file.
    Aliases: -Csv, -Export.

.PARAMETER ExportPath
    String parameter. Destination folder for exported CSV reports.
    Defaults to 'C:\Logs' (or the script directory if C:\Logs is inaccessible).
    Aliases: -OutputPath, -Path.

.PARAMETER Help
    Switch parameter. Displays command syntax, parameter descriptions, prerequisites, and practical examples.
    Aliases: -h.

.EXAMPLE
    # Display built-in help and execution syntax (also displayed when run with no parameters)
    .\Get-TeamsAndUsers.ps1 -Help

.EXAMPLE
    # Audit all Microsoft Teams across the entire tenant
    .\Get-TeamsAndUsers.ps1 -All

.EXAMPLE
    # Audit a specific Team by its display name
    .\Get-TeamsAndUsers.ps1 -Site "Sales Team"

.EXAMPLE
    # Audit a specific Team by its Group GUID and export results to CSV
    .\Get-TeamsAndUsers.ps1 -Site "11111111-2222-3333-4444-555555555555" -ExportCsv

.EXAMPLE
    # Audit a specific Team by its SharePoint Online URL and save CSV to a custom folder
    .\Get-TeamsAndUsers.ps1 -Site "https://contoso.sharepoint.com/sites/Marketing" -ExportCsv -ExportPath "C:\Reports"

.EXAMPLE
    # Query all Teams and pipe results to Out-GridView or downstream filtering
    .\Get-TeamsAndUsers.ps1 -All | Where-Object { $_.Role -eq 'Guest' } | Out-GridView

.NOTES
    Script Name   : Get-TeamsAndUsers.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.0.0
    Last Modified : 2026-10-25
    License       : MIT
    Standards     : Clean Code, Microsoft 365 PowerShell Standards, Zero PSScriptAnalyzer Warnings
#>

[CmdletBinding(SupportsShouldProcess = $false)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [switch]$All,

    [Parameter(Mandatory = $false, Position = 1)]
    [Alias('Team', 'TeamName', 'Identity')]
    [string]$Site,

    [Parameter(Mandatory = $false)]
    [Alias('Csv', 'Export')]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [Alias('OutputPath', 'Path')]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # ==============================================================================
    # SCRIPT METADATA & GLOBAL STATE
    # ==============================================================================
    $Script:Metadata = @{
        Name         = 'Get-TeamsAndUsers.ps1'
        Version      = '1.0.0'
        Author       = 'Roman Pindela'
        Email        = 'roman.pindela@gmail.com'
        GitHub       = 'https://github.com/romanpindela'
        LastModified = '2026-10-25'
        License      = 'MIT'
    }

    $Script:StartTime = Get-Date
    $Script:LogFilePath = $null
    $Script:RequiredScopes = @(
        'Team.ReadBasic.All',
        'Group.Read.All',
        'User.Read.All',
        'Directory.Read.All'
    )
    $Script:AllAuditResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Script:TeamsInventory = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ==============================================================================
    # LOGGING SUBSYSTEM
    # ==============================================================================
    function Initialize-LoggingSystem {
        [CmdletBinding()]
        param()

        $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'LOCALHOST' }
        $userName = if ($env:USERNAME) { $env:USERNAME } else { 'SYSTEM' }
        $logFileName = "${timestamp}-Get-TeamsAndUsers-${computerName}-${userName}.log"

        $targetLogDir = 'C:\Logs'
        $resolvedDir = $targetLogDir

        try {
            if (-not (Test-Path -Path $targetLogDir)) {
                New-Item -Path $targetLogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
        }
        catch {
            # Graceful fallback to %LOCALAPPDATA%\Logs if C:\Logs is inaccessible
            $fallbackDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs"
            if (-not (Test-Path -Path $fallbackDir)) {
                New-Item -Path $fallbackDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
            }
            $resolvedDir = $fallbackDir
        }

        $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName

        $logHeader = @"
================================================================================
Log File Created : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script           : $($Script:Metadata.Name) v$($Script:Metadata.Version)
Author           : $($Script:Metadata.Author) ($($Script:Metadata.Email))
GitHub           : $($Script:Metadata.GitHub)
Host / User      : $computerName / $userName
Target Scope     : $(if ($All) { "All Teams (-All)" } else { "Single Team (-Site '$Site')" })
Log Destination  : $($Script:LogFilePath)
================================================================================
"@
        Add-Content -Path $Script:LogFilePath -Value $logHeader -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    function Write-Log {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '')]
        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$Message,

            [Parameter(Mandatory = $false, Position = 1)]
            [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
            [string]$Level = 'INFO'
        )

        if ($null -eq $Script:LogFilePath) {
            Initialize-LoggingSystem
        }

        $timeString = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $logEntry = "[$timeString] [$Level] $Message"

        # File output
        if ($null -ne $Script:LogFilePath) {
            Add-Content -Path $Script:LogFilePath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
        }

        # Console output with color tagging
        switch ($Level) {
            'INFO'    { Write-Host "[$Level] " -ForegroundColor Cyan -NoNewline; Write-Host $Message -ForegroundColor Gray }
            'SUCCESS' { Write-Host "[$Level] " -ForegroundColor Green -NoNewline; Write-Host $Message -ForegroundColor White }
            'WARN'    { Write-Host "[$Level] " -ForegroundColor Yellow -NoNewline; Write-Host $Message -ForegroundColor Yellow }
            'ERROR'   { Write-Host "[$Level] " -ForegroundColor Red -NoNewline; Write-Host $Message -ForegroundColor Red }
            'DEBUG'   { Write-Host "[$Level] " -ForegroundColor DarkGray -NoNewline; Write-Host $Message -ForegroundColor DarkGray }
        }
    }

    # ==============================================================================
    # SCRIPT HELP DISPLAY
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        Write-Host ""
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "   Get-TeamsAndUsers.ps1 - Microsoft Teams & Users Permission Auditor           " -ForegroundColor Cyan
        Write-Host "   Version: 1.0.0 | Author: Roman Pindela (roman.pindela@gmail.com)             " -ForegroundColor DarkCyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "DESCRIPTION:" -ForegroundColor Yellow
        Write-Host "  Audits Microsoft Teams teams across the tenant or a specific team, mapping"
        Write-Host "  associated SharePoint Online site URLs, privacy settings, and complete user"
        Write-Host "  permission hierarchies (Owners, Members, and External Guests)."
        Write-Host ""
        Write-Host "SYNTAX:" -ForegroundColor Yellow
        Write-Host "  .\Get-TeamsAndUsers.ps1 -All [-ExportCsv] [-ExportPath <String>]"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -Site <String> [-ExportCsv] [-ExportPath <String>]"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -Help"
        Write-Host ""
        Write-Host "PARAMETERS:" -ForegroundColor Yellow
        Write-Host "  -All                 : Audits all Microsoft Teams teams across the entire tenant."
        Write-Host "  -Site <String>       : Audits a specific Team/Site by Team Name, Display Name,"
        Write-Host "                         MailNickName, Group ID (GUID), or SharePoint site URL."
        Write-Host "  -ExportCsv           : Optional switch to export audit results to a CSV report."
        Write-Host "  -ExportPath <String> : Directory path for exported CSV reports (defaults to C:\Logs\)."
        Write-Host "  -Help, -h            : Displays this help screen, parameter details, and examples."
        Write-Host ""
        Write-Host "PREREQUISITES & PERMISSIONS:" -ForegroundColor Yellow
        Write-Host "  1. Microsoft Graph PowerShell SDK (Microsoft.Graph.Authentication, Groups, Teams, Users)."
        Write-Host "     Install: Install-Module Microsoft.Graph -Scope CurrentUser"
        Write-Host "  2. Required Graph API Scopes:"
        Write-Host "     - Team.ReadBasic.All"
        Write-Host "     - Group.Read.All"
        Write-Host "     - User.Read.All"
        Write-Host "     - Directory.Read.All"
        Write-Host "  3. Execution Unblock Command:"
        Write-Host "     Unblock-File -Path .\Get-TeamsAndUsers.ps1"
        Write-Host ""
        Write-Host "EXAMPLES:" -ForegroundColor Yellow
        Write-Host "  # Audit all Teams in tenant and output structured table"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -All" -ForegroundColor Green
        Write-Host ""
        Write-Host "  # Audit a specific team by name and export report to CSV"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -Site ""Sales Team"" -ExportCsv" -ForegroundColor Green
        Write-Host ""
        Write-Host "  # Audit by Group ID GUID and store report in custom path"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -Site ""11111111-2222-3333-4444-555555555555"" -ExportCsv -ExportPath ""C:\Reports""" -ForegroundColor Green
        Write-Host ""
        Write-Host "  # Audit by SharePoint Online site URL"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -Site ""https://contoso.sharepoint.com/sites/Marketing""" -ForegroundColor Green
        Write-Host ""
        Write-Host "  # Pipeline integration: filter only guest users"
        Write-Host "  .\Get-TeamsAndUsers.ps1 -All | Where-Object { `$_.Role -eq 'Guest' } | Out-GridView" -ForegroundColor Green
        Write-Host ""
        Write-Host "================================================================================" -ForegroundColor Cyan
    }

    # ==============================================================================
    # MICROSOFT GRAPH SESSION MANAGEMENT
    # ==============================================================================
    function Connect-GraphSession {
        [CmdletBinding()]
        param()

        Write-Log -Message "Verifying Microsoft Graph PowerShell SDK environment..." -Level 'INFO'

        if (-not (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue)) {
            $err = "Microsoft Graph PowerShell SDK is not installed. Please install it using: Install-Module Microsoft.Graph -Scope CurrentUser"
            Write-Log -Message $err -Level 'ERROR'
            throw $err
        }

        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        $needConnect = $true

        if ($null -ne $mgContext) {
            Write-Log -Message "Existing Microsoft Graph session detected. Account: $($mgContext.Account), TenantId: $($mgContext.TenantId)" -Level 'INFO'

            $currentScopes = @()
            if ($mgContext.Scopes) {
                $currentScopes = @($mgContext.Scopes)
            }

            $missingScopes = @()
            foreach ($scope in $Script:RequiredScopes) {
                if ($currentScopes -notcontains $scope) {
                    $missingScopes += $scope
                }
            }

            if ($missingScopes.Count -eq 0) {
                Write-Log -Message "Existing Graph session contains all required scopes." -Level 'SUCCESS'
                $needConnect = $false
            }
            else {
                Write-Log -Message "Existing Graph session is missing required scope(s): $($missingScopes -join ', '). Reconnecting with required scopes..." -Level 'WARN'
            }
        }

        if ($needConnect) {
            Write-Log -Message "Initiating connection to Microsoft Graph with scopes: $($Script:RequiredScopes -join ', ')..." -Level 'INFO'
            try {
                Connect-MgGraph -Scopes $Script:RequiredScopes -NoWelcome -ErrorAction Stop
                $newContext = Get-MgContext -ErrorAction Stop
                Write-Log -Message "Successfully connected to Microsoft Graph. TenantId: $($newContext.TenantId), Account: $($newContext.Account)" -Level 'SUCCESS'
            }
            catch {
                $err = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
                Write-Log -Message $err -Level 'ERROR'
                throw $err
            }
        }
    }

    # ==============================================================================
    # GRAPH REST REQUEST HELPERS (PAGE-AWARE & BULLETPROOF)
    # ==============================================================================
    function Invoke-GraphPagingRequest {
        [CmdletBinding()]
        [OutputType([System.Object[]])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri
        )

        $items = [System.Collections.Generic.List[PSObject]]::new()
        $nextUri = $Uri

        while (-not [string]::IsNullOrEmpty($nextUri)) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
                if ($null -ne $response -and $response.value) {
                    foreach ($val in $response.value) {
                        $items.Add($val)
                    }
                }
                if ($null -ne $response -and $response.PSObject.Properties['@odata.nextLink'] -and -not [string]::IsNullOrEmpty($response.'@odata.nextLink')) {
                    $nextUri = [string]$response.'@odata.nextLink'
                }
                else {
                    $nextUri = $null
                }
            }
            catch {
                Write-Log -Message "Graph API paging error on '$nextUri': $($_.Exception.Message)" -Level 'ERROR'
                throw
            }
        }

        return $items.ToArray()
    }

    function Invoke-GraphSingleRequest {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri
        )

        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
            return $response
        }
        catch {
            Write-Log -Message "Graph API single request info for '$Uri': $($_.Exception.Message)" -Level 'DEBUG'
            return $null
        }
    }

    # ==============================================================================
    # BUSINESS LOGIC: TEAMS DISCOVERY & INVENTORY
    # ==============================================================================
    function Get-TeamsInventory {
        [CmdletBinding()]
        [OutputType([System.Object[]])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
        param(
            [Parameter(Mandatory = $false)]
            [switch]$AllTeams,

            [Parameter(Mandatory = $false)]
            [string]$TargetSite
        )

        $discoveredTeams = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rawGroups = @()

        if ($AllTeams) {
            Write-Log -Message "Discovering all Microsoft Teams across the tenant..." -Level 'INFO'
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,mailNickname,description,visibility"
            $rawGroups = Invoke-GraphPagingRequest -Uri $uri
        }
        elseif (-not [string]::IsNullOrWhiteSpace($TargetSite)) {
            $trimmedSite = $TargetSite.Trim()
            $isGuid = $trimmedSite -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
            $isUrl = $trimmedSite -match '^https?://'

            if ($isGuid) {
                Write-Log -Message "Searching for Team by Group ID (GUID): $trimmedSite..." -Level 'INFO'
                $uri = "https://graph.microsoft.com/v1.0/groups/$trimmedSite`?`$select=id,displayName,mailNickname,description,visibility,resourceProvisioningOptions"
                $singleGroup = Invoke-GraphSingleRequest -Uri $uri
                if ($null -ne $singleGroup) {
                    $rawGroups = @($singleGroup)
                }
            }
            elseif ($isUrl) {
                Write-Log -Message "Searching for Team by SharePoint Online site URL: $trimmedSite..." -Level 'INFO'
                # Fetch all teams to resolve underlying SharePoint site URLs
                $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,mailNickname,description,visibility"
                $allTenantTeams = Invoke-GraphPagingRequest -Uri $uri

                $cleanTargetUrl = $trimmedSite.TrimEnd('/')
                foreach ($grp in $allTenantTeams) {
                    $spoSiteObj = Invoke-GraphSingleRequest -Uri "https://graph.microsoft.com/v1.0/groups/$($grp.id)/sites/root?`$select=id,webUrl"
                    if ($null -ne $spoSiteObj -and $spoSiteObj.webUrl) {
                        $cleanSpoUrl = [string]$spoSiteObj.webUrl.TrimEnd('/')
                        if ($cleanSpoUrl -eq $cleanTargetUrl) {
                            $rawGroups = @($grp)
                            break
                        }
                    }
                }

                if ($rawGroups.Count -eq 0) {
                    # Fallback match by site path segment (slug)
                    $urlSegments = $cleanTargetUrl -split '/'
                    $slug = $urlSegments[-1]
                    Write-Log -Message "Direct URL match not found. Attempting match by URL slug: '$slug'..." -Level 'DEBUG'
                    $matched = $allTenantTeams | Where-Object {
                        ($_.mailNickname -and $_.mailNickname -eq $slug) -or
                        ($_.displayName -and $_.displayName -eq $slug)
                    }
                    if ($matched) {
                        $rawGroups = @($matched)
                    }
                }
            }
            else {
                Write-Log -Message "Searching for Team by Display Name or MailNickName: '$trimmedSite'..." -Level 'INFO'
                $escapedQuery = $trimmedSite.Replace("'", "''")
                $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and (displayName eq '$escapedQuery' or mailNickname eq '$escapedQuery')&`$select=id,displayName,mailNickname,description,visibility"
                $rawGroups = Invoke-GraphPagingRequest -Uri $uri

                if ($rawGroups.Count -eq 0) {
                    # Fallback to loose case-insensitive matching across all teams
                    Write-Log -Message "Exact match not found. Searching all tenant teams for substring: '$trimmedSite'..." -Level 'DEBUG'
                    $uriAll = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,mailNickname,description,visibility"
                    $allTenantTeams = Invoke-GraphPagingRequest -Uri $uriAll
                    $matched = $allTenantTeams | Where-Object {
                        ($_.displayName -and $_.displayName -like "*$trimmedSite*") -or
                        ($_.mailNickname -and $_.mailNickname -like "*$trimmedSite*")
                    }
                    if ($matched) {
                        $rawGroups = @($matched)
                    }
                }
            }
        }

        foreach ($grp in $rawGroups) {
            if ($null -eq $grp -or [string]::IsNullOrEmpty($grp.id)) { continue }

            $grpId = [string]$grp.id
            $displayName = if ($grp.displayName) { [string]$grp.displayName } else { 'Unnamed Team' }
            $mailNickName = if ($grp.mailNickname) { [string]$grp.mailNickname } else { 'N/A' }
            $visibility = if ($grp.visibility) { [string]$grp.visibility } else { 'Private' }
            $description = if ($grp.description) { [string]$grp.description } else { '' }

            # Retrieve underlying SharePoint site URL
            $spoSiteObj = Invoke-GraphSingleRequest -Uri "https://graph.microsoft.com/v1.0/groups/$grpId/sites/root?`$select=id,webUrl,displayName"
            $siteUrl = 'Not Provisioned'
            if ($null -ne $spoSiteObj -and $spoSiteObj.webUrl) {
                $siteUrl = [string]$spoSiteObj.webUrl
            }

            $teamRecord = [PSCustomObject]@{
                GroupId      = $grpId
                TeamName     = $displayName
                MailNickName = $mailNickName
                Privacy      = $visibility
                SiteUrl      = $siteUrl
                Description  = $description
            }

            $discoveredTeams.Add($teamRecord)
        }

        return $discoveredTeams.ToArray()
    }

    # ==============================================================================
    # BUSINESS LOGIC: USER ROLES & MEMBERSHIP EXTRACTION
    # ==============================================================================
    function Get-TeamUserPermissions {
        [CmdletBinding()]
        [OutputType([System.Object[]])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '')]
        param(
            [Parameter(Mandatory = $true)]
            [PSCustomObject]$TeamInfo
        )

        $userAuditList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $processedUserIds = @{}

        $grpId = $TeamInfo.GroupId
        $teamName = $TeamInfo.TeamName
        $siteUrl = $TeamInfo.SiteUrl
        $privacy = $TeamInfo.Privacy

        # Step 1: Query Owners
        $ownersUri = "https://graph.microsoft.com/v1.0/groups/$grpId/owners?`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled"
        $owners = Invoke-GraphPagingRequest -Uri $ownersUri

        foreach ($owner in $owners) {
            if ($null -eq $owner -or [string]::IsNullOrEmpty($owner.id)) { continue }

            $userId = [string]$owner.id
            $upn = if ($owner.userPrincipalName) { [string]$owner.userPrincipalName } elseif ($owner.mail) { [string]$owner.mail } else { $userId }
            $dispName = if ($owner.displayName) { [string]$owner.displayName } else { $upn }

            $isGuest = ($owner.userType -and $owner.userType -eq 'Guest') -or ($upn -like '*#EXT#*')
            $userTypeStr = if ($isGuest) { 'Guest' } else { 'Member' }

            $accStatus = 'Unknown'
            if ($null -ne $owner.accountEnabled) {
                if ($owner.accountEnabled -eq $true) { $accStatus = 'Enabled' }
                elseif ($owner.accountEnabled -eq $false) { $accStatus = 'Disabled' }
            }

            $record = [PSCustomObject]@{
                TeamName          = $teamName
                GroupId           = $grpId
                SiteUrl           = $siteUrl
                Privacy           = $privacy
                DisplayName       = $dispName
                UserPrincipalName = $upn
                Role              = 'Owner'
                UserType          = $userTypeStr
                AccountStatus     = $accStatus
            }

            $userAuditList.Add($record)
            $processedUserIds[$userId] = $true
        }

        # Step 2: Query Members
        $membersUri = "https://graph.microsoft.com/v1.0/groups/$grpId/members?`$select=id,displayName,userPrincipalName,mail,userType,accountEnabled"
        $members = Invoke-GraphPagingRequest -Uri $membersUri

        foreach ($member in $members) {
            if ($null -eq $member -or [string]::IsNullOrEmpty($member.id)) { continue }

            $userId = [string]$member.id
            if ($processedUserIds.ContainsKey($userId)) {
                # Already registered as Owner; avoid duplicate entry
                continue
            }

            $upn = if ($member.userPrincipalName) { [string]$member.userPrincipalName } elseif ($member.mail) { [string]$member.mail } else { $userId }
            $dispName = if ($member.displayName) { [string]$member.displayName } else { $upn }

            $isGuest = ($member.userType -and $member.userType -eq 'Guest') -or ($upn -like '*#EXT#*')
            $userTypeStr = if ($isGuest) { 'Guest' } else { 'Member' }
            $roleStr = if ($isGuest) { 'Guest' } else { 'Member' }

            $accStatus = 'Unknown'
            if ($null -ne $member.accountEnabled) {
                if ($member.accountEnabled -eq $true) { $accStatus = 'Enabled' }
                elseif ($member.accountEnabled -eq $false) { $accStatus = 'Disabled' }
            }

            $record = [PSCustomObject]@{
                TeamName          = $teamName
                GroupId           = $grpId
                SiteUrl           = $siteUrl
                Privacy           = $privacy
                DisplayName       = $dispName
                UserPrincipalName = $upn
                Role              = $roleStr
                UserType          = $userTypeStr
                AccountStatus     = $accStatus
            }

            $userAuditList.Add($record)
            $processedUserIds[$userId] = $true
        }

        return $userAuditList.ToArray()
    }

    # ==============================================================================
    # REPORTING & EXPORT UTILITIES
    # ==============================================================================
    function Export-TeamsReport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[PSCustomObject]]$Records,

            [Parameter(Mandatory = $false)]
            [string]$DestinationPath
        )

        $targetFolder = $DestinationPath
        if ([string]::IsNullOrWhiteSpace($targetFolder)) {
            $defaultLogDir = 'C:\Logs'
            if (Test-Path -Path $defaultLogDir) {
                $targetFolder = $defaultLogDir
            }
            else {
                $targetFolder = $PSScriptRoot
            }
        }

        try {
            if (-not (Test-Path -Path $targetFolder)) {
                New-Item -Path $targetFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
        }
        catch {
            Write-Log -Message "Could not create target export directory '$targetFolder'. Defaulting to script folder." -Level 'WARN'
            $targetFolder = $PSScriptRoot
        }

        $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
        $csvFileName = "Get-TeamsAndUsers-Audit-${timestamp}.csv"
        $resolvedCsvPath = Join-Path -Path $targetFolder -ChildPath $csvFileName

        try {
            $Records | Export-Csv -Path $resolvedCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Log -Message "Audit report successfully exported to CSV: $resolvedCsvPath" -Level 'SUCCESS'
        }
        catch {
            Write-Log -Message "Failed to export audit report to CSV: $($_.Exception.Message)" -Level 'ERROR'
        }
    }

    # ==============================================================================
    # INITIALIZATION & EXECUTION PRE-FLIGHT
    # ==============================================================================
    # Guard: No parameters or -Help terminates with code 0
    if ($Help.IsPresent -or ($PSBoundParameters.Count -eq 0)) {
        Show-ScriptHelp
        exit 0
    }

    Initialize-LoggingSystem

    Write-Log -Message "Starting execution of $($Script:Metadata.Name) v$($Script:Metadata.Version)..." -Level 'INFO'

    # Guard: Either -All or -Site must be supplied
    if (-not $All -and [string]::IsNullOrWhiteSpace($Site)) {
        Write-Log -Message "Missing required operational parameter. Specify either -All or -Site <String>." -Level 'ERROR'
        Show-ScriptHelp
        exit 1
    }

    if ($All -and -not [string]::IsNullOrWhiteSpace($Site)) {
        Write-Log -Message "Both -All and -Site were supplied. Prioritizing specific Team/Site: '$Site'." -Level 'WARN'
        $All = $false
    }

    # Authenticate and establish Microsoft Graph connection
    Connect-GraphSession
}

process {
    try {
        # Discover Teams matching parameters
        $discovered = Get-TeamsInventory -AllTeams:$All -TargetSite:$Site
        foreach ($item in $discovered) {
            $Script:TeamsInventory.Add($item)
        }

        if ($Script:TeamsInventory.Count -eq 0) {
            $searchTarget = if ($All) { 'All Tenant Teams' } else { $Site }
            Write-Log -Message "No matching Microsoft Teams found in the tenant for query: '$searchTarget'." -Level 'WARN'
            return
        }

        Write-Log -Message "Discovered $($Script:TeamsInventory.Count) Microsoft Team(s) to audit." -Level 'INFO'

        $index = 0
        foreach ($team in $Script:TeamsInventory) {
            $index++
            Write-Log -Message "[$index/$($Script:TeamsInventory.Count)] Auditing Team: '$($team.TeamName)' (Group ID: $($team.GroupId))..." -Level 'INFO'

            $teamMembers = Get-TeamUserPermissions -TeamInfo $team
            Write-Log -Message "Found $($teamMembers.Count) total user permission mapping(s) for Team '$($team.TeamName)'." -Level 'DEBUG'

            foreach ($record in $teamMembers) {
                $Script:AllAuditResults.Add($record)
            }
        }
    }
    catch {
        Write-Log -Message "Execution failed during audit processing: $($_.Exception.Message)" -Level 'ERROR'
        throw
    }
}

end {
    try {
        $totalTeams = $Script:TeamsInventory.Count
        $totalUsers = $Script:AllAuditResults.Count
        $totalOwners = ($Script:AllAuditResults | Where-Object { $_.Role -eq 'Owner' }).Count
        $totalMembers = ($Script:AllAuditResults | Where-Object { $_.Role -eq 'Member' }).Count
        $totalGuests = ($Script:AllAuditResults | Where-Object { $_.Role -eq 'Guest' }).Count
        $totalDisabled = ($Script:AllAuditResults | Where-Object { $_.AccountStatus -eq 'Disabled' }).Count

        if ($Script:AllAuditResults.Count -gt 0) {
            $tableColumns = @(
                @{ Label = 'Team Name'; Expression = { $_.TeamName }; Width = 22 },
                @{ Label = 'Display Name'; Expression = { $_.DisplayName }; Width = 22 },
                @{ Label = 'User Principal Name'; Expression = { $_.UserPrincipalName }; Width = 32 },
                @{ Label = 'Role'; Expression = { $_.Role }; Width = 10 },
                @{ Label = 'User Type'; Expression = { $_.UserType }; Width = 10 },
                @{ Label = 'Status'; Expression = { $_.AccountStatus }; Width = 10 },
                @{ Label = 'Privacy'; Expression = { $_.Privacy }; Width = 10 },
                @{ Label = 'SharePoint Site URL'; Expression = { $_.SiteUrl } }
            )

            Write-Host ""
            $Script:AllAuditResults | Format-Table -Property $tableColumns -AutoSize | Out-Host
        }

        # Executive Console Summary Card
        Write-Host ""
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "                        AUDIT EXECUTION SUMMARY                                 " -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  Total Teams Audited        : $totalTeams" -ForegroundColor White
        Write-Host "  Total User Mappings        : $totalUsers" -ForegroundColor White
        Write-Host "    - Owners                 : $totalOwners" -ForegroundColor Green
        Write-Host "    - Members                : $totalMembers" -ForegroundColor Yellow
        Write-Host "    - Guests                 : $totalGuests" -ForegroundColor Magenta
        Write-Host "  Disabled Accounts Detected : $totalDisabled" -ForegroundColor $(if ($totalDisabled -gt 0) { 'Red' } else { 'DarkGray' })
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host ""

        # Export CSV if requested
        if ($ExportCsv) {
            Export-TeamsReport -Records $Script:AllAuditResults -DestinationPath $ExportPath
        }

        $elapsed = (Get-Date) - $Script:StartTime
        $elapsedFormatted = "{0:mm\:ss\.fff}" -f $elapsed
        Write-Log -Message "Audit completed successfully in $elapsedFormatted. Total records: $totalUsers." -Level 'SUCCESS'

        # Emit structured objects to the PowerShell pipeline for downstream cmdlets
        Write-Output -InputObject $Script:AllAuditResults
    }
    catch {
        Write-Log -Message "Error during report generation / finalization: $($_.Exception.Message)" -Level 'ERROR'
        throw
    }
}
