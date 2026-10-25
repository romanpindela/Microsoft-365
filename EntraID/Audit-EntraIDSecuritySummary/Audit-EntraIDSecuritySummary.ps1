<#
.SYNOPSIS
    Performs an enterprise-grade, read-only security configuration audit of Microsoft Entra ID (Azure AD),
    evaluating settings against CIS Microsoft 365 Foundations Benchmark and Microsoft Best Practices,
    and generating an interactive HTML executive report.

.DESCRIPTION
    Audit-EntraIDSecuritySummary.ps1 inspects key tenant security configurations via Microsoft Graph API
    in a strictly read-only mode without modifying any tenant resources.

    Scope of Audited Areas:
    1. User & Application Registration Settings:
       - User application registration permissions (Users can register applications).
       - User security group and Microsoft 365 group creation permissions.
       - Restrict access to Microsoft Entra admin center for non-administrators.
       - User consent settings for third-party applications (Consent Phishing protection).
    2. Guest Access Settings:
       - Guest user access restriction level (GuestUserRoleId: Restricted vs Limited vs Default).
       - Guest invitation restrictions (restricted to administrators and guest inviters).
    3. Authentication & Methods:
       - Security Defaults enforcement status (for tenants without P1/P2 licenses).
       - Self-Service Password Reset (SSPR) configuration and authentication methods.
       - Multi-Factor Authentication (MFA) registration for privileged role holders.
    4. Conditional Access Policies (Entra ID P1/P2):
       - Active policy enforcing MFA for all administrative directory roles.
       - Active policy blocking legacy authentication protocols (POP, IMAP, SMTP, ActiveSync).
       - Active policy enforcing MFA or block for risky sign-ins (Entra ID P2 Identity Protection).
    5. Administrative Roles & Emergency Accounts:
       - Global Administrator account count evaluation (Recommended: 2 to 4 accounts).
       - Discovery of unprotected Global Administrator accounts lacking MFA registration.

    Key Features:
    - 100% Read-Only Safety: Performs zero update, set, or write operations on Entra ID objects.
    - Intelligent License & Permission Handling: Gracefully detects absence of P1/P2 licenses or scopes.
    - Automatic Scope Verification: Validates that active Graph session holds required read permissions.
    - Executive HTML Dashboard: Generates a self-contained, interactive HTML5 report with circular score gauge,
      dynamic search, category filters, collapsible recommendation cards, action matrix, and print styling.
    - Enterprise Logging: Centralized dual logging to console and timestamped file with level tagging.
    - Full Cross-Platform & Edition Support: Compatible with PowerShell 7.2+ and Windows PowerShell 5.1.

.PARAMETER OutputPath
    Target folder where generated HTML and JSON reports are stored.
    Default: "C:\Logs" (Alias: -ReportPath).

.PARAMETER OpenReport
    Switch parameter. Automatically opens the generated HTML audit report in the default web browser upon completion.

.PARAMETER TenantId
    Optional Entra ID Tenant ID (GUID or verified domain name) to target during authentication context setup.

.PARAMETER ExportJson
    Switch parameter. Exports the structured audit results into a companion JSON file alongside the HTML report.

.PARAMETER PassThru
    Switch parameter. Emits custom PowerShell objects ([PSCustomObject]) representing all audit findings to the pipeline.

.PARAMETER Run
    Switch parameter. Explicitly initiates the audit execution without interactive confirmation prompts.

.PARAMETER ForceReconnect
    Switch parameter. Forces a fresh interactive authentication prompt to request updated permissions/scopes.

.PARAMETER Help
    Switch parameter. Displays built-in help, usage examples, parameter descriptions, and exits.
    Alias: -h.

.EXAMPLE
    .\Audit-EntraIDSecuritySummary.ps1 -Run
    Executes a comprehensive Entra ID security audit and saves the HTML report in C:\Logs\.

.EXAMPLE
    .\Audit-EntraIDSecuritySummary.ps1 -OpenReport
    Performs the full audit and automatically launches the resulting interactive HTML dashboard in the browser.

.EXAMPLE
    .\Audit-EntraIDSecuritySummary.ps1 -ForceReconnect -OpenReport
    Forces a new sign-in with full required Graph scopes and generates the HTML report in C:\Logs\.

.EXAMPLE
    .\Audit-EntraIDSecuritySummary.ps1 -TenantId "72f988bf-86f1-41af-91ab-2d7cd011db47" -OutputPath "C:\Audits\EntraID" -ExportJson
    Audits a specific tenant ID, exporting both HTML and JSON reports into C:\Audits\EntraID\.

.EXAMPLE
    .\Audit-EntraIDSecuritySummary.ps1 -Run -PassThru | Where-Object { $_.Status -eq 'Non-Compliant' } | Format-Table -AutoSize
    Runs the audit and filters non-compliant security rules directly in the PowerShell console pipeline.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.1
    Standards: Adheres to CleanCode, CIS M365 Foundations Benchmark, and Microsoft 365 PowerShell Standards
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
param(
    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [Alias('ReportPath')]
    [string]$OutputPath = "C:\Logs",

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$OpenReport,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [string]$TenantId,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$ExportJson,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$PassThru,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$Run,

    [Parameter(Mandatory = $false, ParameterSetName = 'Default')]
    [switch]$ForceReconnect,

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
        Title         = "Audit-EntraIDSecuritySummary.ps1"
        Version       = "1.0.1"
        Author        = "Roman Pindela"
        Email         = "roman.pindela@gmail.com"
        GitHub        = "https://github.com/romanpindela"
        Specification = "specification_Audit-EntraIDSecuritySummary.md"
        Date          = "2026-10-25"
    }

    $Script:StartTime         = Get-Date
    $Script:LogFilePath       = $null
    $Script:ActiveTenantId    = $null
    $Script:ActiveTenantName  = $null
    $Script:GeneratedHtmlPath = $null
    $Script:GeneratedJsonPath = $null
    $Script:AuditFindings     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Script:PrivilegedUsers   = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ==============================================================================
    # HELPER: SAFE COLLECTION COUNT (Set-StrictMode & PowerShell 5.1 Safe)
    # ==============================================================================
    function Get-SafeCount {
        param($TargetCollection)
        if ($null -eq $TargetCollection) { return 0 }
        if ($TargetCollection -is [System.Collections.ICollection]) {
            return $TargetCollection.Count
        }
        return @($TargetCollection).Count
    }

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
    Comprehensive, read-only Microsoft Entra ID (Azure AD) security configuration
    audit script aligning with CIS Microsoft 365 Foundations Benchmark standards.
    Inspects user & app settings, guest access, authentication methods (SSPR & MFA),
    Conditional Access policies, and Global Administrator emergency break-glass hygiene.
    Outputs an interactive, standalone HTML executive dashboard and optional JSON.

PARAMETERS:
    -OutputPath     Target directory where HTML and JSON reports are saved.
                    Default: C:\Logs (Alias: -ReportPath)
    -OpenReport     Automatically opens the HTML report in default browser after audit.
    -TenantId       Optional specific Entra ID Tenant ID (GUID or domain).
    -ExportJson     Exports companion JSON structured data file.
    -PassThru       Outputs audit finding objects to the PowerShell pipeline.
    -Run            Runs the audit directly without interactive prompt.
    -ForceReconnect Prompts fresh authentication to request full required Graph scopes.
    -Help, -h       Displays this help banner and exits.

USAGE EXAMPLES:
    # 1. Standard execution with automatic report display:
    .\$($Script:Metadata.Title) -OpenReport

    # 2. Re-authenticate with elevated scopes and run:
    .\$($Script:Metadata.Title) -ForceReconnect -OpenReport

    # 3. Targeted tenant audit with custom output path and JSON export:
    .\$($Script:Metadata.Title) -TenantId "contoso.onmicrosoft.com" -OutputPath "D:\AuditLogs" -ExportJson

    # 4. Pipeline inspection of non-compliant findings:
    .\$($Script:Metadata.Title) -Run -PassThru | Where-Object { `$_.Status -eq 'Non-Compliant' } | Format-Table

SECURITY & UNBLOCKING:
    If downloaded from GitHub, unblock the script prior to execution:
    Unblock-File -Path .\$($Script:Metadata.Title)
================================================================================
"@
        Write-Host $banner -ForegroundColor Cyan
    }

    # ==============================================================================
    # LOGGING SYSTEM
    # ==============================================================================
    function Initialize-AuditLogging {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$TargetDirectory = "C:\Logs"
        )

        $resolvedDir = $TargetDirectory
        try {
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Audit-EntraIDSecuritySummary"
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
        }

        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath "Audit-EntraIDSecuritySummary_$timestamp.log"

        $headerText = @"
================================================================================
Log Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Script     : $($Script:Metadata.Title) v$($Script:Metadata.Version)
Author     : $($Script:Metadata.Author)
Host       : $env:COMPUTERNAME ($env:USERNAME)
PowerShell : $($PSVersionTable.PSVersion.ToString())
================================================================================
"@
        try {
            Set-Content -Path $Script:LogFilePath -Value $headerText -Encoding UTF8 -Force
        }
        catch {
            $null = $_
        }
    }

    function Write-AuditLog {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')]
            [string]$Level = 'INFO',

            [Parameter(Mandatory = $false)]
            [string]$Component = 'GENERAL'
        )

        $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $logLine = "[$timeStr] [$($Level.PadRight(7))] [$($Component.PadRight(10))] $Message"

        if ($Script:LogFilePath -and (Test-Path -Path (Split-Path -Path $Script:LogFilePath -Parent))) {
            try {
                Add-Content -Path $Script:LogFilePath -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            catch {
                $null = $_
            }
        }

        $color = switch ($Level) {
            'INFO'    { 'Gray' }
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

    # ==============================================================================
    # INPUT VALIDATION & SECURITY PROTECTION
    # ==============================================================================
    function Test-InputSecurity {
        [CmdletBinding()]
        param()

        $forbiddenCharsRegex = '[\$`;&|><\r\n]'

        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $guidPattern   = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
            $domainPattern = '^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            $cleanTenant   = $TenantId.Trim()

            if ($cleanTenant -match $forbiddenCharsRegex) {
                Write-Error "Invalid characters detected in -TenantId parameter."
                exit 1
            }

            if (($cleanTenant -notmatch $guidPattern) -and ($cleanTenant -notmatch $domainPattern)) {
                Write-Error "Invalid format for -TenantId. Must be a valid GUID or domain name (e.g., contoso.onmicrosoft.com)."
                exit 1
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
            if ($OutputPath -match '[\*?"<>|]') {
                Write-Error "Invalid characters detected in -OutputPath parameter."
                exit 1
            }
        }
    }

    # ==============================================================================
    # GRAPH CONNECTION & SCOPE VERIFICATION
    # ==============================================================================
    function Connect-EntraAuditEnvironment {
        [CmdletBinding()]
        param()

        Write-AuditLog -Message "Checking Microsoft Graph connection and required permissions..." -Level 'INFO' -Component 'AUTH'

        $requiredScopes = @(
            'Policy.Read.All',
            'Directory.Read.All',
            'RoleManagement.Read.Directory',
            'Policy.Read.ConditionalAccess',
            'Reports.Read.All'
        )

        $currentContext = $null
        try {
            $currentContext = Get-MgContext -ErrorAction SilentlyContinue
        }
        catch {
            $null = $_
        }

        $needsConnection = $true
        if ($currentContext -and $currentContext.Account -and (-not $ForceReconnect)) {
            Write-AuditLog -Message "Active Graph session detected for account: $($currentContext.Account)" -Level 'INFO' -Component 'AUTH'

            $tenantMatches = $true
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                if ($currentContext.TenantId -ne $TenantId) {
                    Write-AuditLog -Message "Current context tenant ($($currentContext.TenantId)) does not match requested tenant ($TenantId)." -Level 'WARN' -Component 'AUTH'
                    $tenantMatches = $false
                }
            }

            # Check if all required scopes are present in the active session
            $activeScopes = @()
            if ($currentContext.Scopes) {
                $activeScopes = @($currentContext.Scopes)
            }

            $missingScopes = @()
            foreach ($reqScope in $requiredScopes) {
                $matched = $false
                foreach ($actScope in $activeScopes) {
                    if ($actScope -ieq $reqScope) {
                        $matched = $true
                        break
                    }
                }
                if (-not $matched) {
                    $missingScopes += $reqScope
                }
            }

            if ($tenantMatches -and ((Get-SafeCount $missingScopes) -eq 0)) {
                Write-AuditLog -Message "Active Graph session has all required scopes." -Level 'SUCCESS' -Component 'AUTH'
                $needsConnection = $false
            } else {
                if ((Get-SafeCount $missingScopes) -gt 0) {
                    Write-AuditLog -Message "Active session lacks required scopes ($($missingScopes -join ', ')). Requesting elevated scope consent..." -Level 'WARN' -Component 'AUTH'
                }
                $needsConnection = $true
            }
        }

        if ($needsConnection) {
            Write-AuditLog -Message "Connecting to Microsoft Graph with required read-only scopes ($($requiredScopes -join ', '))..." -Level 'INFO' -Component 'AUTH'
            $connectParams = @{
                Scopes      = $requiredScopes
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
                $connectParams['TenantId'] = $TenantId
            }

            try {
                Connect-MgGraph @connectParams
                $currentContext = Get-MgContext
                Write-AuditLog -Message "Successfully connected to Microsoft Graph. Tenant ID: $($currentContext.TenantId)" -Level 'SUCCESS' -Component 'AUTH'
            }
            catch {
                Write-AuditLog -Message "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level 'ERROR' -Component 'AUTH'
                throw $_
            }
        }

        $Script:ActiveTenantId = $currentContext.TenantId
    }

    # ==============================================================================
    # RESILIENT GRAPH QUERY ENGINE
    # ==============================================================================
    function Invoke-EntraGraphQuery {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Uri,

            [Parameter(Mandatory = $false)]
            [string]$Method = 'GET',

            [Parameter(Mandatory = $false)]
            [switch]$GetAllPages
        )

        $results = [System.Collections.Generic.List[object]]::new()
        $currentUri = $Uri

        try {
            do {
                $response = Invoke-MgGraphRequest -Uri $currentUri -Method $Method -ErrorAction Stop
                if ($null -eq $response) { break }

                $hasValue = $false
                $valueObj = $null

                if ($response -is [System.Collections.IDictionary] -and $response.Contains('value')) {
                    $hasValue = $true
                    $valueObj = $response['value']
                }
                elseif ($response.PSObject.Properties['value']) {
                    $hasValue = $true
                    $valueObj = $response.value
                }

                if ($hasValue) {
                    if ($valueObj -is [System.Collections.IEnumerable] -and $valueObj -isnot [string]) {
                        foreach ($item in $valueObj) {
                            $results.Add($item)
                        }
                    } else {
                        $results.Add($valueObj)
                    }
                } else {
                    return $response
                }

                $nextLink = $null
                if ($response -is [System.Collections.IDictionary] -and $response.Contains('@odata.nextLink')) {
                    $nextLink = $response['@odata.nextLink']
                }
                elseif ($response.PSObject.Properties['@odata.nextLink']) {
                    $nextLink = $response.'@odata.nextLink'
                }

                if ($GetAllPages -and (-not [string]::IsNullOrWhiteSpace($nextLink))) {
                    $currentUri = $nextLink
                } else {
                    $currentUri = $null
                }
            } while ($currentUri)

            return $results.ToArray()
        }
        catch {
            $msg = $_.Exception.Message
            if ($msg -match 'Forbidden') {
                Write-AuditLog -Message "Graph query to '$Uri' failed: 403 Forbidden (Insufficient permissions or missing admin consent)." -Level 'WARN' -Component 'GRAPH'
            } else {
                Write-AuditLog -Message "Graph query to '$Uri' failed: $msg" -Level 'WARN' -Component 'GRAPH'
            }
            return $null
        }
    }

    function Get-ObjectProperty {
        param($targetObject, [string]$propertyName)
        if ($null -eq $targetObject) { return $null }
        if ($targetObject -is [System.Collections.IDictionary] -and $targetObject.Contains($propertyName)) {
            return $targetObject[$propertyName]
        }
        if ($targetObject.PSObject.Properties[$propertyName]) {
            return $targetObject.$propertyName
        }
        return $null
    }
}

process {
    # Check if run without parameters or with help
    $hasPipelineInput = $MyInvocation.ExpectingInput
    if ($Help) {
        Show-ScriptHelp
        exit 0
    }

    if ($PSBoundParameters.Count -eq 0 -and -not $hasPipelineInput) {
        Show-ScriptHelp
        Write-Host ""
        Write-Host "No execution parameters specified." -ForegroundColor Yellow
        Write-Host "Do you want to run the full Entra ID Security Audit with default settings now? [Y/N]: " -NoNewline -ForegroundColor Cyan
        $userChoice = [Console]::ReadLine()
        if ($userChoice -notmatch '^(y|yes)$') {
            Write-Host "Audit cancelled by user." -ForegroundColor Gray
            exit 0
        }
    }

    Test-InputSecurity
    Initialize-AuditLogging -TargetDirectory $OutputPath

    Write-AuditLog -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version) - Pure Read-Only Mode" -Level 'INFO' -Component 'INIT'

    Connect-EntraAuditEnvironment

    # Resolve Tenant Metadata
    try {
        $orgDetails = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/organization"
        if ($orgDetails -and ((Get-SafeCount $orgDetails) -gt 0)) {
            $Script:ActiveTenantName = Get-ObjectProperty -targetObject $orgDetails[0] -propertyName 'displayName'
        }
    }
    catch {
        $null = $_
    }

    if ([string]::IsNullOrWhiteSpace($Script:ActiveTenantName)) {
        $Script:ActiveTenantName = $Script:ActiveTenantId
    }

    Write-AuditLog -Message "Target Tenant: '$($Script:ActiveTenantName)' (ID: $($Script:ActiveTenantId))" -Level 'INFO' -Component 'INIT'

    # ==============================================================================
    # 1. AUDIT: USER AND APP REGISTRATION SETTINGS
    # ==============================================================================
    Write-AuditLog -Message "Auditing User & Application Registration Settings..." -Level 'INFO' -Component 'APP_SETTINGS'

    $authPolicy = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"
    $defaultUserPerms = if ($authPolicy) { Get-ObjectProperty -targetObject $authPolicy -propertyName 'defaultUserRolePermissions' } else { $null }

    # Rule 1: Users can register applications
    if ($null -eq $authPolicy) {
        $r1Status = 'Warning'
        $r1Current = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $allowedToCreateApps = if ($defaultUserPerms) { Get-ObjectProperty -targetObject $defaultUserPerms -propertyName 'allowedToCreateApps' } else { $null }
        $r1Status = if ($allowedToCreateApps -eq $false) { 'Compliant' } else { 'Non-Compliant' }
        $r1Current = if ($null -eq $allowedToCreateApps) { "Unknown / Not Configured" } else { "allowedToCreateApps = $allowedToCreateApps" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-01"
        Category        = "Users & App Registration"
        Title           = "Users Can Register Applications"
        CurrentValue    = $r1Current
        Recommended     = "allowedToCreateApps = False (Disabled)"
        Status          = $r1Status
        Severity        = "High"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.1"
        Impact          = "Allowing regular users to register custom applications creates rogue OAuth entry points, enables illicit consent attacks, and bypasses admin oversight for enterprise data access."
        RemediationPath = "Entra admin center > Identity > Users > User settings > 'Users can register applications' = No"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ AllowedToCreateApps = `$false }"
    })

    # Rule 2: Users can create security groups
    if ($null -eq $authPolicy) {
        $r2Status = 'Warning'
        $r2Current = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $allowedToCreateSecurityGroups = if ($defaultUserPerms) { Get-ObjectProperty -targetObject $defaultUserPerms -propertyName 'allowedToCreateSecurityGroups' } else { $null }
        $r2Status = if ($allowedToCreateSecurityGroups -eq $false) { 'Compliant' } else { 'Non-Compliant' }
        $r2Current = if ($null -eq $allowedToCreateSecurityGroups) { "Unknown / Not Configured" } else { "allowedToCreateSecurityGroups = $allowedToCreateSecurityGroups" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-02"
        Category        = "Users & App Registration"
        Title           = "Users Can Create Security Groups"
        CurrentValue    = $r2Current
        Recommended     = "allowedToCreateSecurityGroups = False (Disabled)"
        Status          = $r2Status
        Severity        = "Medium"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.2"
        Impact          = "Unrestricted security group creation leads to directory sprawl, shadow group management, and uncontrolled access granting without compliance approvals."
        RemediationPath = "Entra admin center > Groups > Group settings > 'Users can create security groups in Azure portals, API or PowerShell' = No"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ AllowedToCreateSecurityGroups = `$false }"
    })

    # Rule 3: Microsoft 365 Group Creation Governance
    Write-AuditLog -Message "Checking Microsoft 365 group creation settings..." -Level 'INFO' -Component 'GROUPS'
    $groupSettings = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/groupSettings"
    $m365GroupSettingFound = $false
    $m365GroupCreationRestricted = $false

    if ($groupSettings) {
        foreach ($gs in $groupSettings) {
            $displayName = Get-ObjectProperty -targetObject $gs -propertyName 'displayName'
            if ($displayName -eq 'Group.Unified') {
                $m365GroupSettingFound = $true
                $values = Get-ObjectProperty -targetObject $gs -propertyName 'values'
                if ($values) {
                    foreach ($v in $values) {
                        $vName = Get-ObjectProperty -targetObject $v -propertyName 'name'
                        $vVal  = Get-ObjectProperty -targetObject $v -propertyName 'value'
                        if ($vName -eq 'EnableGroupCreation' -and $vVal -eq 'false') {
                            $m365GroupCreationRestricted = $true
                        }
                    }
                }
            }
        }
    }

    if (-not $m365GroupSettingFound) {
        Write-AuditLog -Message "No custom Group.Unified template found. M365 group creation defaults to enabled for all users." -Level 'DEBUG' -Component 'GROUPS'
    }

    $r3Status = if ($m365GroupCreationRestricted) { 'Compliant' } else { 'Warning' }
    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-03"
        Category        = "Users & App Registration"
        Title           = "Microsoft 365 Group Creation Governance"
        CurrentValue    = if ($m365GroupCreationRestricted) { "Restricted to designated Security Group" } else { "Unrestricted (All users can create M365 Groups & Teams)" }
        Recommended     = "Restricted to an approved administrator security group"
        Status          = $r3Status
        Severity        = "Medium"
        ActionPriority  = "Architectural Change"
        BenchmarkRef    = "Microsoft Security Best Practices - Group Lifecycle Management"
        Impact          = "Default settings allow any user to create M365 Groups, spawning unmanaged SharePoint sites, Teams, and shared mailboxes that lack DLP classifications."
        RemediationPath = "Microsoft 365 admin center > Settings > Org settings > Microsoft 365 Groups > Restrict group creation to a dedicated security group"
        PowerShellCmd   = "# Configure Group.Unified template in Entra ID via AzureADPreview / Graph API to restrict EnableGroupCreation to specific Group ID"
    })

    # Rule 4: Restrict Access to Microsoft Entra admin center
    if ($null -eq $authPolicy) {
        $r4Status = 'Warning'
        $r4Current = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $blockMsolPowerShell = Get-ObjectProperty -targetObject $authPolicy -propertyName 'blockMsolPowerShell'
        $r4Status = if ($blockMsolPowerShell -eq $true) { 'Compliant' } else { 'Non-Compliant' }
        $r4Current = if ($null -eq $blockMsolPowerShell) { "Unknown" } else { "blockMsolPowerShell = $blockMsolPowerShell" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-04"
        Category        = "Users & App Registration"
        Title           = "Restrict Access to Microsoft Entra Admin Center"
        CurrentValue    = $r4Current
        Recommended     = "blockMsolPowerShell = True (Restricted to Administrators)"
        Status          = $r4Status
        Severity        = "High"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.4"
        Impact          = "Unrestricted access allows non-privileged standard users to navigate administrative portals and enumerate sensitive directory details, tenant configuration, and user metadata."
        RemediationPath = "Entra admin center > Identity > Users > User settings > 'Restrict access to Microsoft Entra admin center' = Yes"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -BlockMsolPowerShell `$true"
    })

    # Rule 5: User Consent Settings for Applications
    if ($null -eq $authPolicy) {
        $r5Status = 'Warning'
        $r5Current = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $permissionGrantPolicies = if ($defaultUserPerms) { Get-ObjectProperty -targetObject $defaultUserPerms -propertyName 'permissionGrantPoliciesAssigned' } else { $null }
        $r5Status = 'Compliant'
        $r5Current = "Do not allow user consent (Strictly Disabled)"
        if ($permissionGrantPolicies -and ((Get-SafeCount $permissionGrantPolicies) -gt 0)) {
            $policiesJoined = $permissionGrantPolicies -join ', '
            if ($policiesJoined -match 'microsoft-user-default-legacy') {
                $r5Status = 'Non-Compliant'
                $r5Current = "Permissive user consent enabled ($policiesJoined)"
            } elseif ($policiesJoined -match 'microsoft-user-default-low') {
                $r5Status = 'Warning'
                $r5Current = "Allow user consent for verified publishers with low-impact permissions"
            } else {
                $r5Current = "Custom consent policies assigned: $policiesJoined"
            }
        }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-05"
        Category        = "Users & App Registration"
        Title           = "User Consent to Third-Party Applications"
        CurrentValue    = $r5Current
        Recommended     = "Do not allow user consent (Require Admin Consent Workflow)"
        Status          = $r5Status
        Severity        = "Critical"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.3"
        Impact          = "Consent phishing attacks trick users into granting malicious applications access to corporate emails, OneDrive files, and contacts without triggering MFA alerts."
        RemediationPath = "Entra admin center > Applications > Enterprise applications > Consent and permissions > 'User consent for applications' = 'Do not allow user consent'"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ PermissionGrantPoliciesAssigned = @() }"
    })

    # ==============================================================================
    # 2. AUDIT: GUEST ACCESS SETTINGS
    # ==============================================================================
    Write-AuditLog -Message "Auditing Guest Access & External Collaboration Settings..." -Level 'INFO' -Component 'GUESTS'

    # Rule 6: Guest user access restrictions
    if ($null -eq $authPolicy) {
        $r6Status = 'Warning'
        $guestRoleDescription = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $guestUserRoleId = Get-ObjectProperty -targetObject $authPolicy -propertyName 'guestUserRoleId'
        $guestRoleDescription = switch ($guestUserRoleId) {
            '2af84b1e-32c8-42b7-82bc-daa82404023b' { 'Restricted access (Most restrictive - Guests can only view own profile)' }
            'a0b1b346-4d3e-4e8b-98f8-753987be4970' { 'Limited access (Default - Guests can view directory objects but not list memberships)' }
            '10dae51f-b6e1-4b42-bf46-604263b49639' { 'Same as member users (Least restrictive - High risk)' }
            Default { "Custom / Unknown ID: $guestUserRoleId" }
        }

        $r6Status = switch ($guestUserRoleId) {
            '2af84b1e-32c8-42b7-82bc-daa82404023b' { 'Compliant' }
            'a0b1b346-4d3e-4e8b-98f8-753987be4970' { 'Warning' }
            Default { 'Non-Compliant' }
        }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-06"
        Category        = "Guest Access"
        Title           = "Guest User Directory Access Restrictions"
        CurrentValue    = $guestRoleDescription
        Recommended     = "Restricted access (GuestUserRoleId = 2af84b1e-32c8-42b7-82bc-daa82404023b)"
        Status          = $r6Status
        Severity        = "High"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.5"
        Impact          = "Default or member-level guest access allows external users to discover internal employees, groups, email addresses, and organizational hierarchy."
        RemediationPath = "Entra admin center > External Identities > External collaboration settings > 'Guest user access restrictions' = 'Restricted access'"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -GuestUserRoleId '2af84b1e-32c8-42b7-82bc-daa82404023b'"
    })

    # Rule 7: Guest invitation restrictions
    if ($null -eq $authPolicy) {
        $r7Status = 'Warning'
        $r7Current = "Unable to query authorizationPolicy (403 Forbidden - Policy.Read.All scope required)"
    } else {
        $allowInvitesFrom = Get-ObjectProperty -targetObject $authPolicy -propertyName 'allowInvitesFrom'
        $r7Status = switch ($allowInvitesFrom) {
            'adminsOnly'             { 'Compliant' }
            'adminsAndGuestInviters' { 'Compliant' }
            'none'                   { 'Compliant' }
            'everyone'               { 'Non-Compliant' }
            Default                  { 'Warning' }
        }
        $r7Current = if ($allowInvitesFrom) { "allowInvitesFrom = $allowInvitesFrom" } else { "Unknown" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-07"
        Category        = "Guest Access"
        Title           = "Guest Invitation Permissions"
        CurrentValue    = $r7Current
        Recommended     = "adminsOnly or adminsAndGuestInviters"
        Status          = $r7Status
        Severity        = "High"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.1.6"
        Impact          = "Allowing any employee or external guest to invite external collaborators causes untracked external access and potential sensitive data exfiltration."
        RemediationPath = "Entra admin center > External Identities > External collaboration settings > 'Guest invite settings' = 'Only users assigned to specific admin roles can invite guest users'"
        PowerShellCmd   = "Update-MgPolicyAuthorizationPolicy -AllowInvitesFrom 'adminsAndGuestInviters'"
    })

    # ==============================================================================
    # 3. AUDIT: AUTHENTICATION & LOGIN METHODS
    # ==============================================================================
    Write-AuditLog -Message "Auditing Authentication Methods & Security Defaults..." -Level 'INFO' -Component 'AUTH_METHODS'

    # Rule 8: Security Defaults Status
    $secDefaultsPolicy = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy"
    $secDefaultsEnabled = $false
    if ($secDefaultsPolicy) {
        $secDefaultsEnabled = [bool](Get-ObjectProperty -targetObject $secDefaultsPolicy -propertyName 'isEnabled')
    }

    # Fetch Conditional Access Policies to assess mutually exclusive configuration
    Write-AuditLog -Message "Querying Conditional Access policies to evaluate baseline MFA..." -Level 'INFO' -Component 'COND_ACCESS'
    $caPolicies = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -GetAllPages
    $hasCaPolicies = ($null -ne $caPolicies -and ((Get-SafeCount $caPolicies) -gt 0))
    $enabledCaPolicies = @(
        if ($hasCaPolicies) {
            $caPolicies | Where-Object { (Get-ObjectProperty -targetObject $_ -propertyName 'state') -eq 'enabled' }
        }
    )

    $r8Status = 'Non-Compliant'
    $r8Current = ""
    if ($null -eq $secDefaultsPolicy -and $null -eq $caPolicies) {
        $r8Status = 'Warning'
        $r8Current = "Unable to query Security Defaults or Conditional Access (403 Forbidden / Insufficient Scopes)"
    } elseif ($secDefaultsEnabled) {
        $r8Status = 'Compliant'
        $r8Current = "Security Defaults is Enabled (Baseline MFA enforced)"
    } else {
        if ((Get-SafeCount $enabledCaPolicies) -gt 0) {
            $r8Status = 'Compliant'
            $r8Current = "Security Defaults Disabled; $((Get-SafeCount $enabledCaPolicies)) active Conditional Access policies detected (Best Practice for P1/P2)"
        } else {
            $r8Status = 'Non-Compliant'
            $r8Current = "Security Defaults is Disabled AND no active Conditional Access policies found!"
        }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-08"
        Category        = "Authentication & SSPR"
        Title           = "Security Defaults / Baseline MFA Enforcement"
        CurrentValue    = $r8Current
        Recommended     = "Enabled (if Entra ID Free) OR Disabled with custom Conditional Access MFA policies (if P1/P2)"
        Status          = $r8Status
        Severity        = "Critical"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.2.1"
        Impact          = "Without Security Defaults or Conditional Access policies, users and administrators are not required to use MFA, leaving accounts vulnerable to basic password spraying."
        RemediationPath = "Entra admin center > Identity > Overview > Properties > 'Manage security defaults' = Enable (or implement Conditional Access)"
        PowerShellCmd   = "# Enable Security Defaults: Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -IsEnabled `$true (or configure CA)"
    })

    # Rule 9: Self-Service Password Reset (SSPR)
    Write-AuditLog -Message "Auditing Self-Service Password Reset (SSPR) status..." -Level 'INFO' -Component 'SSPR'
    $ssprPolicy = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/beta/policies/authorizationPolicy"
    $allowedToUseSspr = $null
    if ($ssprPolicy) {
        $defaultUserPermsBeta = Get-ObjectProperty -targetObject $ssprPolicy -propertyName 'defaultUserRolePermissions'
        if ($defaultUserPermsBeta) {
            $allowedToUseSspr = Get-ObjectProperty -targetObject $defaultUserPermsBeta -propertyName 'allowedToUseSspr'
        }
    }

    $r9Status = if ($allowedToUseSspr -eq $true) { 'Compliant' } elseif ($null -eq $allowedToUseSspr) { 'Compliant' } else { 'Warning' }
    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-09"
        Category        = "Authentication & SSPR"
        Title           = "Self-Service Password Reset (SSPR) Deployment"
        CurrentValue    = if ($null -ne $allowedToUseSspr) { "allowedToUseSspr = $allowedToUseSspr" } else { "SSPR configured in tenant policies" }
        Recommended     = "SSPR enabled for All users (or scoped pilot security group) with strong multi-method registration"
        Status          = $r9Status
        Severity        = "Medium"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.2.2"
        Impact          = "Lack of self-service password reset forces helpdesk reliance for password changes and risks social engineering attacks against support personnel."
        RemediationPath = "Entra admin center > Protection > Password reset > Properties > Set 'Self service password reset enabled' to 'All'"
        PowerShellCmd   = "# Configure SSPR via Entra admin center or Microsoft Graph authenticationMethodsPolicy"
    })

    # ==============================================================================
    # 4. AUDIT: CONDITIONAL ACCESS POLICIES (P1/P2)
    # ==============================================================================
    Write-AuditLog -Message "Auditing Conditional Access Policies for Administrator MFA, Legacy Auth, and Risk..." -Level 'INFO' -Component 'COND_ACCESS'

    $adminMfaPolicyFound = $false
    $legacyAuthBlocked = $false
    $riskPolicyFound = $false

    if ($null -eq $caPolicies) {
        Write-AuditLog -Message "Conditional Access policies could not be retrieved (requires Entra ID P1/P2 and Policy.Read.ConditionalAccess)." -Level 'WARN' -Component 'COND_ACCESS'

        $r11Status = 'Warning'
        $r11Current = "Conditional Access unavailable (403 Forbidden / Not Licensed - Requires Entra ID P1/P2)"

        $r12Status = 'Warning'
        $r12Current = "Conditional Access unavailable (403 Forbidden / Not Licensed - Requires Entra ID P1/P2)"

        $r13Status = 'Warning'
        $r13Current = "Conditional Access unavailable (403 Forbidden / Not Licensed - Requires Entra ID P2)"
    } else {
        if ((Get-SafeCount $enabledCaPolicies) -gt 0) {
            foreach ($policy in $enabledCaPolicies) {
                $pName = Get-ObjectProperty -targetObject $policy -propertyName 'displayName'
                Write-AuditLog -Message "Evaluating CA Policy: $pName" -Level 'DEBUG' -Component 'COND_ACCESS'

                $conditions = Get-ObjectProperty -targetObject $policy -propertyName 'conditions'
                $grantControls = Get-ObjectProperty -targetObject $policy -propertyName 'grantControls'

                $builtInControls = @()
                if ($grantControls) {
                    $bic = Get-ObjectProperty -targetObject $grantControls -propertyName 'builtInControls'
                    if ($bic) { $builtInControls = @($bic) }
                }

                # Check 1: Admin MFA Policy
                if ($builtInControls -contains 'mfa') {
                    if ($conditions) {
                        $usersCondition = Get-ObjectProperty -targetObject $conditions -propertyName 'users'
                        if ($usersCondition) {
                            $incRoles = Get-ObjectProperty -targetObject $usersCondition -propertyName 'includeRoles'
                            $incUsers = Get-ObjectProperty -targetObject $usersCondition -propertyName 'includeUsers'
                            if (((Get-SafeCount $incRoles) -gt 0) -or ($incUsers -contains 'All')) {
                                $adminMfaPolicyFound = $true
                            }
                        }
                    }
                }

                # Check 2: Block Legacy Authentication
                if ($builtInControls -contains 'block') {
                    if ($conditions) {
                        $clientAppTypes = Get-ObjectProperty -targetObject $conditions -propertyName 'clientAppTypes'
                        if ($clientAppTypes -and (($clientAppTypes -contains 'exchangeActiveSync') -or ($clientAppTypes -contains 'other'))) {
                            $legacyAuthBlocked = $true
                        }
                    }
                }

                # Check 3: Risky Sign-ins Policy
                if ($conditions) {
                    $signInRisk = Get-ObjectProperty -targetObject $conditions -propertyName 'signInRiskLevels'
                    if ($signInRisk -and ((Get-SafeCount $signInRisk) -gt 0)) {
                        if ($builtInControls -contains 'mfa' -or $builtInControls -contains 'block') {
                            $riskPolicyFound = $true
                        }
                    }
                }
            }
        }

        # Rule 11: Conditional Access MFA for Admins
        $r11Status = if ($adminMfaPolicyFound) { 'Compliant' } elseif (-not $hasCaPolicies -and $secDefaultsEnabled) { 'Compliant' } else { 'Non-Compliant' }
        $r11Current = if ($adminMfaPolicyFound) { "Active CA policy enforcing MFA for Directory Roles detected" } elseif ($secDefaultsEnabled) { "Covered by Security Defaults baseline" } else { "No active policy enforcing MFA specifically for administrative roles" }

        # Rule 12: Conditional Access Block Legacy Auth
        $r12Status = if ($legacyAuthBlocked) { 'Compliant' } elseif ($secDefaultsEnabled) { 'Compliant' } else { 'Non-Compliant' }
        $r12Current = if ($legacyAuthBlocked) { "Active CA policy blocking Legacy Authentication (ExchangeActiveSync / Other clients)" } elseif ($secDefaultsEnabled) { "Blocked automatically by Security Defaults" } else { "Legacy authentication is NOT blocked via Conditional Access" }

        # Rule 13: Conditional Access Risk-Based Sign-In Policy
        $r13Status = if ($riskPolicyFound) { 'Compliant' } else { 'Warning' }
        $r13Current = if ($riskPolicyFound) { "Active policy enforcing MFA or Block for High/Medium Sign-In Risk" } else { "No active risk-based sign-in policy configured" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-11"
        Category        = "Conditional Access"
        Title           = "Conditional Access: Enforce MFA for All Administrators"
        CurrentValue    = $r11Current
        Recommended     = "Active Conditional Access policy requiring MFA for all directory administrative roles"
        Status          = $r11Status
        Severity        = "Critical"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.3.1"
        Impact          = "Privileged administrative accounts must never be accessed without multi-factor authentication; single-factor compromise gives attackers full control of the tenant."
        RemediationPath = "Entra admin center > Protection > Conditional Access > Create policy > Users: Directory roles (All admins) > Grant: Require multifactor authentication"
        PowerShellCmd   = "# Create Conditional Access Policy targeting Directory Roles with builtInControls = @('mfa')"
    })

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-12"
        Category        = "Conditional Access"
        Title           = "Conditional Access: Block Legacy Authentication Protocols"
        CurrentValue    = $r12Current
        Recommended     = "Active Conditional Access policy blocking client app types: 'ExchangeActiveSync' and 'Other clients'"
        Status          = $r12Status
        Severity        = "High"
        ActionPriority  = "Architectural Change"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.3.2"
        Impact          = "Legacy protocols (POP3, IMAP, SMTP, MAPI) do not support modern MFA challenges, allowing attackers to perform credential spraying and bypass MFA completely."
        RemediationPath = "Entra admin center > Protection > Conditional Access > Create policy > Conditions > Client apps: Exchange ActiveSync & Other clients > Grant: Block access"
        PowerShellCmd   = "# Create Conditional Access Policy targeting clientAppTypes @('exchangeActiveSync', 'other') with builtInControls = @('block')"
    })

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-13"
        Category        = "Conditional Access"
        Title           = "Conditional Access: Risk-Based Sign-In Protection (P2)"
        CurrentValue    = $r13Current
        Recommended     = "Active Conditional Access policy requiring MFA or blocking access when Sign-in Risk is Medium or High"
        Status          = $r13Status
        Severity        = "High"
        ActionPriority  = "Architectural Change"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.3.3"
        Impact          = "Risk-based policies detect anomalous sign-in tokens, unfamiliar locations, and anonymous IP proxies in real time, stopping compromised session replay."
        RemediationPath = "Entra admin center > Protection > Conditional Access > Policies > New policy > Conditions > Sign-in risk: High and Medium > Grant: Require MFA"
        PowerShellCmd   = "# Requires Entra ID P2 license: configure policy with signInRiskLevels = @('high', 'medium')"
    })

    # ==============================================================================
    # 5. AUDIT: ADMINISTRATIVE ROLES & BREAK-GLASS ACCOUNTS
    # ==============================================================================
    Write-AuditLog -Message "Auditing Directory Roles, Global Administrators, and Privileged MFA..." -Level 'INFO' -Component 'ROLES'

    $directoryRoles = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/directoryRoles" -GetAllPages
    $globalAdminRole = $null
    $secAdminRole    = $null
    $privRoleAdmin   = $null

    if ($directoryRoles) {
        foreach ($dr in $directoryRoles) {
            $templateId = Get-ObjectProperty -targetObject $dr -propertyName 'roleTemplateId'
            switch ($templateId) {
                '62e90394-69f5-4237-9190-012177145e10' { $globalAdminRole = $dr }
                '194ae4cb-b126-40b2-bd5b-6091b380977d' { $secAdminRole    = $dr }
                'e8611ab8-c189-44e8-9421-60601ac9c477' { $privRoleAdmin   = $dr }
            }
        }
    }

    $globalAdminUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
    $allPrivilegedUsers = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Get-RoleMembersSafe {
        param($roleObj, [string]$roleName)
        if ($null -eq $roleObj) { return }
        $roleId = Get-ObjectProperty -targetObject $roleObj -propertyName 'id'
        if (-not $roleId) { return }

        $members = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$roleId/members" -GetAllPages
        if ($members) {
            foreach ($m in $members) {
                $odataType = Get-ObjectProperty -targetObject $m -propertyName '@odata.type'
                # Filter specifically for user accounts (exclude service principals)
                if ($odataType -eq '#microsoft.graph.user' -or (-not $odataType -and (Get-ObjectProperty -targetObject $m -propertyName 'userPrincipalName'))) {
                    $uObj = [PSCustomObject]@{
                        Id                = Get-ObjectProperty -targetObject $m -propertyName 'id'
                        DisplayName       = Get-ObjectProperty -targetObject $m -propertyName 'displayName'
                        UserPrincipalName = Get-ObjectProperty -targetObject $m -propertyName 'userPrincipalName'
                        RoleName          = $roleName
                        AccountEnabled    = Get-ObjectProperty -targetObject $m -propertyName 'accountEnabled'
                        IsMfaRegistered   = $false
                    }
                    $allPrivilegedUsers.Add($uObj)
                    if ($roleName -eq 'Global Administrator') {
                        $globalAdminUsers.Add($uObj)
                    }
                }
            }
        }
    }

    Get-RoleMembersSafe -roleObj $globalAdminRole -roleName 'Global Administrator'
    Get-RoleMembersSafe -roleObj $secAdminRole    -roleName 'Security Administrator'
    Get-RoleMembersSafe -roleObj $privRoleAdmin   -roleName 'Privileged Role Administrator'

    # Retrieve User Authentication Methods Registration Details
    Write-AuditLog -Message "Querying MFA registration details for privileged role accounts..." -Level 'INFO' -Component 'MFA_REG'
    $userRegDetails = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails" -GetAllPages

    $mfaLookup = @{}
    if ($userRegDetails) {
        foreach ($reg in $userRegDetails) {
            $regId = Get-ObjectProperty -targetObject $reg -propertyName 'id'
            $isMfa = Get-ObjectProperty -targetObject $reg -propertyName 'isMfaRegistered'
            if ($regId) {
                $mfaLookup[$regId] = [bool]$isMfa
            }
        }
    }

    # Cross-reference privileged accounts with MFA registration
    $unprotectedAdmins = [System.Collections.Generic.List[string]]::new()
    $unprotectedGlobalAdmins = [System.Collections.Generic.List[string]]::new()

    foreach ($pUser in $allPrivilegedUsers) {
        if ($mfaLookup.ContainsKey($pUser.Id)) {
            $pUser.IsMfaRegistered = $mfaLookup[$pUser.Id]
        } else {
            # Secondary check fallback: inspect user authentication methods
            try {
                $uMethods = Invoke-EntraGraphQuery -Uri "https://graph.microsoft.com/v1.0/users/$($pUser.Id)/authentication/methods"
                if ($uMethods -and ((Get-SafeCount $uMethods) -gt 1)) {
                    $pUser.IsMfaRegistered = $true
                }
            }
            catch {
                $null = $_
            }
        }

        if (-not $pUser.IsMfaRegistered) {
            $unprotectedAdmins.Add("$($pUser.DisplayName) ($($pUser.UserPrincipalName)) [Role: $($pUser.RoleName)]")
            if ($pUser.RoleName -eq 'Global Administrator') {
                $unprotectedGlobalAdmins.Add("$($pUser.DisplayName) ($($pUser.UserPrincipalName))")
            }
        }
        $Script:PrivilegedUsers.Add($pUser)
    }

    # Rule 10: MFA Registration for Privileged Roles
    if ($null -eq $directoryRoles) {
        $r10Status = 'Warning'
        $r10Current = "Unable to query directory roles (403 Forbidden - RoleManagement.Read.Directory required)"
    } else {
        $r10Status = if ($unprotectedAdmins.Count -eq 0 -and $allPrivilegedUsers.Count -gt 0) { 'Compliant' } elseif ($allPrivilegedUsers.Count -eq 0) { 'Warning' } else { 'Non-Compliant' }
        $r10Current = if ($unprotectedAdmins.Count -eq 0) { "All privileged role holders ($($allPrivilegedUsers.Count) accounts) are registered for MFA" } else { "$($unprotectedAdmins.Count) of $($allPrivilegedUsers.Count) privileged accounts lack MFA registration: $($unprotectedAdmins -join '; ')" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-10"
        Category        = "Authentication & SSPR"
        Title           = "MFA Registration for Privileged Administrative Roles"
        CurrentValue    = $r10Current
        Recommended     = "100% of privileged directory role holders registered and enforced with MFA"
        Status          = $r10Status
        Severity        = "Critical"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.2.3"
        Impact          = "Any administrative account lacking multi-factor authentication can be compromised via phishing or credential stuffing, granting attackers tenant-wide control."
        RemediationPath = "Entra admin center > Protection > Authentication methods > User registration details > Filter unregistered admins and require MFA setup"
        PowerShellCmd   = "# Enforce MFA via Conditional Access policy requiring MFA for all directory roles"
    })

    # Rule 14: Global Administrator Account Count & Break-Glass Hygiene
    if ($null -eq $directoryRoles) {
        $r14Status = 'Warning'
        $gaCurrentDesc = "Unable to query directory roles (403 Forbidden - RoleManagement.Read.Directory required)"
    } else {
        $gaCount = $globalAdminUsers.Count
        $r14Status = if ($gaCount -ge 2 -and $gaCount -le 4) {
            'Compliant'
        } elseif ($gaCount -lt 2) {
            'Non-Compliant'
        } else {
            'Warning'
        }

        $gaCurrentDesc = if ($gaCount -eq 0) {
            "0 Global Administrators discovered"
        } elseif ($gaCount -lt 2) {
            "Only $gaCount Global Administrator detected (Critical lockout risk - minimum 2 required)"
        } elseif ($gaCount -le 4) {
            "$gaCount Global Administrators detected (Optimal compliance range: 2-4 accounts)"
        } else {
            "$gaCount Global Administrators detected (Excessive count - exceeds maximum recommendation of 4)"
        }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-14"
        Category        = "Roles & Emergency Accounts"
        Title           = "Global Administrator Account Count Hygiene"
        CurrentValue    = $gaCurrentDesc
        Recommended     = "Between 2 and 4 Global Administrators total (including 2 dedicated emergency break-glass accounts)"
        Status          = $r14Status
        Severity        = "High"
        ActionPriority  = "Architectural Change"
        BenchmarkRef    = "CIS M365 Foundations Benchmark v3.0 - 5.4.1"
        Impact          = "Having fewer than 2 Global Admins risks permanent tenant lockout if the sole admin is unavailable. Having more than 4 increases the privileged attack surface."
        RemediationPath = "Entra admin center > Identity > Roles and administrators > Global Administrator > Demote excess accounts to role-specific privileges (e.g. User Admin)"
        PowerShellCmd   = "# Review role assignments: Get-MgDirectoryRoleMember -DirectoryRoleId <GlobalAdminId>"
    })

    # Rule 15: Unprotected Global Administrators
    if ($null -eq $directoryRoles) {
        $r15Status = 'Warning'
        $r15Current = "Unable to query directory roles (403 Forbidden - RoleManagement.Read.Directory required)"
    } else {
        $r15Status = if ($unprotectedGlobalAdmins.Count -eq 0 -and $globalAdminUsers.Count -gt 0) { 'Compliant' } elseif ($globalAdminUsers.Count -eq 0) { 'Warning' } else { 'Non-Compliant' }
        $r15Current = if ($unprotectedGlobalAdmins.Count -eq 0) { "All ($($globalAdminUsers.Count)) Global Administrators have registered MFA methods" } else { "$($unprotectedGlobalAdmins.Count) Global Administrator(s) lack registered MFA: $($unprotectedGlobalAdmins -join ', ')" }
    }

    $Script:AuditFindings.Add([PSCustomObject]@{
        Id              = "ENTRA-SEC-15"
        Category        = "Roles & Emergency Accounts"
        Title           = "Global Administrators Protected with MFA"
        CurrentValue    = $r15Current
        Recommended     = "0 Global Administrators without MFA (All superadministrators must have strong/FIDO2 MFA)"
        Status          = $r15Status
        Severity        = "Critical"
        ActionPriority  = "Quick Win"
        BenchmarkRef    = "Microsoft Security Best Practices - Emergency Access & Privileged Identity"
        Impact          = "A single Global Administrator without MFA represents a total tenant takeover vulnerability. Immediate remediation required."
        RemediationPath = "Entra admin center > Users > Select affected Global Administrator > Authentication methods > Register FIDO2 or Authenticator method"
        PowerShellCmd   = "# Immediately enforce MFA on affected Global Administrator accounts"
    })

    # ==============================================================================
    # 6. EXECUTIVE SCORING & COMPLIANCE METRICS
    # ==============================================================================
    $totalRules     = $Script:AuditFindings.Count
    $compliantRules = @($Script:AuditFindings | Where-Object { $_.Status -eq 'Compliant' }).Count
    $nonCompliant   = @($Script:AuditFindings | Where-Object { $_.Status -eq 'Non-Compliant' }).Count
    $warningRules   = @($Script:AuditFindings | Where-Object { $_.Status -eq 'Warning' }).Count

    $criticalIssues = @($Script:AuditFindings | Where-Object { $_.Severity -eq 'Critical' -and $_.Status -ne 'Compliant' }).Count
    $highIssues     = @($Script:AuditFindings | Where-Object { $_.Severity -eq 'High'     -and $_.Status -ne 'Compliant' }).Count
    $mediumIssues   = @($Script:AuditFindings | Where-Object { $_.Severity -eq 'Medium'   -and $_.Status -ne 'Compliant' }).Count
    $lowIssues      = @($Script:AuditFindings | Where-Object { $_.Severity -eq 'Low'      -and $_.Status -ne 'Compliant' }).Count

    $compliancePercentage = if ($totalRules -gt 0) { [math]::Round(($compliantRules / $totalRules) * 100, 1) } else { 0 }

    Write-AuditLog -Message "Audit complete! Score: $compliancePercentage% ($compliantRules/$totalRules Compliant, $nonCompliant Non-Compliant, $warningRules Warnings)" -Level 'SUCCESS' -Component 'AUDIT'

    # ==============================================================================
    # 7. GENERATE STANDALONE INTERACTIVE HTML DASHBOARD
    # ==============================================================================
    function Export-EntraIDSecurityHtmlReport {
        [CmdletBinding()]
        param()

        Write-AuditLog -Message "Building standalone interactive HTML executive report..." -Level 'INFO' -Component 'REPORT'

        try {
            if (-not (Test-Path -Path $OutputPath)) {
                $null = New-Item -Path $OutputPath -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            $OutputPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Reports\Audit-EntraIDSecuritySummary"
            if (-not (Test-Path -Path $OutputPath)) {
                $null = New-Item -Path $OutputPath -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
        }

        $reportDateStr = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $Script:GeneratedHtmlPath = Join-Path -Path $OutputPath -ChildPath "AuditReport-EntraID-$reportDateStr.html"

        # SVG Circle Calculations for Score Gauge
        $radius = 54
        $circumference = [math]::Round(2 * [math]::PI * $radius, 2)
        $offset = [math]::Round($circumference - (($compliancePercentage / 100) * $circumference), 2)
        $scoreColor = if ($compliancePercentage -ge 80) { '#10b981' } elseif ($compliancePercentage -ge 60) { '#f59e0b' } else { '#ef4444' }

        # Action Matrix Rows sorted by severity/priority
        $sortedFindings = $Script:AuditFindings | Sort-Object @{ Expression = {
            switch ($_.Severity) {
                'Critical' { 1 }
                'High'     { 2 }
                'Medium'   { 3 }
                'Low'      { 4 }
                Default    { 5 }
            }
        }}, @{ Expression = { if ($_.Status -eq 'Non-Compliant') { 1 } elseif ($_.Status -eq 'Warning') { 2 } else { 3 } } }

        $matrixRowsHtml = [System.Text.StringBuilder]::new()
        foreach ($f in $sortedFindings) {
            $sevBadgeClass = switch ($f.Severity) {
                'Critical' { 'badge-critical' }
                'High'     { 'badge-high' }
                'Medium'   { 'badge-medium' }
                'Low'      { 'badge-low' }
                Default    { 'badge-neutral' }
            }

            $statusBadgeClass = switch ($f.Status) {
                'Compliant'     { 'badge-compliant' }
                'Non-Compliant' { 'badge-noncompliant' }
                'Warning'       { 'badge-warning' }
                Default         { 'badge-neutral' }
            }

            $prioClass = if ($f.ActionPriority -eq 'Quick Win') { 'pill-quickwin' } else { 'pill-arch' }

            $matrixRowsHtml.Append(@"
                <tr class="action-row" data-category="$($f.Category)" data-status="$($f.Status)" data-severity="$($f.Severity)">
                    <td><span class="badge $sevBadgeClass">$($f.Severity)</span></td>
                    <td><span class="badge-priority $prioClass">$($f.ActionPriority)</span></td>
                    <td class="font-mono text-xs"><b>$($f.Id)</b></td>
                    <td><b>$($f.Title)</b><div class="text-xs text-muted">$($f.Category)</div></td>
                    <td><span class="badge $statusBadgeClass">$($f.Status)</span></td>
                    <td class="text-xs font-mono">$([System.Web.HttpUtility]::HtmlEncode($f.CurrentValue))</td>
                    <td class="text-xs text-emerald font-semibold">$([System.Web.HttpUtility]::HtmlEncode($f.Recommended))</td>
                </tr>
"@) | Out-Null
        }

        # Recommendation Cards HTML
        $cardsHtml = [System.Text.StringBuilder]::new()
        foreach ($f in $Script:AuditFindings) {
            $sevBadgeClass = switch ($f.Severity) {
                'Critical' { 'badge-critical' }
                'High'     { 'badge-high' }
                'Medium'   { 'badge-medium' }
                'Low'      { 'badge-low' }
                Default    { 'badge-neutral' }
            }

            $statusBadgeClass = switch ($f.Status) {
                'Compliant'     { 'badge-compliant' }
                'Non-Compliant' { 'badge-noncompliant' }
                'Warning'       { 'badge-warning' }
                Default         { 'badge-neutral' }
            }

            $cardBorderClass = switch ($f.Status) {
                'Non-Compliant' { 'card-border-danger' }
                'Warning'       { 'card-border-warning' }
                'Compliant'     { 'card-border-success' }
                Default         { '' }
            }

            $cardsHtml.Append(@"
            <div class="recommendation-card $cardBorderClass" data-category="$($f.Category)" data-status="$($f.Status)" data-severity="$($f.Severity)" data-keywords="$($f.Id) $($f.Title) $($f.Category) $($f.Status) $($f.Severity)">
                <div class="card-header" onclick="toggleCard(this)">
                    <div class="card-header-left">
                        <span class="badge $statusBadgeClass">$($f.Status)</span>
                        <span class="badge $sevBadgeClass">$($f.Severity)</span>
                        <span class="badge-tag">$($f.Id)</span>
                        <span class="card-title">$($f.Title)</span>
                    </div>
                    <div class="card-header-right">
                        <span class="badge-category">$($f.Category)</span>
                        <span class="chevron">▼</span>
                    </div>
                </div>
                <div class="card-body">
                    <div class="comparison-grid">
                        <div class="state-box current-state">
                            <div class="state-title">Stan obecny (Current Tenant State)</div>
                            <div class="state-value">$([System.Web.HttpUtility]::HtmlEncode($f.CurrentValue))</div>
                        </div>
                        <div class="state-box target-state">
                            <div class="state-title">Rekomendacja (Target State)</div>
                            <div class="state-value">$([System.Web.HttpUtility]::HtmlEncode($f.Recommended))</div>
                        </div>
                    </div>

                    <div class="info-section">
                        <div class="info-heading">Dlaczego to ustawienie jest istotne? (Business & Security Rationale)</div>
                        <div class="info-content">$($f.Impact)</div>
                        <div class="benchmark-tag">Referencja: $($f.BenchmarkRef)</div>
                    </div>

                    <div class="remediation-section">
                        <div class="info-heading">Kroki naprawcze (Remediation Guidance)</div>
                        <div class="remediation-path">
                            <span class="path-icon">📂</span>
                            <span>$($f.RemediationPath)</span>
                        </div>
                        <div class="cmd-box">
                            <code>$([System.Web.HttpUtility]::HtmlEncode($f.PowerShellCmd))</code>
                            <button class="copy-btn" onclick="copyToClipboard(this)">Kopiuj</button>
                        </div>
                    </div>
                </div>
            </div>
"@) | Out-Null
        }

        # Privileged Users Table HTML
        $privUsersHtml = [System.Text.StringBuilder]::new()
        foreach ($u in $Script:PrivilegedUsers) {
            $mfaClass = if ($u.IsMfaRegistered) { 'badge-compliant' } else { 'badge-critical' }
            $mfaText  = if ($u.IsMfaRegistered) { 'Enrolled / MFA OK' } else { 'NO MFA DETECTED' }

            $privUsersHtml.Append(@"
                <tr>
                    <td><b>$([System.Web.HttpUtility]::HtmlEncode($u.DisplayName))</b></td>
                    <td class="font-mono text-xs">$([System.Web.HttpUtility]::HtmlEncode($u.UserPrincipalName))</td>
                    <td><span class="badge badge-high">$($u.RoleName)</span></td>
                    <td><span class="badge $mfaClass">$mfaText</span></td>
                    <td>$($u.AccountEnabled)</td>
                </tr>
"@) | Out-Null
        }

        $htmlContent = @"
<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Entra ID Security Summary - $($Script:ActiveTenantName)</title>
    <style>
        :root {
            --bg-body: #f8fafc;
            --bg-card: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border-color: #e2e8f0;
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --info: #0ea5e9;
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.05), 0 2px 4px -2px rgb(0 0 0 / 0.05);
            --radius: 10px;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background: var(--bg-body); color: var(--text-main); line-height: 1.5; padding: 24px; }
        .container { max-width: 1400px; margin: 0 auto; }

        /* HEADER */
        .header-card { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #ffffff; border-radius: var(--radius); padding: 30px; box-shadow: var(--shadow); margin-bottom: 24px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 20px; }
        .header-title h1 { font-size: 26px; font-weight: 700; margin-bottom: 6px; letter-spacing: -0.5px; }
        .header-title p { color: #94a3b8; font-size: 14px; }
        .header-meta { display: flex; gap: 20px; flex-wrap: wrap; font-size: 13px; color: #cbd5e1; }
        .header-meta-item { background: rgba(255,255,255,0.08); padding: 6px 12px; border-radius: 6px; }

        /* KPI SCORE DASHBOARD */
        .dashboard-grid { display: grid; grid-template-columns: 280px repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .kpi-card { background: var(--bg-card); border-radius: var(--radius); padding: 20px; border: 1px solid var(--border-color); box-shadow: var(--shadow); display: flex; flex-direction: column; justify-content: center; }
        .gauge-card { align-items: center; text-align: center; }
        .gauge-svg { width: 140px; height: 140px; }
        .gauge-bg { fill: none; stroke: #e2e8f0; stroke-width: 10; }
        .gauge-bar { fill: none; stroke: $scoreColor; stroke-width: 10; stroke-linecap: round; stroke-dasharray: $circumference; stroke-dashoffset: $offset; transition: stroke-dashoffset 1s ease-in-out; }
        .gauge-text { font-size: 26px; font-weight: 800; fill: var(--text-main); }
        .gauge-subtext { font-size: 11px; fill: var(--text-muted); font-weight: 600; text-transform: uppercase; }

        .kpi-num { font-size: 32px; font-weight: 800; line-height: 1.1; margin-top: 4px; }
        .kpi-label { font-size: 12px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); letter-spacing: 0.5px; }
        .kpi-danger .kpi-num { color: var(--danger); }
        .kpi-warning .kpi-num { color: var(--warning); }
        .kpi-success .kpi-num { color: var(--success); }
        .kpi-info .kpi-num { color: var(--info); }

        /* CONTROLS & FILTER BAR */
        .controls-card { background: var(--bg-card); border-radius: var(--radius); padding: 16px 20px; border: 1px solid var(--border-color); box-shadow: var(--shadow); margin-bottom: 24px; display: flex; gap: 12px; flex-wrap: wrap; align-items: center; justify-content: space-between; }
        .filter-group { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
        .filter-btn { background: #f1f5f9; border: 1px solid var(--border-color); border-radius: 6px; padding: 7px 14px; font-size: 13px; font-weight: 600; color: var(--text-main); cursor: pointer; transition: all 0.2s; }
        .filter-btn:hover { background: #e2e8f0; }
        .filter-btn.active { background: var(--primary); color: #ffffff; border-color: var(--primary); }
        .search-box { flex-grow: 1; max-width: 320px; position: relative; }
        .search-input { width: 100%; padding: 8px 12px; font-size: 13px; border-radius: 6px; border: 1px solid var(--border-color); outline: none; transition: border-color 0.2s; }
        .search-input:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1); }

        /* SECTION HEADINGS */
        .section-header { margin: 30px 0 16px 0; display: flex; justify-content: space-between; align-items: baseline; }
        .section-title { font-size: 20px; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 10px; }
        .section-subtitle { font-size: 13px; color: var(--text-muted); }

        /* TABLES */
        .table-container { background: var(--bg-card); border-radius: var(--radius); border: 1px solid var(--border-color); box-shadow: var(--shadow); overflow-x: auto; margin-bottom: 24px; }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }
        th { background: #f8fafc; padding: 12px 16px; font-weight: 600; color: var(--text-muted); border-bottom: 1px solid var(--border-color); text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; }
        td { padding: 12px 16px; border-bottom: 1px solid var(--border-color); vertical-align: middle; }
        tr:last-child td { border-bottom: none; }
        tr:hover { background: #f8fafc; }

        /* BADGES */
        .badge { display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px; }
        .badge-compliant { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }
        .badge-noncompliant { background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca; }
        .badge-warning { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
        .badge-critical { background: #7f1d1d; color: #fee2e2; }
        .badge-high { background: #ffedd5; color: #c2410c; }
        .badge-medium { background: #fef9c3; color: #854d0e; }
        .badge-low { background: #e0f2fe; color: #0369a1; }
        .badge-neutral { background: #f1f5f9; color: #475569; }
        .badge-tag { font-family: monospace; font-size: 11px; background: #e2e8f0; color: #334155; padding: 2px 6px; border-radius: 4px; }
        .badge-category { font-size: 11px; color: var(--text-muted); font-weight: 600; }
        .badge-priority { font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 4px; display: inline-block; }
        .pill-quickwin { background: #e0e7ff; color: #3730a3; border: 1px solid #c7d2fe; }
        .pill-arch { background: #fae8ff; color: #86198f; border: 1px solid #f5d0fe; }

        /* RECOMMENDATION CARDS */
        .cards-list { display: flex; flex-direction: column; gap: 14px; margin-bottom: 24px; }
        .recommendation-card { background: var(--bg-card); border-radius: var(--radius); border: 1px solid var(--border-color); box-shadow: var(--shadow); overflow: hidden; transition: box-shadow 0.2s; }
        .recommendation-card:hover { box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.08); }
        .card-border-danger { border-left: 5px solid var(--danger); }
        .card-border-warning { border-left: 5px solid var(--warning); }
        .card-border-success { border-left: 5px solid var(--success); }

        .card-header { padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer; background: #ffffff; user-select: none; }
        .card-header:hover { background: #f8fafc; }
        .card-header-left { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
        .card-header-right { display: flex; align-items: center; gap: 14px; }
        .card-title { font-size: 15px; font-weight: 700; color: var(--text-main); }
        .chevron { font-size: 12px; color: var(--text-muted); transition: transform 0.2s; }
        .recommendation-card.open .chevron { transform: rotate(180deg); }

        .card-body { padding: 0 20px 20px 20px; display: none; border-top: 1px solid #f1f5f9; background: #fafafa; }
        .recommendation-card.open .card-body { display: block; padding-top: 16px; }

        .comparison-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px; }
        @media (max-width: 768px) { .comparison-grid { grid-template-columns: 1fr; } }
        .state-box { padding: 12px 14px; border-radius: 6px; font-size: 13px; }
        .current-state { background: #ffffff; border: 1px solid var(--border-color); }
        .target-state { background: #ecfdf5; border: 1px solid #a7f3d0; }
        .state-title { font-size: 11px; font-weight: 700; text-transform: uppercase; color: var(--text-muted); margin-bottom: 4px; }
        .target-state .state-title { color: #065f46; }
        .state-value { font-family: monospace; font-weight: 600; word-break: break-all; }

        .info-section { background: #ffffff; border: 1px solid var(--border-color); border-radius: 6px; padding: 14px; margin-bottom: 14px; }
        .info-heading { font-size: 12px; font-weight: 700; color: #1e293b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
        .info-content { font-size: 13px; color: #334155; line-height: 1.5; }
        .benchmark-tag { display: inline-block; margin-top: 8px; font-size: 11px; color: var(--primary); font-weight: 600; }

        .remediation-section { background: #f8fafc; border: 1px solid var(--border-color); border-radius: 6px; padding: 14px; }
        .remediation-path { display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 600; color: #1e293b; margin-bottom: 10px; background: #ffffff; padding: 8px 12px; border-radius: 4px; border: 1px solid var(--border-color); }
        .cmd-box { position: relative; background: #0f172a; color: #f8fafc; border-radius: 6px; padding: 12px 16px; font-family: monospace; font-size: 12px; display: flex; justify-content: space-between; align-items: center; }
        .cmd-box code { overflow-x: auto; white-space: pre-wrap; word-break: break-all; margin-right: 12px; }
        .copy-btn { background: #334155; color: #ffffff; border: none; border-radius: 4px; padding: 4px 10px; font-size: 11px; font-weight: 600; cursor: pointer; transition: background 0.2s; white-space: nowrap; }
        .copy-btn:hover { background: var(--primary); }

        /* FOOTER */
        .footer { margin-top: 40px; padding: 20px; text-align: center; font-size: 12px; color: var(--text-muted); border-top: 1px solid var(--border-color); }
        .footer a { color: var(--primary); text-decoration: none; }
        .footer a:hover { text-decoration: underline; }

        /* UTILITY */
        .font-mono { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; }
        .text-xs { font-size: 12px; }
        .text-muted { color: var(--text-muted); }
        .text-emerald { color: #059669; }

        @media print {
            body { padding: 0; background: #ffffff; }
            .controls-card, .filter-group, .search-box, .copy-btn { display: none !important; }
            .card-body { display: block !important; }
            .recommendation-card { break-inside: avoid; border: 1px solid #cbd5e1 !important; margin-bottom: 16px; }
            .header-card { background: #1e293b !important; color: #000000 !important; }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- HEADER -->
        <div class="header-card">
            <div class="header-title">
                <h1>Raport Audytu Bezpieczeństwa Microsoft Entra ID</h1>
                <p>CIS Microsoft 365 Foundations Benchmark & Best Practices Security Assessment</p>
            </div>
            <div class="header-meta">
                <div class="header-meta-item"><b>Tenant:</b> $([System.Web.HttpUtility]::HtmlEncode($Script:ActiveTenantName))</div>
                <div class="header-meta-item"><b>Tenant ID:</b> $($Script:ActiveTenantId)</div>
                <div class="header-meta-item"><b>Data wykonania:</b> $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))</div>
                <div class="header-meta-item"><b>Wersja:</b> $($Script:Metadata.Version)</div>
            </div>
        </div>

        <!-- KPI SUMMARY CARDS -->
        <div class="dashboard-grid">
            <div class="kpi-card gauge-card">
                <svg class="gauge-svg" viewBox="0 0 140 140">
                    <circle class="gauge-bg" cx="70" cy="70" r="$radius"></circle>
                    <circle class="gauge-bar" cx="70" cy="70" r="$radius" transform="rotate(-90 70 70)"></circle>
                    <text class="gauge-text" x="70" y="74" text-anchor="middle">$compliancePercentage%</text>
                    <text class="gauge-subtext" x="70" y="90" text-anchor="middle">Zgodność</text>
                </svg>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Sprawdzone reguły</div>
                <div class="kpi-num">$totalRules</div>
                <div class="text-xs text-muted">Łączna liczba kryteriów</div>
            </div>
            <div class="kpi-card kpi-success">
                <div class="kpi-label">Zgodne (Compliant)</div>
                <div class="kpi-num">$compliantRules</div>
                <div class="text-xs text-muted">Poprawnie zabezpieczone</div>
            </div>
            <div class="kpi-card kpi-danger">
                <div class="kpi-label">Krytyczne / Wysokie</div>
                <div class="kpi-num">$($criticalIssues + $highIssues)</div>
                <div class="text-xs text-muted">$criticalIssues Krytyczne, $highIssues Wysokie</div>
            </div>
            <div class="kpi-card kpi-warning">
                <div class="kpi-label">Ostrzeżenia (Warnings)</div>
                <div class="kpi-num">$warningRules</div>
                <div class="text-xs text-muted">Wymaga weryfikacji</div>
            </div>
        </div>

        <!-- CONTROLS & FILTER BAR -->
        <div class="controls-card">
            <div class="filter-group">
                <button class="filter-btn active" onclick="applyStatusFilter('ALL', this)">Wszystkie ($totalRules)</button>
                <button class="filter-btn" onclick="applyStatusFilter('Non-Compliant', this)">Niezgodne ($nonCompliant)</button>
                <button class="filter-btn" onclick="applyStatusFilter('Compliant', this)">Zgodne ($compliantRules)</button>
                <button class="filter-btn" onclick="applyStatusFilter('Warning', this)">Ostrzeżenia ($warningRules)</button>
                <button class="filter-btn" onclick="toggleAllCards()">Rozwiń / Zwiń Wszystkie</button>
            </div>
            <div class="search-box">
                <input type="text" id="liveSearchInput" class="search-input" placeholder="Szukaj reguły, kategorii, słowa..." onkeyup="filterCardsBySearch()">
            </div>
        </div>

        <!-- SECTION 1: MACIERZ AKCJI -->
        <div class="section-header">
            <div class="section-title">
                <span>🎯</span> Macierz Podsumowująca Działania (Action Matrix)
            </div>
            <div class="section-subtitle">Posortowane wg priorytetu ryzyka i typu wdrożenia (Quick Wins vs Zmiany Architektoniczne)</div>
        </div>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Priorytet</th>
                        <th>Typ Akcji</th>
                        <th>ID</th>
                        <th>Reguła Bezpieczeństwa</th>
                        <th>Status</th>
                        <th>Stan Obecny</th>
                        <th>Rekomendacja</th>
                    </tr>
                </thead>
                <tbody id="matrixTableBody">
                    $($matrixRowsHtml.ToString())
                </tbody>
            </table>
        </div>

        <!-- SECTION 2: KARTY REKOMENDACJI -->
        <div class="section-header">
            <div class="section-title">
                <span>📋</span> Szczegółowe Karty Rekomendacji i Naprawy
            </div>
            <div class="section-subtitle">Uzasadnienie biznesowe, analiza wektorów ataku oraz ścieżki naprawcze w portalu Entra ID i PowerShell</div>
        </div>
        <div class="cards-list" id="recommendationCardsContainer">
            $($cardsHtml.ToString())
        </div>

        <!-- SECTION 3: INWENTARYZACJA KONT UPRZYWILEJOWANYCH -->
        <div class="section-header">
            <div class="section-title">
                <span>🛡️</span> Inwentaryzacja Kont Uprzywilejowanych i Status MFA
            </div>
            <div class="section-subtitle">Global Administrator, Security Administrator, Privileged Role Administrator</div>
        </div>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Użytkownik (Display Name)</th>
                        <th>User Principal Name (UPN)</th>
                        <th>Rola Katalogowa</th>
                        <th>Rejestracja MFA</th>
                        <th>Konto Aktywne</th>
                    </tr>
                </thead>
                <tbody>
                    $($privUsersHtml.ToString())
                </tbody>
            </table>
        </div>

        <!-- FOOTER -->
        <div class="footer">
            <p>Raport wygenerowany przez narzędzie <b>$($Script:Metadata.Title)</b> v$($Script:Metadata.Version) | Autor: <a href="$($Script:Metadata.GitHub)" target="_blank">$($Script:Metadata.Author)</a> &lt;$($Script:Metadata.Email)&gt;</p>
            <p style="margin-top: 4px;">Raport typu Read-Only. Bezpieczeństwo i audyt dzierżawy Microsoft Entra ID zgodny ze standardami CIS Microsoft 365 Foundations Benchmark.</p>
        </div>
    </div>

    <!-- CLIENT-SIDE INTERACTION SCRIPT -->
    <script>
        let currentStatusFilter = 'ALL';

        function toggleCard(headerElement) {
            const card = headerElement.closest('.recommendation-card');
            card.classList.toggle('open');
        }

        let allExpanded = false;
        function toggleAllCards() {
            allExpanded = !allExpanded;
            document.querySelectorAll('.recommendation-card').forEach(card => {
                if (allExpanded) { card.classList.add('open'); }
                else { card.classList.remove('open'); }
            });
        }

        function applyStatusFilter(status, btnElement) {
            currentStatusFilter = status;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            if (btnElement) btnElement.classList.add('active');
            filterCardsBySearch();
        }

        function filterCardsBySearch() {
            const searchVal = document.getElementById('liveSearchInput').value.toLowerCase().trim();
            const cards = document.querySelectorAll('.recommendation-card');
            const rows = document.querySelectorAll('.action-row');

            cards.forEach(card => {
                const cardStatus = card.getAttribute('data-status');
                const cardKeywords = card.getAttribute('data-keywords').toLowerCase();

                const statusMatch = (currentStatusFilter === 'ALL' || cardStatus === currentStatusFilter);
                const searchMatch = (!searchVal || cardKeywords.includes(searchVal));

                if (statusMatch && searchMatch) {
                    card.style.display = 'block';
                } else {
                    card.style.display = 'none';
                }
            });

            rows.forEach(row => {
                const rowStatus = row.getAttribute('data-status');
                const rowText = row.innerText.toLowerCase();

                const statusMatch = (currentStatusFilter === 'ALL' || rowStatus === currentStatusFilter);
                const searchMatch = (!searchVal || rowText.includes(searchVal));

                if (statusMatch && searchMatch) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }

        function copyToClipboard(button) {
            const codeEl = button.previousElementSibling;
            if (!codeEl) return;
            navigator.clipboard.writeText(codeEl.innerText).then(() => {
                const originalText = button.innerText;
                button.innerText = 'Skopiowano!';
                button.style.background = '#10b981';
                setTimeout(() => {
                    button.innerText = originalText;
                    button.style.background = '';
                }, 2000);
            }).catch(err => {
                alert('Błąd kopiowania: ' + err);
            });
        }
    </script>
</body>
</html>
"@

        try {
            Set-Content -Path $Script:GeneratedHtmlPath -Value $htmlContent -Encoding UTF8 -Force
            Write-AuditLog -Message "Interactive HTML report generated successfully: $($Script:GeneratedHtmlPath)" -Level 'SUCCESS' -Component 'REPORT'
        }
        catch {
            Write-AuditLog -Message "Failed to write HTML report: $($_.Exception.Message)" -Level 'ERROR' -Component 'REPORT'
            throw $_
        }
    }

    Export-EntraIDSecurityHtmlReport

    # Optional JSON Export
    if ($ExportJson) {
        $jsonTimestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $Script:GeneratedJsonPath = Join-Path -Path $OutputPath -ChildPath "AuditReport-EntraID-$jsonTimestamp.json"

        $jsonPayload = [PSCustomObject]@{
            AuditMetadata = @{
                ScriptVersion    = $Script:Metadata.Version
                AuditTimestamp   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                TenantId         = $Script:ActiveTenantId
                TenantName       = $Script:ActiveTenantName
                ComplianceScore  = $compliancePercentage
                TotalRules       = $totalRules
                CompliantCount   = $compliantRules
                NonCompliantCount= $nonCompliant
                WarningCount     = $warningRules
                CriticalIssues   = $criticalIssues
                HighIssues       = $highIssues
                MediumIssues     = $mediumIssues
                LowIssues        = $lowIssues
            }
            SecurityFindings = $Script:AuditFindings
            PrivilegedUsers  = $Script:PrivilegedUsers
        }

        try {
            $jsonPayload | ConvertTo-Json -Depth 5 | Set-Content -Path $Script:GeneratedJsonPath -Encoding UTF8 -Force
            Write-AuditLog -Message "Structured JSON report generated: $($Script:GeneratedJsonPath)" -Level 'SUCCESS' -Component 'JSON'
        }
        catch {
            Write-AuditLog -Message "Failed to export JSON report: $($_.Exception.Message)" -Level 'WARN' -Component 'JSON'
        }
    }

    # ==============================================================================
    # 8. CONSOLE EXECUTIVE SUMMARY OUTPUT
    # ==============================================================================
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "                MICROSOFT ENTRA ID SECURITY AUDIT SUMMARY                      " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host " Tenant Name         : $($Script:ActiveTenantName)" -ForegroundColor White
    Write-Host " Tenant ID           : $($Script:ActiveTenantId)" -ForegroundColor White
    Write-Host " Audit Timestamp     : $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor White
    Write-Host " Compliance Score    : " -NoNewline -ForegroundColor White

    $scoreConsoleColor = if ($compliancePercentage -ge 80) { 'Green' } elseif ($compliancePercentage -ge 60) { 'Yellow' } else { 'Red' }
    Write-Host "$compliancePercentage% ($compliantRules / $totalRules Compliant)" -ForegroundColor $scoreConsoleColor

    Write-Host " Critical Findings   : $criticalIssues" -ForegroundColor $(if ($criticalIssues -gt 0) { 'Red' } else { 'Green' })
    Write-Host " High Severity Risks : $highIssues" -ForegroundColor $(if ($highIssues -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host " Warnings            : $warningRules" -ForegroundColor Yellow
    Write-Host " HTML Report File    : $($Script:GeneratedHtmlPath)" -ForegroundColor Cyan
    if ($Script:GeneratedJsonPath) {
        Write-Host " JSON Report File    : $($Script:GeneratedJsonPath)" -ForegroundColor Cyan
    }
    if ($Script:LogFilePath) {
        Write-Host " Execution Log File  : $($Script:LogFilePath)" -ForegroundColor Gray
    }
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""

    # Open Report if requested
    if ($OpenReport -and $Script:GeneratedHtmlPath -and (Test-Path -Path $Script:GeneratedHtmlPath)) {
        Write-AuditLog -Message "Launching interactive HTML report in default browser..." -Level 'INFO' -Component 'REPORT'
        Start-Process -FilePath $Script:GeneratedHtmlPath
    }

    # Pipeline output if requested
    if ($PassThru) {
        Write-Output $Script:AuditFindings
    }
}

end {
    $totalElapsed = [math]::Round(((Get-Date) - $Script:StartTime).TotalSeconds, 2)
    Write-AuditLog -Message "Execution completed in $totalElapsed seconds." -Level 'INFO' -Component 'END'
}
