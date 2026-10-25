<#
.SYNOPSIS
    Get-SharePointSitesAndUsers - Audits SharePoint Online site collections and aggregates user roles and permissions.

.DESCRIPTION
    Get-SharePointSitesAndUsers is an enterprise-grade PowerShell automation utility engineered for
    Microsoft 365 administrators, security engineers, and compliance auditors.

    The script performs comprehensive inventory and permissions inspection across SharePoint Online:
    - Supports tenant-wide bulk auditing (-All) or targeted single-site collection auditing (-Site <Url>).
    - Analyzes and standardizes permissions into four standardized logical tiers:
        1. SiteCollectionAdmins : Site Collection Administrators (unconditional ownership).
        2. FullControl          : Owners group members, Full Control direct/group role assignments, M365 Group owners.
        3. Write                : Members group members, Edit/Contribute role assignments, M365 Group members.
        4. Read                 : Visitors group members, Read/View Only role assignments.
    - Resolves nested group claims, Entra ID security groups, and Microsoft 365 Groups.
    - Emits strongly-typed [PSCustomObject] records directly to the PowerShell pipeline for downstream
      processing (e.g., Export-Csv, Out-GridView, Format-Table).
    - Robust resilience: catches 401/403 access denied errors on individual sites without aborting the audit.
    - Handles API rate-limiting and HTTP 429/503 throttling with exponential backoff retries.
    - Real-time progress tracking with Write-Progress in bulk mode.
    - Enterprise-grade synchronous file logging in C:\Logs\Get-SharePointSitesAndUsers\ with automatic
      fallback to $env:LOCALAPPDATA\Logs\Get-SharePointSitesAndUsers.
    - Backward compatible with Windows PowerShell 5.1 with automatic elevation to PowerShell 7+ (pwsh.exe).
    - Dual module support: PnP.PowerShell (primary) and Microsoft.Online.SharePoint.PowerShell (fallback).

.PARAMETER All
    Switch parameter. Triggers bulk discovery and security auditing across all tenant site collections.
    Mutually exclusive with -Site.

.PARAMETER Site
    Absolute URL of a specific SharePoint Online site collection to audit
    (e.g., 'https://contoso.sharepoint.com/sites/Marketing').
    Supports pipeline input by value and property name.
    Aliases: -Url, -SiteUrl.
    Mutually exclusive with -All.

.PARAMETER AdminCenterUrl
    Optional SharePoint Online Admin Center URL (e.g., 'https://contoso-admin.sharepoint.com').
    If omitted, the script dynamically derives it from the target site URL, current session, or tenant domain.

.PARAMETER LogPath
    Directory path for audit execution logs.
    Defaults to 'C:\Logs\Get-SharePointSitesAndUsers'.

.PARAMETER IncludePersonalSites
    Switch parameter. When specified, OneDrive for Business personal sites ('/personal/') are included.
    By default, personal sites are excluded to focus on organizational SharePoint sites.

.PARAMETER ClientId
    Microsoft Entra ID Application (Client) ID for modern authentication.
    If omitted, checks local cache ($env:LOCALAPPDATA\SharePoint-Audit) or environment variable $env:ENTRAID_APP_ID.
    Aliases: -AppId.

.PARAMETER TenantId
    Microsoft Entra ID Directory (Tenant) ID or domain name (e.g., 'contoso.onmicrosoft.com').
    Aliases: -Tenant.

.PARAMETER CertificateThumbprint
    Thumbprint of an X.509 certificate in the local certificate store (Cert:\CurrentUser\My or Cert:\LocalMachine\My)
    for unattended daemon authentication.
    Aliases: -Thumbprint.

.PARAMETER CertificatePath
    File path to an X.509 certificate file (.pfx or .cer) for App-Only authentication.

.PARAMETER CertificatePassword
    Optional SecureString password for the certificate file.

.PARAMETER Interactive
    Switch parameter. Forces interactive browser authentication with MFA support.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts.

.PARAMETER Help
    Displays built-in help, usage examples, parameter descriptions, and exits.
    Aliases: -h.

.EXAMPLE
    # 1. Audit a single SharePoint Online site collection
    .\Get-SharePointSitesAndUsers.ps1 -Site "https://contoso.sharepoint.com/sites/Marketing"

.EXAMPLE
    # 2. Audit all site collections across the tenant and export to CSV
    .\Get-SharePointSitesAndUsers.ps1 -All | Export-Csv -Path "C:\Reports\SharePointAudit.csv" -NoTypeInformation

.EXAMPLE
    # 3. Pipeline input from URL string with interactive Out-GridView
    "https://contoso.sharepoint.com/sites/Finance" | .\Get-SharePointSitesAndUsers.ps1 | Out-GridView

.EXAMPLE
    # 4. Bulk tenant audit including OneDrive personal sites with custom admin endpoint
    .\Get-SharePointSitesAndUsers.ps1 -All -AdminCenterUrl "https://contoso-admin.sharepoint.com" -IncludePersonalSites

.EXAMPLE
    # 5. Automated unattended daemon execution using App-Only certificate authentication
    .\Get-SharePointSitesAndUsers.ps1 -All `
        -AdminCenterUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "contoso.onmicrosoft.com" `
        -CertificateThumbprint "A1B2C3D4E5F67890123456789ABCDEF012345678" |
        Export-Csv -Path "C:\Reports\TenantAudit.csv" -NoTypeInformation

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
    Last Modified: 2026-10-25
    Standards: CleanCode & PSScriptAnalyzer Zero-Warning Standard
#>

[CmdletBinding(DefaultParameterSetName = 'AllSites')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'AllSites', Position = 0)]
    [switch]$All,

    [Parameter(Mandatory = $true, ParameterSetName = 'SingleSite', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [Alias('Url', 'SiteUrl')]
    [string]$Site,

    [Parameter(Mandatory = $false)]
    [string]$AdminCenterUrl,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\Get-SharePointSitesAndUsers",

    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalSites,

    [Parameter(Mandatory = $false)]
    [Alias('AppId')]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [Alias('Tenant')]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [Alias('Thumbprint')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $true, ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # ==============================================================================
    # SCRIPT METADATA & RUNTIME TELEMETRY
    # ==============================================================================
    $Script:Metadata = @{
        Title        = "Get-SharePointSitesAndUsers.ps1"
        CommandName  = "Get-SharePointSitesAndUsers"
        Version      = "1.0.0"
        Author       = "Roman Pindela"
        Email        = "roman.pindela@gmail.com"
        GitHub       = "https://github.com/romanpindela"
        LastModified = "2026-10-25"
        PrimaryMod   = "PnP.PowerShell"
        FallbackMod  = "Microsoft.Online.SharePoint.PowerShell"
    }

    $Script:StartTime        = Get-Date
    $Script:LogFilePath      = $null
    $Script:ActiveConnection = $null
    $Script:ActiveClientId   = $null
    $Script:ActiveTenant     = $null
    $Script:EngineMode       = $null
    $Script:TotalSites       = 0
    $Script:ProcessedSites   = 0
    $Script:FailedSites      = 0

    $Script:ScriptFilePath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $PSCommandPath
    }
    elseif ($MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
        $MyInvocation.MyCommand.Path
    }
    else {
        $null
    }

    # Register Global Alias for convenient invocation
    if (-not [string]::IsNullOrWhiteSpace($Script:ScriptFilePath)) {
        Set-Alias -Name 'Get-SharePointSitesAndUsers' -Value $Script:ScriptFilePath -Scope Global -ErrorAction SilentlyContinue
    }

    # ==============================================================================
    # SHOW-SCRIPTHELP
    # ==============================================================================
    function Show-ScriptHelp {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host " SCRIPT  : $($Script:Metadata.Title) (v$($Script:Metadata.Version))" -ForegroundColor Cyan
        Write-Host " AUTHOR  : $($Script:Metadata.Author) <$($Script:Metadata.Email)>" -ForegroundColor Cyan
        Write-Host " GITHUB  : $($Script:Metadata.GitHub)" -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "`nDESCRIPTION:" -ForegroundColor White
        Write-Host "    Audits SharePoint Online site collections and aggregates user roles and permissions." -ForegroundColor Gray
        Write-Host "    Maps entitled users to 4 standardized tiers: SiteCollectionAdmins, FullControl, Write, Read." -ForegroundColor Gray
        Write-Host "    Emits structured [PSCustomObject] records directly to the pipeline for CSV/Grid export." -ForegroundColor Gray
        Write-Host "`nPARAMETER SETS:" -ForegroundColor Yellow
        Write-Host "    [AllSites]   : Get-SharePointSitesAndUsers.ps1 -All [-AdminCenterUrl <Url>] [-IncludePersonalSites] ..." -ForegroundColor Gray
        Write-Host "    [SingleSite] : Get-SharePointSitesAndUsers.ps1 -Site <Url> [-AdminCenterUrl <Url>] ..." -ForegroundColor Gray
        Write-Host "    [Help]       : Get-SharePointSitesAndUsers.ps1 -Help" -ForegroundColor Gray
        Write-Host "`nPARAMETERS:" -ForegroundColor Yellow
        Write-Host "    -All                   [Switch]  Process all site collections in the tenant (excludes /personal/ by default)." -ForegroundColor Gray
        Write-Host "    -Site (-Url)           [String]  Target site collection URL (e.g. 'https://contoso.sharepoint.com/sites/hr')." -ForegroundColor Gray
        Write-Host "    -AdminCenterUrl        [String]  SharePoint Admin URL (e.g. 'https://contoso-admin.sharepoint.com')." -ForegroundColor Gray
        Write-Host "    -LogPath               [String]  Directory for transcript logs (Default: 'C:\Logs\Get-SharePointSitesAndUsers')." -ForegroundColor Gray
        Write-Host "    -IncludePersonalSites  [Switch]  Include OneDrive for Business personal sites." -ForegroundColor Gray
        Write-Host "    -ClientId (-AppId)     [String]  Microsoft Entra ID Application (Client) ID." -ForegroundColor Gray
        Write-Host "    -TenantId (-Tenant)    [String]  Microsoft Entra ID Tenant ID or domain (contoso.onmicrosoft.com)." -ForegroundColor Gray
        Write-Host "    -CertificateThumbprint [String]  X.509 Certificate thumbprint for unattended daemon auth." -ForegroundColor Gray
        Write-Host "    -CertificatePath       [String]  Path to certificate (.pfx/.cer) file." -ForegroundColor Gray
        Write-Host "    -Interactive           [Switch]  Initiate interactive browser OAuth authentication with MFA." -ForegroundColor Gray
        Write-Host "    -Force                 [Switch]  Bypass confirmation prompts." -ForegroundColor Gray
        Write-Host "    -Help (-h)             [Switch]  Display this help screen and exit." -ForegroundColor Gray
        Write-Host "`nEXAMPLES:" -ForegroundColor Yellow
        Write-Host "    .\Get-SharePointSitesAndUsers.ps1 -Site 'https://contoso.sharepoint.com/sites/Marketing'" -ForegroundColor Green
        Write-Host "    .\Get-SharePointSitesAndUsers.ps1 -All | Export-Csv -Path 'SharePointAudit.csv' -NoTypeInformation" -ForegroundColor Green
        Write-Host "    'https://contoso.sharepoint.com/sites/Finance' | .\Get-SharePointSitesAndUsers.ps1 | Out-GridView" -ForegroundColor Green
        Write-Host "    .\Get-SharePointSitesAndUsers.ps1 -All -IncludePersonalSites | Format-Table -AutoSize" -ForegroundColor Green
        Write-Host "================================================================================`n" -ForegroundColor Cyan
    }

    if ($Help) {
        Show-ScriptHelp
        exit 0
    }

    # ==============================================================================
    # POWERSHELL VERSION CHECK & AUTOMATIC RELAUNCH IN PWSH (PowerShell 7+)
    # ==============================================================================
    function Assert-PowerShellVersion {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $false)]
            [System.Collections.IDictionary]$BoundParams
        )

        if ($PSVersionTable.PSVersion.Major -lt 7) {
            $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($null -ne $pwsh -and (-not [string]::IsNullOrWhiteSpace($Script:ScriptFilePath))) {
                Write-Host "[*] PnP.PowerShell performs optimally on PowerShell 7+ (Core). Current host: Windows PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
                Write-Host "[*] Automatically relaunching script in PowerShell 7 (pwsh.exe)...`n" -ForegroundColor Yellow

                $forwardArgs = [System.Collections.Generic.List[string]]::new()
                $escapedScriptPath = $Script:ScriptFilePath.Replace('"', '`"')
                $forwardArgs.Add("& `"$escapedScriptPath`"")

                if ($null -ne $BoundParams) {
                    foreach ($key in $BoundParams.Keys) {
                        $val = $BoundParams[$key]
                        if ($val -is [switch] -or $val -is [System.Management.Automation.SwitchParameter]) {
                            if ($val.IsPresent) {
                                $forwardArgs.Add("-$key")
                            }
                        }
                        elseif ($val -is [bool]) {
                            $forwardArgs.Add("-$key")
                            $forwardArgs.Add("`$$val")
                        }
                        else {
                            $forwardArgs.Add("-$key")
                            $escapedVal = $val.ToString().Replace('"', '`"')
                            $forwardArgs.Add("`"$escapedVal`"")
                        }
                    }
                }

                $cmdBlock = $forwardArgs -join " "
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($cmdBlock)
                $encoded = [System.Convert]::ToBase64String($bytes)
                $process = Start-Process -FilePath $pwsh.Source -ArgumentList "-NoProfile -EncodedCommand $encoded" -Wait -PassThru -NoNewWindow
                exit $process.ExitCode
            }
        }
    }

    # ==============================================================================
    # LOGGING INFRASTRUCTURE (Specification Section 5)
    # ==============================================================================
    function Initialize-AuditLog {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $false)]
            [string]$Directory = 'C:\Logs\Get-SharePointSitesAndUsers'
        )

        $timestamp    = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $logFileName  = "Get-SharePointSitesAndUsers_${timestamp}.log"
        $resolvedDir  = $Directory

        try {
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop
            }
        }
        catch {
            $fallbackDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs\Get-SharePointSitesAndUsers"
            Write-Host "[-] Notice: Target log directory '$resolvedDir' is inaccessible ($($_.Exception.Message)). Falling back to '$fallbackDir'." -ForegroundColor Yellow
            $resolvedDir = $fallbackDir
            if (-not (Test-Path -Path $resolvedDir)) {
                $null = New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
        }

        $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logFileName

        $header = @"
================================================================================
Audit Log Initialized : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script                : $($Script:Metadata.Title) v$($Script:Metadata.Version)
Author                : $($Script:Metadata.Author) <$($Script:Metadata.Email)>
GitHub                : $($Script:Metadata.GitHub)
Host Machine          : $env:COMPUTERNAME
User Context          : $env:USERNAME
Log File Location     : $Script:LogFilePath
================================================================================
"@
        try {
            Add-Content -Path $Script:LogFilePath -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
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
            [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG', 'SUCCESS')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry  = "[$timestamp] [$Level] $Message"

        if ($null -ne $Script:LogFilePath) {
            try {
                Add-Content -Path $Script:LogFilePath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            catch {
                $null = $_
            }
        }

        switch ($Level) {
            'INFO'    { Write-Host $logEntry -ForegroundColor Cyan }
            'WARN'    { Write-Host $logEntry -ForegroundColor Yellow }
            'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
            'DEBUG'   { Write-Host $logEntry -ForegroundColor DarkGray }
            'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        }
    }

    # ==============================================================================
    # RESILIENCE & THROTTLING RETRY HANDLER (Specification Section 7)
    # ==============================================================================
    function Invoke-WithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [scriptblock]$ScriptBlock,

            [Parameter(Mandatory = $false)]
            [int]$MaxRetries = 3,

            [Parameter(Mandatory = $false)]
            [int]$InitialDelaySeconds = 4
        )

        $attempt = 0
        $delay = $InitialDelaySeconds

        while ($attempt -lt $MaxRetries) {
            $attempt++
            try {
                return (& $ScriptBlock)
            }
            catch {
                $errMsg = $_.Exception.Message
                $isThrottled = ($errMsg -match '429' -or
                                $errMsg -match 'throttle' -or
                                $errMsg -match 'Too Many Requests' -or
                                $errMsg -match '503' -or
                                $errMsg -match 'Server Too Busy')

                if ($isThrottled -and $attempt -lt $MaxRetries) {
                    Write-AuditLog -Level WARN -Message "HTTP throttling encountered (429/503). Retrying attempt $attempt of $MaxRetries after $delay seconds..."
                    Start-Sleep -Seconds $delay
                    $delay = $delay * 2
                }
                else {
                    throw
                }
            }
        }
    }

    # ==============================================================================
    # SAFE PROPERTY RETRIEVAL (StrictMode Resilient)
    # ==============================================================================
    function Get-SafePropertyValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            $InputObject,

            [Parameter(Mandatory = $true)]
            [string]$PropertyName,

            [Parameter(Mandatory = $false)]
            $DefaultValue = $null
        )

        if ($null -eq $InputObject) {
            return $DefaultValue
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            if ($InputObject.Contains($PropertyName)) {
                return $InputObject[$PropertyName]
            }
            return $DefaultValue
        }

        if ($InputObject.PSObject -and $InputObject.PSObject.Properties) {
            $prop = $InputObject.PSObject.Properties[$PropertyName]
            if ($null -ne $prop) {
                return $prop.Value
            }
        }

        return $DefaultValue
    }

    # ==============================================================================
    # IDENTITY EXTRACTION & SANITIZATION (Specification Section 3 & 6)
    # ==============================================================================
    function Get-NormalizedIdentity {
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory = $false)]
            $Principal
        )

        if ($null -eq $Principal) { return "" }

        $email = ""
        $login = ""
        $title = ""

        if ($Principal -is [string]) {
            $login = $Principal
        }
        elseif ($Principal.PSObject) {
            if ($Principal.PSObject.Properties['Email'] -and -not [string]::IsNullOrWhiteSpace($Principal.Email)) {
                $email = [string]$Principal.Email
            }
            elseif ($Principal.PSObject.Properties['Mail'] -and -not [string]::IsNullOrWhiteSpace($Principal.Mail)) {
                $email = [string]$Principal.Mail
            }
            elseif ($Principal.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($Principal.UserPrincipalName)) {
                $email = [string]$Principal.UserPrincipalName
            }

            if ($Principal.PSObject.Properties['LoginName'] -and -not [string]::IsNullOrWhiteSpace($Principal.LoginName)) {
                $login = [string]$Principal.LoginName
            }
            if ($Principal.PSObject.Properties['Title'] -and -not [string]::IsNullOrWhiteSpace($Principal.Title)) {
                $title = [string]$Principal.Title
            }
            elseif ($Principal.PSObject.Properties['DisplayName'] -and -not [string]::IsNullOrWhiteSpace($Principal.DisplayName)) {
                $title = [string]$Principal.DisplayName
            }
        }

        # Resolve claim identities:
        # e.g., 'i:0#.f|membership|user@contoso.com' -> 'user@contoso.com'
        if ([string]::IsNullOrWhiteSpace($email) -and -not [string]::IsNullOrWhiteSpace($login)) {
            if ($login -match '\|membership\|([^|#]+)$') {
                $email = $Matches[1]
            }
            elseif ($login -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                $email = $login
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($email)) {
            return $email.Trim().ToLowerInvariant()
        }
        if (-not [string]::IsNullOrWhiteSpace($title)) {
            return $title.Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace($login)) {
            if ($login -match '\|([^|#]+)$') {
                return $Matches[1].Trim()
            }
            return $login.Trim()
        }

        return ""
    }

    # ==============================================================================
    # NESTED ENTRA ID / MICROSOFT 365 GROUP MEMBERSHIP EXPANSION
    # ==============================================================================
    function Get-NestedGroupUser {
        [CmdletBinding()]
        [OutputType([System.Collections.Generic.List[string]])]
        param(
            [Parameter(Mandatory = $false)]
            [string]$GroupLoginName = "",

            [Parameter(Mandatory = $false)]
            [string]$GroupTitle = "",

            [Parameter(Mandatory = $false)]
            [string]$ConnectedSiteGroupId = "",

            [Parameter(Mandatory = $false)]
            $Connection = $null
        )

        $resolved = [System.Collections.Generic.List[string]]::new()
        $guidPattern = '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})'

        $targetGroupId = $null
        if (-not [string]::IsNullOrWhiteSpace($GroupLoginName) -and $GroupLoginName -match $guidPattern) {
            $targetGroupId = $Matches[1]
        }
        elseif (-not [string]::IsNullOrWhiteSpace($ConnectedSiteGroupId) -and $ConnectedSiteGroupId -ne "00000000-0000-0000-0000-000000000000") {
            $targetGroupId = $ConnectedSiteGroupId
        }

        if ([string]::IsNullOrWhiteSpace($targetGroupId)) {
            return $resolved
        }

        $isOwner = ($GroupLoginName -like "*_o*" -or
                    $GroupTitle -like "*Owner*" -or
                    $GroupTitle -like "*Właściciel*" -or
                    $GroupTitle -like "*Wlasciciel*")

        $rawUsers = $null
        try {
            if ($isOwner) {
                $rawUsers = if ($null -ne $Connection) {
                    Get-PnPEntraIDGroupOwner -Identity $targetGroupId -Connection $Connection -ErrorAction Stop
                } else {
                    Get-PnPEntraIDGroupOwner -Identity $targetGroupId -ErrorAction Stop
                }
            } else {
                $rawUsers = if ($null -ne $Connection) {
                    Get-PnPEntraIDGroupMember -Identity $targetGroupId -Transitive -Connection $Connection -ErrorAction Stop
                } else {
                    Get-PnPEntraIDGroupMember -Identity $targetGroupId -Transitive -ErrorAction Stop
                }
            }
        }
        catch {
            try {
                if ($isOwner) {
                    $rawUsers = if ($null -ne $Connection) {
                        Get-PnPMicrosoft365GroupOwner -Identity $targetGroupId -Connection $Connection -ErrorAction Stop
                    } else {
                        Get-PnPMicrosoft365GroupOwner -Identity $targetGroupId -ErrorAction Stop
                    }
                } else {
                    $rawUsers = if ($null -ne $Connection) {
                        Get-PnPMicrosoft365GroupMember -Identity $targetGroupId -Connection $Connection -ErrorAction Stop
                    } else {
                        Get-PnPMicrosoft365GroupMember -Identity $targetGroupId -ErrorAction Stop
                    }
                }
            }
            catch {
                Write-AuditLog -Level DEBUG -Message "Nested group expansion skipped for '$GroupTitle' ($targetGroupId): $($_.Exception.Message)"
            }
        }

        if ($null -ne $rawUsers) {
            foreach ($u in @($rawUsers)) {
                $ident = Get-NormalizedIdentity -Principal $u
                if (-not [string]::IsNullOrWhiteSpace($ident)) {
                    $resolved.Add($ident)
                }
            }
        }

        return $resolved
    }

    # ==============================================================================
    # URL & TENANT RESOLUTION
    # ==============================================================================
    function Resolve-AdminCenterUrlInternal {
        [CmdletBinding()]
        [OutputType([string])]
        param(
            [Parameter(Mandatory = $false)]
            [string]$ExplicitAdminUrl,

            [Parameter(Mandatory = $false)]
            [string]$TargetSiteUrl,

            [Parameter(Mandatory = $false)]
            [string]$TenantDomain
        )

        if (-not [string]::IsNullOrWhiteSpace($ExplicitAdminUrl)) {
            $trimmed = $ExplicitAdminUrl.Trim().TrimEnd('/')
            if ($trimmed -notmatch '^https?://') {
                $trimmed = "https://$trimmed"
            }
            try {
                $uri = [System.Uri]$trimmed
                $hostName = $uri.Host
                if ($hostName -notmatch '-admin\.sharepoint\.com$' -and $hostName -match '^([a-zA-Z0-9\-]+)\.(sharepoint|onmicrosoft)\.com$') {
                    $tenantName = $Matches[1].Replace('-my', '')
                    $trimmed = "https://${tenantName}-admin.sharepoint.com"
                    Write-AuditLog -Level INFO -Message "Normalized tenant URL to SharePoint Admin URL: $trimmed"
                }
                elseif ($hostName -notmatch '\.') {
                    $trimmed = "https://${hostName}-admin.sharepoint.com"
                    Write-AuditLog -Level INFO -Message "Normalized tenant name to SharePoint Admin URL: $trimmed"
                }
            }
            catch {
                $null = $_
            }
            return $trimmed
        }

        if (-not [string]::IsNullOrWhiteSpace($TargetSiteUrl)) {
            $clean = $TargetSiteUrl.Trim().TrimEnd('/')
            if ($clean -notmatch '^https?://') {
                $clean = "https://$clean"
            }
            try {
                $uri = [System.Uri]$clean
                $hostName = $uri.Host
                if ($hostName -match '^([a-zA-Z0-9\-]+)\.sharepoint\.com$') {
                    $prefix = $Matches[1].Replace('-my', '').Replace('-admin', '')
                    return "https://${prefix}-admin.sharepoint.com"
                }
            }
            catch {
                $null = $_
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($TenantDomain)) {
            $prefix = $TenantDomain.Trim().ToLowerInvariant().Replace('.sharepoint.com', '').Replace('.onmicrosoft.com', '').Replace('-admin', '')
            return "https://${prefix}-admin.sharepoint.com"
        }

        # Check existing PnP Connection URL
        try {
            $currentConn = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($null -ne $currentConn -and -not [string]::IsNullOrWhiteSpace($currentConn.Url)) {
                $connUri = [System.Uri]$currentConn.Url
                if ($connUri.Host -match '^([a-zA-Z0-9\-]+)\.sharepoint\.com$') {
                    $prefix = $Matches[1].Replace('-my', '').Replace('-admin', '')
                    return "https://${prefix}-admin.sharepoint.com"
                }
            }
        }
        catch {
            $null = $_
        }

        return ""
    }

    # ==============================================================================
    # INTERACTIVE PROMPT HELPER
    # ==============================================================================
    function Read-AdminCenterUrlInteractive {
        [CmdletBinding()]
        [OutputType([string])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param()

        Write-Host "`n[?] Enter SharePoint Admin Center URL (e.g. 'https://contoso-admin.sharepoint.com'): " -NoNewline -ForegroundColor Yellow
        $entered = Read-Host
        return $entered
    }

    # ==============================================================================
    # AUTOMATED ENTRA ID APP REGISTRATION (Post-2024 Modern PnP Compliance)
    # ==============================================================================
    function Invoke-EntraIDAppRegistration {
        [CmdletBinding()]
        [OutputType([string])]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Tenant,

            [Parameter(Mandatory = $false)]
            [string]$ApplicationName = "SharePoint-Audit"
        )

        $targetAppName = $ApplicationName
        $maxAttempts   = 3
        $attempt       = 0

        while ($attempt -lt $maxAttempts) {
            $attempt++

            Write-AuditLog -Level INFO -Message "Initiating automated Microsoft Entra ID App registration ('$targetAppName') in tenant '$Tenant'..."
            Write-Host "`n[*] Microsoft Entra ID registration requires a one-time administrator consent flow." -ForegroundColor Yellow
            Write-Host "[*] A browser window will now open for modern interactive sign-in." -ForegroundColor Yellow
            Write-Host "[*] Please sign in with your Microsoft 365 account and grant consent to the requested SharePoint permissions.`n" -ForegroundColor Yellow

            try {
                $regResult = Register-PnPEntraIDAppForInteractiveLogin `
                    -ApplicationName $targetAppName `
                    -Tenant $Tenant `
                    -SharePointDelegatePermissions "AllSites.FullControl" `
                    -GraphDelegatePermissions "Group.Read.All", "User.Read.All" `
                    -ErrorAction Stop

                $newAppId = $null
                if ($null -ne $regResult) {
                    foreach ($item in @($regResult)) {
                        if ($null -eq $item) { continue }
                        $itemStr = $item.ToString()
                        if ($itemStr -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                            $newAppId = $Matches[1]
                            break
                        }
                        if ($item.PSObject) {
                            foreach ($prop in $item.PSObject.Properties) {
                                if ($prop.Value -and $prop.Value.ToString() -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                                    $newAppId = $Matches[1]
                                    break
                                }
                            }
                            if ($newAppId) { break }
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($newAppId)) {
                        $allText = ($regResult | Out-String)
                        if ($allText -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                            $newAppId = $Matches[1]
                        }
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($newAppId)) {
                    Write-AuditLog -Level SUCCESS -Message "Entra ID App '$targetAppName' registered successfully. Application (Client) ID: $newAppId"

                    # Persist to local cache so subsequent runs never need prompt or re-registration
                    try {
                        $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                        if (-not (Test-Path -Path $cacheDir)) {
                            $null = New-Item -Path $cacheDir -ItemType Directory -Force
                        }
                        $cacheFile = Join-Path -Path $cacheDir -ChildPath "${Tenant}_client_id.txt"
                        [System.IO.File]::WriteAllText($cacheFile, $newAppId)
                        Write-AuditLog -Level SUCCESS -Message "Application ID cached locally at: $cacheFile"
                    }
                    catch {
                        Write-AuditLog -Level WARN -Message "Unable to persist Application ID to local cache: $($_.Exception.Message)"
                    }

                    return $newAppId
                }
                else {
                    throw "Could not extract Application (Client) ID from registration output."
                }
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-AuditLog -Level WARN -Message "Error during Microsoft Entra ID app registration: $errMsg"

                if ([Environment]::UserInteractive) {
                    $suggestedName = "SharePoint-Audit-" + (Get-Date).ToString("yyyyMMdd-HHmm")
                    Write-Host "`n[!] Application registration for '$targetAppName' failed: $errMsg" -ForegroundColor Yellow
                    Write-Host "[*] Suggested unique application name: $suggestedName" -ForegroundColor Cyan
                    Write-Host "[?] Press ENTER to try [$suggestedName], enter a different name, or paste an existing Application ID (GUID): " -NoNewline -ForegroundColor Yellow
                    $inputVal = Read-Host

                    if ([string]::IsNullOrWhiteSpace($inputVal)) {
                        $targetAppName = $suggestedName
                        Write-AuditLog -Level INFO -Message "Retrying registration with suggested name: $targetAppName"
                        continue
                    }
                    elseif ($inputVal.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                        $pastedId = $inputVal.Trim()
                        Write-AuditLog -Level SUCCESS -Message "Using provided existing Application ID: $pastedId"
                        try {
                            $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
                            if (-not (Test-Path -Path $cacheDir)) {
                                $null = New-Item -Path $cacheDir -ItemType Directory -Force
                            }
                            $cacheFile = Join-Path -Path $cacheDir -ChildPath "${Tenant}_client_id.txt"
                            [System.IO.File]::WriteAllText($cacheFile, $pastedId)
                            Write-AuditLog -Level SUCCESS -Message "Application ID cached locally at: $cacheFile"
                        }
                        catch {
                            Write-AuditLog -Level WARN -Message "Unable to persist ClientId to cache: $($_.Exception.Message)"
                        }
                        return $pastedId
                    }
                    else {
                        $targetAppName = $inputVal.Trim()
                        Write-AuditLog -Level INFO -Message "Retrying registration with new name: $targetAppName"
                        continue
                    }
                }
                else {
                    if ($errMsg -match 'already exists') {
                        $targetAppName = "SharePoint-Audit-" + (Get-Date).ToString("yyyyMMdd-HHmm")
                        continue
                    }
                    $regErr = "Failed to register Microsoft Entra ID application: $errMsg"
                    Write-AuditLog -Level ERROR -Message $regErr
                    throw $regErr
                }
            }
        }

        throw "Exceeded maximum attempts for Microsoft Entra ID application registration."
    }

    # ==============================================================================
    # MODULE VERIFICATION & CONNECTION MANAGEMENT (PnP + SPO Fallback)
    # ==============================================================================
    function Assert-SharePointEnvironment {
        [CmdletBinding()]
        param()

        $hasPnP = $null -ne (Get-Module -ListAvailable -Name $Script:Metadata.PrimaryMod)
        $hasSPO = $null -ne (Get-Module -ListAvailable -Name $Script:Metadata.FallbackMod)

        if ($hasPnP) {
            $Script:EngineMode = 'PnP'
            if (-not (Get-Module -Name $Script:Metadata.PrimaryMod)) {
                Write-AuditLog -Level INFO -Message "Importing primary module '$($Script:Metadata.PrimaryMod)'..."
                Import-Module -Name $Script:Metadata.PrimaryMod -ErrorAction Stop
            }
            $modVer = (Get-Module -Name $Script:Metadata.PrimaryMod).Version
            Write-AuditLog -Level SUCCESS -Message "Operating with primary module '$($Script:Metadata.PrimaryMod)' (v$modVer)."
            return
        }

        if ($hasSPO) {
            $Script:EngineMode = 'SPO'
            if (-not (Get-Module -Name $Script:Metadata.FallbackMod)) {
                Write-AuditLog -Level INFO -Message "Importing fallback module '$($Script:Metadata.FallbackMod)'..."
                Import-Module -Name $Script:Metadata.FallbackMod -ErrorAction Stop
            }
            $modVer = (Get-Module -Name $Script:Metadata.FallbackMod).Version
            Write-AuditLog -Level WARN -Message "Primary module '$($Script:Metadata.PrimaryMod)' not found. Operating with fallback module '$($Script:Metadata.FallbackMod)' (v$modVer)."
            return
        }

        $err = @"
Neither '$($Script:Metadata.PrimaryMod)' nor '$($Script:Metadata.FallbackMod)' is installed.
Please install the recommended PnP.PowerShell module:
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
"@
        Write-AuditLog -Level ERROR -Message $err
        throw $err
    }

    function Connect-SharePointSession {
        [CmdletBinding()]
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetAdminUrl
        )

        Write-AuditLog -Level INFO -Message "Validating administrator connection for: $TargetAdminUrl"

        if ($Script:EngineMode -eq 'PnP') {
            # 1. Check if active PnP connection already matches or has access
            try {
                $activeConn = Get-PnPConnection -ErrorAction SilentlyContinue
                if ($null -ne $activeConn -and -not [string]::IsNullOrWhiteSpace($activeConn.Url)) {
                    $activeHost = ([System.Uri]$activeConn.Url).Host
                    $targetHost = ([System.Uri]$TargetAdminUrl).Host
                    if ($activeHost -eq $targetHost) {
                        Write-AuditLog -Level SUCCESS -Message "Existing active PnP session verified for: $TargetAdminUrl"
                        $Script:ActiveConnection = $activeConn
                        return
                    }
                }
            }
            catch {
                $null = $_
            }

            # Determine Tenant Domain from AdminUrl if not explicitly provided
            $effectiveTenant = $TenantId
            if ([string]::IsNullOrWhiteSpace($effectiveTenant)) {
                try {
                    $uri = [System.Uri]$TargetAdminUrl
                    $subdomain = $uri.Host.Split('.')[0].Replace('-admin', '').Replace('-my', '')
                    if (-not [string]::IsNullOrWhiteSpace($subdomain)) {
                        $effectiveTenant = "${subdomain}.onmicrosoft.com"
                    }
                }
                catch {
                    $null = $_
                }
            }

            # 2. Check for App-Only Certificate daemon mode
            if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or -not [string]::IsNullOrWhiteSpace($CertificatePath)) {
                Write-AuditLog -Level INFO -Message "Establishing App-Only daemon connection via certificate..."
                $certParams = @{
                    Url              = $TargetAdminUrl
                    ClientId         = $ClientId
                    Tenant           = $effectiveTenant
                    ReturnConnection = $true
                    ErrorAction      = 'Stop'
                }
                if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                    $certParams['Thumbprint'] = $CertificateThumbprint
                }
                elseif (-not [string]::IsNullOrWhiteSpace($CertificatePath)) {
                    $certParams['CertificatePath'] = $CertificatePath
                    if ($null -ne $CertificatePassword) {
                        $certParams['CertificatePassword'] = $CertificatePassword
                    }
                }
                $Script:ActiveConnection = Connect-PnPOnline @certParams
                $Script:ActiveClientId = $ClientId
                $Script:ActiveTenant   = $effectiveTenant
                Write-AuditLog -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via certificate daemon."
                return
            }

            # 3. Resolve Client ID for Interactive Modern Authentication
            $effectiveClientId = $ClientId

            # 3.1 Check environment variables
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                if (-not [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)) {
                    $effectiveClientId = $env:ENTRAID_APP_ID.Trim()
                }
                elseif (-not [string]::IsNullOrWhiteSpace($env:PNP_APP_ID)) {
                    $effectiveClientId = $env:PNP_APP_ID.Trim()
                }
            }

            # 3.2 Check persistent local disk cache (%LOCALAPPDATA%\SharePoint-Audit)
            $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "SharePoint-Audit"
            if ([string]::IsNullOrWhiteSpace($effectiveClientId) -and (-not [string]::IsNullOrWhiteSpace($effectiveTenant))) {
                $cacheFile = Join-Path -Path $cacheDir -ChildPath "${effectiveTenant}_client_id.txt"
                if (Test-Path -Path $cacheFile) {
                    try {
                        $cachedId = [System.IO.File]::ReadAllText($cacheFile).Trim()
                        if ($cachedId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            Write-AuditLog -Level INFO -Message "Loaded cached Entra ID Application ID for '$effectiveTenant': $cachedId"
                            $effectiveClientId = $cachedId
                        }
                    }
                    catch {
                        Write-AuditLog -Level WARN -Message "Unable to read cached ClientId: $($_.Exception.Message)"
                    }
                }
            }

            # 3.3 Wildcard fallback in cache folder if single cache file exists
            if ([string]::IsNullOrWhiteSpace($effectiveClientId) -and (Test-Path -Path $cacheDir)) {
                $cachedFiles = Get-ChildItem -Path (Join-Path -Path $cacheDir -ChildPath "*_client_id.txt") -ErrorAction SilentlyContinue
                if ($cachedFiles -and $cachedFiles.Count -eq 1) {
                    try {
                        $cachedId = [System.IO.File]::ReadAllText($cachedFiles[0].FullName).Trim()
                        if ($cachedId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            Write-AuditLog -Level INFO -Message "Loaded discovered Entra ID Application ID from $($cachedFiles[0].Name): $cachedId"
                            $effectiveClientId = $cachedId
                        }
                    }
                    catch {
                        $null = $_
                    }
                }
            }

            # 3.4 Guide registration on-the-fly if ClientId is still missing
            if ([string]::IsNullOrWhiteSpace($effectiveClientId)) {
                Write-AuditLog -Level WARN -Message "No Entra ID Application ID configured for tenant '$effectiveTenant'."

                $shouldRegister = $false
                if ([Environment]::UserInteractive) {
                    Write-Host "`n[?] Microsoft Entra ID interactive authentication requires an Application (Client) ID." -ForegroundColor Yellow
                    Write-Host "[?] Microsoft permanently retired the legacy shared multi-tenant PnP app on Sept 9, 2024." -ForegroundColor Yellow
                    Write-Host "[?] Would you like to automatically register and configure an Entra ID App in '$effectiveTenant' now? (Y/N): " -NoNewline -ForegroundColor Yellow
                    $answer = Read-Host
                    if ($answer -match '^[yY1tT]$') {
                        $shouldRegister = $true
                    }
                }

                if ($shouldRegister) {
                    $effectiveClientId = Invoke-EntraIDAppRegistration -Tenant $effectiveTenant
                }
                else {
                    $guidance = @"
Microsoft Entra ID interactive authentication requires an Application (Client) ID.
Microsoft permanently retired the legacy shared multi-tenant PnP app on September 9, 2024.

To register a dedicated Entra ID App for your tenant ('$effectiveTenant'), run once in PowerShell 7:
    Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint-Audit" -Tenant "$effectiveTenant"

Then re-run this script with the generated Application ID:
    .\Get-SharePointSitesAndUsers.ps1 -AdminCenterUrl "$TargetAdminUrl" -ClientId "<ApplicationId>" -Interactive
"@
                    Write-AuditLog -Level ERROR -Message $guidance
                    throw $guidance
                }
            }

            # Attempt interactive connection with the effective ClientId and Tenant
            $Script:ActiveClientId = $effectiveClientId
            $Script:ActiveTenant   = $effectiveTenant

            $connectParams = @{
                Url              = $TargetAdminUrl
                Interactive      = $true
                ClientId         = $effectiveClientId
                Tenant           = $effectiveTenant
                ReturnConnection = $true
                ErrorAction      = 'Stop'
            }

            Write-AuditLog -Level INFO -Message "Initiating modern interactive browser authentication (ClientId: $effectiveClientId, Tenant: $effectiveTenant)..."
            try {
                $Script:ActiveConnection = Connect-PnPOnline @connectParams
                Write-AuditLog -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via interactive PnP session."
            }
            catch {
                if ($_.Exception.Message -match 'AADSTS700016' -or $_.Exception.Message -match 'was not found in the directory') {
                    Write-AuditLog -Level WARN -Message "Application ID '$effectiveClientId' was not found or not consented in directory '$effectiveTenant'."

                    if ([Environment]::UserInteractive) {
                        Write-Host "`n[!] Application ID '$effectiveClientId' was not found in directory '$effectiveTenant' (AADSTS700016)." -ForegroundColor Yellow
                        Write-Host "[?] Would you like to register a new Entra ID App in '$effectiveTenant' now? (Y/N): " -NoNewline -ForegroundColor Yellow
                        $reAns = Read-Host
                        if ($reAns -match '^[yY1tT]$') {
                            $newId = Invoke-EntraIDAppRegistration -Tenant $effectiveTenant
                            $connectParams['ClientId'] = $newId
                            $Script:ActiveClientId = $newId
                            Write-AuditLog -Level INFO -Message "Retrying connection with new Application ID: $newId..."
                            $Script:ActiveConnection = Connect-PnPOnline @connectParams
                            Write-AuditLog -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via interactive session."
                            return
                        }
                    }
                }
                Write-AuditLog -Level ERROR -Message "Connect-PnPOnline failed: $($_.Exception.Message)"
                throw
            }
        }
        else {
            # SPO PowerShell Fallback
            Write-AuditLog -Level INFO -Message "Connecting via Microsoft.Online.SharePoint.PowerShell (Connect-SPOService)..."
            try {
                Connect-SPOService -Url $TargetAdminUrl -ErrorAction Stop
                Write-AuditLog -Level SUCCESS -Message "Connected successfully to $TargetAdminUrl via Connect-SPOService."
            }
            catch {
                Write-AuditLog -Level ERROR -Message "Connect-SPOService failed: $($_.Exception.Message)"
                throw
            }
        }
    }

    # ==============================================================================
    # CORE SITE PERMISSION INSPECTION ENGINE (Specification Section 3 & 4)
    # ==============================================================================
    function Get-SiteSecurityRecord {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        param(
            [Parameter(Mandatory = $true)]
            [string]$TargetSiteUrl,

            [Parameter(Mandatory = $false)]
            $SiteInventoryObject = $null
        )

        Write-AuditLog -Level INFO -Message "Auditing permissions on site: $TargetSiteUrl"

        # Initialize defaults
        $siteTitle = "Untitled Site"
        $template  = "Unknown"
        $storageMB = 0

        if ($null -ne $SiteInventoryObject) {
            $siteTitle = [string](Get-SafePropertyValue -InputObject $SiteInventoryObject -PropertyName "Title" -DefaultValue "Untitled Site")
            $template  = [string](Get-SafePropertyValue -InputObject $SiteInventoryObject -PropertyName "Template" -DefaultValue "Unknown")
            $rawStorage = Get-SafePropertyValue -InputObject $SiteInventoryObject -PropertyName "StorageUsageCurrent" -DefaultValue 0
            if ($null -ne $rawStorage) {
                $storageMB = [int]$rawStorage
            }
        }

        $adminsList = [System.Collections.Generic.List[string]]::new()
        $fullList   = [System.Collections.Generic.List[string]]::new()
        $writeList  = [System.Collections.Generic.List[string]]::new()
        $readList   = [System.Collections.Generic.List[string]]::new()

        $siteConn = $null
        try {
            if ($Script:EngineMode -eq 'PnP') {
                # Connect to specific site collection
                $siteConnParams = @{
                    Url              = $TargetSiteUrl
                    ReturnConnection = $true
                    ErrorAction      = 'Stop'
                }
                if ($Interactive -or $null -ne $Script:ActiveClientId) {
                    $siteConnParams['Interactive'] = $true
                    if ($null -ne $Script:ActiveClientId) { $siteConnParams['ClientId'] = $Script:ActiveClientId }
                    if ($null -ne $Script:ActiveTenant) { $siteConnParams['Tenant'] = $Script:ActiveTenant }
                }
                elseif (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                    $siteConnParams['ClientId'] = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $Script:ActiveClientId }
                    $siteConnParams['Tenant'] = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { $Script:ActiveTenant }
                    $siteConnParams['Thumbprint'] = $CertificateThumbprint
                }
                elseif (-not [string]::IsNullOrWhiteSpace($CertificatePath)) {
                    $siteConnParams['ClientId'] = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $Script:ActiveClientId }
                    $siteConnParams['Tenant'] = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { $Script:ActiveTenant }
                    $siteConnParams['CertificatePath'] = $CertificatePath
                    if ($null -ne $CertificatePassword) {
                        $siteConnParams['CertificatePassword'] = $CertificatePassword
                    }
                }

                $siteConn = try {
                    Connect-PnPOnline @siteConnParams
                }
                catch {
                    # Fallback to inheriting current default connection
                    Get-PnPConnection -ErrorAction SilentlyContinue
                }

                # Retrieve site and web properties
                $web = Invoke-WithRetry -ScriptBlock {
                    if ($null -ne $siteConn) {
                        Get-PnPWeb -Connection $siteConn -Includes Title, Url, AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup -ErrorAction Stop
                    } else {
                        Get-PnPWeb -Includes Title, Url, AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup -ErrorAction Stop
                    }
                }

                if ($siteTitle -eq "Untitled Site" -and $null -ne $web -and -not [string]::IsNullOrWhiteSpace($web.Title)) {
                    $siteTitle = $web.Title
                }

                # Query Site Collection Details (Template & GroupId)
                $siteObj = try {
                    if ($null -ne $siteConn) {
                        Get-PnPSite -Connection $siteConn -Includes GroupId, Template, Usage -ErrorAction SilentlyContinue
                    } else {
                        Get-PnPSite -Includes GroupId, Template, Usage -ErrorAction SilentlyContinue
                    }
                } catch {
                    $null = $_
                    $null
                }

                if ($null -ne $siteObj) {
                    if ($template -eq "Unknown" -and $siteObj.Template) {
                        $template = $siteObj.Template
                    }
                    if ($storageMB -eq 0 -and $siteObj.Usage -and $siteObj.Usage.Storage) {
                        $storageMB = [int]$siteObj.Usage.Storage
                    }
                }

                $connectedGroupId = ""
                if ($null -ne $siteObj -and $siteObj.GroupId -and $siteObj.GroupId -ne [System.Guid]::Empty) {
                    $connectedGroupId = $siteObj.GroupId.ToString()
                }

                # 1. AUDIT SITE COLLECTION ADMINISTRATORS
                try {
                    $admins = Invoke-WithRetry -ScriptBlock {
                        if ($null -ne $siteConn) {
                            Get-PnPSiteCollectionAdmin -Connection $siteConn -ErrorAction Stop
                        } else {
                            Get-PnPSiteCollectionAdmin -ErrorAction Stop
                        }
                    }
                    if ($null -ne $admins) {
                        foreach ($adm in @($admins)) {
                            $ident = Get-NormalizedIdentity -Principal $adm
                            if (-not [string]::IsNullOrWhiteSpace($ident)) {
                                $adminsList.Add($ident)
                            }
                        }
                    }
                }
                catch {
                    Write-AuditLog -Level WARN -Message "Unable to query Site Collection Admins for '$TargetSiteUrl': $($_.Exception.Message)"
                }

                # 2. AUDIT ASSOCIATED GROUPS (Owners, Members, Visitors)
                # Owners
                if ($null -ne $web.AssociatedOwnerGroup) {
                    $ownerGroupMembers = try {
                        if ($null -ne $siteConn) {
                            Get-PnPGroupMember -Group $web.AssociatedOwnerGroup.Id -Connection $siteConn -ErrorAction SilentlyContinue
                        } else {
                            Get-PnPGroupMember -Group $web.AssociatedOwnerGroup.Id -ErrorAction SilentlyContinue
                        }
                    } catch {
                        $null = $_
                        $null
                    }

                    $foundOwnerCount = 0
                    if ($null -ne $ownerGroupMembers) {
                        foreach ($ogm in @($ownerGroupMembers)) {
                            $foundOwnerCount++
                            $ident = Get-NormalizedIdentity -Principal $ogm
                            if (-not [string]::IsNullOrWhiteSpace($ident)) { $fullList.Add($ident) }
                        }
                    }

                    if ($foundOwnerCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($connectedGroupId)) {
                        $nestedOwners = Get-NestedGroupUser -GroupTitle $web.AssociatedOwnerGroup.Title -ConnectedSiteGroupId $connectedGroupId -Connection $siteConn
                        foreach ($no in $nestedOwners) { $fullList.Add($no) }
                    }
                }

                # Members
                if ($null -ne $web.AssociatedMemberGroup) {
                    $memberGroupMembers = try {
                        if ($null -ne $siteConn) {
                            Get-PnPGroupMember -Group $web.AssociatedMemberGroup.Id -Connection $siteConn -ErrorAction SilentlyContinue
                        } else {
                            Get-PnPGroupMember -Group $web.AssociatedMemberGroup.Id -ErrorAction SilentlyContinue
                        }
                    } catch {
                        $null = $_
                        $null
                    }

                    $foundMemberCount = 0
                    if ($null -ne $memberGroupMembers) {
                        foreach ($mgm in @($memberGroupMembers)) {
                            $foundMemberCount++
                            $ident = Get-NormalizedIdentity -Principal $mgm
                            if (-not [string]::IsNullOrWhiteSpace($ident)) { $writeList.Add($ident) }
                        }
                    }

                    if ($foundMemberCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($connectedGroupId)) {
                        $nestedMembers = Get-NestedGroupUser -GroupTitle $web.AssociatedMemberGroup.Title -ConnectedSiteGroupId $connectedGroupId -Connection $siteConn
                        foreach ($nm in $nestedMembers) { $writeList.Add($nm) }
                    }
                }

                # Visitors
                if ($null -ne $web.AssociatedVisitorGroup) {
                    $visitorGroupMembers = try {
                        if ($null -ne $siteConn) {
                            Get-PnPGroupMember -Group $web.AssociatedVisitorGroup.Id -Connection $siteConn -ErrorAction SilentlyContinue
                        } else {
                            Get-PnPGroupMember -Group $web.AssociatedVisitorGroup.Id -ErrorAction SilentlyContinue
                        }
                    } catch {
                        $null = $_
                        $null
                    }

                    if ($null -ne $visitorGroupMembers) {
                        foreach ($vgm in @($visitorGroupMembers)) {
                            $ident = Get-NormalizedIdentity -Principal $vgm
                            if (-not [string]::IsNullOrWhiteSpace($ident)) { $readList.Add($ident) }
                        }
                    }
                }

                # 3. AUDIT ROLE ASSIGNMENTS & DEFINITIONS
                try {
                    $null = if ($null -ne $siteConn) {
                        Get-PnPProperty -ClientObject $web -Property RoleAssignments -Connection $siteConn
                    } else {
                        Get-PnPProperty -ClientObject $web -Property RoleAssignments
                    }

                    if ($null -ne $web.RoleAssignments) {
                        foreach ($assignment in $web.RoleAssignments) {
                            $null = if ($null -ne $siteConn) {
                                Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings -Connection $siteConn
                            } else {
                                Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings
                            }

                            $member = $assignment.Member
                            $roles = @($assignment.RoleDefinitionBindings | ForEach-Object { $_.Name })
                            $pType = $member.PrincipalType.ToString()

                            $isFullControl = ($roles -contains "Full Control" -or $roles -contains "Pełna kontrola")
                            $isWrite       = ($roles -contains "Edit" -or $roles -contains "Contribute" -or $roles -contains "Edycja" -or $roles -contains "Współtworzenie" -or $roles -contains "Design" -or $roles -contains "Projektowanie")
                            $isRead        = ($roles -contains "Read" -or $roles -contains "View Only" -or $roles -contains "Odczyt" -or $roles -contains "Tylko wyświetlanie" -or $roles -contains "Restricted View")

                            if ($pType -eq 'SharePointGroup') {
                                # Enumerate members of this group
                                $grpMembers = try {
                                    if ($null -ne $siteConn) {
                                        Get-PnPGroupMember -Group $member.Id -Connection $siteConn -ErrorAction SilentlyContinue
                                    } else {
                                        Get-PnPGroupMember -Group $member.Id -ErrorAction SilentlyContinue
                                    }
                                } catch {
                                    $null = $_
                                    $null
                                }

                                $mCount = 0
                                if ($null -ne $grpMembers) {
                                    foreach ($gm in @($grpMembers)) {
                                        $mCount++
                                        $ident = Get-NormalizedIdentity -Principal $gm
                                        if (-not [string]::IsNullOrWhiteSpace($ident)) {
                                            if ($isFullControl) { $fullList.Add($ident) }
                                            elseif ($isWrite)   { $writeList.Add($ident) }
                                            elseif ($isRead)    { $readList.Add($ident) }
                                        }
                                    }
                                }

                                if ($mCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($connectedGroupId)) {
                                    $nested = Get-NestedGroupUser -GroupTitle $member.Title -ConnectedSiteGroupId $connectedGroupId -Connection $siteConn
                                    foreach ($n in $nested) {
                                        if ($isFullControl) { $fullList.Add($n) }
                                        elseif ($isWrite)   { $writeList.Add($n) }
                                        elseif ($isRead)    { $readList.Add($n) }
                                    }
                                }

                                # Fallback: if group cannot be expanded, record group title
                                if ($mCount -eq 0) {
                                    $grpIdent = Get-NormalizedIdentity -Principal $member
                                    if (-not [string]::IsNullOrWhiteSpace($grpIdent)) {
                                        if ($isFullControl) { $fullList.Add($grpIdent) }
                                        elseif ($isWrite)   { $writeList.Add($grpIdent) }
                                        elseif ($isRead)    { $readList.Add($grpIdent) }
                                    }
                                }
                            }
                            else {
                                # Direct User or Security Group Assignment
                                $ident = Get-NormalizedIdentity -Principal $member
                                if (-not [string]::IsNullOrWhiteSpace($ident)) {
                                    if ($isFullControl) { $fullList.Add($ident) }
                                    elseif ($isWrite)   { $writeList.Add($ident) }
                                    elseif ($isRead)    { $readList.Add($ident) }
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-AuditLog -Level WARN -Message "Unable to process role assignments for '$TargetSiteUrl': $($_.Exception.Message)"
                }
            }
            else {
                # SPO PowerShell Fallback Engine
                $spoUsers = Invoke-WithRetry -ScriptBlock {
                    Get-SPOUser -Site $TargetSiteUrl -Limit All -ErrorAction Stop
                }

                $spoGroups = try {
                    Get-SPOSiteGroup -Site $TargetSiteUrl -ErrorAction SilentlyContinue
                } catch {
                    $null = $_
                    $null
                }

                $ownerGroupNames = [System.Collections.Generic.List[string]]::new()
                $memberGroupNames = [System.Collections.Generic.List[string]]::new()
                $visitorGroupNames = [System.Collections.Generic.List[string]]::new()

                if ($null -ne $spoGroups) {
                    foreach ($g in @($spoGroups)) {
                        $r = @($g.Roles)
                        if ($r -contains "Full Control" -or $r -contains "Pełna kontrola" -or $g.Title -like "*Owner*" -or $g.Title -like "*Właściciel*") {
                            $ownerGroupNames.Add($g.Title)
                        }
                        elseif ($r -contains "Edit" -or $r -contains "Contribute" -or $r -contains "Edycja" -or $r -contains "Współtworzenie" -or $g.Title -like "*Member*" -or $g.Title -like "*Członk*") {
                            $memberGroupNames.Add($g.Title)
                        }
                        elseif ($r -contains "Read" -or $r -contains "Odczyt" -or $g.Title -like "*Visitor*" -or $g.Title -like "*Goście*" -or $g.Title -like "*Czytelnicy*") {
                            $visitorGroupNames.Add($g.Title)
                        }
                    }
                }

                if ($null -ne $spoUsers) {
                    foreach ($u in @($spoUsers)) {
                        $ident = Get-NormalizedIdentity -Principal $u
                        if ([string]::IsNullOrWhiteSpace($ident)) { continue }

                        if ($u.IsSiteAdmin) {
                            $adminsList.Add($ident)
                        }

                        $userGroups = @($u.Groups)
                        $inOwners   = ($userGroups | Where-Object { $ownerGroupNames -contains $_ })
                        $inMembers  = ($userGroups | Where-Object { $memberGroupNames -contains $_ })
                        $inVisitors = ($userGroups | Where-Object { $visitorGroupNames -contains $_ })

                        if ($inOwners)   { $fullList.Add($ident) }
                        if ($inMembers)  { $writeList.Add($ident) }
                        if ($inVisitors) { $readList.Add($ident) }
                    }
                }
            }

            # Deduplicate and sort identities per category
            $adminsClean = @($adminsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object)
            $fullClean   = @($fullList   | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object)
            $writeClean  = @($writeList  | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object)
            $readClean   = @($readList   | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Sort-Object)

            $adminsStr = $adminsClean -join '; '
            $fullStr   = $fullClean   -join '; '
            $writeStr  = $writeClean  -join '; '
            $readStr   = $readClean   -join '; '

            $Script:ProcessedSites++
            Write-AuditLog -Level SUCCESS -Message "Successfully audited '$siteTitle' (Admins: $($adminsClean.Count), Owners: $($fullClean.Count), Members: $($writeClean.Count), Visitors: $($readClean.Count))."

            return [PSCustomObject]@{
                SiteTitle            = $siteTitle
                SiteUrl              = $TargetSiteUrl
                StorageUsedMB        = $storageMB
                Template             = $template
                SiteCollectionAdmins = $adminsStr
                FullControl          = $fullStr
                Write                = $writeStr
                Read                 = $readStr
                AuditDate            = (Get-Date)
            }
        }
        catch {
            # SECTION 7: RESILIENCE REQUIREMENT (401/403 or other single-site error cannot abort the audit)
            $Script:FailedSites++
            $errDetail = $_.Exception.Message
            Write-AuditLog -Level WARN -Message "Access denied or audit error on site '$TargetSiteUrl': $errDetail"

            return [PSCustomObject]@{
                SiteTitle            = $siteTitle
                SiteUrl              = $TargetSiteUrl
                StorageUsedMB        = $storageMB
                Template             = $template
                SiteCollectionAdmins = "Access Denied / Insufficient Permissions"
                FullControl          = "Access Denied"
                Write                = "Access Denied"
                Read                 = "Access Denied"
                AuditDate            = (Get-Date)
            }
        }
        finally {
            if ($null -ne $siteConn -and $siteConn -ne $Script:ActiveConnection) {
                try {
                    Disconnect-PnPOnline -Connection $siteConn -ErrorAction SilentlyContinue
                }
                catch {
                    $null = $_
                }
            }
        }
    }

    # ==============================================================================
    # PREFLIGHT CHECKS & INITIALIZATION
    # ==============================================================================
    Assert-PowerShellVersion -BoundParams $PSBoundParameters
    Initialize-AuditLog -Directory $LogPath
    Write-AuditLog -Level INFO -Message "Starting $($Script:Metadata.Title) v$($Script:Metadata.Version)"
    Write-AuditLog -Level INFO -Message "Audit transcript initialized at: $Script:LogFilePath"

    Assert-SharePointEnvironment

    $targetSite = if ($PSCmdlet.ParameterSetName -eq 'SingleSite') { $Site } else { "" }
    $resolvedAdminUrl = Resolve-AdminCenterUrlInternal `
        -ExplicitAdminUrl $AdminCenterUrl `
        -TargetSiteUrl $targetSite `
        -TenantDomain $TenantId

    if ([string]::IsNullOrWhiteSpace($resolvedAdminUrl)) {
        if ([Environment]::UserInteractive) {
            $inputAdmin = Read-AdminCenterUrlInteractive
            if (-not [string]::IsNullOrWhiteSpace($inputAdmin)) {
                $resolvedAdminUrl = Resolve-AdminCenterUrlInternal -ExplicitAdminUrl $inputAdmin
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolvedAdminUrl)) {
        $msg = "SharePoint Admin Center URL could not be resolved. Please specify -AdminCenterUrl."
        Write-AuditLog -Level ERROR -Message $msg
        throw $msg
    }

    Connect-SharePointSession -TargetAdminUrl $resolvedAdminUrl
}

process {
    if ($Help) {
        return
    }

    # ==============================================================================
    # PARAMETER SET: SINGLESITE (Specification Section 1 & 2)
    # ==============================================================================
    if ($PSCmdlet.ParameterSetName -eq 'SingleSite') {
        $cleanSite = $Site.Trim()
        Write-AuditLog -Level INFO -Message "Running targeted single-site audit on: $cleanSite"

        $inventoryObj = $null
        try {
            if ($Script:EngineMode -eq 'PnP') {
                $inventoryObj = Get-PnPTenantSite -Detailed -Url $cleanSite -Connection $Script:ActiveConnection -ErrorAction SilentlyContinue
            } else {
                $inventoryObj = Get-SPOSite -Identity $cleanSite -Detailed -ErrorAction SilentlyContinue
            }
        } catch {
            $null = $_
            Write-AuditLog -Level DEBUG -Message "Tenant-level inventory lookup skipped for single site: $($_.Exception.Message)"
        }

        $record = Get-SiteSecurityRecord -TargetSiteUrl $cleanSite -SiteInventoryObject $inventoryObj
        Write-Output $record
        return
    }

    # ==============================================================================
    # PARAMETER SET: ALLSITES (Specification Section 1 & 2)
    # ==============================================================================
    if ($PSCmdlet.ParameterSetName -eq 'AllSites') {
        Write-AuditLog -Level INFO -Message "Initiating tenant-wide site collection discovery from: $resolvedAdminUrl"
        Write-Progress -Activity "Auditing SharePoint Sites & Users" -Status "Discovering site collections from tenant endpoint..." -PercentComplete 5

        $discoveredSites = try {
            if ($Script:EngineMode -eq 'PnP') {
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPTenantSite -Detailed -Connection $Script:ActiveConnection -ErrorAction Stop
                }
            } else {
                Invoke-WithRetry -ScriptBlock {
                    Get-SPOSite -Limit All -IncludePersonalSite:$IncludePersonalSites -Detailed -ErrorAction Stop
                }
            }
        }
        catch {
            Write-AuditLog -Level WARN -Message "Detailed site query failed ($($_.Exception.Message)). Retrying standard query..."
            if ($Script:EngineMode -eq 'PnP') {
                Get-PnPTenantSite -Connection $Script:ActiveConnection -ErrorAction Stop
            } else {
                Get-SPOSite -Limit All -IncludePersonalSite:$IncludePersonalSites -ErrorAction Stop
            }
        }

        if ($null -eq $discoveredSites -or (@($discoveredSites).Count -eq 0)) {
            Write-AuditLog -Level WARN -Message "No site collections were returned from the tenant endpoint."
            return
        }

        # Filter out personal sites unless explicitly requested
        $sitesToProcess = [System.Collections.Generic.List[object]]::new()
        foreach ($s in @($discoveredSites)) {
            $sUrl = [string](Get-SafePropertyValue -InputObject $s -PropertyName "Url" -DefaultValue "")
            if (-not $IncludePersonalSites -and ($sUrl -like "*/personal/*" -or $sUrl -like "*-my.sharepoint.com*")) {
                continue
            }
            $sitesToProcess.Add($s)
        }

        $Script:TotalSites = $sitesToProcess.Count
        Write-AuditLog -Level SUCCESS -Message "Discovered $($Script:TotalSites) site collection(s) to audit (Personal sites excluded: $(-not $IncludePersonalSites))."

        $index = 0
        foreach ($siteItem in $sitesToProcess) {
            $index++
            $sUrl   = [string](Get-SafePropertyValue -InputObject $siteItem -PropertyName "Url" -DefaultValue "")
            $sTitle = [string](Get-SafePropertyValue -InputObject $siteItem -PropertyName "Title" -DefaultValue "Untitled Site")

            # Update Progress Bar (Specification Section 7.2)
            $pct = [int](($index / [Math]::Max($Script:TotalSites, 1)) * 100)
            Write-Progress -Activity "Auditing SharePoint Sites & Users" `
                -Status "Processing site $index of $($Script:TotalSites) ($pct%): $sTitle" `
                -PercentComplete $pct

            $secRecord = Get-SiteSecurityRecord -TargetSiteUrl $sUrl -SiteInventoryObject $siteItem
            Write-Output $secRecord
        }

        Write-Progress -Activity "Auditing SharePoint Sites & Users" -Completed
    }
}

end {
    if ($Script:ActiveConnection) {
        try {
            Disconnect-PnPOnline -Connection $Script:ActiveConnection -ErrorAction SilentlyContinue
        }
        catch {
            $null = $_
        }
    }

    $duration = (Get-Date) - $Script:StartTime
    $durFormatted = "{0:mm}m {0:ss}s" -f $duration

    Write-AuditLog -Level SUCCESS -Message "================================================================================"
    Write-AuditLog -Level SUCCESS -Message "SharePoint Sites & Users audit completed in $durFormatted."
    Write-AuditLog -Level SUCCESS -Message "Sites Processed: $Script:ProcessedSites | Failed/Skipped: $Script:FailedSites"
    Write-AuditLog -Level SUCCESS -Message "Detailed execution log saved to: $Script:LogFilePath"
    Write-AuditLog -Level SUCCESS -Message "================================================================================"
}
