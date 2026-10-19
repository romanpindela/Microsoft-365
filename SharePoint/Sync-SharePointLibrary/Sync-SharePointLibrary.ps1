<#
.SYNOPSIS
    Automates SharePoint Online document library synchronization via OneDrive for Business.

.DESCRIPTION
    Sync-SharePointLibrary is an enterprise-grade PowerShell utility designed to map and
    synchronize SharePoint Online document libraries directly to Windows 10/11 File Explorer
    using the native OneDrive for Business sync client ('odopen://sync?' URI protocol).

    Key Capabilities:
    1. Interactive Microsoft 365 Device Code authentication (no heavy SDKs required).
    2. Non-interactive Microsoft Entra ID App Registration authentication (client credentials).
    3. Automatic discovery of all SharePoint sites accessible to the user (via security-trimmed REST Search).
    4. Explicit multi-site / multi-library synchronization via parameters.
    5. Native detection of running OneDrive process and bilingual document library title resolution
       (supporting Polish 'Dokumenty' and English 'Documents').
    6. Duplicate sync prevention by checking local OneDrive registry providers.
    7. Comprehensive execution summary report and detailed logging to C:\Logs.

.PARAMETER SiteName
    One or more site slugs, relative paths, or full SharePoint URLs to synchronize.
    Examples: 'automatyka', @('automatyka', 'zarzad'), or 'https://elektrimont.sharepoint.com/sites/automatyka'.

.PARAMETER TenantHost
    The tenant domain or root SharePoint host.
    Default: 'elektrimont.sharepoint.com'

.PARAMETER DocumentLibraryTitle
    The display title of the target document library. If omitted or not found, the script
    automatically resolves the site's default document library (e.g., 'Dokumenty' / 'Documents').

.PARAMETER AllLibraries
    Switch to automatically synchronize all visible document libraries within the target site(s).

.PARAMETER AutoDiscover
    Switch to automatically search and discover all SharePoint site collections and webs
    the authenticated user has permissions to access across the tenant.

.PARAMETER All
    When used with -AutoDiscover, synchronizes all discovered sites without interactive prompts.

.PARAMETER Force
    Forces sync re-registration even if the library is already detected as synced locally.

.PARAMETER LogDirectory
    Directory where execution logs are stored.
    Default: 'C:\Logs' (falls back to $env:LOCALAPPDATA\Logs if C:\Logs is inaccessible).

.PARAMETER ClientId
    Microsoft Entra ID Application (Client) ID for non-interactive app-only authentication.

.PARAMETER TenantId
    Microsoft Entra ID Directory (Tenant) ID for non-interactive authentication. Default: 'common'.

.PARAMETER ClientSecret
    Microsoft Entra ID Client Secret for non-interactive authentication.

.PARAMETER Help
    Displays detailed usage instructions, examples, Entra ID setup guide, and author details.
    Alias: -h

.EXAMPLE
    # Display help and usage instructions
    .\Sync-SharePointLibrary.ps1 -Help

.EXAMPLE
    # Synchronize a specific site library interactively
    .\Sync-SharePointLibrary.ps1 -SiteName "automatyka"

.EXAMPLE
    # Synchronize multiple sites
    .\Sync-SharePointLibrary.ps1 -SiteName "automatyka", "projekty", "zarzad"

.EXAMPLE
    # Automatically discover accessible sites and choose interactively
    .\Sync-SharePointLibrary.ps1 -AutoDiscover

.EXAMPLE
    # Automatically discover and sync all accessible sites
    .\Sync-SharePointLibrary.ps1 -AutoDiscover -All

.EXAMPLE
    # Synchronize all libraries within a specific site on a custom tenant
    .\Sync-SharePointLibrary.ps1 -TenantHost "contoso.sharepoint.com" -SiteName "engineering" -AllLibraries

.EXAMPLE
    # Non-interactive execution using Microsoft Entra ID App credentials
    .\Sync-SharePointLibrary.ps1 -TenantHost "elektrimont.sharepoint.com" -TenantId "xxx-xxx" -ClientId "yyy-yyy" -ClientSecret "zzz" -SiteName "automatyka"

.NOTES
    Script Name   : Sync-SharePointLibrary.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Company       : Elektrimont Sp. z o.o.
    Version       : 2.0.0
    Created       : 2026-10-16
    Updated       : 2026-10-19
    Target OS     : Windows 10, Windows 11, Windows Server (Desktop Experience)
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName = 'Default', Position = 0, Mandatory = $false)]
    [Alias('Sites', 'SiteUrl')]
    [string[]]$SiteName,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $false)]
    [string]$TenantHost = "elektrimont.sharepoint.com",

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $false)]
    [string]$DocumentLibraryTitle,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $false)]
    [switch]$AllLibraries,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [switch]$AutoDiscover,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [switch]$All,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $false)]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $false)]
    [string]$LogDirectory = "C:\Logs",

    # Non-interactive Entra ID App registration credentials
    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $true)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $true)]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'NonInteractive', Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# GLOBAL SCRIPT METADATA & CONFIGURATION
# ==============================================================================
$Script:Metadata = @{
    Name        = "Sync-SharePointLibrary.ps1"
    Version     = "2.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Company     = "Elektrimont Sp. z o.o."
    LastUpdated = "2026-10-19"
}

# Execution tracking
$Script:StartTime   = Get-Date
$Script:LogFilePath = $null
$Script:SummaryList = [System.Collections.Generic.List[PSObject]]::new()

# Well-known public Microsoft client ID (Azure CLI) pre-consented for SharePoint delegated access
$Script:DefaultPublicClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"

# ==============================================================================
# LOGGING SYSTEM (CleanCode & PowerShell Script Standards)
# ==============================================================================
function Initialize-LoggingSystem {
    [CmdletBinding()]
    param([string]$TargetDir)

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $computer  = $env:COMPUTERNAME
    $user      = $env:USERNAME
    $logName   = "${timestamp}-Sync-SharePointLibrary-${computer}-${user}.log"

    $resolvedDir = $TargetDir
    try {
        if (-not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        # Fallback to LocalAppData\Logs if C:\Logs is inaccessible
        $resolvedDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Logs"
        if (-not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    $Script:LogFilePath = Join-Path -Path $resolvedDir -ChildPath $logName
    $header = @"
================================================================================
Log File Created : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script           : $($Script:Metadata.Name) v$($Script:Metadata.Version)
Author           : $($Script:Metadata.Author) ($($Script:Metadata.Email))
Company          : $($Script:Metadata.Company)
Host / User      : $computer / $user
Target Tenant    : $TenantHost
================================================================================
"@
    Add-Content -Path $Script:LogFilePath -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Write-Log {
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

# ==============================================================================
# HELP & SECURITY GUARD (Standards 1, 7, 9, 11, 13, 14)
# ==============================================================================
function Show-ScriptHelp {
    Write-Host @"
================================================================================
SCRIPT INFORMATION: $($Script:Metadata.Name) v$($Script:Metadata.Version)
Author   : $($Script:Metadata.Author)
Email    : $($Script:Metadata.Email)
GitHub   : $($Script:Metadata.GitHub)
Company  : $($Script:Metadata.Company)
================================================================================

SYNOPSIS:
  Automates the mapping and synchronization of SharePoint Online document libraries
  to the local Windows 10/11 File Explorer via the OneDrive for Business sync client.

PREREQUISITES:
  1. OneDrive Desktop Client:
     - The user MUST be logged into the local OneDrive for Business desktop application.
     - Login credentials: Corporate M365 work email and domain password.
  2. Standard User Execution (NO ELEVATION):
     - Modern OneDrive sync runs strictly inside the active interactive desktop user session.
     - DO NOT run this script as Administrator (UAC elevated).
  3. Script Unblocking (After Downloading from GitHub):
     - Run: Unblock-File -Path .\Sync-SharePointLibrary.ps1

PARAMETERS:
  -SiteName <String[]>
      Name(s), relative slug(s), or full URL(s) of the target SharePoint site(s).
      Examples: 'automatyka', @('automatyka', 'zarzad'), 'https://elektrimont.sharepoint.com/sites/automatyka'

  -TenantHost <String>
      The domain of the target Microsoft 365 SharePoint tenant (e.g., 'elektrimont.sharepoint.com').
      Default: 'elektrimont.sharepoint.com'

  -AutoDiscover
      Automatically searches and enumerates all SharePoint site collections and webs across
      the tenant that the authenticated user has permissions to access (using security-trimmed REST API).

  -All
      When combined with -AutoDiscover, syncs all discovered accessible sites without prompting.

  -DocumentLibraryTitle <String>
      Title of the document library (e.g., 'Dokumenty', 'Documents', 'Projekty').
      If omitted, the script automatically resolves the site's default document library.

  -AllLibraries
      Synchronizes all visible document libraries located within the specified site(s).

  -Force
      Forces synchronization even if the library is already registered in OneDrive registry.

  -LogDirectory <String>
      Custom directory for execution logs. Default: 'C:\Logs'

  -Help, -h
      Displays this comprehensive help documentation.

NON-INTERACTIVE / ENTRA ID APP REGISTRATION MODE:
  To run this script unattended (e.g., via automated provisioning or deployment pipelines):
  1. Register an Application in Microsoft Entra admin center (https://entra.microsoft.com).
  2. Navigate to 'API permissions' -> 'Add a permission' -> 'SharePoint'.
  3. Select 'Application permissions' and grant:
     - Sites.Read.All (or Sites.FullControl.All)
  4. Click 'Grant admin consent for <Tenant>'.
  5. Under 'Certificates & secrets', create a new Client Secret and record the value.
  6. Call this script with:
     .\Sync-SharePointLibrary.ps1 -TenantId "<TenantGUID>" -ClientId "<AppID>" -ClientSecret "<Secret>" -SiteName "automatyka"

USAGE EXAMPLES:
  # 1. Sync a single site's default document library:
  .\Sync-SharePointLibrary.ps1 -SiteName "automatyka"

  # 2. Sync multiple sites simultaneously:
  .\Sync-SharePointLibrary.ps1 -SiteName "automatyka", "projekty", "zarzad"

  # 3. Automatically discover accessible sites and interactively select:
  .\Sync-SharePointLibrary.ps1 -AutoDiscover

  # 4. Automatically discover and sync all accessible sites:
  .\Sync-SharePointLibrary.ps1 -AutoDiscover -All

  # 5. Sync a custom document library from a full URL:
  .\Sync-SharePointLibrary.ps1 -SiteName "https://elektrimont.sharepoint.com/sites/zarzad" -DocumentLibraryTitle "Zarzadzeniam"

  # 6. Display this help:
  .\Sync-SharePointLibrary.ps1 -Help
================================================================================
"@ -ForegroundColor Cyan
}

# Enforce display of help when executed without parameters
if ($Help.IsPresent -or ($PSBoundParameters.Count -eq 0)) {
    Show-ScriptHelp
    exit 0
}

# ==============================================================================
# ENVIRONMENT & PREREQUISITE CHECKS (Standards 9, 11)
# ==============================================================================
function Test-EnvironmentPrerequisites {
    # Verify execution privilege using language-independent SID check
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $adminSid  = New-Object Security.Principal.SecurityIdentifier([Security.Principal.WellKnownSidType]::BuiltInAdministratorsSid, $null)

    if ($principal.IsInRole($adminSid)) {
        Write-Warning "`n[SECURITY / ENVIRONMENT ERROR]: OneDrive sync cannot operate with elevated Administrator privileges."
        Write-Host "OneDrive desktop sync is bound strictly to standard user desktop sessions." -ForegroundColor Yellow
        Write-Host "Please relaunch this script in a standard user PowerShell console (without 'Run as Administrator').`n" -ForegroundColor Yellow
        exit 1
    }

    # Verify OneDrive process and executable availability
    Write-Log -Level INFO -Message "Verifying OneDrive desktop client status..."
    $proc = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if (-not $proc) {
        Write-Log -Level WARN -Message "OneDrive process is not currently running. Attempting to launch..."

        $candidatePaths = @(
            "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
            "C:\Program Files\Microsoft OneDrive\OneDrive.exe",
            "C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe"
        )

        $exeFound = $null
        foreach ($p in $candidatePaths) {
            if (Test-Path -Path $p) {
                $exeFound = $p
                break
            }
        }

        if ($exeFound) {
            Start-Process -FilePath $exeFound
            Start-Sleep -Seconds 4
            $checkProc = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
            if ($checkProc) {
                Write-Log -Level SUCCESS -Message "OneDrive desktop client started successfully."
            } else {
                Write-Log -Level WARN -Message "OneDrive launched but process confirmation timed out. Proceeding..."
            }
        } else {
            throw "OneDrive executable was not found on this system. Please install Microsoft OneDrive."
        }
    } else {
        Write-Log -Level SUCCESS -Message "OneDrive client is active and running in background."
    }
}

# ==============================================================================
# INPUT SANITIZATION & SAFE URL RESOLUTION (Standard 7)
# ==============================================================================
function Get-SanitizedTenantHost {
    param([string]$RawHost)

    if ([string]::IsNullOrWhiteSpace($RawHost)) {
        return "elektrimont.sharepoint.com"
    }

    $clean = $RawHost.Trim().ToLower()
    if ($clean -match '^https?://([^/]+)') {
        $clean = $Matches[1]
    }
    $clean = $clean.TrimEnd('/')

    # Basic protection against illegal characters or injection
    if ($clean -notmatch '^[a-zA-Z0-9.\-]+$') {
        throw "Invalid tenant host format: '$RawHost'. Must be a valid domain name."
    }

    return $clean
}

function Resolve-SharePointSiteCanonicalUrl {
    param(
        [Parameter(Mandatory = $true)][string]$InputSite,
        [Parameter(Mandatory = $true)][string]$CleanTenantHost
    )

    $trimmed = $InputSite.Trim()

    # Reject dangerous input containing command injection or illegal chars
    if ($trimmed -match '[;&|`$<>"]') {
        throw "Dangerous input detected in site parameter: '$InputSite'."
    }

    if ($trimmed -match '^https?://') {
        return $trimmed.TrimEnd('/')
    }

    # Relative path starting with /sites/ or /teams/
    if ($trimmed -match '^/?(sites|teams)/') {
        $slug = $trimmed.TrimStart('/')
        return "https://$CleanTenantHost/$slug".TrimEnd('/')
    }

    # Standard site slug alias
    $slug = $trimmed.TrimStart('/')
    return "https://$CleanTenantHost/sites/$slug".TrimEnd('/')
}

# ==============================================================================
# LOCAL ONEDRIVE REGISTRY & SYNC STATE RECOGNITION
# ==============================================================================
function Get-LocalOneDriveAccounts {
    $accounts = @()
    $accKey = "HKCU:\Software\Microsoft\OneDrive\Accounts"
    if (Test-Path $accKey) {
        Get-ChildItem -Path $accKey -ErrorAction SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($props -and $props.UserEmail) {
                $accounts += [PSCustomObject]@{
                    AccountKey  = $_.PSChildName
                    UserEmail   = $props.UserEmail
                    DisplayName = $props.DisplayName
                    UserFolder  = $props.UserFolder
                    EndpointUri = $props.ServiceEndpointUri
                }
            }
        }
    }
    return $accounts
}

function Get-LocalSyncedLibraries {
    $synced = [System.Collections.Generic.List[PSObject]]::new()
    $providerKey = "HKCU:\Software\SyncEngines\Providers\OneDrive"
    if (Test-Path $providerKey) {
        Get-ChildItem -Path $providerKey -ErrorAction SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($props -and $props.UrlNamespace -and $props.MountPoint) {
                $synced.Add([PSCustomObject]@{
                    MountPoint   = $props.MountPoint
                    UrlNamespace = $props.UrlNamespace.TrimEnd('/')
                    LibraryType  = $props.LibraryType
                })
            }
        }
    }
    return $synced
}

# ==============================================================================
# AUTHENTICATION SUBSYSTEM (Interactive Device Code & App Credentials)
# ==============================================================================
function Get-SharePointAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CleanTenantHost,
        [string]$CustomTenantId,
        [string]$CustomClientId,
        [string]$CustomClientSecret
    )

    # 1. Non-Interactive Client Credentials Grant (Microsoft Entra ID App)
    if (-not [string]::IsNullOrWhiteSpace($CustomClientSecret)) {
        Write-Log -Level INFO -Message "Authenticating non-interactively using Microsoft Entra ID App credentials..."
        $tId = if ([string]::IsNullOrWhiteSpace($CustomTenantId)) { "common" } else { $CustomTenantId }
        $tokenEndpoint = "https://login.microsoftonline.com/$tId/oauth2/v2.0/token"

        $tokenBody = @{
            client_id     = $CustomClientId
            client_secret = $CustomClientSecret
            scope         = "https://$CleanTenantHost/.default"
            grant_type    = "client_credentials"
        }

        try {
            $tokenResp = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $tokenBody -ErrorAction Stop
            if ($tokenResp.access_token) {
                Write-Log -Level SUCCESS -Message "Non-interactive app token acquired successfully."
                return $tokenResp.access_token
            }
        }
        catch {
            throw "Non-interactive Entra ID token acquisition failed: $($_.Exception.Message)"
        }
    }

    # 2. Interactive Microsoft 365 OAuth2 Device Code Flow
    Write-Log -Level INFO -Message "Initiating interactive Microsoft 365 Device Code authentication..."

    $cId = if ([string]::IsNullOrWhiteSpace($CustomClientId)) { $Script:DefaultPublicClientId } else { $CustomClientId }
    $tId = if ([string]::IsNullOrWhiteSpace($CustomTenantId)) { "common" } else { $CustomTenantId }
    $resourceUri = "https://$CleanTenantHost"

    $deviceCodeEndpoint = "https://login.microsoftonline.com/$tId/oauth2/devicecode"
    $deviceCodeParams   = @{
        client_id = $cId
        resource  = $resourceUri
    }

    try {
        $dcResponse = Invoke-RestMethod -Uri $deviceCodeEndpoint -Method Post -Body $deviceCodeParams -ErrorAction Stop
    }
    catch {
        throw "Failed to initiate Device Code authentication: $($_.Exception.Message)"
    }

    # Copy code to clipboard safely
    try {
        Set-Clipboard -Value $dcResponse.user_code -ErrorAction SilentlyContinue
    } catch { }

    Write-Host "`n================================================================================" -ForegroundColor Yellow
    Write-Host " MICROSOFT 365 AUTHENTICATION REQUIRED" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host " 1. Browser authentication page : https://microsoft.com/devicelogin" -ForegroundColor White
    Write-Host " 2. Your One-Time Device Code   : $($dcResponse.user_code)" -ForegroundColor Green
    Write-Host "    (The code has been automatically copied to your clipboard - press Ctrl+V!)" -ForegroundColor Gray
    Write-Host " 3. Sign in with your corporate account (e.g. your @elektrimont.pl credentials)" -ForegroundColor White
    Write-Host "================================================================================`n" -ForegroundColor Yellow

    Start-Sleep -Seconds 2
    try {
        Start-Process "https://microsoft.com/devicelogin"
    } catch { }

    Write-Log -Level INFO -Message "Waiting for browser identity verification (User Code: $($dcResponse.user_code))..."

    $interval  = [Math]::Max([int]$dcResponse.interval, 5)
    $expiresIn = [int]$dcResponse.expires_in
    $elapsed   = 0

    $tokenBody = @{
        client_id  = $cId
        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $dcResponse.device_code
        resource   = $resourceUri
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$tId/oauth2/token"

    while ($elapsed -lt $expiresIn) {
        Start-Sleep -Seconds $interval
        $elapsed += $interval
        try {
            $tokenResp = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $tokenBody -ErrorAction Stop
            if ($tokenResp.access_token) {
                Write-Host ""
                Write-Log -Level SUCCESS -Message "User identity verified successfully. Bearer token acquired."
                return $tokenResp.access_token
            }
        }
        catch {
            Write-Host "." -NoNewline -ForegroundColor DarkGray
        }
    }

    throw "Device code authentication timed out after $expiresIn seconds."
}

# ==============================================================================
# SITE & LIBRARY DISCOVERY VIA REST API (Requirement 2 & Bilingual Support)
# ==============================================================================
function Get-AccessibleSharePointSites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CleanTenantHost,
        [Parameter(Mandatory = $true)][string]$Token
    )

    Write-Log -Level INFO -Message "Discovering accessible SharePoint sites via security-trimmed Search API..."

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Accept"        = "application/json;odata=verbose"
    }

    $discoveredList = [System.Collections.Generic.List[PSObject]]::new()

    # Query SharePoint Search REST API - security trimmed to current user permissions
    $searchQuery   = "contentclass:STS_Site OR contentclass:STS_Web"
    $encodedQuery  = [System.Uri]::EscapeDataString($searchQuery)
    $searchEndpoint = "https://$CleanTenantHost/_api/search/query?querytext='$encodedQuery'&rowlimit=300&selectproperties='Title,Path,SiteId,WebId,contentclass'&trimduplicates=false"

    try {
        $searchResp = Invoke-RestMethod -Uri $searchEndpoint -Headers $headers -Method Get -ErrorAction Stop
        $rows = $null
        if ($searchResp.d -and $searchResp.d.query) {
            $rows = $searchResp.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
            if (-not $rows) {
                $rows = $searchResp.d.query.PrimaryQueryResult.RelevantResults.Table.Rows
            }
        } elseif ($searchResp.PrimaryQueryResult) {
            $rows = $searchResp.PrimaryQueryResult.RelevantResults.Table.Rows
        }

        if ($rows) {
            foreach ($row in $rows) {
                $cells = if ($row.Cells.results) { $row.Cells.results } else { $row.Cells }
                $map = @{}
                foreach ($c in $cells) { $map[$c.Key] = $c.Value }

                $path  = $map['Path']
                $title = $map['Title']

                if ([string]::IsNullOrWhiteSpace($path)) { continue }

                # Filter out personal OneDrive (-my.sharepoint.com), search portals, compliance, app catalogs
                $isPersonal = $path -match '(?i)(-my\.sharepoint\.com|/personal/)'
                $isSystem   = $path -match '(?i)(/search(/pages)?|/portals/|/sites/appcatalog|/sites/compliancehub)'

                if (-not $isPersonal -and -not $isSystem) {
                    $cleanPath = $path.TrimEnd('/')
                    $displayTitle = if ([string]::IsNullOrWhiteSpace($title)) { [System.IO.Path]::GetFileName($cleanPath) } else { $title }
                    $discoveredList.Add([PSCustomObject]@{
                        Title = $displayTitle
                        Url   = $cleanPath
                        Type  = $map['contentclass']
                    })
                }
            }
        }
    }
    catch {
        Write-Log -Level WARN -Message "SharePoint Search API query failed: $($_.Exception.Message)"
    }

    # Query Social Following API (followed sites) as supplemental discovery
    try {
        Write-Log -Level DEBUG -Message "Checking followed sites via Social Following API..."
        $followUrl = "https://$CleanTenantHost/_api/social.following/my/followed(types=4)"
        $followResp = Invoke-RestMethod -Uri $followUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
        $items = if ($followResp.d -and $followResp.d.Followed) { $followResp.d.Followed.results } else { $followResp.value }
        if ($items) {
            foreach ($it in $items) {
                $u = $it.Uri
                if ($u -and ($u -notmatch '(?i)(-my\.sharepoint\.com|/personal/)')) {
                    $cleanU = $u.TrimEnd('/')
                    $name = if ($it.Name) { $it.Name } else { [System.IO.Path]::GetFileName($cleanU) }
                    $discoveredList.Add([PSCustomObject]@{
                        Title = $name
                        Url   = $cleanU
                        Type  = "FollowedSite"
                    })
                }
            }
        }
    }
    catch {
        Write-Log -Level DEBUG -Message "Social following check completed."
    }

    # Deduplicate by URL
    $uniqueMap = @{}
    $uniqueList = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($item in $discoveredList) {
        $canonical = $item.Url.ToLower()
        if (-not $uniqueMap.ContainsKey($canonical)) {
            $uniqueMap[$canonical] = $true
            $uniqueList.Add($item)
        }
    }

    return $uniqueList
}

function Get-SharePointLibraryMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FullSiteUrl,
        [Parameter(Mandatory = $true)][string]$Token,
        [string]$TargetLibraryTitle,
        [switch]$FetchAllLibraries
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Accept"        = "application/json;odata=verbose"
    }

    Write-Log -Level INFO -Message "Querying site and web GUIDs from: $FullSiteUrl"

    # 1. Web Metadata
    $webResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web?`$select=Id,Title,Url,ServerRelativeUrl" -Headers $headers -Method Get
    $webId    = if ($webResp.d) { $webResp.d.Id } else { $webResp.Id }
    $webTitle = if ($webResp.d) { $webResp.d.Title } else { $webResp.Title }
    $webUrl   = if ($webResp.d) { $webResp.d.Url } else { $webResp.Url }

    # 2. Site Collection Metadata
    $siteResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/site?`$select=Id" -Headers $headers -Method Get
    $siteId   = if ($siteResp.d) { $siteResp.d.Id } else { $siteResp.Id }

    # 3. User Identity (UPN / Email)
    $userEmail = $null
    try {
        $userResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web/currentuser?`$select=Email,UserPrincipalName" -Headers $headers -Method Get
        if ($userResp.d) {
            $userEmail = if ($userResp.d.Email) { $userResp.d.Email } else { $userResp.d.UserPrincipalName }
        } else {
            $userEmail = if ($userResp.Email) { $userResp.Email } else { $userResp.UserPrincipalName }
        }
    } catch { }

    # Fallback to local OneDrive registry account email if REST user is blank
    if ([string]::IsNullOrWhiteSpace($userEmail)) {
        $accounts = Get-LocalOneDriveAccounts
        if ($accounts.Count -gt 0) {
            $userEmail = $accounts[0].UserEmail
        }
    }

    # 4. Resolve Document Libraries
    $resolvedLibraries = [System.Collections.Generic.List[PSObject]]::new()

    if ($FetchAllLibraries) {
        Write-Log -Level INFO -Message "Enumerating all visible document libraries for site '$webTitle'..."
        $listsResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web/lists?`$filter=BaseTemplate eq 101 and Hidden eq false&`$select=Id,Title" -Headers $headers -Method Get
        $items = if ($listsResp.d) { $listsResp.d.results } else { $listsResp.value }
        foreach ($lib in $items) {
            $resolvedLibraries.Add([PSCustomObject]@{
                ListId    = $lib.Id
                ListTitle = $lib.Title
            })
        }
    }
    else {
        # Check by requested title if specified
        $found = $false
        if (-not [string]::IsNullOrWhiteSpace($TargetLibraryTitle)) {
            try {
                $encTitle = [System.Uri]::EscapeDataString($TargetLibraryTitle)
                $lResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web/lists/getbytitle('$encTitle')?`$select=Id,Title" -Headers $headers -Method Get
                $lid = if ($lResp.d) { $lResp.d.Id } else { $lResp.Id }
                $ltitle = if ($lResp.d) { $lResp.d.Title } else { $lResp.Title }
                $resolvedLibraries.Add([PSCustomObject]@{ ListId = $lid; ListTitle = $ltitle })
                $found = $true
            }
            catch {
                Write-Log -Level WARN -Message "Library '$TargetLibraryTitle' not found by explicit title. Attempting default document library resolution..."
            }
        }

        # Fallback to default document library (supports both Polish 'Dokumenty' and English 'Documents')
        if (-not $found) {
            try {
                $defResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web/defaultDocumentLibrary?`$select=Id,Title" -Headers $headers -Method Get
                $lid = if ($defResp.d) { $defResp.d.Id } else { $defResp.Id }
                $ltitle = if ($defResp.d) { $defResp.d.Title } else { $defResp.Title }
                $resolvedLibraries.Add([PSCustomObject]@{ ListId = $lid; ListTitle = $ltitle })
                $found = $true
            }
            catch { }
        }

        # Fallback bilingual trial: 'Dokumenty', 'Documents', 'Shared Documents'
        if (-not $found) {
            $trialTitles = @("Dokumenty", "Documents", "Shared Documents", "Dokumenty udostepnione")
            foreach ($t in $trialTitles) {
                try {
                    $encT = [System.Uri]::EscapeDataString($t)
                    $lResp = Invoke-RestMethod -Uri "$FullSiteUrl/_api/web/lists/getbytitle('$encT')?`$select=Id,Title" -Headers $headers -Method Get
                    $lid = if ($lResp.d) { $lResp.d.Id } else { $lResp.Id }
                    $ltitle = if ($lResp.d) { $lResp.d.Title } else { $lResp.Title }
                    $resolvedLibraries.Add([PSCustomObject]@{ ListId = $lid; ListTitle = $ltitle })
                    $found = $true
                    break
                } catch { }
            }
        }
    }

    if ($resolvedLibraries.Count -eq 0) {
        throw "No accessible document libraries could be resolved for site: $FullSiteUrl"
    }

    return @{
        SiteId    = $siteId.ToString().Trim('{}')
        WebId     = $webId.ToString().Trim('{}')
        WebTitle  = $webTitle
        WebUrl    = $webUrl
        UserEmail = $userEmail
        Libraries = $resolvedLibraries
    }
}

# ==============================================================================
# ONEDRIVE SYNCHRONIZATION TRIGGER (Requirement 3)
# ==============================================================================
function Invoke-OneDriveSyncRegistration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$SiteMeta,
        [Parameter(Mandatory = $true)][PSCustomObject]$Library,
        [switch]$ForceSync
    )

    $siteId    = $SiteMeta.SiteId
    $webId     = $SiteMeta.WebId
    $listId    = $Library.ListId.ToString().Trim('{}')
    $userEmail = $SiteMeta.UserEmail
    $webUrl    = $SiteMeta.WebUrl
    $webTitle  = $SiteMeta.WebTitle
    $listTitle = $Library.ListTitle

    Write-Log -Level INFO -Message "Processing sync for site: '$webTitle' -> Library: '$listTitle'"

    # Check local registry to see if this library is already synced
    $syncedLibs = Get-LocalSyncedLibraries
    $alreadySynced = $false
    $matchingPath  = ""

    foreach ($sl in $syncedLibs) {
        if ($sl.UrlNamespace -and ($sl.UrlNamespace -match [regex]::Escape($webUrl) -or $sl.UrlNamespace -match [regex]::Escape($listTitle))) {
            $alreadySynced = $true
            $matchingPath  = $sl.MountPoint
            break
        }
    }

    if ($alreadySynced -and -not $ForceSync) {
        Write-Log -Level SUCCESS -Message "Already synced locally at: $matchingPath"
        $Script:SummaryList.Add([PSCustomObject]@{
            SiteName  = $webTitle
            Library   = $listTitle
            Status    = "Already Synced"
            Details   = $matchingPath
        })
        return
    }

    # Construct the canonical odopen:// sync protocol URI
    # Format: odopen://sync/?siteId=%7B{SiteId}%7D&webId=%7B{WebId}%7D&listId=%7B{ListId}%7D&userEmail=...&webUrl=...&webTitle=...&listTitle=...
    $encEmail = [System.Uri]::EscapeDataString($userEmail)
    $encUrl   = [System.Uri]::EscapeDataString($webUrl)
    $encWeb   = [System.Uri]::EscapeDataString($webTitle)
    $encList  = [System.Uri]::EscapeDataString($listTitle)

    $syncUri = "odopen://sync/?siteId=%7B$siteId%7D&webId=%7B$webId%7D&listId=%7B$listId%7D&userEmail=$encEmail&webUrl=$encUrl&webTitle=$encWeb&listTitle=$encList"

    Write-Log -Level INFO -Message "Invoking OneDrive sync protocol handler..."
    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$syncUri`""
        Write-Log -Level SUCCESS -Message "Sync command successfully delivered to OneDrive client for '$listTitle'!"
        
        $Script:SummaryList.Add([PSCustomObject]@{
            SiteName  = $webTitle
            Library   = $listTitle
            Status    = "Initiated"
            Details   = "Handed off to OneDrive client ($userEmail)"
        })

        # Add pause to allow OneDrive client to process the registration cleanly
        Start-Sleep -Seconds 3
    }
    catch {
        Write-Log -Level ERROR -Message "Failed to invoke OneDrive sync protocol: $($_.Exception.Message)"
        $Script:SummaryList.Add([PSCustomObject]@{
            SiteName  = $webTitle
            Library   = $listTitle
            Status    = "Failed"
            Details   = $_.Exception.Message
        })
    }
}

# ==============================================================================
# SUMMARY & FINAL REPORTING (Requirement 4 & CleanCode)
# ==============================================================================
function Show-ExecutionSummary {
    $duration = (Get-Date) - $Script:StartTime
    $durStr   = [string]::Format("{0:d2}:{1:d2}:{2:d2}", [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds)

    $totalCount     = $Script:SummaryList.Count
    $initiatedCount = ($Script:SummaryList | Where-Object { $_.Status -eq 'Initiated' }).Count
    $syncedCount    = ($Script:SummaryList | Where-Object { $_.Status -eq 'Already Synced' }).Count
    $failedCount    = ($Script:SummaryList | Where-Object { $_.Status -eq 'Failed' }).Count

    Write-Host "`n========================================================================================" -ForegroundColor Cyan
    Write-Host "                                 EXECUTION SUMMARY REPORT                               " -ForegroundColor Cyan
    Write-Host "========================================================================================" -ForegroundColor Cyan
    Write-Host " Total Libraries Processed : $totalCount" -ForegroundColor White
    Write-Host " Sync Newly Initiated      : $initiatedCount" -ForegroundColor Green
    Write-Host " Already Synced Locally    : $syncedCount" -ForegroundColor Yellow
    Write-Host " Failed Registrations      : $failedCount" -ForegroundColor $(if ($failedCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host " Total Execution Time      : $durStr" -ForegroundColor White
    Write-Host " Execution Log File        : $Script:LogFilePath" -ForegroundColor Cyan
    Write-Host "========================================================================================" -ForegroundColor Cyan

    if ($Script:SummaryList.Count -gt 0) {
        $Script:SummaryList | Format-Table -Property @(
            @{ Label = "Site Name";    Expression = { $_.SiteName }; Width = 25 },
            @{ Label = "Library";      Expression = { $_.Library };  Width = 20 },
            @{ Label = "Status";       Expression = { $_.Status };   Width = 16 },
            @{ Label = "Details / Local Path"; Expression = { $_.Details } }
        ) | Out-String | Write-Host
    }

    Write-Host "Log details have been written to: $Script:LogFilePath`n" -ForegroundColor DarkGray
}

# ==============================================================================
# MAIN SCRIPT ENTRYPOINT
# ==============================================================================
try {
    # Initialize logging
    Initialize-LoggingSystem -TargetDir $LogDirectory

    $cleanHost = Get-SanitizedTenantHost -RawHost $TenantHost

    Write-Host "`n================================================================================" -ForegroundColor Cyan
    Write-Host " $($Script:Metadata.Company) - SHAREPOINT ONEDRIVE SYNC UTILITY (v$($Script:Metadata.Version))" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host " Target Tenant Host : https://$cleanHost" -ForegroundColor White
    Write-Host " Log Destination    : $Script:LogFilePath" -ForegroundColor White
    Write-Host " Execution Mode     : $(if ($AutoDiscover) { 'Auto-Discovery' } else { 'Direct Site Specification' })" -ForegroundColor White
    Write-Host "================================================================================`n" -ForegroundColor Cyan

    # Step 1: Validate environment and OneDrive client status
    Test-EnvironmentPrerequisites

    # Step 2: Pre-validate input parameters before initiating authentication
    $targetSiteUrls = [System.Collections.Generic.List[string]]::new()
    if ($SiteName -and $SiteName.Count -gt 0) {
        foreach ($sn in $SiteName) {
            $canonical = Resolve-SharePointSiteCanonicalUrl -InputSite $sn -CleanTenantHost $cleanHost
            $targetSiteUrls.Add($canonical)
        }
    } elseif (-not $AutoDiscover) {
        Show-ScriptHelp
        exit 0
    }

    # Step 3: Acquire Access Token (Interactive Device Code or Non-Interactive Entra ID)
    $token = Get-SharePointAccessToken -CleanTenantHost $cleanHost `
                                       -CustomTenantId $TenantId `
                                       -CustomClientId $ClientId `
                                       -CustomClientSecret $ClientSecret

    # Step 4: Perform site discovery if -AutoDiscover was requested
    if ($AutoDiscover) {
        $discoveredSites = Get-AccessibleSharePointSites -CleanTenantHost $cleanHost -Token $token

        if ($discoveredSites.Count -eq 0) {
            Write-Log -Level WARN -Message "No accessible SharePoint sites found via search or following APIs."
            Write-Host "Tip: You can specify target sites directly using: -SiteName 'site-name'`n" -ForegroundColor Yellow
            exit 0
        }

        Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host " DISCOVERED ACCESSIBLE SHAREPOINT SITES ($($discoveredSites.Count) found):" -ForegroundColor Cyan
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor Cyan
        for ($i = 0; $i -lt $discoveredSites.Count; $i++) {
            Write-Host (" [{0,2}] {1,-30} - {2}" -f ($i + 1), $discoveredSites[$i].Title, $discoveredSites[$i].Url) -ForegroundColor White
        }
        Write-Host "--------------------------------------------------------------------------------`n" -ForegroundColor Cyan

        if ($All) {
            Write-Log -Level INFO -Message "Switch -All specified. Selecting all $($discoveredSites.Count) discovered sites for synchronization."
            foreach ($s in $discoveredSites) { $targetSiteUrls.Add($s.Url) }
        }
        else {
            Write-Host "Select sites to sync:" -ForegroundColor Yellow
            Write-Host "  - Enter comma-separated numbers (e.g. 1, 3, 5)" -ForegroundColor Gray
            Write-Host "  - Type 'ALL' to synchronize all discovered sites" -ForegroundColor Gray
            Write-Host "  - Type 'Q' to quit" -ForegroundColor Gray
            $choice = Read-Host "`nYour selection"

            if ([string]::IsNullOrWhiteSpace($choice) -or ($choice.Trim().ToUpper() -eq 'Q')) {
                Write-Log -Level INFO -Message "User aborted site selection. Exiting."
                exit 0
            }

            if ($choice.Trim().ToUpper() -eq 'ALL') {
                foreach ($s in $discoveredSites) { $targetSiteUrls.Add($s.Url) }
            }
            else {
                $indexes = $choice.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries)
                foreach ($idxStr in $indexes) {
                    $parsedIdx = 0
                    if ([int]::TryParse($idxStr.Trim(), [ref]$parsedIdx) -and ($parsedIdx -ge 1) -and ($parsedIdx -le $discoveredSites.Count)) {
                        $targetSiteUrls.Add($discoveredSites[$parsedIdx - 1].Url)
                    } else {
                        Write-Log -Level WARN -Message "Invalid site index specified: '$idxStr'. Skipping."
                    }
                }
            }
        }
    }

    if ($targetSiteUrls.Count -eq 0) {
        Write-Log -Level WARN -Message "No valid SharePoint sites were selected for synchronization."
        exit 0
    }

    Write-Log -Level INFO -Message "Initiating synchronization for $($targetSiteUrls.Count) SharePoint site(s)..."

    # Step 4: Resolve metadata and trigger sync for each site
    foreach ($siteUrl in $targetSiteUrls) {
        try {
            Write-Host ""
            Write-Log -Level INFO -Message "=========================================================="
            Write-Log -Level INFO -Message "Target Site URL: $siteUrl"

            $meta = Get-SharePointLibraryMetadata -FullSiteUrl $siteUrl `
                                                  -Token $token `
                                                  -TargetLibraryTitle $DocumentLibraryTitle `
                                                  -FetchAllLibraries:$AllLibraries

            foreach ($lib in $meta.Libraries) {
                Invoke-OneDriveSyncRegistration -SiteMeta $meta -Library $lib -ForceSync:$Force
            }
        }
        catch {
            Write-Log -Level ERROR -Message "Failed to process site '$siteUrl': $($_.Exception.Message)"
            $Script:SummaryList.Add([PSCustomObject]@{
                SiteName  = $siteUrl
                Library   = if ($DocumentLibraryTitle) { $DocumentLibraryTitle } else { "Default" }
                Status    = "Failed"
                Details   = $_.Exception.Message
            })
        }
    }

    # Step 5: Render final summary report
    Show-ExecutionSummary
}
catch {
    Write-Host ""
    Write-Log -Level ERROR -Message "Execution terminated due to an unhandled error: $($_.Exception.Message)"
    if ($Script:LogFilePath) {
        Write-Host "Detailed log is available at: $Script:LogFilePath" -ForegroundColor DarkGray
    }
    exit 1
}
