<#
.SYNOPSIS
    Enterprise Helpdesk utility to reset Microsoft Teams cache, WAM token broker, and identity credentials.

.DESCRIPTION
    Reset-TeamsIdentityCache.ps1 is a production-grade PowerShell automation tool designed
    for Microsoft 365 Helpdesk and IT Support teams. It resolves persistent Teams authentication
    glitches, credential loops (AADSTS / SSO errors), stale tenant accounts, token expiration,
    and profile corruption across Microsoft Teams (both New MSTeams MSIX and Classic Teams).

    Key Capabilities:
    1. Interactive and automated execution with detailed parameter control.
    2. Comprehensive feature catalog display via -ListFeatures.
    3. Safe termination of Teams desktop clients and orphaned WebView2 identity broker processes.
    4. Deep cache cleansing for New Microsoft Teams (MSTeams AppX/MSIX LocalCache, TempState, AC).
    5. AppX package state reset via Reset-AppxPackage (Windows 10/11) with fallback.
    6. Classic Teams cache purge (blob_storage, Cache, databases, GPUCache, IndexedDB, Local Storage, tmp).
    7. Microsoft 365 IdentityCache, OneAuth, and Web Account Manager (WAM) TokenBroker cache purge.
    8. Windows Credential Manager cleanup targeting Teams generic credentials and stale tokens.
    9. Optional automatic relaunch of Microsoft Teams upon completion.
    10. Safe simulation mode (-WhatIf / Dry Run) for auditing without deleting data.
    11. Language-independent Windows architecture (English and Polish OS compatibility).
    12. Language-independent Administrator verification and optional automatic UAC elevation (-Elevate).
    13. Optional Microsoft 365 / Exchange Online administrator session connectivity.
    14. Detailed telemetry and structured audit logging to C:\Logs (fallback to %LOCALAPPDATA%\Logs).

.PARAMETER Reset
    Switch parameter. Executes the full, recommended end-to-end reset procedure:
    terminates Teams processes, clears New Teams and Classic Teams caches, purges
    IdentityCache, OneAuth, and TokenBroker caches, and cleans Teams credentials.
    Aliases: -Run, -Execute, -All.

.PARAMETER ClearTeamsNewCache
    Switch parameter. Clears New Microsoft Teams (MSIX / AppX) LocalCache, TempState,
    and resets the MSTeams package state.
    Aliases: -ClearNewTeams, -NewTeams.

.PARAMETER ClearTeamsClassicCache
    Switch parameter. Clears legacy Classic Microsoft Teams (%APPDATA%\Microsoft\Teams) cache subfolders.
    Aliases: -ClearClassicTeams, -ClassicTeams.

.PARAMETER ClearIdentityCache
    Switch parameter. Clears Microsoft 365 IdentityCache (%LOCALAPPDATA%\Microsoft\IdentityCache).
    Aliases: -IdentityCache, -Identity.

.PARAMETER ClearOneAuthCache
    Switch parameter. Clears OneAuth modern authentication token cache (%LOCALAPPDATA%\Microsoft\OneAuth).
    Aliases: -OneAuthCache, -OneAuth.

.PARAMETER ClearTokenBrokerCache
    Switch parameter. Clears Web Account Manager (WAM) Token Broker cache (%LOCALAPPDATA%\Microsoft\TokenBroker\Cache).
    Aliases: -TokenBrokerCache, -TokenBroker, -WAM.

.PARAMETER ClearCredentials
    Switch parameter. Enumerates and removes cached Microsoft Teams generic credentials
    from the Windows Credential Manager.
    Aliases: -ClearCreds, -Credentials.

.PARAMETER CredentialPattern
    Pattern used to match credential target names in Windows Credential Manager when -ClearCredentials is used.
    Default: 'teams' (matches msteams*, *teams*, etc.).
    Aliases: -CredFilter, -Pattern.

.PARAMETER KillTeamsProcesses
    Switch parameter. Safely terminates active Teams processes (ms-teams.exe, teams.exe, and related webview2).
    Aliases: -KillProcesses, -StopTeams.

.PARAMETER RestartTeams
    Switch parameter. Automatically relaunches Microsoft Teams after cache reset completes.
    Aliases: -Relaunch, -StartTeams.

.PARAMETER Force
    Switch parameter. Suppresses interactive user confirmation prompts prior to executing reset actions.
    Aliases: -f, -Quiet, -Silent.

.PARAMETER WhatIf
    Switch parameter. Simulates execution without deleting files or modifying credentials (Dry Run mode).
    Aliases: -DryRun, -Simulate.

.PARAMETER ListFeatures
    Switch parameter. Displays a detailed breakdown of all script features, target paths, and options.
    Aliases: -ShowFeatures, -Features.

.PARAMETER LogDirectory
    Destination directory for structured execution log files.
    Default: 'C:\Logs' (automatically falls back to %LOCALAPPDATA%\Logs if C:\Logs is inaccessible).
    Aliases: -LogPath, -LogDir.

.PARAMETER Elevate
    Switch parameter. Automatically requests Administrator elevation via UAC prompt if not already elevated.
    Aliases: -RunAsAdmin, -Admin.

.PARAMETER ConnectM365
    Switch parameter. Prompts to connect to Microsoft 365 / Exchange Online administrator session.
    Aliases: -ConnectExchange, -M365Admin.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used for explicit authentication with Microsoft 365 / Exchange Online.
    Aliases: -AdminUPN, -AdminEmail.

.PARAMETER SkipM365Connection
    Switch parameter. Explicitly suppresses Microsoft 365 administrator connection prompts.
    Aliases: -NoM365, -SkipExchange.

.PARAMETER Help
    Switch parameter. Displays comprehensive console help, parameter descriptions, prerequisites, and examples.
    Aliases: -h.

.EXAMPLE
    # Display help and usage instructions (also shown when run without parameters)
    .\Reset-TeamsIdentityCache.ps1 -Help

.EXAMPLE
    # Display detailed list of all features and architectural options
    .\Reset-TeamsIdentityCache.ps1 -ListFeatures

.EXAMPLE
    # Perform full recommended Teams identity & cache cleanup with confirmation prompt
    .\Reset-TeamsIdentityCache.ps1 -Reset

.EXAMPLE
    # Perform full cleanup silently and automatically relaunch Microsoft Teams
    .\Reset-TeamsIdentityCache.ps1 -Reset -RestartTeams -Force

.EXAMPLE
    # Perform dry-run simulation of the full cleanup without modifying any files or credentials
    .\Reset-TeamsIdentityCache.ps1 -Reset -WhatIf

.EXAMPLE
    # Granularly clear only New Teams cache and WAM TokenBroker cache
    .\Reset-TeamsIdentityCache.ps1 -ClearTeamsNewCache -ClearTokenBrokerCache -KillTeamsProcesses

.EXAMPLE
    # Clear Teams credentials from Windows Credential Manager silently
    .\Reset-TeamsIdentityCache.ps1 -ClearCredentials -Force

.EXAMPLE
    # Run full reset with optional Microsoft 365 Exchange Online administrator authentication
    .\Reset-TeamsIdentityCache.ps1 -Reset -ConnectM365 -AdminUserPrincipalName "admin@contoso.com"

.NOTES
    Script Name   : Reset-TeamsIdentityCache.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.0.0
    Release Date  : 2026-10-19
    License       : MIT
    Target OS     : Windows 10, Windows 11, Windows Server (with Desktop Experience)
    Language      : Polish & English Windows OS compatible (language-independent SIDs & paths)
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # --- Execution Trigger Modes ---
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('Run', 'Execute', 'All')]
    [switch]$Reset,

    # --- Granular Feature Switches ---
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('ClearNewTeams', 'NewTeams')]
    [switch]$ClearTeamsNewCache,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('ClearClassicTeams', 'ClassicTeams')]
    [switch]$ClearTeamsClassicCache,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('IdentityCache', 'Identity')]
    [switch]$ClearIdentityCache,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('OneAuthCache', 'OneAuth')]
    [switch]$ClearOneAuthCache,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('TokenBrokerCache', 'TokenBroker', 'WAM')]
    [switch]$ClearTokenBrokerCache,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('ClearCreds', 'Credentials')]
    [switch]$ClearCredentials,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('CredFilter', 'Pattern')]
    [string]$CredentialPattern = 'teams',

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('KillProcesses', 'StopTeams')]
    [switch]$KillTeamsProcesses,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('Relaunch', 'StartTeams')]
    [switch]$RestartTeams,

    # --- Operational Flags ---
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('f', 'Quiet', 'Silent')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('DryRun', 'Simulate')]
    [switch]$WhatIf,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('RunAsAdmin', 'Admin')]
    [switch]$Elevate,

    # --- Telemetry & Logging ---
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('LogPath', 'LogDir')]
    [string]$LogDirectory = 'C:\Logs',

    # --- Microsoft 365 Admin Connection (Standard 14) ---
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('ConnectExchange', 'M365Admin')]
    [switch]$ConnectM365,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('AdminUPN', 'AdminEmail')]
    [string]$AdminUserPrincipalName,

    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ResetMode')]
    [Alias('NoM365', 'SkipExchange')]
    [switch]$SkipM365Connection,

    # --- Feature Discovery Mode ---
    [Parameter(ParameterSetName = 'ListFeaturesMode')]
    [Alias('ShowFeatures', 'Features')]
    [switch]$ListFeatures,

    # --- Help Mode ---
    [Parameter(ParameterSetName = 'HelpMode')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# GLOBAL SCRIPT METADATA & CONFIGURATION (CleanCode & Standards)
# ==============================================================================
$Script:Metadata = @{
    Name        = 'Reset-TeamsIdentityCache.ps1'
    Version     = '1.0.0'
    Author      = 'Roman Pindela'
    Email       = 'roman.pindela@gmail.com'
    GitHub      = 'https://github.com/romanpindela'
    ReleaseDate = '2026-10-19'
    License     = 'MIT'
    TargetOS    = 'Windows 10, Windows 11, Windows Server'
}

# Execution tracking
$Script:StartTime   = Get-Date
$Script:LogFilePath = $null
$Script:SummaryList = [System.Collections.Generic.List[PSObject]]::new()

# ==============================================================================
# LOGGING SYSTEM (CleanCode & Standards)
# ==============================================================================
function Initialize-LoggingSystem {
    [CmdletBinding()]
    param([string]$TargetDir)

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $computer  = $env:COMPUTERNAME
    $user      = $env:USERNAME
    $logName   = "${timestamp}-Reset-TeamsIdentityCache-${computer}-${user}.log"

    $resolvedDir = $TargetDir
    try {
        if (-not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        # Fallback to LocalAppData\Logs if C:\Logs is inaccessible or permission is denied
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
GitHub           : $($Script:Metadata.GitHub)
Host / User      : $computer / $user
Target Mode      : $(if ($WhatIf) { "SIMULATION (WhatIf / Dry Run)" } else { "EXECUTION" })
Log Path         : $($Script:LogFilePath)
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

function Record-Summary {
    param(
        [string]$Component,
        [string]$Action,
        [string]$Details,
        [string]$Status
    )

    $item = [PSCustomObject]@{
        Component = $Component
        Action    = $Action
        Details   = $Details
        Status    = $Status
    }
    $Script:SummaryList.Add($item)
}

# ==============================================================================
# FEATURE CATALOG DEFINITION (Requirement 3: List Features)
# ==============================================================================
function Get-ScriptFeaturesCatalog {
    return @(
        [PSCustomObject]@{
            Id          = 1
            Feature     = 'Process Termination'
            Parameter   = '-KillTeamsProcesses'
            Target      = 'ms-teams.exe, teams.exe, msedgewebview2.exe'
            Description = 'Gracefully stops running Microsoft Teams instances (Classic & New) and associated background WebView2 identity broker processes to release locked cache databases.'
        },
        [PSCustomObject]@{
            Id          = 2
            Feature     = 'New Teams Cache Reset'
            Parameter   = '-ClearTeamsNewCache'
            Target      = '%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache, TempState, AC'
            Description = 'Purges New Microsoft Teams (MSIX/AppX) cache databases, web storage, and invokes Reset-AppxPackage when available in Windows 10/11 to restore pristine package state.'
        },
        [PSCustomObject]@{
            Id          = 3
            Feature     = 'Classic Teams Cache Reset'
            Parameter   = '-ClearTeamsClassicCache'
            Target      = '%APPDATA%\Microsoft\Teams\*'
            Description = 'Purges legacy Win32/Electron Teams cache subdirectories: blob_storage, Cache, databases, GPUCache, IndexedDB, Local Storage, and tmp.'
        },
        [PSCustomObject]@{
            Id          = 4
            Feature     = 'IdentityCache Purge'
            Parameter   = '-ClearIdentityCache'
            Target      = '%LOCALAPPDATA%\Microsoft\IdentityCache'
            Description = 'Removes cached Microsoft 365 OAuth tokens, tenant metadata, and cached identity records responsible for account sign-in loops and stale tenant switching.'
        },
        [PSCustomObject]@{
            Id          = 5
            Feature     = 'OneAuth Cache Purge'
            Parameter   = '-ClearOneAuthCache'
            Target      = '%LOCALAPPDATA%\Microsoft\OneAuth'
            Description = 'Clears Microsoft modern authentication (OneAuth) client token repositories and session contexts shared across Microsoft 365 desktop productivity apps.'
        },
        [PSCustomObject]@{
            Id          = 6
            Feature     = 'WAM TokenBroker Cache Purge'
            Parameter   = '-ClearTokenBrokerCache'
            Target      = '%LOCALAPPDATA%\Microsoft\TokenBroker\Cache'
            Description = 'Purges Windows Web Account Manager (WAM) TokenBroker cache files to eliminate corrupted broker tickets, AADSTS refresh failures, and 0x80070002 errors.'
        },
        [PSCustomObject]@{
            Id          = 7
            Feature     = 'Credential Manager Cleanup'
            Parameter   = '-ClearCredentials'
            Target      = 'Windows Credential Manager (Generic: msteams*, TokenCache*)'
            Description = 'Enumerates and safely deletes cached Microsoft Teams generic credentials and tokens from the Windows Credential Manager without impacting Git or Windows credentials.'
        },
        [PSCustomObject]@{
            Id          = 8
            Feature     = 'Full End-to-End Reset'
            Parameter   = '-Reset'
            Target      = 'All components (1 through 7 above)'
            Description = 'Executes the complete recommended Microsoft 365 Helpdesk remediation workflow: stops processes, cleans all caches (New/Classic/Identity/OneAuth/WAM), and wipes Teams credentials.'
        },
        [PSCustomObject]@{
            Id          = 9
            Feature     = 'Automatic Teams Relaunch'
            Parameter   = '-RestartTeams'
            Target      = 'ms-teams: URI or Teams executable'
            Description = 'Automatically detects the installed Microsoft Teams edition and relaunches the application with a clean identity state once maintenance finishes.'
        },
        [PSCustomObject]@{
            Id          = 10
            Feature     = 'Simulation / Dry Run'
            Parameter   = '-WhatIf'
            Target      = 'Non-destructive audit'
            Description = 'Audits all target files, directories, processes, and credentials without making any system changes, displaying what would be removed.'
        },
        [PSCustomObject]@{
            Id          = 11
            Feature     = 'UAC Auto-Elevation'
            Parameter   = '-Elevate'
            Target      = 'Elevated Administrator Token'
            Description = 'Automatically prompts for UAC elevation and relaunches the script using language-independent Well-Known SIDs if elevated privileges are requested.'
        },
        [PSCustomObject]@{
            Id          = 12
            Feature     = 'M365 Admin Integration'
            Parameter   = '-ConnectM365'
            Target      = 'Exchange Online / Graph PowerShell'
            Description = 'Enables Helpdesk administrators to connect to Microsoft 365 / Exchange Online for tenant diagnostics, mailbox checks, and license verification.'
        },
        [PSCustomObject]@{
            Id          = 13
            Feature     = 'Enterprise Audit Logging'
            Parameter   = '-LogDirectory'
            Target      = 'C:\Logs (Fallback: %LOCALAPPDATA%\Logs)'
            Description = 'Records full execution telemetry, timestamps, status codes, and summary tables to persistent log files for Helpdesk ticketing compliance.'
        }
    )
}

function Show-FeaturesDetails {
    Write-Host @"
================================================================================
  $($Script:Metadata.Name) - DETAILED FEATURES CATALOG
  Version: $($Script:Metadata.Version) | Author: $($Script:Metadata.Author)
================================================================================
"@ -ForegroundColor Cyan

    $features = Get-ScriptFeaturesCatalog

    foreach ($f in $features) {
        Write-Host "[$($f.Id)] $($f.Feature)" -ForegroundColor Yellow
        Write-Host "    Parameter   : $($f.Parameter)" -ForegroundColor Green
        Write-Host "    Target Path : $($f.Target)" -ForegroundColor Gray
        Write-Host "    Description : $($f.Description)" -ForegroundColor White
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    }

    Write-Host @"
QUICK FEATURE INVOCATION EXAMPLES:
  # 1. Full Recommended Reset:
  .\Reset-TeamsIdentityCache.ps1 -Reset

  # 2. Granular Cache Clean (New Teams + WAM Broker):
  .\Reset-TeamsIdentityCache.ps1 -ClearTeamsNewCache -ClearTokenBrokerCache -KillTeamsProcesses

  # 3. Credential Manager Only:
  .\Reset-TeamsIdentityCache.ps1 -ClearCredentials -Force

  # 4. Dry Run (Test what will be cleaned):
  .\Reset-TeamsIdentityCache.ps1 -Reset -WhatIf
================================================================================
"@ -ForegroundColor Cyan
}

# ==============================================================================
# SCRIPT HELP DISPLAY (Standard 1, 7, 10, 13, 14)
# ==============================================================================
function Show-ScriptHelp {
    Write-Host @"
================================================================================
  SCRIPT INFORMATION: $($Script:Metadata.Name) v$($Script:Metadata.Version)
  Author       : $($Script:Metadata.Author)
  Email        : $($Script:Metadata.Email)
  GitHub       : $($Script:Metadata.GitHub)
  Release Date : $($Script:Metadata.ReleaseDate)
  License      : $($Script:Metadata.License)
================================================================================

SYNOPSIS:
  Enterprise-grade Helpdesk utility to reset Microsoft Teams identity caches,
  WAM token broker sessions, and Credential Manager records to resolve sign-in loops,
  AADSTS authentication errors, stale tenant accounts, and profile corruption.

PREREQUISITES & ENVIRONMENT:
  1. Standard User Profile Scope:
     - Microsoft Teams caches and Windows Credential Manager entries reside within
       the active user's Windows profile (%LOCALAPPDATA% and %APPDATA%).
     - By default, run this script inside the user session experiencing the issue.
  2. Language Independence:
     - Fully compatible with Polish ("Administratorzy") and English ("Administrators")
       versions of Windows 10, 11, and Windows Server (utilizes Well-Known SIDs).
  3. Script Unblocking (After downloading from GitHub):
     - Before first execution, unblock the script to permit PowerShell execution:
       Unblock-File -Path .\Reset-TeamsIdentityCache.ps1

PARAMETERS:
  -Reset, -Run, -All
      Executes the full, recommended identity reset workflow across all modules.

  -ClearTeamsNewCache, -ClearNewTeams
      Cleans New Teams (MSTeams AppX) LocalCache, TempState, and triggers package reset.

  -ClearTeamsClassicCache, -ClearClassicTeams
      Cleans legacy Classic Teams (%APPDATA%\Microsoft\Teams) cache folders.

  -ClearIdentityCache, -Identity
      Clears %LOCALAPPDATA%\Microsoft\IdentityCache.

  -ClearOneAuthCache, -OneAuth
      Clears %LOCALAPPDATA%\Microsoft\OneAuth modern auth token repository.

  -ClearTokenBrokerCache, -TokenBroker, -WAM
      Clears %LOCALAPPDATA%\Microsoft\TokenBroker\Cache to resolve broker loops.

  -ClearCredentials, -ClearCreds
      Removes Teams generic credentials from Windows Credential Manager.

  -CredentialPattern <String>
      Search pattern for credentials in Credential Manager. Default: 'teams'.

  -KillTeamsProcesses, -StopTeams
      Gracefully stops ms-teams.exe, teams.exe, and orphaned webview2 broker processes.

  -RestartTeams, -Relaunch
      Automatically launches Microsoft Teams after cache reset completes.

  -Force, -f, -Quiet
      Suppresses interactive confirmation prompts.

  -WhatIf, -DryRun
      Simulates actions without deleting any files, processes, or credentials.

  -ListFeatures, -Features
      Lists detailed technical specifications and paths for all script features.

  -LogDirectory <String>
      Custom directory for execution logs. Default: 'C:\Logs' (fallback: %LOCALAPPDATA%\Logs).

  -Elevate, -RunAsAdmin
      Requests UAC Administrator elevation and relaunches the script.

  -ConnectM365, -ConnectExchange
      Connects to Microsoft 365 / Exchange Online administrator session.

  -AdminUserPrincipalName <String>
      Administrator UPN for Microsoft 365 / Exchange Online authentication.

  -SkipM365Connection
      Explicitly bypasses Microsoft 365 admin connection checks.

  -Help, -h
      Displays this comprehensive help documentation.

USAGE EXAMPLES:
  # 1. Display Help (default when executed without parameters):
  .\Reset-TeamsIdentityCache.ps1 -Help

  # 2. List all features with technical details:
  .\Reset-TeamsIdentityCache.ps1 -ListFeatures

  # 3. Full recommended cleanup with confirmation prompt:
  .\Reset-TeamsIdentityCache.ps1 -Reset

  # 4. Silent full cleanup and automatic Teams restart:
  .\Reset-TeamsIdentityCache.ps1 -Reset -RestartTeams -Force

  # 5. Simulation / Dry Run (Audit what would be deleted without changes):
  .\Reset-TeamsIdentityCache.ps1 -Reset -WhatIf

  # 6. Granular reset (New Teams + WAM + Credentials):
  .\Reset-TeamsIdentityCache.ps1 -ClearTeamsNewCache -ClearTokenBrokerCache -ClearCredentials -Force

  # 7. Helpdesk Admin with Exchange Online login:
  .\Reset-TeamsIdentityCache.ps1 -Reset -ConnectM365 -AdminUserPrincipalName "admin@contoso.com"
================================================================================
"@ -ForegroundColor Yellow
}

# ==============================================================================
# NO-PARAMETER & HELP GUARD (Standards 1 & 7)
# ==============================================================================
if ($Help.IsPresent -or ($PSBoundParameters.Count -eq 0)) {
    Show-ScriptHelp
    exit 0
}

if ($ListFeatures.IsPresent) {
    Show-FeaturesDetails
    exit 0
}

# ==============================================================================
# LANGUAGE-INDEPENDENT ADMIN VERIFICATION & OPTIONAL ELEVATION (Standards 9 & 11)
# ==============================================================================
function Test-IsAdministrator {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $adminSid  = New-Object Security.Principal.SecurityIdentifier([Security.Principal.WellKnownSidType]::BuiltInAdministratorsSid, $null)
    return $principal.IsInRole($adminSid)
}

if ($Elevate.IsPresent) {
    $isAdmin = Test-IsAdministrator
    if (-not $isAdmin) {
        Write-Host "Elevation requested. Requesting Administrator privileges..." -ForegroundColor Yellow

        $argumentsList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
        foreach ($param in $PSBoundParameters.GetEnumerator()) {
            if ($param.Key -eq 'Elevate') { continue }
            if ($param.Value -is [switch]) {
                if ($param.Value.IsPresent) {
                    $argumentsList += "-$($param.Key)"
                }
            } elseif ($param.Value -is [array]) {
                $argumentsList += "-$($param.Key)"
                $argumentsList += ($param.Value | ForEach-Object { "`"$_`"" }) -join ','
            } else {
                $argumentsList += "-$($param.Key)"
                $argumentsList += "`"$($param.Value)`""
            }
        }

        try {
            Start-Process powershell.exe -ArgumentList ($argumentsList -join ' ') -Verb RunAs
            exit 0
        }
        catch {
            Write-Error "Failed to elevate process. Please run PowerShell as Administrator."
            exit 1
        }
    }
}

# ==============================================================================
# INITIALIZE LOGGING & ENVIRONMENT
# ==============================================================================
Initialize-LoggingSystem -TargetDir $LogDirectory

Write-Log -Level INFO -Message "Starting $($Script:Metadata.Name) v$($Script:Metadata.Version)"
Write-Log -Level INFO -Message "Running as user: $env:COMPUTERNAME\$env:USERNAME (IsAdmin: $(Test-IsAdministrator))"

# Warn if running as elevated administrator when clearing user profile cache
if (Test-IsAdministrator) {
    Write-Log -Level WARN -Message "Running with Administrator privileges. Note that cache and credentials reside in the current user profile ($env:USERPROFILE)."
}

# ==============================================================================
# MICROSOFT 365 / EXCHANGE ONLINE ADMIN INTEGRATION (Standard 14)
# ==============================================================================
function Connect-M365HelpdeskSession {
    [CmdletBinding()]
    param(
        [string]$Upn,
        [switch]$Skip
    )

    if ($Skip) {
        Write-Log -Level DEBUG -Message "M365 admin connection explicitly skipped (-SkipM365Connection)."
        return
    }

    Write-Log -Level INFO -Message "Initiating Microsoft 365 / Exchange Online administrator connection..."
    try {
        if (-not (Get-Module -ListAvailable -Name "ExchangeOnlineManagement")) {
            Write-Log -Level WARN -Message "ExchangeOnlineManagement PowerShell module is not installed locally."
            Write-Log -Level INFO -Message "To install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
            Record-Summary -Component "M365 Connection" -Action "Connect-ExchangeOnline" -Details "Module missing" -Status "Skipped"
            return
        }

        if ($Upn) {
            Write-Log -Level INFO -Message "Connecting to Exchange Online as: $Upn"
            Connect-ExchangeOnline -UserPrincipalName $Upn -ShowBanner:$false -ErrorAction Stop
        } else {
            Write-Log -Level INFO -Message "Connecting to Exchange Online with interactive prompt..."
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }

        Write-Log -Level SUCCESS -Message "Successfully established Microsoft 365 administrator session."
        Record-Summary -Component "M365 Connection" -Action "Connect-ExchangeOnline" -Details "Connected as $(if ($Upn) {$Upn} else {'Interactive'})" -Status "Success"
    }
    catch {
        Write-Log -Level ERROR -Message "Failed to connect to Microsoft 365 Exchange Online: $($_.Exception.Message)"
        Record-Summary -Component "M365 Connection" -Action "Connect-ExchangeOnline" -Details $_.Exception.Message -Status "Failed"
    }
}

if ($ConnectM365 -or [string]::IsNullOrWhiteSpace($AdminUserPrincipalName) -eq $false) {
    Connect-M365HelpdeskSession -Upn $AdminUserPrincipalName -Skip:$SkipM365Connection
}

# ==============================================================================
# CONFIRMATION GUARD
# ==============================================================================
# Determine if any reset action is requested
$shouldClearNew      = $Reset -or $ClearTeamsNewCache
$shouldClearClassic  = $Reset -or $ClearTeamsClassicCache
$shouldClearIdentity = $Reset -or $ClearIdentityCache
$shouldClearOneAuth  = $Reset -or $ClearOneAuthCache
$shouldClearToken    = $Reset -or $ClearTokenBrokerCache
$shouldClearCreds    = $Reset -or $ClearCredentials
$shouldKillProcesses = $Reset -or $KillTeamsProcesses
$shouldRestart       = $RestartTeams

$hasAction = $shouldClearNew -or $shouldClearClassic -or $shouldClearIdentity -or
             $shouldClearOneAuth -or $shouldClearToken -or $shouldClearCreds -or
             $shouldKillProcesses -or $shouldRestart

if (-not $hasAction) {
    Write-Log -Level WARN -Message "No action parameters specified. Please specify -Reset or a granular switch."
    Write-Host "`nUse: .\Reset-TeamsIdentityCache.ps1 -Help to view available options." -ForegroundColor Yellow
    exit 0
}

if (-not $Force -and -not $WhatIf) {
    Write-Host "`n[CONFIRMATION REQUIRED]" -ForegroundColor Yellow
    Write-Host "This operation will reset Microsoft Teams cache, session tokens, and credentials for user: $env:USERNAME" -ForegroundColor White
    Write-Host "Selected actions:" -ForegroundColor Cyan
    if ($shouldKillProcesses) { Write-Host " - Terminate Microsoft Teams processes" -ForegroundColor Gray }
    if ($shouldClearNew)      { Write-Host " - Clear New Teams (MSTeams) cache and reset AppX package" -ForegroundColor Gray }
    if ($shouldClearClassic)  { Write-Host " - Clear Classic Teams (%APPDATA%\Microsoft\Teams) cache" -ForegroundColor Gray }
    if ($shouldClearIdentity) { Write-Host " - Purge Microsoft 365 IdentityCache" -ForegroundColor Gray }
    if ($shouldClearOneAuth)  { Write-Host " - Purge OneAuth token cache" -ForegroundColor Gray }
    if ($shouldClearToken)    { Write-Host " - Purge WAM TokenBroker cache" -ForegroundColor Gray }
    if ($shouldClearCreds)    { Write-Host " - Delete Teams credentials from Windows Credential Manager" -ForegroundColor Gray }
    if ($shouldRestart)       { Write-Host " - Relaunch Microsoft Teams after reset" -ForegroundColor Gray }

    $confirm = Read-Host "`nDo you want to proceed? (Y/N)"
    if ($confirm -notmatch '^[YyTt]$') {
        Write-Log -Level WARN -Message "Operation cancelled by user."
        exit 0
    }
}

# ==============================================================================
# HELPER FUNCTIONS (CleanCode & Safety)
# ==============================================================================
function Clear-FolderContentSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not (Test-Path -Path $Path)) {
        Write-Log -Level DEBUG -Message "$Description path does not exist: $Path"
        Record-Summary -Component $Description -Action "Clean Files" -Details "Path not present" -Status "Skipped"
        return
    }

    Write-Log -Level INFO -Message "Scanning $Description at: $Path"
    $items = @(Get-ChildItem -Path $Path -Force -Recurse -ErrorAction SilentlyContinue)

    if ($items.Count -eq 0) {
        Write-Log -Level INFO -Message "$Description is already clean."
        Record-Summary -Component $Description -Action "Clean Files" -Details "0 items" -Status "Clean"
        return
    }

    $fileCount = @($items | Where-Object { -not $_.PSIsContainer }).Count
    $dirCount  = @($items | Where-Object { $_.PSIsContainer }).Count

    Write-Log -Level INFO -Message "Found $fileCount file(s) and $dirCount folder(s) in $Description."

    if ($WhatIf) {
        Write-Log -Level WARN -Message "[WHATIF] Would delete $fileCount file(s) and $dirCount folder(s) in $Path"
        Record-Summary -Component $Description -Action "[WhatIf] Delete" -Details "$fileCount files, $dirCount dirs" -Status "Simulated"
        return
    }

    $deletedCount = 0
    $errorCount   = 0

    # Remove child files and folders
    Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop
            $deletedCount++
        }
        catch {
            $errorCount++
            Write-Log -Level WARN -Message "Unable to delete: $($_.FullName) ($($_.Exception.Message))"
        }
    }

    if ($errorCount -eq 0) {
        Write-Log -Level SUCCESS -Message "Successfully cleared $Description."
        Record-Summary -Component $Description -Action "Clean Files" -Details "$deletedCount root item(s) purged" -Status "Success"
    } else {
        Write-Log -Level WARN -Message "$Description partially cleared ($deletedCount succeeded, $errorCount locked/failed)."
        Record-Summary -Component $Description -Action "Clean Files" -Details "$deletedCount cleared, $errorCount locked" -Status "Partial"
    }
}

# ==============================================================================
# FEATURE 1: PROCESS TERMINATION
# ==============================================================================
function Stop-TeamsProcessesSafe {
    Write-Log -Level INFO -Message "Checking for running Microsoft Teams and related broker processes..."

    $processNames = @('ms-teams', 'teams')
    $targets = @(Get-Process -Name $processNames -ErrorAction SilentlyContinue)
    $targetPids = @($targets | Select-Object -ExpandProperty Id)

    # Quickly identify dedicated WebView2 broker processes spawned by Teams
    $wv2Pids = @()
    try {
        $wv2Cim = @(Get-CimInstance Win32_Process -Filter "Name = 'msedgewebview2.exe'" -ErrorAction SilentlyContinue | Where-Object {
            $_.ParentProcessId -in $targetPids -or $_.CommandLine -like "*teams*"
        })
        $wv2Pids = @($wv2Cim | Select-Object -ExpandProperty ProcessId)
    } catch { }

    $wv2Processes = @()
    if ($wv2Pids.Count -gt 0) {
        $wv2Processes = @(Get-Process -Id $wv2Pids -ErrorAction SilentlyContinue)
    }

    $allProcesses = @($targets) + @($wv2Processes)

    if ($allProcesses.Count -eq 0) {
        Write-Log -Level INFO -Message "No active Microsoft Teams processes detected."
        Record-Summary -Component "Process Termination" -Action "Stop-Process" -Details "No running processes" -Status "Clean"
        return
    }

    Write-Log -Level INFO -Message "Detected $($allProcesses.Count) process instance(s) to terminate."

    foreach ($proc in $allProcesses) {
        if ($WhatIf) {
            Write-Log -Level WARN -Message "[WHATIF] Would stop process: $($proc.Name) (PID: $($proc.Id))"
            Record-Summary -Component "Process Termination" -Action "[WhatIf] Stop" -Details "$($proc.Name) (PID: $($proc.Id))" -Status "Simulated"
            continue
        }

        try {
            Write-Log -Level INFO -Message "Stopping process: $($proc.Name) (PID: $($proc.Id))..."
            $proc | Stop-Process -Force -ErrorAction Stop
            Record-Summary -Component "Process Termination" -Action "Stop-Process" -Details "$($proc.Name) (PID: $($proc.Id))" -Status "Success"
        }
        catch {
            Write-Log -Level WARN -Message "Could not terminate process $($proc.Name) (PID: $($proc.Id)): $($_.Exception.Message)"
            Record-Summary -Component "Process Termination" -Action "Stop-Process" -Details "$($proc.Name) (PID: $($proc.Id))" -Status "Failed"
        }
    }

    if (-not $WhatIf) {
        Start-Sleep -Seconds 2
    }
}

if ($shouldKillProcesses) {
    Stop-TeamsProcessesSafe
}

# ==============================================================================
# FEATURE 2: NEW TEAMS (MSTEAMS MSIX/APPX) CACHE RESET
# ==============================================================================
function Reset-NewTeamsCacheSafe {
    Write-Log -Level INFO -Message "Processing New Microsoft Teams (MSTeams) cache reset..."

    $msteamsPkgPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Packages\MSTeams_8wekyb3d8bbwe"

    if (Test-Path -Path $msteamsPkgPath) {
        $localCacheDir = Join-Path -Path $msteamsPkgPath -ChildPath "LocalCache"
        $tempStateDir  = Join-Path -Path $msteamsPkgPath -ChildPath "TempState"
        $acDir         = Join-Path -Path $msteamsPkgPath -ChildPath "AC"

        Clear-FolderContentSafe -Path $localCacheDir -Description "New Teams LocalCache"
        Clear-FolderContentSafe -Path $tempStateDir  -Description "New Teams TempState"
        Clear-FolderContentSafe -Path $acDir         -Description "New Teams AC Cache"
    } else {
        Write-Log -Level DEBUG -Message "New Teams package folder not found at: $msteamsPkgPath"
        Record-Summary -Component "New Teams Cache" -Action "Locate Package" -Details "Package directory not found" -Status "Skipped"
    }

    # Attempt native AppX package reset via Reset-AppxPackage (Windows 10 2004+ and Windows 11)
    if (Get-Command -Name "Reset-AppxPackage" -ErrorAction SilentlyContinue) {
        $pkg = Get-AppxPackage -Name "*MSTeams*" -ErrorAction SilentlyContinue
        if ($pkg) {
            if ($WhatIf) {
                Write-Log -Level WARN -Message "[WHATIF] Would invoke Reset-AppxPackage for: $($pkg.PackageFamilyName)"
                Record-Summary -Component "New Teams AppX" -Action "[WhatIf] Reset-AppxPackage" -Details $pkg.PackageFamilyName -Status "Simulated"
            } else {
                try {
                    Write-Log -Level INFO -Message "Invoking Reset-AppxPackage for $($pkg.Name)..."
                    $pkg | Reset-AppxPackage -ErrorAction Stop
                    Write-Log -Level SUCCESS -Message "Reset-AppxPackage executed successfully."
                    Record-Summary -Component "New Teams AppX" -Action "Reset-AppxPackage" -Details $pkg.PackageFamilyName -Status "Success"
                }
                catch {
                    Write-Log -Level WARN -Message "Reset-AppxPackage failed: $($_.Exception.Message). LocalCache files were cleared manually."
                    Record-Summary -Component "New Teams AppX" -Action "Reset-AppxPackage" -Details $_.Exception.Message -Status "Warning"
                }
            }
        }
    }
}

if ($shouldClearNew) {
    Reset-NewTeamsCacheSafe
}

# ==============================================================================
# FEATURE 3: CLASSIC TEAMS CACHE RESET
# ==============================================================================
function Reset-ClassicTeamsCacheSafe {
    Write-Log -Level INFO -Message "Processing Classic Microsoft Teams cache reset..."

    $classicTeamsPath = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Teams"

    if (-not (Test-Path -Path $classicTeamsPath)) {
        Write-Log -Level DEBUG -Message "Classic Teams directory not present: $classicTeamsPath"
        Record-Summary -Component "Classic Teams Cache" -Action "Locate Directory" -Details "Directory not present" -Status "Skipped"
        return
    }

    $subFolders = @(
        'blob_storage',
        'Cache',
        'databases',
        'GPUCache',
        'IndexedDB',
        'Local Storage',
        'tmp',
        'Application Cache\Cache'
    )

    foreach ($folder in $subFolders) {
        $targetPath = Join-Path -Path $classicTeamsPath -ChildPath $folder
        Clear-FolderContentSafe -Path $targetPath -Description "Classic Teams ($folder)"
    }
}

if ($shouldClearClassic) {
    Reset-ClassicTeamsCacheSafe
}

# ==============================================================================
# FEATURE 4: MICROSOFT 365 IDENTITYCACHE PURGE
# ==============================================================================
function Reset-IdentityCacheSafe {
    Write-Log -Level INFO -Message "Processing Microsoft 365 IdentityCache reset..."
    $identityCachePath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\IdentityCache"
    Clear-FolderContentSafe -Path $identityCachePath -Description "Microsoft IdentityCache"
}

if ($shouldClearIdentity) {
    Reset-IdentityCacheSafe
}

# ==============================================================================
# FEATURE 5: ONEAUTH CACHE PURGE
# ==============================================================================
function Reset-OneAuthCacheSafe {
    Write-Log -Level INFO -Message "Processing OneAuth modern auth cache reset..."
    $oneAuthPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\OneAuth"
    Clear-FolderContentSafe -Path $oneAuthPath -Description "OneAuth Cache"
}

if ($shouldClearOneAuth) {
    Reset-OneAuthCacheSafe
}

# ==============================================================================
# FEATURE 6: WAM TOKENBROKER CACHE PURGE
# ==============================================================================
function Reset-TokenBrokerCacheSafe {
    Write-Log -Level INFO -Message "Processing Web Account Manager (WAM) TokenBroker cache reset..."
    $tokenBrokerCachePath = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\TokenBroker\Cache"
    Clear-FolderContentSafe -Path $tokenBrokerCachePath -Description "WAM TokenBroker Cache"
}

if ($shouldClearToken) {
    Reset-TokenBrokerCacheSafe
}

# ==============================================================================
# FEATURE 7: WINDOWS CREDENTIAL MANAGER CLEANUP
# ==============================================================================
function Reset-CredentialManagerEntriesSafe {
    [CmdletBinding()]
    param([string]$Pattern = 'teams')

    Write-Log -Level INFO -Message "Scanning Windows Credential Manager for entries matching '$Pattern'..."

    # Use cmdkey /list with locale-independent extraction (supports English 'Target:' and Polish 'Element docelowy:')
    $cmdkeyOutput = cmdkey.exe /list 2>&1

    if (-not $cmdkeyOutput) {
        Write-Log -Level WARN -Message "Unable to query Windows Credential Manager."
        Record-Summary -Component "Credential Manager" -Action "cmdkey /list" -Details "Query returned no output" -Status "Warning"
        return
    }

    $targets = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $cmdkeyOutput) {
        # Match English "Target:" or Polish "Element docelowy:"
        if ($line -match '^\s*(?:Target|Element docelowy)\s*:\s*(?:LegacyGeneric:target=)?(?<target>.+?)\s*$') {
            $extractedTarget = $Matches['target'].Trim()

            # Filter for requested pattern (case-insensitive)
            if ($extractedTarget -like "*$Pattern*" -or
                $extractedTarget -like "msteams*" -or
                $extractedTarget -like "Windows.Security.Authentication.Web.TokenCache*teams*") {
                $targets.Add($extractedTarget)
            }
        }
    }

    if ($targets.Count -eq 0) {
        Write-Log -Level INFO -Message "No stored Teams credentials found matching pattern '$Pattern'."
        Record-Summary -Component "Credential Manager" -Action "Enumerate" -Details "No matching entries" -Status "Clean"
        return
    }

    Write-Log -Level INFO -Message "Found $($targets.Count) matching credential target(s)."

    foreach ($t in $targets) {
        if ($WhatIf) {
            Write-Log -Level WARN -Message "[WHATIF] Would delete credential target: $t"
            Record-Summary -Component "Credential Manager" -Action "[WhatIf] Delete" -Details $t -Status "Simulated"
            continue
        }

        try {
            Write-Log -Level INFO -Message "Deleting credential: $t"
            $deleteResult = cmdkey.exe /delete:$t 2>&1
            Write-Log -Level SUCCESS -Message "Deleted: $t"
            Record-Summary -Component "Credential Manager" -Action "cmdkey /delete" -Details $t -Status "Success"
        }
        catch {
            Write-Log -Level WARN -Message "Failed to delete credential '$t': $($_.Exception.Message)"
            Record-Summary -Component "Credential Manager" -Action "cmdkey /delete" -Details "$t failed" -Status "Failed"
        }
    }
}

if ($shouldClearCreds) {
    Reset-CredentialManagerEntriesSafe -Pattern $CredentialPattern
}

# ==============================================================================
# FEATURE 8: RELAUNCH TEAMS
# ==============================================================================
function Start-TeamsApplicationSafe {
    Write-Log -Level INFO -Message "Attempting to relaunch Microsoft Teams..."

    if ($WhatIf) {
        Write-Log -Level WARN -Message "[WHATIF] Would launch Microsoft Teams application."
        Record-Summary -Component "Relaunch Teams" -Action "[WhatIf] Start" -Details "Teams application" -Status "Simulated"
        return
    }

    # Priority 1: Launch New Microsoft Teams via protocol handler
    try {
        Write-Log -Level INFO -Message "Invoking New Microsoft Teams protocol URI (ms-teams:)..."
        Start-Process -FilePath "ms-teams:" -ErrorAction Stop
        Write-Log -Level SUCCESS -Message "Launched New Microsoft Teams successfully."
        Record-Summary -Component "Relaunch Teams" -Action "Start-Process" -Details "Protocol ms-teams:" -Status "Success"
        return
    }
    catch {
        Write-Log -Level DEBUG -Message "ms-teams: protocol failed ($($_.Exception.Message)). Trying Classic Teams executable..."
    }

    # Priority 2: Classic Teams Update.exe runner
    $classicTeamsUpdate = Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\Teams\Update.exe"
    if (Test-Path -Path $classicTeamsUpdate) {
        try {
            Write-Log -Level INFO -Message "Launching Classic Teams from: $classicTeamsUpdate"
            Start-Process -FilePath $classicTeamsUpdate -ArgumentList '--processStart "Teams.exe"' -ErrorAction Stop
            Write-Log -Level SUCCESS -Message "Classic Microsoft Teams launched."
            Record-Summary -Component "Relaunch Teams" -Action "Start-Process" -Details "Classic Teams Update.exe" -Status "Success"
            return
        }
        catch {
            Write-Log -Level WARN -Message "Failed to launch Classic Teams: $($_.Exception.Message)"
        }
    }

    Write-Log -Level WARN -Message "Microsoft Teams client executable not found or could not be started."
    Record-Summary -Component "Relaunch Teams" -Action "Start-Process" -Details "Client not found" -Status "Warning"
}

if ($shouldRestart) {
    Start-TeamsApplicationSafe
}

# ==============================================================================
# EXECUTION SUMMARY REPORT (CleanCode & Standards)
# ==============================================================================
function Show-ExecutionSummaryReport {
    $duration = (Get-Date) - $Script:StartTime
    $durStr   = "{0:D2}m {1:D2}s" -f [int]$duration.TotalMinutes, $duration.Seconds

    Write-Host @"
================================================================================
                    EXECUTION SUMMARY & AUDIT REPORT
================================================================================
Script       : $($Script:Metadata.Name) v$($Script:Metadata.Version)
Executed By  : $env:COMPUTERNAME\$env:USERNAME
Duration     : $durStr
Log File     : $($Script:LogFilePath)
Mode         : $(if ($WhatIf) { "SIMULATION (WhatIf)" } else { "LIVE EXECUTION" })
================================================================================
"@ -ForegroundColor Cyan

    if ($Script:SummaryList.Count -gt 0) {
        $Script:SummaryList | Format-Table -Property Component, Action, Status, Details -AutoSize | Out-String | Write-Host -ForegroundColor White
    } else {
        Write-Host "No actions were recorded during this execution." -ForegroundColor Yellow
    }

    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Log -Level INFO -Message "Execution finished in $durStr. Log saved to: $($Script:LogFilePath)"
}

Show-ExecutionSummaryReport
