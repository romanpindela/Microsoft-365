<#
.SYNOPSIS
    Synchronizes files between an Office 365 user's OneDrive for Business folder and a local directory with real-time transfer progress, automated folder creation, detailed logging, desktop/webhook notifications, and support for interactive or non-interactive Entra ID authentication.

.DESCRIPTION
    Sync-OneDriveFolderAndNotify is an enterprise-grade PowerShell synchronization tool designed for Microsoft 365 and OneDrive for Business.
    It connects via the Microsoft Graph API / SDK and synchronizes files between a designated cloud folder and a local filesystem path.

    Key Features:
    - Dual Authentication Modes:
      * Interactive sign-in (delegated MFA-ready session with session reuse).
      * Non-Interactive / Unattended automation via Entra ID App Registration (Service Principal using Client Secret or Certificate).
    - Multi-Directional Sync:
      * TwoWay (Bidirectional): Synchronizes new and modified files in both directions (newer timestamp wins).
      * Upload: Synchronizes local files to OneDrive only.
      * Download: Synchronizes OneDrive files to local storage only.
    - Smart Folder Validation & Auto-Creation:
      * Detects missing OneDrive or local folders, displays a warning, and prompts the administrator or automatically provisions them (-AutoCreate, -AutoCreateOneDriveFolder, -AutoCreateLocalFolder).
    - High-Performance File Transfers:
      * Simple uploads for files up to 4 MB.
      * Chunked Upload Sessions (RFC compliant) for large files (> 4 MB up to 250 GB) with granular chunk-level progress.
    - Granular Progress Reporting:
      * Multi-level Write-Progress indicators showing file count, percentage, file names, bytes transferred, and upload session chunks.
    - Comprehensive Enterprise Logging:
      * Color-coded console feedback and persistent timestamped UTF-8 log files with execution summaries.
    - Notification System:
      * Native Windows desktop toast/balloon notification upon completion or failure.
      * Optional webhook integration for Microsoft Teams, Slack, or Power Automate.
    - Strict Security & Input Sanitization:
      * Regex email/UPN validation, illegal path character filtering, and traversal protection against malicious input.

.PARAMETER UserPrincipalName
    The User Principal Name (UPN) or email address of the target Office 365 / Entra ID user whose OneDrive is being synchronized.
    Aliases: -User, -UPN, -Email, -Account.

.PARAMETER OneDriveFolderPath
    The target folder path in the user's OneDrive (e.g. 'Backups/Finance' or 'Documents/Reports').
    To synchronize the root of OneDrive, specify '/' or leave blank.
    Aliases: -OneDriveFolder, -RemoteFolder, -CloudFolder, -CloudPath.

.PARAMETER LocalFolderPath
    The target directory path on the local filesystem (e.g. 'C:\Backups\Finance').
    Aliases: -LocalFolder, -LocalPath, -Path.

.PARAMETER Direction
    The direction of file synchronization.
    Values:
      - 'TwoWay': Bidirectional sync. Uploads newer/missing local files to OneDrive and downloads newer/missing cloud files to local.
      - 'Upload': One-way sync. Uploads new and updated local files to OneDrive.
      - 'Download': One-way sync. Downloads new and updated OneDrive files to the local folder.
    Default: 'TwoWay'.
    Aliases: -SyncDirection, -Mode.

.PARAMETER AutoCreate
    Switch parameter. Automatically provisions both missing OneDrive and Local folders without interactive prompting.

.PARAMETER AutoCreateOneDriveFolder
    Switch parameter. Automatically provisions the OneDrive cloud folder if it does not exist, without interactive prompting.
    Alias: -AutoCreateRemoteFolder.

.PARAMETER AutoCreateLocalFolder
    Switch parameter. Automatically provisions the local folder if it does not exist, without interactive prompting.

.PARAMETER Recurse
    Switch parameter. Synchronizes files in subdirectories recursively.
    Default: $true.

.PARAMETER TenantId
    The Entra ID (Azure AD) Directory (Tenant) ID for non-interactive daemon authentication.
    Aliases: -Tenant, -DirectoryId.

.PARAMETER ClientId
    The Entra ID Application (Client) ID for non-interactive daemon authentication.
    Aliases: -AppId, -ApplicationId.

.PARAMETER ClientSecret
    The Entra ID Client Secret string for non-interactive daemon authentication.
    Alias: -Secret.

.PARAMETER ClientSecretSecure
    Optional SecureString containing the Entra ID Client Secret.

.PARAMETER CertificateThumbprint
    The thumbprint of an X.509 Certificate installed in the local certificate store (Cert:\CurrentUser\My or Cert:\LocalMachine\My) used for non-interactive authentication.
    Alias: -Thumbprint.

.PARAMETER CertificateFilePath
    Path to a certificate file (.pfx or .cer) used for certificate-based non-interactive authentication.
    Alias: -CertPath.

.PARAMETER CertificatePassword
    Optional SecureString password for decrypting a .pfx certificate file.

.PARAMETER LogPath
    Custom path for the output log file. If omitted, a timestamped log file is created in a 'logs' subdirectory next to the script.

.PARAMETER NoLog
    Switch parameter. Disables writing execution logs to disk.

.PARAMETER WebhookUrl
    Optional incoming webhook URL (Microsoft Teams, Slack, or Power Automate) to receive an execution summary notification.
    Alias: -TeamsWebhookUrl.

.PARAMETER DisableNotification
    Switch parameter. Suppresses the native Windows desktop balloon/toast notification upon completion.
    Alias: -NoNotify.

.PARAMETER Force
    Switch parameter. Bypasses all interactive confirmation prompts (auto-approves folder creation, module installations, and file overwrites).

.PARAMETER Help
    Displays custom help, usage instructions, author information, Entra ID setup guide, and exits.
    Alias: -h.

.EXAMPLE
    # 1. Display full interactive help and Entra ID configuration guide
    .\Sync-OneDriveFolderAndNotify.ps1 -Help

.EXAMPLE
    # 2. Interactive Two-Way Synchronization with auto-folder creation
    .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "john.doe@company.com" `
        -OneDriveFolderPath "ProjectFiles" `
        -LocalFolderPath "C:\Data\ProjectFiles" `
        -Direction TwoWay `
        -AutoCreate

.EXAMPLE
    # 3. Unattended Nightly Backup (Upload Only) via Entra ID App Client Secret
    .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "backup.service@company.com" `
        -OneDriveFolderPath "DailyBackups/SQL" `
        -LocalFolderPath "D:\SQLBackups" `
        -Direction Upload `
        -AutoCreate `
        -TenantId "11111111-2222-3333-4444-555555555555" `
        -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -ClientSecret "SecretValueHere" `
        -WebhookUrl "https://outlook.office.com/webhook/..."

.EXAMPLE
    # 4. Download files from OneDrive to local directory with custom log path
    .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "sarah.smith@company.com" `
        -OneDriveFolderPath "Reports/2026" `
        -LocalFolderPath "C:\Reports\2026" `
        -Direction Download `
        -AutoCreateLocalFolder `
        -LogPath "C:\Logs\OneDriveSync_Reports.log"

.EXAMPLE
    # 5. Non-Interactive Sync via Installed Certificate
    .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "finance.user@company.com" `
        -OneDriveFolderPath "Audits" `
        -LocalFolderPath "E:\FinanceAudits" `
        -TenantId "11111111-2222-3333-4444-555555555555" `
        -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -CertificateThumbprint "A1B2C3D4E5F60718293A4B5C6D7E8F9012345678" `
        -AutoCreate

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 0)]
    [Alias('User', 'UPN', 'Email', 'Account')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 1)]
    [Alias('OneDriveFolder', 'RemoteFolder', 'CloudFolder', 'CloudPath')]
    [string]$OneDriveFolderPath,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 2)]
    [Alias('LocalFolder', 'LocalPath', 'Path')]
    [string]$LocalFolderPath,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('SyncDirection', 'Mode')]
    [ValidateSet('TwoWay', 'Upload', 'Download')]
    [string]$Direction = 'TwoWay',

    [Parameter(ParameterSetName = 'Default')]
    [switch]$AutoCreate,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AutoCreateRemoteFolder')]
    [switch]$AutoCreateOneDriveFolder,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$AutoCreateLocalFolder,

    [Parameter(ParameterSetName = 'Default')]
    [bool]$Recurse = $true,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('Tenant', 'DirectoryId')]
    [string]$TenantId,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AppId', 'ApplicationId')]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('Secret')]
    [string]$ClientSecret,

    [Parameter(ParameterSetName = 'Default')]
    [System.Security.SecureString]$ClientSecretSecure,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('Thumbprint')]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('CertPath')]
    [string]$CertificateFilePath,

    [Parameter(ParameterSetName = 'Default')]
    [System.Security.SecureString]$CertificatePassword,

    [Parameter(ParameterSetName = 'Default')]
    [string]$LogPath,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$NoLog,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('TeamsWebhookUrl')]
    [string]$WebhookUrl,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('NoNotify')]
    [switch]$DisableNotification,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# GLOBAL STATE & LOGGING INITIALIZATION
# ==============================================================================
$global:ScriptVersion = "1.0.0"
$global:ScriptAuthor  = "Roman Pindela"
$global:ScriptEmail   = "roman.pindela@gmail.com"
$global:ScriptGitHub  = "https://github.com/romanpindela"
$global:ScriptLogPath = $null
$global:AuthMode      = "Interactive"
$global:AccessToken   = $null
$global:UseRestDirect = $false

function Initialize-ScriptLogger {
    if ($NoLog) {
        $global:ScriptLogPath = $null
        return
    }

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $logDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $global:ScriptLogPath = Join-Path -Path $logDir -ChildPath "Sync-OneDriveFolderAndNotify_$timestamp.log"
    } else {
        $resolvedDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($resolvedDir) -and -not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force | Out-Null
        }
        $global:ScriptLogPath = $LogPath
    }

    # Write log header
    $header = @"
================================================================================
Sync-OneDriveFolderAndNotify Log Session
Started   : $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss UTCzzz"))
Version   : $global:ScriptVersion
Author    : $global:ScriptAuthor ($global:ScriptEmail)
================================================================================
"@
    try {
        Set-Content -Path $global:ScriptLogPath -Value $header -Encoding UTF8 -Force
    } catch {
        Write-Warning "Failed to initialize log file at '$global:ScriptLogPath': $_"
        $global:ScriptLogPath = $null
    }
}

function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug', 'Progress')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$Component = "Sync"
    )

    $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $logLine = "[$timeStr] [$($Level.ToUpper().PadRight(7))] [$Component] $Message"

    # Console display with appropriate colors
    switch ($Level) {
        'Success'  { Write-Host $logLine -ForegroundColor Green }
        'Warning'  { Write-Host $logLine -ForegroundColor Yellow }
        'Error'    { Write-Host $logLine -ForegroundColor Red }
        'Progress' { Write-Host $logLine -ForegroundColor Magenta }
        'Debug'    { Write-Host $logLine -ForegroundColor DarkGray }
        default    { Write-Host $logLine -ForegroundColor Cyan }
    }

    # Write to persistent log file if configured
    if ($global:ScriptLogPath) {
        try {
            Add-Content -Path $global:ScriptLogPath -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {
            # Ignore log write contention
        }
    }
}

# ==============================================================================
# NOTIFICATION SYSTEM (Desktop Toast & Webhook)
# ==============================================================================
function Send-SyncNotification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )

    # 1. Desktop Notification (Windows Forms / Balloon Tip)
    if (-not $DisableNotification) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop

            $notify = New-Object System.Windows.Forms.NotifyIcon
            $iconType = switch ($Type) {
                'Warning' { [System.Windows.Forms.ToolTipIcon]::Warning }
                'Error'   { [System.Windows.Forms.ToolTipIcon]::Error }
                default   { [System.Windows.Forms.ToolTipIcon]::Info }
            }
            $notify.Icon = [System.Drawing.SystemIcons]::Information
            $notify.BalloonTipIcon = $iconType
            $notify.BalloonTipTitle = $Title
            $notify.BalloonTipText = $Message
            $notify.Visible = $true
            $notify.ShowBalloonTip(5000)
            [System.Threading.Tasks.Task]::Delay(1200).Wait()
            $notify.Dispose()
        } catch {
            # Headless or non-GUI environments will safely skip balloon tips
            Write-LogMessage -Level Debug -Message "Desktop notification skipped (non-interactive or GUI service unavailable)."
        }
    }

    # 2. Webhook Notification (Teams / Slack / Power Automate)
    if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
        try {
            Write-LogMessage -Level Info -Message "Sending webhook notification to configured endpoint..."
            $color = switch ($Type) {
                'Warning' { "d9a74a" }
                'Error'   { "d9534f" }
                default   { "28a745" }
            }
            $payload = @{
                "@type"      = "MessageCard"
                "@context"   = "http://schema.org/extensions"
                "themeColor" = $color
                "summary"    = $Title
                "title"      = $Title
                "text"       = $Message
                "sections"   = @(
                    @{
                        "facts" = @(
                            @{ "name" = "Timestamp"; "value" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
                            @{ "name" = "Script";    "value" = "Sync-OneDriveFolderAndNotify v$global:ScriptVersion" }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 4

            Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body $payload -ErrorAction Stop | Out-Null
            Write-LogMessage -Level Success -Message "Webhook notification delivered successfully."
        } catch {
            Write-LogMessage -Level Warning -Message "Failed to deliver webhook notification: $_"
        }
    }
}

# ==============================================================================
# SCRIPT HELP & DOCUMENTATION FUNCTION
# ==============================================================================
function Show-ScriptHelp {
    $helpText = @'
================================================================================
  Sync-OneDriveFolderAndNotify.ps1 | Enterprise OneDrive Cloud Synchronizer
  Version : 1.0.0
  Author  : Roman Pindela
  Contact : roman.pindela@gmail.com | https://github.com/romanpindela
================================================================================

DESCRIPTION:
    Synchronizes files between an Office 365 user's OneDrive for Business folder
    and a local folder. Features real-time granular progress bars, multi-chunk
    large file upload sessions (>4 MB), auto-creation of missing folders,
    persistent logging, desktop notifications, and both interactive & non-interactive
    Entra ID daemon authentication.

--------------------------------------------------------------------------------
HOW TO CONFIGURE ENTRA ID FOR NON-INTERACTIVE / UNATTENDED AUTHENTICATION:
--------------------------------------------------------------------------------
To run this script unattended (e.g. in Windows Task Scheduler, CI/CD, or automated
backup pipelines) without interactive browser logins, configure an Entra ID App:

Step 1: Register an Application in Microsoft Entra ID
    1. Sign in to the Microsoft Entra admin center (https://entra.microsoft.com).
    2. Navigate to: Identity > Applications > App registrations > New registration.
    3. Name: 'OneDrive-Sync-Automation' (or any meaningful name).
    4. Supported account types: 'Accounts in this organizational directory only'.
    5. Redirect URI: Leave blank (not required for service daemons).
    6. Click 'Register'. Note down:
       - Application (client) ID
       - Directory (tenant) ID

Step 2: Assign Microsoft Graph API Application Permissions
    1. Under the registered app, navigate to: Manage > API permissions.
    2. Click 'Add a permission' > Select 'Microsoft Graph'.
    3. Choose 'Application permissions' (CRITICAL: Do NOT choose Delegated).
    4. Search and select the following permissions:
       * Files.ReadWrite.All  (Grants read and write access to all files across all user OneDrives)
       * User.Read.All        (Enables resolving target user IDs and OneDrive endpoints via UPN)
    5. Click 'Add permissions'.
    6. Click 'Grant admin consent for <Your Organization>' and verify green status checks.

Step 3: Generate Credentials (Choose Option A or Option B)
    Option A: Client Secret (Quickest)
      - Go to: Manage > Certificates & secrets > Client secrets > New client secret.
      - Set Description (e.g., 'OneDriveSyncSecret') and expiration.
      - Copy the 'Value' immediately (it will be permanently hidden upon leaving the page).
      - Pass to script via: -TenantId <ID> -ClientId <ID> -ClientSecret '<Value>'

    Option B: X.509 Certificate (Recommended for Enterprise Production)
      - Generate a self-signed certificate in PowerShell:
        $cert = New-SelfSignedCertificate -Subject 'CN=OneDriveSyncDaemon' -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy Exportable
        Export-Certificate -Cert $cert -FilePath 'C:\Certs\OneDriveSyncDaemon.cer'
      - In Entra ID App: Manage > Certificates & secrets > Certificates > Upload certificate.
      - Pass to script via: -TenantId <ID> -ClientId <ID> -CertificateThumbprint $cert.Thumbprint

--------------------------------------------------------------------------------
USAGE EXAMPLES:
--------------------------------------------------------------------------------
  # 1. Interactive Two-Way Synchronization with auto-folder creation:
  .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName 'john.doe@company.com' `
      -OneDriveFolderPath 'ProjectFiles' `
      -LocalFolderPath 'C:\Data\ProjectFiles' `
      -Direction TwoWay `
      -AutoCreate

  # 2. Unattended Nightly Backup (Upload Only) via Entra ID Client Secret:
  .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName 'backup.user@company.com' `
      -OneDriveFolderPath 'DailyBackups/SQL' `
      -LocalFolderPath 'D:\SQLBackups' `
      -Direction Upload `
      -AutoCreate `
      -TenantId '11111111-2222-3333-4444-555555555555' `
      -ClientId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' `
      -ClientSecret 'YourClientSecretHere' `
      -WebhookUrl 'https://outlook.office.com/webhook/...' 

  # 3. Download Files from OneDrive to Local directory with Custom Log Path:
  .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName 'sarah.smith@company.com' `
      -OneDriveFolderPath 'Reports/2026' `
      -LocalFolderPath 'C:\Reports\2026' `
      -Direction Download `
      -AutoCreateLocalFolder `
      -LogPath 'C:\Logs\OneDriveSync_Reports.log'

  # 4. Non-Interactive Sync via Certificate Thumbprint:
  .\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName 'finance@company.com' `
      -OneDriveFolderPath 'Audits' `
      -LocalFolderPath 'E:\FinanceAudits' `
      -TenantId '11111111-2222-3333-4444-555555555555' `
      -ClientId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' `
      -CertificateThumbprint 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678' `
      -AutoCreate

--------------------------------------------------------------------------------
PARAMETERS:
--------------------------------------------------------------------------------
  -UserPrincipalName, -User, -UPN  Office 365 user email/UPN owning the OneDrive.
  -OneDriveFolderPath, -RemoteFolder Target cloud folder path in OneDrive (root if empty).
  -LocalFolderPath, -LocalFolder    Local directory path for synchronization.
  -Direction, -SyncDirection        Sync direction: 'TwoWay' (default), 'Upload', 'Download'.
  -AutoCreate                       Auto-create both missing OneDrive and Local folders.
  -AutoCreateOneDriveFolder         Auto-create missing OneDrive folder only.
  -AutoCreateLocalFolder            Auto-create missing local folder only.
  -Recurse                          Recurse into subdirectories (Default: $true).
  -TenantId, -DirectoryId           Entra ID Tenant ID for unattended daemon authentication.
  -ClientId, -AppId                 Entra ID Application Client ID for daemon authentication.
  -ClientSecret, -Secret            Entra ID Client Secret for daemon authentication.
  -CertificateThumbprint            X.509 Certificate thumbprint for daemon authentication.
  -LogPath                          Custom log file destination path.
  -NoLog                            Disable file logging.
  -WebhookUrl                       Microsoft Teams / Slack incoming webhook endpoint.
  -DisableNotification, -NoNotify   Suppress Windows desktop balloon/toast notifications.
  -Force                            Bypass interactive confirmation prompts.
  -Help, -h                         Display this help manual.

SECURITY & UNBLOCKING:
  After downloading or cloning from GitHub, remember to unblock the PowerShell script:
  Unblock-File -Path .\Sync-OneDriveFolderAndNotify.ps1
================================================================================
'@
    Write-Host $helpText -ForegroundColor Cyan
}

# ------------------------------------------------------------------------------
# Check for -Help or missing parameters (Launch help when run without arguments)
# ------------------------------------------------------------------------------
if ($Help -or $PSBoundParameters.Count -eq 0) {
    Show-ScriptHelp
    exit 0
}

# ==============================================================================
# INPUT SANITIZATION & DEFENSIVE VALIDATION
# ==============================================================================
# 1. Validate UserPrincipalName
if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    Write-Error "Missing mandatory parameter: -UserPrincipalName. Run with -Help to view examples."
    exit 1
}

$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
if ($UserPrincipalName -notmatch $emailPattern) {
    Write-Error "Invalid format for -UserPrincipalName: '$UserPrincipalName'. Please specify a valid email or UPN (e.g. user@domain.com)."
    exit 1
}

# 2. Validate LocalFolderPath
if ([string]::IsNullOrWhiteSpace($LocalFolderPath)) {
    Write-Error "Missing mandatory parameter: -LocalFolderPath. Run with -Help to view examples."
    exit 1
}

# Detect malicious path traversals or invalid drive specifications
if ($LocalFolderPath.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) {
    Write-Error "Local folder path contains illegal filesystem characters: '$LocalFolderPath'."
    exit 1
}

# 3. Clean and normalize OneDrive folder path
$normalizedOneDrivePath = ""
if (-not [string]::IsNullOrWhiteSpace($OneDriveFolderPath)) {
    $cleanPath = $OneDriveFolderPath.Trim().Replace('\', '/')
    while ($cleanPath.Contains('//')) { $cleanPath = $cleanPath.Replace('//', '/') }
    $cleanPath = $cleanPath.Trim('/')
    if ($cleanPath -eq "root" -or $cleanPath -eq ".") {
        $normalizedOneDrivePath = ""
    } else {
        $normalizedOneDrivePath = $cleanPath
    }
}

# 4. Validate Entra ID parameter combinations
$hasSecret = (-not [string]::IsNullOrWhiteSpace($ClientSecret)) -or ($null -ne $ClientSecretSecure)
$hasCert   = (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) -or (-not [string]::IsNullOrWhiteSpace($CertificateFilePath))
$hasClient = -not [string]::IsNullOrWhiteSpace($ClientId)
$hasTenant = -not [string]::IsNullOrWhiteSpace($TenantId)

if ($hasSecret -or $hasCert -or $hasClient -or $hasTenant) {
    if (-not $hasTenant -or -not $hasClient) {
        Write-Error "Incomplete Entra ID authentication parameters. Both -TenantId and -ClientId must be provided for non-interactive login."
        exit 1
    }
    if (-not $hasSecret -and -not $hasCert) {
        Write-Error "Incomplete Entra ID authentication parameters. Provide either a Client Secret (-ClientSecret) or a Certificate (-CertificateThumbprint / -CertificateFilePath)."
        exit 1
    }
    if ($hasSecret -and $hasCert) {
        Write-Error "Conflicting parameters: Specify either a Client Secret OR a Certificate, not both."
        exit 1
    }
}

# Initialize Logger
Initialize-ScriptLogger

# ==============================================================================
# DISPLAY RUNTIME PARAMETERS BANNER
# ==============================================================================
$banner = @"
================================================================================
  Sync-OneDriveFolderAndNotify | Execution Parameters
================================================================================
Target User (UPN)     : $UserPrincipalName
OneDrive Path         : $(if ($normalizedOneDrivePath) { $normalizedOneDrivePath } else { '/ (OneDrive Root)' })
Local Path            : $LocalFolderPath
Sync Direction        : $Direction
Recursive Sync        : $Recurse
Auto-Create Remote    : $(if ($AutoCreate -or $AutoCreateOneDriveFolder) { 'Enabled' } else { 'Prompt if missing' })
Auto-Create Local     : $(if ($AutoCreate -or $AutoCreateLocalFolder) { 'Enabled' } else { 'Prompt if missing' })
Auth Mode             : $(if ($hasSecret) { 'Entra ID App (Client Secret)' } elseif ($hasCert) { 'Entra ID App (Certificate)' } else { 'Interactive / Delegated' })
Log Output            : $(if ($global:ScriptLogPath) { $global:ScriptLogPath } else { 'Disabled (-NoLog)' })
Notifications         : $(if ($DisableNotification) { 'Disabled' } else { 'Desktop Toast Enabled' })
Webhook               : $(if ($WebhookUrl) { 'Configured' } else { 'None' })
================================================================================
"@
Write-Host $banner -ForegroundColor Cyan
if ($global:ScriptLogPath) {
    Add-Content -Path $global:ScriptLogPath -Value $banner -Encoding UTF8 -ErrorAction SilentlyContinue
}

# ==============================================================================
# MICROSOFT GRAPH AUTHENTICATION ENGINE
# ==============================================================================
function Connect-GraphEnvironment {
    Write-LogMessage -Level Info -Message "Verifying Microsoft Graph connectivity..."

    # Determine if non-interactive daemon auth was requested
    if ($hasSecret -or $hasCert) {
        $global:AuthMode = "NonInteractive"
        Write-LogMessage -Level Info -Message "Initiating non-interactive Entra ID authentication (Tenant: $TenantId, Client: $ClientId)..."

        # Check if Microsoft.Graph.Authentication is installed
        $hasAuthModule = [bool](Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)

        if ($hasSecret) {
            $secretPlain = $ClientSecret
            if ($null -ne $ClientSecretSecure) {
                $secretPlain = [System.Net.NetworkCredential]::new('', $ClientSecretSecure).Password
            }

            if ($hasAuthModule) {
                try {
                    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                    $secretCred = [PSCredential]::new($ClientId, (ConvertTo-SecureString $secretPlain -AsPlainText -Force))
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -ClientSecretCredential $secretCred -NoWelcome -ErrorAction Stop
                    Write-LogMessage -Level Success -Message "Connected to Microsoft Graph via ClientSecretCredential."
                    return
                } catch {
                    Write-LogMessage -Level Warning -Message "Connect-MgGraph failed with ClientSecretCredential ($($_)). Falling back to direct REST token..."
                }
            }

            # Direct REST fallback (high resilience when Graph SDK module is not present or throws)
            try {
                $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                $body = @{
                    grant_type    = "client_credentials"
                    client_id     = $ClientId
                    client_secret = $secretPlain
                    scope         = "https://graph.microsoft.com/.default"
                }
                $tokenResp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ErrorAction Stop
                $global:AccessToken   = $tokenResp.access_token
                $global:UseRestDirect = $true
                Write-LogMessage -Level Success -Message "Acquired Microsoft Graph OAuth2 token directly via REST API."
                return
            } catch {
                Write-Error "Failed to acquire Entra ID OAuth token via REST: $_"
                exit 1
            }
        }

        if ($hasCert) {
            if (-not $hasAuthModule) {
                Write-Error "Certificate authentication requires the 'Microsoft.Graph.Authentication' module. Install it via: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
                exit 1
            }
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            try {
                if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
                } elseif (-not [string]::IsNullOrWhiteSpace($CertificateFilePath)) {
                    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateName $CertificateFilePath -NoWelcome -ErrorAction Stop
                }
                Write-LogMessage -Level Success -Message "Connected to Microsoft Graph via Certificate authentication."
                return
            } catch {
                Write-Error "Failed to connect to Microsoft Graph via certificate: $_"
                exit 1
            }
        }
    }

    # Interactive Authentication (Delegated)
    $global:AuthMode = "Interactive"
    $requiredScopes = @("Files.ReadWrite.All", "User.Read.All")

    if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
        Write-LogMessage -Level Warning -Message "Required module 'Microsoft.Graph.Authentication' is missing."
        $prompt = if ($Force) { "Y" } else { Read-Host "[?] Do you want to install 'Microsoft.Graph.Authentication' from PSGallery now? (Y/N)" }
        if ($prompt -match '^(Y|y|T|t)$') {
            try {
                Write-LogMessage -Level Info -Message "Installing Microsoft.Graph.Authentication from PSGallery..."
                Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
                Write-LogMessage -Level Success -Message "Module installed successfully."
            } catch {
                Write-Error "Failed to install Microsoft.Graph.Authentication: $_"
                exit 1
            }
        } else {
            Write-Error "Execution aborted. Microsoft.Graph.Authentication module is required for interactive login."
            exit 1
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Check active session reuse
    $mgContext = $null
    try {
        $mgContext = Get-MgContext -ErrorAction Stop
    } catch {
        $mgContext = $null
    }

    if ($null -ne $mgContext) {
        $scopes = @($mgContext.Scopes)
        $hasScopes = ($scopes -contains "Files.ReadWrite.All" -or $scopes -contains "Files.ReadWrite")
        if ($hasScopes) {
            Write-LogMessage -Level Success -Message "Reusing active Microsoft Graph session (Account: $($mgContext.Account))."
            return
        }
    }

    Write-LogMessage -Level Info -Message "Initiating interactive sign-in for Microsoft Graph..."
    try {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        Write-LogMessage -Level Success -Message "Successfully connected to Microsoft Graph."
    } catch {
        Write-Error "Failed to connect to Microsoft Graph interactively: $_"
        exit 1
    }
}

# ==============================================================================
# GRAPH API HTTP CLIENT WRAPPER (Auto-Retry on Throttling & SDK/REST abstraction)
# ==============================================================================
function Invoke-GraphRequestSafe {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [object]$Body = $null,

        [string]$ContentType = "application/json",

        [string]$OutFile = $null,

        [byte[]]$BinaryBody = $null,

        [hashtable]$CustomHeaders = @{}
    )

    $maxRetries = 4
    $retryCount = 0
    $success    = $false
    $result     = $null

    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            if ($global:UseRestDirect) {
                # Direct REST execution
                $headers = @{
                    "Authorization" = "Bearer $global:AccessToken"
                }
                foreach ($k in $CustomHeaders.Keys) { $headers[$k] = $CustomHeaders[$k] }

                $splat = @{
                    Method      = $Method
                    Uri         = $Uri
                    Headers     = $headers
                    ErrorAction = 'Stop'
                }

                if ($OutFile) {
                    $splat['OutFile'] = $OutFile
                } elseif ($null -ne $BinaryBody) {
                    $splat['Body']        = $BinaryBody
                    $splat['ContentType'] = $ContentType
                } elseif ($null -ne $Body) {
                    if ($Body -is [string]) {
                        $splat['Body'] = $Body
                    } else {
                        $splat['Body'] = ($Body | ConvertTo-Json -Depth 6)
                    }
                    $splat['ContentType'] = $ContentType
                }

                $result  = Invoke-RestMethod @splat
                $success = $true
            } else {
                # Microsoft Graph SDK execution
                $headers = @{}
                foreach ($k in $CustomHeaders.Keys) { $headers[$k] = $CustomHeaders[$k] }

                $splat = @{
                    Method      = $Method
                    Uri         = $Uri
                    Headers     = $headers
                    ErrorAction = 'Stop'
                }

                if ($OutFile) {
                    $splat['OutputFilePath'] = $OutFile
                } elseif ($null -ne $BinaryBody) {
                    $splat['Body']        = $BinaryBody
                    $splat['ContentType'] = $ContentType
                } elseif ($null -ne $Body) {
                    if ($Body -is [string]) {
                        $splat['Body'] = $Body
                    } else {
                        $splat['Body'] = ($Body | ConvertTo-Json -Depth 6)
                    }
                    $splat['ContentType'] = $ContentType
                }

                $result  = Invoke-MgGraphRequest @splat
                $success = $true
            }
        } catch {
            $ex = $_
            $statusCode = 0
            if ($ex.Exception -and $ex.Exception.Response) {
                $statusCode = [int]$ex.Exception.Response.StatusCode
            }

            # Handle HTTP 429 Too Many Requests / Throttling
            if ($statusCode -eq 429 -or ($ex.Message -match "429")) {
                $retryCount++
                $retryAfter = 5
                if ($ex.Exception.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]$ex.Exception.Response.Headers["Retry-After"]
                }
                Write-LogMessage -Level Warning -Message "Graph API throttled (HTTP 429). Retrying after $retryAfter seconds (Attempt $retryCount/$maxRetries)..."
                Start-Sleep -Seconds $retryAfter
            } else {
                # Non-retryable error
                throw $ex
            }
        }
    }

    if (-not $success) {
        throw "Failed Graph API call to $Uri after $maxRetries retries."
    }

    return $result
}

# ==============================================================================
# ONEDRIVE METADATA & FOLDER MANAGEMENT HELPERS
# ==============================================================================
function Get-UserOneDriveDrive {
    param([string]$UserUPN)

    Write-LogMessage -Level Info -Message "Retrieving OneDrive drive for user '$UserUPN'..."
    $uri = "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($UserUPN))/drive"

    try {
        $drive = Invoke-GraphRequestSafe -Method GET -Uri $uri
        Write-LogMessage -Level Success -Message "Located OneDrive: Id '$($drive.id)', Type '$($drive.driveType)'."
        return $drive
    } catch {
        Write-Error "Failed to access OneDrive for user '$UserUPN'. Ensure the account exists and has a provisioned OneDrive license. Details: $_"
        exit 1
    }
}

function Ensure-LocalFolderHierarchy {
    param([string]$Path)

    if (-not (Test-Path -Path $Path)) {
        Write-LogMessage -Level Warning -Message "Local target folder does not exist: '$Path'"
        $create = $false
        if ($AutoCreate -or $AutoCreateLocalFolder -or $Force) {
            $create = $true
            Write-LogMessage -Level Info -Message "Auto-creating local folder per parameter configuration..."
        } else {
            $prompt = Read-Host "[?] Local folder '$Path' does not exist. Create it now? (Y/N)"
            if ($prompt -match '^(Y|y|T|t)$') {
                $create = $true
            }
        }

        if ($create) {
            try {
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
                Write-LogMessage -Level Success -Message "Created local folder: '$Path'"
            } catch {
                Write-Error "Failed to create local directory '$Path': $_"
                exit 1
            }
        } else {
            Write-Error "Operation aborted. Local folder was not created."
            exit 1
        }
    } else {
        Write-LogMessage -Level Success -Message "Local target folder verified: '$Path'"
    }
}

function Ensure-OneDriveFolderHierarchy {
    param(
        [string]$DriveId,
        [string]$FolderPath
    )

    # If folder is root, it always exists
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $rootUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root"
        $rootItem = Invoke-GraphRequestSafe -Method GET -Uri $rootUri
        Write-LogMessage -Level Success -Message "Targeting OneDrive Root folder (DriveId: $DriveId)."
        return $rootItem
    }

    # Check if folder exists directly by path
    $escapedSegments = @($FolderPath.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) })
    $pathUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($escapedSegments -join '/')"

    $folderItem = $null
    $exists     = $false

    try {
        $folderItem = Invoke-GraphRequestSafe -Method GET -Uri $pathUrl
        if ($null -ne $folderItem -and $null -ne $folderItem.folder) {
            $exists = $true
            Write-LogMessage -Level Success -Message "OneDrive folder verified: '$FolderPath' (ItemId: $($folderItem.id))."
            return $folderItem
        }
    } catch {
        # 404 Not Found is expected if folder doesn't exist
        $exists = $false
    }

    if (-not $exists) {
        Write-LogMessage -Level Warning -Message "OneDrive folder does not exist: '$FolderPath' for user '$UserPrincipalName'"
        $create = $false
        if ($AutoCreate -or $AutoCreateOneDriveFolder -or $Force) {
            $create = $true
            Write-LogMessage -Level Info -Message "Auto-creating OneDrive folder per parameter configuration..."
        } else {
            $prompt = Read-Host "[?] OneDrive folder '$FolderPath' does not exist. Create it now in cloud? (Y/N)"
            if ($prompt -match '^(Y|y|T|t)$') {
                $create = $true
            }
        }

        if ($create) {
            Write-LogMessage -Level Info -Message "Provisioning folder hierarchy '$FolderPath' in OneDrive..."
            # Traverse/create path segments from root down
            $segments = $FolderPath.Split('/')
            $currentParentId = "root"

            foreach ($segment in $segments) {
                if ([string]::IsNullOrWhiteSpace($segment)) { continue }

                # Check if child already exists
                $childrenUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$currentParentId/children"
                $childrenResp = Invoke-GraphRequestSafe -Method GET -Uri $childrenUri
                $match = @($childrenResp.value | Where-Object { $_.name -ieq $segment -and $null -ne $_.folder })

                if ($match.Count -gt 0) {
                    $currentParentId = $match[0].id
                } else {
                    # Create folder segment
                    $createBody = @{
                        "name"                              = $segment
                        "folder"                            = @{}
                        "@microsoft.graph.conflictBehavior" = "fail"
                    }
                    $createUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$currentParentId/children"
                    $created = Invoke-GraphRequestSafe -Method POST -Uri $createUri -Body $createBody
                    $currentParentId = $created.id
                    Write-LogMessage -Level Success -Message "Created OneDrive folder segment: '$segment'"
                }
            }

            # Return the leaf item
            $finalItem = Invoke-GraphRequestSafe -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$currentParentId"
            Write-LogMessage -Level Success -Message "OneDrive folder hierarchy ready: '$FolderPath'"
            return $finalItem
        } else {
            Write-Error "Operation aborted. OneDrive cloud folder was not created."
            exit 1
        }
    }

    return $folderItem
}

# ==============================================================================
# INVENTORY SCANNING (Remote & Local)
# ==============================================================================
function Get-RemoteOneDriveItems {
    param(
        [string]$DriveId,
        [string]$FolderItemId,
        [string]$RelativePath = "",
        [bool]$IsRecursive = $true
    )

    $items = @()
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderItemId/children?`$top=200"

    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $resp = Invoke-GraphRequestSafe -Method GET -Uri $uri
        $batch = @($resp.value)

        foreach ($entry in $batch) {
            $rel = if ([string]::IsNullOrWhiteSpace($RelativePath)) { $entry.name } else { "$RelativePath/$($entry.name)" }

            if ($null -ne $entry.folder) {
                if ($IsRecursive) {
                    $subItems = Get-RemoteOneDriveItems -DriveId $DriveId -FolderItemId $entry.id -RelativePath $rel -IsRecursive $true
                    $items += $subItems
                }
            } elseif ($null -ne $entry.file) {
                $lastModUtc = [DateTime]::UtcNow
                if ($entry.lastModifiedDateTime) {
                    $lastModUtc = [DateTime]::Parse($entry.lastModifiedDateTime).ToUniversalTime()
                }

                $dlUrl = $entry.'@microsoft.graph.downloadUrl'

                $items += [PSCustomObject]@{
                    Name                 = $entry.name
                    RelativePath         = $rel.Replace('\', '/')
                    ItemId               = $entry.id
                    ParentItemId         = $FolderItemId
                    Size                 = [int64]$entry.size
                    LastModifiedDateTime = $lastModUtc
                    DownloadUrl          = $dlUrl
                }
            }
        }

        # Follow pagination link if present
        $uri = $resp.'@odata.nextLink'
    }

    return $items
}

function Get-LocalFolderInventory {
    param(
        [string]$FolderPath,
        [bool]$IsRecursive = $true
    )

    $items = @()
    $searchOpt = if ($IsRecursive) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
    $dirInfo = New-Object System.IO.DirectoryInfo($FolderPath)

    $files = $dirInfo.GetFiles("*", $searchOpt)
    $rootLen = $dirInfo.FullName.TrimEnd('\').Length + 1

    foreach ($f in $files) {
        $rel = $f.FullName.Substring($rootLen).Replace('\', '/')
        $items += [PSCustomObject]@{
            Name                 = $f.Name
            FullName             = $f.FullName
            RelativePath         = $rel
            Size                 = [int64]$f.Length
            LastModifiedDateTime = $f.LastWriteTimeUtc
        }
    }

    return $items
}

# ==============================================================================
# FILE TRANSFER ENGINES (Upload & Download with Progress Bars)
# ==============================================================================
function Upload-FileToOneDrive {
    param(
        [string]$DriveId,
        [string]$TargetFolderItemId,
        [string]$BaseOneDrivePath,
        [PSCustomObject]$LocalFile,
        [int]$FileIndex,
        [int]$TotalFiles
    )

    $fileName = $LocalFile.Name
    $filePath = $LocalFile.FullName
    $fileSize = $LocalFile.Size
    $relPath  = $LocalFile.RelativePath

    $sizeMb = [math]::Round($fileSize / 1MB, 2)
    Write-LogMessage -Level Progress -Message "[$FileIndex/$TotalFiles] Uploading '$relPath' ($sizeMb MB)..."

    # Ensure remote parent subfolder if relative path has directories
    $remoteParentId = $TargetFolderItemId
    if ($relPath.Contains('/')) {
        $subDirPath = Split-Path -Path $relPath -Parent
        $subSegments = $subDirPath.Split('/')
        $curParent = $TargetFolderItemId

        foreach ($seg in $subSegments) {
            if ([string]::IsNullOrWhiteSpace($seg)) { continue }
            $childCheckUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$curParent/children"
            $cResp = Invoke-GraphRequestSafe -Method GET -Uri $childCheckUri
            $matchedFolder = @($cResp.value | Where-Object { $_.name -ieq $seg -and $null -ne $_.folder })

            if ($matchedFolder.Count -gt 0) {
                $curParent = $matchedFolder[0].id
            } else {
                $cBody = @{
                    "name"                              = $seg
                    "folder"                            = @{}
                    "@microsoft.graph.conflictBehavior" = "replace"
                }
                $newFolder = Invoke-GraphRequestSafe -Method POST -Uri $childCheckUri -Body $cBody
                $curParent = $newFolder.id
            }
        }
        $remoteParentId = $curParent
    }

    $escapedFileName = [Uri]::EscapeDataString($fileName)

    # --------------------------------------------------------------------------
    # SMALL FILE UPLOAD (<= 4 MB)
    # --------------------------------------------------------------------------
    $fourMb = 4 * 1024 * 1024
    if ($fileSize -le $fourMb) {
        $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($remoteParentId):/$($escapedFileName):/content"
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)

        $uploadedItem = Invoke-GraphRequestSafe -Method PUT -Uri $uploadUrl -BinaryBody $fileBytes -ContentType "application/octet-stream"
        Write-LogMessage -Level Success -Message "Uploaded '$relPath' ($sizeMb MB) successfully."
        return $uploadedItem
    }

    # --------------------------------------------------------------------------
    # LARGE FILE UPLOAD SESSION (> 4 MB up to 250 GB)
    # --------------------------------------------------------------------------
    Write-LogMessage -Level Info -Message "Large file detected ($sizeMb MB). Initializing chunked upload session for '$fileName'..."
    $sessionUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($remoteParentId):/$($escapedFileName):/createUploadSession"
    $sessionBody = @{
        "item" = @{
            "@microsoft.graph.conflictBehavior" = "replace"
            "name"                              = $fileName
        }
    }
    $session = Invoke-GraphRequestSafe -Method POST -Uri $sessionUri -Body $sessionBody
    $uploadSessionUrl = $session.uploadUrl

    # Chunk size must be a multiple of 320 KiB (327,680 bytes). Use 3.2 MB chunks.
    $chunkSize = 3276800
    $buffer = New-Object byte[] $chunkSize
    $fileStream = [System.IO.File]::OpenRead($filePath)

    try {
        $totalBytesRead = 0
        while ($totalBytesRead -lt $fileSize) {
            $bytesToRead = [int][Math]::Min($chunkSize, ($fileSize - $totalBytesRead))
            $bytesRead = $fileStream.Read($buffer, 0, $bytesToRead)
            if ($bytesRead -le 0) { break }

            $startByte = $totalBytesRead
            $endByte   = $totalBytesRead + $bytesRead - 1

            $headers = @{
                "Content-Length" = "$bytesRead"
                "Content-Range"  = "bytes $startByte-$endByte/$fileSize"
            }

            $chunkBytes = New-Object byte[] $bytesRead
            [Array]::Copy($buffer, $chunkBytes, $bytesRead)

            # Upload chunk directly to session URL
            Invoke-RestMethod -Method Put -Uri $uploadSessionUrl -Headers $headers -Body $chunkBytes -ContentType "application/octet-stream" -ErrorAction Stop | Out-Null
            $totalBytesRead += $bytesRead

            # Update nested chunk progress bar
            $chunkPercent = [math]::Round(($totalBytesRead / $fileSize) * 100)
            $transferredMb = [math]::Round($totalBytesRead / 1MB, 2)
            Write-Progress -Id 2 -ParentId 1 -Activity "Chunked Upload: $fileName" `
                -Status "Uploaded $transferredMb MB of $sizeMb MB ($chunkPercent%)" `
                -PercentComplete $chunkPercent
        }
    } finally {
        $fileStream.Close()
        $fileStream.Dispose()
        Write-Progress -Id 2 -Completed
    }

    Write-LogMessage -Level Success -Message "Large file upload completed: '$relPath' ($sizeMb MB)."
}

function Download-FileFromOneDrive {
    param(
        [string]$DriveId,
        [string]$LocalBasePath,
        [PSCustomObject]$RemoteItem,
        [int]$FileIndex,
        [int]$TotalFiles
    )

    $relPath   = $RemoteItem.RelativePath
    $destPath  = Join-Path -Path $LocalBasePath -ChildPath ($relPath.Replace('/', '\'))
    $destDir   = Split-Path -Path $destPath -Parent
    $sizeMb    = [math]::Round($RemoteItem.Size / 1MB, 2)

    Write-LogMessage -Level Progress -Message "[$FileIndex/$TotalFiles] Downloading '$relPath' ($sizeMb MB)..."

    if (-not (Test-Path -Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
    }

    # Fetch fresh download URL if not present
    $downloadUrl = $RemoteItem.DownloadUrl
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
        $metaUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($RemoteItem.ItemId)"
        $meta = Invoke-GraphRequestSafe -Method GET -Uri $metaUri
        $downloadUrl = $meta.'@microsoft.graph.downloadUrl'
    }

    if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
        # Fallback to Graph item content endpoint
        $contentUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($RemoteItem.ItemId)/content"
        Invoke-GraphRequestSafe -Method GET -Uri $contentUri -OutFile $destPath
    } else {
        # High speed CDN download
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destPath -ErrorAction Stop | Out-Null
    }

    # Match local timestamp with cloud LastModifiedDateTime
    try {
        [System.IO.File]::SetLastWriteTimeUtc($destPath, $RemoteItem.LastModifiedDateTime)
    } catch {
        # Ignore timestamp modification errors if restricted
    }

    Write-LogMessage -Level Success -Message "Downloaded '$relPath' ($sizeMb MB) to '$destPath'."
}

# ==============================================================================
# MAIN SYNCHRONIZATION EXECUTION ROUTINE
# ==============================================================================
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # 1. Establish Graph Connectivity
    Connect-GraphEnvironment

    # 2. Resolve User OneDrive Drive
    $userDrive = Get-UserOneDriveDrive -UserUPN $UserPrincipalName
    $driveId   = $userDrive.id

    # 3. Ensure Local Folder Hierarchy
    Ensure-LocalFolderHierarchy -Path $LocalFolderPath

    # 4. Ensure OneDrive Folder Hierarchy
    $cloudFolderItem = Ensure-OneDriveFolderHierarchy -DriveId $driveId -FolderPath $normalizedOneDrivePath
    $cloudFolderId   = $cloudFolderItem.id

    # 5. Inventory Scan
    Write-LogMessage -Level Info -Message "Scanning inventory on local disk..."
    $localItems = Get-LocalFolderInventory -FolderPath $LocalFolderPath -IsRecursive $Recurse
    Write-LogMessage -Level Info -Message "Discovered $($localItems.Count) local file(s)."

    Write-LogMessage -Level Info -Message "Scanning inventory in OneDrive cloud..."
    $remoteItems = Get-RemoteOneDriveItems -DriveId $driveId -FolderItemId $cloudFolderId -RelativePath "" -IsRecursive $Recurse
    Write-LogMessage -Level Info -Message "Discovered $($remoteItems.Count) remote file(s)."

    # Index items by normalized RelativePath for fast lookup
    $localMap = @{}
    foreach ($item in $localItems) { $localMap[$item.RelativePath.ToLowerInvariant()] = $item }

    $remoteMap = @{}
    foreach ($item in $remoteItems) { $remoteMap[$item.RelativePath.ToLowerInvariant()] = $item }

    # 6. Synchronization Plan Planning
    $uploadsPending   = @()
    $downloadsPending = @()
    $unchangedCount   = 0

    $timeToleranceSeconds = 2.0

    # Evaluate Local Files
    foreach ($relKey in $localMap.Keys) {
        $loc = $localMap[$relKey]
        if (-not $remoteMap.ContainsKey($relKey)) {
            # File exists locally but not in OneDrive
            if ($Direction -eq 'TwoWay' -or $Direction -eq 'Upload') {
                $uploadsPending += $loc
            }
        } else {
            # File exists on both sides
            $rem = $remoteMap[$relKey]
            $timeDiff = ($loc.LastModifiedDateTime - $rem.LastModifiedDateTime).TotalSeconds

            if ([math]::Abs($timeDiff) -le $timeToleranceSeconds -and $loc.Size -eq $rem.Size) {
                $unchangedCount++
            } elseif ($timeDiff -gt $timeToleranceSeconds) {
                # Local file is newer
                if ($Direction -eq 'TwoWay' -or $Direction -eq 'Upload') {
                    $uploadsPending += $loc
                }
            } elseif ($timeDiff -lt (-1 * $timeToleranceSeconds)) {
                # Remote file is newer
                if ($Direction -eq 'TwoWay' -or $Direction -eq 'Download') {
                    $downloadsPending += $rem
                }
            } else {
                # Same timestamp but different size -> default to local upload
                if ($Direction -eq 'TwoWay' -or $Direction -eq 'Upload') {
                    $uploadsPending += $loc
                } elseif ($Direction -eq 'Download') {
                    $downloadsPending += $rem
                }
            }
        }
    }

    # Evaluate Remote Files Missing Locally
    foreach ($relKey in $remoteMap.Keys) {
        if (-not $localMap.ContainsKey($relKey)) {
            if ($Direction -eq 'TwoWay' -or $Direction -eq 'Download') {
                $downloadsPending += $remoteMap[$relKey]
            }
        }
    }

    $totalActions = $uploadsPending.Count + $downloadsPending.Count
    Write-LogMessage -Level Info -Message "Sync plan: $($uploadsPending.Count) to upload, $($downloadsPending.Count) to download, $unchangedCount unchanged."

    # 7. Execute Transfers with Granular Write-Progress
    $currentActionIndex = 0
    $filesUploadedCount = 0
    $filesDownloadedCount = 0
    $filesFailedCount   = 0
    $totalBytesTransferred = 0

    # Process Uploads
    foreach ($upFile in $uploadsPending) {
        $currentActionIndex++
        $overallPercent = [math]::Round(($currentActionIndex / [math]::Max($totalActions, 1)) * 100)

        Write-Progress -Id 1 -Activity "Synchronizing with OneDrive ($UserPrincipalName)" `
            -Status "Upload [$currentActionIndex/$totalActions] ($overallPercent%): $($upFile.Name)" `
            -CurrentOperation "Target: $normalizedOneDrivePath/$($upFile.RelativePath)" `
            -PercentComplete $overallPercent

        try {
            Upload-FileToOneDrive -DriveId $driveId -TargetFolderItemId $cloudFolderId `
                -BaseOneDrivePath $normalizedOneDrivePath -LocalFile $upFile `
                -FileIndex $currentActionIndex -TotalFiles $totalActions
            $filesUploadedCount++
            $totalBytesTransferred += $upFile.Size
        } catch {
            $filesFailedCount++
            Write-LogMessage -Level Error -Message "Failed to upload '$($upFile.RelativePath)': $_"
        }
    }

    # Process Downloads
    foreach ($downFile in $downloadsPending) {
        $currentActionIndex++
        $overallPercent = [math]::Round(($currentActionIndex / [math]::Max($totalActions, 1)) * 100)

        Write-Progress -Id 1 -Activity "Synchronizing with OneDrive ($UserPrincipalName)" `
            -Status "Download [$currentActionIndex/$totalActions] ($overallPercent%): $($downFile.Name)" `
            -CurrentOperation "Source: $normalizedOneDrivePath/$($downFile.RelativePath)" `
            -PercentComplete $overallPercent

        try {
            Download-FileFromOneDrive -DriveId $driveId -LocalBasePath $LocalFolderPath `
                -RemoteItem $downFile -FileIndex $currentActionIndex -TotalFiles $totalActions
            $filesDownloadedCount++
            $totalBytesTransferred += $downFile.Size
        } catch {
            $filesFailedCount++
            Write-LogMessage -Level Error -Message "Failed to download '$($downFile.RelativePath)': $_"
        }
    }

    Write-Progress -Id 1 -Completed
    $stopwatch.Stop()

    # 8. Execution Summary Report
    $transferredMb = [math]::Round($totalBytesTransferred / 1MB, 2)
    $elapsedFormatted = $stopwatch.Elapsed.ToString("hh\:mm\:ss")
    $syncStatus = if ($filesFailedCount -eq 0) { "SUCCESS" } else { "COMPLETED WITH WARNINGS ($filesFailedCount failed)" }

    $summaryReport = @"
================================================================================
  SYNCHRONIZATION COMPLETED SUMMARY
================================================================================
Target User           : $UserPrincipalName
OneDrive Folder       : $(if ($normalizedOneDrivePath) { $normalizedOneDrivePath } else { '/ (OneDrive Root)' })
Local Folder          : $LocalFolderPath
Sync Direction        : $Direction
Total Files Analyzed  : $($localItems.Count + $remoteItems.Count)
Files Uploaded        : $filesUploadedCount
Files Downloaded      : $filesDownloadedCount
Files Unchanged       : $unchangedCount
Failed Operations     : $filesFailedCount
Total Data Transferred: $transferredMb MB
Elapsed Duration      : $elapsedFormatted
Log File Path         : $(if ($global:ScriptLogPath) { $global:ScriptLogPath } else { 'None' })
Final Status          : [$syncStatus]
================================================================================
"@
    Write-Host $summaryReport -ForegroundColor Green
    if ($global:ScriptLogPath) {
        Add-Content -Path $global:ScriptLogPath -Value $summaryReport -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    # 9. Send Notifications
    $notifyTitle = "OneDrive Sync: $syncStatus"
    $notifyBody  = "Processed $totalActions items ($filesUploadedCount uploaded, $filesDownloadedCount downloaded) for $UserPrincipalName in $elapsedFormatted."
    $notifyType  = if ($filesFailedCount -gt 0) { 'Warning' } else { 'Info' }

    Send-SyncNotification -Title $notifyTitle -Message $notifyBody -Type $notifyType

} catch {
    Write-Progress -Id 1 -Completed
    $stopwatch.Stop()
    $errMessage = "Fatal error during synchronization execution: $_"
    Write-LogMessage -Level Error -Message $errMessage

    Send-SyncNotification -Title "OneDrive Sync Failed" -Message $errMessage -Type 'Error'
    exit 1
}
