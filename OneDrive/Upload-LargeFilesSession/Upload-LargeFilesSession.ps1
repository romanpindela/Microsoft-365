<#
.SYNOPSIS
    High-performance, resilient upload session script for uploading large files, directories,
    and wildcard sets to Microsoft 365 OneDrive using Microsoft Graph API / SDK.

.DESCRIPTION
    Upload-LargeFilesSession is an enterprise-grade PowerShell tool designed to upload individual files,
    entire directory trees, or files matching wildcard masks (e.g. *.pdf, p?indela.pdf, p*.docx, or folder names)
    directly to a targeted Microsoft 365 user's OneDrive for Business.

    Key Features:
    - Microsoft Graph Upload Session: Uses the official Graph createUploadSession API to chunk files
      (required for files > 4MB, supporting large files up to 250 GB) in compliant 320 KiB multiples.
    - Automatic Folder Hierarchy Creation: If the destination folder path in OneDrive does not exist,
      the script recursively verifies and provisions every intermediate parent and child folder.
    - Flexible Input Matching: Accepts single files, directories, arrays of paths, and search masks/wildcards
      (e.g., *.pdf, p?indela.pdf, p*.docx, or folder names), with optional recursion (-Recurse) and
      directory structure preservation (-PreserveStructure).
    - Detailed Real-Time Progress: Displays startup parameter summaries and dual Write-Progress indicators
      (overall payload % and current file chunk %, transferred bytes, transfer speed in MB/s, and ETA).
    - Enterprise Logging: Timestamped color-coded console telemetry and persistent UTF-8 log file recording.
    - Dual Authentication Modes:
        1. Non-interactive (headless / daemon): Authenticates via Entra ID App Registration using
           Client Credentials (Client Secret or Certificate) without interactive prompts.
        2. Interactive: Authenticates via modern browser / MFA prompt using Microsoft.Graph PowerShell SDK.
    - Conflict Resolution: Configurable file conflict behavior (-ConflictBehavior: replace, rename, fail).
    - Defensive Engineering: Input sanitization, path traversal prevention, automatic chunk retries with
      exponential backoff for HTTP 429 (throttling) / 5xx server errors, and memory-efficient streaming.

.PARAMETER UserPrincipalName
    The primary SMTP address / UPN of the target Microsoft 365 user whose OneDrive will receive the files
    (e.g., user@domain.com).
    Aliases: -User, -UPN, -UserIdentity, -Identity, -u.

.PARAMETER TargetFolder
    The target folder path inside the user's OneDrive (e.g., 'Backups/Database' or 'Documents/2026').
    If the folder path does not exist, the script automatically creates all missing parent and child folders.
    Aliases: -Folder, -DestinationFolder, -RemoteFolder, -Target, -f.

.PARAMETER Path
    One or more file paths, directory paths, or wildcard patterns to upload (e.g., 'C:\Data\archive.zip',
    'C:\Reports\*.pdf', 'p?indela.pdf', 'p*.docx', or 'my_folder').
    Aliases: -FilePath, -Source, -Files, -InputPath, -p.

.PARAMETER Filter
    Optional filter mask applied when scanning directories (e.g., '*.pdf', '*.docx').
    Aliases: -Mask.

.PARAMETER Recurse
    Switch parameter. Enables recursive scanning of directories when folders or patterns are specified.
    Aliases: -r.

.PARAMETER PreserveStructure
    Switch parameter. Preserves local subfolder hierarchy inside the OneDrive destination folder when
    uploading directory trees.
    Aliases: -KeepHierarchy, -MaintainFolders.

.PARAMETER ConflictBehavior
    Specifies how Microsoft Graph handles filename collisions in OneDrive.
    Allowed values: 'replace' (default), 'rename', 'fail'.

.PARAMETER ChunkSizeMB
    Upload chunk size in megabytes (1-100 MB). Default is 10 MB.
    Automatically rounded to the nearest compliant multiple of 320 KiB (327,680 bytes) as required by
    Microsoft Graph API.
    Aliases: -ChunkMB.

.PARAMETER TenantId
    Microsoft Entra ID Directory (tenant) ID (GUID) used for non-interactive App-Only authentication.
    Aliases: -Tenant.

.PARAMETER ClientId
    Microsoft Entra ID Application (client) ID (GUID) used for non-interactive App-Only authentication.
    Aliases: -AppId, -ApplicationId.

.PARAMETER ClientSecret
    Plaintext client secret value associated with the Entra ID App Registration.
    Aliases: -Secret.

.PARAMETER ClientSecretSecure
    SecureString representation of the client secret.

.PARAMETER CertificateThumbprint
    Certificate thumbprint from the local Windows Certificate Store for certificate-based App-Only login.
    Aliases: -Thumbprint.

.PARAMETER CertificatePath
    Path to a .pfx certificate file for certificate-based App-Only login.
    Aliases: -CertPath.

.PARAMETER Interactive
    Switch parameter. Forces interactive browser authentication via Microsoft.Graph SDK, bypassing any
    stored or environmental app credentials.
    Aliases: -InteractiveLogin, -i.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used to pre-fill the interactive sign-in prompt.
    Aliases: -AdminUPN.

.PARAMETER LogPath
    Custom file path for the output log file. If omitted, logs default to
    '.\logs\Upload-LargeFilesSession_YYYYMMDD_HHmmss.log'.
    Aliases: -LogFile.

.PARAMETER LogDirectory
    Custom directory where log files should be created.
    Aliases: -LogDir.

.PARAMETER LogLevel
    Minimum logging level. Allowed values: 'Debug', 'Info' (default), 'Warning', 'Error'.

.PARAMETER NoLog
    Switch parameter. Disables writing log entries to disk (console output remains active).

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts for unattended execution.
    Aliases: -f.

.PARAMETER PassThru
    Switch parameter. Outputs custom PowerShell objects containing upload results to the pipeline.

.PARAMETER Help
    Displays detailed script help, usage examples, Entra ID setup instructions, and author metadata.
    Aliases: -h.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1
    Running without parameters displays the built-in help banner, usage examples, author info, and version.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 -UserPrincipalName "john.doe@company.com" -TargetFolder "Backups" -Path "C:\Data\database_backup.vhdx"
    Uploads a large 40 GB VHDX virtual disk to John Doe's OneDrive 'Backups' folder using interactive login.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -Folder "Reports" -Path "C:\Monthly\*.pdf"
    Uploads all PDF files matching the wildcard pattern '*.pdf' into the 'Reports' folder.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -Folder "Contracts" -Path "p?indela.pdf"
    Uploads files matching single-character wildcard 'p?indela.pdf' (e.g. ppindela.pdf, p1indela.pdf) to OneDrive.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -Folder "ProjectDocs" -Path "p*.docx" -ConflictBehavior rename
    Uploads all docx files starting with 'p' and automatically renames any files that conflict with existing items.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -Folder "Archive/2026" -Path "C:\ProjectData" -Recurse -PreserveStructure
    Recursively uploads the 'ProjectData' directory tree, automatically replicating local folder structures in OneDrive.

.EXAMPLE
    .\Upload-LargeFilesSession.ps1 `
        -UserPrincipalName "backup-service@company.com" `
        -TargetFolder "SQL_Backups" `
        -Path "E:\SQL\*.bak" `
        -TenantId "11111111-2222-3333-4444-555555555555" `
        -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -ClientSecret "SecretValueHere~1234" `
        -ChunkSizeMB 20
    Executes unattended, headless non-interactive upload using Entra ID Client Credentials with 20 MB chunks.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Target User & Destination
    [Parameter(ParameterSetName = 'Default', Position = 0)]
    [Alias('User', 'UPN', 'UserIdentity', 'Identity', 'u')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'Default', Position = 1)]
    [Alias('Folder', 'DestinationFolder', 'RemoteFolder', 'Target')]
    [string]$TargetFolder,

    # Source Files / Folders / Masks
    [Parameter(ParameterSetName = 'Default', Position = 2)]
    [Alias('FilePath', 'Source', 'Files', 'InputPath', 'p')]
    [string[]]$Path,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('Mask')]
    [string]$Filter,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('r')]
    [switch]$Recurse,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('KeepHierarchy', 'MaintainFolders')]
    [switch]$PreserveStructure,

    # Upload Session Options
    [Parameter(ParameterSetName = 'Default')]
    [ValidateSet('replace', 'rename', 'fail')]
    [string]$ConflictBehavior = 'replace',

    [Parameter(ParameterSetName = 'Default')]
    [ValidateRange(1, 100)]
    [Alias('ChunkMB')]
    [int]$ChunkSizeMB = 10,

    # Entra ID / Microsoft Graph Non-Interactive Authentication
    [Parameter(ParameterSetName = 'Default')]
    [Alias('Tenant')]
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
    [string]$CertificatePath,

    # Interactive Authentication
    [Parameter(ParameterSetName = 'Default')]
    [Alias('InteractiveLogin', 'i')]
    [switch]$Interactive,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

    # Logging Options
    [Parameter(ParameterSetName = 'Default')]
    [Alias('LogFile')]
    [string]$LogPath,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('LogDir')]
    [string]$LogDirectory,

    [Parameter(ParameterSetName = 'Default')]
    [ValidateSet('Debug', 'Info', 'Warning', 'Error')]
    [string]$LogLevel = 'Info',

    [Parameter(ParameterSetName = 'Default')]
    [switch]$NoLog,

    # Execution Flow & Output
    [Parameter(ParameterSetName = 'Default')]
    [Alias('f')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$PassThru,

    # Built-in Help
    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata declaration
$ScriptMetadata = @{
    Title       = "Upload-LargeFilesSession.ps1"
    Version     = "1.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    LastUpdated = "2026-10-15"
}

# Global State Variables
$script:ActiveLogFile = $null
$script:GraphToken    = $null
$script:AuthType      = "Unknown"
$script:HttpClient    = $null

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
        "    High-performance PowerShell automation tool designed to upload files of any size",
        "    (including multi-gigabyte files up to 250 GB) to a target user's Microsoft 365 OneDrive.",
        "    Implements official Microsoft Graph Upload Session chunking (createUploadSession),",
        "    automatic remote folder creation, recursive directory uploads, wildcard pattern matching",
        "    (e.g., *.pdf, p?indela.pdf, p*.docx, or folder names), detailed dual-level progress bars,",
        "    and enterprise-grade logging.",
        "",
        "KEY CAPABILITIES:",
        "    - Upload Sessions: Chunks large files in compliant 320 KiB multiples with automatic retry.",
        "    - Auto-Provision Folders: Missing folders in OneDrive target path are created automatically.",
        "    - Mask & Wildcard Support: Matches files via *, ?, directory names, or explicit file paths.",
        "    - Recursive Uploads: -Recurse and -PreserveStructure replicate local trees in OneDrive.",
        "    - Detailed Telemetry: Displays parameter tables, real-time chunk speed, ETA, and dual progress.",
        "    - Non-Interactive Daemon Ready: Supports headless execution via Entra ID App Registration.",
        "",
        "HOW TO CONFIGURE ENTRA ID FOR NON-INTERACTIVE / HEADLESS EXECUTION:",
        "    To run this script automatically via Task Scheduler, Azure DevOps, CI/CD, or service accounts",
        "    without any interactive browser prompts, follow these steps:",
        "",
        "    1. Register Application in Microsoft Entra ID (Azure AD):",
        "       - Open Microsoft Entra admin center (https://entra.microsoft.com) or Azure Portal.",
        "       - Go to: Identity > Applications > App registrations > Click 'New registration'.",
        "       - Name: e.g., 'OneDrive-LargeFiles-Uploader'.",
        "       - Supported account types: 'Accounts in this organizational directory only (Single tenant)'.",
        "       - Click 'Register'.",
        "",
        "    2. Collect Application Identifiers:",
        "       - Note the 'Application (client) ID' (GUID).",
        "       - Note the 'Directory (tenant) ID' (GUID).",
        "",
        "    3. Create Client Secret or Upload Certificate:",
        "       - Go to: Manage > Certificates & secrets > Click 'New client secret'.",
        "       - Add description, choose expiration (e.g., 180 or 730 days), and click 'Add'.",
        "       - Copy the 'Value' immediately (it will not be shown again!).",
        "       - (Alternative) Upload a public certificate (.cer) and use -CertificateThumbprint.",
        "",
        "    4. Configure API Permissions (CRITICAL: MUST BE APPLICATION PERMISSIONS):",
        "       - Go to: Manage > API permissions > Click 'Add a permission'.",
        "       - Select 'Microsoft Graph' > Choose 'Application permissions' (NOT Delegated!).",
        "       - Search and select:",
        "           * Files.ReadWrite.All  (Grants full read/write access to all OneDrive/SharePoint files)",
        "           * User.Read.All        (Allows resolving user identities and OneDrive drive endpoints)",
        "       - Click 'Add permissions'.",
        "",
        "    5. Grant Admin Consent:",
        "       - Click 'Grant admin consent for [Your Organization]'.",
        "       - Confirm with 'Yes'. Verify green checkmarks 'Granted for [Tenant]' appear for both.",
        "",
        '    6. Non-Interactive Execution Command:',
        '       .\Upload-LargeFilesSession.ps1 `',
        '           -UserPrincipalName "user@company.com" `',
        '           -TargetFolder "Backups/Daily" `',
        '           -Path "D:\Backups\*.zip" `',
        '           -TenantId "YOUR_TENANT_ID" `',
        '           -ClientId "YOUR_CLIENT_ID" `',
        '           -ClientSecret "YOUR_SECRET_VALUE" `',
        '           -ChunkSizeMB 20',
        '',
        '       Tip: You can also define environment variables:',
        '       $env:M365_TENANT_ID = "YOUR_TENANT_ID"',
        '       $env:M365_CLIENT_ID = "YOUR_CLIENT_ID"',
        '       $env:M365_CLIENT_SECRET = "YOUR_SECRET_VALUE"',
        '',
        'USAGE EXAMPLES:',
        '    # 1. Upload single large file interactively (MFA login prompt):',
        '    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -TargetFolder "Backups" -Path "C:\Large\video.mp4"',
        '',
        '    # 2. Upload multiple files using wildcard mask (*.pdf):',
        '    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -TargetFolder "Reports/2026" -Path "C:\Data\*.pdf"',
        '',
        '    # 3. Upload files matching single-character wildcard (p?indela.pdf):',
        '    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -TargetFolder "Contracts" -Path "p?indela.pdf"',
        '',
        '    # 4. Upload files matching prefix wildcard (p*.docx) with rename conflict rule:',
        '    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -TargetFolder "Docs" -Path "p*.docx" -ConflictBehavior rename',
        '',
        '    # 5. Upload an entire directory with all subfolders preserved:',
        '    .\Upload-LargeFilesSession.ps1 -User "john.doe@company.com" -TargetFolder "Archive" -Path "C:\Projects\Alpha" -Recurse -PreserveStructure',
        '',
        '    # 6. Non-interactive upload with custom 25 MB chunks and explicit log path:',
        '    .\Upload-LargeFilesSession.ps1 `',
        '        -User "user@company.com" `',
        '        -TargetFolder "SQL_Dumps" `',
        '        -Path "E:\Dumps\*.bak" `',
        '        -TenantId "11111111-2222-3333-4444-555555555555" `',
        '        -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `',
        '        -ClientSecret "SecretValue..." `',
        '        -ChunkSizeMB 25 `',
        '        -LogPath "C:\Logs\OneDrive_Upload.log"',
        "",
        "PARAMETERS:",
        "    -UserPrincipalName, -User, -UPN   Target Microsoft 365 user email / UPN (OneDrive owner).",
        "    -TargetFolder, -Folder, -Target   Target folder inside OneDrive (created if missing).",
        "    -Path, -FilePath, -Source         One or more files, folders, or wildcard masks (*.pdf, p*.docx, etc.).",
        "    -Filter, -Mask                    Secondary file filter applied when scanning folders.",
        "    -Recurse, -r                      Scan folders recursively.",
        "    -PreserveStructure                Replicate local subfolder paths inside destination OneDrive folder.",
        "    -ConflictBehavior                 Action on filename collision: 'replace' (default), 'rename', 'fail'.",
        "    -ChunkSizeMB                      Upload chunk size in MB (1-100, default: 10 MB, rounded to 320 KiB).",
        "    -TenantId                         Entra ID Directory (tenant) ID for App-Only non-interactive login.",
        "    -ClientId, -AppId                 Entra ID Application (client) ID for App-Only non-interactive login.",
        "    -ClientSecret, -Secret            Client secret plaintext value for App-Only authentication.",
        "    -ClientSecretSecure               Client secret as SecureString.",
        "    -CertificateThumbprint            Certificate thumbprint from local store for App-Only login.",
        "    -CertificatePath                  Path to .pfx certificate file for App-Only login.",
        "    -Interactive, -i                  Forces interactive browser login via Microsoft.Graph SDK.",
        "    -AdminUserPrincipalName           Admin UPN to pre-populate interactive login prompt.",
        "    -LogPath                          Custom log file path.",
        "    -LogDirectory                     Custom directory for log files.",
        "    -LogLevel                         Minimum log level ('Debug', 'Info', 'Warning', 'Error').",
        "    -NoLog                            Disable file logging to disk.",
        "    -Force, -f                        Bypass interactive confirmations.",
        "    -PassThru                         Output completed file objects to pipeline.",
        "    -Help, -h                         Display this help manual.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Upload-LargeFilesSession.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Auto-show help when executed without parameters or when -Help / -h is specified
$hasExecutionParams = (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) -or `
                      (-not [string]::IsNullOrWhiteSpace($TargetFolder)) -or `
                      ($null -ne $Path -and $Path.Count -gt 0)

if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or (-not $hasExecutionParams)) {
    Show-ScriptHelp
    exit 0
}

# ==============================================================================
# Helper Functions: Formatting, Logging, Defensive Sanitization
# ==============================================================================

function Format-ByteSize {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0} B" -f $Bytes)
}

function Format-TimeSpan {
    param([TimeSpan]$Span)
    if ($Span.TotalHours -ge 1) {
        return ("{0:D2}:{1:D2}:{2:D2}" -f [int]$Span.TotalHours, $Span.Minutes, $Span.Seconds)
    }
    return ("{0:D2}:{1:D2}" -f [int]$Span.TotalMinutes, $Span.Seconds)
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $levelHierarchy = @{ 'DEBUG' = 1; 'INFO' = 2; 'WARN' = 3; 'ERROR' = 4; 'SUCCESS' = 2 }
    $configThreshold = switch ($script:ConfiguredLogLevel.ToUpperInvariant()) {
        'DEBUG'   { 1 }
        'INFO'    { 2 }
        'WARNING' { 3 }
        'ERROR'   { 4 }
        default   { 2 }
    }

    $msgLevelVal = if ($levelHierarchy.ContainsKey($Level.ToUpperInvariant())) {
        $levelHierarchy[$Level.ToUpperInvariant()]
    } else { 2 }

    if ($msgLevelVal -lt $configThreshold) { return }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"

    $consoleColor = switch ($Level.ToUpperInvariant()) {
        'DEBUG'   { [ConsoleColor]::DarkGray }
        'INFO'    { [ConsoleColor]::Cyan }
        'SUCCESS' { [ConsoleColor]::Green }
        'WARN'    { [ConsoleColor]::Yellow }
        'ERROR'   { [ConsoleColor]::Red }
        default   { [ConsoleColor]::White }
    }

    $prefix = switch ($Level.ToUpperInvariant()) {
        'DEBUG'   { "  [DEBUG] " }
        'INFO'    { "  [*] " }
        'SUCCESS' { "  [+] " }
        'WARN'    { "  [!] " }
        'ERROR'   { "  [-] " }
        default   { "      " }
    }

    try { [Console]::CursorLeft = 0 } catch {}
    Write-Host "$prefix$Message" -ForegroundColor $consoleColor

    if (-not [string]::IsNullOrWhiteSpace($script:ActiveLogFile)) {
        try {
            [System.IO.File]::AppendAllText($script:ActiveLogFile, "$logLine`r`n", [System.Text.Encoding]::UTF8)
        } catch {}
    }
}

function Initialize-LoggingSystem {
    param(
        [string]$CustomLogPath,
        [string]$CustomLogDir,
        [switch]$DisableLog
    )

    if ($DisableLog) {
        $script:ActiveLogFile = $null
        return
    }

    try {
        $resolvedLogFile = $null
        if (-not [string]::IsNullOrWhiteSpace($CustomLogPath)) {
            $parentDir = [System.IO.Path]::GetDirectoryName($CustomLogPath)
            if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -Path $parentDir)) {
                New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
            }
            $resolvedLogFile = $CustomLogPath
        } else {
            $baseDir = if (-not [string]::IsNullOrWhiteSpace($CustomLogDir)) {
                $CustomLogDir
            } else {
                $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
                Join-Path -Path $scriptDir -ChildPath "logs"
            }

            if (-not (Test-Path -Path $baseDir)) {
                New-Item -ItemType Directory -Path $baseDir -Force | Out-Null
            }

            $dateStamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
            $resolvedLogFile = Join-Path -Path $baseDir -ChildPath "Upload-LargeFilesSession_$dateStamp.log"
        }

        $script:ActiveLogFile = $resolvedLogFile

        # Write Log Header
        $header = @(
            "================================================================================",
            "LOG SESSION INITIALIZED: $($ScriptMetadata.Title) v$($ScriptMetadata.Version)",
            "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTCzz')",
            "Author: $($ScriptMetadata.Author) ($($ScriptMetadata.Email))",
            "PowerShell Version: $($PSVersionTable.PSVersion)",
            "OS: $([System.Environment]::OSVersion.VersionString)",
            "User: $([System.Environment]::UserName)",
            "================================================================================"
        )
        [System.IO.File]::WriteAllText($script:ActiveLogFile, (($header -join "`r`n") + "`r`n`r`n"), [System.Text.Encoding]::UTF8)
    } catch {
        Write-Host "  [!] Warning: Could not initialize log file: $($_.Exception.Message)" -ForegroundColor Yellow
        $script:ActiveLogFile = $null
    }
}

function Assert-SafeInputText {
    param(
        [string]$InputText,
        [string]$ParamName,
        [int]$MaxLength = 512
    )
    if ([string]::IsNullOrWhiteSpace($InputText)) { return }
    if ($InputText.Length -gt $MaxLength) {
        throw "Parameter -$ParamName exceeds maximum allowed length of $MaxLength characters."
    }
    if ($InputText -match '[\0`$<>|;&]') {
        throw "Parameter -$ParamName contains forbidden or dangerous characters."
    }
}

# ==============================================================================
# Input Validation & Sanitization
# ==============================================================================

$script:ConfiguredLogLevel = $LogLevel
Initialize-LoggingSystem -CustomLogPath $LogPath -CustomLogDir $LogDirectory -DisableLog:$NoLog

Write-Log "Initializing Upload-LargeFilesSession v$($ScriptMetadata.Version)..." -Level INFO

# Validate UserPrincipalName
if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    Write-Log "Missing required parameter: -UserPrincipalName. Specify a valid Microsoft 365 user." -Level ERROR
    exit 1
}
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
$cleanUPN = $UserPrincipalName.Trim().ToLowerInvariant()
if ($cleanUPN -notmatch $emailPattern) {
    Write-Log "Invalid UserPrincipalName format: '$UserPrincipalName'. Must be a valid email or UPN." -Level ERROR
    exit 1
}
$UserPrincipalName = $cleanUPN

# Validate TargetFolder
if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
    Write-Log "Missing required parameter: -TargetFolder. Specify target destination folder in OneDrive." -Level ERROR
    exit 1
}
Assert-SafeInputText -InputText $TargetFolder -ParamName "TargetFolder"
# Normalize target folder: replace backslashes, strip leading and trailing slashes
$normalizedTargetFolder = $TargetFolder.Replace('\', '/').Trim('/')
if ($normalizedTargetFolder.Contains('//') -or $normalizedTargetFolder.Contains('..')) {
    Write-Log "TargetFolder path contains illegal traversal elements ('..' or '//')." -Level ERROR
    exit 1
}

# Validate Source Paths
if ($null -eq $Path -or $Path.Count -eq 0) {
    Write-Log "Missing required parameter: -Path. Specify at least one file, directory, or wildcard mask." -Level ERROR
    exit 1
}

# Compliant Chunk Size Calculation (Multiple of 320 KiB = 327,680 bytes)
$CHUNK_UNIT = 327680 # 320 KiB
$requestedChunkBytes = [long]$ChunkSizeMB * 1024 * 1024
$compliantChunkSize = [long]([Math]::Max(1, [Math]::Floor($requestedChunkBytes / $CHUNK_UNIT)) * $CHUNK_UNIT)

# ==============================================================================
# File Discovery & Mask Matching Engine
# ==============================================================================

function Resolve-UploadCandidates {
    param(
        [string[]]$SourcePaths,
        [string]$FileFilter,
        [switch]$IsRecursive,
        [switch]$KeepHierarchy
    )

    $candidateList = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processedFiles = @{}

    foreach ($src in $SourcePaths) {
        if ([string]::IsNullOrWhiteSpace($src)) { continue }
        $cleanSrc = $src.Trim()

        $hasWildcard = $cleanSrc.Contains('*') -or $cleanSrc.Contains('?')

        if ($hasWildcard) {
            # Wildcard pattern (e.g. *.pdf, p?indela.pdf, p*.docx, or C:\Folder\*.zip)
            $parentDir = [System.IO.Path]::GetDirectoryName($cleanSrc)
            $pattern = [System.IO.Path]::GetFileName($cleanSrc)
            
            if ([string]::IsNullOrWhiteSpace($parentDir)) {
                $parentDir = (Get-Location).Path
            } elseif (-not [System.IO.Path]::IsPathRooted($parentDir)) {
                $parentDir = Join-Path -Path (Get-Location).Path -ChildPath $parentDir
            }

            if (-not (Test-Path -Path $parentDir)) {
                Write-Log "Directory does not exist for pattern: '$cleanSrc'" -Level WARN
                continue
            }

            $gciParams = @{
                Path        = $parentDir
                Filter      = $pattern
                File        = $true
                ErrorAction = 'SilentlyContinue'
            }
            if ($IsRecursive) { $gciParams['Recurse'] = $true }

            $foundFiles = Get-ChildItem @gciParams
            foreach ($fileItem in $foundFiles) {
                if ($processedFiles.ContainsKey($fileItem.FullName.ToLowerInvariant())) { continue }
                $processedFiles[$fileItem.FullName.ToLowerInvariant()] = $true

                $relativeSub = ""
                if ($KeepHierarchy -and $fileItem.FullName.Length -gt $parentDir.Length) {
                    $relativeSub = [System.IO.Path]::GetDirectoryName($fileItem.FullName.Substring($parentDir.Length).TrimStart('\', '/'))
                }

                $candidateList.Add([PSCustomObject]@{
                    FullName      = $fileItem.FullName
                    Name          = $fileItem.Name
                    Length        = $fileItem.Length
                    RelativeSub   = $relativeSub
                    DirectoryName = $fileItem.DirectoryName
                })
            }
        } else {
            # Literal Path (File or Directory)
            $resolvedPath = if ([System.IO.Path]::IsPathRooted($cleanSrc)) {
                $cleanSrc
            } else {
                Join-Path -Path (Get-Location).Path -ChildPath $cleanSrc
            }

            if (-not (Test-Path -Path $resolvedPath)) {
                # Test if cleanSrc could be a relative wildcard in current directory
                $fallbackItems = Get-ChildItem -Path (Get-Location).Path -Filter $cleanSrc -File -ErrorAction SilentlyContinue
                if ($fallbackItems -and $fallbackItems.Count -gt 0) {
                    foreach ($fileItem in $fallbackItems) {
                        if ($processedFiles.ContainsKey($fileItem.FullName.ToLowerInvariant())) { continue }
                        $processedFiles[$fileItem.FullName.ToLowerInvariant()] = $true
                        $candidateList.Add([PSCustomObject]@{
                            FullName      = $fileItem.FullName
                            Name          = $fileItem.Name
                            Length        = $fileItem.Length
                            RelativeSub   = ""
                            DirectoryName = $fileItem.DirectoryName
                        })
                    }
                    continue
                }

                Write-Log "Source path not found: '$cleanSrc'" -Level WARN
                continue
            }

            $fsItem = Get-Item -Path $resolvedPath -ErrorAction Stop
            if ($fsItem.PSIsContainer) {
                # It's a directory
                $gciParams = @{
                    Path        = $fsItem.FullName
                    File        = $true
                    ErrorAction = 'SilentlyContinue'
                }
                if (-not [string]::IsNullOrWhiteSpace($FileFilter)) {
                    $gciParams['Filter'] = $FileFilter
                }
                if ($IsRecursive) {
                    $gciParams['Recurse'] = $true
                }

                $dirFiles = Get-ChildItem @gciParams
                foreach ($fileItem in $dirFiles) {
                    if ($processedFiles.ContainsKey($fileItem.FullName.ToLowerInvariant())) { continue }
                    $processedFiles[$fileItem.FullName.ToLowerInvariant()] = $true

                    $relativeSub = ""
                    if ($KeepHierarchy -and $fileItem.FullName.Length -gt $fsItem.FullName.Length) {
                        $relativeSub = [System.IO.Path]::GetDirectoryName($fileItem.FullName.Substring($fsItem.FullName.Length).TrimStart('\', '/'))
                    }

                    $candidateList.Add([PSCustomObject]@{
                        FullName      = $fileItem.FullName
                        Name          = $fileItem.Name
                        Length        = $fileItem.Length
                        RelativeSub   = $relativeSub
                        DirectoryName = $fileItem.DirectoryName
                    })
                }
            } else {
                # Single file
                if (-not $processedFiles.ContainsKey($fsItem.FullName.ToLowerInvariant())) {
                    $processedFiles[$fsItem.FullName.ToLowerInvariant()] = $true
                    $candidateList.Add([PSCustomObject]@{
                        FullName      = $fsItem.FullName
                        Name          = $fsItem.Name
                        Length        = $fsItem.Length
                        RelativeSub   = ""
                        DirectoryName = $fsItem.DirectoryName
                    })
                }
            }
        }
    }

    return $candidateList
}

Write-Log "Scanning and resolving source files matching parameters..." -Level INFO
$filesToUpload = Resolve-UploadCandidates -SourcePaths $Path -FileFilter $Filter -IsRecursive:$Recurse -KeepHierarchy:$PreserveStructure

if ($filesToUpload.Count -eq 0) {
    Write-Log "No matching files found to upload for the specified path/filter parameters." -Level ERROR
    exit 1
}

$totalPayloadBytes = [long]0
foreach ($f in $filesToUpload) { $totalPayloadBytes += $f.Length }

# ==============================================================================
# Authentication Subsystem: Entra ID App-Only / REST / SDK
# ==============================================================================

function Get-EntraIdToken {
    param(
        [string]$Tenant,
        [string]$Client,
        [string]$Secret
    )

    $tokenUrl = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
    $body = @{
        client_id     = $Client
        client_secret = $Secret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    return $response.access_token
}

function Connect-AuthenticationBackend {
    # Check for environmental credentials fallback
    $effTenant = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { $env:M365_TENANT_ID }
    $effClient = if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $ClientId } else { $env:M365_CLIENT_ID }
    
    $rawSecret = $null
    if ($ClientSecretSecure) {
        $rawSecret = [System.Net.NetworkCredential]::new('', $ClientSecretSecure).Password
    } elseif (-not [string]::IsNullOrWhiteSpace($ClientSecret)) {
        $rawSecret = $ClientSecret
    } elseif (-not [string]::IsNullOrWhiteSpace($env:M365_CLIENT_SECRET)) {
        $rawSecret = $env:M365_CLIENT_SECRET
    }

    $isAppOnly = (-not [string]::IsNullOrWhiteSpace($effTenant)) -and `
                 (-not [string]::IsNullOrWhiteSpace($effClient)) -and `
                 (-not [string]::IsNullOrWhiteSpace($rawSecret))

    if ($isAppOnly -and -not $Interactive) {
        Write-Log "Authenticating via Entra ID App Registration (Client Credentials)..." -Level INFO
        try {
            $token = Get-EntraIdToken -Tenant $effTenant -Client $effClient -Secret $rawSecret
            $script:GraphToken = $token
            $script:AuthType = "Entra ID App-Only (Client Credentials)"
            Write-Log "Successfully acquired Microsoft Graph Bearer token via Entra ID App ($effClient)." -Level SUCCESS
            return
        } catch {
            Write-Log "Failed to acquire token via Entra ID Client Credentials: $($_.Exception.Message)" -Level ERROR
            throw $_
        }
    }

    # Certificate-based App-Only login via Microsoft.Graph SDK
    if ((-not [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or -not [string]::IsNullOrWhiteSpace($CertificatePath)) -and -not [string]::IsNullOrWhiteSpace($effTenant) -and -not [string]::IsNullOrWhiteSpace($effClient)) {
        Write-Log "Authenticating via Entra ID Certificate..." -Level INFO
        if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
            Write-Log "Missing required module 'Microsoft.Graph.Authentication' for certificate auth." -Level ERROR
            throw "Module Microsoft.Graph.Authentication is required for certificate login."
        }
        $certParams = @{
            TenantId = $effTenant
            ClientId = $effClient
            NoWelcome = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
            $certParams['CertificateThumbprint'] = $CertificateThumbprint
        } else {
            $certParams['CertificatePath'] = $CertificatePath
        }
        Connect-MgGraph @certParams -ErrorAction Stop
        $script:AuthType = "Entra ID App-Only (Certificate)"
        Write-Log "Connected to Microsoft Graph via Certificate ($effClient)." -Level SUCCESS
        return
    }

    # Check for active Microsoft Graph SDK session
    $hasExistingSession = $false
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -ne $ctx -and -not [string]::IsNullOrWhiteSpace($ctx.Account)) {
                $hasExistingSession = $true
                $script:AuthType = "Microsoft.Graph Existing Session ($($ctx.Account))"
                Write-Log "Reusing active Microsoft Graph session: $($ctx.Account)" -Level SUCCESS
                return
            }
        } catch {}
    }

    # Interactive Authentication
    Write-Log "No non-interactive credentials provided. Initiating interactive sign-in..." -Level INFO
    $requiredModules = @("Microsoft.Graph.Authentication")
    $missingMods = @($requiredModules | Where-Object { -not (Get-Module -Name $_ -ListAvailable) })
    if ($missingMods.Count -gt 0) {
        Write-Log "Module 'Microsoft.Graph.Authentication' is required for interactive login." -Level WARN
        $consent = if ($Force) { "Y" } else { Read-Host "[?] Install Microsoft.Graph.Authentication from PSGallery? (Y/N)" }
        if ($consent -match "^(Y|y|T|t)$") {
            Write-Log "Installing Microsoft.Graph.Authentication module..." -Level INFO
            Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
        } else {
            throw "Interactive authentication cannot proceed without Microsoft.Graph.Authentication."
        }
    }

    $connectParams = @{
        Scopes    = @("Files.ReadWrite.All", "User.Read.All")
        NoWelcome = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
        $connectParams['UserPrincipalName'] = $AdminUserPrincipalName
    }

    Connect-MgGraph @connectParams -ErrorAction Stop
    $ctx = Get-MgContext
    $script:AuthType = "Interactive Login ($($ctx.Account))"
    Write-Log "Successfully authenticated interactively as $($ctx.Account)." -Level SUCCESS
}

Connect-AuthenticationBackend

# Initialize reusable HttpClient
$script:HttpClient = [System.Net.Http.HttpClient]::new()
$script:HttpClient.Timeout = [TimeSpan]::FromMinutes(10)

# Unified Graph API invoker
function Invoke-GraphRequestDirect {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Method = 'GET',

        [Parameter(Mandatory = $false)]
        [object]$Body = $null
    )

    if (-not [string]::IsNullOrWhiteSpace($script:GraphToken)) {
        $headers = @{
            "Authorization" = "Bearer $($script:GraphToken)"
        }
        $params = @{
            Uri         = $Uri
            Method      = $Method
            Headers     = $headers
            ErrorAction = 'Stop'
        }
        if ($null -ne $Body) {
            $params['Body'] = if ($Body -is [string]) { $Body } else { ($Body | ConvertTo-Json -Depth 10 -Compress) }
            $params['ContentType'] = 'application/json'
        }
        return (Invoke-RestMethod @params)
    }

    if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $params = @{
            Uri         = $Uri
            Method      = $Method
            ErrorAction = 'Stop'
        }
        if ($null -ne $Body) {
            $params['Body'] = if ($Body -is [string]) { $Body } else { ($Body | ConvertTo-Json -Depth 10 -Compress) }
            $params['ContentType'] = 'application/json'
        }
        return (Invoke-MgGraphRequest @params)
    }

    throw "No active authentication mechanism available to query Microsoft Graph."
}

# ==============================================================================
# OneDrive Drive & Folder Resolution (with Auto-Provisioning)
# ==============================================================================

Write-Log "Querying Microsoft 365 OneDrive drive endpoint for: $UserPrincipalName..." -Level INFO
$driveEndpoint = "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($UserPrincipalName))/drive"
$userDrive = $null
try {
    $userDrive = Invoke-GraphRequestDirect -Uri $driveEndpoint -Method GET
} catch {
    Write-Log "Failed to access OneDrive for user '$UserPrincipalName'. Verify user has an active OneDrive license and permissions: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$driveId = $userDrive.id
$driveType = $userDrive.driveType
Write-Log "Resolved OneDrive Drive ID: $driveId (Type: $driveType)" -Level SUCCESS

# Auto-provision folder hierarchy
function Ensure-RemoteFolderHierarchy {
    param(
        [string]$TargetDriveId,
        [string]$FolderPath
    )

    $segments = $FolderPath.Split([char[]]@('/', '\'), [System.StringSplitOptions]::RemoveEmptyEntries)
    $currentParentId = "root"

    foreach ($seg in $segments) {
        $encodedName = [Uri]::EscapeDataString($seg)
        # Search if child folder exists under current parent
        $checkUri = "https://graph.microsoft.com/v1.0/drives/$TargetDriveId/items/$currentParentId/children?`$filter=name eq '$encodedName' and folder ne null&`$select=id,name"
        
        $match = $null
        try {
            $resp = Invoke-GraphRequestDirect -Uri $checkUri -Method GET
            if ($resp.value -and $resp.value.Count -gt 0) {
                $match = $resp.value[0]
            }
        } catch {
            Write-Log "Folder search query warning for '$seg': $($_.Exception.Message)" -Level DEBUG
        }

        if ($null -ne $match) {
            $currentParentId = $match.id
        } else {
            # Create the folder
            Write-Log "Folder '$seg' not found in OneDrive hierarchy. Creating folder..." -Level INFO
            $createUri = "https://graph.microsoft.com/v1.0/drives/$TargetDriveId/items/$currentParentId/children"
            $folderPayload = @{
                "name"   = $seg
                "folder" = @{}
                "@microsoft.graph.conflictBehavior" = "fail"
            }
            try {
                $createdItem = Invoke-GraphRequestDirect -Uri $createUri -Method POST -Body $folderPayload
                $currentParentId = $createdItem.id
                Write-Log "Successfully created remote OneDrive folder '$seg' (ID: $currentParentId)." -Level SUCCESS
            } catch {
                # Handle possible race condition if folder was created concurrently
                $verifyResp = Invoke-GraphRequestDirect -Uri $checkUri -Method GET
                if ($verifyResp.value -and $verifyResp.value.Count -gt 0) {
                    $currentParentId = $verifyResp.value[0].id
                } else {
                    throw "Failed to create remote folder '$seg' in OneDrive: $($_.Exception.Message)"
                }
            }
        }
    }

    return $currentParentId
}

Write-Log "Verifying and provisioning destination folder hierarchy in OneDrive: '$normalizedTargetFolder'..." -Level INFO
$rootTargetFolderId = Ensure-RemoteFolderHierarchy -TargetDriveId $driveId -FolderPath $normalizedTargetFolder
Write-Log "Destination folder resolved in OneDrive (Folder ID: $rootTargetFolderId)." -Level SUCCESS

# ==============================================================================
# Execution Parameters Display Banner
# ==============================================================================

$bannerLines = @(
    "================================================================================",
    "                   UPLOAD EXECUTION PARAMETERS & PROFILE",
    "================================================================================",
    ("  Target User (UPN)     : {0}" -f $UserPrincipalName),
    ("  OneDrive Drive ID     : {0} ({1})" -f $driveId, $driveType),
    ("  Destination Folder    : /{0}" -f $normalizedTargetFolder),
    ("  Total Files Matched   : {0}" -f $filesToUpload.Count),
    ("  Total Payload Size    : {0} ({1:N0} bytes)" -f (Format-ByteSize $totalPayloadBytes), $totalPayloadBytes),
    ("  Configured Chunk Size : {0} ({1:N0} bytes, compliant 320 KiB multiple)" -f (Format-ByteSize $compliantChunkSize), $compliantChunkSize),
    ("  Collision Behavior    : {0}" -f $ConflictBehavior),
    ("  Preserve Structure    : {0}" -f [string]$PreserveStructure.IsPresent),
    ("  Authentication Mode   : {0}" -f $script:AuthType),
    ("  Log File Path         : {0}" -f (if ($script:ActiveLogFile) { $script:ActiveLogFile } else { "Disabled" })),
    "================================================================================"
)
Write-Host ($bannerLines -join "`n") -ForegroundColor DarkCyan

# ==============================================================================
# Chunk Upload Helper (HttpClient Stream Streaming)
# ==============================================================================

function Send-UploadSessionChunk {
    param(
        [System.Net.Http.HttpClient]$Client,
        [string]$UploadUrl,
        [byte[]]$Buffer,
        [int]$Count,
        [long]$RangeStart,
        [long]$RangeEnd,
        [long]$TotalFileBytes,
        [int]$MaxRetries = 3
    )

    $retryCount = 0
    $success = $false
    $lastError = $null
    $responseObj = $null

    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Put, $UploadUrl)
            
            $byteContent = [System.Net.Http.ByteArrayContent]::new($Buffer, 0, $Count)
            $byteContent.Headers.ContentRange = [System.Net.Http.Headers.ContentRangeHeaderValue]::new($RangeStart, $RangeEnd, $TotalFileBytes)
            $byteContent.Headers.ContentLength = $Count
            $request.Content = $byteContent

            $response = $Client.SendAsync($request).GetAwaiter().GetResult()
            $statusCode = [int]$response.StatusCode
            $contentStr = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

            if ($statusCode -in @(200, 201, 202)) {
                $success = $true
                if (-not [string]::IsNullOrWhiteSpace($contentStr)) {
                    try { $responseObj = ($contentStr | ConvertFrom-Json) } catch {}
                }
                return @{
                    Success    = $true
                    StatusCode = $statusCode
                    Data       = $responseObj
                    Error      = $null
                }
            } elseif ($statusCode -eq 429 -or $statusCode -ge 500) {
                $retryCount++
                $delaySec = 2 * $retryCount
                if ($response.Headers.RetryAfter -and $response.Headers.RetryAfter.Delta) {
                    $delaySec = [Math]::Max(1, [int]$response.Headers.RetryAfter.Delta.TotalSeconds)
                }
                Write-Log "Chunk upload received HTTP $statusCode. Retrying in $delaySec s (Attempt $retryCount of $MaxRetries)..." -Level WARN
                Start-Sleep -Seconds $delaySec
            } else {
                throw "HTTP error $($statusCode): $contentStr"
            }
        } catch {
            $retryCount++
            $lastError = $_
            if ($retryCount -lt $MaxRetries) {
                $delaySec = 2 * $retryCount
                Write-Log "Chunk upload attempt $retryCount failed: $($_.Exception.Message). Retrying in $delaySec s..." -Level WARN
                Start-Sleep -Seconds $delaySec
            }
        }
    }

    return @{
        Success    = $false
        StatusCode = 0
        Data       = $null
        Error      = $lastError
    }
}

# ==============================================================================
# Main File Upload Execution Loop
# ==============================================================================

$overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$uploadedTotalBytes = [long]0
$resultsList = [System.Collections.Generic.List[PSCustomObject]]::new()
$successfulFiles = 0
$failedFiles = 0

Write-Log "Beginning upload session sequence for $($filesToUpload.Count) file(s)..." -Level INFO

for ($fileIndex = 0; $fileIndex -lt $filesToUpload.Count; $fileIndex++) {
    $currentCandidate = $filesToUpload[$fileIndex]
    $fileNumber = $fileIndex + 1
    $fileName = $currentCandidate.Name
    $filePath = $currentCandidate.FullName
    $fileSize = $currentCandidate.Length

    # Resolve target subfolder if preserving directory structure
    $destinationFolderId = $rootTargetFolderId
    if ($PreserveStructure -and -not [string]::IsNullOrWhiteSpace($currentCandidate.RelativeSub)) {
        $subPathNormalized = $currentCandidate.RelativeSub.Replace('\', '/').Trim('/')
        if (-not [string]::IsNullOrWhiteSpace($subPathNormalized)) {
            $fullSubPath = "$normalizedTargetFolder/$subPathNormalized"
            $destinationFolderId = Ensure-RemoteFolderHierarchy -TargetDriveId $driveId -FolderPath $fullSubPath
        }
    }

    Write-Log "[$fileNumber/$($filesToUpload.Count)] Processing: $fileName ($(Format-ByteSize $fileSize))..." -Level INFO
    $fileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Case A: 0-byte file (Upload sessions do not accept 0-byte files, use direct simple PUT)
    if ($fileSize -eq 0) {
        try {
            $encodedFileName = [Uri]::EscapeDataString($fileName)
            $directPutUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/${destinationFolderId}:/${encodedFileName}:/content"
            $uploadResult = Invoke-GraphRequestDirect -Uri $directPutUri -Method PUT -Body ""
            
            $fileStopwatch.Stop()
            $successfulFiles++
            Write-Log "[$fileNumber/$($filesToUpload.Count)] Succeeded: $fileName (Empty 0-byte file created in $($fileStopwatch.Elapsed.TotalSeconds.ToString('F1'))s)" -Level SUCCESS
            
            $resultsList.Add([PSCustomObject]@{
                FileName         = $fileName
                LocalPath        = $filePath
                Size             = 0
                SizeBytes        = 0
                DurationSeconds  = [Math]::Round($fileStopwatch.Elapsed.TotalSeconds, 2)
                AverageSpeedMBps = 0
                Status           = 'Success'
                DriveItemId      = $uploadResult.id
                WebUrl           = $uploadResult.webUrl
                Error            = $null
            })
            continue
        } catch {
            $fileStopwatch.Stop()
            $failedFiles++
            Write-Log "[$fileNumber/$($filesToUpload.Count)] Failed 0-byte upload for ${fileName}: $($_.Exception.Message)" -Level ERROR
            $resultsList.Add([PSCustomObject]@{
                FileName         = $fileName
                LocalPath        = $filePath
                Size             = 0
                SizeBytes        = 0
                DurationSeconds  = [Math]::Round($fileStopwatch.Elapsed.TotalSeconds, 2)
                AverageSpeedMBps = 0
                Status           = 'Failed'
                DriveItemId      = $null
                WebUrl           = $null
                Error            = $_.Exception.Message
            })
            continue
        }
    }

    # Case B: Standard / Large File Upload via createUploadSession
    $uploadSessionUrl = $null
    try {
        $encodedFileName = [Uri]::EscapeDataString($fileName)
        $sessionEndpoint = "https://graph.microsoft.com/v1.0/drives/$driveId/items/${destinationFolderId}:/${encodedFileName}:/createUploadSession"
        $sessionBody = @{
            "item" = @{
                "@microsoft.graph.conflictBehavior" = $ConflictBehavior
            }
        }
        $sessionResponse = Invoke-GraphRequestDirect -Uri $sessionEndpoint -Method POST -Body $sessionBody
        $uploadSessionUrl = $sessionResponse.uploadUrl
        Write-Log "  [*] Upload session created for $fileName (Expires: $($sessionResponse.expirationDateTime))" -Level DEBUG
    } catch {
        $fileStopwatch.Stop()
        $failedFiles++
        Write-Log "[$fileNumber/$($filesToUpload.Count)] Failed to create upload session for ${fileName}: $($_.Exception.Message)" -Level ERROR
        $resultsList.Add([PSCustomObject]@{
            FileName         = $fileName
            LocalPath        = $filePath
            Size             = (Format-ByteSize $fileSize)
            SizeBytes        = $fileSize
            DurationSeconds  = [Math]::Round($fileStopwatch.Elapsed.TotalSeconds, 2)
            AverageSpeedMBps = 0
            Status           = 'Failed'
            DriveItemId      = $null
            WebUrl           = $null
            Error            = "Failed to create upload session: $($_.Exception.Message)"
        })
        continue
    }

    # Stream & Chunk Transmission Loop
    $totalChunks = [int][Math]::Ceiling($fileSize / $compliantChunkSize)
    $fileBytesTransferred = [long]0
    $fileStream = $null
    $fileUploadSuccess = $false
    $completedDriveItem = $null
    $errorMessage = $null

    try {
        $fileStream = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $buffer = New-Object byte[] $compliantChunkSize

        for ($chunkIndex = 0; $chunkIndex -lt $totalChunks; $chunkIndex++) {
            $chunkNumber = $chunkIndex + 1
            $rangeStart = $fileStream.Position
            $bytesToRead = [int][Math]::Min($compliantChunkSize, $fileSize - $rangeStart)

            $readActual = 0
            while ($readActual -lt $bytesToRead) {
                $n = $fileStream.Read($buffer, $readActual, $bytesToRead - $readActual)
                if ($n -eq 0) { break }
                $readActual += $n
            }

            $rangeEnd = $rangeStart + $readActual - 1
            $chunkStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            $chunkResult = Send-UploadSessionChunk -Client $script:HttpClient `
                                                   -UploadUrl $uploadSessionUrl `
                                                   -Buffer $buffer `
                                                   -Count $readActual `
                                                   -RangeStart $rangeStart `
                                                   -RangeEnd $rangeEnd `
                                                   -TotalFileBytes $fileSize

            $chunkStopwatch.Stop()

            if (-not $chunkResult.Success) {
                throw "Chunk $chunkNumber of $totalChunks upload failed: $($chunkResult.Error.Exception.Message)"
            }

            $fileBytesTransferred += $readActual
            $uploadedTotalBytes += $readActual

            # Performance Telemetry Calculations
            $chunkSec = [Math]::Max(0.001, $chunkStopwatch.Elapsed.TotalSeconds)
            $chunkSpeedMBps = [Math]::Round(($readActual / 1MB) / $chunkSec, 2)
            $filePercent = [Math]::Round(($fileBytesTransferred / $fileSize) * 100, 1)

            $overallPercent = if ($totalPayloadBytes -gt 0) {
                [Math]::Min(100, [Math]::Round(($uploadedTotalBytes / $totalPayloadBytes) * 100, 1))
            } else { 100 }

            $totalElapsedSec = [Math]::Max(0.1, $overallStopwatch.Elapsed.TotalSeconds)
            $overallAvgSpeedMBps = [Math]::Round(($uploadedTotalBytes / 1MB) / $totalElapsedSec, 2)
            $remainingBytes = [Math]::Max(0, $totalPayloadBytes - $uploadedTotalBytes)
            $etaSeconds = if ($overallAvgSpeedMBps -gt 0) {
                [int]($remainingBytes / ($overallAvgSpeedMBps * 1MB))
            } else { 0 }
            $etaSpan = [TimeSpan]::FromSeconds($etaSeconds)

            # Update Detailed Console Progress Bars
            $overallStatusText = ("File {0}/{1} ({2}% overall) | Transferred: {3}/{4} | Avg: {5} MB/s | ETA: {6}" -f `
                $fileNumber, $filesToUpload.Count, $overallPercent, `
                (Format-ByteSize $uploadedTotalBytes), (Format-ByteSize $totalPayloadBytes), `
                $overallAvgSpeedMBps, (Format-TimeSpan $etaSpan))

            $fileStatusText = ("Chunk {0}/{1} ({2}%) | Range: {3:N0}-{4:N0} / {5:N0} bytes | Speed: {6} MB/s" -f `
                $chunkNumber, $totalChunks, $filePercent, `
                $rangeStart, $rangeEnd, $fileSize, $chunkSpeedMBps)

            Write-Progress -Id 1 -Activity "Uploading to OneDrive ($UserPrincipalName)" -Status $overallStatusText -PercentComplete $overallPercent
            Write-Progress -Id 2 -ParentId 1 -Activity "Current File: $fileName" -Status $fileStatusText -PercentComplete $filePercent

            # Log periodic milestone updates
            if ($chunkNumber -eq 1 -or $chunkNumber -eq $totalChunks -or ($chunkNumber % 5 -eq 0)) {
                Write-Log ("  --> [{0}/{1}] Chunk {2}/{3} ({4}%) uploaded at {5} MB/s" -f `
                    $fileNumber, $filesToUpload.Count, $chunkNumber, $totalChunks, $filePercent, $chunkSpeedMBps) -Level DEBUG
            }

            if ($chunkResult.StatusCode -in @(200, 201)) {
                $fileUploadSuccess = $true
                $completedDriveItem = $chunkResult.Data
            }
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "[$fileNumber/$($filesToUpload.Count)] Upload aborted for ${fileName}: $errorMessage" -Level ERROR
    } finally {
        if ($null -ne $fileStream) {
            $fileStream.Dispose()
            $fileStream = $null
        }
    }

    $fileStopwatch.Stop()

    if ($fileUploadSuccess) {
        $successfulFiles++
        $fileDurSec = [Math]::Max(0.01, $fileStopwatch.Elapsed.TotalSeconds)
        $avgSpeedMBps = [Math]::Round(($fileSize / 1MB) / $fileDurSec, 2)
        Write-Log ("[$fileNumber/{0}] Completed: {1} ({2} in {3}s, avg {4} MB/s) -> Item ID: {5}" -f `
            $filesToUpload.Count, $fileName, (Format-ByteSize $fileSize), `
            $fileDurSec.ToString('F1'), $avgSpeedMBps, $completedDriveItem.id) -Level SUCCESS

        $resultsList.Add([PSCustomObject]@{
            FileName         = $fileName
            LocalPath        = $filePath
            Size             = (Format-ByteSize $fileSize)
            SizeBytes        = $fileSize
            DurationSeconds  = [Math]::Round($fileDurSec, 2)
            AverageSpeedMBps = $avgSpeedMBps
            Status           = 'Success'
            DriveItemId      = $completedDriveItem.id
            WebUrl           = $completedDriveItem.webUrl
            Error            = $null
        })
    } else {
        $failedFiles++
        # Attempt to cancel abandoned upload session
        if (-not [string]::IsNullOrWhiteSpace($uploadSessionUrl)) {
            try {
                $cancelReq = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Delete, $uploadSessionUrl)
                $script:HttpClient.SendAsync($cancelReq).GetAwaiter().GetResult() | Out-Null
            } catch {}
        }

        $resultsList.Add([PSCustomObject]@{
            FileName         = $fileName
            LocalPath        = $filePath
            Size             = (Format-ByteSize $fileSize)
            SizeBytes        = $fileSize
            DurationSeconds  = [Math]::Round($fileStopwatch.Elapsed.TotalSeconds, 2)
            AverageSpeedMBps = 0
            Status           = 'Failed'
            DriveItemId      = $null
            WebUrl           = $null
            Error            = $errorMessage
        })
    }
}

# Complete and clear progress bars
try {
    Write-Progress -Id 2 -Activity "Current File" -Completed
    Write-Progress -Id 1 -Activity "Uploading to OneDrive" -Completed
} catch {}

$overallStopwatch.Stop()

# Dispose HttpClient
if ($null -ne $script:HttpClient) {
    $script:HttpClient.Dispose()
    $script:HttpClient = $null
}

# ==============================================================================
# Execution Summary Report
# ==============================================================================

$totalDurationSec = [Math]::Max(0.1, $overallStopwatch.Elapsed.TotalSeconds)
$overallThroughputMBps = [Math]::Round(($uploadedTotalBytes / 1MB) / $totalDurationSec, 2)

$summaryReport = @(
    "",
    "================================================================================",
    "                    UPLOAD EXECUTION SUMMARY REPORT",
    "================================================================================",
    ("  Target User (OneDrive)  : {0}" -f $UserPrincipalName),
    ("  Destination Folder      : /{0}" -f $normalizedTargetFolder),
    ("  Total Files Attempted   : {0}" -f $filesToUpload.Count),
    ("  Files Succeeded         : {0}" -f $successfulFiles),
    ("  Files Failed            : {0}" -f $failedFiles),
    ("  Total Payload Uploaded  : {0} ({1:N0} bytes)" -f (Format-ByteSize $uploadedTotalBytes), $uploadedTotalBytes),
    ("  Total Execution Time    : {0}" -f (Format-TimeSpan $overallStopwatch.Elapsed)),
    ("  Average Transfer Speed  : {0} MB/s" -f $overallThroughputMBps),
    ("  Authentication Mode     : {0}" -f $script:AuthType),
    ("  Log File Location       : {0}" -f (if ($script:ActiveLogFile) { $script:ActiveLogFile } else { "Disabled" })),
    "================================================================================"
)

$summaryColor = if ($failedFiles -eq 0) { [ConsoleColor]::Green } else { [ConsoleColor]::Yellow }
Write-Host ($summaryReport -join "`n") -ForegroundColor $summaryColor

foreach ($line in $summaryReport) {
    if (-not [string]::IsNullOrWhiteSpace($script:ActiveLogFile)) {
        try { [System.IO.File]::AppendAllText($script:ActiveLogFile, "$line`r`n", [System.Text.Encoding]::UTF8) } catch {}
    }
}

if ($PassThru) {
    Write-Output $resultsList
}

if ($failedFiles -gt 0) {
    Write-Log "Execution finished with $failedFiles failed upload(s). Review log file for details." -Level WARN
    exit 2
} else {
    Write-Log "All files uploaded successfully to OneDrive." -Level SUCCESS
    exit 0
}
