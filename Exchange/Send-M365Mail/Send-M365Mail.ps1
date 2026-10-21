<#
.SYNOPSIS
    Sends emails via Microsoft 365 (Office 365) using Microsoft Graph API with support for interactive sign-in or non-interactive certificate-based authentication, attachments, and comprehensive enterprise logging.

.DESCRIPTION
    Send-M365Mail is a production-ready PowerShell automation tool designed for Microsoft 365 and Exchange Online.
    It allows users and automation workflows to authenticate securely via modern authentication (Entra ID / OAuth 2.0)
    and send emails with custom subjects, plain text or HTML bodies, and multiple file attachments.

    Key Features:
    - Dual Authentication Modes:
      * Automated Certificate-based sign-in (scans Windows Certificate Store for installed M365/Graph certificates or accepts -CertificateThumbprint).
      * Interactive sign-in fallback (delegated browser/MFA session via Connect-MgGraph).
    - Client Configuration & Deployment (-ConfigureClient, -GenerateCertificate):
      * Automatically deploys and caches client configuration (TenantId, ClientId, DefaultSender) in %APPDATA%\Send-M365Mail\config.json.
      * Includes built-in certificate generator for unattended background operations.
    - Attachment Handling:
      * Validates file existence, MIME types, and sizes.
      * Automatically encodes files into RFC-compliant Base64 Graph attachments.
    - Enterprise Logging:
      * Automatically writes detailed execution logs to C:\Logs\<timestamp>-Send-M365Mail-<computer>-<user>.logs.
      * Color-coded console tracking of every single pipeline step.
    - Security & Robustness:
      * Strict email address validation using RFC 5322 regex.
      * Language-independent Windows checks (supports Polish and English Windows editions).
      * Safe input handling and detailed execution summary with actionable user guidance.

.PARAMETER To
    One or more recipient email addresses.
    Aliases: -Recipient, -RecipientEmail, -User.

.PARAMETER Subject
    The subject line of the email message.
    Alias: -Title.

.PARAMETER Body
    The message content of the email. Supports plain text or rich HTML.
    Aliases: -Content, -Message.

.PARAMETER AttachmentPath
    One or more local file paths to attach to the email message.
    Aliases: -Attachment, -Attachments, -FilePath.

.PARAMETER IsHtmlBody
    Switch parameter. Indicates that the message body content is formatted as HTML.
    Aliases: -Html, -HtmlBody.

.PARAMETER From
    The sender mailbox address. If running in Application/Certificate mode, specifies which mailbox sends the message.
    In Delegated/Interactive mode, defaults to the signed-in user's mailbox.
    Aliases: -Sender, -SenderEmail.

.PARAMETER Cc
    Optional CC (Carbon Copy) recipient email addresses.

.PARAMETER Bcc
    Optional BCC (Blind Carbon Copy) recipient email addresses.

.PARAMETER Importance
    Message importance level: 'Low', 'Normal', or 'High'. Default is 'Normal'.

.PARAMETER SaveToSentItems
    Whether to save a copy of the sent email in the sender's 'Sent Items' folder. Default is $true.

.PARAMETER TenantId
    The Entra ID Directory (Tenant) ID used for certificate-based non-interactive authentication.
    Aliases: -Tenant, -DirectoryId.

.PARAMETER ClientId
    The Entra ID Application (Client) ID registered in Microsoft Entra.
    Aliases: -AppId, -ApplicationId.

.PARAMETER CertificateThumbprint
    The thumbprint of an X.509 certificate installed in Cert:\CurrentUser\My or Cert:\LocalMachine\My.
    Alias: -Thumbprint.

.PARAMETER CertificateFilePath
    Path to a .pfx certificate file for certificate-based authentication.
    Alias: -CertPath.

.PARAMETER CertificatePassword
    Optional SecureString password for decrypting a .pfx certificate file.

.PARAMETER Interactive
    Switch parameter. Forces interactive browser/device-code authentication even if a certificate is detected.

.PARAMETER ConfigureClient
    Switch parameter. Deploys client configuration, saves default TenantId, ClientId, and Sender to local profile config.

.PARAMETER GenerateCertificate
    Switch parameter. Generates a self-signed X.509 certificate for Entra ID app authentication and exports the .cer file.

.PARAMETER LogPath
    Custom path for the log file. If omitted, logs default to C:\Logs\<timestamp>-Send-M365Mail-<computer>-<user>.logs.

.PARAMETER NoLog
    Switch parameter. Disables writing execution logs to disk.

.PARAMETER Help
    Displays custom help, usage instructions, author information, Entra ID setup guide, and exits.
    Alias: -h.

.EXAMPLE
    # 1. Display full interactive help and Entra ID configuration guide
    .\Send-M365Mail.ps1 -Help

.EXAMPLE
    # 2. Interactive Email Sending (auto-prompts login if not authenticated)
    .\Send-M365Mail.ps1 -To "john.doe@contoso.com" `
        -Subject "Monthly Financial Report" `
        -Body "Hello John, attached is the report." `
        -AttachmentPath "C:\Reports\Finance_Oct2026.pdf"

.EXAMPLE
    # 3. HTML Email with Multiple Attachments and High Importance
    .\Send-M365Mail.ps1 -To "board@company.com" -Cc "ceo@company.com" `
        -Subject "Urgent: System Maintenance Announcement" `
        -Body "<h3>Notice</h3><p>Maintenance scheduled for tonight.</p>" `
        -IsHtmlBody `
        -AttachmentPath "C:\Notices\Plan.pdf", "C:\Notices\Schedule.xlsx" `
        -Importance High

.EXAMPLE
    # 4. Unattended / Automated Email using Certificate
    .\Send-M365Mail.ps1 -To "alerts@company.com" `
        -From "notifications@company.com" `
        -Subject "Server Backup Status: Success" `
        -Body "All database backups finished successfully." `
        -TenantId "11111111-2222-3333-4444-555555555555" `
        -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -CertificateThumbprint "A1B2C3D4E5F60718293A4B5C6D7E8F9012345678"

.EXAMPLE
    # 5. One-time Client Setup (Store credentials & generate certificate)
    .\Send-M365Mail.ps1 -ConfigureClient -TenantId "11111111-2222-3333-4444-555555555555" -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
    .\Send-M365Mail.ps1 -GenerateCertificate

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Send')]
param(
    [Parameter(ParameterSetName = 'Send', Position = 0)]
    [Alias('Recipient', 'RecipientEmail', 'User')]
    [string[]]$To,

    [Parameter(ParameterSetName = 'Send', Position = 1)]
    [Alias('Title')]
    [string]$Subject,

    [Parameter(ParameterSetName = 'Send', Position = 2)]
    [Alias('Content', 'Message')]
    [string]$Body,

    [Parameter(ParameterSetName = 'Send')]
    [Alias('Attachment', 'Attachments', 'FilePath')]
    [string[]]$AttachmentPath,

    [Parameter(ParameterSetName = 'Send')]
    [Alias('Html', 'HtmlBody')]
    [switch]$IsHtmlBody,

    [Alias('Sender', 'SenderEmail')]
    [string]$From,

    [Parameter(ParameterSetName = 'Send')]
    [string[]]$Cc,

    [Parameter(ParameterSetName = 'Send')]
    [string[]]$Bcc,

    [Parameter(ParameterSetName = 'Send')]
    [ValidateSet('Low', 'Normal', 'High')]
    [string]$Importance = 'Normal',

    [Parameter(ParameterSetName = 'Send')]
    [bool]$SaveToSentItems = $true,

    # Authentication Parameters
    [Alias('Tenant', 'DirectoryId')]
    [string]$TenantId,

    [Alias('AppId', 'ApplicationId')]
    [string]$ClientId,

    [Alias('Thumbprint')]
    [string]$CertificateThumbprint,

    [Alias('CertPath')]
    [string]$CertificateFilePath,

    [System.Security.SecureString]$CertificatePassword,

    [Alias('Secret')]
    [string]$ClientSecret,

    [switch]$Interactive,

    [Alias('DeviceCode')]
    [switch]$UseDeviceAuthentication,

    # Client Configuration & Setup
    [Parameter(ParameterSetName = 'Configure', Mandatory = $true)]
    [switch]$ConfigureClient,

    [Parameter(ParameterSetName = 'CertGen', Mandatory = $true)]
    [switch]$GenerateCertificate,

    # Logging Parameters
    [string]$LogPath,

    [switch]$NoLog,

    # Help Parameter
    [Parameter(ParameterSetName = 'Help', Mandatory = $true)]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# ENVIRONMENT & MODULE PATH BOOTSTRAP
# ==============================================================================
# Ensure %LOCALAPPDATA%\WindowsPowerShell\Modules is always in $env:PSModulePath
# (bypasses Windows Defender Controlled Folder Access block on Documents)
$localAppData = [Environment]::GetFolderPath('LocalApplicationData')
$global:LocalAppDataModules = $null
if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
    $global:LocalAppDataModules = Join-Path -Path $localAppData -ChildPath "WindowsPowerShell\Modules"
    if (-not (Test-Path -Path $global:LocalAppDataModules)) {
        try { New-Item -Path $global:LocalAppDataModules -ItemType Directory -Force | Out-Null } catch {}
    }
    if ($env:PSModulePath -notlike "*$global:LocalAppDataModules*") {
        $env:PSModulePath = "$global:LocalAppDataModules;$env:PSModulePath"
    }
}

# ==============================================================================
# GLOBAL CONSTANTS & SCRIPT METADATA
# ==============================================================================
$global:ScriptVersion     = "1.0.0"
$global:ScriptAuthor      = "Roman Pindela"
$global:ScriptEmail       = "roman.pindela@gmail.com"
$global:ScriptGitHub      = "https://github.com/romanpindela"
$global:ScriptLogPath     = $null
$global:ConfigDirectory   = Join-Path -Path $env:APPDATA -ChildPath "Send-M365Mail"
$global:ConfigFilePath    = Join-Path -Path $global:ConfigDirectory -ChildPath "config.json"
$global:AuthMethodUsed    = "Unknown"
$global:ConnectedUserUpn  = $null
$global:ExecutionStart    = Get-Date

# ==============================================================================
# LOGGING SUBSYSTEM
# ==============================================================================
function Initialize-ScriptLogger {
    if ($NoLog) {
        $global:ScriptLogPath = $null
        return
    }

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        # Default standard path: C:\Logs\<dataigodzina>-nazwaskryptu-nazwakomputera-uzytkownik.logs
        $logsBaseDir = "C:\Logs"
        if (-not (Test-Path -Path $logsBaseDir)) {
            try {
                New-Item -Path $logsBaseDir -ItemType Directory -Force | Out-Null
            } catch {
                # Fallback to local logs directory if C:\Logs is inaccessible
                $logsBaseDir = Join-Path -Path $PSScriptRoot -ChildPath "logs"
                if (-not (Test-Path -Path $logsBaseDir)) {
                    New-Item -Path $logsBaseDir -ItemType Directory -Force | Out-Null
                }
            }
        }
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $compName  = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { "LOCAL" }
        $userName  = if ($env:USERNAME) { $env:USERNAME } else { "USER" }
        $fileName  = "${timestamp}-Send-M365Mail-${compName}-${userName}.logs"
        $global:ScriptLogPath = Join-Path -Path $logsBaseDir -ChildPath $fileName
    } else {
        $resolvedDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($resolvedDir) -and -not (Test-Path -Path $resolvedDir)) {
            New-Item -Path $resolvedDir -ItemType Directory -Force | Out-Null
        }
        $global:ScriptLogPath = $LogPath
    }

    $header = @"
================================================================================
Send-M365Mail Execution Log Session
Started     : $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss UTCzzz"))
Script Ver  : $global:ScriptVersion
Author      : $global:ScriptAuthor ($global:ScriptEmail)
GitHub      : $global:ScriptGitHub
Machine     : $env:COMPUTERNAME
User        : $env:USERNAME
Log Path    : $global:ScriptLogPath
================================================================================
"@
    try {
        Set-Content -Path $global:ScriptLogPath -Value $header -Encoding UTF8 -Force
    } catch {
        Write-Warning "Unable to create log file at '$global:ScriptLogPath': $_"
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

        [string]$Component = "M365Mail"
    )

    $timeStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $logLine = "[$timeStr] [$($Level.ToUpper().PadRight(7))] [$Component] $Message"

    switch ($Level) {
        'Success'  { Write-Host $logLine -ForegroundColor Green }
        'Warning'  { Write-Host $logLine -ForegroundColor Yellow }
        'Error'    { Write-Host $logLine -ForegroundColor Red }
        'Progress' { Write-Host $logLine -ForegroundColor Magenta }
        'Debug'    { Write-Host $logLine -ForegroundColor DarkGray }
        default    { Write-Host $logLine -ForegroundColor Cyan }
    }

    if ($global:ScriptLogPath) {
        try {
            Add-Content -Path $global:ScriptLogPath -Value $logLine -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {
            # Avoid breaking script execution on log write contention
        }
    }
}

# ==============================================================================
# CLIENT CONFIGURATION MANAGER
# ==============================================================================
function Get-ClientConfiguration {
    if (Test-Path -Path $global:ConfigFilePath) {
        try {
            $json = Get-Content -Path $global:ConfigFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
            return $json
        } catch {
            Write-LogMessage -Level Warning -Message "Failed to parse client config at '$global:ConfigFilePath': $_" -Component "Config"
        }
    }
    return $null
}

function Save-ClientConfiguration {
    param(
        [string]$Tenant,
        [string]$Client,
        [string]$DefaultSender,
        [string]$Thumbprint
    )

    if (-not (Test-Path -Path $global:ConfigDirectory)) {
        New-Item -Path $global:ConfigDirectory -ItemType Directory -Force | Out-Null
    }

    $current = Get-ClientConfiguration
    $configObj = [PSCustomObject]@{
        TenantId              = if ($Tenant) { $Tenant } elseif ($current) { $current.TenantId } else { "" }
        ClientId              = if ($Client) { $Client } elseif ($current) { $current.ClientId } else { "" }
        DefaultSender         = if ($DefaultSender) { $DefaultSender } elseif ($current) { $current.DefaultSender } else { "" }
        CertificateThumbprint = if ($Thumbprint) { $Thumbprint } elseif ($current) { $current.CertificateThumbprint } else { "" }
        UpdatedAt             = (Get-Date).ToString("s")
    }

    $json = $configObj | ConvertTo-Json -Depth 3
    Set-Content -Path $global:ConfigFilePath -Value $json -Encoding UTF8 -Force
    Write-LogMessage -Level Success -Message "Saved client configuration to '$global:ConfigFilePath'" -Component "Config"
}

# ==============================================================================
# CERTIFICATE GENERATOR (Matches New-GraphCertificate standard)
# ==============================================================================
function New-M365GraphCertificate {
    param(
        [string]$CertificateName = "M365-Graph-EmailSender",
        [string]$StoreLocation = "CurrentUser",
        [string]$ExportDirectory = "C:\M365_Certs",
        [int]$ValidityYears = 2
    )

    Write-LogMessage -Level Info -Message "Starting automated X.509 Certificate generation for Entra ID authentication..." -Component "CertGen"

    if (-not (Test-Path -Path $ExportDirectory)) {
        New-Item -ItemType Directory -Path $ExportDirectory -Force | Out-Null
        Write-LogMessage -Level Success -Message "Created export folder: $ExportDirectory" -Component "CertGen"
    }

    $subjectName = "CN=$CertificateName"
    $certStore   = "Cert:\$StoreLocation\My"

    Write-LogMessage -Level Info -Message "Generating 2048-bit RSA SHA-256 certificate '$subjectName' in $certStore..." -Component "CertGen"

    $cert = New-SelfSignedCertificate -Subject $subjectName `
        -CertStoreLocation $certStore `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($ValidityYears)

    $thumbprint = $cert.Thumbprint
    $cerFilePath = Join-Path -Path $ExportDirectory -ChildPath "$CertificateName.cer"
    Export-Certificate -Cert $cert -FilePath $cerFilePath -Type CERT | Out-Null

    Write-LogMessage -Level Success -Message "Certificate generated successfully! Thumbprint: $thumbprint" -Component "CertGen"
    Write-LogMessage -Level Success -Message "Public key exported to: $cerFilePath" -Component "CertGen"

    # Save to client config automatically
    Save-ClientConfiguration -Thumbprint $thumbprint

    Write-Host @"

================================================================================
  AUTOMATED CERTIFICATE DEPLOYMENT SUMMARY
================================================================================
  Thumbprint : $thumbprint
  Store      : $certStore
  Public Key : $cerFilePath
  Valid For  : $ValidityYears years (Expires: $((Get-Date).AddYears($ValidityYears).ToString("yyyy-MM-dd")))

NEXT STEPS:
1. Open Microsoft Entra Admin Center (https://entra.microsoft.com).
2. Go to: App registrations -> Your App -> Certificates & secrets -> Certificates.
3. Click 'Upload certificate' and select: $cerFilePath
4. Ensure Application Permission 'Mail.Send' is granted with Admin Consent.
================================================================================
"@ -ForegroundColor Yellow
}

# ==============================================================================
# INPUT VALIDATION HELPERS
# ==============================================================================
function Test-EmailAddressFormat {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    $regex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return ($Email -match $regex)
}

function Get-FileMimeType {
    param([string]$FilePath)
    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    switch ($ext) {
        '.pdf'  { return 'application/pdf' }
        '.txt'  { return 'text/plain' }
        '.html' { return 'text/html' }
        '.htm'  { return 'text/html' }
        '.csv'  { return 'text/csv' }
        '.json' { return 'application/json' }
        '.zip'  { return 'application/zip' }
        '.png'  { return 'image/png' }
        '.jpg'  { return 'image/jpeg' }
        '.jpeg' { return 'image/jpeg' }
        '.docx' { return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
        '.xlsx' { return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        '.pptx' { return 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }
        default { return 'application/octet-stream' }
    }
}

# ==============================================================================
# MICROSOFT GRAPH AUTHENTICATION ENGINE
# ==============================================================================
function Find-InstalledGraphCertificate {
    param([string]$ExplicitThumbprint)

    # 1. Search by explicit thumbprint
    if (-not [string]::IsNullOrWhiteSpace($ExplicitThumbprint)) {
        $cleanThumb = $ExplicitThumbprint.Replace(" ", "").Trim().ToUpper()
        foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
            if (Test-Path -Path $store) {
                $found = Get-ChildItem -Path $store | Where-Object { $_.Thumbprint.ToUpper() -eq $cleanThumb }
                if ($found) {
                    Write-LogMessage -Level Debug -Message "Found certificate by thumbprint in $store" -Component "Auth"
                    return $found[0]
                }
            }
        }
        return $null
    }

    # 2. Check local client config
    $clientConfig = Get-ClientConfiguration
    if ($clientConfig -and -not [string]::IsNullOrWhiteSpace($clientConfig.CertificateThumbprint)) {
        $cleanThumb = $clientConfig.CertificateThumbprint.Replace(" ", "").Trim().ToUpper()
        foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
            if (Test-Path -Path $store) {
                $found = Get-ChildItem -Path $store | Where-Object { $_.Thumbprint.ToUpper() -eq $cleanThumb }
                if ($found) {
                    Write-LogMessage -Level Info -Message "Found certificate from client config: $($found[0].Subject) [$($found[0].Thumbprint)]" -Component "Auth"
                    return $found[0]
                }
            }
        }
    }

    # 3. Auto-discover non-expired certificates with relevant subjects
    $candidateSubjects = @("M365-Graph-EmailSender", "M365-Graph", "Send-M365Mail", "Graph")
    foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
        if (Test-Path -Path $store) {
            $certs = Get-ChildItem -Path $store | Where-Object {
                $_.HasPrivateKey -and
                $_.NotAfter -gt (Get-Date) -and
                ($candidateSubjects | ForEach-Object { $sub = $_; $_certSub = $args[0].Subject; $_certSub -like "*$sub*" } -ArgumentList $_)
            }
            if ($certs -and $certs.Count -gt 0) {
                $selected = $certs[0]
                Write-LogMessage -Level Info -Message "Auto-discovered valid M365 certificate in $($store): $($selected.Subject) [$($selected.Thumbprint)]" -Component "Auth"
                return $selected
            }
        }
    }

    return $null
}

function Connect-M365Service {
    param(
        [string]$Tenant,
        [string]$Client,
        [string]$Thumbprint,
        [switch]$ForceInteractive
    )

    Write-LogMessage -Level Info -Message "Verifying Microsoft Graph connection requirements..." -Component "Auth"

    # Load defaults from config if not supplied
    $config = Get-ClientConfiguration
    if ([string]::IsNullOrWhiteSpace($Tenant) -and $config -and $config.TenantId) {
        $Tenant = $config.TenantId
        Write-LogMessage -Level Debug -Message "Loaded TenantId from config: $Tenant" -Component "Auth"
    }
    if ([string]::IsNullOrWhiteSpace($Client) -and $config -and $config.ClientId) {
        $Client = $config.ClientId
        Write-LogMessage -Level Debug -Message "Loaded ClientId from config: $Client" -Component "Auth"
    }

    # Check for Microsoft.Graph module availability
    $moduleInstalled = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication
    if (-not $moduleInstalled) {
        Write-LogMessage -Level Warning -Message "Microsoft.Graph.Authentication module not found." -Component "Auth"
        Write-Host "[!] The 'Microsoft.Graph.Authentication' module is recommended for Microsoft 365 communication." -ForegroundColor Yellow
        $installPrompt = Read-Host "Do you want to install Microsoft.Graph.Authentication now from PSGallery? (Y/N)"
        if ($installPrompt -match '^[YyTt]') {
            Write-LogMessage -Level Info -Message "Installing Microsoft.Graph.Authentication for CurrentUser..." -Component "Auth"
            Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        } else {
            throw "Microsoft.Graph.Authentication is required to proceed. Please run: Install-Module Microsoft.Graph.Authentication"
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Check for existing active session
    try {
        $currentContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($currentContext -and -not $ForceInteractive -and [string]::IsNullOrWhiteSpace($Thumbprint)) {
            Write-LogMessage -Level Success -Message "Active Microsoft Graph session detected! Account: $($currentContext.Account), Scopes: $($currentContext.Scopes -join ', ')" -Component "Auth"
            $global:AuthMethodUsed   = "Active-Session ($($currentContext.AuthType))"
            $global:ConnectedUserUpn = $currentContext.Account
            return $currentContext
        }
    } catch {
        # Proceed with new authentication
    }

    # Step A: Attempt Certificate-based authentication if not forced interactive
    if (-not $ForceInteractive) {
        $cert = Find-InstalledGraphCertificate -ExplicitThumbprint $Thumbprint
        if ($cert) {
            if ([string]::IsNullOrWhiteSpace($Tenant) -or [string]::IsNullOrWhiteSpace($Client)) {
                Write-LogMessage -Level Warning -Message "Found certificate $($cert.Thumbprint), but -TenantId and -ClientId are missing. Falling back to interactive login." -Component "Auth"
            } else {
                Write-LogMessage -Level Info -Message "Authenticating to Microsoft Graph via Certificate (App ID: $Client, Tenant: $Tenant)..." -Component "Auth"
                try {
                    Connect-MgGraph -ClientId $Client -TenantId $Tenant -CertificateThumbprint $cert.Thumbprint -NoWelcome -ErrorAction Stop
                    $global:AuthMethodUsed   = "Certificate ($($cert.Thumbprint.Substring(0, 8))...)"
                    $global:ConnectedUserUpn = "App:$Client"
                    Write-LogMessage -Level Success -Message "Successfully connected via Entra ID App Certificate!" -Component "Auth"
                    return (Get-MgContext)
                } catch {
                    Write-LogMessage -Level Error -Message "Certificate authentication failed: $_" -Component "Auth"
                    Write-LogMessage -Level Warning -Message "Attempting interactive sign-in fallback..." -Component "Auth"
                }
            }
        } else {
            Write-LogMessage -Level Info -Message "No matching X.509 certificate found in system store." -Component "Auth"
        }
    }

    # Step B: Interactive Authentication Fallback
    Write-LogMessage -Level Info -Message "Initiating Interactive browser sign-in with required scope 'Mail.Send'..." -Component "Auth"
    try {
        $scopes = @("Mail.Send", "User.Read")
        if (-not [string]::IsNullOrWhiteSpace($Tenant)) {
            Connect-MgGraph -Scopes $scopes -TenantId $Tenant -NoWelcome -ErrorAction Stop
        } else {
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
        }

        $context = Get-MgContext
        $global:AuthMethodUsed   = "Interactive (Delegated MFA)"
        $global:ConnectedUserUpn = $context.Account
        Write-LogMessage -Level Success -Message "Successfully connected interactively as: $($context.Account)" -Component "Auth"
        return $context
    } catch {
        Write-LogMessage -Level Error -Message "Failed to authenticate interactively: $_" -Component "Auth"
        throw "Authentication to Microsoft 365 failed: $_"
    }
}

# ==============================================================================
# EMAIL PREPARATION & DISPATCH ENGINE
# ==============================================================================
function Send-GraphEmailMessage {
    param(
        [string[]]$Recipients,
        [string]$MailSubject,
        [string]$MailBody,
        [string[]]$Attachments,
        [bool]$IsHtml,
        [string]$SenderAddress,
        [string[]]$CcRecipients,
        [string[]]$BccRecipients,
        [string]$Priority,
        [bool]$SaveSent
    )

    Write-LogMessage -Level Info -Message "Validating email parameters..." -Component "EmailEngine"

    # Validate recipients
    $validRecipients = @()
    foreach ($rec in $Recipients) {
        $trimmed = $rec.Trim()
        if (-not (Test-EmailAddressFormat -Email $trimmed)) {
            throw "Invalid recipient email address format: '$trimmed'"
        }
        $validRecipients += @{
            emailAddress = @{
                address = $trimmed
            }
        }
    }

    # Validate CC recipients
    $validCc = @()
    if ($CcRecipients) {
        foreach ($ccAddr in $CcRecipients) {
            $trimmed = $ccAddr.Trim()
            if (-not (Test-EmailAddressFormat -Email $trimmed)) {
                throw "Invalid CC email address format: '$trimmed'"
            }
            $validCc += @{
                emailAddress = @{
                    address = $trimmed
                }
            }
        }
    }

    # Validate BCC recipients
    $validBcc = @()
    if ($BccRecipients) {
        foreach ($bccAddr in $BccRecipients) {
            $trimmed = $bccAddr.Trim()
            if (-not (Test-EmailAddressFormat -Email $trimmed)) {
                throw "Invalid BCC email address format: '$trimmed'"
            }
            $validBcc += @{
                emailAddress = @{
                    address = $trimmed
                }
            }
        }
    }

    # Process attachments
    $processedAttachments = @()
    $totalAttachmentBytes = 0
    if ($Attachments) {
        foreach ($attPath in $Attachments) {
            $resolvedPath = Resolve-Path -Path $attPath -ErrorAction SilentlyContinue
            if (-not $resolvedPath -or -not (Test-Path -Path $resolvedPath)) {
                throw "Attachment file not found: '$attPath'"
            }

            $fileItem = Get-Item -Path $resolvedPath
            $fileBytes = [System.IO.File]::ReadAllBytes($fileItem.FullName)
            $totalAttachmentBytes += $fileBytes.Length

            if ($fileBytes.Length -gt 3145728) {
                Write-LogMessage -Level Warning -Message "Attachment '$($fileItem.Name)' is large ($([math]::Round($fileBytes.Length/1MB, 2)) MB). Inline Graph sendMail limit is typically 3-4 MB." -Component "EmailEngine"
            }

            $base64 = [Convert]::ToBase64String($fileBytes)
            $mimeType = Get-FileMimeType -FilePath $fileItem.FullName

            $processedAttachments += @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = $fileItem.Name
                contentType   = $mimeType
                contentBytes  = $base64
            }

            Write-LogMessage -Level Debug -Message "Prepared attachment: $($fileItem.Name) ($($fileBytes.Length) bytes, MIME: $mimeType)" -Component "EmailEngine"
        }
    }

    # Build Message Object
    $messageObj = @{
        subject      = $MailSubject
        body         = @{
            contentType = if ($IsHtml) { "HTML" } else { "Text" }
            content     = $MailBody
        }
        toRecipients = $validRecipients
        importance   = $Priority
    }

    if ($validCc.Count -gt 0) {
        $messageObj["ccRecipients"] = $validCc
    }
    if ($validBcc.Count -gt 0) {
        $messageObj["bccRecipients"] = $validBcc
    }
    if ($processedAttachments.Count -gt 0) {
        $messageObj["attachments"] = $processedAttachments
    }

    $requestBody = @{
        message         = $messageObj
        saveToSentItems = if ($SaveSent) { "true" } else { "false" }
    } | ConvertTo-Json -Depth 6

    # Determine endpoint: /me/sendMail or /users/{id}/sendMail
    $context = Get-MgContext
    $targetUri = "v1.0/me/sendMail"

    if (-not [string]::IsNullOrWhiteSpace($SenderAddress)) {
        $targetUri = "v1.0/users/$SenderAddress/sendMail"
        Write-LogMessage -Level Info -Message "Dispatching email on behalf of mailbox: $SenderAddress" -Component "EmailEngine"
    } elseif ($context.AuthType -eq "AppOnly") {
        # When using AppOnly (Certificate / Service Principal), /me does not exist. A sender mailbox must be provided.
        $config = Get-ClientConfiguration
        if ($config -and $config.DefaultSender) {
            $targetUri = "v1.0/users/$($config.DefaultSender)/sendMail"
            Write-LogMessage -Level Info -Message "Using configured default sender for App-Only session: $($config.DefaultSender)" -Component "EmailEngine"
        } else {
            throw "When running in non-interactive / App-Only mode (Certificate), you must specify the sender mailbox via -From 'user@domain.com' or configure default sender via -ConfigureClient."
        }
    }

    Write-LogMessage -Level Info -Message "Transmitting email via Microsoft Graph API ($targetUri)..." -Component "EmailEngine"

    try {
        Invoke-MgGraphRequest -Method POST -Uri $targetUri -Body $requestBody -ContentType "application/json" -ErrorAction Stop
        Write-LogMessage -Level Success -Message "Email successfully submitted to Microsoft 365! (HTTP 202 Accepted)" -Component "EmailEngine"
        return @{
            Success         = $true
            Endpoint        = $targetUri
            AttachmentCount = $processedAttachments.Count
            AttachmentBytes = $totalAttachmentBytes
        }
    } catch {
        Write-LogMessage -Level Error -Message "Failed to send email via Microsoft Graph API: $_" -Component "EmailEngine"
        throw $_
    }
}

# ==============================================================================
# HELP & DOCUMENTATION DISPLAY
# ==============================================================================
function Show-ScriptHelp {
    $banner = @"
================================================================================
  SEND-M365MAIL :: Microsoft 365 Automated Email Dispatcher
  Version  : $global:ScriptVersion
  Author   : $global:ScriptAuthor
  Contact  : $global:ScriptEmail
  GitHub   : $global:ScriptGitHub
================================================================================

SYNOPSIS:
    Sends emails via Microsoft 365 (Office 365) using Microsoft Graph API with
    support for interactive sign-in or non-interactive certificate-based authentication.

SYNTAX:
    .\Send-M365Mail.ps1 -To <email[]> -Subject <string> -Body <string>
                        [-AttachmentPath <path[]>] [-IsHtmlBody] [-From <email>]
                        [-Cc <email[]>] [-Bcc <email[]>] [-Importance <Low|Normal|High>]
                        [-TenantId <id>] [-ClientId <id>] [-CertificateThumbprint <thumbprint>]
                        [-Interactive] [-LogPath <path>] [-NoLog]

    .\Send-M365Mail.ps1 -ConfigureClient [-TenantId <id>] [-ClientId <id>] [-From <email>]
    .\Send-M365Mail.ps1 -GenerateCertificate [-CertificateName <name>]

EXAMPLES:
    1. Interactive Send:
       .\Send-M365Mail.ps1 -To "user@company.com" -Subject "Daily Brief" -Body "Hello Team!"

    2. Send with Attachment & HTML formatting:
       .\Send-M365Mail.ps1 -To "finance@company.com" -Subject "Invoice #1024" `
           -Body "<h3>Invoice Attached</h3><p>Please review.</p>" -IsHtmlBody `
           -AttachmentPath "C:\Invoices\1024.pdf"

    3. Unattended Automated Dispatch via Certificate:
       .\Send-M365Mail.ps1 -To "alerts@company.com" -From "service@company.com" `
           -Subject "Backup Completed" -Body "System backup finished." `
           -TenantId "YOUR-TENANT-ID" -ClientId "YOUR-CLIENT-ID" `
           -CertificateThumbprint "A1B2C3D4..."

ENTRA ID (AZURE AD) NON-INTERACTIVE SETUP GUIDE:
    1. Go to Microsoft Entra admin center (https://entra.microsoft.com).
    2. Register an application under 'App registrations'.
    3. Under 'API permissions', add Microsoft Graph -> Application permissions:
       - Mail.Send (Allows sending emails from any designated tenant mailbox)
       - User.Read.All (Recommended for identity resolution)
       Click 'Grant admin consent for <Tenant>'.
    4. Generate and attach certificate:
       Run: .\Send-M365Mail.ps1 -GenerateCertificate
       Upload the resulting C:\M365_Certs\M365-Graph-EmailSender.cer under
       'Certificates & secrets' -> 'Certificates'.

SCRIPT UNBLOCKING NOTICE:
    If downloaded from GitHub, unblock before execution:
    Unblock-File -Path .\Send-M365Mail.ps1
================================================================================
"@
    Write-Host $banner -ForegroundColor Cyan
}

# ==============================================================================
# MAIN EXECUTION ROUTINE
# ==============================================================================
try {
    # 1. Check if run without parameters or -Help requested
    if ($Help -or ($PSBoundParameters.Count -eq 0)) {
        Show-ScriptHelp
        exit 0
    }

    # Initialize Logger
    Initialize-ScriptLogger
    Write-LogMessage -Level Info -Message "Starting Send-M365Mail v$global:ScriptVersion on $($env:COMPUTERNAME) ($($env:USERNAME))..." -Component "Core"

    # 2. Handle Client Configuration Routine (-ConfigureClient)
    if ($ConfigureClient) {
        Write-LogMessage -Level Info -Message "Configuring local client settings..." -Component "Config"
        Save-ClientConfiguration -Tenant $TenantId -Client $ClientId -DefaultSender $From -Thumbprint $CertificateThumbprint
        Write-Host "`n[+] Client configuration successfully updated in $global:ConfigFilePath" -ForegroundColor Green
        exit 0
    }

    # 3. Handle Certificate Generation Routine (-GenerateCertificate)
    if ($GenerateCertificate) {
        New-M365GraphCertificate
        exit 0
    }

    # 4. Validate mandatory parameters for email dispatch
    if ([string]::IsNullOrWhiteSpace($To) -or [string]::IsNullOrWhiteSpace($Subject) -or [string]::IsNullOrWhiteSpace($Body)) {
        Write-LogMessage -Level Error -Message "Missing required parameters (-To, -Subject, -Body). Run with -Help for syntax." -Component "Core"
        Show-ScriptHelp
        exit 1
    }

    # 5. Connect to Microsoft 365 / Microsoft Graph
    $context = Connect-M365Service -Tenant $TenantId -Client $ClientId -Thumbprint $CertificateThumbprint -ForceInteractive:$Interactive

    # 6. Dispatch Email Message
    $dispatchResult = Send-GraphEmailMessage `
        -Recipients $To `
        -MailSubject $Subject `
        -MailBody $Body `
        -Attachments $AttachmentPath `
        -IsHtml $IsHtmlBody `
        -SenderAddress $From `
        -CcRecipients $Cc `
        -BccRecipients $Bcc `
        -Priority $Importance `
        -SaveSent $SaveToSentItems

    $duration = (Get-Date) - $global:ExecutionStart
    $formattedDuration = "{0:mm}m {0:ss}s" -f $duration

    # 7. Comprehensive Execution Summary
    Write-LogMessage -Level Success -Message "Email pipeline completed successfully in $formattedDuration!" -Component "Core"

    $attSummary = if ($AttachmentPath) { "$($dispatchResult.AttachmentCount) file(s) ($([math]::Round($dispatchResult.AttachmentBytes / 1KB, 1)) KB)" } else { "None" }

    $summaryCard = @"

================================================================================
                    SEND-M365MAIL EXECUTION SUMMARY
================================================================================
  Status         : SUCCESS
  Duration       : $formattedDuration
  Sender         : $(if ($From) { $From } else { $global:ConnectedUserUpn })
  Recipient(s)   : $($To -join ', ')
  Subject        : $Subject
  Importance     : $Importance
  Attachments    : $attSummary
  Authentication : $global:AuthMethodUsed
  Log File       : $(if ($global:ScriptLogPath) { $global:ScriptLogPath } else { 'Disabled' })
================================================================================

QUICK INSTRUCTIONS FOR USER:
- The email has been submitted to Microsoft 365 cloud transport.
- You can inspect the sent message in your 'Sent Items' folder in Outlook.
- For scheduled task automation without user prompts, use certificate authentication:
  .\Send-M365Mail.ps1 -To "alert@domain.com" -Subject "Alert" -Body "Text" -CertificateThumbprint "THUMBPRINT"
- Full execution log saved to: $global:ScriptLogPath
================================================================================
"@
    Write-Host $summaryCard -ForegroundColor Green

    if ($global:ScriptLogPath) {
        Add-Content -Path $global:ScriptLogPath -Value $summaryCard -Encoding UTF8 -ErrorAction SilentlyContinue
    }

} catch {
    $duration = (Get-Date) - $global:ExecutionStart
    $formattedDuration = "{0:mm}m {0:ss}s" -f $duration

    Write-LogMessage -Level Error -Message "Critical error during Send-M365Mail execution: $_" -Component "Core"

    $failureCard = @"

================================================================================
                    SEND-M365MAIL EXECUTION FAILED
================================================================================
  Status         : FAILED
  Duration       : $formattedDuration
  Error Details  : $_
  Log File       : $(if ($global:ScriptLogPath) { $global:ScriptLogPath } else { 'Disabled' })
================================================================================
INSTRUCTIONS TO RESOLVE:
1. Ensure the Microsoft.Graph.Authentication module is installed.
2. If using Certificate auth, verify that the application has 'Mail.Send' Application Permission granted with admin consent.
3. If running interactively, confirm you have completed the browser sign-in prompt and approved requested permissions.
4. Review detailed errors in log file: $global:ScriptLogPath
================================================================================
"@
    Write-Host $failureCard -ForegroundColor Red

    if ($global:ScriptLogPath) {
        Add-Content -Path $global:ScriptLogPath -Value $failureCard -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    exit 1
}


