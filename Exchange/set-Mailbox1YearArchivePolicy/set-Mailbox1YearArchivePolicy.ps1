<#
.SYNOPSIS
    Enables Online Archive (if disabled) and applies a 1-year archive retention policy to an Exchange Online mailbox.

.DESCRIPTION
    This script automates the complete lifecycle of enabling online archiving and enforcing a 1-year retention policy
    for Exchange Online mailboxes (User Mailboxes or Shared Mailboxes).
    
    Key Actions Performed:
      1. Verifies and installs the ExchangeOnlineManagement module if missing.
      2. Authenticates to Exchange Online securely via interactive modern auth (MFA supported) if not connected.
      3. Validates the existence of the targeted mailbox.
      4. AUTOMATICALLY ENABLES ONLINE ARCHIVE (Enable-Mailbox -Archive) if the mailbox archive status is not 'Active'.
      5. Idempotently creates the 'Default 1 Year Move to Archive' retention tag (365 days, MoveToArchive) if missing.
      6. Idempotently creates the 'Shared Mailbox 1 Year Archive' retention policy if missing.
      7. Safely assigns the retention policy to the targeted mailbox without altering existing permissions or deleting items.
      8. Triggers the Managed Folder Assistant to initiate background migration to the archive.

.PARAMETER Identity
    The primary SMTP address, UPN, or alias of the mailbox to process.

.PARAMETER Help
    Displays syntax instructions, parameters, prerequisites, and author information.

.EXAMPLE
    .\set-Mailbox1YearArchivePolicy.ps1 -Identity "jurnaling@elektrimont.pl"
    Enables Online Archive (if needed), assigns the 1-year retention policy, and triggers the assistant.

.EXAMPLE
    .\set-Mailbox1YearArchivePolicy.ps1 -Help
    Displays help details and usage examples.

.NOTES
    Author  : Roman Pindela
    Email   : roman.pindela@gmail.com
    GitHub  : https://github.com/romanpindela
    Version : 1.1.0
#>

[CmdletBinding(DefaultParameterSetName = 'Execute')]
param(
    [Parameter(ParameterSetName = 'Execute', Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,

    [Parameter(ParameterSetName = 'Help', Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

# ----------------------------------------------------------------------
# Helper: Show Help & Exit
# ----------------------------------------------------------------------
function Show-ScriptHelp {
    Write-Host ""
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host " set-Mailbox1YearArchivePolicy.ps1 - Version 1.1.0" -ForegroundColor Cyan
    Write-Host " Author: Roman Pindela (roman.pindela@gmail.com | github.com/romanpindela)" -ForegroundColor Gray
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Automates Exchange Online archiving configuration. It automatically:"
    Write-Host "    - Checks if Online Archive is active; if disabled, ENABLES IT AUTOMATICALLY."
    Write-Host "    - Creates/verifies a 1-year (365-day) 'Move to Archive' retention policy."
    Write-Host "    - Assigns the policy to the target mailbox and invokes the assistant."
    Write-Host ""
    Write-Host "SYNTAX:" -ForegroundColor Yellow
    Write-Host "  .\set-Mailbox1YearArchivePolicy.ps1 -Identity <MailboxAddress>"
    Write-Host "  .\set-Mailbox1YearArchivePolicy.ps1 -Help | -h"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\set-Mailbox1YearArchivePolicy.ps1 -Identity `"jurnaling@elektrimont.pl`""
    Write-Host "  .\set-Mailbox1YearArchivePolicy.ps1 -h"
    Write-Host ""
    Write-Host "PREREQUISITES & AUTHENTICATION:" -ForegroundColor Yellow
    Write-Host "  - Windows PowerShell 5.1 or PowerShell 7+"
    Write-Host "  - Automatically verifies active Exchange Online connection; if not connected,"
    Write-Host "    prompts for Administrator credentials (OAuth2 / MFA supported)."
    Write-Host "  - Automatically installs ExchangeOnlineManagement module if missing."
    Write-Host "  - Exchange Administrator or Global Administrator privileges required."
    Write-Host ""
}

# Show help if invoked with -Help/-h or if Identity parameter is missing
if ($Help.IsPresent -or [string]::IsNullOrWhiteSpace($Identity)) {
    Show-ScriptHelp
    Exit 0
}

# ----------------------------------------------------------------------
# Input Sanitation & Protection
# ----------------------------------------------------------------------
$Identity = $Identity.Trim().Trim('"').Trim("'")

if ($Identity -match '[\$`;&|><\r\n]') {
    Write-Error "Invalid mailbox identifier format detected. Execution aborted."
    Exit 1
}

$emailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
if ($Identity -notmatch$emailRegex) {
    Write-Warning "Identity '$Identity' is not in standard UPN/email format. Checking alias..."
}

# ----------------------------------------------------------------------
# Dependency Verification: ExchangeOnlineManagement
# ----------------------------------------------------------------------
$moduleName = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Host "[$moduleName] module is missing. Attempting installation..." -ForegroundColor Yellow
    try {
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[$moduleName] installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install [$moduleName]. Error: $_"
        Exit 1
    }
}

Import-Module $moduleName -ErrorAction SilentlyContinue

# ----------------------------------------------------------------------
# Microsoft 365 Authentication
# ----------------------------------------------------------------------
try {
    $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $existingConnection) {
        Write-Host "Connecting to Exchange Online (interactive administrator sign-in)..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Successfully authenticated and connected to Exchange Online." -ForegroundColor Green
    }
    else {
        Write-Host "Active Exchange Online session detected." -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to establish connection with Exchange Online: $_"
    Exit 1
}

# ----------------------------------------------------------------------
# Core Execution Logic & Safety Checks
# ----------------------------------------------------------------------
$tagName    = "Default 1 Year Move to Archive"
$policyName = "Shared Mailbox 1 Year Archive"

try {
    # 1. Verify Target Mailbox
    Write-Host "Verifying mailbox: $Identity..." -ForegroundColor Cyan
    $mailbox = Get-Mailbox -Identity$Identity -ErrorAction Stop
    Write-Host "Mailbox identified: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))" -ForegroundColor Green

    # 2. Automated Online Archive Check and Enablement
    if ($mailbox.ArchiveStatus -ne 'Active') {
        Write-Host "Online Archive is currently '$($mailbox.ArchiveStatus)'. Enabling Online Archive automatically..." -ForegroundColor Cyan
        Enable-Mailbox -Identity $mailbox.Identity -Archive -ErrorAction Stop
        Write-Host "Online Archive enabled successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Online Archive is already Active." -ForegroundColor DarkGray
    }

    # 3. Check / Create Retention Policy Tag (MoveToArchive only - Non-destructive)
    $existingTag = Get-RetentionPolicyTag -Identity$tagName -ErrorAction SilentlyContinue
    if (-not $existingTag) {
        Write-Host "Creating retention tag: '$tagName' (365 days)..." -ForegroundColor Cyan
        $null = New-RetentionPolicyTag -Name$tagName `
                                       -Type All `
                                       -RetentionAction MoveToArchive `
                                       -AgeLimitForRetention 365 `
                                       -ErrorAction Stop
        Write-Host "Retention tag created successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Retention tag '$tagName' already exists. Skipping creation." -ForegroundColor DarkGray
    }

    # 4. Check / Create Retention Policy
    $existingPolicy = Get-RetentionPolicy -Identity$policyName -ErrorAction SilentlyContinue
    if (-not $existingPolicy) {
        Write-Host "Creating retention policy: '$policyName'..." -ForegroundColor Cyan
        $null = New-RetentionPolicy -Name$policyName `
                                    -RetentionPolicyTagLinks $tagName, "Junk Email" `
                                    -ErrorAction Stop
        Write-Host "Retention policy created successfully." -ForegroundColor Green
    }
    else {
        Write-Host "Retention policy '$policyName' already exists. Skipping creation." -ForegroundColor DarkGray
    }

    # 5. Safely Assign Policy to Target Mailbox
    Write-Host "Assigning policy '$policyName' to '$($mailbox.PrimarySmtpAddress)'..." -ForegroundColor Cyan
    Set-Mailbox -Identity $mailbox.Identity -RetentionPolicy$policyName -ErrorAction Stop
    Write-Host "Retention policy assigned successfully." -ForegroundColor Green

    # 6. Trigger Managed Folder Assistant
    Write-Host "Initiating Managed Folder Assistant for '$($mailbox.PrimarySmtpAddress)'..." -ForegroundColor Cyan
    try {
        Start-ManagedFolderAssistant -Identity $mailbox.Identity -ErrorAction Stop
        Write-Host "Managed Folder Assistant triggered successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Managed Folder Assistant could not be triggered immediately (cloud replication pending)."
        Write-Warning "Exchange Online will automatically process this mailbox during the next scheduled cycle."
    }

    Write-Host ""
    Write-Host "Operation completed successfully for $($mailbox.PrimarySmtpAddress)." -ForegroundColor Green
}
catch {
    Write-Error "A terminating error occurred during processing: $_"
    Exit 1
}