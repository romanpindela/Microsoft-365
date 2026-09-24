<#
.SYNOPSIS
    Grants FullAccess permissions on a target mailbox to a user with optional AutoMapping control.

.DESCRIPTION
    This script connects to Exchange Online (prompting for administrator authentication if needed),
    validates target and user identities, displays current permissions, requests confirmation,
    and applies FullAccess permissions. AutoMapping can be enabled or disabled via parameter.
    Also features a read-only mode to display existing permissions.

.PARAMETER MailboxIdentity
    The identity (UPN, primary email, or alias) of the target mailbox (User or Shared).

.PARAMETER UserIdentity
    The identity (UPN or primary email) of the user receiving FullAccess permissions.
    Required unless -ShowPermissionsOnly is used.

.PARAMETER DisableAutoMapping
    Switch parameter. If specified, disables AutoMapping (-AutoMapping $false), preventing
    Outlook from automatically loading the mailbox via Autodiscover.

.PARAMETER ShowPermissionsOnly
    Switch parameter. Displays the current mailbox permissions and exits without modifying anything.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Add-MailboxPermissionToUser.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "admin@domain.com" -DisableAutoMapping
    Grants FullAccess to shared@domain.com for admin@domain.com with AutoMapping disabled.

.EXAMPLE
    .\Add-MailboxPermissionToUser.ps1 -MailboxIdentity "shared@domain.com" -ShowPermissionsOnly
    Displays explicit permissions assigned to shared@domain.com.

.EXAMPLE
    .\Add-MailboxPermissionToUser.ps1 -h
    Displays detailed script documentation.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 2.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'GrantAccess')]
param(
    [Parameter(ParameterSetName = 'GrantAccess', Mandatory = $false, Position = 0)]
    [Parameter(ParameterSetName = 'ViewPermissions', Mandatory = $false, Position = 0)]
    [string]$MailboxIdentity,

    [Parameter(ParameterSetName = 'GrantAccess', Mandatory = $false, Position = 1)]
    [string]$UserIdentity,

    [Parameter(ParameterSetName = 'GrantAccess')]
    [switch]$DisableAutoMapping,

    [Parameter(ParameterSetName = 'ViewPermissions')]
    [switch]$ShowPermissionsOnly,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-ScriptHelp {
    Write-Host @"
================================================================================
SCRIPT: Add-MailboxPermissionToUser.ps1
VERSION: 2.0.0
AUTHOR: Roman Pindela
CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela
================================================================================

DESCRIPTION:
    Grants FullAccess rights to a target mailbox (User or Shared).
    Allows toggling AutoMapping on or off, and viewing mailbox permissions.

AUTHENTICATION:
    Requires the 'ExchangeOnlineManagement' PowerShell module. Automatically detects active
    sessions or prompts for Microsoft 365 Exchange Administrator sign-in.

USAGE EXAMPLES:
    # Grant access with AutoMapping disabled:
    .\Add-MailboxPermissionToUser.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -DisableAutoMapping

    # Grant access with standard AutoMapping enabled:
    .\Add-MailboxPermissionToUser.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com"

    # View existing permissions only:
    .\Add-MailboxPermissionToUser.ps1 -MailboxIdentity "shared@domain.com" -ShowPermissionsOnly

PARAMETERS:
    -MailboxIdentity       The email address/UPN/alias of the target mailbox.
    -UserIdentity          The email address/UPN of the user receiving access.
    -DisableAutoMapping    Switch to disable Outlook Autodiscover automapping.
    -ShowPermissionsOnly   Switch to display current permissions without making changes.
    -Help, -h              Display this help documentation.

SECURITY & UNBLOCKING:
    After downloading from GitHub, remember to unblock the script before execution:
    Unblock-File -Path .\Add-MailboxPermissionToUser.ps1
================================================================================
"@ -ForegroundColor Cyan
}

# Display help if requested or if no arguments provided
if ($Help -or [string]::IsNullOrWhiteSpace($MailboxIdentity)) {
    Show-ScriptHelp
    exit 0
}

if (-not $ShowPermissionsOnly -and [string]::IsNullOrWhiteSpace($UserIdentity)) {
    Write-Error "Missing parameter: -UserIdentity is required when granting permissions. Run with -h for help."
    exit 1
}

# Input Sanitization and Validation
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
if ($MailboxIdentity -notmatch $emailPattern -and $MailboxIdentity -notmatch "^[a-zA-Z0-9._-]+$") {
    Write-Error "Invalid format for -MailboxIdentity. Provide a valid email address, UPN, or alias."
    exit 1
}

if (-not $ShowPermissionsOnly -and $UserIdentity -notmatch $emailPattern) {
    Write-Error "Invalid format for -UserIdentity. Provide a valid internal email address or UPN."
    exit 1
}

# 1. Verify ExchangeOnlineManagement module is available
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Error "ExchangeOnlineManagement module is not installed.`nRun: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
    exit 1
}

# 2. Connect to Exchange Online if not already connected
try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[+] Active Exchange Online session detected." -ForegroundColor Green
} catch {
    Write-Host "[*] No active Exchange Online session detected. Initiating administrator login..." -ForegroundColor Cyan
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

# 3. Verify that the target mailbox exists
Write-Host "`n[*] Validating target mailbox: $MailboxIdentity..." -ForegroundColor Cyan
try {
    $targetMailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
    Write-Host "    Found mailbox: $($targetMailbox.DisplayName) ($($targetMailbox.RecipientTypeDetails))" -ForegroundColor Gray
} catch {
    Write-Error "Unable to locate mailbox '$MailboxIdentity'. Verify identity and tenant connectivity. Details: $_"
    exit 1
}

# Function: Display current permissions
function Get-MailboxCurrentPermissions {
    param([string]$Identity)
    Write-Host "`n[*] Current permissions for mailbox '$Identity':" -ForegroundColor Cyan
    $perms = Get-MailboxPermission -Identity $Identity | 
        Where-Object { -not $_.IsInherited -and$_.User -notmatch "^NT AUTHORITY" } | 
        Select-Object User, AccessRights, Deny

    if ($perms) {$perms | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Yellow
    } else {
        Write-Host "    No explicit non-inherited permissions assigned." -ForegroundColor Gray
    }
}

# 4. Read-Only Mode Execution
if ($ShowPermissionsOnly) {
    Get-MailboxCurrentPermissions -Identity $MailboxIdentity
    exit 0
}

# 5. Verify user identity exists in directory
Write-Host "`n[*] Validating user identity: $UserIdentity..." -ForegroundColor Cyan
try {
    $targetUser = Get-Recipient -Identity $UserIdentity -ErrorAction Stop
    Write-Host "    Found recipient: $($targetUser.DisplayName) ($($targetUser.RecipientTypeDetails))" -ForegroundColor Gray
} catch {
    Write-Error "User '$UserIdentity' was not found in Exchange Online. Ensure this is an internal mailbox/account in your tenant."
    exit 1
}

# Display Current Permissions before modification
Get-MailboxCurrentPermissions -Identity $MailboxIdentity

# 6. Confirmation Prompt
$autoMappingSetting = if ($DisableAutoMapping) { "FALSE (Disabled)" } else { "TRUE (Enabled)" }

Write-Host "`nAction Summary:" -ForegroundColor White
Write-Host "  Target Mailbox : $MailboxIdentity" -ForegroundColor White
Write-Host "  Grant To User  : $UserIdentity" -ForegroundColor White
Write-Host "  Right to Grant : FullAccess" -ForegroundColor White
Write-Host "  AutoMapping    : $autoMappingSetting" -ForegroundColor White
Write-Host ""

$confirmation = Read-Host "Are you sure you want to apply these permissions? (Y/N)"
if ($confirmation -notmatch "^(Y|y|T|t)$") {
    Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
    exit 0
}

# 7. Apply permission changes
try {
    Write-Host "`n[*] Checking and clearing existing FullAccess permission for '$UserIdentity'..." -ForegroundColor Cyan
    $existingAccess = Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited }

    if ($existingAccess) {
        Remove-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        Write-Host "    Existing FullAccess permission removed." -ForegroundColor Gray
    }

    $autoMappingParam = -not $DisableAutoMapping

    Write-Host "[*] Applying FullAccess (AutoMapping = $autoMappingParam)..." -ForegroundColor Cyan
    Add-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -AutoMapping $autoMappingParam -Confirm:$false -ErrorAction Stop

    Write-Host "`n[SUCCESS] Successfully assigned FullAccess to '$UserIdentity'!" -ForegroundColor Green

    # Display Updated Permissions
    Get-MailboxCurrentPermissions -Identity $MailboxIdentity

    # Output Guidance on Access
    if ($DisableAutoMapping) {
        Write-Host @"
================================================================================
HOW TO OPEN THIS MAILBOX (AutoMapping is DISABLED):
================================================================================
1. Outlook on the Web (OWA):
   - Sign in to https://outlook.office.com
   - Click your profile photo/initials in the top right corner.
   - Select 'Open another mailbox...'.
   - Type '$MailboxIdentity' and click Open.

2. Outlook Desktop (Manual Addition):
   - Open Outlook -> File -> Account Settings -> Account Settings.
   - Select your primary account -> Click 'Change'.
   - Click 'More Settings' -> 'Advanced' tab.
   - Under 'Open these additional mailboxes', click 'Add...' and enter '$MailboxIdentity'.
   - Click OK -> Next -> Finish.
================================================================================
"@ -ForegroundColor Yellow
    } else {
        Write-Host @"
================================================================================
AUTOMAPPING IS ENABLED:
The mailbox '$MailboxIdentity' will automatically map in Outlook Desktop within
15-60 minutes once Autodiscover updates.
================================================================================
"@ -ForegroundColor Cyan
    }

} catch {
    Write-Error "An error occurred while modifying permissions: $_"
    exit 1
}