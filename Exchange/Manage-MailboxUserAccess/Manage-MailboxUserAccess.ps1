<#
.SYNOPSIS
    Manages FullAccess and SendAs permissions (grant or revoke) on a target mailbox with optional AutoMapping and Force controls.

.DESCRIPTION
    This script connects to Exchange Online (prompting for administrator authentication if needed),
    validates target mailbox and user identities, displays current permissions, and manages FullAccess and SendAs access.
    Supports granting access (with optional SendAs permission and AutoMapping suppression), revoking access, running in read-only
    audit mode, and bypassing confirmation prompts via -Force.

.PARAMETER MailboxIdentity
    The identity (UPN, primary email, or alias) of the target mailbox (User or Shared).

.PARAMETER UserIdentity
    The identity (UPN or primary email) of the user receiving or losing FullAccess permissions.
    Required unless -ShowPermissionsOnly is used.

.PARAMETER SendAs
    Switch parameter (alias: -GrantSendAs). When granting access, assigns 'SendAs' recipient permission
    in addition to FullAccess, allowing the user to send or forward emails directly from the mailbox.
    When revoking access (-RemoveAccess), also revokes the 'SendAs' permission.

.PARAMETER RemoveAccess
    Switch parameter. If specified, revokes FullAccess (and SendAs if specified) permissions from the specified user.

.PARAMETER DisableAutoMapping
    Switch parameter. When granting access, disables AutoMapping (-AutoMapping $false) to prevent
    Outlook from automatically mounting the mailbox via Autodiscover.

.PARAMETER Force
    Switch parameter. Bypasses the interactive confirmation prompt (Y/N) for automated executions.

.PARAMETER ShowPermissionsOnly
    Switch parameter. Displays current mailbox permissions and exits without making changes.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -DisableAutoMapping
    Grants FullAccess to shared@domain.com for user@domain.com with AutoMapping disabled.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -SendAs
    Grants FullAccess and SendAs to shared@domain.com for user@domain.com.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -DisableAutoMapping -SendAs
    Grants FullAccess and SendAs to shared@domain.com for user@domain.com with AutoMapping disabled.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -RemoveAccess
    Revokes FullAccess from user@domain.com.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -UserIdentity "user@domain.com" -RemoveAccess -SendAs -Force
    Revokes FullAccess and SendAs silently without confirmation.

.EXAMPLE
    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity "shared@domain.com" -ShowPermissionsOnly
    Displays explicit permissions assigned to the mailbox.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 2.2.0
#>

[CmdletBinding(DefaultParameterSetName = 'ManageAccess')]
param(
    [Parameter(ParameterSetName = 'ManageAccess', Mandatory = $false, Position = 0)]
    [Parameter(ParameterSetName = 'ViewPermissions', Mandatory = $false, Position = 0)]
    [string]$MailboxIdentity,

    [Parameter(ParameterSetName = 'ManageAccess', Mandatory = $false, Position = 1)]
    [string]$UserIdentity,

    [Parameter(ParameterSetName = 'ManageAccess')]
    [Alias('GrantSendAs')]
    [switch]$SendAs,

    [Parameter(ParameterSetName = 'ManageAccess')]
    [switch]$RemoveAccess,

    [Parameter(ParameterSetName = 'ManageAccess')]
    [switch]$DisableAutoMapping,

    [Parameter(ParameterSetName = 'ManageAccess')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'ViewPermissions')]
    [switch]$ShowPermissionsOnly,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: Manage-MailboxUserAccess.ps1",
        "VERSION: 2.2.0",
        "AUTHOR: Roman Pindela",
        "CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Manages FullAccess and SendAs rights on a target mailbox (User or Shared).",
        "    Supports granting access (with optional SendAs and AutoMapping suppression), revoking access,",
        "    and auditing current mailbox permissions.",
        "",
        "AUTHENTICATION:",
        "    Requires the 'ExchangeOnlineManagement' PowerShell module. Automatically detects active",
        "    sessions or prompts for Microsoft 365 Exchange Administrator sign-in.",
        "",
        "USAGE EXAMPLES:",
        "    # Grant FullAccess with confirmation and AutoMapping disabled:",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -UserIdentity `"user@domain.com`" -DisableAutoMapping",
        "",
        "    # Grant FullAccess and SendAs (allows sending/forwarding as the mailbox):",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -UserIdentity `"user@domain.com`" -SendAs",
        "",
        "    # Grant FullAccess and SendAs silently without confirmation:",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -UserIdentity `"user@domain.com`" -SendAs -Force",
        "",
        "    # Revoke FullAccess with confirmation:",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -UserIdentity `"user@domain.com`" -RemoveAccess",
        "",
        "    # Revoke FullAccess and SendAs silently:",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -UserIdentity `"user@domain.com`" -RemoveAccess -SendAs -Force",
        "",
        "    # View existing permissions only:",
        "    .\Manage-MailboxUserAccess.ps1 -MailboxIdentity `"shared@domain.com`" -ShowPermissionsOnly",
        "",
        "PARAMETERS:",
        "    -MailboxIdentity       The email address/UPN/alias of the target mailbox.",
        "    -UserIdentity          The email address/UPN of the user to grant or revoke access.",
        "    -SendAs, -GrantSendAs  Switch to grant (or revoke) SendAs permission alongside FullAccess.",
        "    -RemoveAccess          Switch to revoke FullAccess (and SendAs if specified) permissions.",
        "    -DisableAutoMapping    Switch to disable Outlook Autodiscover automapping when granting access.",
        "    -Force                 Switch to bypass confirmation prompts (Y/N).",
        "    -ShowPermissionsOnly   Switch to display current permissions without making changes.",
        "    -Help, -h              Display this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Manage-MailboxUserAccess.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

if ($Help -or [string]::IsNullOrWhiteSpace($MailboxIdentity)) {
    Show-ScriptHelp
    exit 0
}

if (-not $ShowPermissionsOnly -and [string]::IsNullOrWhiteSpace($UserIdentity)) {
    Write-Error "Missing parameter: -UserIdentity is required when managing permissions. Run with -h for help."
    exit 1
}

$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
if ($MailboxIdentity -notmatch $emailPattern -and $MailboxIdentity -notmatch "^[a-zA-Z0-9._-]+$") {
    Write-Error "Invalid format for -MailboxIdentity. Provide a valid email address, UPN, or alias."
    exit 1
}

if (-not $ShowPermissionsOnly -and $UserIdentity -notmatch $emailPattern) {
    Write-Error "Invalid format for -UserIdentity. Provide a valid internal email address or UPN."
    exit 1
}

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Error "ExchangeOnlineManagement module is not installed.`nRun: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
    exit 1
}

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

Write-Host "`n[*] Validating target mailbox: $MailboxIdentity..." -ForegroundColor Cyan
try {
    $targetMailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
    Write-Host "    Found mailbox: $($targetMailbox.DisplayName) ($($targetMailbox.RecipientTypeDetails))" -ForegroundColor Gray
} catch {
    Write-Error "Unable to locate mailbox '$MailboxIdentity'. Details: $_"
    exit 1
}

function Get-MailboxCurrentPermissions {
    param([string]$Identity)
    Write-Host "`n[*] Current permissions for mailbox '$Identity':" -ForegroundColor Cyan
    Write-Host "  -- FullAccess Permissions --" -ForegroundColor DarkCyan
    $perms = Get-MailboxPermission -Identity $Identity -ErrorAction SilentlyContinue | 
        Where-Object { -not $_.IsInherited -and $_.User -notmatch "^NT AUTHORITY" } | 
        Select-Object User, AccessRights, Deny

    if ($perms) {
        $perms | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Yellow
    } else {
        Write-Host "    No explicit non-inherited FullAccess permissions assigned." -ForegroundColor Gray
    }

    Write-Host "  -- SendAs Permissions --" -ForegroundColor DarkCyan
    $sendAsPerms = Get-RecipientPermission -Identity $Identity -ErrorAction SilentlyContinue | 
        Where-Object { -not $_.IsInherited -and $_.Trustee -notmatch "^NT AUTHORITY" } | 
        Select-Object Trustee, AccessRights

    if ($sendAsPerms) {
        $sendAsPerms | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Yellow
    } else {
        Write-Host "    No explicit non-inherited SendAs permissions assigned." -ForegroundColor Gray
    }
}

if ($ShowPermissionsOnly) {
    Get-MailboxCurrentPermissions -Identity $MailboxIdentity
    exit 0
}

Write-Host "`n[*] Validating user identity: $UserIdentity..." -ForegroundColor Cyan
try {
    $targetUser = Get-Recipient -Identity $UserIdentity -ErrorAction Stop
    Write-Host "    Found recipient: $($targetUser.DisplayName) ($($targetUser.RecipientTypeDetails))" -ForegroundColor Gray
} catch {
    Write-Error "User '$UserIdentity' was not found in Exchange Online. Ensure this is an internal mailbox/account in your tenant."
    exit 1
}

Get-MailboxCurrentPermissions -Identity $MailboxIdentity

$actionType = if ($RemoveAccess) {
    if ($SendAs) { "REVOKE (Remove FullAccess and SendAs)" } else { "REVOKE (Remove FullAccess)" }
} else {
    if ($SendAs) { "GRANT (Assign FullAccess and SendAs)" } else { "GRANT (Assign FullAccess)" }
}
$autoMappingSetting = if ($RemoveAccess) { "N/A" } elseif ($DisableAutoMapping) { "FALSE (Disabled)" } else { "TRUE (Enabled)" }
$sendAsSetting = if ($SendAs) { "YES (SendAs included)" } else { "NO (FullAccess only)" }

Write-Host "`nAction Summary:" -ForegroundColor White
Write-Host "  Action Type    : $actionType" -ForegroundColor White
Write-Host "  Target Mailbox : $MailboxIdentity" -ForegroundColor White
Write-Host "  User Identity  : $UserIdentity" -ForegroundColor White
Write-Host "  SendAs Rights  : $sendAsSetting" -ForegroundColor White
if (-not $RemoveAccess) {
    Write-Host "  AutoMapping    : $autoMappingSetting" -ForegroundColor White
}
Write-Host ""

if (-not $Force) {
    $promptQuestion = if ($RemoveAccess) {
        if ($SendAs) {
            "Are you sure you want to REVOKE FullAccess and SendAs permissions for this user? (Y/N)"
        } else {
            "Are you sure you want to REVOKE permissions for this user? (Y/N)"
        }
    } else {
        if ($SendAs) {
            "Are you sure you want to APPLY FullAccess and SendAs permissions? (Y/N)"
        } else {
            "Are you sure you want to APPLY these permissions? (Y/N)"
        }
    }
    
    $confirmation = Read-Host $promptQuestion
    if ($confirmation -notmatch "^(Y|y|T|t)$") {
        Write-Host "[!] Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-Host "[*] -Force switch detected. Proceeding without confirmation prompt." -ForegroundColor Cyan
}

try {
    $existingAccess = Get-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited }

    $existingSendAs = Get-RecipientPermission -Identity $MailboxIdentity -Trustee $UserIdentity -ErrorAction SilentlyContinue |
        Where-Object { $_.AccessRights -contains "SendAs" -and -not $_.IsInherited }

    if ($RemoveAccess) {
        if ($SendAs) {
            if (-not $existingSendAs) {
                Write-Host "[!] User '$UserIdentity' does not have explicit SendAs permission on '$MailboxIdentity'. Nothing to remove for SendAs." -ForegroundColor Yellow
            } else {
                Write-Host "`n[*] Revoking SendAs permission for '$UserIdentity'..." -ForegroundColor Cyan
                Remove-RecipientPermission -Identity $MailboxIdentity -Trustee $UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "`n[SUCCESS] Successfully revoked SendAs from '$UserIdentity'!" -ForegroundColor Green
            }
        }

        if (-not $existingAccess) {
            Write-Host "[!] User '$UserIdentity' does not have explicit FullAccess to '$MailboxIdentity'. Nothing to remove for FullAccess." -ForegroundColor Yellow
        } else {
            Write-Host "`n[*] Revoking FullAccess permission for '$UserIdentity'..." -ForegroundColor Cyan
            Remove-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
            Write-Host "`n[SUCCESS] Successfully revoked FullAccess from '$UserIdentity'!" -ForegroundColor Green
        }

        if (-not $SendAs -and $existingSendAs) {
            Write-Host "[i] Note: User '$UserIdentity' still retains explicit SendAs permission on '$MailboxIdentity'. Run with -SendAs -RemoveAccess to revoke SendAs as well." -ForegroundColor Yellow
        }
    } else {
        if ($existingAccess) {
            Write-Host "`n[*] Clearing existing FullAccess assignment for clean re-application..." -ForegroundColor Cyan
            Remove-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
        }

        $autoMappingParam = -not $DisableAutoMapping
        Write-Host "`n[*] Applying FullAccess (AutoMapping = $autoMappingParam)..." -ForegroundColor Cyan
        Add-MailboxPermission -Identity $MailboxIdentity -User $UserIdentity -AccessRights FullAccess -AutoMapping $autoMappingParam -Confirm:$false -ErrorAction Stop

        Write-Host "`n[SUCCESS] Successfully assigned FullAccess to '$UserIdentity'!" -ForegroundColor Green

        if ($SendAs) {
            if ($existingSendAs) {
                Write-Host "`n[*] Clearing existing SendAs assignment for clean re-application..." -ForegroundColor Cyan
                Remove-RecipientPermission -Identity $MailboxIdentity -Trustee $UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            }

            Write-Host "`n[*] Applying SendAs permission..." -ForegroundColor Cyan
            Add-RecipientPermission -Identity $MailboxIdentity -Trustee $UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop

            Write-Host "`n[SUCCESS] Successfully assigned SendAs to '$UserIdentity'!" -ForegroundColor Green
        }
    }

    Get-MailboxCurrentPermissions -Identity $MailboxIdentity

    if ($RemoveAccess) {
        $revokeMsg = @(
            "================================================================================",
            "ACCESS REVOKED:",
            "If the mailbox was previously mapped in Outlook Desktop, allow 15-60 minutes for",
            "Autodiscover caches to expire or restart Outlook.",
            "================================================================================"
        )
        Write-Host ($revokeMsg -join "`n") -ForegroundColor Yellow
    } else {
        if ($SendAs) {
            $sendAsMsg = @(
                "================================================================================",
                "SEND AS PERMISSION CONFIGURED:",
                "User '$UserIdentity' can now send and forward messages as '$MailboxIdentity'.",
                "In Outlook (Desktop / Web):",
                "1. When composing or forwarding an email, ensure the 'From' field is displayed.",
                "2. Click 'From' -> select 'Other Email Address...' and enter '$MailboxIdentity'.",
                "Note: Address book / token propagation in Exchange Online may take 15-60 minutes.",
                "================================================================================"
            )
            Write-Host ($sendAsMsg -join "`n") -ForegroundColor Green
        }

        if ($DisableAutoMapping) {
            $noAutoMsg = @(
                "================================================================================",
                "HOW TO OPEN THIS MAILBOX (AutoMapping is DISABLED):",
                "================================================================================",
                "1. Outlook on the Web (OWA):",
                "   - Sign in to https://outlook.office.com",
                "   - Click your profile photo/initials in the top right corner.",
                "   - Select 'Open another mailbox...'.",
                "   - Type '$MailboxIdentity' and click Open.",
                "",
                "2. Outlook Desktop (Manual Addition):",
                "   - Open Outlook -> File -> Account Settings -> Account Settings.",
                "   - Select your primary account -> Click 'Change'.",
                "   - Click 'More Settings' -> 'Advanced' tab.",
                "   - Under 'Open these additional mailboxes', click 'Add...' and enter '$MailboxIdentity'.",
                "   - Click OK -> Next -> Finish.",
                "================================================================================"
            )
            Write-Host ($noAutoMsg -join "`n") -ForegroundColor Yellow
        } else {
            $autoMsg = @(
                "================================================================================",
                "AUTOMAPPING IS ENABLED:",
                "The mailbox '$MailboxIdentity' will automatically map in Outlook Desktop within",
                "15-60 minutes once Autodiscover updates.",
                "================================================================================"
            )
            Write-Host ($autoMsg -join "`n") -ForegroundColor Cyan
        }
    }

} catch {
    Write-Error "An error occurred while modifying permissions: $_"
    exit 1
}