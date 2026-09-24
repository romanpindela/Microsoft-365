<#
.SYNOPSIS
    Manages Microsoft 365 (Exchange/Entra ID) user accounts: password reset, sign-in blocking/unblocking, session revocation, and comprehensive account auditing.

.DESCRIPTION
    This script connects to Microsoft Graph / M365 (prompting for administrator authentication if needed),
    validates target user identity, audits detailed account status, and performs administrative actions:
    - Detailed view (-View) of password change date, email aliases, friendly license plan names,
      recent client applications used (Outlook Desktop, Web, Mobile), and sign-in activity
    - Secure password reset via SecureString or cryptographically secure auto-generation (enforcing change at next sign-in)
    - Blocking user sign-in (AccountEnabled = $false) with automatic session revocation (instant logout from Outlook, Teams, OWA)
    - Unblocking user sign-in (AccountEnabled = $true)
    - Standalone session revocation (-RevokeSessions)
    Includes input validation, confirmation prompts, SecureString handling, tenant-license fault tolerance,
    and an automated execution switch (-Force).

.PARAMETER UserIdentity
    The identity (UPN or primary SMTP email) of the target Microsoft 365 user.

.PARAMETER View
    Switch parameter. Displays detailed user audit information (sign-in state, last password change,
    email aliases, assigned license plans, detected Outlook/M365 client applications) and exits. Aliases: -v, -ShowStatusOnly.

.PARAMETER ResetPassword
    Switch parameter. Initiates a password reset for the specified account.

.PARAMETER NewPassword
    Optional SecureString password. If omitted when -ResetPassword is used, the administrator will be
    prompted interactively to either generate a strong random password or type one masked in console.

.PARAMETER BlockSignIn
    Switch parameter. Disables user account sign-in (AccountEnabled = $false) and immediately revokes all active sessions.

.PARAMETER UnblockSignIn
    Switch parameter. Enables user account sign-in (AccountEnabled = $true).

.PARAMETER RevokeSessions
    Switch parameter. Revokes all active refresh tokens/sessions without modifying account sign-in status.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N) for unattended automation.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Manage-M365UserAccount.ps1 -UserIdentity "john.doe@domain.com" -View
    Displays detailed audit information (last password change, aliases, license names, detected Outlook clients).

.EXAMPLE
    .\Manage-M365UserAccount.ps1 -UserIdentity "john.doe@domain.com" -BlockSignIn
    Blocks sign-in access and automatically terminates all active sessions across Outlook and Teams.

.EXAMPLE
    .\Manage-M365UserAccount.ps1 -UserIdentity "john.doe@domain.com" -RevokeSessions -Force
    Immediately revokes all active sessions silently without changing account enablement.

.EXAMPLE
    .\Manage-M365UserAccount.ps1 -UserIdentity "john.doe@domain.com" -UnblockSignIn -Force
    Unblocks sign-in silently without confirmation prompt.

.EXAMPLE
    .\Manage-M365UserAccount.ps1 -UserIdentity "john.doe@domain.com" -ResetPassword
    Prompts interactively to auto-generate a random password or input a masked secure password.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.4.2
#>

[CmdletBinding(DefaultParameterSetName = 'ManageAccount')]
param(
    [Parameter(ParameterSetName = 'ManageAccount', Mandatory = $false, Position = 0)]
    [Parameter(ParameterSetName = 'ViewAccount', Mandatory = $false, Position = 0)]     [string]$UserIdentity,

    [Parameter(ParameterSetName = 'ViewAccount')]
    [Alias('v', 'ShowStatusOnly')]
    [switch]$View,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [switch]$ResetPassword,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [System.Security.SecureString]$NewPassword,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [switch]$BlockSignIn,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [switch]$UnblockSignIn,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [switch]$RevokeSessions,

    [Parameter(ParameterSetName = 'ManageAccount')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: Manage-M365UserAccount.ps1",
        "VERSION: 1.4.2",
        "AUTHOR: Roman Pindela",
        "CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Manages and audits Microsoft 365 / Entra ID / Exchange user accounts.",
        "    Supports viewing account details (aliases, friendly licenses, client apps used,",
        "    last password change, sign-in activity), resetting passwords securely,",
        "    toggling sign-in enablement, and revoking active user sessions.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    Requires 'Microsoft.Graph.Users' and 'Microsoft.Graph.Authentication'.",
        "    Required Microsoft Graph scopes: 'User.ReadWrite.All', 'AuditLog.Read.All', 'Organization.Read.All'.",
        "",
        "USAGE EXAMPLES:",
        "    # View detailed account info (aliases, licenses, client apps, last login):",
        "    .\Manage-M365UserAccount.ps1 -UserIdentity `"user@domain.com`" -View",
        "",
        "    # Block sign-in and instantly terminate all active sessions (Outlook/Teams):",
        "    .\Manage-M365UserAccount.ps1 -UserIdentity `"user@domain.com`" -BlockSignIn",
        "",
        "    # Revoke sessions only (without blocking):",
        "    .\Manage-M365UserAccount.ps1 -UserIdentity `"user@domain.com`" -RevokeSessions",
        "",
        "    # Unblock sign-in silently without confirmation:",
        "    .\Manage-M365UserAccount.ps1 -UserIdentity `"user@domain.com`" -UnblockSignIn -Force",
        "",
        "    # Reset password with interactive prompt (auto-generate vs secure masked input):",
        "    .\Manage-M365UserAccount.ps1 -UserIdentity `"user@domain.com`" -ResetPassword",
        "",
        "PARAMETERS:",
        "    -UserIdentity     The email address/UPN of the target user.",
        "    -View, -v         Switch to display detailed audit info (aliases, licenses, clients, etc.).",
        "    -ResetPassword    Switch to trigger a password reset (forces change at next login).",
        "    -NewPassword      Optional SecureString password object.",
        "    -BlockSignIn      Switch to block user sign-in and automatically revoke active sessions.",
        "    -UnblockSignIn    Switch to unblock user sign-in.",
        "    -RevokeSessions   Switch to invalidate active sessions without modifying account status.",
        "    -Force            Switch to bypass confirmation prompts (Y/N).",
        "    -Help, -h         Display this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Manage-M365UserAccount.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

if ($Help -or [string]::IsNullOrWhiteSpace($UserIdentity)) {
    Show-ScriptHelp
    exit 0
}

# Input sanitation
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
if ($UserIdentity -notmatch $emailPattern) {
    Write-Error "Invalid format for -UserIdentity. Provide a valid email address or UPN."
    exit 1
}

if ($BlockSignIn -and $UnblockSignIn) {
    Write-Error "Conflicting parameters: You cannot specify both -BlockSignIn and -UnblockSignIn at the same time."
    exit 1
}

if (-not $View -and -not $ResetPassword -and -not $BlockSignIn -and -not $UnblockSignIn -and -not $RevokeSessions) {
    Write-Error "No action specified. Provide an action flag (-View, -ResetPassword, -BlockSignIn, -UnblockSignIn, -RevokeSessions)."
    exit 1
}

# Module prerequisites check
$missingModules = @()
if (-not (Get-Module -Name Microsoft.Graph.Users -ListAvailable)) { $missingModules += "Microsoft.Graph.Users" }
if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) { $missingModules += "Microsoft.Graph.Authentication" }

if ($missingModules.Count -gt 0) {
    Write-Host "[!] Missing required module(s): $($missingModules -join ', ')" -ForegroundColor Yellow
    $installPrompt = if ($Force) { "Y" } else { Read-Host "[?] Do you want to install missing modules now from PSGallery? (Y/N)" }
    if ($installPrompt -match "^(Y|y|T|t)$") {
        try {
            Write-Host "[*] Installing missing modules ($($missingModules -join ', '))..." -ForegroundColor Cyan
            Install-Module -Name $missingModules -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            Write-Host "[+] Modules installed successfully!" -ForegroundColor Green
        } catch {
            Write-Error "Failed to install required modules: $_`nInstall manually: Install-Module -Name Microsoft.Graph.Users, Microsoft.Graph.Authentication -Scope CurrentUser"
            exit 1
        }
    } else {
        Write-Error "Execution stopped. Required modules must be installed to continue."
        exit 1
    }
}

# Connection handler
$requiredScopes = @("User.ReadWrite.All", "AuditLog.Read.All", "Organization.Read.All")
try {
    $mgContext = Get-MgContext -ErrorAction Stop
    if ($null -eq$mgContext) { throw "No active context" }
    Write-Host "[+] Active Microsoft Graph session detected." -ForegroundColor Green
} catch {
    Write-Host "[*] No active Microsoft Graph session detected. Initiating administrator login..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        Write-Host "[+] Successfully connected to Microsoft Graph." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Microsoft Graph: $_"
        exit 1
    }
}

# Retrieve base user properties
Write-Host "`n[*] Retrieving user details for: $UserIdentity..." -ForegroundColor Cyan
$baseProperties = @(
    "Id", "DisplayName", "UserPrincipalName", "Mail", "AccountEnabled",
    "LastPasswordChangeDateTime", "CreatedDateTime",
    "UserType", "JobTitle", "Department", "AssignedLicenses", "ProxyAddresses"
)

try {
    $user = Get-MgUser -UserId $UserIdentity -Property $baseProperties -ErrorAction Stop
    Write-Host "    Found user: $($user.DisplayName) (UPN: $($user.UserPrincipalName))" -ForegroundColor Gray
} catch {
    Write-Error "Unable to locate user '$UserIdentity' in tenant. Details: $_"
    exit 1
}

# Query SignInActivity (requires Entra ID P1/P2 license)
$signInActivityData = $null
try {
    $activityQuery = Get-MgUser -UserId $user.Id -Property "SignInActivity" -ErrorAction Stop
    if ($activityQuery.SignInActivity) {
        $signInActivityData = $activityQuery.SignInActivity
    }
} catch {
    # Non-terminating fallback for non-P1/P2 tenants
}

# Query Subscribed Skus (License Names)
$friendlyLicenses = @()
try {
    $subscribedSkus = @(Get-MgSubscribedSku -ErrorAction Stop)
    if ($user.AssignedLicenses) {
        foreach ($lic in @($user.AssignedLicenses)) {
            $matchingSku = $subscribedSkus | Where-Object { $_.SkuId -eq $lic.SkuId } | Select-Object -First 1
            if ($matchingSku -and $matchingSku.SkuPartNumber) {
                $friendlyLicenses += $matchingSku.SkuPartNumber
            } else {
                $friendlyLicenses += $lic.SkuId
            }
        }
    }
} catch {
    if ($user.AssignedLicenses) {
        $friendlyLicenses = @($user.AssignedLicenses | ForEach-Object { $_.SkuId })
    }
}

# Query Recent Sign-in Client Apps (AuditLogs)
$detectedClients = @()
try {
    $signIns = @(Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)'" -Top 20 -ErrorAction Stop)
    if ($signIns.Length -gt 0) {
        $detectedClients = @(
            $signIns | 
            Select-Object -ExpandProperty AppDisplayName -Unique | 
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
    }
} catch {
    # Fallback if AuditLog query is not accessible
}

function Show-DetailedUserAudit {
    param(
        [object]$TargetUser,
        [object]$SignInActivity,
        [array]$LicenseNames,
        [array]$ClientApps
    )

    $statusColor = if ($TargetUser.AccountEnabled) { "Green" } else { "Red" }
    $signInStatusText = if ($TargetUser.AccountEnabled) { "ALLOWED (Enabled)" } else { "BLOCKED (Disabled)" }

    $pwdChanged = if ($TargetUser.LastPasswordChangeDateTime) {
        $TargetUser.LastPasswordChangeDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
    } else {
        "Unknown / Not Available"
    }

    $created = if ($TargetUser.CreatedDateTime) {
        $TargetUser.CreatedDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
    } else {
        "Unknown"
    }

    $lastInteractive = "No sign-in recorded (or missing Entra P1/P2 license)"
    $lastNonInteractive = "No sign-in recorded (or missing Entra P1/P2 license)"

    if ($SignInActivity) {
        if ($SignInActivity.LastSignInDateTime) {
            $lastInteractive = $SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
        }
        if ($SignInActivity.LastNonInteractiveSignInDateTime) {
            $lastNonInteractive = $SignInActivity.LastNonInteractiveSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
        }
    }

    # Extract secondary email aliases (smtp:), ignoring primary (SMTP:)
    $aliases = @()
    if ($TargetUser.ProxyAddresses) {
        $aliases = @(
            $TargetUser.ProxyAddresses | 
            Where-Object { $_ -clike "smtp:*" } | 
            ForEach-Object { $_.Substring(5) }
        )
    }

    $licenseArray = @($LicenseNames)
    $clientArray  = @($ClientApps)

    Write-Host "`n========================= USER AUDIT & DETAILS =========================" -ForegroundColor Cyan
    Write-Host ("  Display Name           : " + $TargetUser.DisplayName) -ForegroundColor White
    Write-Host ("  User Principal Name    : " + $TargetUser.UserPrincipalName) -ForegroundColor White
    Write-Host ("  Primary Email          : " + $(if ($TargetUser.Mail) {$TargetUser.Mail } else { "N/A" })) -ForegroundColor White
    
    # Display Aliases
    if ($aliases.Length -gt 0) {
        Write-Host ("  Email Aliases          : " + ($aliases -join ", ")) -ForegroundColor Yellow
    } else {
        Write-Host "  Email Aliases          : (None assigned)" -ForegroundColor Gray
    }

    Write-Host ("  Object ID              : " + $TargetUser.Id) -ForegroundColor Gray
    Write-Host -NoNewline "  Sign-In Status         : " -ForegroundColor White
    Write-Host $signInStatusText -ForegroundColor$statusColor
    Write-Host ("  Last Password Change   : " + $pwdChanged) -ForegroundColor Yellow
    Write-Host ("  Last Interactive Login : " + $lastInteractive) -ForegroundColor Yellow
    Write-Host ("  Last Non-Interactive   : " + $lastNonInteractive) -ForegroundColor Yellow

    # Display Detected Clients
    if ($clientArray.Length -gt 0) {
        Write-Host ("  Recent Client Apps     : " + ($clientArray -join ", ")) -ForegroundColor Cyan
    } else {
        Write-Host "  Recent Client Apps     : No recent client app activity logged" -ForegroundColor Gray
    }

    # Display Licenses
    if ($licenseArray.Length -gt 0) {
        Write-Host ("  Assigned Licenses      : " + ($licenseArray -join ", ")) -ForegroundColor White
    } else {
        Write-Host "  Assigned Licenses      : (No active licenses)" -ForegroundColor Gray
    }

    Write-Host ("  Account Created Date   : " + $created) -ForegroundColor White
    Write-Host ("  User Type              : " + $(if ($TargetUser.UserType) {$TargetUser.UserType } else { "Member" })) -ForegroundColor White
    Write-Host ("  Department / Title     : " + "$($TargetUser.Department) / $($TargetUser.JobTitle)") -ForegroundColor White
    Write-Host "========================================================================`n" -ForegroundColor Cyan
}

if ($View) {
    Show-DetailedUserAudit -TargetUser $user -SignInActivity $signInActivityData -LicenseNames $friendlyLicenses -ClientApps $detectedClients
    exit 0
}

Show-DetailedUserAudit -TargetUser $user -SignInActivity $signInActivityData -LicenseNames $friendlyLicenses -ClientApps $detectedClients

# Cryptographically strong password generator
function New-RandomPassword {
    param([int]$Length = 16)
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+"
    $rnd = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] $Length
    $rnd.GetBytes($bytes)
    $result = New-Object char[] $Length
    for ($i = 0; $i -lt $Length; $i++) {
        $result[$i] = $chars[$bytes[$i] % $chars.Length]
    }
    return -join $result
}

function Convert-SecureStringToString {
    param([System.Security.SecureString]$Secure)
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

# Handle password preparation securely
$plainPasswordToApply = $null
$isAutoGenerated = $false

if ($ResetPassword) {
    if ($null -ne $NewPassword) {
        $plainPasswordToApply = Convert-SecureStringToString -Secure $NewPassword
    } else {
        Write-Host "[?] Password Reset Options:" -ForegroundColor Cyan
        Write-Host "    [1] Auto-generate strong 16-character random password (Recommended)" -ForegroundColor White
        Write-Host "    [2] Manually type password in console (Masked input)" -ForegroundColor White
        $choice = Read-Host "Select option (1/2) [Default: 1]"
        
        if ($choice -eq "2") {
            $inputSecPass = Read-Host -AsSecureString "Enter new password for $($user.UserPrincipalName)"
            $plainPasswordToApply = Convert-SecureStringToString -Secure $inputSecPass
            if ([string]::IsNullOrWhiteSpace($plainPasswordToApply)) {
                Write-Error "Password cannot be empty or whitespace."
                exit 1
            }
        } else {
            $plainPasswordToApply = New-RandomPassword -Length 16
            $isAutoGenerated = $true
        }
    }
}

# Action Preparation
$plannedActions = @()
if ($BlockSignIn) { 
    $plannedActions += "BLOCK Sign-In" 
    $plannedActions += "REVOKE Active Sessions (Instant logout from Outlook/Teams)" 
}
if ($UnblockSignIn) { $plannedActions += "UNBLOCK Sign-In" }
if ($RevokeSessions -and -not $BlockSignIn) { $plannedActions += "REVOKE Active Sessions" }
if ($ResetPassword) { $plannedActions += "RESET Password (Forces change at next sign-in)" }

Write-Host "Action Summary:" -ForegroundColor White
Write-Host "  Target User : $($user.UserPrincipalName)" -ForegroundColor White
Write-Host "  Actions     : $($plannedActions -join ', ')" -ForegroundColor White
Write-Host ""

if (-not $Force) {
    $confirmation = Read-Host "Are you sure you want to execute these actions on '$($user.UserPrincipalName)'? (Y/N)"
    if ($confirmation -notmatch "^(Y|y|T|t)$") {
        Write-Host "[!] Operation cancelled by administrator." -ForegroundColor Yellow
        exit 0
    }
} else {
    Write-Host "[*] -Force switch detected. Proceeding without confirmation prompt." -ForegroundColor Cyan
}

# Execute modifications
try {
    $updateParams = @{}

    if ($BlockSignIn) {
        $updateParams["AccountEnabled"] = $false
    } elseif ($UnblockSignIn) {
        $updateParams["AccountEnabled"] = $true
    }

    if ($ResetPassword) {
        $passwordProfile = @{
            ForceChangePasswordNextSignIn = $true
            Password = $plainPasswordToApply
        }
        $updateParams["PasswordProfile"] = $passwordProfile
    }

    if ($updateParams.Count -gt 0) {
        Write-Host "`n[*] Applying updates to user account..." -ForegroundColor Cyan
        Update-MgUser -UserId $user.Id @updateParams -ErrorAction Stop
        Write-Host "[SUCCESS] Account parameters updated successfully!" -ForegroundColor Green
    }

    # Revoke sessions if requested or if blocking sign-in
    if ($BlockSignIn -or$RevokeSessions) {
        Write-Host "`n[*] Revoking all active sign-in sessions (terminating Outlook, Teams, OWA access)..." -ForegroundColor Cyan
        try {
            Revoke-MgUserAllRefreshToken -UserId $user.Id -ErrorAction Stop | Out-Null
            Write-Host "[SUCCESS] Successfully revoked active user sessions!" -ForegroundColor Green
        } catch {
            Write-Host "[!] Warning: Unable to revoke active sessions: $_" -ForegroundColor Yellow
        }
    }

    # Fetch updated user status
    $updatedUser = Get-MgUser -UserId $user.Id -Property $baseProperties -ErrorAction Stop
    Show-DetailedUserAudit -TargetUser $updatedUser -SignInActivity $signInActivityData -LicenseNames $friendlyLicenses -ClientApps $detectedClients

    if ($ResetPassword) {
        $pwdNotice = @(
            "================================================================================",
            "PASSWORD RESET COMPLETE:",
            "  User:     $($updatedUser.UserPrincipalName)"
        )
        if ($isAutoGenerated) {
            $pwdNotice += "  Password: $plainPasswordToApply"
        } else {
            $pwdNotice += "  Password: [Custom SecureString Password Applied]"
        }
        $pwdNotice += @(
            "  Notice:   User will be prompted to set a new password upon next sign-in.",
            "================================================================================"
        )
        Write-Host ($pwdNotice -join "`n") -ForegroundColor Yellow
    }

} catch {
    Write-Error "An error occurred while modifying the user account: $_"
    exit 1
}
