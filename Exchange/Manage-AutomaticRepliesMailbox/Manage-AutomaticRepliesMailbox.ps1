<#
.SYNOPSIS
    Manages Microsoft 365 Exchange Online Automatic Replies (Out of Office / OOF) for user mailboxes.

.DESCRIPTION
    Production-ready PowerShell script designed for Microsoft 365 Exchange Online administrators to inspect,
    enable, disable, schedule, and configure Automatic Replies (Out of Office) for user mailboxes.
    
    Key capabilities include:
    1. Automatic detection of active Exchange Online sessions with modern interactive administrator login.
    2. Inspecting current auto-reply state (Enabled, Disabled, Scheduled) and external audience scope.
    3. Viewing and verifying existing message content for both inside (internal) and outside (external) senders.
    4. Enabling or disabling automatic replies globally or specifically for outside senders (ExternalAudience: None, Known, All).
    5. Setting or updating message content separately for internal (inside) and external (outside) senders.
    6. Loading auto-reply messages directly from text/HTML template files.
    7. Configuring scheduled auto-reply time windows (-StartTime and -EndTime).
    8. Interactive console menu mode (-Interactive) for guided administrative tasks.
    9. Input sanitization, safety pre-validations, change summary reviews, and interactive confirmation prompts (with -Force bypass).

.PARAMETER Identity
    The email address, User Principal Name (UPN), or mailbox alias of the target user mailbox.

.PARAMETER GetStatus
    Switch parameter. Displays the current automatic replies configuration, status for inside/outside senders,
    and message previews without making any changes. This is the default action when only -Identity is specified.

.PARAMETER Enable
    Switch parameter. Enables automatic replies for the specified mailbox.

.PARAMETER Disable
    Switch parameter. Disables automatic replies for the specified mailbox.

.PARAMETER AutoReplyState
    Explicitly sets the automatic reply state. Valid values: 'Enabled', 'Disabled', 'Scheduled'.

.PARAMETER ExternalAudience
    Controls which outside (external) senders receive automatic replies. Valid values:
    - 'None': Outside senders are disabled (auto-replies sent only to inside/internal senders).
    - 'Known': Auto-replies sent only to external senders present in the user's Contacts folder.
    - 'All': Auto-replies sent to all external senders.

.PARAMETER InsideOnly
    Switch parameter. Convenience switch to enable automatic replies for inside organization senders only
    (sets ExternalAudience to 'None').

.PARAMETER InternalMessage
    The text or HTML message body sent to internal senders (inside the organization).
    Alias: -InsideMessage.

.PARAMETER ExternalMessage
    The text or HTML message body sent to external senders (outside the organization).
    Alias: -OutsideMessage.

.PARAMETER InternalMessageFile
    Path to a local text or HTML file containing the message body for internal senders.

.PARAMETER ExternalMessageFile
    Path to a local text or HTML file containing the message body for external senders.

.PARAMETER StartTime
    The start date and time when automatic replies should begin sending (used with 'Scheduled' state).

.PARAMETER EndTime
    The end date and time when automatic replies should stop sending (used with 'Scheduled' state).

.PARAMETER ShowRawHtml
    Switch parameter. When viewing messages, displays the unformatted raw HTML markup instead of clean text preview.

.PARAMETER Interactive
    Switch parameter. Launches an interactive menu wizard guiding the administrator through all operations.
    Alias: -Menu.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used to pre-populate the Microsoft 365 Exchange Online login prompt.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N) for automated non-interactive scripts.

.PARAMETER Help
    Displays built-in help, usage examples, author info, and exits.
    Alias: -h.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1
    Running without parameters displays built-in help, version, author info, and usage examples.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com"
    Inspects and displays the current auto-reply status, inside/outside state, and message content for the mailbox.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -Enable -ExternalAudience All -InternalMessage "I am out of the office until Monday." -ExternalMessage "Thank you for contacting me. I am currently out of office."
    Enables automatic replies for both inside and outside senders with designated message contents.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -Enable -InsideOnly -InternalMessage "Internal OOF message."
    Enables automatic replies for inside (internal) organization senders only, disabling replies to outside senders.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -Disable -Force
    Disables automatic replies immediately without interactive confirmation prompts.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -ExternalAudience None
    Disables automatic replies for outside senders while leaving internal settings intact.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -InternalMessageFile "C:\Templates\InternalOOF.html" -ExternalMessageFile "C:\Templates\ExternalOOF.html"
    Updates auto-reply message contents for inside and outside senders from local template files.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Identity "john.doe@contoso.com" -AutoReplyState Scheduled -StartTime "2026-10-15 08:00" -EndTime "2026-10-22 17:00" -InternalMessage "On vacation."
    Schedules an automatic reply window for a specific date and time period.

.EXAMPLE
    .\Manage-AutomaticRepliesMailbox.ps1 -Interactive
    Launches the guided console wizard for interactive mailbox administration.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Primary Target Mailbox Identity
    [Parameter(Position = 0, Mandatory = $false)]
    [Alias('Mailbox', 'User')]
    [string]$Identity,

    # Operation Mode Switches
    [Parameter(Mandatory = $false)]
    [Alias('View', 'Status')]
    [switch]$GetStatus,

    [Parameter(Mandatory = $false)]
    [switch]$Enable,

    [Parameter(Mandatory = $false)]
    [switch]$Disable,

    # Configuration Parameters
    [Parameter(Mandatory = $false)]
    [ValidateSet('Enabled', 'Disabled', 'Scheduled')]
    [string]$AutoReplyState,

    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'Known', 'All')]
    [string]$ExternalAudience,

    [Parameter(Mandatory = $false)]
    [switch]$InsideOnly,

    [Parameter(Mandatory = $false)]
    [Alias('InsideMessage')]
    [string]$InternalMessage,

    [Parameter(Mandatory = $false)]
    [Alias('OutsideMessage')]
    [string]$ExternalMessage,

    [Parameter(Mandatory = $false)]
    [string]$InternalMessageFile,

    [Parameter(Mandatory = $false)]
    [string]$ExternalMessageFile,

    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartTime,

    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$EndTime,

    # Display options
    [Parameter(Mandatory = $false)]
    [switch]$ShowRawHtml,

    # Interactive Menu
    [Parameter(Mandatory = $false)]
    [Alias('Menu')]
    [switch]$Interactive,

    # Common parameters
    [Parameter(Mandatory = $false)]
    [string]$AdminUserPrincipalName,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    # Help Set
    [Parameter(Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata declaration
$ScriptMetadata = @{
    Title       = "Manage-AutomaticRepliesMailbox.ps1"
    Version     = "1.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Description = "Manage Microsoft 365 Exchange Online User Mailbox Automatic Replies (Out of Office)"
}

# Display custom formatted help documentation
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
        "    Production-ready automation script for Microsoft 365 Exchange Online administrators.",
        "    Manages Automatic Replies (Out of Office / OOF) for user mailboxes. Allows administrators",
        "    to check status and message content, toggle replies for inside/outside senders, update",
        "    message bodies independently, and schedule Out of Office windows.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    - Requires PowerShell 5.1 or PowerShell 7+.",
        "    - Requires 'ExchangeOnlineManagement' PowerShell module (auto-installed if missing).",
        "    - Requires an Exchange Online administrator role (Exchange Admin or Global Admin).",
        "    - Automatically detects active sessions; if not connected, prompts for modern administrator sign-in.",
        "    - Optional -AdminUserPrincipalName can pre-populate the administrator identity.",
        "",
        "USAGE EXAMPLES:",
        "    1. Display Built-In Help & Info (default behavior with no parameters):",
        "       .\Manage-AutomaticRepliesMailbox.ps1",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -h",
        "",
        "    2. Check Current Status & Message Contents for Inside and Outside Senders:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity `"user@contoso.com`"",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity `"user@contoso.com`" -GetStatus",
        "",
        "    3. Enable Automatic Replies for Both Inside and Outside Senders:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -Enable -ExternalAudience All -InternalMessage 'I am away until Friday.' -ExternalMessage 'Thank you for reaching out. I am away.'",
        "",
        "    4. Enable Automatic Replies for Inside Organization Senders Only (Disable Outside):",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -Enable -InsideOnly -InternalMessage 'Internal auto-reply: On annual leave.'",
        "",
        "    5. Disable Automatic Replies Completely (Turn Off for Everyone):",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -Disable",
        "",
        "    6. Disable Automatic Replies Silently Without Confirmation Prompts (-Force):",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -Disable -Force",
        "",
        "    7. Disable Outside Senders Only (Keep Inside Auto-Reply Active):",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -ExternalAudience None",
        "",
        "    8. Update Message Content for Inside Senders Only:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -InternalMessage 'Updated internal notice.'",
        "",
        "    9. Update Message Content for Outside Senders Only:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -ExternalMessage 'Updated external notice.'",
        "",
        "    10. Set Auto-Reply Messages from Local Text or HTML Template Files:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -InternalMessageFile 'C:\Templates\InternalOOF.html' -ExternalMessageFile 'C:\Templates\ExternalOOF.html'",
        "",
        "    11. Schedule Automatic Replies for a Future Time Window:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Identity 'user@contoso.com' -AutoReplyState Scheduled -StartTime '2026-10-15 08:00' -EndTime '2026-10-22 17:00' -InternalMessage 'On vacation.'",
        "",
        "    12. Interactive Console Menu Wizard:",
        "       .\Manage-AutomaticRepliesMailbox.ps1 -Interactive",
        "",
        "PARAMETERS:",
        "    -Identity               Target user mailbox email address, UPN, or alias (Required for direct operations).",
        "    -GetStatus, -View       Displays auto-reply state, audience scope, and message contents.",
        "    -Enable                 Enables automatic replies.",
        "    -Disable                Disables automatic replies.",
        "    -AutoReplyState         Explicitly sets state: 'Enabled', 'Disabled', or 'Scheduled'.",
        "    -ExternalAudience       Configures outside senders: 'None' (inside only), 'Known' (contacts), 'All' (everyone).",
        "    -InsideOnly             Convenience switch to restrict replies strictly to inside organization senders.",
        "    -InternalMessage        Message body sent to inside senders (Alias: -InsideMessage).",
        "    -ExternalMessage        Message body sent to outside senders (Alias: -OutsideMessage).",
        "    -InternalMessageFile    Path to a text/HTML template file for inside senders.",
        "    -ExternalMessageFile    Path to a text/HTML template file for outside senders.",
        "    -StartTime              Start timestamp for scheduled auto-replies.",
        "    -EndTime                End timestamp for scheduled auto-replies.",
        "    -ShowRawHtml            Displays raw HTML markup instead of plain text preview.",
        "    -Interactive, -Menu     Launches interactive menu mode.",
        "    -AdminUserPrincipalName Optional admin UPN for Exchange Online authentication.",
        "    -Force                  Bypasses confirmation prompts for silent automation.",
        "    -Help, -h               Displays this documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Manage-AutomaticRepliesMailbox.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Convert HTML markup to clean, formatted multi-line plain text for readable console display
function Convert-HtmlToPlainText {
    param([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $text = $Html -replace '<br\s*/?>', "`n"
    $text = $text -replace '(?i)</p>', "`n`n"
    $text = $text -replace '(?i)</div>', "`n"
    $text = $text -replace '<[^>]+>', ''
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return $text.Trim()
}

# Ensure ExchangeOnlineManagement module and active administrator session
function Assert-ExchangeOnlineSession {
    param([string]$AdminUPN)

    # 1. Prerequisite module check
    if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Write-Host "[!] ExchangeOnlineManagement module is not installed." -ForegroundColor Yellow
        Write-Host "[*] Installing ExchangeOnlineManagement module for CurrentUser..." -ForegroundColor Cyan
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "[+] ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to install ExchangeOnlineManagement module: $_`nPlease install manually via: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
            exit 1
        }
    }

    # 2. Check active session or prompt for modern login
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        Write-Host "[+] Active Exchange Online session detected." -ForegroundColor Green
    } catch {
        Write-Host "[*] No active Exchange Online session detected. Initiating administrator login..." -ForegroundColor Cyan
        try {
            if (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
                Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
            Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
        } catch {
            Write-Error "Failed to connect to Exchange Online: $_"
            exit 1
        }
    }
}

# Validate mailbox format and existence in tenant
function Validate-MailboxIdentity {
    param([string]$MailboxId)

    if ([string]::IsNullOrWhiteSpace($MailboxId)) {
        Write-Error "Mailbox identity cannot be empty. Please provide an email address, UPN, or alias."
        exit 1
    }

    $emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    $aliasPattern = "^[a-zA-Z0-9._-]+$"
    if ($MailboxId -notmatch $emailPattern -and $MailboxId -notmatch $aliasPattern) {
        Write-Error "Invalid mailbox identity format: '$MailboxId'. Provide a valid email address, UPN, or alias."
        exit 1
    }

    Write-Host "[*] Verifying mailbox existence in Exchange Online: $MailboxId..." -ForegroundColor Cyan
    try {
        $mbx = Get-Mailbox -Identity $MailboxId -ErrorAction Stop
        Write-Host "    Found mailbox: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress)) [Type: $($mbx.RecipientTypeDetails)]" -ForegroundColor Gray
        return $mbx
    } catch {
        Write-Error "Unable to locate mailbox '$MailboxId' in Exchange Online. Details: $_"
        exit 1
    }
}

# Read content securely from a file with size validation
function Get-MessageFromFile {
    param([string]$FilePath, [string]$MessageType)

    if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
        Write-Error "The specified $MessageType message file was not found: '$FilePath'"
        exit 1
    }

    $fileItem = Get-Item -Path $FilePath
    if ($fileItem.Length -gt 1MB) {
        Write-Error "The file '$FilePath' exceeds maximum supported size (1 MB)."
        exit 1
    }

    try {
        $content = Get-Content -Path $FilePath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-Warning "The file '$FilePath' is empty."
        }
        return $content
    } catch {
        Write-Error "Failed to read content from file '$FilePath': $_"
        exit 1
    }
}

# Retrieve and format automatic replies configuration and message content
function Show-MailboxAutoReplyStatus {
    param(
        [Parameter(Mandatory = $true)]
        $MailboxObject,
        [switch]$RawHtml
    )

    $mbxIdentity = $MailboxObject.PrimarySmtpAddress
    Write-Host "`n================================================================================" -ForegroundColor White
    Write-Host "AUTOMATIC REPLIES CONFIGURATION: $($MailboxObject.DisplayName)" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor White
    Write-Host "Mailbox Identity        : $mbxIdentity" -ForegroundColor Gray
    Write-Host "Recipient Type          : $($MailboxObject.RecipientTypeDetails)" -ForegroundColor Gray

    try {
        $config = Get-MailboxAutoReplyConfiguration -Identity $mbxIdentity -ErrorAction Stop
    } catch {
        Write-Error "Failed to retrieve automatic replies configuration for '$mbxIdentity': $_"
        return
    }

    # Format Overall State
    $stateColor = switch ($config.AutoReplyState.ToString()) {
        'Enabled'   { [ConsoleColor]::Green }
        'Scheduled' { [ConsoleColor]::Yellow }
        'Disabled'  { [ConsoleColor]::Red }
        Default     { [ConsoleColor]::White }
    }
    Write-Host "Overall AutoReply State : " -NoNewline -ForegroundColor White
    Write-Host "$($config.AutoReplyState)" -ForegroundColor $stateColor

    # Inside Senders (Internal)
    Write-Host "Inside Senders (Internal): " -NoNewline -ForegroundColor White
    if ($config.AutoReplyState.ToString() -eq 'Disabled') {
        Write-Host "INACTIVE (Overall state is Disabled)" -ForegroundColor DarkGray
    } else {
        Write-Host "ACTIVE (Replies enabled for inside organization senders)" -ForegroundColor Green
    }

    # Outside Senders (External Audience)
    Write-Host "Outside Senders (External): " -NoNewline -ForegroundColor White
    if ($config.AutoReplyState.ToString() -eq 'Disabled') {
        Write-Host "INACTIVE (Overall state is Disabled | Audience: $($config.ExternalAudience))" -ForegroundColor DarkGray
    } else {
        switch ($config.ExternalAudience.ToString()) {
            'None' {
                Write-Host "DISABLED (Audience: None - Outside senders will NOT receive replies)" -ForegroundColor Red
            }
            'Known' {
                Write-Host "ACTIVE FOR CONTACTS ONLY (Audience: Known - Senders in user contacts only)" -ForegroundColor Yellow
            }
            'All' {
                Write-Host "ACTIVE FOR ALL SDOUTERS (Audience: All - All external senders receive replies)" -ForegroundColor Green
            }
            Default {
                Write-Host "$($config.ExternalAudience)" -ForegroundColor White
            }
        }
    }

    # Scheduled Window (if applicable)
    if ($config.AutoReplyState.ToString() -eq 'Scheduled') {
        Write-Host "Scheduled Start Time    : $($config.StartTime)" -ForegroundColor Yellow
        Write-Host "Scheduled End Time      : $($config.EndTime)" -ForegroundColor Yellow
    }

    Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[1] INSIDE SENDERS (INTERNAL) MESSAGE CONTENT:" -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($config.InternalMessage)) {
        Write-Host "    (No internal message configured / Empty)" -ForegroundColor DarkGray
    } else {
        if ($RawHtml) {
            Write-Host $config.InternalMessage -ForegroundColor White
        } else {
            $plainInternal = Convert-HtmlToPlainText -Html $config.InternalMessage
            Write-Host $plainInternal -ForegroundColor White
        }
    }

    Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[2] OUTSIDE SENDERS (EXTERNAL) MESSAGE CONTENT:" -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($config.ExternalMessage)) {
        Write-Host "    (No external message configured / Empty)" -ForegroundColor DarkGray
    } else {
        if ($RawHtml) {
            Write-Host $config.ExternalMessage -ForegroundColor White
        } else {
            $plainExternal = Convert-HtmlToPlainText -Html $config.ExternalMessage
            Write-Host $plainExternal -ForegroundColor White
        }
    }
    Write-Host "================================================================================`n" -ForegroundColor White

    return $config
}

# Apply new configuration to Exchange Online
function Set-MailboxAutoReplySettings {
    param(
        [Parameter(Mandatory = $true)]
        $MailboxObject,
        [string]$NewState,
        [string]$NewAudience,
        [string]$NewInternalMsg,
        [string]$NewExternalMsg,
        [Nullable[DateTime]]$NewStart,
        [Nullable[DateTime]]$NewEnd,
        [switch]$BypassConfirmation
    )

    $mbxIdentity = $MailboxObject.PrimarySmtpAddress

    # Retrieve current configuration to calculate diff
    try {
        $currentConfig = Get-MailboxAutoReplyConfiguration -Identity $mbxIdentity -ErrorAction Stop
    } catch {
        Write-Error "Failed to retrieve current configuration for '$mbxIdentity': $_"
        exit 1
    }

    # Resolve target values
    $targetState = if (-not [string]::IsNullOrWhiteSpace($NewState)) { $NewState } else { $currentConfig.AutoReplyState.ToString() }
    $targetAudience = if (-not [string]::IsNullOrWhiteSpace($NewAudience)) { $NewAudience } else { $currentConfig.ExternalAudience.ToString() }
    
    $updateInternal = ($null -ne $NewInternalMsg)
    $updateExternal = ($null -ne $NewExternalMsg)

    # Validate Scheduled state dates
    if ($targetState -eq 'Scheduled') {
        $resolvedStart = if ($null -ne $NewStart) { $NewStart } else { $currentConfig.StartTime }
        $resolvedEnd = if ($null -ne $NewEnd) { $NewEnd } else { $currentConfig.EndTime }

        if ($resolvedStart -ge $resolvedEnd) {
            Write-Error "Invalid schedule time window: StartTime ($resolvedStart) must be earlier than EndTime ($resolvedEnd)."
            exit 1
        }
    }

    # Display Change Summary
    Write-Host "`n================================================================================" -ForegroundColor Yellow
    Write-Host "ACTION SUMMARY: PROPOSED AUTOMATIC REPLIES MODIFICATIONS" -ForegroundColor Yellow
    Write-Host "================================================================================" -ForegroundColor Yellow
    Write-Host "Target Mailbox          : $($MailboxObject.DisplayName) ($mbxIdentity)" -ForegroundColor White
    
    $stateDiff = if ($targetState -ne $currentConfig.AutoReplyState.ToString()) { "$targetState [CHANGED from $($currentConfig.AutoReplyState)]" } else { "$targetState [Unchanged]" }
    Write-Host "AutoReply State         : $stateDiff" -ForegroundColor White

    $audienceDiff = if ($targetAudience -ne $currentConfig.ExternalAudience.ToString()) { "$targetAudience [CHANGED from $($currentConfig.ExternalAudience)]" } else { "$targetAudience [Unchanged]" }
    Write-Host "External Audience Scope : $audienceDiff" -ForegroundColor White

    $intMsgDiff = if ($updateInternal) { "[WILL UPDATE] Message length: $($NewInternalMsg.Length) chars" } else { "[Retain current]" }
    Write-Host "Internal (Inside) Msg   : $intMsgDiff" -ForegroundColor White

    $extMsgDiff = if ($updateExternal) { "[WILL UPDATE] Message length: $($NewExternalMsg.Length) chars" } else { "[Retain current]" }
    Write-Host "External (Outside) Msg  : $extMsgDiff" -ForegroundColor White

    if ($targetState -eq 'Scheduled') {
        Write-Host "Scheduled Window        : $resolvedStart -> $resolvedEnd" -ForegroundColor White
    }
    Write-Host "================================================================================" -ForegroundColor Yellow

    # Interactive confirmation prompt (language independent check for Y/N or T/N)
    if (-not $BypassConfirmation) {
        $confirm = Read-Host "Are you sure you want to apply these changes to '$mbxIdentity'? (Y/N)"
        if ($confirm -notmatch "^(?i)(y|yes|t|tak)$") {
            Write-Host "[!] Operation cancelled by user. No modifications were made." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "[*] -Force switch detected. Proceeding without interactive confirmation..." -ForegroundColor Gray
    }

    # Build parameter splatting for Set-MailboxAutoReplyConfiguration
    $cmdParams = @{
        Identity = $mbxIdentity
    }

    if (-not [string]::IsNullOrWhiteSpace($NewState)) {
        $cmdParams['AutoReplyState'] = $NewState
    }

    if (-not [string]::IsNullOrWhiteSpace($NewAudience)) {
        $cmdParams['ExternalAudience'] = $NewAudience
    }

    if ($updateInternal) {
        $cmdParams['InternalMessage'] = $NewInternalMsg
    }

    if ($updateExternal) {
        $cmdParams['ExternalMessage'] = $NewExternalMsg
    }

    if ($targetState -eq 'Scheduled') {
        if ($null -ne $NewStart) { $cmdParams['StartTime'] = $NewStart }
        if ($null -ne $NewEnd) { $cmdParams['EndTime'] = $NewEnd }
    }

    Write-Host "`n[*] Applying updated automatic replies configuration..." -ForegroundColor Cyan
    try {
        Set-MailboxAutoReplyConfiguration @cmdParams -ErrorAction Stop
        Write-Host "[+] Automatic replies configuration updated successfully for '$mbxIdentity'." -ForegroundColor Green
    } catch {
        Write-Error "Failed to update automatic replies configuration for '$mbxIdentity': $_"
        exit 1
    }

    # Verify and show updated live status
    Write-Host "`n[*] Verifying updated live configuration:" -ForegroundColor Cyan
    Show-MailboxAutoReplyStatus -MailboxObject $MailboxObject | Out-Null
}

# Guided interactive console wizard mode
function Invoke-InteractiveWizard {
    param([string]$AdminUPN)

    Assert-ExchangeOnlineSession -AdminUPN $AdminUPN

    Write-Host "`n================================================================================" -ForegroundColor Cyan
    Write-Host "  INTERACTIVE WIZARD: EXCHANGE ONLINE AUTOMATIC REPLIES MANAGEMENT" -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan

    $targetEmail = Read-Host "Enter the user email address, UPN, or mailbox alias"
    if ([string]::IsNullOrWhiteSpace($targetEmail)) {
        Write-Warning "No mailbox identity entered. Exiting wizard."
        return
    }

    $mbx = Validate-MailboxIdentity -MailboxId $targetEmail

    while ($true) {
        Write-Host "`n========================================================" -ForegroundColor Yellow
        Write-Host "MAILBOX: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -ForegroundColor Yellow
        Write-Host "========================================================" -ForegroundColor Yellow
        Write-Host "1. View Current Status & Message Contents (Inside & Outside)"
        Write-Host "2. Enable Automatic Replies (Both Inside & Outside)"
        Write-Host "3. Enable Automatic Replies for INSIDE Senders Only (Disable Outside)"
        Write-Host "4. Disable Automatic Replies Completely (Turn Off for Everyone)"
        Write-Host "5. Change Outside Senders Scope (External Audience: None / Known / All)"
        Write-Host "6. Update Message Content for Inside Senders"
        Write-Host "7. Update Message Content for Outside Senders"
        Write-Host "8. Schedule Automatic Replies Window (Start / End Date)"
        Write-Host "0. Exit"
        Write-Host "========================================================" -ForegroundColor Yellow

        $choice = Read-Host "Select an option [0-8]"

        switch ($choice) {
            '1' {
                Show-MailboxAutoReplyStatus -MailboxObject $mbx | Out-Null
            }
            '2' {
                $intMsg = Read-Host "Enter Internal (Inside) message (press Enter to keep current)"
                $extMsg = Read-Host "Enter External (Outside) message (press Enter to keep current)"
                $pInt = if ([string]::IsNullOrWhiteSpace($intMsg)) { $null } else { $intMsg }
                $pExt = if ([string]::IsNullOrWhiteSpace($extMsg)) { $null } else { $extMsg }
                
                Set-MailboxAutoReplySettings -MailboxObject $mbx `
                    -NewState 'Enabled' `
                    -NewAudience 'All' `
                    -NewInternalMsg $pInt `
                    -NewExternalMsg $pExt
            }
            '3' {
                $intMsg = Read-Host "Enter Internal (Inside) message (press Enter to keep current)"
                $pInt = if ([string]::IsNullOrWhiteSpace($intMsg)) { $null } else { $intMsg }
                
                Set-MailboxAutoReplySettings -MailboxObject $mbx `
                    -NewState 'Enabled' `
                    -NewAudience 'None' `
                    -NewInternalMsg $pInt `
                    -NewExternalMsg $null
            }
            '4' {
                Set-MailboxAutoReplySettings -MailboxObject $mbx `
                    -NewState 'Disabled' `
                    -NewAudience $null `
                    -NewInternalMsg $null `
                    -NewExternalMsg $null
            }
            '5' {
                Write-Host "Select External Audience:"
                Write-Host "  1. None  (Disable replies to outside senders)"
                Write-Host "  2. Known (Only outside senders in user's Contacts)"
                Write-Host "  3. All   (All outside senders)"
                $audChoice = Read-Host "Choose [1-3]"
                $targetAud = switch ($audChoice) {
                    '1' { 'None' }
                    '2' { 'Known' }
                    '3' { 'All' }
                    Default { $null }
                }
                if ($targetAud) {
                    Set-MailboxAutoReplySettings -MailboxObject $mbx `
                        -NewState $null `
                        -NewAudience $targetAud `
                        -NewInternalMsg $null `
                        -NewExternalMsg $null
                } else {
                    Write-Warning "Invalid choice. Audience was not changed."
                }
            }
            '6' {
                $intMsg = Read-Host "Enter new message body for Inside Senders"
                if (-not [string]::IsNullOrWhiteSpace($intMsg)) {
                    Set-MailboxAutoReplySettings -MailboxObject $mbx `
                        -NewState $null `
                        -NewAudience $null `
                        -NewInternalMsg $intMsg `
                        -NewExternalMsg $null
                } else {
                    Write-Warning "Empty message entered. Cancelled."
                }
            }
            '7' {
                $extMsg = Read-Host "Enter new message body for Outside Senders"
                if (-not [string]::IsNullOrWhiteSpace($extMsg)) {
                    Set-MailboxAutoReplySettings -MailboxObject $mbx `
                        -NewState $null `
                        -NewAudience $null `
                        -NewInternalMsg $null `
                        -NewExternalMsg $extMsg
                } else {
                    Write-Warning "Empty message entered. Cancelled."
                }
            }
            '8' {
                $startStr = Read-Host "Enter Start Time (e.g. 2026-10-15 08:00)"
                $endStr = Read-Host "Enter End Time (e.g. 2026-10-22 17:00)"
                try {
                    $dtStart = [DateTime]::Parse($startStr)
                    $dtEnd = [DateTime]::Parse($endStr)
                    Set-MailboxAutoReplySettings -MailboxObject $mbx `
                        -NewState 'Scheduled' `
                        -NewAudience $null `
                        -NewInternalMsg $null `
                        -NewExternalMsg $null `
                        -NewStart $dtStart `
                        -NewEnd $dtEnd
                } catch {
                    Write-Error "Invalid date format entered: $_"
                }
            }
            '0' {
                Write-Host "Exiting interactive wizard." -ForegroundColor Green
                return
            }
            Default {
                Write-Warning "Invalid choice. Please choose a number between 0 and 8."
            }
        }
    }
}

# ==============================================================================
# SCRIPT ENTRY POINT & PARAMETER PROCESSING
# ==============================================================================

# 1. Parameterless execution or -Help/-h triggers built-in help and exits cleanly
if ($Help -or ($PSBoundParameters.Count -eq 0 -and -not $Interactive)) {
    Show-ScriptHelp
    exit 0
}

# 2. Interactive Wizard Mode
if ($Interactive) {
    Invoke-InteractiveWizard -AdminUPN $AdminUserPrincipalName
    exit 0
}

# 3. Check for required Identity in direct executions
if ([string]::IsNullOrWhiteSpace($Identity)) {
    Write-Error "Missing required parameter: -Identity is required. Provide a valid email address, UPN, or mailbox alias. Run with -h for help."
    exit 1
}

# 4. Input sanitization & mutual exclusion validations
if ($Enable -and $Disable) {
    Write-Error "Conflicting parameters: Cannot specify both -Enable and -Disable simultaneously."
    exit 1
}

if (-not [string]::IsNullOrWhiteSpace($InternalMessage) -and -not [string]::IsNullOrWhiteSpace($InternalMessageFile)) {
    Write-Error "Conflicting parameters: Cannot specify both -InternalMessage and -InternalMessageFile simultaneously."
    exit 1
}

if (-not [string]::IsNullOrWhiteSpace($ExternalMessage) -and -not [string]::IsNullOrWhiteSpace($ExternalMessageFile)) {
    Write-Error "Conflicting parameters: Cannot specify both -ExternalMessage and -ExternalMessageFile simultaneously."
    exit 1
}

if ($InsideOnly -and -not [string]::IsNullOrWhiteSpace($ExternalAudience) -and $ExternalAudience -ne 'None') {
    Write-Error "Conflicting parameters: -InsideOnly sets ExternalAudience to 'None'. Cannot conflict with -ExternalAudience '$ExternalAudience'."
    exit 1
}

# 5. Format validation for Target Mailbox Identity
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
$aliasPattern = "^[a-zA-Z0-9._-]+$"
if ($Identity -notmatch $emailPattern -and $Identity -notmatch $aliasPattern) {
    Write-Error "Invalid mailbox identity format: '$Identity'. Provide a valid email address, UPN, or mailbox alias."
    exit 1
}

# 6. Pre-validate and resolve message files if specified
$resolvedInternalMsg = if (-not [string]::IsNullOrWhiteSpace($InternalMessageFile)) {
    Get-MessageFromFile -FilePath $InternalMessageFile -MessageType "internal"
} elseif ($PSBoundParameters.ContainsKey('InternalMessage')) {
    $InternalMessage
} else {
    $null
}

$resolvedExternalMsg = if (-not [string]::IsNullOrWhiteSpace($ExternalMessageFile)) {
    Get-MessageFromFile -FilePath $ExternalMessageFile -MessageType "external"
} elseif ($PSBoundParameters.ContainsKey('ExternalMessage')) {
    $ExternalMessage
} else {
    $null
}

# 7. Pre-validate schedule date range if provided
if (($null -ne $StartTime) -and ($null -ne $EndTime)) {
    if ($StartTime -ge $EndTime) {
        Write-Error "Invalid schedule time window: StartTime ($StartTime) must be earlier than EndTime ($EndTime)."
        exit 1
    }
}

# 8. Direct Execution: Ensure Exchange Online Connection
Assert-ExchangeOnlineSession -AdminUPN $AdminUserPrincipalName

# 9. Validate Target Mailbox in tenant
$mailbox = Validate-MailboxIdentity -MailboxId $Identity

# 10. Resolve Target AutoReplyState
$resolvedState = if ($Enable) {
    if (($null -ne $StartTime) -and ($null -ne $EndTime)) { 'Scheduled' } else { 'Enabled' }
} elseif ($Disable) {
    'Disabled'
} elseif (-not [string]::IsNullOrWhiteSpace($AutoReplyState)) {
    $AutoReplyState
} else {
    $null
}

# 11. Resolve External Audience
$resolvedAudience = if ($InsideOnly) {
    'None'
} elseif (-not [string]::IsNullOrWhiteSpace($ExternalAudience)) {
    $ExternalAudience
} else {
    $null
}

# 12. Determine if this is a Read-Only Query or a Configuration Change
$isChangeRequested = (
    $Enable -or 
    $Disable -or 
    (-not [string]::IsNullOrWhiteSpace($resolvedState)) -or 
    (-not [string]::IsNullOrWhiteSpace($resolvedAudience)) -or 
    ($null -ne $resolvedInternalMsg) -or 
    ($null -ne $resolvedExternalMsg) -or 
    ($null -ne $StartTime) -or 
    ($null -ne $EndTime)
)

if (-not $isChangeRequested -or $GetStatus) {
    # Read-Only Inspection Mode: Check status and message contents
    Show-MailboxAutoReplyStatus -MailboxObject $mailbox -RawHtml:$ShowRawHtml | Out-Null
    exit 0
}

# 9. Perform Configuration Change
Set-MailboxAutoReplySettings `
    -MailboxObject $mailbox `
    -NewState $resolvedState `
    -NewAudience $resolvedAudience `
    -NewInternalMsg $resolvedInternalMsg `
    -NewExternalMsg $resolvedExternalMsg `
    -NewStart $StartTime `
    -NewEnd $EndTime `
    -BypassConfirmation:$Force
