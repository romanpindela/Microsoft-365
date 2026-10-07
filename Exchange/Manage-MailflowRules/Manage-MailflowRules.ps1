<#
.SYNOPSIS
    Manages Microsoft 365 Exchange Online Mail Flow (Transport) rules: audits rules, toggles state, and creates BCC (UDW) redirection rules.

.DESCRIPTION
    Production-ready PowerShell script designed for Microsoft 365 Exchange Online administrators to manage Mail Flow
    (Transport) rules with enterprise-grade safety, input validation, and interactive or automated execution.
    Capabilities include:
    1. Automatic Exchange Online administrator session detection and modern interactive login.
    2. Comprehensive listing of all Mail Flow rules in a formatted table with detailed status (Priority, Name, State, Mode, Comments).
    3. Enabling and disabling specific mail flow rules with interactive confirmation (or silent automation via -Force).
    4. Creating new transport rules to copy/redirect messages from a specified mailbox to designated BCC (UDW) recipients.
    5. Interactive console menu mode (-Interactive) for guided administration.

.PARAMETER ListRules
    Switch parameter. Lists all Mail Flow (Transport) rules in a structured table.

.PARAMETER StateFilter
    Filters the listed rules by state. Valid options: 'All', 'Enabled', 'Disabled'. Default is 'All'.

.PARAMETER NameFilter
    Optional wildcard string to filter rules by name (e.g. '*BCC*' or 'Forward*').

.PARAMETER Detailed
    Switch parameter. Displays in-depth rule configuration including conditions, actions, and auto-generated descriptions.

.PARAMETER EnableRule
    Switch parameter. Enables the specified mail flow rule.

.PARAMETER DisableRule
    Switch parameter. Disables the specified mail flow rule.

.PARAMETER RuleIdentity
    The Name, Identity, or GUID of the mail flow rule to enable, disable, or inspect.

.PARAMETER NewBccRule
    Switch parameter. Creates a new mail flow rule that blind carbon copies (BCC / UDW) messages to specified recipients.

.PARAMETER RuleName
    The unique name for the new mail flow rule.

.PARAMETER SourceMailbox
    The email address or UPN of the mailbox to monitor (e.g., 'sales@domain.com').

.PARAMETER BccRecipients
    One or more email addresses to receive the blind carbon copy (BCC / UDW) messages. Can be passed as an array or comma-separated string.

.PARAMETER Direction
    Specifies message traffic direction to trigger the rule:
    - 'Incoming' (default): Triggered when messages are sent TO the source mailbox (-SentTo).
    - 'Outgoing': Triggered when messages are sent FROM the source mailbox (-From).

.PARAMETER Comments
    Optional administrative notes or explanation to store inside the rule.

.PARAMETER Priority
    Optional integer defining rule evaluation order (higher priority rules are evaluated first).

.PARAMETER Interactive
    Switch parameter. Launches an interactive console menu guiding the administrator through all available actions.

.PARAMETER AdminUserPrincipalName
    Optional administrator UPN used to pre-fill the Microsoft 365 Exchange Online login prompt.

.PARAMETER Force
    Switch parameter. Bypasses interactive confirmation prompts (Y/N) for automated non-interactive scripts.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Manage-MailflowRules.ps1
    Running without parameters displays this help documentation and usage examples.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -ListRules
    Lists all mail flow rules in a clean formatted table showing Priority, Name, State, Mode, and Comments.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -ListRules -StateFilter Enabled -Detailed
    Lists only enabled mail flow rules with full rule details and action summaries.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -EnableRule -RuleIdentity "BCC - Financial Compliance"
    Enables the specified rule with interactive confirmation.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -DisableRule -RuleIdentity "BCC - Financial Compliance" -Force
    Disables the specified rule silently without confirmation prompts.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -NewBccRule -RuleName "BCC - Inbound Sales Copy" -SourceMailbox "sales@domain.com" -BccRecipients "audit@domain.com"
    Creates a new rule that automatically copies all incoming emails received by sales@domain.com to audit@domain.com via BCC (UDW).

.EXAMPLE
    .\Manage-MailflowRules.ps1 -NewBccRule -RuleName "BCC - Outbound Executive Audit" -SourceMailbox "ceo@domain.com" -BccRecipients "compliance@domain.com","archive@domain.com" -Direction Outgoing -Comments "Executive audit rule"
    Creates a rule BCC'ing all outgoing emails sent by ceo@domain.com to compliance and archive addresses.

.EXAMPLE
    .\Manage-MailflowRules.ps1 -Interactive
    Launches an interactive, menu-driven console wizard for managing rules.

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Set: ListRules
    [Parameter(ParameterSetName = 'ListRules', Mandatory = $false)]
    [switch]$ListRules,

    [Parameter(ParameterSetName = 'ListRules', Mandatory = $false)]
    [ValidateSet('All', 'Enabled', 'Disabled')]
    [string]$StateFilter = 'All',

    [Parameter(ParameterSetName = 'ListRules', Mandatory = $false)]
    [string]$NameFilter,

    [Parameter(ParameterSetName = 'ListRules', Mandatory = $false)]
    [switch]$Detailed,

    # Set: EnableRule
    [Parameter(ParameterSetName = 'EnableRule', Mandatory = $true)]
    [switch]$EnableRule,

    # Set: DisableRule
    [Parameter(ParameterSetName = 'DisableRule', Mandatory = $true)]
    [switch]$DisableRule,

    [Parameter(ParameterSetName = 'EnableRule', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'DisableRule', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'ListRules', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$RuleIdentity,

    # Set: NewBccRule
    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $true)]
    [switch]$NewBccRule,

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$RuleName,

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceMailbox,

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$BccRecipients,

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $false)]
    [ValidateSet('Incoming', 'Outgoing')]
    [string]$Direction = 'Incoming',

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $false)]
    [string]$Comments,

    [Parameter(ParameterSetName = 'NewBccRule', Mandatory = $false)]
    [int]$Priority,

    # Set: Interactive Menu
    [Parameter(ParameterSetName = 'Interactive', Mandatory = $false)]
    [Alias('Menu')]
    [switch]$Interactive,

    # Common parameters
    [Parameter(Mandatory = $false)]
    [string]$AdminUserPrincipalName,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    # Set: Help
    [Parameter(ParameterSetName = 'Help', Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Metadata declaration
$ScriptMetadata = @{
    Title   = "Manage-MailflowRules.ps1"
    Version = "1.0.0"
    Author  = "Roman Pindela"
    Email   = "roman.pindela@gmail.com"
    GitHub  = "https://github.com/romanpindela"
}

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
        "    Manages Mail Flow (Transport) rules: audits rules, toggles state (Enable/Disable),",
        "    creates automated BCC (UDW) redirection/monitoring rules, and supports interactive console usage.",
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
        "       .\Manage-MailflowRules.ps1",
        "       .\Manage-MailflowRules.ps1 -h",
        "",
        "    2. List All Mail Flow Rules (Summary Table):",
        "       .\Manage-MailflowRules.ps1 -ListRules",
        "",
        "    3. List Only Enabled Rules with Wildcard Filter & Detailed Configuration:",
        "       .\Manage-MailflowRules.ps1 -ListRules -StateFilter Enabled -NameFilter `"*Audit*`" -Detailed",
        "",
        "    4. Enable a Specific Mail Flow Rule (Interactive Confirmation):",
        "       .\Manage-MailflowRules.ps1 -EnableRule -RuleIdentity `"BCC - Legal Hold`"",
        "",
        "    5. Disable a Mail Flow Rule Silently (-Force):",
        "       .\Manage-MailflowRules.ps1 -DisableRule -RuleIdentity `"BCC - Legal Hold`" -Force",
        "",
        "    6. Create a New BCC (UDW) Redirection Rule (Incoming Messages):",
        "       .\Manage-MailflowRules.ps1 -NewBccRule -RuleName `"BCC - Inbound Sales`" -SourceMailbox `"sales@domain.com`" -BccRecipients `"compliance@domain.com`"",
        "",
        "    7. Create a BCC Rule for Outgoing Messages with Multiple Recipients:",
        "       .\Manage-MailflowRules.ps1 -NewBccRule -RuleName `"BCC - Executive Outbound`" -SourceMailbox `"ceo@domain.com`" -BccRecipients `"archive@domain.com`",`"auditor@domain.com`" -Direction Outgoing -Comments `"Mandatory executive audit`"",
        "",
        "    8. Interactive Console Menu Wizard:",
        "       .\Manage-MailflowRules.ps1 -Interactive",
        "",
        "PARAMETERS:",
        "    -ListRules                Lists Mail Flow rules in a structured table.",
        "    -StateFilter              Filter listed rules by state: 'All', 'Enabled', or 'Disabled'.",
        "    -NameFilter               Wildcard name pattern to filter listed rules.",
        "    -Detailed                 Displays deep configuration details (actions, conditions, description).",
        "    -EnableRule               Enables the specified mail flow rule.",
        "    -DisableRule              Disables the specified mail flow rule.",
        "    -RuleIdentity             Name or Identity/GUID of the rule to toggle or view.",
        "    -NewBccRule               Creates a new mail flow rule adding BCC (UDW) recipients.",
        "    -RuleName                 Unique name for the new rule.",
        "    -SourceMailbox            Source mailbox email address to monitor.",
        "    -BccRecipients            One or more recipient email addresses for BCC (UDW).",
        "    -Direction                Traffic direction: 'Incoming' (SentTo) or 'Outgoing' (From). Default: 'Incoming'.",
        "    -Comments                 Optional administrator comments/description for the rule.",
        "    -Priority                 Optional integer priority order for rule evaluation.",
        "    -Interactive, -Menu       Launches an interactive menu-driven console wizard.",
        "    -AdminUserPrincipalName   (Optional) Administrator UPN for Exchange Online connection.",
        "    -Force                    Bypasses interactive confirmation prompts (Y/N).",
        "    -Help, -h                 Displays this help screen.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, unblock the script before execution:",
        "    Unblock-File -Path .\Manage-MailflowRules.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Auto-show help when executed without parameters, or with -Help / -h
$hasExplicitAction = $ListRules -or $EnableRule -or $DisableRule -or $NewBccRule -or $Interactive
if ($Help -or ($PSCmdlet.ParameterSetName -eq 'Help') -or ($PSCmdlet.ParameterSetName -eq 'Default' -and -not $hasExplicitAction)) {
    Show-ScriptHelp
    exit 0
}

# Validation helper for email addresses (defensive protection against invalid or malicious input)
function Test-ValidEmailAddress {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    $emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return ($Email.Trim() -match $emailPattern)
}

# Upfront input sanitization and verification before connecting to cloud
$cleanBccRecipients = @()
if ($NewBccRule) {
    if ([string]::IsNullOrWhiteSpace($RuleName) -or $RuleName.Length -gt 64 -or $RuleName -match '["<>|]') {
        Write-Error "Invalid RuleName: Name cannot contain quotes, angle brackets (<>), pipe symbols (|), and must not exceed 64 characters."
        exit 1
    }

    if (-not (Test-ValidEmailAddress -Email $SourceMailbox)) {
        Write-Error "Invalid format for -SourceMailbox: '$SourceMailbox'. Must be a valid email address."
        exit 1
    }

    foreach ($entry in $BccRecipients) {
        $splitItems = $entry -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($item in $splitItems) {
            if (-not (Test-ValidEmailAddress -Email $item)) {
                Write-Error "Invalid format for BCC recipient: '$item'. Must be a valid email address."
                exit 1
            }
            $cleanBccRecipients += $item
        }
    }

    if ($cleanBccRecipients.Count -eq 0) {
        Write-Error "At least one valid recipient email address must be provided in -BccRecipients."
        exit 1
    }

    $cleanBccRecipients = $cleanBccRecipients | Select-Object -Unique
}

if (($EnableRule -or $DisableRule) -and [string]::IsNullOrWhiteSpace($RuleIdentity)) {
    Write-Error "RuleIdentity cannot be empty when enabling or disabling a rule."
    exit 1
}

# Exchange Online Connection Manager
function Connect-ExchangeEnvironment {
    param([string]$AdminUPN)

    # 1. Check prerequisite module
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

    # 2. Check active session or prompt for administrator login
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

# Helper to locate a transport rule by Name or Identity/GUID
function Find-TransportRule {
    param([string]$Identity)

    if ([string]::IsNullOrWhiteSpace($Identity)) {
        throw "RuleIdentity cannot be empty."
    }

    try {
        $found = Get-TransportRule -Identity $Identity -ErrorAction SilentlyContinue
        if ($found) { return $found }
    } catch {
        # Fall through to search by Name
    }

    $allRules = Get-TransportRule -ErrorAction Stop
    $matches = $allRules | Where-Object { $_.Name -eq $Identity -or $_.Identity -eq $Identity }

    if ($matches.Count -eq 1) {
        return $matches[0]
    } elseif ($matches.Count -gt 1) {
        throw "Multiple rules found matching '$Identity'. Please specify the unique Rule Identity/GUID."
    }

    return $null
}

# Feature 2: List Mail Flow Rules
function Get-MailflowRulesList {
    param(
        [string]$State = 'All',
        [string]$FilterName,
        [switch]$DetailedView
    )

    Write-Host "`n[*] Retrieving Mail Flow (Transport) rules from Exchange Online..." -ForegroundColor Cyan
    try {
        $rules = Get-TransportRule -ErrorAction Stop
    } catch {
        Write-Error "Failed to retrieve transport rules: $_"
        return
    }

    if (-not $rules -or $rules.Count -eq 0) {
        Write-Host "[i] No Mail Flow rules found in this tenant." -ForegroundColor Yellow
        return
    }

    # Apply State filter
    if ($State -ne 'All') {
        $rules = $rules | Where-Object { $_.State.ToString() -eq $State }
    }

    # Apply Name filter
    if (-not [string]::IsNullOrWhiteSpace($FilterName)) {
        $rules = $rules | Where-Object { $_.Name -like $FilterName }
    }

    $ruleCount = @($rules).Count
    Write-Host "[+] Found $ruleCount rule(s) matching criteria (State: $State, NameFilter: $(if ($FilterName) { $FilterName } else { 'None' })):" -ForegroundColor Green

    if ($ruleCount -eq 0) {
        Write-Host "    No rules matched the specified filter criteria." -ForegroundColor Gray
        return
    }

    # Sort by Priority
    $sortedRules = $rules | Sort-Object Priority

    if ($DetailedView) {
        foreach ($r in $sortedRules) {
            Write-Host "`n--------------------------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host " Rule Name   : $($r.Name)" -ForegroundColor White
            $stateColor = if ($r.State -eq 'Enabled') { 'Green' } else { 'DarkYellow' }
            Write-Host " State       : $($r.State)" -ForegroundColor $stateColor
            Write-Host " Priority    : $($r.Priority)" -ForegroundColor White
            Write-Host " Mode        : $($r.Mode)" -ForegroundColor White
            Write-Host " Identity    : $($r.Identity)" -ForegroundColor DarkGray
            if ($r.Comments) {
                Write-Host " Comments    : $($r.Comments)" -ForegroundColor Gray
            }
            if ($r.Description) {
                Write-Host " Description :" -ForegroundColor Yellow
                Write-Host "   $($r.Description)" -ForegroundColor Gray
            }
        }
        Write-Host "--------------------------------------------------------------------------------`n" -ForegroundColor Cyan
    } else {
        $tableData = $sortedRules | Select-Object `
            @{Name = 'Priority'; Expression = { $_.Priority } }, `
            @{Name = 'Name';     Expression = { $_.Name } }, `
            @{Name = 'State';    Expression = { $_.State } }, `
            @{Name = 'Mode';     Expression = { $_.Mode } }, `
            @{Name = 'Comments'; Expression = { if ($_.Comments) { $_.Comments } else { '-' } } }

        $tableData | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor White
    }
}

# Feature 3: Enable or Disable Rule
function Set-MailflowRuleState {
    param(
        [Parameter(Mandatory = $true)] [string]$Identity,
        [Parameter(Mandatory = $true)] [ValidateSet('Enable', 'Disable')] [string]$Action,
        [switch]$ForceAction
    )

    Write-Host "`n[*] Searching for mail flow rule: '$Identity'..." -ForegroundColor Cyan
    try {
        $rule = Find-TransportRule -Identity $Identity
    } catch {
        Write-Error $_.Exception.Message
        exit 1
    }

    if (-not $rule) {
        Write-Error "Mail flow rule '$Identity' was not found in Exchange Online."
        exit 1
    }

    $targetState = if ($Action -eq 'Enable') { 'Enabled' } else { 'Disabled' }
    if ($rule.State.ToString() -eq $targetState) {
        Write-Host "[i] Rule '$($rule.Name)' is already in '$targetState' state. No change needed." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host " MAIL FLOW RULE STATE MODIFICATION" -ForegroundColor Cyan
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host " Rule Name       : $($rule.Name)" -ForegroundColor White
    Write-Host " Current State   : $($rule.State)" -ForegroundColor Yellow
    Write-Host " Requested Action: $Action rule (New state -> $targetState)" -ForegroundColor Cyan
    Write-Host " Priority        : $($rule.Priority)" -ForegroundColor White
    Write-Host " Mode            : $($rule.Mode)" -ForegroundColor White
    Write-Host "=======================================================" -ForegroundColor Cyan

    if (-not $ForceAction) {
        $prompt = Read-Host "`n[?] Are you sure you want to $Action rule '$($rule.Name)'? (Y/N)"
        if ($prompt -notmatch '^(Y|y|T|t)$') {
            Write-Host "[-] Action canceled by administrator." -ForegroundColor Gray
            return
        }
    }

    try {
        if ($Action -eq 'Enable') {
            Enable-TransportRule -Identity $rule.Identity -Confirm:$false -ErrorAction Stop
        } else {
            Disable-TransportRule -Identity $rule.Identity -Confirm:$false -ErrorAction Stop
        }

        # Verify updated state
        $updatedRule = Get-TransportRule -Identity $rule.Identity -ErrorAction Stop
        $finalColor = if ($updatedRule.State -eq 'Enabled') { 'Green' } else { 'Yellow' }
        Write-Host "[+] Successfully $($Action)d rule '$($updatedRule.Name)'!" -ForegroundColor Green
        Write-Host "    New State: $($updatedRule.State)" -ForegroundColor $finalColor
    } catch {
        Write-Error "Failed to $Action mail flow rule '$($rule.Name)': $_"
        exit 1
    }
}

# Feature 4: Create BCC (UDW) Redirection Rule
function New-BccMailflowRule {
    param(
        [Parameter(Mandatory = $true)] [string]$RuleName,
        [Parameter(Mandatory = $true)] [string]$SourceMailbox,
        [Parameter(Mandatory = $true)] [string[]]$Recipients,
        [Parameter(Mandatory = $false)] [string]$Direction = 'Incoming',
        [Parameter(Mandatory = $false)] [string]$Comments,
        [Parameter(Mandatory = $false)] [int]$Priority,
        [switch]$ForceAction
    )

    # Check if rule with this name already exists
    Write-Host "[*] Checking for duplicate rule name: '$RuleName'..." -ForegroundColor Cyan
    $existing = Find-TransportRule -Identity $RuleName
    if ($existing) {
        Write-Error "A mail flow rule named '$RuleName' already exists (State: $($existing.State), Priority: $($existing.Priority)). Choose a unique name."
        exit 1
    }

    # Pre-validate source mailbox existence in Exchange Online
    Write-Host "[*] Validating source mailbox '$SourceMailbox'..." -ForegroundColor Cyan
    try {
        $sourceObj = Get-Recipient -Identity $SourceMailbox -ErrorAction Stop
        Write-Host "    Found source recipient: $($sourceObj.DisplayName) ($($sourceObj.RecipientTypeDetails))" -ForegroundColor Gray
    } catch {
        Write-Warning "Unable to verify recipient '$SourceMailbox' in Exchange Online directory: $_"
        if (-not $ForceAction) {
            $proceed = Read-Host "[?] Recipient could not be verified. Do you want to proceed anyway? (Y/N)"
            if ($proceed -notmatch '^(Y|y|T|t)$') {
                Write-Host "[-] Rule creation aborted." -ForegroundColor Gray
                return
            }
        }
    }

    # Present summary preview
    $directionDesc = if ($Direction -eq 'Outgoing') {
        "Outgoing (Messages sent FROM '$SourceMailbox')"
    } else {
        "Incoming (Messages sent TO '$SourceMailbox')"
    }

    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host " CREATE NEW BCC (UDW) REDIRECTION RULE" -ForegroundColor Cyan
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host " Rule Name       : $RuleName" -ForegroundColor White
    Write-Host " Direction       : $directionDesc" -ForegroundColor White
    Write-Host " Source Mailbox  : $SourceMailbox" -ForegroundColor White
    Write-Host " BCC Recipients  : $($Recipients -join ', ')" -ForegroundColor White
    Write-Host " Enforcement     : Enforce" -ForegroundColor White
    if ($Comments) {
        Write-Host " Comments        : $Comments" -ForegroundColor White
    }
    if ($PSBoundParameters.ContainsKey('Priority')) {
        Write-Host " Priority        : $Priority" -ForegroundColor White
    }
    Write-Host "=======================================================" -ForegroundColor Cyan

    if (-not $ForceAction) {
        $confirm = Read-Host "`n[?] Do you want to create this rule in Exchange Online? (Y/N)"
        if ($confirm -notmatch '^(Y|y|T|t)$') {
            Write-Host "[-] Rule creation canceled by administrator." -ForegroundColor Gray
            return
        }
    }

    # Build parameter hashtable for New-TransportRule
    $ruleParams = @{
        Name        = $RuleName
        BlindCopyTo = $Recipients
        Mode        = 'Enforce'
        Confirm     = $false
        ErrorAction = 'Stop'
    }

    if ($Direction -eq 'Outgoing') {
        $ruleParams['From'] = $SourceMailbox
    } else {
        $ruleParams['SentTo'] = $SourceMailbox
    }

    if (-not [string]::IsNullOrWhiteSpace($Comments)) {
        $ruleParams['Comments'] = $Comments
    }

    if ($PSBoundParameters.ContainsKey('Priority')) {
        $ruleParams['Priority'] = $Priority
    }

    Write-Host "`n[*] Creating mail flow rule '$RuleName' in Exchange Online..." -ForegroundColor Cyan
    try {
        $newRule = New-TransportRule @ruleParams
        Write-Host "[+] Successfully created mail flow rule '$($newRule.Name)'!" -ForegroundColor Green

        Write-Host "`nRule Summary:" -ForegroundColor Yellow
        $newRule | Select-Object Name, State, Priority, Mode, Identity | Format-List | Out-String | Write-Host -ForegroundColor White
    } catch {
        Write-Error "Failed to create mail flow rule '$RuleName': $_"
        exit 1
    }
}

# Feature 5: Interactive Console Menu
function Start-InteractiveMenu {
    do {
        Write-Host "`n=======================================================" -ForegroundColor Cyan
        Write-Host "   Exchange Online Mail Flow Rules Manager (Interactive)" -ForegroundColor Cyan
        Write-Host "=======================================================" -ForegroundColor Cyan
        Write-Host " [1] List all mail flow rules (Summary Table)" -ForegroundColor Yellow
        Write-Host " [2] List all mail flow rules (Detailed Configuration)" -ForegroundColor Yellow
        Write-Host " [3] Enable a mail flow rule" -ForegroundColor Yellow
        Write-Host " [4] Disable a mail flow rule" -ForegroundColor Yellow
        Write-Host " [5] Create new BCC (UDW) redirection rule" -ForegroundColor Yellow
        Write-Host " [Q] Quit" -ForegroundColor DarkGray
        Write-Host "=======================================================" -ForegroundColor Cyan

        $choice = Read-Host "Select an option [1-5, Q]"
        if ([string]::IsNullOrWhiteSpace($choice)) { continue }

        switch ($choice.Trim().ToUpper()) {
            '1' {
                Get-MailflowRulesList -State 'All'
            }
            '2' {
                Get-MailflowRulesList -State 'All' -DetailedView
            }
            '3' {
                $target = Read-Host "`nEnter Rule Name or Identity/GUID to ENABLE"
                if (-not [string]::IsNullOrWhiteSpace($target)) {
                    Set-MailflowRuleState -Identity $target.Trim() -Action 'Enable'
                }
            }
            '4' {
                $target = Read-Host "`nEnter Rule Name or Identity/GUID to DISABLE"
                if (-not [string]::IsNullOrWhiteSpace($target)) {
                    Set-MailflowRuleState -Identity $target.Trim() -Action 'Disable'
                }
            }
            '5' {
                Write-Host "`n--- Create BCC / UDW Redirection Rule ---" -ForegroundColor Cyan
                $name = Read-Host "Enter unique Rule Name (e.g. 'BCC - Inbound Sales Monitoring')"
                $src = Read-Host "Enter Source Mailbox email (e.g. 'sales@domain.com')"
                $recips = Read-Host "Enter BCC (UDW) Recipient email(s) separated by comma (e.g. 'audit@domain.com, archive@domain.com')"
                $dirChoice = Read-Host "Traffic Direction [1 = Incoming (SentTo), 2 = Outgoing (From)] (Default: 1)"
                $dir = if ($dirChoice.Trim() -eq '2') { 'Outgoing' } else { 'Incoming' }
                $cmt = Read-Host "Enter optional comment/description (Press Enter to skip)"

                if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($src) -or [string]::IsNullOrWhiteSpace($recips)) {
                    Write-Warning "Rule Name, Source Mailbox, and BCC Recipients are required."
                } else {
                    $recipList = $recips -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                    # Validate before passing
                    if (-not (Test-ValidEmailAddress -Email $src.Trim())) {
                        Write-Warning "Source mailbox '$src' is not a valid email address."
                        continue
                    }
                    $allValid = $true
                    foreach ($r in $recipList) {
                        if (-not (Test-ValidEmailAddress -Email $r)) {
                            Write-Warning "Recipient '$r' is not a valid email address."
                            $allValid = $false
                            break
                        }
                    }
                    if ($allValid) {
                        New-BccMailflowRule -RuleName $name.Trim() -SourceMailbox $src.Trim() -Recipients $recipList -Direction $dir -Comments $cmt
                    }
                }
            }
            'Q' {
                Write-Host "`nExiting interactive mode." -ForegroundColor Cyan
                return
            }
            default {
                Write-Warning "Invalid selection. Please choose an option from 1 to 5, or 'Q'."
            }
        }
    } while ($true)
}

# Connect to Exchange Online
Connect-ExchangeEnvironment -AdminUPN $AdminUserPrincipalName

# Execution Entrypoint
if ($Interactive) {
    Start-InteractiveMenu
    exit 0
}

if ($ListRules -or ($PSCmdlet.ParameterSetName -eq 'ListRules' -and -not [string]::IsNullOrWhiteSpace($RuleIdentity))) {
    if (-not [string]::IsNullOrWhiteSpace($RuleIdentity)) {
        # Inspect specific rule
        Write-Host "`n[*] Inspecting rule: '$RuleIdentity'..." -ForegroundColor Cyan
        $targetRule = Find-TransportRule -Identity $RuleIdentity
        if ($targetRule) {
            $targetRule | Format-List * | Out-String | Write-Host -ForegroundColor White
        } else {
            Write-Error "Rule '$RuleIdentity' not found."
            exit 1
        }
    } else {
        Get-MailflowRulesList -State $StateFilter -FilterName $NameFilter -DetailedView:$Detailed
    }
    exit 0
}

if ($EnableRule) {
    Set-MailflowRuleState -Identity $RuleIdentity -Action 'Enable' -ForceAction:$Force
    exit 0
}

if ($DisableRule) {
    Set-MailflowRuleState -Identity $RuleIdentity -Action 'Disable' -ForceAction:$Force
    exit 0
}

if ($NewBccRule) {
    $newRuleArgs = @{
        RuleName      = $RuleName
        SourceMailbox = $SourceMailbox
        Recipients    = $cleanBccRecipients
        Direction     = $Direction
        ForceAction   = $Force
    }
    if (-not [string]::IsNullOrWhiteSpace($Comments)) {
        $newRuleArgs['Comments'] = $Comments
    }
    if ($PSBoundParameters.ContainsKey('Priority')) {
        $newRuleArgs['Priority'] = $Priority
    }
    New-BccMailflowRule @newRuleArgs
    exit 0
}
