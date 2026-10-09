<#
.SYNOPSIS
    Traces email message flow (Message Trace) and manages quarantine in Exchange Online.

.DESCRIPTION
    This tool allows you to:
    1. Trace inbound and outbound messages (Get-MessageTraceV2).
    2. Detect blocked, spam-filtered, or quarantined messages.
    3. Display the unique Quarantine Identity.
    4. Release/restore quarantined messages directly to the recipient mailbox.

.PARAMETER User
    Target mailbox email address (e.g., user@domain.com).

.PARAMETER Days
    Number of days back to search (1 to 10). Default: 7.

.PARAMETER Direction
    Mail flow direction: 'Inbound' or 'Outbound'. Default: 'Inbound'.

.PARAMETER OnlyBlockedOrSpam
    Filters results exclusively to blocked messages, spam, or quarantine.

.PARAMETER ReleaseId
    Quarantine message identifier (Identity from QuarantineId column) to release.

.PARAMETER Help
    Displays help and usage examples.

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -User "user@domain.com" -OnlyBlockedOrSpam

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -ReleaseId "c9e782e4-xxxx-xxxx-xxxx-xxxxxxxxxxxx\00000000-0000-0000-0000-000000000000"

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 2.0.1
#>

[CmdletBinding(DefaultParameterSetName = "Trace")]
param(
    [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Trace")]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$User,

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [ValidateRange(1, 10)]
    [int]$Days = 7,

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [ValidateSet("Inbound", "Outbound")]
    [string]$Direction = "Inbound",

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [switch]$OnlyBlockedOrSpam,

    [Parameter(Mandatory = $true, ParameterSetName = "Release")]
    [string]$ReleaseId,

    [Parameter(Mandatory = $false)]
    [Alias("h")]
    [switch]$Help
)

function Show-ScriptHelp {
    Clear-Host
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host "  M365 Message Trace & Quarantine Manager - v2.0.1" -ForegroundColor Cyan
    Write-Host "  Author: Roman Pindela (roman.pindela@gmail.com) | github.com/romanpindela" -ForegroundColor DarkGray
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Audits email traffic in Exchange Online and allows releasing messages from quarantine."
    Write-Host ""
    Write-Host "SYNTAX:" -ForegroundColor Yellow
    Write-Host "  Message Trace:" -ForegroundColor Green
    Write-Host "    .\Get-M365MessageTrace.ps1 -User <email> [-Days <1-10>] [-Direction <Inbound|Outbound>] [-OnlyBlockedOrSpam]"
    Write-Host "  Quarantine Release:" -ForegroundColor Green
    Write-Host "    .\Get-M365MessageTrace.ps1 -ReleaseId <QuarantineIdentity>"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -User               Target mailbox email address."
    Write-Host "  -Days               Number of days back (1-10). Default: 7."
    Write-Host "  -Direction          Direction: 'Inbound' or 'Outbound'. Default: 'Inbound'."
    Write-Host "  -OnlyBlockedOrSpam  Filters only failures, spam, and quarantine."
    Write-Host "  -ReleaseId          Releases quarantined message with specified Identity."
    Write-Host "  -Help, -h           Displays this help screen."
    Write-Host "=======================================================================" -ForegroundColor Cyan
}

if ($Help -or ($PSCmdlet.ParameterSetName -eq "Trace" -and [string]::IsNullOrWhiteSpace($User))) {
    Show-ScriptHelp
    Exit 0
}

# 1. Check Exchange Online connection
Write-Host "[*] Checking active Exchange Online session..." -ForegroundColor Cyan
$exoSession = Get-ConnectionInformation | Where-Object { $_.Name -like "*ExchangeOnline*" }

if (-not $exoSession) {
    Write-Host "[!] No active session found. Initiating administrator sign-in..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Error "[-] Exchange Online authentication failed: $_"
        Exit 1
    }
} else {
    Write-Host "[+] Existing Exchange Online session detected." -ForegroundColor Green
}

# 2. Handle action: RELEASE MESSAGE FROM QUARANTINE
if ($PSCmdlet.ParameterSetName -eq "Release") {
    Write-Host "[*] Attempting to release message from quarantine..." -ForegroundColor Cyan
    Write-Host "    Target Identity: $ReleaseId" -ForegroundColor Gray
    try {
        Release-QuarantineMessage -Identity $ReleaseId -ReleaseToAll -Confirm:$false -ErrorAction Stop
        Write-Host "[+] Success: Message was successfully released and delivered to recipient mailbox!" -ForegroundColor Green
    }
    catch {
        Write-Error "[-] Error releasing message: $_"
        Write-Host "[!] Verify the Identity is correct and that the message has not expired." -ForegroundColor Yellow
        Exit 1
    }
    Exit 0
}

# 3. Handle action: MESSAGE TRACE
$startDate = (Get-Date).AddDays(-$Days)
$endDate = Get-Date

$traceParams = @{
    StartDate = $startDate
    EndDate   = $endDate
}

if ($Direction -eq "Inbound") {
    $traceParams["RecipientAddress"] = $User
    $displayTarget = "SenderAddress"
    $targetHeader = "Sender"
} else {
    $traceParams["SenderAddress"] = $User
    $displayTarget = "RecipientAddress"
    $targetHeader = "Recipient"
}

Write-Host "[*] Retrieving message trace ($Direction) for '$User' (last $Days days)..." -ForegroundColor Cyan

try {
    $results = Get-MessageTraceV2 @traceParams -ErrorAction Stop
}
catch {
    Write-Error "[-] Error executing Get-MessageTraceV2: $_"
    Exit 1
}

if (-not $results -or $results.Count -eq 0) {
    Write-Host "[!] No entries found for $User in the specified period." -ForegroundColor Yellow
    Exit 0
}

# 4. Filter unwanted events / spam
if ($OnlyBlockedOrSpam) {
    Write-Host "[*] Filtering events: blocked, errors, spam, and quarantine..." -ForegroundColor Cyan
    $spamKeywords = @("Quarantined", "Failed", "FilteredAsSpam", "Blocked", "Spam")

    $results = $results | Where-Object {
        $status = $_.Status
        $matched = $false
        foreach ($k in $spamKeywords) {
            if ($status -like "*$k*") {
                $matched = $true
                break
            }
        }
        $matched
    }

    if (-not $results -or $results.Count -eq 0) {
        Write-Host "[+] Clean: No blocked messages or spam found." -ForegroundColor Green
        Exit 0
    }
}

# 5. Retrieve quarantine metadata for message correlation
$quarantineMap = @{}
if ($Direction -eq "Inbound") {
    Write-Host "[*] Checking EOP quarantine for recipient '$User'..." -ForegroundColor Cyan
    try {
        $quarantineItems = Get-QuarantineMessage -RecipientAddress $User -StartReceivedDate $startDate -EndReceivedDate $endDate -PageSize 1000 -ErrorAction SilentlyContinue
        if ($quarantineItems) {
            foreach ($q in $quarantineItems) {
                if ($q.MessageId) {
                    $quarantineMap[$q.MessageId] = $q.Identity
                }
                if ($q.NetworkMessageId) {
                    $quarantineMap[$q.NetworkMessageId] = $q.Identity
                }
            }
            Write-Host "[+] Found $($quarantineItems.Count) messages in quarantine." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "[!] Failed to retrieve quarantine details: $_"
    }
}

# 6. Results presentation
Write-Host "[+] Found $($results.Count) matching events:" -ForegroundColor Green

$index = 1
$formattedResults = foreach ($item in $results) {
    $qId = $null
    if ($item.MessageId -and $quarantineMap.ContainsKey($item.MessageId)) {
        $qId = $quarantineMap[$item.MessageId]
    } elseif ($item.NetworkMessageId -and $quarantineMap.ContainsKey($item.NetworkMessageId)) {
        $qId = $quarantineMap[$item.NetworkMessageId]
    }

    [PSCustomObject]@{
        "#"            = $index++
        "Received"     = $item.Received
        $targetHeader  = $item.$displayTarget
        "Subject"      = if ($item.Subject.Length -gt 45) { $item.Subject.Substring(0, 42) + "..." } else { $item.Subject }
        "Status"       = $item.Status
        "InQuarantine" = if ($qId) { "YES" } elseif ($item.Status -like "*Quarantined*") { "YES (search)" } else { "NO" }
        "QuarantineId" = if ($qId) { $qId } else { "-" }
    }
}

# Main table
$formattedResults | Format-Table -Property "#", "Received", $targetHeader, "Subject", "Status", "InQuarantine" -AutoSize

# 7. Display quarantined messages list
$releasable = $formattedResults | Where-Object { $_.QuarantineId -ne "-" }

if ($releasable) {
    Write-Host "
=======================================================================" -ForegroundColor Yellow
    Write-Host " MESSAGES AVAILABLE FOR RELEASE FROM QUARANTINE" -ForegroundColor Yellow
    Write-Host "=======================================================================" -ForegroundColor Yellow

    foreach ($r in $releasable) {
        Write-Host "[$($r.'#')] From: $($r.$targetHeader) | Subject: $($r.Subject)" -ForegroundColor White
        Write-Host "    Identity: $($r.QuarantineId)" -ForegroundColor DarkCyan
        Write-Host "    To release/restore, execute:" -ForegroundColor Gray
        Write-Host "    .\Get-M365MessageTrace.ps1 -ReleaseId `"$($r.QuarantineId)`"" -ForegroundColor Green
        Write-Host ""
    }
} else {
    Write-Host "
[i] No directly mapped items found in quarantine." -ForegroundColor Gray
    Write-Host "    If status is 'FilteredAsSpam', the message was delivered directly to the Junk Email folder in the user's mailbox." -ForegroundColor Gray
}
