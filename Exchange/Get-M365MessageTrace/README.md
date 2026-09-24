Akkuma qajeelfama keessaniitti, skriiptii kana maqaa **`Get-M365MessageTrace.ps1`** jedhuun, qajeelfama qulqullina koodii (CleanCode) fi ulaagaalee hunda guutee qophaa'ee jira.

---

### 1. Faayila Skriiptii: `Get-M365MessageTrace.ps1`

```powershell
<#
.SYNOPSIS
    Exchange Online (V2) keessatti dhangala'aa xalayaa (message trace) sakatta'a fi qorata.

.DESCRIPTION
    Skriiptiin kun ergamoota fi xalayaalee dhufan yookiin deeman sakatta'uuf gargaara.
    Xalayaalee cufaman (blocked), kwarantiinii seenan yookiin akka 'spam' lakkaa'aman 
    addaan baasuufis tajaajila kenna.

.PARAMETER User
    Teessoo imeelii fayyadamaa sakatta'amu (e.g., user@domain.com).

.PARAMETER Days
    Guyyoota meeqa duraa akka sakatta'amu (1 hanga 10 gidduutti). Qophii duraatiin: 7.

.PARAMETER Direction
    Kallattii xalayaa: 'Inbound' (kan dhufe) yookiin 'Outbound' (kan ergame). Qophii duraatiin: 'Inbound'.

.PARAMETER OnlyBlockedOrSpam
    Xalayaalee danqaman, cufaman yookiin akka spam/quarantine ta'an qofa calaluuf.

.PARAMETER Help
    Gargaarsa fi qajeelfama fayyadama skriiptichaa agarsiisa.

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl" -Days 5

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl" -Days 10 -OnlyBlockedOrSpam

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -Help

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$User,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$Days = 7,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Inbound", "Outbound")]
    [string]$Direction = "Inbound",

    [Parameter(Mandatory = $false)]
    [switch]$OnlyBlockedOrSpam,

    [Parameter(Mandatory = $false)]
    [Alias("h")]
    [switch]$Help
)

function Show-ScriptHelp {
    Clear-Host
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host "  M365 Message Trace Utility (Exchange Online) - v1.0.0" -ForegroundColor Cyan
    Write-Host "  Author: Roman Pindela (roman.pindela@gmail.com) | github.com/romanpindela" -ForegroundColor DarkGray
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Traces email messages for a specific Microsoft 365 mailbox using ExchangeOnlineManagement V2."
    Write-Host "  Allows filtering for inbound/outbound emails and targeted detection of blocked/spam/quarantined items."
    Write-Host ""
    Write-Host "SYNTAX:" -ForegroundColor Yellow
    Write-Host "  .\Get-M365MessageTrace.ps1 -User <email> [-Days <1-10>] [-Direction <Inbound|Outbound>] [-OnlyBlockedOrSpam]"
    Write-Host "  .\Get-M365MessageTrace.ps1 [-Help|-h]"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -User               Target mailbox email address (required to trace)."
    Write-Host "  -Days               Number of days to search back (1-10). Default: 7."
    Write-Host "  -Direction          Filter by message flow ('Inbound' or 'Outbound'). Default: 'Inbound'."
    Write-Host "  -OnlyBlockedOrSpam  Filter only messages that failed, were blocked, quarantined, or tagged as spam."
    Write-Host "  -Help, -h           Displays this interactive help manual."
    Write-Host ""
    Write-Host "AUTHENTICATION & PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "  - Requires 'ExchangeOnlineManagement' PowerShell module."
    Write-Host "  - The script checks active Exchange Online sessions; prompts for admin login if disconnected."
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\Get-M365MessageTrace.ps1 -User user@domain.com -Days 3"
    Write-Host "  .\Get-M365MessageTrace.ps1 -User user@domain.com -Direction Outbound -Days 10"
    Write-Host "  .\Get-M365MessageTrace.ps1 -User user@domain.com -OnlyBlockedOrSpam"
    Write-Host "=======================================================================" -ForegroundColor Cyan
}

if ($Help -or [string]::IsNullOrWhiteSpace($User)) {
    Show-ScriptHelp
    Exit 0
}

# 1. Checking Exchange Online Connection & Authenticating Admin
Write-Host "[*] Checking Exchange Online session state..." -ForegroundColor Cyan
$exoSession = Get-ConnectionInformation | Where-Object { $_.Name -like "*ExchangeOnline*" }

if (-not $exoSession) {
    Write-Host "[!] Active Exchange Online connection not found. Initiating administrator sign-in..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Error "[-] Failed to authenticate to Exchange Online: $_"
        Exit 1
    }
} else {
    Write-Host "[+] Existing Exchange Online session detected." -ForegroundColor Green
}

# 2. Date and Scope Definition
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

Write-Host "[*] Fetching $Direction message trace for '$User' (Last $Days days)..." -ForegroundColor Cyan

# 3. Querying Message Trace V2
try {
    $results = Get-MessageTraceV2 @traceParams -ErrorAction Stop
}
catch {
    Write-Error "[-] Error executing Get-MessageTraceV2: $_"
    Exit 1
}

if (-not $results -or $results.Count -eq 0) {
    Write-Host "[!] No message trace entries found for $User in the specified timeframe." -ForegroundColor Yellow
    Exit 0
}

# 4. Filter for Spam / Blocked / Quarantined
if ($OnlyBlockedOrSpam) {
    Write-Host "[*] Filtering for blocked, failed, spam, and quarantined events..." -ForegroundColor Cyan
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
        Write-Host "[+] Clean record: No blocked, failed, or spam messages found." -ForegroundColor Green
        Exit 0
    }
}

# 5. Output Results to Console
Write-Host "[+] Found $($results.Count) matching message trace event(s):" -ForegroundColor Green

$results | Select-Object `
    Received,
    @{Name = $targetHeader; Expression = { $_.$displayTarget }},
    Subject,
    Status,
    FromIP,
    Size,
    MessageId |
    Format-Table -AutoSize

```

---

### 2. Faayila README.md

```markdown
# M365 Message Trace Utility (`Get-M365MessageTrace.ps1`)

A robust, enterprise-grade PowerShell tool designed for Microsoft 365 administrators to inspect, trace, and audit email traffic in Exchange Online via the modern `Get-MessageTraceV2` interface.

---

## Author & Metadata

- **Author:** Roman Pindela
- **Email:** [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub:** [https://github.com/romanpindela](https://github.com/romanpindela)
- **Version:** `1.0.0`
- **Module Requirements:** `ExchangeOnlineManagement` (v3.0.0 or later)

---

## Features

- **Deprecation Resilient:** Utilizes `Get-MessageTraceV2` exclusively.
- **Directional Tracing:** Supports bidirectional message tracking (`Inbound` vs. `Outbound`).
- **Security Auditing:** Filter immediately for rejected, failed, blocked, quarantined, or spam-flagged deliveries using `-OnlyBlockedOrSpam`.
- **Automatic Session Verification:** Checks for existing Exchange Online sessions and initiates modern administrator authentication when required.
- **Safety & Input Sanitization:** Regex-validated email formats and bounded day ranges (1–10 days).
- **Default Help Mode:** Displays clean console help and syntax parameters when invoked without arguments.

---

## Prerequisites & Installation

### 1. Unblock Downloaded Script
If you clone or download this script directly from GitHub, Windows will mark it as untrusted. Unblock it before running:
```powershell
Unblock-File -Path .\Get-M365MessageTrace.ps1

```

### 2. Install Required Module

Make sure the Exchange Online management module is installed on your workstation:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force

```

---

## Usage Examples

### 1. Display Interactive Help

```powershell
.\Get-M365MessageTrace.ps1 -Help
# or
.\Get-M365MessageTrace.ps1 -h

```

### 2. Inspect Inbound Mail (Default: Last 7 Days)

```powershell
.\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl"

```

### 3. Check Messages Received in the Last 3 Days

```powershell
.\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl" -Days 3

```

### 4. Audit Outbound Emails

```powershell
.\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl" -Days 10 -Direction Outbound

```

### 5. Check for Blocked, Quarantined, or Spam Messages

```powershell
.\Get-M365MessageTrace.ps1 -User "anna.szywala@elektrimont.pl" -Days 10 -OnlyBlockedOrSpam

```

---

## License

This project is licensed under the MIT License - see the LICENSE file for details.



## Screenshots & Examples

### Standard run
![PowerShell Output](assets/Standard-run.jpg)

### Checking sent messages
![HTML Report](assets/Checking_sent_messages.jpg)

### Checking spam messages
![HTML Report](assets/Checking_spam_messages.jpg)

## Author

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](https://www.google.com/search?q=mailto%3Aroman.pindela%40gmail.com)
* **GitHub**: [romanpindela](https://github.com/romanpindela?utm_source=gemini)
* **Version**: 1.3.1