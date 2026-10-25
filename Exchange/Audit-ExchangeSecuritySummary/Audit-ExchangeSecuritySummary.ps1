<#
.SYNOPSIS
    Performs a non-intrusive, read-only security posture audit of Microsoft Exchange Online and generates an executive-ready HTML dashboard.

.DESCRIPTION
    Audit-ExchangeSecuritySummary is an enterprise-grade, non-intrusive (read-only) PowerShell auditing tool
    engineered to inspect Microsoft Exchange Online configuration against recognized security baselines (CIS Microsoft 365,
    CISA, and Microsoft Cloud Security Baselines).

    The tool assesses 5 critical defense-in-depth areas:
    1. Modern Authentication & Legacy Protocols:
       - Tenant-wide OAuth 2.0 Modern Authentication status.
       - SMTP Client Authentication (SMTP AUTH) transport configuration.
       - POP3 and IMAP4 legacy protocols on default CAS mailbox plans.
       - Exchange ActiveSync (EAS) legacy mobile protocol posture.
    2. Mail Routing & Exfiltration Protections:
       - External automatic mail forwarding posture on Remote Domains.
       - Outbound Anti-Spam auto-forwarding policy rules.
       - Anti-Malware policy common attachment filter and Zero-hour Auto Purge (ZAP).
       - Anti-Spam / Anti-Phishing policies and Zero-hour Auto Purge (ZAP).
       - Anti-Phishing spoof intelligence and First Contact Safety Tips.
    3. Auditing & Governance:
       - Tenant-wide mailbox auditing status (AuditDisabled).
       - Unified Audit Log (UAL) ingestion status.
       - Administrator audit logging status (AdminAuditLogEnabled).
       - Default mailbox audit action coverage for high-risk operations.
    4. Mailbox Delegation & Suspicious Rules:
       - Detection of external forwarding addresses on privileged mailboxes and user accounts.
       - Privileged mailbox delegation sprawl (FullAccess, SendAs, SendOnBehalf).
    5. Role-Based Access Control (RBAC):
       - Standing membership of the 'Organization Management' role group.
       - Security Administrator and Compliance Management role assignments.
       - Help Desk and Mail-Enabled Recipients role assignments.

    Output Capabilities:
    - Standardized ANSI console feedback ([OK], [WARN], [FAIL], [INFO]).
    - Interactive, standalone HTML report with pure embedded CSS and SVG icons (zero external CDN dependencies).
    - Executive metric KPI cards (Posture Score %, Checks Evaluated, Compliant, Warnings, Risks).
    - Searchable and filterable audit findings table with plain-English risk rationales.
    - Priority Remediation Roadmap sorted by severity (Critical first) with copy-pasteable PowerShell cmdlets.
    - Structured [PSCustomObject] pipeline emission for automated reporting.

.PARAMETER ReportPath
    The directory where the standalone HTML report will be generated.
    Default is 'C:\Logs'. If the directory does not exist, it will be automatically created.

.PARAMETER OpenReport
    Switch parameter. Automatically launches the generated HTML report in the default web browser upon completion.

.PARAMETER ShowDetailedOutput
    Switch parameter. Displays in-depth risk rationales and remediation commands directly in the console output.

.PARAMETER AdminUserPrincipalName
    Optional User Principal Name (UPN) of the administrator account to pre-fill when initiating an interactive Exchange Online login.

.PARAMETER Help
    Displays syntax help, architectural overview, parameter reference, usage examples, and exits cleanly.
    Aliases: -h.

.EXAMPLE
    # 1. Standard execution with default parameters (saves HTML report to C:\Logs)
    .\Audit-ExchangeSecuritySummary.ps1

.EXAMPLE
    # 2. Audit tenant and immediately open the executive HTML dashboard in default browser
    .\Audit-ExchangeSecuritySummary.ps1 -OpenReport

.EXAMPLE
    # 3. Specify custom report directory with detailed console output
    .\Audit-ExchangeSecuritySummary.ps1 -ReportPath "C:\SecOps\Reports" -ShowDetailedOutput -OpenReport

.EXAMPLE
    # 4. Pipeline integration: export all failed checks to CSV
    .\Audit-ExchangeSecuritySummary.ps1 | Where-Object Status -eq 'FAIL' | Export-Csv -Path "C:\Logs\FailedChecks.csv" -NoTypeInformation

.EXAMPLE
    # 5. Display built-in help documentation
    .\Audit-ExchangeSecuritySummary.ps1 -Help

.NOTES
    Author: Roman Pindela
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
    License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $false)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPositionalParameters', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath = "C:\Logs",

    [switch]$OpenReport,

    [switch]$ShowDetailedOutput,

    [string]$AdminUserPrincipalName,

    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Global Script State for Connection Lifecycle
$script:SessionWasActive = $false

# -----------------------------------------------------------------------------
# Function: Show-ScriptHelp
# -----------------------------------------------------------------------------
function Show-ScriptHelp {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    [CmdletBinding()]
    param()

    $banner = @"
================================================================================
   ___             _ _ _      ______          _
  / _ \           | (_) |     | ___ \        | |
 / /_\ \_   _  __| |_| |_ ___ | |_/ /___  ___| |__   __ _ _ __   __ _  ___
 |  _  | | | |/ _` | | __/ __||    // _ \/ __| '_ \ / _` | '_ \ / _` |/ _ \
 | | | | |_| | (_| | | |_\__ \| |\ \  __/ (__| | | | (_| | | | | (_| |  __/
 \_| |_/\__,_|\__,_|_|\__|___/\_| \_\___|\___|_| |_|\__,_|_| |_|\__, |\___|
                                                                  __/ |
 Microsoft 365 Exchange Online Security Posture Auditor          |___/
 Version: 1.0.0 | Author: Roman Pindela | License: MIT
================================================================================
"@
    Write-Host $banner -ForegroundColor Cyan

    $helpContent = @(
        "OVERVIEW:",
        "    Audit-ExchangeSecuritySummary evaluates tenant security settings against",
        "    recognized Exchange Online hardening baselines, provides plain-English risk",
        "    explanations, and exports a standalone HTML dashboard to C:\Logs.",
        "",
        "CORE AUDIT AREAS:",
        "    1. Modern Authentication & Legacy Protocols (OAuth 2.0, SMTP AUTH, POP3, IMAP4, EAS)",
        "    2. Mail Routing & Exfiltration Protections (Remote Domains, Outbound Spam, ZAP, Anti-Phish)",
        "    3. Auditing & Governance (Mailbox Auditing, Unified Audit Log, Admin Audit Log)",
        "    4. Mailbox Delegation & Suspicious Rules (External Forwarding, Privileged Delegation)",
        "    5. Role-Based Access Control (Organization Management, Compliance, Security Admins)",
        "",
        "PARAMETERS:",
        "    -ReportPath <string>       Directory for HTML report output (Default: 'C:\Logs').",
        "    -OpenReport                Automatically opens the HTML report in default browser.",
        "    -ShowDetailedOutput        Prints risk rationales and remediation commands to console.",
        "    -AdminUserPrincipalName    Administrator UPN for login pre-fill.",
        "    -Help, -h                  Displays this documentation and exits.",
        "",
        "PREREQUISITES & ROLES:",
        "    - PowerShell 5.1 or PowerShell 7+",
        "    - ExchangeOnlineManagement module (v3.0.0+)",
        "    - Microsoft 365 Role: Global Reader, Security Reader, or Exchange Administrator",
        "",
        "USAGE EXAMPLES:",
        "    .\Audit-ExchangeSecuritySummary.ps1",
        "    .\Audit-ExchangeSecuritySummary.ps1 -OpenReport",
        "    .\Audit-ExchangeSecuritySummary.ps1 -ReportPath 'C:\Security\Reports' -OpenReport",
        "    .\Audit-ExchangeSecuritySummary.ps1 | Where-Object Status -eq 'FAIL'",
        "================================================================================"
    )
    Write-Host ($helpContent -join "`n") -ForegroundColor Gray
}

# -----------------------------------------------------------------------------
# Early Exit for Help Parameter
# -----------------------------------------------------------------------------
if ($Help) {
    Show-ScriptHelp
    exit 0
}

# -----------------------------------------------------------------------------
# Console Logging Helpers
# -----------------------------------------------------------------------------
function Write-AuditBanner {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param([string]$Title)
    Write-Host "`n[*] $Title" -ForegroundColor Cyan
}

function Write-AuditPass {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param([string]$CheckId, [string]$Title, [string]$Detail)
    Write-Host "  [OK]   [$CheckId] $Title" -ForegroundColor Green
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host "         Value: $Detail" -ForegroundColor DarkGray
    }
}

function Write-AuditWarn {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param([string]$CheckId, [string]$Title, [string]$Detail)
    Write-Host "  [WARN] [$CheckId] $Title" -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host "         Value: $Detail" -ForegroundColor DarkGray
    }
}

function Write-AuditFail {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param([string]$CheckId, [string]$Title, [string]$Detail)
    Write-Host "  [FAIL] [$CheckId] $Title" -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host "         Value: $Detail" -ForegroundColor DarkGray
    }
}

function Write-AuditDetail {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
    param([string]$Rationale, [string]$Remediation)
    Write-Host "         Why: $Rationale" -ForegroundColor Gray
    Write-Host "         Fix: $Remediation" -ForegroundColor DarkCyan
}

# -----------------------------------------------------------------------------
# Function: Test-ExchangePrerequisite
# -----------------------------------------------------------------------------
function Test-ExchangePrerequisite {
    [CmdletBinding()]
    param()

    $installedModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

    if ($null -eq $installedModule) {
        Write-Host "[!] ExchangeOnlineManagement module is not installed." -ForegroundColor Yellow
        Write-Host "[*] Installing ExchangeOnlineManagement from PSGallery (CurrentUser)..." -ForegroundColor Cyan
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "[+] ExchangeOnlineManagement installed successfully." -ForegroundColor Green
        }
        catch {
            throw "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message). Please run: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
        }
    }
    else {
        if ($installedModule.Version -lt [System.Version]'3.0.0') {
            Write-Host "[!] Installed ExchangeOnlineManagement version is $($installedModule.Version). Version 3.0.0+ recommended." -ForegroundColor Yellow
        }
    }
}

# -----------------------------------------------------------------------------
# Function: Connect-ExchangeSession
# -----------------------------------------------------------------------------
function Connect-ExchangeSession {
    [CmdletBinding()]
    param([string]$AdminUPN)

    # 1. Inspect existing active connections
    $existingConnections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
    if ($existingConnections.Count -gt 0) {
        $script:SessionWasActive = $true
        Write-Host "[+] Reusing existing active Exchange Online session." -ForegroundColor Green
        return
    }

    # 2. Test Get-OrganizationConfig to verify whether session is already open
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        $script:SessionWasActive = $true
        Write-Host "[+] Active Exchange Online session confirmed." -ForegroundColor Green
        return
    }
    catch {
        Write-Verbose "No active session detected, initiating login: $($_.Exception.Message)"
    }

    Write-Host "[*] No active Exchange Online session detected. Initiating connection..." -ForegroundColor Cyan
    try {
        if (-not [string]::IsNullOrWhiteSpace($AdminUPN)) {
            Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        $script:SessionWasActive = $false
        Write-Host "[+] Connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        throw "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }
}

# -----------------------------------------------------------------------------
# Function: Disconnect-ExchangeSession
# -----------------------------------------------------------------------------
function Disconnect-ExchangeSession {
    [CmdletBinding()]
    param()

    if (-not $script:SessionWasActive) {
        Write-Host "`n[*] Disconnecting temporary Exchange Online session..." -ForegroundColor Cyan
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Host "[+] Disconnected cleanly." -ForegroundColor Green
        }
        catch {
            Write-Verbose "Disconnection error suppressed: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "`n[*] Preserving pre-existing Exchange Online session." -ForegroundColor DarkGray
    }
}

# -----------------------------------------------------------------------------
# Helper: New-AuditFindingObject
# -----------------------------------------------------------------------------
function New-AuditFindingObject {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param(
        [string]$Id,
        [string]$Category,
        [string]$Title,
        [ValidateSet('PASS', 'WARNING', 'FAIL')]
        [string]$Status,
        [ValidateSet('Critical', 'High', 'Medium', 'Low', 'Informational')]
        [string]$Severity,
        [string]$CurrentValue,
        [string]$BaselineValue,
        [string]$RiskRationale,
        [string]$Remediation,
        [string]$Details = ""
    )

    [PSCustomObject]@{
        Id            = $Id
        Category      = $Category
        Title         = $Title
        Status        = $Status
        Severity      = $Severity
        CurrentValue  = $CurrentValue
        BaselineValue = $BaselineValue
        RiskRationale = $RiskRationale
        Remediation   = $Remediation
        Details       = $Details
    }
}

# -----------------------------------------------------------------------------
# Function: Export-ExchangeSecurityHtmlReport
# -----------------------------------------------------------------------------
function Export-ExchangeSecurityHtmlReport {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$Findings,

        [Parameter(Mandatory = $true)]
        [string]$TenantName,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [double]$PostureScore,

        [Parameter(Mandatory = $true)]
        [hashtable]$Stats
    )

    $generatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME

    # Generate Finding Cards HTML
    $cardsHtml = [System.Text.StringBuilder]::new()
    foreach ($item in $Findings) {
        $statusClass = $item.Status.ToLower()
        $sevClass = $item.Severity.ToLower()

        $card = @"
        <div class="finding-card status-$statusClass severity-$sevClass" data-category="$([System.Web.HttpUtility]::HtmlEncode($item.Category))" data-status="$($item.Status)" data-severity="$($item.Severity)">
            <div class="card-header">
                <div class="header-left">
                    <span class="badge badge-$statusClass">$($item.Status)</span>
                    <span class="badge badge-sev badge-sev-$sevClass">$($item.Severity)</span>
                    <span class="finding-id">$($item.Id)</span>
                    <span class="category-tag">$([System.Web.HttpUtility]::HtmlEncode($item.Category))</span>
                </div>
            </div>
            <h3 class="finding-title">$([System.Web.HttpUtility]::HtmlEncode($item.Title))</h3>

            <div class="metric-comparison">
                <div class="metric-col">
                    <span class="metric-label">Current Setting</span>
                    <span class="metric-val current-val">$([System.Web.HttpUtility]::HtmlEncode($item.CurrentValue))</span>
                </div>
                <div class="metric-col">
                    <span class="metric-label">Security Baseline Target</span>
                    <span class="metric-val target-val">$([System.Web.HttpUtility]::HtmlEncode($item.BaselineValue))</span>
                </div>
            </div>

            <div class="section-box rationale-box">
                <div class="section-box-title">
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M11 15h2v2h-2zm0-8h2v6h-2zm1-5C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/></svg>
                    Why This Matters (Risk Analysis)
                </div>
                <p class="section-box-text">$([System.Web.HttpUtility]::HtmlEncode($item.RiskRationale))</p>
            </div>

            <div class="section-box remediation-box">
                <div class="section-box-title">
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M22.7 19l-9.1-9.1c.9-2.3.4-5-1.5-6.9-2-2-5-2.4-7.4-1.3L9 6 6 9 1.6 4.7C.4 7.1.9 10.1 2.9 12.1c1.9 1.9 4.6 2.4 6.9 1.5l9.1 9.1c.4.4 1 .4 1.4 0l2.3-2.3c.5-.4.5-1.1.1-1.4z"/></svg>
                    Remediation Action
                    <button class="copy-btn" onclick="copyCode(this)" data-code="$([System.Web.HttpUtility]::HtmlAttributeEncode($item.Remediation))">Copy Cmdlet</button>
                </div>
                <pre class="code-block"><code>$([System.Web.HttpUtility]::HtmlEncode($item.Remediation))</code></pre>
            </div>
        </div>
"@
        [void]$cardsHtml.AppendLine($card)
    }

    # Generate Priority Remediation Roadmap Rows
    $nonPassing = @($Findings | Where-Object { $_.Status -ne 'PASS' })
    $sevOrder = @{ 'Critical' = 1; 'High' = 2; 'Medium' = 3; 'Low' = 4; 'Informational' = 5 }
    $sortedNonPassing = @($nonPassing | Sort-Object -Property @{ Expression = { $sevOrder[$_.Severity] } })

    $roadmapHtml = [System.Text.StringBuilder]::new()
    $rank = 1
    foreach ($item in $sortedNonPassing) {
        $sevClass = $item.Severity.ToLower()
        $row = @"
        <tr class="roadmap-row sev-$sevClass">
            <td class="col-rank">#$rank</td>
            <td class="col-sev"><span class="badge badge-sev badge-sev-$sevClass">$($item.Severity)</span></td>
            <td class="col-id"><strong>$($item.Id)</strong></td>
            <td class="col-title">$([System.Web.HttpUtility]::HtmlEncode($item.Title))</td>
            <td class="col-cmd">
                <code class="mini-code">$([System.Web.HttpUtility]::HtmlEncode($item.Remediation))</code>
                <button class="copy-btn mini-copy" onclick="copyCode(this)" data-code="$([System.Web.HttpUtility]::HtmlAttributeEncode($item.Remediation))">Copy</button>
            </td>
        </tr>
"@
        [void]$roadmapHtml.AppendLine($row)
        $rank++
    }

    if ($sortedNonPassing.Count -eq 0) {
        $roadmapHtml.Append("<tr><td colspan='5' style='text-align:center; padding: 24px; color: #10b981; font-weight:600;'>All evaluated controls meet baseline standards. No immediate remediation actions required!</td></tr>")
    }

    $scoreColor = if ($PostureScore -ge 85.0) { '#10b981' } elseif ($PostureScore -ge 70.0) { '#f59e0b' } else { '#ef4444' }

    # Standalone HTML Template with embedded CSS & SVG (zero external CDN links)
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Online Security Posture Audit - $TenantName</title>
    <style>
        :root {
            --bg-page: #0f172a;
            --bg-surface: #1e293b;
            --bg-surface-elevated: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border-color: #334155;
            --pass-color: #10b981;
            --pass-bg: rgba(16, 185, 129, 0.15);
            --warn-color: #f59e0b;
            --warn-bg: rgba(245, 158, 11, 0.15);
            --fail-color: #ef4444;
            --fail-bg: rgba(239, 68, 68, 0.15);
            --info-color: #0ea5e9;
            --info-bg: rgba(14, 165, 233, 0.15);
            --crit-color: #ec4899;
            --crit-bg: rgba(236, 72, 153, 0.15);
            --accent-blue: #0078d4;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-page);
            color: var(--text-main);
            line-height: 1.5;
            padding-bottom: 60px;
        }

        header {
            background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
            border-bottom: 1px solid var(--border-color);
            padding: 32px 40px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.4);
        }

        .header-container {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .header-branding h1 {
            font-size: 26px;
            font-weight: 700;
            letter-spacing: -0.5px;
            display: flex;
            align-items: center;
            gap: 12px;
            color: #ffffff;
        }

        .header-branding h1 svg {
            width: 32px;
            height: 32px;
            fill: #38bdf8;
        }

        .header-subtitle {
            color: var(--text-muted);
            font-size: 14px;
            margin-top: 6px;
        }

        .header-meta {
            text-align: right;
            font-size: 13px;
            color: var(--text-muted);
            background: rgba(30, 41, 59, 0.7);
            padding: 12px 18px;
            border-radius: 8px;
            border: 1px solid var(--border-color);
        }

        .header-meta strong {
            color: #ffffff;
        }

        main {
            max-width: 1400px;
            margin: 32px auto 0 auto;
            padding: 0 40px;
        }

        /* Metric KPI Grid */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 18px;
            margin-bottom: 32px;
        }

        .kpi-card {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            transition: transform 0.2s ease, border-color 0.2s ease;
        }

        .kpi-card:hover {
            transform: translateY(-2px);
            border-color: #475569;
        }

        .kpi-title {
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .kpi-value {
            font-size: 36px;
            font-weight: 800;
            margin-top: 10px;
            color: #ffffff;
        }

        .kpi-sub {
            font-size: 12px;
            color: var(--text-muted);
            margin-top: 6px;
        }

        .score-val { color: $scoreColor; }
        .pass-val { color: var(--pass-color); }
        .warn-val { color: var(--warn-color); }
        .fail-val { color: var(--fail-color); }

        /* Filter Toolbar */
        .toolbar {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 16px 20px;
            margin-bottom: 24px;
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            align-items: center;
            justify-content: space-between;
        }

        .search-box {
            position: relative;
            flex: 1;
            min-width: 280px;
        }

        .search-box input {
            width: 100%;
            background: #0f172a;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 10px 16px 10px 38px;
            color: #ffffff;
            font-size: 14px;
            outline: none;
            transition: border-color 0.2s;
        }

        .search-box input:focus {
            border-color: #38bdf8;
        }

        .search-box svg {
            position: absolute;
            left: 12px;
            top: 11px;
            width: 18px;
            height: 18px;
            fill: var(--text-muted);
        }

        .filter-group {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .filter-btn {
            background: #0f172a;
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 8px 14px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
        }

        .filter-btn:hover {
            color: #ffffff;
            border-color: #475569;
        }

        .filter-btn.active {
            background: #2563eb;
            color: #ffffff;
            border-color: #3b82f6;
        }

        /* Findings Grid */
        .findings-list {
            display: flex;
            flex-direction: column;
            gap: 18px;
            margin-bottom: 40px;
        }

        .finding-card {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            transition: border-color 0.2s;
        }

        .finding-card.status-fail { border-left: 5px solid var(--fail-color); }
        .finding-card.status-warning { border-left: 5px solid var(--warn-color); }
        .finding-card.status-pass { border-left: 5px solid var(--pass-color); }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            flex-wrap: wrap;
            gap: 10px;
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }

        .badge {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            padding: 4px 10px;
            border-radius: 9999px;
            letter-spacing: 0.5px;
        }

        .badge-pass { background: var(--pass-bg); color: var(--pass-color); border: 1px solid rgba(16, 185, 129, 0.3); }
        .badge-warning { background: var(--warn-bg); color: var(--warn-color); border: 1px solid rgba(245, 158, 11, 0.3); }
        .badge-fail { background: var(--fail-bg); color: var(--fail-color); border: 1px solid rgba(239, 68, 68, 0.3); }

        .badge-sev-critical { background: var(--crit-bg); color: var(--crit-color); border: 1px solid rgba(236, 72, 153, 0.3); }
        .badge-sev-high { background: var(--fail-bg); color: var(--fail-color); border: 1px solid rgba(239, 68, 68, 0.3); }
        .badge-sev-medium { background: var(--warn-bg); color: var(--warn-color); border: 1px solid rgba(245, 158, 11, 0.3); }
        .badge-sev-low { background: var(--info-bg); color: var(--info-color); border: 1px solid rgba(14, 165, 233, 0.3); }
        .badge-sev-informational { background: rgba(148, 163, 184, 0.15); color: #94a3b8; border: 1px solid rgba(148, 163, 184, 0.3); }

        .finding-id {
            font-family: monospace;
            font-size: 13px;
            color: #cbd5e1;
            background: #0f172a;
            padding: 3px 8px;
            border-radius: 6px;
            border: 1px solid #334155;
        }

        .category-tag {
            font-size: 12px;
            color: var(--text-muted);
            background: rgba(51, 65, 85, 0.4);
            padding: 3px 8px;
            border-radius: 6px;
        }

        .finding-title {
            font-size: 18px;
            font-weight: 600;
            color: #ffffff;
            margin-bottom: 16px;
        }

        .metric-comparison {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            background: #0f172a;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 14px 18px;
            margin-bottom: 16px;
        }

        .metric-col {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .metric-label {
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
            color: var(--text-muted);
        }

        .metric-val {
            font-size: 14px;
            font-weight: 600;
            word-break: break-all;
        }

        .current-val { color: #f1f5f9; }
        .target-val { color: #38bdf8; }

        .section-box {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 14px 18px;
            margin-top: 12px;
        }

        .section-box-title {
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #94a3b8;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
        }

        .icon-inline {
            width: 16px;
            height: 16px;
            display: inline-block;
            vertical-align: middle;
            fill: #38bdf8;
        }

        .section-box-text {
            font-size: 13.5px;
            color: #cbd5e1;
            line-height: 1.6;
        }

        .code-block {
            background: #090d16;
            border: 1px solid #1e293b;
            border-radius: 6px;
            padding: 12px 14px;
            font-family: Consolas, Monaco, "Courier New", monospace;
            font-size: 13px;
            color: #38bdf8;
            overflow-x: auto;
            white-space: pre-wrap;
            word-break: break-all;
        }

        .copy-btn {
            background: #1e293b;
            border: 1px solid #475569;
            color: #e2e8f0;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 11px;
            cursor: pointer;
            transition: all 0.2s;
        }

        .copy-btn:hover {
            background: #334155;
            color: #ffffff;
            border-color: #64748b;
        }

        /* Roadmap Table */
        .roadmap-section {
            background: var(--bg-surface);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            margin-top: 32px;
        }

        .section-header-title {
            font-size: 20px;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .section-header-desc {
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 20px;
        }

        .roadmap-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        .roadmap-table th {
            text-align: left;
            padding: 12px 14px;
            background: #0f172a;
            color: var(--text-muted);
            border-bottom: 1px solid var(--border-color);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 11px;
        }

        .roadmap-table td {
            padding: 14px;
            border-bottom: 1px solid #293548;
            vertical-align: middle;
        }

        .col-rank { width: 50px; font-weight: 700; color: #94a3b8; }
        .col-sev { width: 110px; }
        .col-id { width: 130px; font-family: monospace; }
        .col-title { font-weight: 500; color: #f8fafc; }
        .col-cmd { width: 45%; }

        .mini-code {
            display: block;
            background: #090d16;
            padding: 6px 10px;
            border-radius: 4px;
            font-family: monospace;
            font-size: 12px;
            color: #38bdf8;
            margin-bottom: 6px;
            word-break: break-all;
        }

        .mini-copy {
            padding: 3px 8px;
            font-size: 10px;
        }

        footer {
            max-width: 1400px;
            margin: 40px auto 0 auto;
            padding: 0 40px;
            text-align: center;
            font-size: 13px;
            color: #64748b;
        }

        footer a {
            color: #38bdf8;
            text-decoration: none;
        }

        footer a:hover {
            text-decoration: underline;
        }

        /* Print Layout */
        @media print {
            body { background: #ffffff; color: #000000; padding-bottom: 0; }
            header { background: #ffffff; border-bottom: 2px solid #000000; box-shadow: none; padding: 20px 0; }
            .header-branding h1 { color: #000000; }
            .header-meta { background: #f8fafc; border: 1px solid #cbd5e1; }
            .toolbar { display: none; }
            .kpi-card { background: #ffffff; border: 1px solid #cbd5e1; }
            .finding-card { background: #ffffff; border: 1px solid #cbd5e1; page-break-inside: avoid; }
            .finding-title { color: #000000; }
            .code-block { background: #f1f5f9; color: #0f172a; border: 1px solid #cbd5e1; }
            .copy-btn { display: none; }
            .roadmap-section { background: #ffffff; border: 1px solid #cbd5e1; }
        }
    </style>
</head>
<body>

    <header>
        <div class="header-container">
            <div class="header-branding">
                <h1>
                    <svg viewBox="0 0 24 24"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm-2 16l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z"/></svg>
                    Exchange Online Security Posture Summary
                </h1>
                <div class="header-subtitle">
                    Non-intrusive baseline audit, threat exposure assessment, and executive remediation roadmap.
                </div>
            </div>
            <div class="header-meta">
                <div><strong>Tenant:</strong> $TenantName</div>
                <div><strong>Audited:</strong> $generatedDate</div>
                <div><strong>Host:</strong> $computerName | <strong>Auditor:</strong> $userName</div>
            </div>
        </div>
    </header>

    <main>
        <!-- KPI Overview Cards -->
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">
                    Overall Posture Score
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 14l-5-5 1.41-1.41L12 14.17l7.59-7.59L21 8l-9 9z"/></svg>
                </div>
                <div class="kpi-value score-val">$PostureScore%</div>
                <div class="kpi-sub">Weighted baseline alignment</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">
                    Checks Evaluated
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
                </div>
                <div class="kpi-value">$($Stats.Total)</div>
                <div class="kpi-sub">Total defense-in-depth rules</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">
                    Compliant (Pass)
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
                </div>
                <div class="kpi-value pass-val">$($Stats.Pass)</div>
                <div class="kpi-sub">Settings meeting baseline</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">
                    Review Required (Warn)
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>
                </div>
                <div class="kpi-value warn-val">$($Stats.Warn)</div>
                <div class="kpi-sub">Potential exposure points</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">
                    Security Risks (Fail)
                    <svg class="icon-inline" viewBox="0 0 24 24"><path fill="currentColor" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
                </div>
                <div class="kpi-value fail-val">$($Stats.Fail)</div>
                <div class="kpi-sub">$($Stats.Critical) Critical, $($Stats.High) High</div>
            </div>
        </div>

        <!-- Filter & Search Toolbar -->
        <div class="toolbar">
            <div class="search-box">
                <svg viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
                <input type="text" id="searchInput" placeholder="Filter by keyword, title, ID, or rationale..." onkeyup="filterCards()">
            </div>
            <div class="filter-group">
                <button class="filter-btn active" onclick="setStatusFilter('ALL', this)">All ($($Stats.Total))</button>
                <button class="filter-btn" onclick="setStatusFilter('FAIL', this)">Failures ($($Stats.Fail))</button>
                <button class="filter-btn" onclick="setStatusFilter('WARNING', this)">Warnings ($($Stats.Warn))</button>
                <button class="filter-btn" onclick="setStatusFilter('PASS', this)">Passed ($($Stats.Pass))</button>
            </div>
            <div class="filter-group">
                <button class="filter-btn" onclick="setSeverityFilter('Critical', this)">Critical ($($Stats.Critical))</button>
                <button class="filter-btn" onclick="setSeverityFilter('High', this)">High ($($Stats.High))</button>
                <button class="filter-btn" onclick="setSeverityFilter('ALL', this)">All Severities</button>
            </div>
        </div>

        <!-- Findings List -->
        <div class="findings-list" id="findingsList">
            $cardsHtml
        </div>

        <!-- Priority Remediation Roadmap -->
        <div class="roadmap-section">
            <div class="section-header-title">
                <svg class="icon-inline" style="width:22px; height:22px;" viewBox="0 0 24 24"><path fill="#38bdf8" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/></svg>
                Priority Remediation Roadmap
            </div>
            <div class="section-header-desc">
                Ordered action plan targeting identified security gaps, prioritized by exploitability and impact (Critical & High first).
            </div>
            <table class="roadmap-table">
                <thead>
                    <tr>
                        <th>Rank</th>
                        <th>Severity</th>
                        <th>ID</th>
                        <th>Finding & Objective</th>
                        <th>Recommended Action / PowerShell Cmdlet</th>
                    </tr>
                </thead>
                <tbody>
                    $roadmapHtml
                </tbody>
            </table>
        </div>
    </main>

    <footer>
        <p>Generated by <strong>Audit-ExchangeSecuritySummary v1.0.0</strong> | Author: <a href="https://github.com/romanpindela" target="_blank" rel="noopener">Roman Pindela</a> | License: MIT</p>
        <p style="margin-top:4px;">Non-intrusive Microsoft 365 Exchange Online Security Posture Auditor</p>
    </footer>

    <script>
        let currentStatus = 'ALL';
        let currentSeverity = 'ALL';

        function setStatusFilter(status, btn) {
            currentStatus = status;
            btn.parentElement.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filterCards();
        }

        function setSeverityFilter(sev, btn) {
            currentSeverity = sev;
            btn.parentElement.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filterCards();
        }

        function filterCards() {
            const query = document.getElementById('searchInput').value.toLowerCase();
            const cards = document.querySelectorAll('.finding-card');

            cards.forEach(card => {
                const text = card.textContent.toLowerCase();
                const cardStatus = card.getAttribute('data-status');
                const cardSeverity = card.getAttribute('data-severity');

                const matchesSearch = text.includes(query);
                const matchesStatus = (currentStatus === 'ALL' || cardStatus === currentStatus);
                const matchesSeverity = (currentSeverity === 'ALL' || cardSeverity === currentSeverity);

                if (matchesSearch && matchesStatus && matchesSeverity) {
                    card.style.display = 'block';
                } else {
                    card.style.display = 'none';
                }
            });
        }

        function copyCode(btn) {
            const code = btn.getAttribute('data-code');
            if (!code) return;
            navigator.clipboard.writeText(code).then(() => {
                const originalText = btn.innerText;
                btn.innerText = 'Copied!';
                btn.style.backgroundColor = '#10b981';
                btn.style.color = '#ffffff';
                setTimeout(() => {
                    btn.innerText = originalText;
                    btn.style.backgroundColor = '';
                    btn.style.color = '';
                }, 2000);
            }).catch(err => {
                console.error('Copy failed:', err);
            });
        }
    </script>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($FilePath, $htmlContent, [System.Text.UTF8Encoding]::new($false))
}

# -----------------------------------------------------------------------------
# Core Audit Routines
# -----------------------------------------------------------------------------

function Get-AuditModernAuth {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$FindingsList,
        [Parameter(Mandatory = $true)]
        $OrgConfig,
        [switch]$ShowDetails
    )

    Write-AuditBanner "Auditing Modern Authentication & Legacy Protocols..."

    # Check 1.1: OAuth2ClientProfileEnabled
    try {
        $modernAuthEnabled = $OrgConfig.OAuth2ClientProfileEnabled
        if ($modernAuthEnabled) {
            Write-AuditPass "EXO-AUTH-001" "Modern Authentication (OAuth 2.0)" "Enabled (OAuth2ClientProfileEnabled: True)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-001" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Modern Authentication Posture (OAuth 2.0)" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "Enabled (OAuth2ClientProfileEnabled: True)" `
                -BaselineValue "OAuth2ClientProfileEnabled: True" `
                -RiskRationale "Modern Authentication enables OAuth 2.0 token-based flows, conditional access, and multi-factor authentication (MFA). It eliminates password transmission in plain text and blocks basic auth spraying." `
                -Remediation "Compliant. Maintain OAuth2ClientProfileEnabled = $true." `
                -Details "Configured on OrganizationConfig."))
        }
        else {
            Write-AuditFail "EXO-AUTH-001" "Modern Authentication (OAuth 2.0)" "Disabled (OAuth2ClientProfileEnabled: False)"
            if ($ShowDetails) {
                Write-AuditDetail "Legacy basic auth leaves tenant vulnerable to credential stuffing and password spraying without MFA." "Set-OrganizationConfig -OAuth2ClientProfileEnabled `$true"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-001" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Modern Authentication Posture (OAuth 2.0)" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "Disabled (OAuth2ClientProfileEnabled: False)" `
                -BaselineValue "OAuth2ClientProfileEnabled: True" `
                -RiskRationale "Disabling Modern Authentication forces mail clients to authenticate using legacy basic authentication protocols without MFA support, rendering accounts vulnerable to automated password spraying." `
                -Remediation "Set-OrganizationConfig -OAuth2ClientProfileEnabled `$true" `
                -Details "OrganizationConfig"))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUTH-001" "Modern Authentication (OAuth 2.0)" "Could not query OrganizationConfig: $($_.Exception.Message)"
    }

    # Check 1.2: SmtpClientAuthenticationDisabled
    try {
        $transConfig = Get-TransportConfig -ErrorAction Stop
        $smtpAuthDisabled = $transConfig.SmtpClientAuthenticationDisabled

        if ($smtpAuthDisabled) {
            Write-AuditPass "EXO-AUTH-002" "SMTP Client Authentication (SMTP AUTH)" "Disabled globally (SmtpClientAuthenticationDisabled: True)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-002" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Default SMTP Client Authentication Posture" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "Disabled globally (SmtpClientAuthenticationDisabled: True)" `
                -BaselineValue "SmtpClientAuthenticationDisabled: True" `
                -RiskRationale "Disabling tenant-wide SMTP AUTH prevents compromised credentials from being used to authenticate via port 587 without MFA to relay phishing or spam." `
                -Remediation "Compliant. Keep SmtpClientAuthenticationDisabled = $true." `
                -Details "Configured on TransportConfig."))
        }
        else {
            Write-AuditWarn "EXO-AUTH-002" "SMTP Client Authentication (SMTP AUTH)" "Enabled globally (SmtpClientAuthenticationDisabled: False)"
            if ($ShowDetails) {
                Write-AuditDetail "SMTP AUTH allows basic auth authentication bypassing MFA for mail dispatch." "Set-TransportConfig -SmtpClientAuthenticationDisabled `$true"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-002" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Default SMTP Client Authentication Posture" `
                -Status "WARNING" `
                -Severity "High" `
                -CurrentValue "Enabled globally (SmtpClientAuthenticationDisabled: False)" `
                -BaselineValue "SmtpClientAuthenticationDisabled: True" `
                -RiskRationale "When SMTP AUTH is enabled tenant-wide, any account with a compromised password can be abused to send unauthorized emails from your domain without triggering MFA prompts." `
                -Remediation "Set-TransportConfig -SmtpClientAuthenticationDisabled `$true" `
                -Details "TransportConfig"))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUTH-002" "SMTP Client Authentication" "Could not query TransportConfig: $($_.Exception.Message)"
    }

    # Check 1.3 & 1.4: POP3 and IMAP4 on CASMailboxPlan
    try {
        $casPlans = @(Get-CASMailboxPlan -ErrorAction Stop)
        $popPlans = @($casPlans | Where-Object { $_.PopEnabled -eq $true })
        $imapPlans = @($casPlans | Where-Object { $_.ImapEnabled -eq $true })

        if ($popPlans.Count -eq 0) {
            Write-AuditPass "EXO-AUTH-003" "POP3 Protocol in CAS Mailbox Plans" "Disabled across all mailbox plans"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-003" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "POP3 Protocol in Default CAS Mailbox Plans" `
                -Status "PASS" `
                -Severity "Medium" `
                -CurrentValue "Disabled across all plans" `
                -BaselineValue "PopEnabled: False" `
                -RiskRationale "POP3 does not support bidirectional folder synchronization and exposes a legacy protocol endpoint. Disabling it on plans ensures new mailboxes are secured by default." `
                -Remediation "Compliant. Keep PopEnabled = $false on CAS plans." `
                -Details "Total plans checked: $($casPlans.Count)"))
        }
        else {
            $planNames = ($popPlans | ForEach-Object { $_.Name }) -join ', '
            Write-AuditWarn "EXO-AUTH-003" "POP3 Protocol in CAS Mailbox Plans" "Enabled on: $planNames"
            if ($ShowDetails) {
                Write-AuditDetail "Newly provisioned mailboxes inherit POP3 enabled status." "Get-CASMailboxPlan | Set-CASMailboxPlan -PopEnabled `$false"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-003" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "POP3 Protocol in Default CAS Mailbox Plans" `
                -Status "WARNING" `
                -Severity "Medium" `
                -CurrentValue "Enabled on: $planNames" `
                -BaselineValue "PopEnabled: False" `
                -RiskRationale "POP3 enables legacy mail retrieval without modern security controls. New mailboxes created under these plans will expose POP3 interfaces." `
                -Remediation "Get-CASMailboxPlan | Set-CASMailboxPlan -PopEnabled `$false" `
                -Details "Affected plans: $planNames"))
        }

        if ($imapPlans.Count -eq 0) {
            Write-AuditPass "EXO-AUTH-004" "IMAP4 Protocol in CAS Mailbox Plans" "Disabled across all mailbox plans"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-004" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "IMAP4 Protocol in Default CAS Mailbox Plans" `
                -Status "PASS" `
                -Severity "Medium" `
                -CurrentValue "Disabled across all plans" `
                -BaselineValue "ImapEnabled: False" `
                -RiskRationale "IMAP4 is a frequent target for credential harvesting and password spray attacks. Disabling it on mailbox plans prevents legacy protocol attack surfaces." `
                -Remediation "Compliant. Keep ImapEnabled = $false on CAS plans." `
                -Details "Total plans checked: $($casPlans.Count)"))
        }
        else {
            $planNames = ($imapPlans | ForEach-Object { $_.Name }) -join ', '
            Write-AuditWarn "EXO-AUTH-004" "IMAP4 Protocol in CAS Mailbox Plans" "Enabled on: $planNames"
            if ($ShowDetails) {
                Write-AuditDetail "Newly created mailboxes have IMAP4 enabled by default." "Get-CASMailboxPlan | Set-CASMailboxPlan -ImapEnabled `$false"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-004" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "IMAP4 Protocol in Default CAS Mailbox Plans" `
                -Status "WARNING" `
                -Severity "Medium" `
                -CurrentValue "Enabled on: $planNames" `
                -BaselineValue "ImapEnabled: False" `
                -RiskRationale "IMAP4 clients often bypass modern conditional access evaluations if legacy auth is not strictly blocked. Disabling IMAP4 at the plan level ensures secure mailbox defaults." `
                -Remediation "Get-CASMailboxPlan | Set-CASMailboxPlan -ImapEnabled `$false" `
                -Details "Affected plans: $planNames"))
        }

        # Check 1.5: ActiveSync on Plans
        $activeSyncPlans = @($casPlans | Where-Object { $_.ActiveSyncEnabled -eq $true })
        if ($activeSyncPlans.Count -gt 0) {
            $planNames = ($activeSyncPlans | ForEach-Object { $_.Name }) -join ', '
            Write-AuditWarn "EXO-AUTH-005" "Exchange ActiveSync (EAS) Protocol Posture" "ActiveSync is enabled on: $planNames"
            if ($ShowDetails) {
                Write-AuditDetail "Unmanaged EAS mobile clients can bypass MAM compliance unless restricted via Intune." "Enforce Intune App Protection policies or disable EAS: Get-CASMailboxPlan | Set-CASMailboxPlan -ActiveSyncEnabled `$false"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-005" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Exchange ActiveSync (EAS) Legacy Mobile Access Posture" `
                -Status "WARNING" `
                -Severity "Low" `
                -CurrentValue "Enabled on: $planNames" `
                -BaselineValue "ActiveSyncEnabled: False (or controlled via Intune MAM)" `
                -RiskRationale "Exchange ActiveSync can be utilized by legacy native mobile mail apps that lack modern Conditional Access compliance checks or Intune Mobile Application Management (MAM) protections." `
                -Remediation "Enforce Outlook Mobile with Intune MAM and optionally disable EAS on plans: Get-CASMailboxPlan | Set-CASMailboxPlan -ActiveSyncEnabled `$false" `
                -Details "Plans with EAS: $planNames"))
        }
        else {
            Write-AuditPass "EXO-AUTH-005" "Exchange ActiveSync (EAS) Protocol Posture" "Disabled across all mailbox plans"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUTH-005" `
                -Category "Modern Authentication & Legacy Protocols" `
                -Title "Exchange ActiveSync (EAS) Legacy Mobile Access Posture" `
                -Status "PASS" `
                -Severity "Low" `
                -CurrentValue "Disabled across all plans" `
                -BaselineValue "ActiveSyncEnabled: False" `
                -RiskRationale "ActiveSync is disabled by default, ensuring native unmanaged mobile clients cannot connect without modern app controls." `
                -Remediation "Compliant. Maintain ActiveSync disabled or managed." `
                -Details "All plans compliant."))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUTH-003" "CAS Mailbox Plans" "Could not query CASMailboxPlan: $($_.Exception.Message)"
    }
}

function Get-AuditMailRouting {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$FindingsList,
        [switch]$ShowDetails
    )

    Write-AuditBanner "Auditing Mail Routing & Exfiltration Protections..."

    # Check 2.1: Remote Domains AutoForwardEnabled
    try {
        $remoteDomains = @(Get-RemoteDomain -ErrorAction Stop)
        $autoForwardDomains = @($remoteDomains | Where-Object { $_.AutoForwardEnabled -eq $true })

        if ($autoForwardDomains.Count -eq 0) {
            Write-AuditPass "EXO-ROUT-001" "External Automatic Mail Forwarding (RemoteDomains)" "Disabled on all remote domains"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-001" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "External Automatic Mail Forwarding on Remote Domains" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "AutoForwardEnabled: False across all remote domains" `
                -BaselineValue "AutoForwardEnabled: False" `
                -RiskRationale "Cybercriminals routinely configure hidden auto-forwarding rules on compromised mailboxes to exfiltrate proprietary communications. Blocking auto-forwarding on remote domains stops exfiltration at the transport perimeter." `
                -Remediation "Compliant. Maintain AutoForwardEnabled = $false on all RemoteDomains." `
                -Details "Domains checked: $($remoteDomains.Count)"))
        }
        else {
            $domainNames = ($autoForwardDomains | ForEach-Object { $_.DomainName }) -join ', '
            Write-AuditFail "EXO-ROUT-001" "External Automatic Mail Forwarding (RemoteDomains)" "Allowed on: $domainNames"
            if ($ShowDetails) {
                Write-AuditDetail "Attackers can set mailbox rules that silently copy emails to external inboxes." "Set-RemoteDomain -Identity 'Default' -AutoForwardEnabled `$false"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-001" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "External Automatic Mail Forwarding on Remote Domains" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "AutoForwardEnabled: True on domain(s): $domainNames" `
                -BaselineValue "AutoForwardEnabled: False" `
                -RiskRationale "Allowing automatic external forwarding enables threat actors to configure silent data exfiltration rules, exposing sensitive corporate data and executive correspondence to external drop boxes." `
                -Remediation "Get-RemoteDomain | Where-Object AutoForwardEnabled | Set-RemoteDomain -AutoForwardEnabled `$false" `
                -Details "Vulnerable remote domains: $domainNames"))
        }
    }
    catch {
        Write-AuditWarn "EXO-ROUT-001" "Remote Domains" "Could not query RemoteDomain: $($_.Exception.Message)"
    }

    # Check 2.2: Outbound Anti-Spam AutoForwardingMode
    try {
        $outboundPolicies = @(Get-HostedOutboundSpamFilterPolicy -ErrorAction Stop)
        $unrestrictedPolicies = @($outboundPolicies | Where-Object { $_.AutoForwardingMode -eq 'On' })
        $automaticPolicies = @($outboundPolicies | Where-Object { $_.AutoForwardingMode -eq 'Automatic' })

        if ($unrestrictedPolicies.Count -gt 0) {
            $names = ($unrestrictedPolicies | ForEach-Object { $_.Name }) -join ', '
            Write-AuditFail "EXO-ROUT-002" "Outbound Anti-Spam Auto-Forwarding Controls" "AutoForwardingMode is set to 'On' on: $names"
            if ($ShowDetails) {
                Write-AuditDetail "Outbound spam policy explicitly allows users to forward emails externally." "Get-HostedOutboundSpamFilterPolicy | Set-HostedOutboundSpamFilterPolicy -AutoForwardingMode Off"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-002" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Outbound Anti-Spam Auto-Forwarding Controls" `
                -Status "FAIL" `
                -Severity "High" `
                -CurrentValue "AutoForwardingMode: On (Allowed) on policy: $names" `
                -BaselineValue "AutoForwardingMode: Off" `
                -RiskRationale "When Outbound Anti-Spam auto-forwarding is explicitly set to 'On', external forwarding is permitted, bypassing default tenant protections and enabling data leakage." `
                -Remediation "Get-HostedOutboundSpamFilterPolicy | Set-HostedOutboundSpamFilterPolicy -AutoForwardingMode Off" `
                -Details "Policies with forwarding enabled: $names"))
        }
        elseif ($automaticPolicies.Count -gt 0) {
            $names = ($automaticPolicies | ForEach-Object { $_.Name }) -join ', '
            Write-AuditWarn "EXO-ROUT-002" "Outbound Anti-Spam Auto-Forwarding Controls" "AutoForwardingMode is set to 'Automatic' (System-controlled) on: $names"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-002" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Outbound Anti-Spam Auto-Forwarding Controls" `
                -Status "WARNING" `
                -Severity "Medium" `
                -CurrentValue "AutoForwardingMode: Automatic (System-controlled)" `
                -BaselineValue "AutoForwardingMode: Off (Explicit hard block)" `
                -RiskRationale "Automatic mode delegates forwarding decisions to Microsoft system intelligence. CIS Security baselines mandate an explicit 'Off' setting to prevent unauthorized forwarding under all conditions." `
                -Remediation "Get-HostedOutboundSpamFilterPolicy | Set-HostedOutboundSpamFilterPolicy -AutoForwardingMode Off" `
                -Details "Configured on: $names"))
        }
        else {
            Write-AuditPass "EXO-ROUT-002" "Outbound Anti-Spam Auto-Forwarding Controls" "AutoForwardingMode is set to 'Off' across all policies"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-002" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Outbound Anti-Spam Auto-Forwarding Controls" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "AutoForwardingMode: Off" `
                -BaselineValue "AutoForwardingMode: Off" `
                -RiskRationale "External automatic mail forwarding is strictly blocked at the outbound filter level, preventing data exfiltration." `
                -Remediation "Compliant. Maintain AutoForwardingMode = Off." `
                -Details "All policies configured to Off."))
        }
    }
    catch {
        Write-AuditWarn "EXO-ROUT-002" "Outbound Anti-Spam Policy" "Could not query HostedOutboundSpamFilterPolicy: $($_.Exception.Message)"
    }

    # Check 2.3 & 2.4: Anti-Malware Common Attachment Filter & ZAP
    try {
        $malwarePolicies = @(Get-MalwareFilterPolicy -ErrorAction Stop)
        $noFileFilter = @($malwarePolicies | Where-Object { $_.EnableFileFilter -ne $true })
        $noZap = @($malwarePolicies | Where-Object { $_.ZapEnabled -ne $true })

        if ($noFileFilter.Count -eq 0) {
            Write-AuditPass "EXO-ROUT-003" "Anti-Malware Common Attachment Filter" "Enabled across all anti-malware policies"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-003" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Anti-Malware Common Attachment Types Filter" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "EnableFileFilter: True on all policies" `
                -BaselineValue "EnableFileFilter: True" `
                -RiskRationale "Blocking high-risk executable, script, and container file extensions (.exe, .scr, .iso, .vbs, .js, .bat) stops common initial access malware payloads." `
                -Remediation "Compliant. Maintain EnableFileFilter = $true." `
                -Details "Policies evaluated: $($malwarePolicies.Count)"))
        }
        else {
            $names = ($noFileFilter | ForEach-Object { $_.Name }) -join ', '
            Write-AuditFail "EXO-ROUT-003" "Anti-Malware Common Attachment Filter" "Disabled on: $names"
            if ($ShowDetails) {
                Write-AuditDetail "Executable attachments can bypass filters if common attachment filter is disabled." "Get-MalwareFilterPolicy | Set-MalwareFilterPolicy -EnableFileFilter `$true"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-003" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Anti-Malware Common Attachment Types Filter" `
                -Status "FAIL" `
                -Severity "High" `
                -CurrentValue "EnableFileFilter: False on: $names" `
                -BaselineValue "EnableFileFilter: True" `
                -RiskRationale "High-risk executable and script attachments are not blocked by default, leaving users exposed to macro viruses, script loaders, and ransomware droppers." `
                -Remediation "Get-MalwareFilterPolicy | Set-MalwareFilterPolicy -EnableFileFilter `$true" `
                -Details "Policies missing filter: $names"))
        }

        if ($noZap.Count -eq 0) {
            Write-AuditPass "EXO-ROUT-004" "Zero-Hour Auto Purge (ZAP) for Malware" "Enabled across all anti-malware policies"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-004" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Zero-Hour Auto Purge (ZAP) for Malware" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "ZapEnabled: True on all policies" `
                -BaselineValue "ZapEnabled: True" `
                -RiskRationale "Zero-hour Auto Purge continuously retroactively neutralizes and isolates delivered malware when threat signatures are identified post-delivery." `
                -Remediation "Compliant. Maintain ZapEnabled = $true." `
                -Details "All policies compliant."))
        }
        else {
            $names = ($noZap | ForEach-Object { $_.Name }) -join ', '
            Write-AuditFail "EXO-ROUT-004" "Zero-Hour Auto Purge (ZAP) for Malware" "Disabled on: $names"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-004" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Zero-Hour Auto Purge (ZAP) for Malware" `
                -Status "FAIL" `
                -Severity "High" `
                -CurrentValue "ZapEnabled: False on: $names" `
                -BaselineValue "ZapEnabled: True" `
                -RiskRationale "Without ZAP for malware, messages with newly classified malware variants remain sitting in user mailboxes, waiting to be executed by victims." `
                -Remediation "Get-MalwareFilterPolicy | Set-MalwareFilterPolicy -ZapEnabled `$true" `
                -Details "Policies with ZAP disabled: $names"))
        }
    }
    catch {
        Write-AuditWarn "EXO-ROUT-003" "Anti-Malware Policy" "Could not query MalwareFilterPolicy: $($_.Exception.Message)"
    }

    # Check 2.5: Zero-Hour Auto Purge (ZAP) for Phishing & Spam
    try {
        $contentPolicies = @(Get-HostedContentFilterPolicy -ErrorAction Stop)
        $noPhishZap = @($contentPolicies | Where-Object { $_.EnablePhishZap -ne $true -or $_.EnableZap -ne $true })

        if ($noPhishZap.Count -eq 0) {
            Write-AuditPass "EXO-ROUT-005" "Zero-Hour Auto Purge (ZAP) for Phishing & Spam" "Enabled across all content filter policies"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-005" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Zero-Hour Auto Purge (ZAP) for Phishing & Spam" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "EnableZap & EnablePhishZap: True on all policies" `
                -BaselineValue "EnableZap: True, EnablePhishZap: True" `
                -RiskRationale "Phish ZAP retroactively quarantines phishing emails that evaded initial perimeter scanning once Microsoft threat intelligence identifies weaponized URLs." `
                -Remediation "Compliant. Maintain ZAP enabled on all content filter policies." `
                -Details "Policies checked: $($contentPolicies.Count)"))
        }
        else {
            $names = ($noPhishZap | ForEach-Object { $_.Name }) -join ', '
            Write-AuditFail "EXO-ROUT-005" "Zero-Hour Auto Purge (ZAP) for Phishing & Spam" "Disabled or incomplete on: $names"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-ROUT-005" `
                -Category "Mail Routing & Exfiltration Protections" `
                -Title "Zero-Hour Auto Purge (ZAP) for Phishing & Spam" `
                -Status "FAIL" `
                -Severity "High" `
                -CurrentValue "ZAP for Phishing/Spam is disabled or partially configured on: $names" `
                -BaselineValue "EnableZap: True, EnablePhishZap: True" `
                -RiskRationale "Phishing campaigns that evade boundary filters will persist indefinitely in user inboxes without retroactive purging." `
                -Remediation "Get-HostedContentFilterPolicy | Set-HostedContentFilterPolicy -EnableZap `$true -EnablePhishZap `$true -EnableHighConfidencePhishZap `$true" `
                -Details "Policies affected: $names"))
        }
    }
    catch {
        Write-AuditWarn "EXO-ROUT-005" "Content Filter Policy" "Could not query HostedContentFilterPolicy: $($_.Exception.Message)"
    }

    # Check 2.6 & 2.7: Anti-Phishing Policies (Spoof Intelligence & Safety Tips)
    try {
        $antiPhishPolicies = @(Get-AntiPhishPolicy -ErrorAction SilentlyContinue)
        if ($antiPhishPolicies.Count -gt 0) {
            $noSpoof = @($antiPhishPolicies | Where-Object { $_.EnableSpoofIntelligence -ne $true })
            $noSafetyTips = @($antiPhishPolicies | Where-Object { $_.EnableFirstContactSafetyTips -ne $true })

            if ($noSpoof.Count -eq 0) {
                Write-AuditPass "EXO-ROUT-006" "Anti-Phishing Spoof Intelligence" "Enabled across all anti-phish policies"
                $FindingsList.Add((New-AuditFindingObject `
                    -Id "EXO-ROUT-006" `
                    -Category "Mail Routing & Exfiltration Protections" `
                    -Title "Anti-Phishing Spoof Intelligence" `
                    -Status "PASS" `
                    -Severity "High" `
                    -CurrentValue "EnableSpoofIntelligence: True" `
                    -BaselineValue "EnableSpoofIntelligence: True" `
                    -RiskRationale "Spoof intelligence verifies sender authenticity (SPF, DKIM, DMARC) and user reputation models to prevent sender address spoofing." `
                    -Remediation "Compliant. Maintain EnableSpoofIntelligence = $true." `
                    -Details "Policies checked: $($antiPhishPolicies.Count)"))
            }
            else {
                $names = ($noSpoof | ForEach-Object { $_.Name }) -join ', '
                Write-AuditFail "EXO-ROUT-006" "Anti-Phishing Spoof Intelligence" "Disabled on: $names"
                $FindingsList.Add((New-AuditFindingObject `
                    -Id "EXO-ROUT-006" `
                    -Category "Mail Routing & Exfiltration Protections" `
                    -Title "Anti-Phishing Spoof Intelligence" `
                    -Status "FAIL" `
                    -Severity "High" `
                    -CurrentValue "EnableSpoofIntelligence: False on: $names" `
                    -BaselineValue "EnableSpoofIntelligence: True" `
                    -RiskRationale "Without spoof intelligence, external adversaries can impersonate internal domains and trusted partners to conduct executive fraud." `
                    -Remediation "Get-AntiPhishPolicy | Set-AntiPhishPolicy -EnableSpoofIntelligence `$true" `
                    -Details "Policies affected: $names"))
            }

            if ($noSafetyTips.Count -eq 0) {
                Write-AuditPass "EXO-ROUT-007" "Anti-Phishing First Contact Safety Tips" "Enabled across all anti-phish policies"
                $FindingsList.Add((New-AuditFindingObject `
                    -Id "EXO-ROUT-007" `
                    -Category "Mail Routing & Exfiltration Protections" `
                    -Title "Anti-Phishing First Contact Safety Tips" `
                    -Status "PASS" `
                    -Severity "Medium" `
                    -CurrentValue "EnableFirstContactSafetyTips: True" `
                    -BaselineValue "EnableFirstContactSafetyTips: True" `
                    -RiskRationale "Warns users when an email arrives from a sender they rarely or never interact with, curbing business email compromise (BEC)." `
                    -Remediation "Compliant. Maintain EnableFirstContactSafetyTips = $true." `
                    -Details "All policies compliant."))
            }
            else {
                $names = ($noSafetyTips | ForEach-Object { $_.Name }) -join ', '
                Write-AuditWarn "EXO-ROUT-007" "Anti-Phishing First Contact Safety Tips" "Disabled on: $names"
                $FindingsList.Add((New-AuditFindingObject `
                    -Id "EXO-ROUT-007" `
                    -Category "Mail Routing & Exfiltration Protections" `
                    -Title "Anti-Phishing First Contact Safety Tips" `
                    -Status "WARNING" `
                    -Severity "Medium" `
                    -CurrentValue "EnableFirstContactSafetyTips: False on: $names" `
                    -BaselineValue "EnableFirstContactSafetyTips: True" `
                    -RiskRationale "Users do not receive visual safety banners on novel sender interactions, increasing susceptibility to typo-squatted domains." `
                    -Remediation "Get-AntiPhishPolicy | Set-AntiPhishPolicy -EnableFirstContactSafetyTips `$true" `
                    -Details "Policies affected: $names"))
            }
        }
        else {
            Write-AuditWarn "EXO-ROUT-006" "Anti-Phishing Policies" "No custom anti-phish policies retrieved or Defender for Office 365 license not present."
        }
    }
    catch {
        Write-AuditWarn "EXO-ROUT-006" "Anti-Phishing Policy" "Could not query AntiPhishPolicy: $($_.Exception.Message)"
    }
}

function Get-AuditGovernance {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$FindingsList,
        [Parameter(Mandatory = $true)]
        $OrgConfig,
        [switch]$ShowDetails
    )

    Write-AuditBanner "Auditing Governance & Audit Logging..."

    # Check 3.1: Organization Mailbox Auditing
    try {
        $auditDisabled = $OrgConfig.AuditDisabled
        if (-not $auditDisabled) {
            Write-AuditPass "EXO-AUDIT-001" "Organization Mailbox Auditing" "Enabled tenant-wide (AuditDisabled: False)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-001" `
                -Category "Auditing & Governance" `
                -Title "Organization-Wide Mailbox Auditing Posture" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "Enabled tenant-wide (AuditDisabled: False)" `
                -BaselineValue "AuditDisabled: False" `
                -RiskRationale "Mailbox auditing records actions taken by owners, delegates, and administrators, providing critical forensic evidence during incident response." `
                -Remediation "Compliant. Maintain AuditDisabled = $false." `
                -Details "Tenant-wide setting verified."))
        }
        else {
            Write-AuditFail "EXO-AUDIT-001" "Organization Mailbox Auditing" "Disabled tenant-wide (AuditDisabled: True)"
            if ($ShowDetails) {
                Write-AuditDetail "Mailbox audit actions are not recorded across the organization." "Set-OrganizationConfig -AuditDisabled `$false"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-001" `
                -Category "Auditing & Governance" `
                -Title "Organization-Wide Mailbox Auditing Posture" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "Disabled tenant-wide (AuditDisabled: True)" `
                -BaselineValue "AuditDisabled: False" `
                -RiskRationale "Disabling mailbox auditing blindfolds security analysts, making it impossible to determine if compromised mailboxes were accessed, searched, or exported." `
                -Remediation "Set-OrganizationConfig -AuditDisabled `$false" `
                -Details "OrganizationConfig"))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUDIT-001" "Mailbox Auditing" "Could not verify OrganizationConfig audit status: $($_.Exception.Message)"
    }

    # Check 3.2: Unified Audit Log Ingestion Status
    try {
        $adminAuditConfig = Get-AdminAuditLogConfig -ErrorAction Stop
        $ualIngestion = $adminAuditConfig.UnifiedAuditLogIngestionEnabled

        if ($ualIngestion) {
            Write-AuditPass "EXO-AUDIT-002" "Unified Audit Log (UAL) Ingestion" "Enabled (UnifiedAuditLogIngestionEnabled: True)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-002" `
                -Category "Auditing & Governance" `
                -Title "Unified Audit Log (UAL) Ingestion Status" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "Enabled (UnifiedAuditLogIngestionEnabled: True)" `
                -BaselineValue "UnifiedAuditLogIngestionEnabled: True" `
                -RiskRationale "The Unified Audit Log centralizes audit events across Exchange Online, Entra ID, SharePoint, and Teams for forensic ingestion and SIEM integration." `
                -Remediation "Compliant. Maintain UnifiedAuditLogIngestionEnabled = $true." `
                -Details "AdminAuditLogConfig"))
        }
        else {
            Write-AuditFail "EXO-AUDIT-002" "Unified Audit Log (UAL) Ingestion" "Disabled (UnifiedAuditLogIngestionEnabled: False)"
            if ($ShowDetails) {
                Write-AuditDetail "Audit events are not recorded to the Microsoft 365 Unified Audit Log." "Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-002" `
                -Category "Auditing & Governance" `
                -Title "Unified Audit Log (UAL) Ingestion Status" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "Disabled (UnifiedAuditLogIngestionEnabled: False)" `
                -BaselineValue "UnifiedAuditLogIngestionEnabled: True" `
                -RiskRationale "Without Unified Audit Log ingestion, user access events, administrator changes, and exfiltration activities are permanently discarded." `
                -Remediation "Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true" `
                -Details "AdminAuditLogConfig"))
        }

        # Check 3.3: Administrator Audit Logging
        $adminAuditEnabled = $adminAuditConfig.AdminAuditLogEnabled
        if ($adminAuditEnabled) {
            Write-AuditPass "EXO-AUDIT-003" "Administrator Audit Logging" "Enabled (AdminAuditLogEnabled: True)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-003" `
                -Category "Auditing & Governance" `
                -Title "Administrator Audit Logging Posture" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "Enabled (AdminAuditLogEnabled: True)" `
                -BaselineValue "AdminAuditLogEnabled: True" `
                -RiskRationale "Records administrative cmdlets executed via PowerShell or EAC, creating a tamper-evident audit trail for regulatory compliance." `
                -Remediation "Compliant. Maintain AdminAuditLogEnabled = $true." `
                -Details "AdminAuditLogConfig"))
        }
        else {
            Write-AuditFail "EXO-AUDIT-003" "Administrator Audit Logging" "Disabled (AdminAuditLogEnabled: False)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-003" `
                -Category "Auditing & Governance" `
                -Title "Administrator Audit Logging Posture" `
                -Status "FAIL" `
                -Severity "High" `
                -CurrentValue "Disabled (AdminAuditLogEnabled: False)" `
                -BaselineValue "AdminAuditLogEnabled: True" `
                -RiskRationale "Administrative actions are unmonitored, allowing rogue or compromised admin accounts to modify settings without leaving a trace." `
                -Remediation "Set-AdminAuditLogConfig -AdminAuditLogEnabled `$true" `
                -Details "AdminAuditLogConfig"))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUDIT-002" "Admin Audit Log Config" "Could not query Get-AdminAuditLogConfig: $($_.Exception.Message)"
    }

    # Check 3.4: Mailbox Audit Actions Coverage
    try {
        $defaultAuditSet = @($OrgConfig.DefaultAuditSet)
        if ($defaultAuditSet.Count -gt 0) {
            Write-AuditPass "EXO-AUDIT-004" "Default Mailbox Audit Action Coverage" "Default audit set configured ($($defaultAuditSet.Count) actions audited)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-004" `
                -Category "Auditing & Governance" `
                -Title "Mailbox Audit Action Coverage" `
                -Status "PASS" `
                -Severity "Medium" `
                -CurrentValue "DefaultAuditSet covers $($defaultAuditSet.Count) actions" `
                -BaselineValue "Critical actions (SendAs, SendOnBehalf, Delete, Update) covered" `
                -RiskRationale "Capturing high-risk actions such as SendAs and message deletion ensures unauthorized access and message tampering are auditable." `
                -Remediation "Compliant. Maintain default audit set coverage." `
                -Details "Actions: $($defaultAuditSet -join ', ')"))
        }
        else {
            Write-AuditWarn "EXO-AUDIT-004" "Default Mailbox Audit Action Coverage" "DefaultAuditSet property is empty or custom"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-AUDIT-004" `
                -Category "Auditing & Governance" `
                -Title "Mailbox Audit Action Coverage" `
                -Status "WARNING" `
                -Severity "Medium" `
                -CurrentValue "Default audit set has limited action coverage" `
                -BaselineValue "Comprehensive action coverage on mailboxes" `
                -RiskRationale "Incomplete audit action coverage may fail to log message access (MailItemsAccessed) or deletion actions." `
                -Remediation "Verify mailbox audit actions using Get-OrganizationConfig | Select-Object -ExpandProperty DefaultAuditSet." `
                -Details "Check DefaultAuditSet on OrganizationConfig."))
        }
    }
    catch {
        Write-AuditWarn "EXO-AUDIT-004" "Audit Actions Coverage" "Could not inspect DefaultAuditSet: $($_.Exception.Message)"
    }
}

function Get-AuditDelegation {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$FindingsList,
        [switch]$ShowDetails
    )

    Write-AuditBanner "Auditing Mailbox Forwarding & Privileged Delegation..."

    # Check 4.1: Detection of External Auto-Forwarding Addresses
    try {
        $forwardingMailboxes = @(Get-Mailbox -ResultSize 1000 -Filter "ForwardingAddress -ne `$null -or ForwardingSmtpAddress -ne `$null" -ErrorAction SilentlyContinue)

        if ($forwardingMailboxes.Count -eq 0) {
            Write-AuditPass "EXO-DELEG-001" "Mailbox External Forwarding Rules" "No mailboxes configured with forwarding addresses"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-DELEG-001" `
                -Category "Mailbox Delegation & Suspicious Rules" `
                -Title "Mailbox Auto-Forwarding Addresses" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "0 mailboxes with forwarding addresses configured" `
                -BaselineValue "0 unauthorized external forwarders" `
                -RiskRationale "Mailbox-level forwarding bypasses standard inbox rules and copies incoming mail directly to an external destination, a prime method for data theft." `
                -Remediation "Compliant. Regularly audit Get-Mailbox -Filter 'ForwardingAddress -ne `$null'." `
                -Details "Sample evaluated: up to 1000 mailboxes."))
        }
        else {
            $fwdDetails = [System.Collections.Generic.List[string]]::new()
            foreach ($m in $forwardingMailboxes) {
                $target = if ($m.ForwardingSmtpAddress) { $m.ForwardingSmtpAddress } else { $m.ForwardingAddress }
                $fwdDetails.Add("$($m.UserPrincipalName) -> $target")
            }
            $sampleSummary = ($fwdDetails | Select-Object -First 5) -join '; '
            Write-AuditFail "EXO-DELEG-001" "Mailbox External Forwarding Rules" "Detected $($forwardingMailboxes.Count) mailbox(es) with forwarding: $sampleSummary"
            if ($ShowDetails) {
                Write-AuditDetail "Mailboxes are configured with forwarding addresses." "Set-Mailbox -Identity <UPN> -ForwardingSmtpAddress `$null -ForwardingAddress `$null"
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-DELEG-001" `
                -Category "Mailbox Delegation & Suspicious Rules" `
                -Title "Mailbox Auto-Forwarding Addresses" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "Detected $($forwardingMailboxes.Count) forwarders: $sampleSummary" `
                -BaselineValue "0 unauthorized external forwarders" `
                -RiskRationale "Forwarding addresses can exfiltrate confidential emails, MFA verification messages, and proprietary intellectual property to external adversary-controlled servers." `
                -Remediation "Audit and remove forwarding: Set-Mailbox -Identity <UPN> -ForwardingSmtpAddress `$null -ForwardingAddress `$null" `
                -Details "Forwarding accounts: $($fwdDetails -join ' | ')"))
        }
    }
    catch {
        Write-AuditWarn "EXO-DELEG-001" "Mailbox Forwarding" "Could not inspect mailbox forwarding addresses: $($_.Exception.Message)"
    }

    # Check 4.2: High-Risk Delegation Sprawl on Privileged Mailboxes
    try {
        $orgAdmins = @(Get-RoleGroupMember -Identity "Organization Management" -ErrorAction SilentlyContinue)
        $suspiciousDelegates = [System.Collections.Generic.List[string]]::new()

        foreach ($admin in $orgAdmins) {
            if ($admin.RecipientType -eq 'UserMailbox') {
                $perms = @(Get-MailboxPermission -Identity $admin.Identity -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.User -notmatch 'NT AUTHORITY\\SELF' -and
                        $_.User -notmatch 'Discovery Management' -and
                        $_.User -notmatch 'Exchange Administrative Group' -and
                        $_.User -notmatch 'S-1-5-21' -and
                        $_.AccessRights -contains 'FullAccess'
                    })

                foreach ($p in $perms) {
                    $suspiciousDelegates.Add("$($admin.DisplayName) has FullAccess granted to: $($p.User)")
                }
            }
        }

        if ($suspiciousDelegates.Count -eq 0) {
            Write-AuditPass "EXO-DELEG-002" "Privileged Mailbox Delegation Sprawl" "No unauthorized third-party FullAccess delegates on admin accounts"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-DELEG-002" `
                -Category "Mailbox Delegation & Suspicious Rules" `
                -Title "Privileged Mailbox Delegation Sprawl" `
                -Status "PASS" `
                -Severity "High" `
                -CurrentValue "0 third-party delegates on Organization Management mailboxes" `
                -BaselineValue "0 third-party delegates on admin accounts" `
                -RiskRationale "Delegation on administrative accounts allows delegates to read, delete, or impersonate administrators without possessing their credentials." `
                -Remediation "Compliant. Maintain strict separation of administrative mailboxes." `
                -Details "Admins audited: $($orgAdmins.Count)"))
        }
        else {
            $summary = ($suspiciousDelegates | Select-Object -First 3) -join '; '
            Write-AuditWarn "EXO-DELEG-002" "Privileged Mailbox Delegation Sprawl" "Found delegates: $summary"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-DELEG-002" `
                -Category "Mailbox Delegation & Suspicious Rules" `
                -Title "Privileged Mailbox Delegation Sprawl" `
                -Status "WARNING" `
                -Severity "High" `
                -CurrentValue "Delegation detected: $summary" `
                -BaselineValue "0 third-party delegates on admin accounts" `
                -RiskRationale "Excessive delegation grants lateral movement access into administrative mailboxes, bypassing Conditional Access controls applied to the administrator account." `
                -Remediation "Review and revoke unnecessary delegation: Remove-MailboxPermission -Identity <AdminMailbox> -User <Delegate> -AccessRights FullAccess -Confirm:`$false" `
                -Details "Delegates: $($suspiciousDelegates -join ' | ')"))
        }
    }
    catch {
        Write-AuditWarn "EXO-DELEG-002" "Privileged Delegation" "Could not inspect admin mailbox permissions: $($_.Exception.Message)"
    }
}

function Get-AuditRbac {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[PSCustomObject]]$FindingsList,
        [switch]$ShowDetails
    )

    Write-AuditBanner "Auditing Role-Based Access Control (RBAC)..."

    # Check 5.1: Organization Management Membership
    try {
        $orgAdmins = @(Get-RoleGroupMember -Identity "Organization Management" -ErrorAction Stop)
        $adminCount = $orgAdmins.Count
        $names = ($orgAdmins | ForEach-Object { $_.DisplayName }) -join ', '

        if ($adminCount -le 4) {
            Write-AuditPass "EXO-RBAC-001" "Organization Management Role Group Membership" "Controlled membership ($adminCount accounts): $names"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-001" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Organization Management Role Group Membership" `
                -Status "PASS" `
                -Severity "Critical" `
                -CurrentValue "$adminCount accounts ($names)" `
                -BaselineValue "2-4 dedicated emergency / PIM-managed accounts" `
                -RiskRationale "Organization Management grants complete administrative sovereignty over Exchange Online. Limiting standing membership reduces the attack surface." `
                -Remediation "Compliant. Enforce Entra ID Privileged Identity Management (PIM) for elevation." `
                -Details "Members: $names"))
        }
        elseif ($adminCount -le 8) {
            Write-AuditWarn "EXO-RBAC-001" "Organization Management Role Group Membership" "Elevated standing membership ($adminCount accounts): $names"
            if ($ShowDetails) {
                Write-AuditDetail "Standing Organization Management accounts expand credential compromise risks." "Reduce standing membership to 2-4 break-glass accounts and enforce Entra ID PIM."
            }
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-001" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Organization Management Role Group Membership" `
                -Status "WARNING" `
                -Severity "High" `
                -CurrentValue "$adminCount accounts assigned standing privileges: $names" `
                -BaselineValue "2-4 dedicated accounts" `
                -RiskRationale "Having more than 4 standing Organization Management members exposes the tenant to heightened credential theft and phishing risks." `
                -Remediation "Reduce standing accounts to 2-4 break-glass administrators and use PIM for just-in-time elevation." `
                -Details "Members: $names"))
        }
        else {
            Write-AuditFail "EXO-RBAC-001" "Organization Management Role Group Membership" "Excessive standing membership ($adminCount accounts)!"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-001" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Organization Management Role Group Membership" `
                -Status "FAIL" `
                -Severity "Critical" `
                -CurrentValue "Excessive standing privileges: $adminCount members ($names)" `
                -BaselineValue "2-4 dedicated accounts" `
                -RiskRationale "Excessive permanent Organization Management roles create a critical vulnerability, dramatically increasing blast radius upon credential compromise." `
                -Remediation "Audit and revoke standing Organization Management memberships immediately: Remove-RoleGroupMember -Identity 'Organization Management' -Member <Account>" `
                -Details "Members: $names"))
        }
    }
    catch {
        Write-AuditWarn "EXO-RBAC-001" "Organization Management" "Could not query Organization Management members: $($_.Exception.Message)"
    }

    # Check 5.2: Compliance & Security Administrator Role Groups
    try {
        $compMembers = @(Get-RoleGroupMember -Identity "Compliance Management" -ErrorAction SilentlyContinue)
        $secMembers = @(Get-RoleGroupMember -Identity "Security Administrator" -ErrorAction SilentlyContinue)
        $totalHighRoles = $compMembers.Count + $secMembers.Count

        if ($totalHighRoles -le 8) {
            Write-AuditPass "EXO-RBAC-002" "Compliance & Security Role Group Membership" "Controlled membership ($($compMembers.Count) Compliance, $($secMembers.Count) Security Admins)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-002" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Compliance & Security Role Group Membership" `
                -Status "PASS" `
                -Severity "Medium" `
                -CurrentValue "$($compMembers.Count) Compliance Admins, $($secMembers.Count) Security Admins" `
                -BaselineValue "Dedicated compliance and security staff only" `
                -RiskRationale "Compliance Management provides eDiscovery search and export capabilities over all mailboxes, which must be strictly limited to legal and compliance officers." `
                -Remediation "Compliant. Maintain quarterly access reviews." `
                -Details "Compliance: $($compMembers.DisplayName -join ', ') | Security: $($secMembers.DisplayName -join ', ')"))
        }
        else {
            Write-AuditWarn "EXO-RBAC-002" "Compliance & Security Role Group Membership" "Elevated membership ($($compMembers.Count) Compliance, $($secMembers.Count) Security Admins)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-002" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Compliance & Security Role Group Membership" `
                -Status "WARNING" `
                -Severity "Medium" `
                -CurrentValue "$($compMembers.Count) Compliance Admins, $($secMembers.Count) Security Admins" `
                -BaselineValue "Minimum necessary staff" `
                -RiskRationale "Unmonitored assignment of compliance roles risks unauthorized eDiscovery searches into confidential executive emails." `
                -Remediation "Conduct access recertification for Compliance Management and Security Administrator role groups." `
                -Details "Compliance: $($compMembers.DisplayName -join ', ') | Security: $($secMembers.DisplayName -join ', ')"))
        }
    }
    catch {
        Write-AuditWarn "EXO-RBAC-002" "Compliance & Security Roles" "Could not inspect Compliance/Security roles: $($_.Exception.Message)"
    }

    # Check 5.3: Mail-Enabled Recipients & Help Desk Role Groups
    try {
        $recipMembers = @(Get-RoleGroupMember -Identity "Mail-Enabled Recipients" -ErrorAction SilentlyContinue)
        $helpDeskMembers = @(Get-RoleGroupMember -Identity "Help Desk" -ErrorAction SilentlyContinue)
        $totalSupport = $recipMembers.Count + $helpDeskMembers.Count

        if ($totalSupport -le 15) {
            Write-AuditPass "EXO-RBAC-003" "Help Desk & Recipient Role Group Membership" "Controlled delegation ($($recipMembers.Count) Recipient Admins, $($helpDeskMembers.Count) Help Desk)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-003" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Help Desk & Recipient Role Group Membership" `
                -Status "PASS" `
                -Severity "Low" `
                -CurrentValue "$($recipMembers.Count) Recipient Admins, $($helpDeskMembers.Count) Help Desk" `
                -BaselineValue "Controlled support personnel" `
                -RiskRationale "Recipient administrators can modify proxy addresses and forwarding rules, potentially enabling password reset interception." `
                -Remediation "Compliant. Regularly recertify Help Desk access." `
                -Details "Recipient Admins: $($recipMembers.DisplayName -join ', ')"))
        }
        else {
            Write-AuditWarn "EXO-RBAC-003" "Help Desk & Recipient Role Group Membership" "Broad delegation ($totalSupport total accounts across Recipient & Help Desk)"
            $FindingsList.Add((New-AuditFindingObject `
                -Id "EXO-RBAC-003" `
                -Category "Role-Based Access Control (RBAC)" `
                -Title "Help Desk & Recipient Role Group Membership" `
                -Status "WARNING" `
                -Severity "Low" `
                -CurrentValue "Broad delegation: $totalSupport accounts" `
                -BaselineValue "Controlled support personnel" `
                -RiskRationale "Sprawling recipient administration privileges increases the likelihood of insider mistakes or account compromise." `
                -Remediation "Recertify and prune obsolete Help Desk and Recipient Management assignments." `
                -Details "Total accounts: $totalSupport"))
        }
    }
    catch {
        Write-AuditWarn "EXO-RBAC-003" "Help Desk & Recipient Roles" "Could not inspect Help Desk/Recipient roles: $($_.Exception.Message)"
    }
}

# -----------------------------------------------------------------------------
# MAIN SCRIPT EXECUTION LIFECYCLE (begin, process, end)
# -----------------------------------------------------------------------------
$findings = [System.Collections.Generic.List[PSCustomObject]]::new()
$orgConfig = $null
$tenantDomain = "ExchangeOnline"

try {
    # -------------------------------------------------------------------------
    # BEGIN BLOCK
    # -------------------------------------------------------------------------
    Write-Host @"
================================================================================
   ___             _ _ _      ______          _
  / _ \           | (_) |     | ___ \        | |
 / /_\ \_   _  __| |_| |_ ___ | |_/ /___  ___| |__   __ _ _ __   __ _  ___
 |  _  | | | |/ _` | | __/ __||    // _ \/ __| '_ \ / _` | '_ \ / _` |/ _ \
 | | | | |_| | (_| | | |_\__ \| |\ \  __/ (__| | | | (_| | | | | (_| |  __/
 \_| |_/\__,_|\__,_|_|\__|___/\_| \_\___|\___|_| |_|\__,_|_| |_|\__, |\___|
                                                                  __/ |
 Microsoft 365 Exchange Online Security Posture Auditor          |___/
 Version: 1.0.0 | Author: Roman Pindela | License: MIT
================================================================================
"@ -ForegroundColor Cyan

    # 1. Validate & create output directory
    if (-not (Test-Path -Path $ReportPath)) {
        Write-Host "[*] Creating report directory: '$ReportPath'..." -ForegroundColor Cyan
        New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null
    }

    # 2. Check and connect to Exchange Online
    Test-ExchangePrerequisite
    Connect-ExchangeSession -AdminUPN $AdminUserPrincipalName

    # 3. Retrieve Organization Config
    Write-Host "[*] Gathering tenant organization metadata..." -ForegroundColor Cyan
    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
    if ($null -ne $orgConfig -and $orgConfig.Identity) {
        $tenantDomain = $orgConfig.Identity.ToString().Split('\')[0]
        if ([string]::IsNullOrWhiteSpace($tenantDomain)) {
            $tenantDomain = "ExchangeOnline"
        }
    }

    # -------------------------------------------------------------------------
    # PROCESS BLOCK (Execute Audit Suites)
    # -------------------------------------------------------------------------
    Get-AuditModernAuth -FindingsList $findings -OrgConfig $orgConfig -ShowDetails:$ShowDetailedOutput
    Get-AuditMailRouting -FindingsList $findings -ShowDetails:$ShowDetailedOutput
    Get-AuditGovernance -FindingsList $findings -OrgConfig $orgConfig -ShowDetails:$ShowDetailedOutput
    Get-AuditDelegation -FindingsList $findings -ShowDetails:$ShowDetailedOutput
    Get-AuditRbac -FindingsList $findings -ShowDetails:$ShowDetailedOutput

    # -------------------------------------------------------------------------
    # END BLOCK (Metrics, Report Generation, Pipeline Output)
    # -------------------------------------------------------------------------
    $totalChecks = $findings.Count
    $passChecks = @($findings | Where-Object { $_.Status -eq 'PASS' }).Count
    $warnChecks = @($findings | Where-Object { $_.Status -eq 'WARNING' }).Count
    $failChecks = @($findings | Where-Object { $_.Status -eq 'FAIL' }).Count

    $critRisks = @($findings | Where-Object { $_.Status -eq 'FAIL' -and $_.Severity -eq 'Critical' }).Count
    $highRisks = @($findings | Where-Object { $_.Status -eq 'FAIL' -and $_.Severity -eq 'High' }).Count

    # Calculate Weighted Security Posture Score
    $weights = @{
        'Critical'      = 10.0
        'High'          = 7.0
        'Medium'        = 4.0
        'Low'           = 2.0
        'Informational' = 1.0
    }

    $maxWeight = 0.0
    $earnedWeight = 0.0

    foreach ($f in $findings) {
        $w = $weights[$f.Severity]
        $maxWeight += $w
        if ($f.Status -eq 'PASS') {
            $earnedWeight += $w
        }
        elseif ($f.Status -eq 'WARNING') {
            $earnedWeight += ($w * 0.5)
        }
    }

    $postureScore = if ($maxWeight -gt 0) {
        [Math]::Round(($earnedWeight / $maxWeight) * 100.0, 1)
    }
    else {
        100.0
    }

    $statsTable = @{
        Total    = $totalChecks
        Pass     = $passChecks
        Warn     = $warnChecks
        Fail     = $failChecks
        Critical = $critRisks
        High     = $highRisks
    }

    # Generate HTML Dashboard
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $safeTenant = ($tenantDomain -replace '[^a-zA-Z0-9_\-]', '_')
    $reportFileName = "Audit-ExchangeSecuritySummary_${safeTenant}_${timestamp}.html"
    $reportFullPath = Join-Path -Path $ReportPath -ChildPath $reportFileName

    Write-Host "`n[*] Generating standalone executive HTML report..." -ForegroundColor Cyan
    Export-ExchangeSecurityHtmlReport `
        -Findings $findings `
        -TenantName $tenantDomain `
        -FilePath $reportFullPath `
        -PostureScore $postureScore `
        -Stats $statsTable

    Write-Host "[+] Report generated successfully:" -ForegroundColor Green
    Write-Host "    $reportFullPath" -ForegroundColor White

    # Console Executive Summary Card
    $scoreColor = if ($postureScore -ge 85.0) { 'Green' } elseif ($postureScore -ge 70.0) { 'Yellow' } else { 'Red' }
    Write-Host @"

================================================================================
                           EXECUTIVE AUDIT SUMMARY
================================================================================
 Tenant Identity:         $tenantDomain
 Overall Posture Score:   $postureScore%
 Total Checks Evaluated:  $totalChecks
 Passed (Compliant):      $passChecks [OK]
 Warnings (Review):       $warnChecks [WARN]
 Failures (Risks):        $failChecks [FAIL]

 Critical Risks:          $critRisks
 High Risks:              $highRisks

 Report Path:             $reportFullPath
================================================================================
"@ -ForegroundColor $scoreColor

    if ($OpenReport) {
        Write-Host "[*] Launching report in default browser..." -ForegroundColor Cyan
        Start-Process $reportFullPath
    }

    # Emit PSCustomObject stream to pipeline
    $findings | Write-Output
}
catch {
    Write-Host "`n[!] Critical error during audit execution: $($_.Exception.Message)" -ForegroundColor Red
    throw $_
}
finally {
    Disconnect-ExchangeSession
}
