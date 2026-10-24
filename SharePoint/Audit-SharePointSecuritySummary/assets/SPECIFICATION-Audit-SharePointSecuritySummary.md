# TASK SPECIFICATION: Audit-SharePointSecuritySummary

> **Target Agent**: Antigravity / Senior PowerShell Systems Architect  
> **Target Date**: 2026-10-24  
> **Author**: Roman Pindela ([roman.pindela@gmail.com](mailto:roman.pindela@gmail.com))  
> **Status**: Ready for Implementation

---

## 1\. PROJECT CONTEXT & REPOSITORY SCAFFOLDING

### 1.1 Directory & File Layout

The script and its associated assets must reside strictly within the following directory tree:

c:\\Dane\\rpindela\\GitHub\\Microsoft-365\\SharePoint\\Audit-SharePointSecuritySummary\\

├── assets\\

│   └── architecture-flow.png

├── Audit-SharePointSecuritySummary.ps1

└── README.md

### 1.2 Scaffold Initialization Snippet (PowerShell)

Execute the following snippet to ensure the target directory structure exists prior to file generation:

\$repoPath \= "c:\\Dane\\rpindela\\GitHub\\Microsoft-365\\SharePoint\\Audit-SharePointSecuritySummary"

\$assetsPath \= Join-Path \-Path \$repoPath \-ChildPath "assets"

if (-not (Test-Path \-Path \$assetsPath)) {

    New-Item \-Path \$assetsPath \-ItemType Directory \-Force \| Out-Null

    Write-Host "\[+\] Directory structure initialized at: \$repoPath" \-ForegroundColor Green

} else {

    Write-Host "\[i\] Target directory already exists: \$repoPath" \-ForegroundColor Yellow

}

---

## 2\. TASK REQUIREMENTS & BUSINESS LOGIC

### 2.1 Core Objective

Develop a production-grade PowerShell automation tool (`Audit-SharePointSecuritySummary.ps1`) to audit, evaluate, and report the security posture across SharePoint Online (SPO) site collections within a Microsoft 365 tenant. The output must provide an executive summary as well as deep-dive operational risk metrics.

### 2.2 Functional Requirements & Audit Scope

The script must connect to SharePoint Online (via `PnP.PowerShell` or `Microsoft.Online.SharePoint.PowerShell`) and analyze the following risk vectors:

1. **External Sharing & Guest Governance:**  
   - Audit site collection sharing capabilities (`ExternalUserSharingDisabled`, `ExistingExternalUserSharingOnly`, `ExternalUserAndGuestSharing`).  
   - Highlight sites allowing anonymous/anyone link access.  
2. **Permission Inheritance & Unique Access:**  
   - Detect site collections, document libraries, and high-risk lists with broken inheritance / unique permissions.  
3. **Site Ownership & Administrative Footprint:**  
   - Flag orphaned site collections (0 active Site Collection Administrators/Owners).  
   - Flag sites with excessive administrative accounts (more than 5 Site Collection Administrators).  
4. **Anonymous & Anyone Links:**  
   - Audit presence of active anonymous sharing links and check if tenant default link expiration policies are enforced.  
5. **Sensitivity Labels & Privacy Settings:**  
   - Capture M365 Group privacy status (`Public` vs `Private`) and applied Sensitivity Labels.  
6. **Reporting & Multi-Format Export:**  
   - **Console Dashboard:** Color-coded status table (Green \= Secure, Yellow \= Warning, Red \= High Risk / Non-Compliant).  
   - **CSV Export:** Comprehensive raw dataset for operational remediation.  
   - **HTML Executive Summary:** Standalone responsive HTML report with summary cards (Total Sites, External Sharing Enabled, Orphaned Sites, High Risk Sites).  
   - **Structured Pipeline:** Return `[PSCustomObject]` collection to the pipeline.

### 2.3 Authentication Models

The script must support both interactive and automated enterprise workflows:

- **Interactive Mode:** Modern browser OAuth / Device Code (`-Interactive`).  
- **App-Only / Unattended Mode:** App registration authentication using Microsoft Entra ID (`-TenantId`, `-ClientId`, `-CertificateThumbprint` or `-CertificatePath`).  
- **Graceful Disconnection:** Always ensure sessions are disconnected in the `finally {}` block (`Disconnect-PnPOnline` / `Disconnect-SPOService`).

---

## 3\. MANDATORY CODING STANDARDS & ARCHITECTURE RULES

### 3.1 Language & Style

- **100% English:** All code, comments, documentation, console output, and help screens must be in English.  
- **PSScriptAnalyzer:** Zero warnings or errors against default and strict rule sets.  
- **Strict Mode:**  
    
  Set-StrictMode \-Version Latest  
    
  \$ErrorActionPreference \= 'Stop'  
    
- **Cmdlet Binding:**  
    
  \[CmdletBinding(SupportsShouldProcess \= \$true, ConfirmImpact \= 'Medium')\]

### 3.2 Script Metadata Header

Every script file must begin with standard author metadata:

\<\#

.SYNOPSIS

    Audit-SharePointSecuritySummary \- Audits SharePoint Online security posture and generates executive reports.

.DESCRIPTION

    Comprehensive security audit tool for Microsoft 365 SharePoint Online. Evaluates external sharing,

    orphaned site collections, excessive admin assignments, anonymous links, and broken permission inheritance.

    Exports findings to CSV, standalone HTML executive report, and structured PowerShell pipeline objects.

.PARAMETER AdminUrl

    The SharePoint Online Admin URL.

.PARAMETER ClientId

    Microsoft Entra ID Application (Client) ID for unattended App-Only authentication.

.PARAMETER TenantId

    Microsoft Entra ID Directory (Tenant) ID.

.PARAMETER CertificateThumbprint

    Certificate thumbprint in Cert:\\CurrentUser\\My or Cert:\\LocalMachine\\My for App-Only authentication.

.PARAMETER Interactive

    Switch to initiate interactive modern authentication.

.PARAMETER OutputPath

    Directory where CSV, HTML, and log reports will be saved. Defaults to 'c:\\Logs'.

.PARAMETER GenerateHtmlReport

    Switch to generate a standalone HTML executive summary report.

.PARAMETER LogPath

    Path for the transcript/execution log. Defaults to 'C:\\Logs'.

.PARAMETER Help

    Displays the script help screen and exits.

.EXAMPLE

    .\\Audit-SharePointSecuritySummary.ps1 \-AdminUrl \$TenantAdminUrl \-Interactive \-GenerateHtmlReport

.EXAMPLE

    .\\Audit-SharePointSecuritySummary.ps1 \-AdminUrl \$TenantAdminUrl \-ClientId "00000000-0000-0000-0000-000000000000" \-TenantId "11111111-1111-1111-1111-111111111111" \-CertificateThumbprint "A1B2C3D4E5..." \-OutputPath "C:\\Reports"

.NOTES

    Author: Roman Pindela

    Email: roman.pindela@gmail.com

    GitHub: romanpindela

    Version: 1.0.0

    Last Modified: 2026-10-24

\#\>

### 3.3 Zero-Parameter & `-Help` Safety Guard

The script must implement an immediate help screen trigger if run without arguments or when explicitly passed `-Help` / `-h`:

\[Alias('h')\]

\[switch\]\$Help

\# In begin {} block:

if (\$Help \-or (\$PSBoundParameters.Count \-eq 0)) {

    Show-ScriptHelp

    exit 0

}

### 3.4 Logging Infrastructure (`C:\Logs\`)

The script must log all actions to `C:\Logs\`:

- Format: `C:\Logs\<yyyyMMdd-HHmmss>-Audit-SharePointSecuritySummary-<COMPUTERNAME>-<USERNAME>.log`  
- Implement a thread-safe `Write-Log` internal function:

function Write-Log {

    \[CmdletBinding()\]

    param(

        \[Parameter(Mandatory \= \$true)\]

        \[string\]\$Message,

        \[ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')\]

        \[string\]\$Level \= 'INFO'

    )

    \$timestamp \= Get-Date \-Format "yyyy-MM-dd HH:mm:ss"

    \$logEntry \= "\[\$timestamp\] \[\$Level\] \$Message"

    

    \# Write to log file

    if (\$script:LogFilePath) {

        Add-Content \-Path \$script:LogFilePath \-Value \$logEntry \-Encoding UTF8

    }

    

    \# Console colorized output

    switch (\$Level) {

        'INFO'    { Write-Host \$logEntry \-ForegroundColor Cyan }

        'WARN'    { Write-Host \$logEntry \-ForegroundColor Yellow }

        'ERROR'   { Write-Host \$logEntry \-ForegroundColor Red }

        'SUCCESS' { Write-Host \$logEntry \-ForegroundColor Green }

    }

}

### 3.5 Clean Code & Error Handling

- Encapsulate all logic within `begin {}`, `process {}`, and `end {}` blocks.  
- Wrap external module invocations and API connections inside `try { ... } catch { ... } finally { ... }`.  
- Ensure all discovered issues are classified cleanly with risk levels:  
  - `Critical`: Anonymous access enabled on confidential sites or 0 active owners.  
  - `High`: External sharing enabled without expiration; more than 5 site collection administrators.  
  - `Medium`: Broken permission inheritance detected.  
  - `Low`: Compliant / Standard governance.

---

## 4\. DOCUMENTATION SPECIFICATION (`README.md`)

The generated `README.md` must adhere to professional GitHub repository standards and include:

1. **Header & Badges:**  
   - PowerShell 7+ / 5.1 compatibility badges.  
   - PSScriptAnalyzer Clean badge.  
   - Version 1.0.0 badge.  
2. **Overview:** High-level executive description of why SharePoint security auditing is critical (compliance, data exfiltration prevention, M365 posture management).  
3. **Architecture Diagram:**  
   - Mermaid sequence/flow diagram showing the audit lifecycle (Authentication \-\> Inventory \-\> Deep Inspection \-\> Risk Assessment \-\> Dashboard & Report Generation \-\> Session Teardown).  
   - ASCII schematic representation.  
4. **Prerequisites & Module Installation:**  
   - Required PowerShell modules (`PnP.PowerShell` or `Microsoft.Online.SharePoint.PowerShell`).  
   - Required RBAC / Entra roles (SharePoint Administrator, Global Reader, or Sites.FullControl.All app role).  
   - Execution Policy & Unblocking command:  
       
     Unblock-File \-Path ".\\Audit-SharePointSecuritySummary.ps1"

     
5. **Parameter Reference Table:** Markdown table defining each parameter, type, default value, and requirement status.  
6. **Execution Examples:** Minimum of 4 realistic scenarios (Interactive audit, App-Only scheduled audit, HTML report generation, Pipeline filtering for high-risk sites).  
7. **Logging & Remediation Guide:** Explanation of log format, log rotation, and remediation tips for flagged items.  
8. **Changelog (SemVer):** Starting at `v1.0.0 (2026-10-24)`.  
9. **Author Info Card:** Roman Pindela.

---

## 5\. ROOT REPOSITORY INDEXING

The script must be indexed into the parent category documentation:

- File: `c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\README.md`  
- Table entry format:  
    
  \| \[Audit-SharePointSecuritySummary\](./Audit-SharePointSecuritySummary/) \| Audits SharePoint Online external sharing, permissions, orphaned sites, and exports security dashboards (HTML/CSV). \| \`v1.0.0\` \| Active \|  
    
- **PowerShell Idempotent Registration Snippet:**  
    
  \$parentReadme \= "c:\\Dane\\rpindela\\GitHub\\Microsoft-365\\SharePoint\\README.md"  
    
  \$entry \= "\| \[Audit-SharePointSecuritySummary\](./Audit-SharePointSecuritySummary/) \| Audits SharePoint Online external sharing, permissions, orphaned sites, and exports security dashboards (HTML/CSV). \| \`v1.0.0\` \| Active \|"  
    
  if (Test-Path \-Path \$parentReadme) {  
    
      \$content \= Get-Content \-Path \$parentReadme \-Raw  
    
      if (\$content \-notmatch "Audit-SharePointSecuritySummary") {  
    
          Add-Content \-Path \$parentReadme \-Value "\`n\$entry" \-Encoding UTF8  
    
          Write-Host "\[+\] Indexed into parent README: \$parentReadme" \-ForegroundColor Green  
    
      } else {  
    
          Write-Host "\[i\] Entry already present in parent README." \-ForegroundColor Yellow  
    
      }  
    
  }

---

## 6\. GIT CONVENTIONAL COMMITS

When committing this work to Git, use the following conventional commit message:

feat(sharepoint): add Audit-SharePointSecuritySummary script and documentation

\- Implement SharePoint Online security audit automation (Audit-SharePointSecuritySummary.ps1)

\- Add external sharing, orphaned sites, and admin footprint inspection

\- Implement dual authentication (Interactive & Certificate App-Only)

\- Implement multi-format reporting (Console ANSI, CSV, and standalone HTML Dashboard)

\- Add centralized C:\\Logs\\ transcript logging

\- Add comprehensive repository documentation and architecture diagrams (README.md)

---

## 7\. EXECUTION DIRECTIVE FOR ANTIGRAVITY

DIRECTIVE TO AGENT:

Generate the complete, production-ready code files for this specification without shortcuts:

1\. "Audit-SharePointSecuritySummary.ps1" in UTF-8 with BOM:

   \- Complete implementation with all parameters, validation attributes, Show-ScriptHelp, Write-Log, PnP connection handling, audit algorithms, ANSI console rendering, CSV export, and standalone HTML dashboard generator.

   \- Do NOT omit any function bodies or use placeholder comments.

2\. "README.md" in UTF-8:

   \- Full documentation including Mermaid flowchart, ASCII architecture, installation steps, Unblock-File instructions, usage examples, and author card.