# TASK SPECIFICATION: Audit-SharePoint-Site

## 1. PROJECT CONTEXT & REPOSITORY SCAFFOLDING
- **Root Repository Path:** `c:\GitHub\Microsoft 365\SharePoint\`
- **Target Folder:** `c:\GitHub\Microsoft 365\SharePoint\Audit-SharePoint-Site`
- **Assets Folder:** `c:\GitHub\Microsoft 365\SharePoint\Audit-SharePoint-Site\assets`
- **Script File:** `Audit-SharePoint-Site.ps1` (Encoding: UTF-8 with BOM)
- **Documentation:** `README.md` (Encoding: UTF-8)
- **Log Output Directory:** `C:\Logs\`

### Scaffold Initialization Snippet:
```powershell
$baseDir = "c:\GitHub\Microsoft 365\SharePoint\Audit-SharePoint-Site"
$assetsDir = Join-Path -Path $baseDir -ChildPath "assets"
$null = New-Item -ItemType Directory -Path $assetsDir -Force
$null = New-Item -ItemType File -Path (Join-Path -Path $baseDir -ChildPath "Audit-SharePoint-Site.ps1") -Force
$null = New-Item -ItemType File -Path (Join-Path -Path $baseDir -ChildPath "README.md") -Force
```

---

## 2. BUSINESS LOGIC & TECHNICAL REQUIREMENTS
- **Target Environment:** SharePoint Online / Microsoft 365.
- **Module Requirements:** `PnP.PowerShell` (recommended for deep permission inspection) with fallback / dual-support checks, or Microsoft Graph SDK.
- **Core Functionality:**
  - Audit all levels of permissions within a specified SharePoint site (`-SiteUrl`).
  - Extract site-level role assignments (Owners, Members, Visitors, custom SharePoint groups, Azure AD / Entra ID security groups, M365 groups).
  - Recursively inspect document libraries, lists, folders, and items with broken permission inheritance (`HasUniqueRoleAssignments -eq $true`).
  - Enumerate users, assigned permission levels (Full Control, Edit, Read, Custom levels), and distinguish direct user assignments vs. group membership.
  - Export audit results to an interactive formatted console view (`[PSCustomObject]`) and support automated CSV/JSON export via dedicated parameters (`-ExportPath`, `-ExportFormat`).
- **Input Parameters:**
  - `[Parameter(Mandatory = $true, ValueFromPipeline = $true)][ValidateNotNullOrEmpty()][string]$SiteUrl`
  - `[Parameter(Mandatory = $false)][string]$ClientId` (for app-only authentication)
  - `[Parameter(Mandatory = $false)][string]$Tenant` (for app-only authentication)
  - `[Parameter(Mandatory = $false)][string]$CertificateThumbprint` (for certificate auth)
  - `[Parameter(Mandatory = $false)][switch]$Interactive` (interactive web login)
  - `[Parameter(Mandatory = $false)][switch]$IncludeInherited` (switch to audit inherited vs only unique permissions)
  - `[Parameter(Mandatory = $false)][string]$ExportPath`
  - `[Parameter(Mandatory = $false)][ValidateSet('CSV', 'JSON')][string]$ExportFormat = 'CSV'`
  - `[Parameter(Mandatory = $false)][Alias('h')][switch]$Help`
- **Output:** Structured PowerShell objects + detailed log entries in `C:\Logs\`.

---

## 3. MANDATORY ARCHITECTURAL RULES & STANDARDS
- **Language:** 100% English for code, variables, function names, comments, console messages, and documentation.
- **Script Execution & Help Handling:**
  - Include parameter `[Alias('h')][switch]$Help`.
  - When invoked with `-Help` OR when `$PSBoundParameters.Count -eq 0`, call `Show-ScriptHelp` and safely terminate with `exit 0`.
- **Clean Code & PSScriptAnalyzer:**
  - Zero PSScriptAnalyzer warnings.
  - Strict mode enabled: `Set-StrictMode -Version Latest`.
  - Error preference: `$ErrorActionPreference = 'Stop'`.
  - Advanced cmdlet declaration: `[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]`.
  - Structured pipeline: `begin {}`, `process {}`, `end {}`.
  - Strong typing on all parameters and function definitions.
- **Author Metadata Header:**
  - Author: Roman Pindela
  - Email: roman.pindela@gmail.com
  - GitHub: https://github.com/romanpindela
  - Version: 1.0.0
  - Last Modified: 2026-10-23
- **Progress Reporting:**
  - Implement dynamic `Write-Progress` tracking site scanning, library parsing, and item-level permission enumeration.
- **Logging Architecture:**
  - Directory: `C:\Logs\` (auto-created if missing).
  - Format: `<yyyyMMdd-HHmmss>-Audit-SharePoint-Site-<COMPUTERNAME>-<USERNAME>.logs`.
  - Function: `Write-Log -Message <string> -Level <INFO|WARN|ERROR|SUCCESS>`.
  - Dual stream: Output to log file and colored console stream (`Write-Host`).
- **Session Cleanup:**
  - Ensure all active connections (`Disconnect-PnPOnline` / Graph sessions) are cleanly terminated inside a `finally {}` block.

---

## 4. DOCUMENTATION SPECIFICATION (README.md)
The `README.md` must be created in `c:\GitHub\Microsoft 365\SharePoint\Audit-SharePoint-Site\README.md` containing:
1. **Badges:** PowerShell 7+ / 5.1, Platform: SharePoint Online, License: MIT.
2. **Overview:** High-level summary of permission auditing capabilities.
3. **Architecture:** ASCII and Mermaid diagram visualizing the scan workflow (Site -> Lists -> Folders/Items -> Unique Perms -> Export).
4. **Prerequisites & Installation:**
   - Required PowerShell modules (`PnP.PowerShell`).
   - PowerShell execution policy note: `Unblock-File -Path ".\Audit-SharePoint-Site.ps1"`.
5. **Parameters Table:** Detailed table with Name, Type, Required, and Description.
6. **Usage Examples:**
   - Example 1: Interactive authentication and console output.
   - Example 2: Certificate/App-Only connection auditing broken permissions and exporting to CSV.
7. **Logging & Diagnostics:** Location and format of logs.
8. **Changelog:** Initial SemVer release 1.0.0.
9. **Author:** Contact and GitHub details.

---

## 5. ROOT REPOSITORY INDEXING
Ensure the script is registered in the parent repository README (`c:\GitHub\Microsoft 365\SharePoint\README.md`).

### Idempotent PowerShell Indexing Snippet:
```powershell
$rootReadme = "c:\GitHub\Microsoft 365\SharePoint\README.md"
$entry = "| [Audit-SharePoint-Site](./Audit-SharePoint-Site) | Audits site, library, folder, and item-level permissions with broken inheritance detection. | 1.0.0 |"

if (Test-Path -Path $rootReadme) {
    $content = Get-Content -Path $rootReadme -Raw
    if ($content -notmatch "Audit-SharePoint-Site") {
        Add-Content -Path $rootReadme -Value "`n$entry"
    }
}
```

---

## 6. CONVENTIONAL COMMIT TEMPLATE
```
feat(sharepoint): introduce Audit-SharePoint-Site script and spec

- Add permission auditing for sites, libraries, and items with broken inheritance
- Implement PnP.PowerShell connection handling (interactive and app-only)
- Add C:\Logs logging module and progress reporting
- Add documentation and repository indexing
```

---

## 7. EXECUTION DIRECTIVE FOR ANTIGRAVITY
You are instructed to generate the full, production-ready implementation of `Audit-SharePoint-Site.ps1` and `README.md` without any placeholders, truncated loops, or `TODO` markers. Every function (`Write-Log`, `Show-ScriptHelp`, `Connect-TargetSite`, `Get-SitePermissionsAudit`, `Export-AuditData`) must be fully written, strongly typed, and pass PSScriptAnalyzer with zero warnings.