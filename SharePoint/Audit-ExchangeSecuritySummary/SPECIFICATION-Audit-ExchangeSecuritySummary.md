# TASK SPECIFICATION: Audit-ExchangeSecuritySummary (SharePoint Security Assessment)

## 1. PROJECT CONTEXT & REPOSITORY SCAFFOLDING

### Target Directory Structure
The script, documentation, and asset artifacts must reside within the following repository layout:
- Root repository path: `c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\Audit-ExchangeSecuritySummary`
- Assets subfolder: `c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\Audit-ExchangeSecuritySummary\assets`
- Main script file: `Audit-ExchangeSecuritySummary.ps1` (UTF-8 with BOM)
- Documentation: `README.md` (UTF-8)
- Specification file: `SPECIFICATION-Audit-ExchangeSecuritySummary.md`

### Directory Provisioning Snippet
```powershell
$targetPath = "c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\Audit-ExchangeSecuritySummary"
$assetsPath = Join-Path -Path $targetPath -ChildPath "assets"

if (-not (Test-Path -Path $assetsPath)) {
    New-Item -ItemType Directory -Path $assetsPath -Force | Out-Null
    Write-Host "[OK] Created directory structure at: $targetPath" -ForegroundColor Green
}
```

---

## 2. TASK REQUIREMENTS & BUSINESS LOGIC

### Primary Objective
Develop an enterprise-grade PowerShell audit script to extract and summarize key SharePoint Online security configurations, site permissions, and external sharing policies across tenant site collections. The script generates both structured console metrics and exportable artifacts (CSV / JSON) summarizing the security posture.

### Scope & Core Operations
1. **Tenant & Site Discovery:**
   - Connect to SharePoint Online using either interactive authentication or App-Only credentials (Certificate Thumbprint / Client Secret with Application Permissions).
   - Enumerate SharePoint Online Site Collections (excluding system / app catalog personal sites unless explicitly specified).
2. **Security & Governance Assessment:**
   - **Sharing Capabilities:** Audit `SharingCapability` (e.g., `Disabled`, `ExistingGuests`, `ExternalUserSharingOnly`, `ExternalUserAndGuestSharing`).
   - **Access Governance:** Check for anonymous access links, guest permissions, and whether non-owners are permitted to invite external users (`MembersCanShare`).
   - **Site Administrators & Sensitive Permissions:** Inventory Site Collection Administrators (Primary, Secondary, and elevated admins).
   - **Conditional Access & Device Access Policies:** Report if unmanaged devices or conditional restrictions apply to the site.
3. **Output & Artifact Generation:**
   - Display a color-coded security score / summary matrix directly in the console.
   - Export comprehensive results to CSV/JSON in the execution folder or specified path.
4. **Resilience & Safe Termination:**
   - Support throttling handling (HTTP 429 / 503 backoff logic).
   - Ensure explicit session disconnect and resource cleanup in `finally {}` blocks.

---

## 3. MANDATORY CODING STANDARDS & ARCHITECTURE RULES

### Language & Output
- **100% English:** All script logic, function names, parameter documentation, console output, comments, log messages, and README must be exclusively in English.
- **Rich Console Formatting:** Use clean, structured, and color-coded console outputs with clear separators and status tags (`[INFO]`, `[WARN]`, `[ERROR]`, `[SUCCESS]`).

### Clean Code & PSScriptAnalyzer Compliance
- **Cmdlet Binding:**
  ```powershell
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param( ... )
  ```
- **Strict Execution & Error Handling:**
  ```powershell
  Set-StrictMode -Version Latest
  $ErrorActionPreference = 'Stop'
  ```
- **Pipeline Architecture:** Proper modular division into `begin {}`, `process {}`, and `end {}` blocks.
- **Strict Verb Usage:** Use only approved PowerShell verbs (`Get-Verb`), e.g., `Invoke-`, `Get-`, `Test-`, `Export-`.
- **Pipeline Output:** Return structured `[PSCustomObject]` instances; avoid returning raw unstructured strings.

### Author & Version Metadata (SemVer)
- **Author:** Roman Pindela
- **Email:** roman.pindela@gmail.com
- **GitHub:** https://github.com/romanpindela
- **Version:** 1.0.0
- **Last Modified:** 2026-10-24

### Help & CLI UX
- Implement parameter `[Alias('h')][switch]$Help`.
- Running with `-Help` OR executing without bound parameters (`if ($PSBoundParameters.Count -eq 0)`) must invoke `Show-ScriptHelp` and exit cleanly (`exit 0`).

### Logging Framework (`C:\Logs\`)
- Log directory: `C:\Logs`
- Format: `C:\Logs\<yyyyMMdd-HHmmss>-Audit-ExchangeSecuritySummary-<COMPUTERNAME>-<USERNAME>.logs`
- Function: Implement `Write-Log -Message <string> -Level <Info|Warn|Error|Debug>`. Ensure the directory is created if missing.

### Microsoft 365 / SharePoint Authentication
- Support interactive browser logon as well as certificate-based Service Principal logon (`-TenantId`, `-ClientId`, `-CertificateThumbprint`).
- Ensure disconnect cmdlets (`Disconnect-PnPOnline` / `Disconnect-SPOService`) are invoked in `finally {}`.

---

## 4. DOCUMENTATION SPECIFICATION (`README.md`)

The generated `README.md` must be placed in `c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\Audit-ExchangeSecuritySummary\README.md` and contain:
1. **Header & Badges:** PowerShell 5.1/7+, License, Version, PSScriptAnalyzer.
2. **Overview:** Clear executive and operational summary.
3. **Architecture:** Both an ASCII flowchart and a Mermaid.js diagram illustrating authentication, enumeration, evaluation, and reporting.
4. **Prerequisites & Unblock Directive:**
   ```powershell
   Unblock-File -Path ".\Audit-ExchangeSecuritySummary.ps1"
   ```
5. **Parameters Reference Table:** Listing name, type, required status, and descriptions.
6. **Usage Examples:** Interactive run, automated run via app registration, filtering specific site templates.
7. **Troubleshooting & Logging:** Paths, common error codes, and permission troubleshooting.
8. **Changelog & Author Card.**

---

## 5. ROOT REPOSITORY INDEXING

Ensure the root SharePoint README file (`c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\README.md`) registers this utility:

### Markdown Entry
```markdown
| [Audit-ExchangeSecuritySummary](./Audit-ExchangeSecuritySummary) | Audits SharePoint Online site collections for external sharing, site admins, and security postures. | v1.0.0 |
```

### Idempotent PowerShell Indexing Snippet
```powershell
$rootReadme = "c:\Dane\rpindela\GitHub\Microsoft-365\SharePoint\README.md"
$entry = "| [Audit-ExchangeSecuritySummary](./Audit-ExchangeSecuritySummary) | Audits SharePoint Online site collections for external sharing, site admins, and security postures. | v1.0.0 |"

if (Test-Path -Path $rootReadme) {
    $content = Get-Content -Path $rootReadme -Raw
    if ($content -notmatch "Audit-ExchangeSecuritySummary") {
        Add-Content -Path $rootReadme -Value "`n$entry"
        Write-Host "[OK] Added script reference to root README." -ForegroundColor Green
    }
}
```

---

## 6. CONVENTIONAL COMMITS SPECIFICATION

Use the following Git commit template upon completion:
```text
feat(sharepoint): add Audit-ExchangeSecuritySummary script and documentation

- Implement SharePoint Online security posture assessment script
- Add external sharing capabilities and site admin auditing logic
- Implement structured console reporting and C:\Logs logging
- Add comprehensive README with Mermaid architecture diagram
```

---

## 7. EXECUTION DIRECTIVE FOR ANTIGRAVITY

> **INSTRUCTION FOR AGENT / ANTIGRAVITY:**  
> Generate the full, production-ready script `Audit-ExchangeSecuritySummary.ps1` and documentation `README.md` strictly adhering to the specifications above. Do NOT use placeholder comments, incomplete logic, or `TODO` annotations. All code must be completely written out, strictly validated for PSScriptAnalyzer, fully commented in English, and ready for immediate deployment.