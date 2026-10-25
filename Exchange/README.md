# Microsoft 365 Exchange Online Automation Toolkit

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-v3.0%2B-0078D4.svg)](https://learn.microsoft.com/powershell/exchange/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v2.0%2B-008272.svg)](https://learn.microsoft.com/graph/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A centralized, enterprise-grade PowerShell automation repository engineered for Microsoft 365 and Exchange Online administrators. Designed, developed, and maintained by **Roman Pindela**, this toolkit provides robust, production-tested solutions for identity administration, mailbox provisioning, archiving lifecycles, mailflow transport rules, threat protection, and security auditing.

---

## Architecture & Code Quality Assessment

### Professional Review Summary
> **Verdict: Enterprise-Grade / Production-Ready**  
> Unlike conventional ad-hoc administration scripts, this repository represents a cohesive, battle-tested software suite. Each tool reflects strict software engineering principles: clean modular functions, strict variable scoping, idempotent operations, robust pre-flight input validation, and defensive error handling.

### Key Architectural Strengths

1. **Intelligent Session Reuse & Modern MFA Authentication**:
   - Every script incorporates an active session detection algorithm (e.g., verifying `Get-OrganizationConfig` or `Get-MgContext`) before prompting for credentials. Administrators working in interactive sessions authenticate once, eliminating repetitive login prompts.
   - Built for Modern Authentication and Multi-Factor Authentication (MFA) using official `ExchangeOnlineManagement` (v3+) and `Microsoft.Graph` SDKs.
   - Automatically installs required prerequisite modules for the `CurrentUser` scope from the PowerShell Gallery if missing.

2. **Safety-First Engineering & Conflict Detection**:
   - **Pre-Flight Validation**: Validates RFC compliance of domains and email addresses before invoking cloud APIs.
   - **Conflict Detection**: Actively prevents administrative mistakes (for instance, detecting and warning when an administrator attempts to whitelist a domain that is currently on the tenant blacklist).
   - **Bilingual Confirmation Flow**: Interactive prompts support both English (`Y`/`N`) and Polish (`T`/`N`) input with silent bypass flags (`-Force`) for unattended scripting.

3. **Multi-Tier Execution Capabilities**:
   - **CLI First-Class Citizen**: Fully parameterized for CLI use, pipeline integration (`-PassThru`), and structured exports (`-ExportCsv` with UTF-8 encoding).
   - **Defender Portal Alignment**: Console outputs and summaries mirror the layout of modern Microsoft Defender and Exchange Admin Center portals.
   - **Interactive Console Wizards**: Advanced scripts offer dedicated interactive menus (`-Interactive` / `-Menu`) with numbered options for everyday operational tasks.

4. **Exemplary Documentation & Visual Grounding**:
   - Each script is housed in its own self-contained directory with a dedicated, detailed `README.md` and an `assets/` directory containing actual console execution screenshots.
   - Built-in help handlers (`-Help` / `-h` or executing without arguments) provide syntax diagrams, parameter descriptions, usage examples, and author metadata.

---

## Tool Catalog & Repository Index

| Tool / Directory | Focus Area | Primary Capabilities |
| :--- | :--- | :--- |
| [**Manage-M365TenantFilterList**](Manage-M365TenantFilterList/) | Security & Anti-Spam | Audits, searches, adds, and removes sender domains and emails across Anti-Spam inbound policies (`HostedContentFilterPolicy`) and Tenant Allow/Block Lists (`TABL`). |
| [**Get-M365Users**](Get-M365Users/) | User Auditing & Licensing | Generates detailed reports of tenant users, assigned licenses, mailbox configurations, account statuses, and SKU breakdowns. |
| [**Manage-M365UserAccount**](Manage-M365UserAccount/) | Identity & Account Lifecycle | End-to-end user management via Microsoft Graph: password resets, account locking/unlocking, MFA auditing, and Entra ID attributes. |
| [**Manage-MailboxUserAccess**](Manage-MailboxUserAccess/) | Mailbox Security & Delegation | Audits, grants, and revokes FullAccess, SendAs, and SendOnBehalf delegation permissions across Exchange Online mailboxes. |
| [**Manage-MailflowRules**](Manage-MailflowRules/) | Transport Rules & Routing | Inspects, searches, creates, toggles, and exports Exchange Online transport/mailflow rules with safety prompts. |
| [**Manage-M365Groups**](Manage-M365Groups/) | Collaboration & Distribution | Audits and manages Unified (M365) Groups, Distribution Lists, Mail-Enabled Security Groups, memberships, and owners. |
| [**Get-MailboxSizeAndArchiveDetails**](Get-MailboxSizeAndArchiveDetails/) | Storage & Capacity Planning | Generates forensic reports on mailbox sizes, item counts, quota consumption, and online archive storage statistics. |
| [**Enable-OnlineArchive**](Enable-OnlineArchive/) | Mailbox Archiving | Provisions and enables Exchange Online Archive mailboxes with status verification and Managed Folder Assistant triggers. |
| [**set-Mailbox1YearArchivePolicy**](set-Mailbox1YearArchivePolicy/) | Retention & Compliance | Creates and enforces a 1-year archive retention policy (`MoveToArchive`) and schedules migration of older messages. |
| [**Get-M365MessageTrace**](Get-M365MessageTrace/) | Diagnostics & Delivery Tracking | Queries historical message traces, tracks delivery status, inspects quarantine states, and troubleshoots spam verdicts. |
| [**Manage-AutomaticRepliesMailbox**](Manage-AutomaticRepliesMailbox/) | Mailbox Availability | Audits, schedules, configures, and disables Out-Of-Office (OOF) automatic replies and internal/external message bodies. |
| [**Send-M365Mail**](Send-M365Mail/) | Mail Dispatch & Automation | Automated Microsoft 365 email dispatcher via Microsoft Graph with certificate (App-Only) and interactive (MFA) authentication, attachments, and logging. |
| [Audit-ExchangeSecuritySummary](./Audit-ExchangeSecuritySummary/) | Non-intrusive Exchange Online security posture audit with plain-English risk rationales and standalone HTML reporting. | v1.0.0 |

---

## Detailed Tool Descriptions

### 1. Security & Mailflow Protection
- **[Audit-ExchangeSecuritySummary](Audit-ExchangeSecuritySummary/Audit-ExchangeSecuritySummary.ps1)**:
  Non-intrusive Exchange Online security posture auditor assessing modern authentication, mail routing exfiltration protections, unified audit logging, mailbox delegation sprawl, and high-privilege RBAC assignments with standalone executive HTML reporting.
- **[Manage-M365TenantFilterList](Manage-M365TenantFilterList/Manage-M365TenantFilterList.ps1)**:
  Manages tenant-wide whitelists and blacklists. Fully synchronizes Anti-Spam Inbound Policies (`AllowedSenders`, `AllowedSenderDomains`, `BlockedSenders`, `BlockedSenderDomains`) with Defender's Tenant Allow/Block List (TABL). Includes wildcard search, duplicate protection, and automated CSV reporting.
- **[Manage-MailflowRules](Manage-MailflowRules/Manage-MailflowRules.ps1)**:
  Provides comprehensive lifecycle management for tenant transport rules. Allows administrators to inspect rule priorities, enable/disable rules safely, create new routing rules, and back up rule configurations to CSV.
- **[Get-M365MessageTrace](Get-M365MessageTrace/Get-M365MessageTrace.ps1)**:
  Automates message delivery forensics. Tracks delivered, failed, quarantined, and filtered emails over specified time windows, eliminating manual queries in the Security & Compliance center.
- **[Send-M365Mail](Send-M365Mail/Send-M365Mail.ps1)**:
  Modern email dispatcher via Microsoft Graph API. Supports automated X.509 certificate authentication (App-Only daemon) or interactive sign-in with MFA, client configuration deployment, Base64 MIME attachment encoding, rich HTML bodies, and mandatory execution logging to C:\Logs.

### 2. User & Identity Administration
- **[Get-M365Users](Get-M365Users/Get-M365Users.ps1)**:
  Fast, comprehensive tenant user directory auditing. Correlates user mailbox properties with Entra ID licensing plans, sign-in states, and department hierarchies.
- **[Manage-M365UserAccount](Manage-M365UserAccount/Manage-M365UserAccount.ps1)**:
  Modern Microsoft Graph administration tool for user accounts. Safely executes password resets with temporary password generation, locks/unlocks compromised accounts, and audits sign-in logs.

### 3. Mailbox Access, Delegation & Availability
- **[Manage-MailboxUserAccess](Manage-MailboxUserAccess/Manage-MailboxUserAccess.ps1)**:
  Streamlines mailbox permission auditing and assignment. Manages `FullAccess`, `SendAs`, and `SendOnBehalf` permissions on user and shared mailboxes with clear reporting of existing delegate access.
- **[Manage-AutomaticRepliesMailbox](Manage-AutomaticRepliesMailbox/Manage-AutomaticRepliesMailbox.ps1)**:
  Configures out-of-office autoreplies for absent employees or departures. Supports immediate activation, scheduled date ranges, separate internal/external responses, and complete deactivation.

### 4. Storage, Archiving & Compliance
- **[Get-MailboxSizeAndArchiveDetails](Get-MailboxSizeAndArchiveDetails/Get-MailboxSizeAndArchiveDetails.ps1)**:
  Delivers deep capacity analytics. Measures primary mailbox size, archive usage, quota percentage, and item counts to prevent quota exhaustion.
- **[Enable-OnlineArchive](Enable-OnlineArchive/Enable-OnlineArchive.ps1)**:
  Automates the provisioning of online archives for mailboxes reaching capacity limits. Verifies licensing eligibility and triggers the assistant.
- **[set-Mailbox1YearArchivePolicy](set-Mailbox1YearArchivePolicy/set-Mailbox1YearArchivePolicy.ps1)**:
  Enforces regulatory and storage optimization policies. Automatically sets up retention tags and retention policies to move items older than 365 days into the archive.

### 5. Group Collaboration
- **[Manage-M365Groups](Manage-M365Groups/Manage-M365Groups.ps1)**:
  Provides a single pane of glass for all group types: Microsoft 365 Groups (Unified), Distribution Lists, Security Groups, and Mail-Enabled Security Groups. Audits owners, membership counts, and delivery settings.

---

## Engineering Standards & Best Practices

All scripts in this toolkit adhere to the formal **PowerShell Script Standards**:

1. **Self-Documenting Help**:
   - Each script contains complete comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, `.NOTES`).
   - Running any script without arguments displays an executive summary, usage patterns, author contact, and security instructions.
2. **Strict Mode Safety**:
   - `Set-StrictMode -Version Latest` is utilized across core logic. All pipeline collections and multi-item queries are encapsulated in array subexpressions (`@(...)`) to prevent `PropertyNotFoundException` on `.Count` checks.
3. **Defense Against Invalid Input**:
   - Strict parameter typing, `[ValidateSet()]`, `[ValidateNotNullOrEmpty()]`, and custom regex format checks ensure safe execution.
4. **Idempotency**:
   - Actions (such as creating retention policies or adding filter entries) check for existing configurations before committing changes, preventing redundant or conflicting operations.
5. **No Hardcoded Credentials**:
   - Authentication delegates strictly to Microsoft modern interactive browser/device login sessions.

---

## Getting Started

### Prerequisites
- **PowerShell**: Windows PowerShell 5.1 or PowerShell Core 7.x (Windows Desktop / Windows Server).
- **Execution Policy**: Must allow local script execution (e.g. `RemoteSigned` or `Bypass`).
- **Required Modules**:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force
  Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force
  Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
  ```
- **Administrative Roles**: Global Administrator, Exchange Administrator, or Security Administrator in Entra ID / Microsoft 365.

### Security: Unblocking Scripts
When downloading or cloning scripts from GitHub on Windows, remove the Zone.Identifier alternate data stream:
```powershell
# Unblock all scripts in the repository
Get-ChildItem -Path ".\Microsoft 365" -Recurse -Filter "*.ps1" | Unblock-File
```

---

## Future Roadmap & Recommendations

To further enhance this toolkit, the following architectural additions are recommended:
1. **PascalCase Standardization**:
   - Rename `set-Mailbox1YearArchivePolicy` to `Set-Mailbox1YearArchivePolicy` to maintain 100% naming harmony across all 11 tools.
2. **Master Launcher (`Start-M365Toolkit.ps1`)**:
   - Create a centralized root launcher script that provides a master dashboard connecting all sub-scripts into an interactive administrative console.
3. **Continuous Integration (CI) Linter**:
   - Integrate a GitHub Actions workflow utilizing `PSScriptAnalyzer` to continuously validate PowerShell formatting, syntax consistency, and strict mode compliance.

---

## Author & Contact

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **Repository**: [https://github.com/romanpindela/Microsoft-365](https://github.com/romanpindela/Microsoft-365)

