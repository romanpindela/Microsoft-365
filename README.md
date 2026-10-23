# Microsoft 365 PowerShell Automation Toolkit

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2F%20v2.0-008272.svg)](https://learn.microsoft.com/graph/)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-v3.0%2B-0078D4.svg)](https://learn.microsoft.com/powershell/exchange/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Cloud%20Storage-0078D4.svg)](https://onedrive.live.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A centralized, enterprise-grade PowerShell automation suite engineered for **Microsoft 365**, **Entra ID**, **Exchange Online**, and **OneDrive for Business** administration. Designed, developed, and maintained by **Roman Pindela**, this repository provides battle-tested, production-ready solutions for identity management, security auditing, mailflow hygiene, mailbox provisioning, large file uploads, and cloud synchronization.

---

## 🏛️ Repository Architecture

This toolkit is organized into modular functional domains:

`	ext
Microsoft 365/
├── EntraID/                         # Identity, user accounts, licensing, and security groups
│   ├── Get-M365Users/               # Comprehensive user auditing and executive license summary
│   ├── Manage-M365Groups/           # Lifecycle management of M365 Groups, DLs, and Security Groups
│   └── Manage-M365UserAccount/      # User password reset, account lock/unlock, and alias control
├── Exchange/                        # Exchange Online administration, mailflow, and archiving
│   ├── Enable-OnlineArchive/        # Provisioning online archive mailboxes
│   ├── Get-MailboxSizeAndArchive... # Deep mailbox storage forensics and quota monitoring
│   ├── Get-M365MessageTrace/        # Message delivery tracking and quarantine diagnostics
│   ├── Manage-AutomaticReplies...   # Out-Of-Office (OOF) autoreply scheduling
│   ├── Manage-MailboxUserAccess/    # FullAccess, SendAs, and SendOnBehalf delegation audits
│   ├── Manage-MailflowRules/        # Transport rules lifecycle and backup
│   ├── Manage-M365TenantFilterList/ # Anti-Spam whitelists/blacklists and TABL synchronization
│   └── set-Mailbox1YearArchive...   # Retention policy configuration
├── Helpdesk/                        # Helpdesk troubleshooting, endpoint remediation, and cache resets
│   └── Reset-TeamsIdentityCache     # Reset Teams caches, WAM token broker, and identity credentials
├── OneDrive/                        # Cloud storage administration, synchronization, and uploads
│   ├── Sync-OneDriveFolderAndNotify # Cloud-to-local and local-to-cloud sync with progress and notifications
│   └── Upload-LargeFilesSession     # Chunked Graph upload session for files up to 250 GB
├── Send-M365Mail/                   # Microsoft Graph automated email dispatcher with certificate & interactive auth
└── SharePoint/                      # SharePoint Online governance, auditing, and synchronization
    ├── Audit-SharePoint-Site        # Deep permission auditing, broken inheritance detection, and export
    ├── Manage-SharePoint-Site       # Full site administration, role modification, and user management
    ├── New-SharePoint-Site          # Modern preview and empty site provisioning
    └── Sync-SharePointLibrary       # Automatic mapping of SharePoint libraries via OneDrive sync client
```

---

## ☁️ Module Catalogs & Featured Tools

### 1. OneDrive for Business (/OneDrive)
* [**Sync-OneDriveFolderAndNotify**](OneDrive/Sync-OneDriveFolderAndNotify/README.md) — Synchronizes files between an Office 365 user's OneDrive for Business folder and a local filesystem directory. Features multi-directional sync (TwoWay, Upload, Download), automatic creation of missing cloud/local folders, granular Write-Progress status bars, multi-chunk upload sessions (>4 MB), persistent logging, native Windows balloon notifications, and headless Entra ID daemon authentication.
* [**Upload-LargeFilesSession**](OneDrive/Upload-LargeFilesSession/README.md) — Uploads large files (up to 250 GB), entire directory trees, and wildcard file sets (*.pdf, p?indela.pdf, p*.docx, folder names) to user OneDrive libraries using Microsoft Graph createUploadSession chunking. Features automatic remote folder creation, dual progress bars (MB/s, ETA), UTF-8 file logging, and Entra ID App-Only non-interactive authentication. Accompanied by three in-depth PDF guides: Graph SDK Compendium, Security Hardening Guide, and X.509 Certificate Walkthrough.

### 2. Exchange Online (/Exchange)
* [**Manage-M365TenantFilterList**](Exchange/Manage-M365TenantFilterList/README.md) — Audits, searches, adds, and removes sender domains and emails across Anti-Spam inbound policies and Tenant Allow/Block Lists (TABL).
* [**Get-MailboxSizeAndArchiveDetails**](Exchange/Get-MailboxSizeAndArchiveDetails/README.md) — Capacity auditing of primary mailboxes, archive quotas, item counts, and storage usage.
* [**Manage-AutomaticRepliesMailbox**](Exchange/Manage-AutomaticRepliesMailbox/README.md) — Auditing and scheduling internal/external Out-Of-Office (OOF) auto-replies.
* [**Manage-MailflowRules**](Exchange/Manage-MailflowRules/README.md) — Management, testing, toggling, and auditing of Exchange transport and routing rules.
* [**Manage-MailboxUserAccess**](Exchange/Manage-MailboxUserAccess/README.md) — Grants, revokes, and audits FullAccess, SendAs, and SendOnBehalf mailbox permissions.
* [**Enable-OnlineArchive**](Exchange/Enable-OnlineArchive/README.md) — Provisions Exchange Online Archive mailboxes and triggers Managed Folder Assistant.
* [**set-Mailbox1YearArchivePolicy**](Exchange/set-Mailbox1YearArchivePolicy/README.md) — Enforces retention policies for 1-year archive migration.
* [**Get-M365MessageTrace**](Exchange/Get-M365MessageTrace/README.md) — Queries message delivery traces, spam verdicts, and quarantine statuses.
* [**Send-M365Mail**](Exchange/Send-M365Mail/README.md) — Enterprise email dispatcher for Microsoft 365 / Exchange Online via Microsoft Graph API. Supports automated X.509 certificate authentication (App-Only daemon) or interactive sign-in with MFA, client configuration deployment, Base64 MIME attachment encoding, rich HTML bodies, and mandatory execution logging to `C:\Logs`.

### 3. Microsoft Entra ID (/EntraID)
* [**Manage-M365UserAccount**](EntraID/Manage-M365UserAccount/README.md) — Password resets, sign-in blocking/unblocking, active session revocation, alias management, and user auditing via Microsoft Graph.
* [**Get-M365Users**](EntraID/Get-M365Users/README.md) — Comprehensive tenant-wide user auditing, licensing SKU breakdowns, and mailbox status reporting.
* [**Manage-M365Groups**](EntraID/Manage-M365Groups/README.md) — Management and membership auditing for M365 Unified Groups, Distribution Lists, and Security Groups.

### 4. Helpdesk & Endpoint Remediation (/Helpdesk)
* [**Reset-TeamsIdentityCache**](Helpdesk/Reset-TeamsIdentityCache/README.md) — Comprehensive remediation tool for Microsoft Teams authentication loops, token expiration, stale tenant profiles, and WAM/IdentityCache corruption. Purges New Teams (MSIX) and Classic Teams caches, OneAuth tokens, Web Account Manager broker state, and Windows Credential Manager entries with detailed feature cataloging (-ListFeatures), simulation mode (-WhatIf), and structured logging.

### 5. SharePoint Online (`/SharePoint`)
* [**Audit-SharePoint-Site**](SharePoint/Audit-SharePoint-Site/README.md) — Audits site, library, folder, and item-level permissions with broken inheritance detection, nested group expansion, and automated CSV/JSON exports.
* [**Manage-SharePoint-Site**](SharePoint/Manage-SharePoint-Site/README.md) — Automates SharePoint Online site administration: site inventory listing, categorized permissions auditing, defining site modifiers, changing user roles, and safe user removal.
* [**New-SharePoint-Site**](SharePoint/New-SharePoint-Site/README.md) — Provisions modern Company-Wide preview sites or isolated empty sites with designated owners, members, and visitors.
* [**Sync-SharePointLibrary**](SharePoint/Sync-SharePointLibrary/README.md) — Automates SharePoint Online document library mapping and synchronization to Windows 10/11 File Explorer using native OneDrive sync client (`odopen://sync?`), supporting tenant auto-discovery, interactive device code login, and non-interactive Entra ID app credentials.

### 6. Email & Dispatching (/Send-M365Mail)
* [**Send-M365Mail**](Exchange/Send-M365Mail/README.md) — Enterprise email dispatcher for Microsoft 365 / Exchange Online via Microsoft Graph API. Supports automated X.509 certificate authentication (App-Only daemon) or interactive sign-in with MFA, client configuration deployment, Base64 MIME attachment encoding, rich HTML bodies, and mandatory execution logging to C:\Logs.

---

## 🚀 Quick Start: Sync-OneDriveFolderAndNotify

`powershell
# 1. Unblock downloaded script
Unblock-File -Path .\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1

# 2. View built-in help and Entra ID daemon configuration guide
.\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1 -Help

# 3. Execute two-way synchronization with auto-folder creation
.\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "john.doe@company.com" 
    -OneDriveFolderPath "ProjectFiles" 
    -LocalFolderPath "C:\Data\ProjectFiles" 
    -Direction TwoWay 
    -AutoCreate

# 4. Execute unattended background backup via Entra ID Client Secret
.\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "backup@company.com" 
    -OneDriveFolderPath "DailyBackups/SQL" 
    -LocalFolderPath "D:\SQLBackups" 
    -Direction Upload 
    -AutoCreate 
    -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" 
    -ClientId "11111111-2222-3333-4444-555555555555" 
    -ClientSecret "SecretValueHere"
`

---

## 🛠️ Prerequisites & Installation

- **PowerShell**: Version 5.1+ (Windows PowerShell) or PowerShell 7.x (Core).
- **Core Modules**:
  `powershell
  # Exchange Online Management
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force

  # Microsoft Graph Authentication & Users
  Install-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force
  `

---

## 🛡️ Security Best Practices

1. **Unblock Downloaded Scripts**:
   `powershell
   Get-ChildItem -Path . -Filter *.ps1 -Recurse | Unblock-File
   `
2. **Defensive Input Validation**:
   All scripts in this repository implement strict input sanitization, directory traversal protection (..), and RFC email format validation.
3. **Language-Independent Execution**:
   Scripts and confirmation prompts support both English (Y/N) and Polish (T/N) operational environments.

---

## 👨‍💻 Author & Contact

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **License**: MIT

