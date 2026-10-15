# Microsoft 365 PowerShell Automation Suite

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%7C%20v2.0-008272.svg)](https://learn.microsoft.com/graph/)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-v3.0%2B-0078D4.svg)](https://learn.microsoft.com/powershell/exchange/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Supported-0078D4.svg)](https://www.microsoft.com/microsoft-365/onedrive/online-cloud-storage)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A comprehensive, enterprise-grade PowerShell automation repository designed and maintained by **Roman Pindela** for Microsoft 365 administrators, cloud architects, and SecOps teams.

This repository delivers production-hardened tools adhering to strict software engineering standards, clean code architecture, defensive validation, multi-factor authentication (MFA), and unattended Entra ID App-Only automation.

---

## 🗂️ Repository Architecture & Modules

The toolkit is organized into specialized domain modules:

```text
Microsoft 365/
├── EntraID/       # Microsoft Entra ID (Azure AD) identity, licensing, and user management
├── Exchange/      # Exchange Online mailflow, archiving, security filtering, and mailbox administration
└── OneDrive/      # OneDrive for Business large file upload sessions, synchronization, and storage tools
```

---

## ☁️ Module Catalog

### 1. OneDrive for Business (`/OneDrive`)
* [**Upload-LargeFilesSession**](OneDrive/Upload-LargeFilesSession/README.md) — Uploads large files (up to 250 GB), entire directory trees, and wildcard file sets (`*.pdf`, `p?indela.pdf`, `p*.docx`, folder names) to user OneDrive libraries using Microsoft Graph `createUploadSession` chunking. Features automatic remote folder creation, dual progress bars (MB/s, ETA), UTF-8 file logging, and Entra ID App-Only non-interactive authentication.
* **Sync-OneDriveFolderAndNotify** — Synchronizes local directories to OneDrive libraries with automated alert notifications.

### 2. Exchange Online (`/Exchange`)
* [**Manage-M365TenantFilterList**](Exchange/Manage-M365TenantFilterList/README.md) — Audits, searches, adds, and removes sender domains and emails across Anti-Spam inbound policies and Tenant Allow/Block Lists (TABL).
* [**Get-MailboxSizeAndArchiveDetails**](Exchange/Get-MailboxSizeAndArchiveDetails/README.md) — Capacity auditing of primary mailboxes, archive quotas, item counts, and storage usage.
* [**Manage-AutomaticRepliesMailbox**](Exchange/Manage-AutomaticRepliesMailbox/README.md) — Auditing and scheduling internal/external Out-Of-Office (OOF) auto-replies.
* [**Manage-MailflowRules**](Exchange/Manage-MailflowRules/README.md) — Management, testing, toggling, and auditing of Exchange transport and routing rules.
* [**Manage-MailboxUserAccess**](Exchange/Manage-MailboxUserAccess/README.md) — Grants, revokes, and audits FullAccess, SendAs, and SendOnBehalf mailbox permissions.
* [**Enable-OnlineArchive**](Exchange/Enable-OnlineArchive/README.md) — Provisions Exchange Online Archive mailboxes and triggers Managed Folder Assistant.
* [**set-Mailbox1YearArchivePolicy**](Exchange/set-Mailbox1YearArchivePolicy/README.md) — Enforces retention policies for 1-year archive migration.
* [**Get-M365MessageTrace**](Exchange/Get-M365MessageTrace/README.md) — Queries message delivery traces, spam verdicts, and quarantine statuses.

### 3. Microsoft Entra ID (`/EntraID`)
* [**Manage-M365UserAccount**](EntraID/Manage-M365UserAccount/README.md) — Password resets, sign-in blocking/unblocking, active session revocation, alias management, and user auditing via Microsoft Graph.
* [**Get-M365Users**](EntraID/Get-M365Users/README.md) — Comprehensive tenant-wide user auditing, licensing SKU breakdowns, and mailbox status reporting.
* [**Manage-M365Groups**](EntraID/Manage-M365Groups/README.md) — Management and membership auditing for M365 Unified Groups, Distribution Lists, and Security Groups.

---

## 🔒 Security & Best Practices

1. **Unblock Downloaded Scripts**:
   ```powershell
   Unblock-File -Path .\OneDrive\Upload-LargeFilesSession\Upload-LargeFilesSession.ps1
   ```
2. **Defensive Input Validation**: Every script strictly validates input values against regex rules and path traversal attacks (`..`).
3. **Bilingual Compatibility**: Fully tested across both Polish and English Windows environments.
4. **App-Only Automation**: Non-interactive tools support Entra ID Application Permissions (`Files.ReadWrite.All`, `User.Read.All`) for secure unattended execution.

---

## 👤 Author & Contact

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
* **License**: [MIT License](https://opensource.org/licenses/MIT)
