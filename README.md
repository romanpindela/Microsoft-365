# Microsoft 365 PowerShell Automation Toolkit

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2F%20v2.0-008272.svg)](https://learn.microsoft.com/graph/)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-v3.0%2B-0078D4.svg)](https://learn.microsoft.com/powershell/exchange/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A centralized, enterprise-grade PowerShell automation suite engineered for **Microsoft 365**, **Entra ID**, **Exchange Online**, and **OneDrive for Business** administration. Designed, developed, and maintained by **Roman Pindela**, this repository provides battle-tested, production-ready solutions for identity management, security auditing, mailflow hygiene, mailbox provisioning, and cloud synchronization.

---

## 🏛️ Repository Architecture

This toolkit is organized into modular functional domains:

```text
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
└── OneDrive/                        # Cloud storage administration and synchronization
    ├── Upload-LargeFilesSession     # Chunked Graph upload sessions for large files up to 250 GB and wildcards
    └── Sync-OneDriveFolderAndNotify # Cloud-to-local and local-to-cloud sync with progress and notifications
```

---

## 🚀 Featured Tools

### [Upload-LargeFilesSession](OneDrive/Upload-LargeFilesSession/README.md)
* **Path**: [`OneDrive/Upload-LargeFilesSession/Upload-LargeFilesSession.ps1`](OneDrive/Upload-LargeFilesSession/Upload-LargeFilesSession.ps1)
* **Author**: Roman Pindela (`roman.pindela@gmail.com`)
* **Focus**: High-performance multi-gigabyte upload sessions (up to 250 GB) to OneDrive for Business using Microsoft Graph chunking.
* **Capabilities**:
  * Uploads individual files, directory trees, or wildcard sets (`*.pdf`, `p?indela.pdf`, `p*.docx`, folder names).
  * Auto-provisions missing remote folder hierarchies in OneDrive.
  * Real-time dual progress bars (MB/s speed, ETA, chunk byte ranges).
  * Unattended daemon execution via Entra ID App Registrations (Client Secret or Certificate).

### [Sync-OneDriveFolderAndNotify](OneDrive/Sync-OneDriveFolderAndNotify/README.md)
* **Path**: [`OneDrive/Sync-OneDriveFolderAndNotify/Sync-OneDriveFolderAndNotify.ps1`](OneDrive/Sync-OneDriveFolderAndNotify/Sync-OneDriveFolderAndNotify.ps1)
* **Author**: Roman Pindela (`roman.pindela@gmail.com`)
* **Focus**: Automated file synchronization between Microsoft 365 OneDrive accounts and local storage.
* **Capabilities**:
  * Multi-directional sync (`TwoWay`, `Upload`, `Download`).
  * Auto-folder creation, balloon notifications, and Teams/Slack webhooks.

### Key Highlights:
* **Interactive & Daemon Authentication**: Supports interactive logins with modern MFA session reuse or headless, non-interactive execution via **Microsoft Entra ID App Registrations** (Client Secret or X.509 Certificate).
* **Flexible Directionality**: Synchronize in `TwoWay` (bidirectional), `Upload` (local to cloud), or `Download` (cloud to local) modes.
* **Smart Folder Handling**: Detects missing OneDrive or local directories, prompts interactively, or auto-provisions them with `-AutoCreate`.
* **Large File Chunking**: Handles files larger than 4 MB (up to 250 GB) using official Graph chunked upload sessions with nested progress bars.
* **Granular Observability**: Multi-level `Write-Progress` status bars, persistent timestamped UTF-8 log files, native Windows balloon notifications, and Microsoft Teams / Slack webhook support.

#### Quick Example:
```powershell
# Unblock after cloning
Unblock-File -Path .\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1

# Execute unattended backup to OneDrive
.\OneDrive\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "backup@company.com" `
    -OneDriveFolderPath "DailyBackups/SQL" `
    -LocalFolderPath "D:\SQLBackups" `
    -Direction Upload `
    -AutoCreate `
    -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
    -ClientId "11111111-2222-3333-4444-555555555555" `
    -ClientSecret "SecretValueHere"
```

---

## 🛠️ Prerequisites & Installation

- **PowerShell**: Version 5.1+ (Windows PowerShell) or PowerShell 7.x (Core).
- **Core Modules**:
  ```powershell
  # Exchange Online Management
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force

  # Microsoft Graph Authentication & Users
  Install-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force
  ```

---

## 🛡️ Security Best Practices

1. **Unblock Downloaded Scripts**:
   ```powershell
   Get-ChildItem -Path . -Filter *.ps1 -Recurse | Unblock-File
   ```
2. **Defensive Input Validation**:
   All scripts in this repository implement strict input sanitization, directory traversal protection, and RFC email format validation.
3. **Language-Independent Execution**:
   Scripts and prompts support both English (`Y`/`N`) and Polish (`T`/`N`) operational environments.

---

## 👨‍💻 Author & Contact

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **License**: MIT
