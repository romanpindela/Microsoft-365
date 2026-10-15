# Microsoft 365 OneDrive for Business Automation Toolkit

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%2B-008272.svg)](https://learn.microsoft.com/graph/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Cloud%20Storage-0078D4.svg)](https://onedrive.live.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A centralized, enterprise-grade PowerShell automation suite engineered for Microsoft 365 OneDrive for Business administration, cloud synchronization, high-performance file uploads, and automated data protection. Designed, developed, and maintained by **Roman Pindela**, this toolkit provides robust, production-tested solutions for cloud storage administration, scheduled synchronization, and backup automation.

---

## 🏗️ Architecture & Code Quality Assessment

### Professional Review Summary
> **Verdict: Enterprise-Grade / Production-Ready**  
> Rather than relying on simple file-copy utilities or unmonitored scripts, this repository provides battle-tested automation software. Tools adhere strictly to software engineering best practices: Clean Code, defensive error handling, multi-chunk RFC-compliant Graph upload sessions, and comprehensive logging.

### Key Architectural Strengths

1. **Dual Modern Authentication (Interactive & Entra ID Daemon)**:
   - Supports interactive administrator logins with modern MFA and automatic active session detection via `Get-MgContext`.
   - Native support for headless, non-interactive execution via **Microsoft Entra ID App Registrations** using Client Secrets or X.509 Certificates (ideal for Windows Task Scheduler and CI/CD).
   - Resilient execution with a built-in direct REST OAuth2 fallback for environments where Microsoft Graph SDK modules cannot be installed.

2. **Enterprise Reliability & Conflict Handling**:
   - **Smart Folder Discovery & Auto-Creation**: Pre-flight validation checks both local and cloud directories. If missing, it provides bilingual confirmation prompts (`[Y/N]` / `[T/N]`) or provisions them automatically (`-AutoCreate`, `-AutoCreateOneDriveFolder`, `-AutoCreateLocalFolder`).
   - **Large File Streaming**: Bypasses Graph API 4 MB upload thresholds by leveraging multi-chunk Upload Sessions (RFC compliant) for files up to 250 GB.

3. **Multi-Channel Observability & Progress Reporting**:
   - Multi-tier visual feedback using PowerShell's native `Write-Progress` (tracking overall file progress, current file, bytes transferred, transfer throughput in MB/s, and upload session chunks).
   - Persistent timestamped UTF-8 log files with structured severity levels (`[INFO]`, `[SUCCESS]`, `[WARNING]`, `[ERROR]`, `[PROGRESS]`).
   - Native Windows desktop notifications (balloon toast) and Microsoft Teams / Slack incoming webhook support.

4. **Standards & Language-Neutral Engineering**:
   - Conforms strictly to the official PowerShell Script Standards.
   - Built-in help handlers (`-Help` / `-h` or zero arguments) providing complete syntax, parameter definitions, and detailed Entra ID configuration guides.
   - Language-independent input processing compatible with both English and Polish Windows operating systems.

---

## 📦 Tool Catalog & Category Index

| Tool / Directory | Focus Area | Primary Capabilities |
| :--- | :--- | :--- |
| [**Sync-OneDriveFolderAndNotify**](Sync-OneDriveFolderAndNotify/) | Synchronization & Backup | Synchronizes cloud OneDrive folders with local filesystem directories (TwoWay, Upload, Download), chunked large file uploads (>4MB), real-time progress bars, automated folder creation, persistent logging, and desktop/webhook notifications. |
| [**Upload-LargeFilesSession**](Upload-LargeFilesSession/) | High-Performance Cloud File Upload | Uses Microsoft Graph `createUploadSession` chunking to upload files of any size (up to 250 GB), directory trees, and wildcard file sets (`*.pdf`, `p?indela.pdf`, `p*.docx`, folder names) to user OneDrive libraries. Features auto-creation of missing remote folders, real-time dual progress bars (speed in MB/s, ETA), UTF-8 file logging, and non-interactive Entra ID App-Only authentication. |

---

## 🛠️ Featured Script Highlights

### 1. [Sync-OneDriveFolderAndNotify](Sync-OneDriveFolderAndNotify/README.md)
* **Script**: [`Sync-OneDriveFolderAndNotify.ps1`](Sync-OneDriveFolderAndNotify/Sync-OneDriveFolderAndNotify.ps1)
* **Author**: Roman Pindela (`roman.pindela@gmail.com`)
* **Description**: Enterprise-grade cloud synchronizer and backup engine for Microsoft 365 OneDrive for Business accounts.
  * **Multi-Directional Modes**: `TwoWay` (bidirectional sync), `Upload` (local to cloud), or `Download` (cloud to local).
  * **Auto-Folder Provisioning**: Detects missing remote or local folders and auto-creates them on demand (`-AutoCreate`).
  * **Large File Uploads**: Chunked upload session handler for files larger than 4 MB (up to 250 GB).
  * **Unattended Daemon Execution**: Native support for Microsoft Entra ID App Registrations using Client Secret (`-ClientSecret`) or Certificate Thumbprint (`-CertificateThumbprint`).
  * **Notifications**: Native Windows balloon tips and Microsoft Teams / Slack webhook alerts.
* **Quick Example**:
  ```powershell
  # Unattended upload backup to OneDrive with auto-folder creation
  .\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "backup@company.com" `
      -OneDriveFolderPath "DailyBackups/SQL" `
      -LocalFolderPath "D:\SQLBackups" `
      -Direction Upload `
      -AutoCreate `
      -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
      -ClientId "11111111-2222-3333-4444-555555555555" `
      -ClientSecret "SecretValueHere"
  ```

### 2. [Upload-LargeFilesSession](Upload-LargeFilesSession/README.md)
* **Script**: [`Upload-LargeFilesSession.ps1`](Upload-LargeFilesSession/Upload-LargeFilesSession.ps1)
* **Author**: Roman Pindela (`roman.pindela@gmail.com`)
* **Description**: High-speed chunked file uploader for Microsoft 365 OneDrive for Business accounts.
  * **Upload Sessions via Microsoft Graph**: Chunks files larger than 4 MB (and up to 250 GB) in compliant 320 KiB multiples with automatic retry on throttling (HTTP 429).
  * **Missing Folder Provisioning**: Automatically inspects and creates missing parent and child folders in OneDrive if the destination folder path does not exist.
  * **Flexible Input & Wildcards**: Accepts individual files, folders, or pattern masks (e.g. `*.pdf`, `p?indela.pdf`, `p*.docx`, or `ten_folder`), with optional directory recursion (`-Recurse`) and local folder tree replication (`-PreserveStructure`).
  * **Dual Progress Bars & Telemetry**: Detailed real-time progress indicators displaying overall transferred bytes, active chunk progress, instantaneous transfer throughput (MB/s), and estimated completion time (ETA).
  * **Accompanying PDF Guides**:
    * 📕 [Kompendium Microsoft Graph SDK i PowerShell (PDF)](Upload-LargeFilesSession/Kompendium%20Microsoft%20Graph%20SDK%20i%20PowerShell_%20Automatyzacja%20OneDrive%20i%20Exchange.pdf)
    * 📕 [Przewodnik Bezpieczeństwa - Upload-LargeFilesSession.ps1 (PDF)](Upload-LargeFilesSession/Przewodnik%20Bezpiecze%C5%84stwa%20-%20Upload-LargeFilesSession.ps1.pdf)
    * 📕 [Przewodnik Certyfikatów X.509 - Upload-LargeFilesSession.ps1 (PDF)](Upload-LargeFilesSession/Przewodnik%20Certyfikat%C3%B3w%20X.509%20-%20Upload-LargeFilesSession.ps1.pdf)
* **Quick Example**:
  ```powershell
  # Upload all PDF files matching mask *.pdf to user's OneDrive with Entra ID App Credentials:
  .\Upload-LargeFilesSession\Upload-LargeFilesSession.ps1 `
      -UserPrincipalName "user@company.com" `
      -TargetFolder "Backups/2026" `
      -Path "C:\Data\*.pdf" `
      -TenantId "11111111-2222-3333-4444-555555555555" `
      -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
      -ClientSecret "SecretValueHere..." `
      -ChunkSizeMB 20
  ```

---

## 🛡️ Security & Prerequisites

- **PowerShell**: Version 5.1+ (Windows PowerShell) or PowerShell 7+ (Core).
- **Module Requirements**:
  ```powershell
  Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
  ```
- **Unblocking Downloaded Scripts**:
  ```powershell
  Unblock-File -Path .\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1
  Unblock-File -Path .\Upload-LargeFilesSession\Upload-LargeFilesSession.ps1
  ```

---

## 👨‍💻 Author & Support

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **License**: MIT
