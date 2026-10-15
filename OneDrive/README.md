# Microsoft 365 OneDrive Automation Toolkit

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0-008272.svg)](https://learn.microsoft.com/graph/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Enterprise-0078D4.svg)](https://www.microsoft.com/microsoft-365/onedrive/online-cloud-storage)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A specialized collection of enterprise-grade PowerShell automation tools for managing, synchronizing, and transferring files to and from **Microsoft 365 OneDrive for Business** using official **Microsoft Graph APIs** and modern Entra ID authentication.

Designed, developed, and maintained by **Roman Pindela**, these scripts solve real-world system administration challenges, including multi-gigabyte large file uploads, automated folder hierarchy provisioning, scheduled background synchronization, and headless daemon execution.

---

## 📂 Tools Catalog

| Script / Directory | Primary Focus | Key Capabilities |
| :--- | :--- | :--- |
| [**Upload-LargeFilesSession**](Upload-LargeFilesSession/README.md) | High-Performance Cloud File Upload | Uses Microsoft Graph `createUploadSession` chunking to upload files of any size (up to 250 GB), directory trees, and wildcard file sets (`*.pdf`, `p?indela.pdf`, `p*.docx`, folder names) to user OneDrive libraries. Features auto-creation of missing remote folders, real-time dual progress bars (speed in MB/s, ETA), UTF-8 file logging, and non-interactive Entra ID App-Only authentication. |
| **Sync-OneDriveFolderAndNotify** | Automated Folder Synchronization | Monitors, synchronizes local folder changes to OneDrive for Business, and dispatches administrative alerts and execution notifications. |

---

## 🚀 Featured Tool: Upload-LargeFilesSession

### Overview
[`Upload-LargeFilesSession.ps1`](Upload-LargeFilesSession/Upload-LargeFilesSession.ps1) is engineered for robust file transfers to OneDrive for Business:
* **Upload Sessions via Microsoft Graph**: Chunks files larger than 4 MB (and up to 250 GB) in compliant 320 KiB multiples with automatic retry on throttling (HTTP 429).
* **Missing Folder Provisioning**: Automatically inspects and creates missing parent and child folders in OneDrive if the destination folder path does not exist.
* **Flexible Input & Wildcards**: Accepts individual files, folders, or pattern masks (e.g. `*.pdf`, `p?indela.pdf`, `p*.docx`, or `ten_folder`), with optional directory recursion (`-Recurse`) and local folder tree replication (`-PreserveStructure`).
* **Dual Progress Bars & Telemetry**: Detailed real-time progress indicators displaying overall transferred bytes, active chunk progress, instantaneous transfer throughput (MB/s), and estimated completion time (ETA).
* **Entra ID App-Only Non-Interactive Mode**: Full support for headless execution in Task Scheduler, CI/CD, and server jobs via Client ID, Tenant ID, and Client Secret or Certificate Thumbprint.
* **Defensive Clean Code & Logging**: Strict parameter validation, input sanitization against command injection, and persistent timestamped UTF-8 file logging.

### Quick Usage Example
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

For complete documentation, syntax reference, and step-by-step Entra ID setup instructions, see the [Upload-LargeFilesSession README](Upload-LargeFilesSession/README.md).

---

## 👤 Author & Contact

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
