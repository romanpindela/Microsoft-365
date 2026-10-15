# Sync-OneDriveFolderAndNotify

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0-008272.svg)](https://learn.microsoft.com/graph/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Enterprise-0078D4.svg)](https://onedrive.live.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

An enterprise-grade PowerShell file synchronization and backup utility designed for **Microsoft 365** and **OneDrive for Business**. Developed and maintained by **Roman Pindela**, this tool bridges cloud storage and local filesystem environments with multi-chunk parallel upload sessions, granular progress reporting, automated missing folder creation, persistent logging, native desktop notifications, and enterprise Entra ID daemon authentication.

---

## 📋 Overview

`Sync-OneDriveFolderAndNotify.ps1` enables administrators, power users, and automated background jobs to reliably synchronize files between any Office 365 user's OneDrive for Business account and a local folder directory. 

Built on top of modern **Microsoft Graph API / SDK**, the tool eliminates common hurdles in OneDrive automation—such as complex MFA interactive login loops, large file upload limitations (>4 MB), missing cloud/local folder exceptions, and lack of visibility during file transfers.

### Key Capabilities:
1. **Interactive & Non-Interactive Daemon Execution**: Run interactively with browser/device-code authentication or fully unattended via Entra ID App Registration credentials (Client Secret or X.509 Certificate).
2. **Flexible Multi-Directional Sync**:
   - `TwoWay` (Default): Bidirectional sync. Synchronizes new and modified files in both directions; newer file timestamps supersede older copies.
   - `Upload`: One-way sync. Uploads new and updated local files into OneDrive.
   - `Download`: One-way sync. Downloads new and updated OneDrive cloud files to the local directory.
3. **Smart Folder Detection & Auto-Creation**: Detects missing remote or local folders, displays visual warnings, and prompts the user or automatically provisions the required folder hierarchy (`-AutoCreate`, `-AutoCreateOneDriveFolder`, `-AutoCreateLocalFolder`).
4. **RFC-Compliant Chunked Upload Sessions**: Automatically handles large files (> 4 MB up to 250 GB) using official Graph upload sessions in 3.2 MB byte chunks with continuous nested progress reporting.
5. **Real-Time Visual Progress Indicators**: Granular `Write-Progress` status bars providing file count, overall completion percentage, active file name, transfer size, and sub-chunk progress.
6. **Enterprise Logging**: Dual-output logging system printing color-coded status messages to the console while appending structured, timestamped UTF-8 records to a persistent log file (`-LogPath`).
7. **Multi-Channel Notifications**: Fires native Windows desktop notifications (balloon toast) upon job completion or failure, and optionally posts formatted status cards to Microsoft Teams or Slack incoming webhooks (`-WebhookUrl`).
8. **Defensive Input Sanitization**: Prevents malicious directory traversal attacks (`..`), filters illegal path characters, and validates email/UPN formats.

---

## 🔐 How to Configure Entra ID for Non-Interactive Authentication

To execute this script unattended (e.g., via **Windows Task Scheduler**, Azure DevOps, GitHub Actions, or corporate backup pipelines) without interactive MFA prompts, set up an Entra ID App Registration:

### Step 1: Register an Application in Microsoft Entra ID
1. Log in to the [Microsoft Entra admin center](https://entra.microsoft.com).
2. Navigate to **Identity** > **Applications** > **App registrations** > **New registration**.
3. Set the application name (e.g. `OneDrive-Sync-Daemon`).
4. Select **Accounts in this organizational directory only (Single tenant)**.
5. Leave the **Redirect URI** blank.
6. Click **Register**.
7. From the application **Overview** page, copy:
   - **Application (client) ID** (e.g., `11111111-2222-3333-4444-555555555555`)
   - **Directory (tenant) ID** (e.g., `aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee`)

### Step 2: Grant Microsoft Graph API Permissions
1. In your registered app, go to **Manage** > **API permissions**.
2. Click **Add a permission** > Select **Microsoft Graph**.
3. Choose **Application permissions** *(Important: Do NOT select Delegated permissions for unattended scripts)*.
4. Search for and add the following permissions:
   - `Files.ReadWrite.All` — Required to read and write files across all user OneDrive drives.
   - `User.Read.All` — Required to resolve target user identities and locate user drive endpoints by email/UPN.
5. Click **Add permissions**.
6. Click **Grant admin consent for <Your Tenant>** and confirm with **Yes**. Ensure the status shows green checkmarks.

### Step 3: Configure Authentication Credentials

#### Option A: Client Secret (Recommended for fast deployment)
1. Go to **Manage** > **Certificates & secrets** > **Client secrets** tab.
2. Click **New client secret**, enter a description (e.g., `OneDriveSyncSecret`), and choose an expiration period.
3. Click **Add**.
4. **Immediately copy the Secret Value** (it will be permanently masked once you leave the blade).
5. Pass this secret to the script using `-TenantId`, `-ClientId`, and `-ClientSecret`.

#### Option B: X.509 Certificate (Recommended for high-security production)
1. Generate a self-signed certificate in PowerShell:
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=OneDriveSyncDaemon" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable
   Export-Certificate -Cert $cert -FilePath "C:\Certs\OneDriveSyncDaemon.cer"
   ```
2. In your Entra ID App, go to **Manage** > **Certificates & secrets** > **Certificates** tab.
3. Click **Upload certificate** and upload `OneDriveSyncDaemon.cer`.
4. Run the script using `-TenantId`, `-ClientId`, and `-CertificateThumbprint $cert.Thumbprint`.

---

## 🛠️ Prerequisites

- **Operating System**: Windows 10, Windows 11, Windows Server 2016 / 2019 / 2022 / 2025.
- **PowerShell Version**: PowerShell 5.1 (Built into Windows) or PowerShell 7.x (Core).
- **Microsoft Graph SDK Module**:
  The script automatically checks for `Microsoft.Graph.Authentication`. If missing during interactive runs, it will offer to install it from PSGallery:
  ```powershell
  Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
  ```
  *(Note: For non-interactive Client Secret runs, the script also includes a standalone native REST fallback that can operate even without the SDK module installed!)*

---

## 🛡️ Security & Script Unblocking

When cloning or downloading PowerShell scripts from GitHub, Windows places an execution block on them. Before running, unblock the script:

```powershell
Unblock-File -Path .\Sync-OneDriveFolderAndNotify.ps1
```

---

## 🚀 Usage Examples

### 1. View Built-in Help & Syntax Reference
Running the script without parameters or with `-Help` / `-h` immediately displays comprehensive documentation:

```powershell
.\Sync-OneDriveFolderAndNotify.ps1
# Or explicitly:
.\Sync-OneDriveFolderAndNotify.ps1 -h
.\Sync-OneDriveFolderAndNotify.ps1 -Help
```

### 2. Interactive Two-Way Synchronization (with Auto-Folder Creation)
Synchronizes files between the user's OneDrive folder `ProjectFiles` and local folder `C:\Data\ProjectFiles`. If folders are missing, `-AutoCreate` automatically provisions them:

```powershell
.\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "john.doe@company.com" `
    -OneDriveFolderPath "ProjectFiles" `
    -LocalFolderPath "C:\Data\ProjectFiles" `
    -Direction TwoWay `
    -AutoCreate
```

### 3. Unattended Nightly Backup (Upload Only) via Entra ID Client Secret
Runs silently in the background, uploads modified/new local SQL backup dumps to the user's cloud folder, and sends a notification to Microsoft Teams:

```powershell
.\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "backup.service@company.com" `
    -OneDriveFolderPath "DailyBackups/SQL" `
    -LocalFolderPath "D:\SQLBackups" `
    -Direction Upload `
    -AutoCreate `
    -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
    -ClientId "11111111-2222-3333-4444-555555555555" `
    -ClientSecret "YourClientSecretValueHere" `
    -WebhookUrl "https://outlook.office.com/webhook/..."
```

### 4. One-Way Cloud Download with Custom Log Path
Downloads updated and new reports from OneDrive into a local reporting cache, writing logs to a specific drive:

```powershell
.\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "sarah.smith@company.com" `
    -OneDriveFolderPath "Reports/2026" `
    -LocalFolderPath "C:\Reports\2026" `
    -Direction Download `
    -AutoCreateLocalFolder `
    -LogPath "C:\Logs\OneDriveSync_Reports.log"
```

### 5. High-Security Unattended Sync via Installed Certificate
Uses an installed X.509 Certificate in the Windows Certificate Store (`Cert:\CurrentUser\My`) to authenticate without storing plaintext passwords or secrets:

```powershell
.\Sync-OneDriveFolderAndNotify.ps1 -UserPrincipalName "finance.user@company.com" `
    -OneDriveFolderPath "Audits" `
    -LocalFolderPath "E:\FinanceAudits" `
    -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
    -ClientId "11111111-2222-3333-4444-555555555555" `
    -CertificateThumbprint "A1B2C3D4E5F60718293A4B5C6D7E8F9012345678" `
    -AutoCreate
```

### 6. Windows Task Scheduler Automation (Example Command)
To schedule a nightly sync at 02:00 AM, configure Task Scheduler Action:
- **Program/script**: `powershell.exe`
- **Add arguments**:
  ```text
  -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Sync-OneDriveFolderAndNotify\Sync-OneDriveFolderAndNotify.ps1" -UserPrincipalName "backup@company.com" -OneDriveFolderPath "Nightly" -LocalFolderPath "D:\Nightly" -Direction Upload -AutoCreate -TenantId "..." -ClientId "..." -ClientSecret "..."
  ```

---

## ⚙️ Parameter Reference

| Parameter | Type | Aliases | Description |
| :--- | :--- | :--- | :--- |
| `-UserPrincipalName` | `string` | `-User`, `-UPN`, `-Email`, `-Account` | **Required**. Office 365 UPN or email owning the target OneDrive. |
| `-OneDriveFolderPath` | `string` | `-OneDriveFolder`, `-RemoteFolder`, `-CloudFolder`, `-CloudPath` | Target folder in OneDrive (e.g. `'Backups/SQL'`). Specify `'/'` or leave empty for root. |
| `-LocalFolderPath` | `string` | `-LocalFolder`, `-LocalPath`, `-Path` | Target local filesystem directory path. |
| `-Direction` | `string` | `-SyncDirection`, `-Mode` | Sync direction: `'TwoWay'` (default), `'Upload'`, or `'Download'`. |
| `-AutoCreate` | `switch` | — | Automatically provisions missing OneDrive and Local folders without prompting. |
| `-AutoCreateOneDriveFolder` | `switch` | `-AutoCreateRemoteFolder` | Automatically creates missing OneDrive folder in cloud. |
| `-AutoCreateLocalFolder` | `switch` | — | Automatically creates missing local directory on disk. |
| `-Recurse` | `bool` | — | Recurse into subdirectories (Default: `$true`). |
| `-TenantId` | `string` | `-Tenant`, `-DirectoryId` | Entra ID Directory (Tenant) ID for unattended daemon login. |
| `-ClientId` | `string` | `-AppId`, `-ApplicationId` | Entra ID Application (Client) ID for unattended daemon login. |
| `-ClientSecret` | `string` | `-Secret` | Entra ID Client Secret string for daemon authentication. |
| `-ClientSecretSecure` | `SecureString` | — | Optional SecureString representation of client secret. |
| `-CertificateThumbprint`| `string` | `-Thumbprint` | X.509 Certificate thumbprint installed in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`. |
| `-CertificateFilePath` | `string` | `-CertPath` | Path to a `.cer` or `.pfx` certificate file. |
| `-LogPath` | `string` | — | Custom path to store execution log file (defaults to `.\logs\Sync-OneDriveFolderAndNotify_<timestamp>.log`). |
| `-NoLog` | `switch` | — | Disables writing log files to disk. |
| `-WebhookUrl` | `string` | `-TeamsWebhookUrl` | Incoming webhook URL (Teams, Slack, Power Automate) to receive completion notifications. |
| `-DisableNotification` | `switch` | `-NoNotify` | Suppresses native Windows desktop balloon/toast notifications. |
| `-Force` | `switch` | — | Bypasses all interactive confirmation prompts (`Y/N`). |
| `-Help` | `switch` | `-h` | Displays custom interactive console help manual. |

---

## 📊 Console & Progress Bar Output

When executed, `Sync-OneDriveFolderAndNotify.ps1` produces structured, real-time visual output:

```text
================================================================================
  Sync-OneDriveFolderAndNotify | Execution Parameters
================================================================================
Target User (UPN)     : john.doe@company.com
OneDrive Path         : Backups/ProjectFiles
Local Path            : C:\Data\ProjectFiles
Sync Direction        : TwoWay
Recursive Sync        : True
Auto-Create Remote    : Enabled
Auto-Create Local     : Enabled
Auth Mode             : Entra ID App (Client Secret)
Log Output            : C:\Scripts\logs\Sync-OneDriveFolderAndNotify_20261015_120000.log
Notifications         : Desktop Toast Enabled
Webhook               : Configured
================================================================================
[2026-10-15 12:00:01.120] [INFO   ] [Sync] Verifying Microsoft Graph connectivity...
[2026-10-15 12:00:02.340] [SUCCESS] [Sync] Connected to Microsoft Graph via ClientSecretCredential.
[2026-10-15 12:00:03.110] [SUCCESS] [Sync] Located OneDrive: Id 'b!d182...', Type 'business'.
[2026-10-15 12:00:03.540] [SUCCESS] [Sync] Local target folder verified: 'C:\Data\ProjectFiles'
[2026-10-15 12:00:04.220] [SUCCESS] [Sync] OneDrive folder verified: 'Backups/ProjectFiles'
[2026-10-15 12:00:05.100] [INFO   ] [Sync] Sync plan: 4 to upload, 1 to download, 18 unchanged.
[2026-10-15 12:00:06.120] [PROGRESS] [Sync] [1/5] Uploading 'Reports/Q3_Summary.xlsx' (1.42 MB)...
[2026-10-15 12:00:07.450] [SUCCESS] [Sync] Uploaded 'Reports/Q3_Summary.xlsx' (1.42 MB) successfully.
[2026-10-15 12:00:08.010] [PROGRESS] [Sync] [2/5] Uploading 'Database/Backup.bak' (45.80 MB)...
[2026-10-15 12:00:08.150] [INFO   ] [Sync] Large file detected (45.80 MB). Initializing chunked upload session for 'Backup.bak'...
[2026-10-15 12:00:15.820] [SUCCESS] [Sync] Large file upload completed: 'Database/Backup.bak' (45.80 MB).
================================================================================
  SYNCHRONIZATION COMPLETED SUMMARY
================================================================================
Target User           : john.doe@company.com
OneDrive Folder       : Backups/ProjectFiles
Local Folder          : C:\Data\ProjectFiles
Sync Direction        : TwoWay
Total Files Analyzed  : 23
Files Uploaded        : 4
Files Downloaded      : 1
Files Unchanged       : 18
Failed Operations     : 0
Total Data Transferred: 48.25 MB
Elapsed Duration      : 00:00:18
Log File Path         : C:\Scripts\logs\Sync-OneDriveFolderAndNotify_20261015_120000.log
Final Status          : [SUCCESS]
================================================================================
```

---

## 👨‍💻 Author & Support

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **Version**: 1.0.0
- **License**: MIT
