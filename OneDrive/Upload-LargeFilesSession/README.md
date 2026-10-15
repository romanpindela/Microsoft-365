# Upload-LargeFilesSession

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0-008272.svg)](https://learn.microsoft.com/graph/)
[![OneDrive for Business](https://img.shields.io/badge/OneDrive%20for%20Business-Supported-0078D4.svg)](https://www.microsoft.com/microsoft-365/onedrive/online-cloud-storage)
[![Entra ID](https://img.shields.io/badge/Entra%20ID-App--Only%20Auth-0078D4.svg)](https://entra.microsoft.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

An enterprise-grade PowerShell automation script engineered to upload individual files, entire directories, or wildcard-matched file sets (e.g., `*.pdf`, `p?indela.pdf`, `p*.docx`, or folder names) directly to a specified Microsoft 365 user's **OneDrive for Business** library using official **Microsoft Graph Upload Sessions** (`createUploadSession`).

Designed, developed, and maintained by **Roman Pindela**, this tool supports multi-gigabyte file transfers (up to 250 GB), automatic remote folder tree provisioning, real-time dual progress indicators, comprehensive logging, and unattended non-interactive execution via **Microsoft Entra ID** (Client Credentials).

---

## 🚀 Key Features

* **Official Microsoft Graph Upload Session**: Uses Microsoft Graph's `createUploadSession` API to slice files into compliant 320 KiB multiples (default 10 MB chunks, customizable 1–100 MB). Overcomes standard 4 MB single-request API limits and supports files up to 250 GB.
* **Automatic Remote Folder Creation**: Verifies if the target OneDrive folder path exists (e.g. `/Backups/2026/October`). If any parent or child subfolder is missing, the script automatically provisions the entire hierarchy before starting file transfers.
* **Versatile File Matching & Wildcards**:
  * Single files: e.g. `C:\Archive\backup.vhdx`
  * Directory trees: e.g. `C:\Reports` or `.\ten_folder`
  * Global and single-character masks: e.g. `*.pdf`, `p?indela.pdf`, `p*.docx`, `report_202*.xlsx`
  * Multiple paths / patterns combined in a single execution.
* **Structure Preservation & Recursion**: Supports recursive folder discovery (`-Recurse`) and local folder tree replication inside the destination OneDrive folder (`-PreserveStructure`).
* **Real-Time Dual Progress & Telemetry**:
  * **Overall Progress**: Total files processed, total transferred payload vs total size, average transfer throughput in MB/s, and estimated time remaining (ETA).
  * **File Chunk Progress**: Active chunk number, byte range (`bytes 10,485,760 - 20,971,519 / 104,857,600`), instantaneous upload speed (MB/s), and file completion percentage.
* **Dual Authentication Modes**:
  1. **Headless / Non-Interactive (App-Only)**: Connects silently via Entra ID App Registration using Client ID, Tenant ID, and Client Secret or Certificate Thumbprint. Ideal for scheduled tasks, backup scripts, Azure DevOps, and CI/CD pipelines.
  2. **Interactive Administrator**: Connects using standard Modern Authentication / MFA browser prompt via Microsoft.Graph PowerShell SDK.
* **Resilient Retry Engine**: Built-in exponential backoff handles network blips, HTTP 429 throttling (`Retry-After`), and temporary cloud service interruptions (HTTP 5xx).
* **Enterprise Logging**: Dual output to a color-coded console and timestamped UTF-8 log files (`[yyyy-MM-dd HH:mm:ss] [LEVEL] Message`).
* **Defensive Clean Code**: Strict parameter validation, input sanitization, path traversal prevention, and language-independent compatibility with both Polish and English Windows environments.

---

## 🛠️ Prerequisites

* **Operating System**: Windows 10, Windows 11, Windows Server 2016/2019/2022/2025.
* **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+ (Core).
* **Network**: Outbound HTTPS connectivity on port 443 to:
  * `https://login.microsoftonline.com`
  * `https://graph.microsoft.com`
  * `https://*.up.1drv.com` / `https://*.sharepoint.com`
* **Microsoft 365 License**: The target user must have an active license with OneDrive for Business enabled.

---

## 🔐 Entra ID Non-Interactive Authentication Setup

To run this script unattended in scheduled tasks, automated backups, or service scripts without interactive login prompts, register an application in **Microsoft Entra ID** and assign Application Permissions:

### Step 1: Register Application in Microsoft Entra ID
1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) or [Azure Portal](https://portal.azure.com).
2. Navigate to **Identity** > **Applications** > **App registrations**.
3. Click **New registration**:
   * **Name**: `OneDrive-LargeFiles-Uploader` (or your preferred name)
   * **Supported account types**: `Accounts in this organizational directory only (Single tenant)`
   * **Redirect URI**: Leave blank.
4. Click **Register**.

### Step 2: Record Identifiers
* Copy the **Application (client) ID** (e.g., `11111111-2222-3333-4444-555555555555`).
* Copy the **Directory (tenant) ID** (e.g., `aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee`).

### Step 3: Generate a Client Secret (or Certificate)
1. In your registered application, go to **Manage** > **Certificates & secrets**.
2. Under the **Client secrets** tab, click **New client secret**.
3. Add a description (e.g., `OneDriveAutomationKey`) and select an expiration period.
4. Click **Add**.
5. **CRITICAL**: Copy the string in the **Value** column immediately. It will be hidden permanently once you navigate away!

### Step 4: Configure API Permissions
1. Go to **Manage** > **API permissions** > Click **Add a permission**.
2. Select **Microsoft Graph**.
3. Click **Application permissions** (*Do NOT select Delegated permissions*).
4. Search and check the following permissions:
   * `Files.ReadWrite.All` — Grants full read and write access to all OneDrive and SharePoint files across the tenant.
   * `User.Read.All` — Allows resolving user identities and OneDrive drive endpoints.
5. Click **Add permissions**.

### Step 5: Grant Admin Consent
1. On the **API permissions** page, click **Grant admin consent for [Your Organization]**.
2. Confirm the prompt by clicking **Yes**.
3. Verify that both permissions display a green checkmark in the **Status** column: `Granted for [Your Organization]`.

---

## 📦 Download & Security Unblocking

When downloading the script from GitHub, Windows attaches an alternate data stream marking it as untrusted. Unblock the file before running:

```powershell
Unblock-File -Path .\Upload-LargeFilesSession.ps1
```

---

## 📚 Technical Documentation & PDF Guides

This tool is accompanied by three comprehensive, in-depth PDF handbooks providing architectural reference, enterprise security hardening, and passwordless certificate configuration:

| Document | Focus Area | Description | Format |
| :--- | :--- | :--- | :--- |
| [**Kompendium Microsoft Graph SDK i PowerShell: Automatyzacja OneDrive i Exchange**](Kompendium%20Microsoft%20Graph%20SDK%20i%20PowerShell_%20Automatyzacja%20OneDrive%20i%20Exchange.pdf) | Graph API & SDK Architecture | Comprehensive 324 KB engineering compendium detailing Microsoft Graph SDK internals, REST endpoints, rate limiting (HTTP 429), upload sessions, and cross-service automation across OneDrive and Exchange Online. | 📕 PDF |
| [**Przewodnik Bezpieczeństwa - Upload-LargeFilesSession.ps1**](Przewodnik%20Bezpiecze%C5%84stwa%20-%20Upload-LargeFilesSession.ps1.pdf) | Security & Governance | Hardening guide (555 KB) covering least-privilege scoping, Entra ID application governance, secure credential storage, audit log monitoring, and protection against accidental data leakage. | 📕 PDF |
| [**Przewodnik Certyfikatów X.509 - Upload-LargeFilesSession.ps1**](Przewodnik%20Certyfikat%C3%B3w%20X.509%20-%20Upload-LargeFilesSession.ps1.pdf) | X.509 Certificate Authentication | Practical walkthrough (295 KB) for passwordless authentication: generating RSA/X.509 certificates (self-signed or PKI), uploading public keys to Entra ID, and configuring `-CertificateThumbprint` for unattended headless execution. | 📕 PDF |

---

## 📋 Parameters & Syntax

```powershell
.\Upload-LargeFilesSession.ps1
    [-UserPrincipalName] <string>
    [-TargetFolder] <string>
    [-Path] <string[]>
    [-Filter <string>]
    [-Recurse]
    [-PreserveStructure]
    [-ConflictBehavior <string>]
    [-ChunkSizeMB <int>]
    [-TenantId <string>]
    [-ClientId <string>]
    [-ClientSecret <string>]
    [-ClientSecretSecure <SecureString>]
    [-CertificateThumbprint <string>]
    [-CertificatePath <string>]
    [-Interactive]
    [-AdminUserPrincipalName <string>]
    [-LogPath <string>]
    [-LogDirectory <string>]
    [-LogLevel <string>]
    [-NoLog]
    [-Force]
    [-PassThru]
    [-Help]
```

### Parameter Reference

| Parameter | Aliases | Type | Default | Description |
| :--- | :--- | :--- | :--- | :--- |
| `-UserPrincipalName` | `-User`, `-UPN`, `-UserIdentity`, `-u` | String | *Mandatory* | Primary SMTP/UPN of the target Microsoft 365 user. |
| `-TargetFolder` | `-Folder`, `-DestinationFolder`, `-Target` | String | *Mandatory* | Destination folder inside OneDrive (auto-created if missing). |
| `-Path` | `-FilePath`, `-Source`, `-Files`, `-p` | String[] | *Mandatory* | File paths, folder paths, or wildcard masks (`*.pdf`, `p*.docx`, etc.). |
| `-Filter` | `-Mask` | String | `$null` | Filter mask when scanning directories (e.g. `*.pdf`). |
| `-Recurse` | `-r` | Switch | `$false` | Recursively scans source directories. |
| `-PreserveStructure` | `-KeepHierarchy`, `-MaintainFolders` | Switch | `$false` | Replicates local subfolder structure in the remote OneDrive folder. |
| `-ConflictBehavior` | — | String | `'replace'` | Action on filename conflict: `replace`, `rename`, `fail`. |
| `-ChunkSizeMB` | `-ChunkMB` | Int | `10` | Chunk size in MB (1–100 MB, rounded to multiple of 320 KiB). |
| `-TenantId` | `-Tenant` | String | `$null` | Entra ID Directory (tenant) ID for App-Only login. |
| `-ClientId` | `-AppId`, `-ApplicationId` | String | `$null` | Entra ID Application (client) ID for App-Only login. |
| `-ClientSecret` | `-Secret` | String | `$null` | Plaintext client secret for App-Only login. |
| `-ClientSecretSecure` | — | SecureString | `$null` | SecureString representation of the client secret. |
| `-CertificateThumbprint`| `-Thumbprint` | String | `$null` | Certificate thumbprint from Windows Cert Store for App-Only login. |
| `-CertificatePath` | `-CertPath` | String | `$null` | Path to `.pfx` certificate file for App-Only login. |
| `-Interactive` | `-InteractiveLogin`, `-i` | Switch | `$false` | Forces interactive MFA browser sign-in via Graph SDK. |
| `-AdminUserPrincipalName`| `-AdminUPN` | String | `$null` | Admin UPN to pre-populate interactive login prompt. |
| `-LogPath` | `-LogFile` | String | Auto | Explicit custom path for the log file. |
| `-LogDirectory` | `-LogDir` | String | `.\logs` | Custom directory where log files should be created. |
| `-LogLevel` | — | String | `'Info'` | Logging threshold: `Debug`, `Info`, `Warning`, `Error`. |
| `-NoLog` | — | Switch | `$false` | Disables disk logging (console output remains active). |
| `-Force` | `-f` | Switch | `$false` | Bypasses confirmation prompts for unattended execution. |
| `-PassThru` | — | Switch | `$false` | Emits custom objects of uploaded files to the PowerShell pipeline. |
| `-Help` | `-h` | Switch | `$false` | Displays built-in help and execution guidelines. |

---

## 💡 Practical Usage Examples

### 1. Zero-Parameter Execution (Help Banner)
Running the script without parameters or with `-h` immediately displays the help manual, parameters, and Entra ID configuration guide:
```powershell
.\Upload-LargeFilesSession.ps1
```

### 2. Upload Single Large File Interactively
Uploads a large database backup or disk image to user `john.doe@company.com` using interactive MFA sign-in:
```powershell
.\Upload-LargeFilesSession.ps1 `
    -UserPrincipalName "john.doe@company.com" `
    -TargetFolder "Backups/Database" `
    -Path "D:\Backups\Production_20261015.bak"
```

### 3. Upload Files by Wildcard Mask (`*.pdf`)
Uploads all PDF files matching the mask `*.pdf` from `C:\MonthlyReports` into the `Reports/2026` OneDrive folder:
```powershell
.\Upload-LargeFilesSession.ps1 `
    -User "john.doe@company.com" `
    -Folder "Reports/2026" `
    -Path "C:\MonthlyReports\*.pdf"
```

### 4. Upload Files by Single-Character Wildcard (`p?indela.pdf`)
Uploads all variations matching single-character pattern `p?indela.pdf` (e.g. `ppindela.pdf`, `p1indela.pdf`):
```powershell
.\Upload-LargeFilesSession.ps1 `
    -User "john.doe@company.com" `
    -Folder "Contracts" `
    -Path "p?indela.pdf"
```

### 5. Upload Files by Prefix Wildcard (`p*.docx`) with Rename Conflict Policy
Uploads all Word documents starting with `p`. If a file with the same name exists in OneDrive, Microsoft Graph automatically renames the incoming file (e.g. `p_report 1.docx`):
```powershell
.\Upload-LargeFilesSession.ps1 `
    -User "john.doe@company.com" `
    -Folder "Documents" `
    -Path "p*.docx" `
    -ConflictBehavior rename
```

### 6. Upload an Entire Directory Tree Replicating Structure
Uploads `C:\Projects\Alpha` recursively, recreating all subdirectories inside `Archive/Alpha` in OneDrive:
```powershell
.\Upload-LargeFilesSession.ps1 `
    -User "john.doe@company.com" `
    -Folder "Archive/Alpha" `
    -Path "C:\Projects\Alpha" `
    -Recurse `
    -PreserveStructure
```

### 7. Non-Interactive Unattended Upload via Entra ID Client Secret
Runs headless in Task Scheduler, Azure DevOps, or PowerShell jobs using App-Only client credentials with 25 MB chunks:
```powershell
.\Upload-LargeFilesSession.ps1 `
    -UserPrincipalName "backup-service@company.com" `
    -TargetFolder "SQL_Dumps" `
    -Path "E:\SQLBackups\*.bak" `
    -TenantId "11111111-2222-3333-4444-555555555555" `
    -ClientId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
    -ClientSecret "SecretValueHere~12345" `
    -ChunkSizeMB 25 `
    -LogPath "C:\Logs\OneDrive_Upload.log"
```

### 8. Non-Interactive Upload via Environment Variables
Set credentials in your environment once (e.g. in CI/CD pipeline secrets):
```powershell
$env:M365_TENANT_ID     = "11111111-2222-3333-4444-555555555555"
$env:M365_CLIENT_ID     = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
$env:M365_CLIENT_SECRET = "SecretValueHere~12345"

.\Upload-LargeFilesSession.ps1 `
    -User "user@company.com" `
    -Folder "DailyBackups" `
    -Path "D:\Data\*.zip"
```

### 9. Pipeline Output with `-PassThru`
Captures upload metadata objects (ID, WebUrl, Size, Upload Duration, Speed) into a PowerShell variable or exports to CSV:
```powershell
$uploaded = .\Upload-LargeFilesSession.ps1 `
    -User "john.doe@company.com" `
    -Folder "Exports" `
    -Path "C:\Exports\*.csv" `
    -PassThru

$uploaded | Export-Csv -Path ".\UploadReport.csv" -NoTypeInformation -Encoding utf8
```

---

## 📊 Console Output & Telemetry Example

```text
================================================================================
                   UPLOAD EXECUTION PARAMETERS & PROFILE
================================================================================
  Target User (UPN)     : roman.pindela@company.com
  OneDrive Drive ID     : b!AbCdEf1234567890 (business)
  Destination Folder    : /Backups/2026/October
  Total Files Matched   : 3
  Total Payload Size    : 12.45 GB (13,368,393,728 bytes)
  Configured Chunk Size : 10.00 MB (10,485,760 bytes, compliant 320 KiB multiple)
  Collision Behavior    : replace
  Preserve Structure    : False
  Authentication Mode   : Entra ID App-Only (Client Credentials)
  Log File Path         : C:\GitHub\Microsoft 365\OneDrive\Upload-LargeFilesSession\logs\Upload-LargeFilesSession_20261015_123500.log
================================================================================
  [*] Beginning upload session sequence for 3 file(s)...
  [*] [1/3] Processing: database_archive.7z (4.20 GB)...
  [+] [1/3] Completed: database_archive.7z (4.20 GB in 138.2s, avg 31.12 MB/s) -> Item ID: 01ABCDEF...
  [*] [2/3] Processing: documentation.pdf (15.30 MB)...
  [+] [2/3] Completed: documentation.pdf (15.30 MB in 0.8s, avg 19.12 MB/s) -> Item ID: 01ABCDEG...
  [*] [3/3] Processing: vm_snapshot.vhdx (8.23 GB)...
  [+] [3/3] Completed: vm_snapshot.vhdx (8.23 GB in 265.4s, avg 31.75 MB/s) -> Item ID: 01ABCDEH...

================================================================================
                    UPLOAD EXECUTION SUMMARY REPORT
================================================================================
  Target User (OneDrive)  : roman.pindela@company.com
  Destination Folder      : /Backups/2026/October
  Total Files Attempted   : 3
  Files Succeeded         : 3
  Files Failed            : 0
  Total Payload Uploaded  : 12.45 GB (13,368,393,728 bytes)
  Total Execution Time    : 00:06:44
  Average Transfer Speed  : 31.54 MB/s
  Authentication Mode     : Entra ID App-Only (Client Credentials)
  Log File Location       : C:\GitHub\Microsoft 365\OneDrive\Upload-LargeFilesSession\logs\Upload-LargeFilesSession_20261015_123500.log
================================================================================
  [+] All files uploaded successfully to OneDrive.
```

---

## 📝 Logging Format

Logs are stored in UTF-8 format with standard RFC 3339 timestamps and severity levels:

```text
[2026-10-15 12:35:01] [INFO] Initializing Upload-LargeFilesSession v1.0.0...
[2026-10-15 12:35:02] [INFO] Scanning and resolving source files matching parameters...
[2026-10-15 12:35:03] [INFO] Authenticating via Entra ID App Registration (Client Credentials)...
[2026-10-15 12:35:04] [SUCCESS] Successfully acquired Microsoft Graph Bearer token via Entra ID App (aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee).
[2026-10-15 12:35:05] [INFO] Querying Microsoft 365 OneDrive drive endpoint for: roman.pindela@company.com...
[2026-10-15 12:35:06] [SUCCESS] Resolved OneDrive Drive ID: b!AbCdEf1234567890 (Type: business)
[2026-10-15 12:35:07] [INFO] Verifying and provisioning destination folder hierarchy in OneDrive: 'Backups/2026/October'...
[2026-10-15 12:35:08] [SUCCESS] Destination folder resolved in OneDrive (Folder ID: 01FOLDERID...).
[2026-10-15 12:35:09] [INFO] [1/3] Processing: database_archive.7z (4.20 GB)...
[2026-10-15 12:37:27] [SUCCESS] [1/3] Completed: database_archive.7z (4.20 GB in 138.2s, avg 31.12 MB/s) -> Item ID: 01ABCDEF...
[2026-10-15 12:41:45] [SUCCESS] All files uploaded successfully to OneDrive.
```

---

## 👤 Author & Support

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)

---

## 📄 License & Standards

This project complies with the **PowerShell Script Standards** specification:
* Fully parameterized and clean-code structured without hardcoded secrets.
* Bilingual Windows support (English and Polish Windows Desktop/Server).
* Built-in security unblocking and input sanitization against command injection.
* Licensed under the [MIT License](https://opensource.org/licenses/MIT).
