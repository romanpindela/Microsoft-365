# Get-M365Users

A production-ready PowerShell audit and reporting script for Microsoft 365 administrators. Connects directly to **Exchange Online** to discover, categorize, and report on all licensed users across your tenant in a clean, grouped tabular layout.

---

## Overview

`Get-M365Users.ps1` connects to Microsoft 365 Exchange Online (using modern REST cmdlets), identifies all users with active licenses or mailbox plans, maps their licensing plans to human-readable names (e.g., *Exchange Online Plan 2 (Enterprise / E3 / E5)*, *Exchange Online Plan 1 (Essentials / Business)*, *Exchange Online Kiosk*), and displays them organized into grouped tables by license type.

The script offers two display modes:
1. **Simple Mode**: Fast, clean overview of user identity, email, and mailbox type.
2. **Detailed Mode (`-Details`)**: Extended report including organizational metadata (Department, Job Title, Office, City, Country, Creation Date, In-Place Archive status, and Address List visibility).

---

## Features

- **Automated Exchange Online Connection**: Automatically detects an active Exchange Online session or prompts for modern interactive administrator sign-in (following the robust pattern from `Manage-MailboxUserAccess.ps1`).
- **Interactive Help on Blank Execution**: Executing the script without parameters immediately displays a formatted guide with syntax, usage examples, version, and author contact.
- **Grouped Tabular Reporting**: Users are automatically grouped by assigned license/mailbox plan with executive counts and a summary breakdown.
- **Partial License Filter (`-LicenseFilter`)**: Quickly filter users by partial plan or tier names (e.g., `-LicenseFilter "Enterprise"`, `-LicenseFilter "Deskless"`, `-LicenseFilter "Essentials"`).
- **Simple vs. Detailed Mode (`-Details`)**: Toggle between a lightweight identity table and a deep audit view with department, location, job title, and archiving details.
- **User Search (`-Search`)**: Quickly find users by DisplayName, UPN, email, or department.
- **CSV Export (`-ExportCsv`)**: Save clean, UTF-8 encoded audit reports directly to disk.
- **Pipeline Integration (`-PassThru`)**: Emit custom PowerShell objects to the pipeline for downstream sorting, filtering, or `Out-GridView`.
- **Optional Microsoft Graph Support (`-UseGraph`)**: Query Microsoft Graph for tenant-wide Entra ID license SKUs and non-mailbox accounts.
- **Security & Input Sanitization**: Robust defense against command injection or illegal parameter characters.
- **Cross-Platform & Bilingual Compatibility**: Compatible with PowerShell 5.1 and 7+ across English and Polish Windows operating systems.

---

## Prerequisites

- **PowerShell**: Version 5.1 or PowerShell 7+.
- **PowerShell Module**: `ExchangeOnlineManagement` (v3.0.0 or higher recommended):
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```
- **Administrative Role**: Exchange Administrator or Global Administrator in your Microsoft 365 tenant.
- *(Optional)* `Microsoft.Graph.Users` and `Microsoft.Graph.Authentication` if using `-UseGraph`.

---

## Security & Unblocking

After downloading the script from GitHub, unblock the file before executing:

```powershell
Unblock-File -Path .\Get-M365Users.ps1
```

---

## Usage Examples

### 1. Display Built-In Help & Usage
Running without arguments displays the built-in help banner, parameters, and examples:

```powershell
.\Get-M365Users.ps1
# Or explicitly:
.\Get-M365Users.ps1 -h
.\Get-M365Users.ps1 -Help
```

### 2. List All Licensed Users (Simple Grouped View)
Retrieves all licensed accounts and displays them in tables grouped by license plan:

```powershell
.\Get-M365Users.ps1 -All
```

### 3. List All Licensed Users with Extended Details
Includes organizational columns (Department, Job Title, Office/City, Country, Archive status):

```powershell
.\Get-M365Users.ps1 -All -Details
```

### 4. Filter by Partial License Name
Filter for Enterprise (E3/E5) users:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Enterprise"
```

Filter for Frontline/Kiosk (Deskless/F1/F3) users with full details:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Deskless" -Details
```

### 5. Search for a Specific User
Search by name, email, or UPN and inspect their licensing details:

```powershell
.\Get-M365Users.ps1 -Search "smith" -Details
```

### 6. Export Report to CSV File
Audit all licensed users and export the report to a CSV file:

```powershell
.\Get-M365Users.ps1 -All -ExportCsv "C:\Reports\M365_LicensedUsers.csv"
```

### 7. Interactive GridView / Pipeline Output
Output flat records directly into an interactive GUI grid:

```powershell
.\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView
```

### 8. Enrich with Microsoft Graph SKUs
Query both Exchange Online and Microsoft Graph for tenant-wide Entra ID license SKUs:

```powershell
.\Get-M365Users.ps1 -All -UseGraph -Details
```

---

## Output Modes & Columns

| Column Name | Simple Mode | Detailed Mode (`-Details`) | Description |
| :--- | :---: | :---: | :--- |
| **Display Name** | Yes | Yes | User's full display name |
| **User Principal Name** | Yes | Yes | User's sign-in identifier (UPN) |
| **Primary SMTP** | Yes | Yes | Primary email address |
| **Recipient Type** | Yes | Yes | Mailbox classification (UserMailbox, SharedMailbox, etc.) |
| **License Plan** | Group Header / Column | Group Header / Column | Human-friendly license/mailbox plan name |
| **Department** | - | Yes | User's assigned department |
| **Job Title** | - | Yes | Job title/position |
| **Office / City** | - | Yes | Physical office location or city |
| **Country** | - | Yes | Two-letter country/region code |
| **Created Date** | - | Yes | Mailbox or account creation timestamp |
| **Archive** | - | Yes | In-Place Archive status (`None`, `Active`, `Local`) |
| **Hidden from GAL** | - | Yes | Address book visibility status |

---

## Parameters Reference

| Parameter | Alias | Type | Description |
| :--- | :--- | :--- | :--- |
| `-All` | - | `Switch` | Lists all licensed users in the tenant. |
| `-LicenseFilter` | `-License`, `-Plan` | `String` | Filters users matching a partial license or mailbox plan name. |
| `-Details` | `-d`, `-Detailed` | `Switch` | Displays extended user attributes in additional columns. |
| `-Search` | `-FilterUser`, `-User` | `String` | Searches users by DisplayName, UPN, primary email, or department. |
| `-ExportCsv` | `-CsvPath`, `-Export` | `String` | Exports the retrieved records to a UTF-8 CSV report. |
| `-NoGrouping` | - | `Switch` | Displays a single flat table instead of grouped plan sections. |
| `-PassThru` | - | `Switch` | Emits custom `PSCustomObject` items to the PowerShell pipeline. |
| `-AdminUserPrincipalName` | `-AdminUPN` | `String` | Administrator UPN for pre-populating Exchange Online login. |
| `-UseGraph` | - | `Switch` | Integrates Microsoft Graph for tenant-wide Entra ID license SKUs. |
| `-Help` | `-h` | `Switch` | Displays the help screen, examples, version, and author info. |

---

## Author & Contact

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **Version**: 1.0.0
