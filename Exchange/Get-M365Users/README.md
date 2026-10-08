# Get-M365Users

A production-ready PowerShell audit and reporting script for Microsoft 365 administrators. Connects directly to **Exchange Online** and **Microsoft Graph** to discover, categorize, and report on all licensed users across your tenant in a clean, grouped tabular layout.

---

## Overview

`Get-M365Users.ps1` connects to Microsoft 365 Exchange Online (using modern REST cmdlets) and Microsoft Graph to:
1. **Intelligent License Classification**: Distinguishes between **Base Commercial Plans** (e.g. *Microsoft 365 Business Standard*, *Business Premium*, *Office 365 E3/E5*) and **Add-ons / Free Services** (*Defender for Office 365*, *Power Automate Free*, *Power BI Free*, *Azure Rights Management*).
2. **Unified Grouping**: Groups users by their primary commercial base plan, eliminating fragmentation where users with the same base plan were previously scattered across secondary add-on groups.
3. **Sequential Row Numbering (`Nr`)**: Provides continuous numbering at the beginning of each user list and summary table.
4. **Consolidated Email & Display Name**: Single clear `Email` column and full name (`Display Name`).
5. **Assigned Licenses in Every View**: Clearly details both Base Plan and Add-on licenses across all views.
6. **Executive Licensing Summary**:
   - **Commercial Subscriptions (Paid)**: Assigned (Used), Available (Free), Total purchased quotas, utilization percentage, and capacity status tags (`[OK]`, `[LOW]`, `[FULL - 0 FREE]`).
   - **Complimentary & Free Pools**: Isolated from commercial subscriptions so multi-seat viral/free pools (e.g. 1M Power BI Free) do not distort financial metrics.
   - **User Count by Base Plan**: Exact user count and percentage share per primary subscription.

---

## Features

- **Automated Connection Handling**: Automatically detects active Exchange Online and Microsoft Graph sessions or prompts for modern administrator sign-in.
- **Base Plan vs. Add-on Separation**: Clearly isolates primary commercial licenses from security add-ons and complimentary cloud services.
- **Row Numbering (`Nr`)**: Consecutive index column at the start of each user list.
- **Consolidated Email Column**: Cleaned up layout with a single `Email` column and full display name.
- **Executive Licensing Summary**: Dual-table inventory reporting paid subscriptions and free pools with real-time capacity tags.
- **Summary-Only Mode (`-LicenseSummary`)**: Instantly inspect tenant license quotas and utilization without dumping individual user accounts.
- **Grouped Tabular Reporting**: Users are automatically grouped by base license plan with summary statistics and zero header wrapping.
- **Partial License Filter (`-LicenseFilter`)**: Quickly filter users by partial plan names (e.g., `-LicenseFilter "Business"`, `-LicenseFilter "Defender"`).
- **Simple vs. Detailed Mode (`-Details`)**: Toggle between a streamlined view and an extended audit view with Department, Job Title, Location, Creation Date, and Archive status.
- **User Search (`-Search`)**: Quickly find users by DisplayName, email, department, or job title.
- **CSV Export (`-ExportCsv`)**: Save clean, UTF-8 encoded audit reports directly to disk.
- **Pipeline Integration (`-PassThru`)**: Emit custom PowerShell objects to the pipeline for downstream commands or `Out-GridView`.
- **Offline / Skip Graph Option (`-SkipGraph`)**: Audit solely via Exchange Online without connecting to Microsoft Graph.
- **Security & Input Sanitization**: Robust defense against command injection or illegal parameter characters.
- **Cross-Platform Compatibility**: Fully compatible with Windows PowerShell 5.1 and PowerShell 7+ across English and Polish environments.

---

## Prerequisites

- **PowerShell**: Version 5.1 or PowerShell 7+.
- **Required Module**: `ExchangeOnlineManagement` (v3.0.0 or higher recommended):
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```
- **Optional Module** (for tenant-wide free/used license pools and assigned SKU details): `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users`:
  ```powershell
  Install-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users -Scope CurrentUser
  ```
- **Administrative Role**: Exchange Administrator, User Administrator, or Global Administrator.

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
Retrieves all licensed accounts grouped by base plan with assigned licenses, row numbers, and executive licensing summary:

```powershell
.\Get-M365Users.ps1 -All
```

### 3. List All Licensed Users with Extended Details
Includes organizational columns (Department / Title, Location, Mailbox Type, Creation Date, Archive Status):

```powershell
.\Get-M365Users.ps1 -All -Details
```

### 4. Display Executive License Summary Only
Displays only the tenant quota tables and user breakdown without printing user accounts:

```powershell
.\Get-M365Users.ps1 -LicenseSummary
```

### 5. Filter by Partial License Name
Filter for Business-tier users:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Business"
```

Filter for Defender-licensed users with full details:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Defender" -Details
```

### 6. Search for a Specific User
Search by name, email, department, or job title:

```powershell
.\Get-M365Users.ps1 -Search "kowalski" -Details
```

### 7. Export Report to CSV File
Audit all licensed users and export the report to a CSV file:

```powershell
.\Get-M365Users.ps1 -All -ExportCsv "C:\Reports\M365_LicensedUsers.csv"
```

### 8. Interactive GridView / Pipeline Output
Output flat records directly into an interactive GUI grid:

```powershell
.\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView
```

### 9. Run Strictly via Exchange Online (Skip Microsoft Graph)
```powershell
.\Get-M365Users.ps1 -All -SkipGraph
```

---

## Output Modes & Columns

### Simple Mode (`-All`)

| Column Name | Width | Description |
| :--- | :---: | :--- |
| **Nr** | 4 | Sequential row index (1..N) |
| **Display Name** | 25 | Full user display name |
| **Email** | 34 | Consolidated primary email address |
| **Base License** | 32 | Core commercial subscription plan (e.g. *Microsoft 365 Business Standard*) |
| **Add-on Licenses** | 26 | Assigned security/feature add-ons (e.g. *Defender for Office 365 (Plan 1)*) |
| **Department** | 20 | User's organizational department |

### Detailed Mode (`-All -Details`)

| Column Name | Width | Description |
| :--- | :---: | :--- |
| **Nr** | 4 | Sequential row index (1..N) |
| **Display Name** | 22 | Full user display name |
| **Email** | 32 | Consolidated primary email address |
| **Base License** | 32 | Core commercial subscription plan |
| **Add-on Licenses** | 22 | Secondary add-on and complimentary licenses |
| **Dept / Job Title** | 20 | Consolidated department and job title |
| **Location** | 15 | City, office, or country |
| **Type** | 12 | Recipient type (`UserMailbox`, `EntraUser`) |
| **Created** | 11 | Creation date (`yyyy-MM-dd`) |
| **Archive** | 7 | In-Place Archive status (`None`, `Active`) |

---

## Executive Licensing Summary

Rendered at the end of every user report or as the standalone output of `-LicenseSummary`:

```text
================================================================================
EXECUTIVE LICENSING SUMMARY (PODSUMOWANIE LICENCJI)
================================================================================
  Total Licensed Users Found : 55
  Primary Base License Plans : 2
  Total Subscription Pools   : 8

[1] COMMERCIAL SUBSCRIPTIONS (PLATNE LICENCJE I DODATKI):
  Nr License / Plan Name                    Used     Free    Total   Util % Status
  -- -------------------                    ----     ----    -----   ------ ------
   1 Microsoft 365 Business Standard          53        4       57    93.0% [OK]
   2 Microsoft 365 Business Premium            1        0        1   100.0% [FULL - 0 FREE]
   3 Defender for Office 365 (Plan 1)         23        2       25    92.0% [OK]
  -----------------------------------------------------------------------------
     COMMERCIAL TOTALS: Used = 77 | Free = 6 | Total = 83 (92.8% Utilized)

[2] COMPLIMENTARY & FREE CLOUD SERVICES (BEZPLATNE USLUGI W CHMURZE):
  Nr Service / Pool Name                    Used     Free    Total Status
  -- -------------------                    ----     ----    ----- ------
   1 Power Automate (Free)                    15    9,985   10,000 [Available]
   2 Power BI (Free)                           1  999,999 1,000,000 [Available]
   3 Azure Rights Management (Free)            1   49,999   50,000 [Available]

[3] USER COUNT BY PRIMARY BASE LICENSE (LICZBA UZYTKOWNIKOW WG LICENCJI BAZOWEJ):
  Nr Primary Base License Plan        User Count   Share %
  -- -------------------------        ----------   -------
   1 Microsoft 365 Business Standard          53     96.4%
   2 Microsoft 365 Business Premium            1      1.8%
   3 EntraUser (No Mailbox)                    1      1.8%
  --------------------------------------------------------
     TOTAL LICENSED USERS:                    55    100.0%
================================================================================
```

---

## Parameters Reference

| Parameter | Alias | Type | Description |
| :--- | :--- | :--- | :--- |
| `-All` | - | `Switch` | Lists all licensed users in the tenant. |
| `-LicenseFilter` | `-License`, `-Plan` | `String` | Filters users matching a partial license or plan name. |
| `-Details` | `-d`, `-Detailed` | `Switch` | Displays extended user attributes in additional columns. |
| `-Search` | `-FilterUser`, `-User` | `String` | Searches users by DisplayName, email, department, or title. |
| `-LicenseSummary` | `-SummaryOnly`, `-Quota` | `Switch` | Displays only the executive licensing summary tables. |
| `-ExportCsv` | `-CsvPath`, `-Export` | `String` | Exports the retrieved records to a UTF-8 CSV report. |
| `-NoGrouping` | - | `Switch` | Displays a single flat table instead of grouped plan sections. |
| `-PassThru` | - | `Switch` | Emits custom `PSCustomObject` items to the PowerShell pipeline. |
| `-SkipGraph` | `-NoGraph` | `Switch` | Bypasses Microsoft Graph connection and audits solely via Exchange Online. |
| `-AdminUserPrincipalName` | `-AdminUPN` | `String` | Administrator UPN for pre-populating Exchange Online login. |
| `-Help` | `-h` | `Switch` | Displays the help screen, examples, version, and author info. |

---

## Author & Contact

- **Author**: Roman Pindela
- **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **Version**: 1.3.0
