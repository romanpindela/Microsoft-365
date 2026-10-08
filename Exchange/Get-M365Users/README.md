# Get-M365Users

A production-ready PowerShell audit and reporting script for Microsoft 365 administrators. Connects directly to **Exchange Online** and **Microsoft Graph** to discover, categorize, and report on all licensed users across your tenant in a clean, grouped tabular layout.

---

## Overview

`Get-M365Users.ps1` connects to Microsoft 365 Exchange Online (using modern REST cmdlets) and Microsoft Graph to:
1. List all licensed users with sequential row numbering (`Nr`).
2. Present a clean, deduplicated table with a single `Email` column and full name (`Imię i nazwisko` / `Display Name`).
3. Display the user's **Assigned Licenses** (`Przypisane licencje`) in **every** tabular view (both Simple and Detailed modes).
4. Provide an **Executive Summary** at the bottom displaying:
   - Total number of licensed users (`Łączna liczba licencjonowanych użytkowników`).
   - Number of consumed/used licenses per type (`Wykorzystane wg typu`).
   - Number of free/available licenses per type (`Wolne wg typu`).
   - Total purchased subscription units (`Łącznie zakupione`).

---

## Features

- **Automated Connection Handling**: Automatically detects active Exchange Online and Microsoft Graph sessions or prompts for modern administrator sign-in.
- **Row Numbering (`Nr`)**: Consecutive index column at the start of each user list.
- **Consolidated Email Column**: Cleaned up layout with a single `Email` column and full display name.
- **Assigned Licenses Column in All Views**: Every table view clearly shows which licenses/plans are assigned to each user.
- **Tenant License Inventory Summary**: Displays total users, consumed licenses, and available (free) pool units per subscription SKU.
- **Grouped Tabular Reporting**: Users are automatically grouped by assigned license plan with summary statistics.
- **Partial License Filter (`-LicenseFilter`)**: Quickly filter users by partial plan names (e.g., `-LicenseFilter "Business"`, `-LicenseFilter "Enterprise"`, `-LicenseFilter "Kiosk"`).
- **Simple vs. Detailed Mode (`-Details`)**: Toggle between a streamlined identity view and an extended audit view with Department, Job Title, Office, City, Country, Creation Date, and Archive status.
- **User Search (`-Search`)**: Quickly find users by DisplayName, email, or department.
- **CSV Export (`-ExportCsv`)**: Save clean, UTF-8 encoded audit reports directly to disk.
- **Pipeline Integration (`-PassThru`)**: Emit custom PowerShell objects to the pipeline for downstream commands or `Out-GridView`.
- **Offline / Skip Graph Option (`-SkipGraph`)**: Audit solely via Exchange Online without connecting to Microsoft Graph.
- **Security & Input Sanitization**: Robust defense against command injection or illegal parameter characters.
- **Cross-Platform & Bilingual Compatibility**: Fully compatible with PowerShell 5.1 and 7+ across English and Polish Windows operating systems.

---

## Prerequisites

- **PowerShell**: Version 5.1 or PowerShell 7+.
- **Required Module**: `ExchangeOnlineManagement` (v3.0.0 or higher recommended):
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```
- **Optional Module** (for tenant-wide free/used license pools): `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users`:
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
Retrieves all licensed accounts with row numbers, single email column, and license summary:

```powershell
.\Get-M365Users.ps1 -All
```

### 3. List All Licensed Users with Extended Details
Includes organizational columns (Department, Job Title, Office/City, Country, Creation Date, Archive):

```powershell
.\Get-M365Users.ps1 -All -Details
```

### 4. Filter by Partial License Name
Filter for Business-tier users:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Business"
```

Filter for Enterprise (E3/E5) users with full details:

```powershell
.\Get-M365Users.ps1 -LicenseFilter "Enterprise" -Details
```

### 5. Search for a Specific User
Search by name, email, or department:

```powershell
.\Get-M365Users.ps1 -Search "kowalski" -Details
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

### 8. Run Strictly via Exchange Online (Skip Microsoft Graph)
```powershell
.\Get-M365Users.ps1 -All -SkipGraph
```

---

## Output Modes & Columns

| Column Name | Simple Mode | Detailed Mode (`-Details`) | Description |
| :--- | :---: | :---: | :--- |
| **Nr** | Yes | Yes | Sequential row number (1..N) |
| **Imię i nazwisko** | Yes | Yes | User's full display name |
| **Email** | Yes | Yes | Consolidated primary email address / UPN |
| **Przypisane licencje** | Yes | Yes | Assigned M365 license(s) or Exchange mailbox plan |
| **Dział** | - | Yes | User's assigned department |
| **Stanowisko** | - | Yes | Job title / position |
| **Biuro / Miasto** | - | Yes | Physical office location or city |
| **Kraj** | - | Yes | Country / region code |
| **Typ konta** | Yes | Yes | Mailbox classification (`UserMailbox`, `SharedMailbox`, etc.) |
| **Utworzono** | - | Yes | Account or mailbox creation timestamp |
| **Archiwum** | - | Yes | In-Place Archive status (`None`, `Active`, `Local`) |

---

## Summary Section at the Bottom

At the end of every execution, the script renders a structured summary:

```text
================================================================================
PODSUMOWANIE LICENCJI I UŻYTKOWNIKÓW (EXECUTIVE SUMMARY)
================================================================================
  Łączna liczba licencjonowanych użytkowników : 54
  Liczba typów licencji / planów w zestawieniu: 3

  ZESTAWIENIE LICENCJI TENANTA (SUBKRYPCJE M365):
Typ licencji / SKU                  Wykorzystane (Used)   Wolne (Free)   Łącznie (Total)
------------------                  -------------------   ------------   ---------------
Microsoft 365 Business Premium                       25              5                30
Exchange Online Plan 2                               15              2                17
Microsoft 365 Apps for business                      14              1                15

  SUMA POZYCJI SUBSKRYPCJI: Wykorzystane = 54 | Wolne = 8 | Łącznie = 62
================================================================================
```

---

## Parameters Reference

| Parameter | Alias | Type | Description |
| :--- | :--- | :--- | :--- |
| `-All` | - | `Switch` | Lists all licensed users in the tenant. |
| `-LicenseFilter` | `-License`, `-Plan` | `String` | Filters users matching a partial license or mailbox plan name. |
| `-Details` | `-d`, `-Detailed` | `Switch` | Displays extended user attributes in additional columns. |
| `-Search` | `-FilterUser`, `-User` | `String` | Searches users by DisplayName, email, or department. |
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
- **Version**: 1.2.0
