# New-M365User

Enterprise-grade PowerShell provisioning script for creating, licensing, and configuring new user accounts in **Microsoft Entra ID / Microsoft 365** and assigning them to **Microsoft Teams** / Microsoft 365 Groups.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2F%20v2.0-008272.svg)](https://learn.microsoft.com/graph/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

---

## Overview

`New-M365User.ps1` standardizes and automates onboarding and identity provisioning in Microsoft Entra ID. It enforces organizational security baselines:
- **Interactive Console Help (`-Help`, `-h`, or executing without arguments)**: Running the script without parameters displays comprehensive usage instructions, parameter details, and copy-paste ready examples.
- **Mandatory password reset on first sign-in** (`forceChangePasswordNextSignIn = $true`).
- **Profile Configuration**: Populates GivenName, Surname, Department, and automatically calculates DisplayName and MailNickname if omitted.
- **Least privilege model**: New accounts are created strictly as standard users without administrative Entra directory roles.
- **Flexible licensing assignment**: Supports both friendly license plan names (`SkuPartNumber` such as `ENTERPRISEPACK`, `SPE_E3`, `O365_BUSINESS_PREMIUM`) and direct GUID identifiers (`SkuId`).
- **Automated UsageLocation**: Enforces the two-letter ISO 3166-1 alpha-2 country code (defaults to `PL`).
- **Microsoft Teams Onboarding**: Automatically resolves Microsoft 365 / Teams groups by `DisplayName` and adds the new user as a member.
- **Audit Logging**: Real-time color-coded console telemetry and persistent audit logging to `C:\Logs\Microsoft-365\New-M365User_YYYY-MM-DD.log`.

---

## Key Features

- **Built-in Interactive Help**: Running `.\New-M365User.ps1` without arguments, or using `-Help` / `-h`, displays usage, parameter descriptions, and examples without throwing errors.
- **Strict Credential Security**: Initial password must be provided as a `[SecureString]`. Passwords never leak into plain-text PowerShell console history (`PSReadLine`), Script Block Logging, or disk log files.
- **Tenant License Capacity Validation**: Queries available tenant licenses (`Get-MgSubscribedSku`), verifies plan existence, and alerts administrators if all purchased units are consumed.
- **Fault-Tolerant Execution**:
  - Independent `try/catch` boundaries for identity creation, licensing, and group assignments.
  - If group membership or license assignment encounters an issue, the user account remains safely created with full diagnostics logged.
  - Non-existing or misspelled Teams group names emit structured `[WARN]` entries without breaking subsequent group assignments.
- **Support for Simulation (`-WhatIf`)**: Native `SupportsShouldProcess` implementation allows administrators to preview account provisioning and group additions before execution.
- **Formatted Executive Summary**: Generates a clear status report in the console upon completion.

---

## Workflow Architecture

```
[Script Execution]
       │
       ├──> [No Parameters or -Help/-h] ──> [Display Formatted Help Screen & Exit]
       ▼
[Verify/Initialize Logging Directory: C:\Logs\Microsoft-365\]
       │
       ▼
[Verify Active Microsoft Graph Session & Scopes]
       │
       ▼
[Validate License SKU (SkuPartNumber or SkuId) & Capacity]
       │
       ▼
[Verify UserPrincipalName Uniqueness in Tenant]
       │
       ├──> [UPN Exists: Terminating error, log, exit]
       ▼
[Create Entra ID User (New-MgUser)]
   - AccountEnabled: $true
   - Profile: GivenName, Surname, Department, DisplayName
   - ForceChangePasswordNextSignIn: $true
   - Role: Standard User (no privileged roles)
       │
       ▼
[Set UsageLocation & Assign License (Set-MgUserLicense)]
       │
       ▼
[Iterate over Microsoft Teams Groups (TeamGroupNames)]
   - Find Group ID by DisplayName
   - Add Member (New-MgGroupMemberByRef)
       │
       ▼
[Display Formatted Console Summary]
       │
       ▼
[Finalize Audit Log & Return Exit Code]
```

---

## Prerequisites

### 1. PowerShell Version
- PowerShell 7+ (recommended) or Windows PowerShell 5.1.

### 2. Microsoft Graph PowerShell SDK Modules
```powershell
Install-Module -Name Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups -Scope CurrentUser
```

### 3. Microsoft Graph API Permissions (Scopes)
The executing administrator requires consent for the following scopes:
- `User.ReadWrite.All`
- `Directory.ReadWrite.All`
- `Group.ReadWrite.All`

Connect to Microsoft Graph prior to running the script:
```powershell
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Group.ReadWrite.All"
```

---

## Parameters

| Parameter | Type | Required | Default | Aliases | Description |
| :--- | :--- | :---: | :--- | :--- | :--- |
| `UserPrincipalName` | `[string]` | **Yes** | - | `Email`, `UPN`, `User` | UPN / primary email address of the new user. |
| `GivenName` | `[string]` | No | `$null` | `FirstName`, `Imie` | First name of the user. |
| `Surname` | `[string]` | No | `$null` | `LastName`, `Nazwisko` | Last name / surname of the user. |
| `Department` | `[string]` | No | `$null` | `Dzial` | Organizational department or business unit. |
| `DisplayName` | `[string]` | No | Auto | - | Full display name (auto-composed if omitted). |
| `MailNickname` | `[string]` | No | Derived | - | Mail alias (default: prefix before `@` in UPN). |
| `UsageLocation` | `[string]` | No | `'PL'` | - | ISO 3166-1 alpha-2 country code for licensing. |
| `InitialPassword` | `[securestring]` | **Yes** | - | `Password`, `Haslo` | Initial password provided as a `SecureString`. |
| `LicenseSku` | `[string]` | **Yes** | - | `License`, `Sku`, `Licencja` | License plan (`SkuPartNumber` or `SkuId` GUID). |
| `TeamGroupNames` | `[string[]]` | No | `@()` | `TeamGroupName`, `TeamsGroup`, `GrupyTeams` | Array of Teams / M365 Group display names. |
| `LogDirectory` | `[string]` | No | `'C:\Logs\Microsoft-365'` | - | Directory path for persistent audit logs. |
| `WhatIf` | `[switch]` | No | `$false` | - | Simulates provisioning without making real changes. |
| `Help` | `[switch]` | No | `$false` | `h` | Displays custom help screen with examples and exits. |

---

## Usage Examples

### 1. View Interactive Help (No parameters or -Help)
```powershell
# Display help:
.\New-M365User.ps1 -Help
# Or with alias:
.\New-M365User.ps1 -h
# Or simply execute without parameters:
.\New-M365User.ps1
```

### 2. Complete User Provisioning with Department & Teams Groups
```powershell
$secPass = "InitialPass123!" | ConvertTo-SecureString -AsPlainText -Force

.\New-M365User.ps1 -UserPrincipalName "john.doe@domain.com" `
                   -GivenName "John" `
                   -Surname "Doe" `
                   -Department "IT / Helpdesk" `
                   -TeamGroupNames "Technical Support", "General" `
                   -LicenseSku "SPE_E3" `
                   -InitialPassword $secPass
```

### 3. Interactive Masked Password Input
```powershell
$secPass = Read-Host "Enter initial password" -AsSecureString

.\New-M365User.ps1 -Email "jane.smith@domain.com" `
                   -GivenName "Jane" `
                   -Surname "Smith" `
                   -Department "Sales" `
                   -TeamGroupNames "Global Sales", "Commercial" `
                   -LicenseSku "ENTERPRISEPACK" `
                   -InitialPassword $secPass
```

### 4. Dry-Run Simulation (`-WhatIf`)
```powershell
$secPass = "Test1234!" | ConvertTo-SecureString -AsPlainText -Force

.\New-M365User.ps1 -UserPrincipalName "alex.turner@domain.com" `
                   -GivenName "Alex" `
                   -Surname "Turner" `
                   -Department "Finance" `
                   -TeamGroupNames "Finance and Accounting" `
                   -LicenseSku "SPE_E3" `
                   -InitialPassword $secPass `
                   -WhatIf
```

---

## Logging & Output Format

### Persistent File Log (`C:\Logs\Microsoft-365\New-M365User_YYYY-MM-DD.log`)
```text
[2026-11-10 08:35:12] [INFO]    Initializing user provisioning for: john.doe@domain.com
[2026-11-10 08:35:13] [SUCCESS] User account successfully created (ID: 11111111-2222-3333-4444-555555555555)
[2026-11-10 08:35:14] [INFO]    Assigning license: ENTERPRISEPACK...
[2026-11-10 08:35:15] [SUCCESS] License assigned successfully.
[2026-11-10 08:35:16] [WARN]    Group 'Sales Department' was not found in the tenant.
[2026-11-10 08:35:17] [INFO]    Process completed. Final status: SUCCESS (WITH WARNINGS).
```

### Console Executive Summary
```text
================================================================================
                    USER PROVISIONING SUMMARY
================================================================================
UserPrincipalName    : john.doe@domain.com
DisplayName          : John Doe
UserId (GUID)        : 11111111-2222-3333-4444-555555555555
Department           : IT / Helpdesk
Role                 : Standard User (no privileged directory roles)
License              : SPE_E3 (Enforced UsageLocation: PL)
Force Password Change: Yes (at next sign-in)
Teams Groups         : [Technical Support: ADDED], [General: ADDED]
Log File             : C:\Logs\Microsoft-365\New-M365User_2026-11-10.log
Final Status         : SUCCESS
================================================================================
```

---

## Author & License

- **Author**: Roman Pindela
- **Email**: roman.pindela@gmail.com
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **License**: MIT
