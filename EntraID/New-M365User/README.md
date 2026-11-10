# New-M365User

Enterprise-grade PowerShell provisioning script for creating, licensing, and configuring new user accounts in **Microsoft Entra ID / Microsoft 365** and assigning them to **Microsoft Teams** / Microsoft 365 Groups.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2F%20v2.0-008272.svg)](https://learn.microsoft.com/graph/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

---

## Overview

`New-M365User.ps1` standardizes and automates onboarding and identity provisioning in Microsoft Entra ID. It enforces organizational security baselines:
- Mandatory password reset on first sign-in (`forceChangePasswordNextSignIn = $true`).
- Least privilege model: new accounts are created strictly as standard users without administrative Entra directory roles.
- Flexible licensing assignment supporting both friendly license plan names (`SkuPartNumber` such as `ENTERPRISEPACK`, `SPE_E3`, `O365_BUSINESS_PREMIUM`) and direct GUID identifiers (`SkuId`).
- Automated enforcement of the two-letter ISO 3166-1 alpha-2 `UsageLocation` country code (defaults to `PL`).
- Group and team onboarding: automatically resolves Microsoft 365 / Teams groups by `DisplayName` and adds the new user as a member.
- Real-time color-coded console telemetry and persistent audit logging to `C:\Logs\Microsoft-365\New-M365User_YYYY-MM-DD.log`.

---

## Key Features

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
[Script Initialization & Environment Validation]
                   │
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

| Parameter | Type | Required | Default | Description |
| :--- | :--- | :---: | :--- | :--- |
| `UserPrincipalName` | `[string]` | **Yes** | - | UPN / primary email of the new user (e.g. `jan.kowalski@domena.pl`). |
| `DisplayName` | `[string]` | **Yes** | - | Full display name (e.g. `Jan Kowalski`). |
| `GivenName` | `[string]` | No | `$null` | First name (e.g. `Jan`). |
| `Surname` | `[string]` | No | `$null` | Last name / surname (e.g. `Kowalski`). |
| `MailNickname` | `[string]` | No | Derived from UPN | Mail alias (default: prefix before `@` in UPN). |
| `UsageLocation` | `[string]` | No | `'PL'` | Two-letter ISO 3166-1 alpha-2 country code required for licensing. |
| `InitialPassword` | `[securestring]` | **Yes** | - | Initial password provided as a `SecureString`. |
| `LicenseSku` | `[string]` | **Yes** | - | License plan name (`SkuPartNumber` e.g. `SPE_E3`) or GUID (`SkuId`). |
| `TeamGroupNames` | `[string[]]` | No | `@()` | Array of Teams / M365 Group display names to add the user to. |
| `LogDirectory` | `[string]` | No | `'C:\Logs\Microsoft-365'` | Directory path for persistent audit logs. |
| `WhatIf` | `[switch]` | No | `$false` | Simulates provisioning without making real changes. |

---

## Usage Examples

### 1. Basic User Provisioning with License
```powershell
$secPass = "P@ssw0rd2026!" | ConvertTo-SecureString -AsPlainText -Force

.\New-M365User.ps1 -UserPrincipalName "jan.kowalski@domena.pl" `
                   -DisplayName "Jan Kowalski" `
                   -InitialPassword $secPass `
                   -LicenseSku "SPE_E3"
```

### 2. Full Provisioning with GivenName, Surname, and Teams Groups
```powershell
$secPass = Read-Host "Wprowadz haslo poczatkowe" -AsSecureString

.\New-M365User.ps1 -UserPrincipalName "anna.nowak@domena.pl" `
                   -DisplayName "Anna Nowak" `
                   -GivenName "Anna" `
                   -Surname "Nowak" `
                   -InitialPassword $secPass `
                   -LicenseSku "ENTERPRISEPACK" `
                   -TeamGroupNames "Dzial Handlowy", "Ogolny"
```

### 3. Dry-Run Simulation (`-WhatIf`)
```powershell
$secPass = "Startowe123!" | ConvertTo-SecureString -AsPlainText -Force

.\New-M365User.ps1 -UserPrincipalName "tomasz.wisniewski@domena.pl" `
                   -DisplayName "Tomasz Wisniewski" `
                   -InitialPassword $secPass `
                   -LicenseSku "c707743b-110d-45b0-802c-1c83b2162703" `
                   -WhatIf
```

---

## Logging & Output Format

### Persistent File Log (`C:\Logs\Microsoft-365\New-M365User_YYYY-MM-DD.log`)
```text
[2026-11-10 08:35:12] [INFO]    Inicjalizacja tworzenia uzytkownika: jan.kowalski@domena.pl
[2026-11-10 08:35:13] [SUCCESS] Konto uzytkownika utworzone (ID: 11111111-2222-3333-4444-555555555555)
[2026-11-10 08:35:14] [INFO]    Przypisywanie licencji: ENTERPRISEPACK...
[2026-11-10 08:35:15] [SUCCESS] Licencja przypisana pomyslnie.
[2026-11-10 08:35:16] [WARN]    Grupa 'Dzial Sprzedazy' nie zostala znaleziona w dzierzawie.
[2026-11-10 08:35:17] [INFO]    Zakonczenie procesu. Status koncowy: Sukces (z ostrzezeniami).
```

### Console Executive Summary
```text
================================================================================
                    PODSUMOWANIE PROVISIONINGU UZYTKOWNIKA
================================================================================
UserPrincipalName    : jan.kowalski@domena.pl
DisplayName          : Jan Kowalski
UserId (GUID)        : 11111111-2222-3333-4444-555555555555
Rola                 : Standard User (brak rol uprzywilejowanych)
Licencja             : SPE_E3 (Wymuszone UsageLocation: PL)
Wymuszenie hasla     : Tak (przy pierwszym logowaniu)
Grupy Teams          : [Dzial Handlowy: DODANY], [Ogolny: DODANY]
Plik Logu            : C:\Logs\Microsoft-365\New-M365User_2026-11-10.log
Status Koncowy       : SUKCES
================================================================================
```

---

## Author & License

- **Author**: Roman Pindela
- **Email**: roman.pindela@gmail.com
- **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
- **License**: MIT
