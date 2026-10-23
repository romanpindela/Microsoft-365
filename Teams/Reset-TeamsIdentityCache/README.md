# Reset-TeamsIdentityCache

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-blue.svg)](https://microsoft.com/powershell)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20%7C%20Server-lightgrey.svg)](https://www.microsoft.com/windows)
[![Version](https://img.shields.io/badge/Version-1.0.0-green.svg)](https://github.com/romanpindela)
[![License](https://img.shields.io/badge/License-MIT-brightgreen.svg)](LICENSE)

An enterprise-grade, clean-code PowerShell remediation tool built for **Microsoft 365 Helpdesk** technicians and system administrators. It resolves persistent Microsoft Teams authentication errors, credential loops (e.g. AADSTS, OAuth token expiration, WAM `0x80070002`), tenant-switching glitches, and local profile corruption across both **New Microsoft Teams (MSIX/AppX)** and legacy **Classic Teams (Win32/Electron)**.

---

## Author & Project Metadata

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
* **Script Version**: `1.0.0`
* **Release Date**: 2026-10-19
* **License**: MIT
* **Target Platforms**: Windows 10, Windows 11, Windows Server (Desktop Experience)
* **OS Localization**: Fully language-independent (Polish & English Windows compatible via Well-Known SIDs)

---

## Why Reset Teams Identity Cache?

In enterprise Microsoft 365 environments, Microsoft Teams relies on a multi-layered identity stack:
1. **New Teams AppX Cache**: Web storage, indexed databases, and LocalCache in `%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe`.
2. **Classic Teams Cache**: Legacy subdirectories in `%APPDATA%\Microsoft\Teams`.
3. **Microsoft 365 IdentityCache**: `%LOCALAPPDATA%\Microsoft\IdentityCache` storing tenant authentication payloads.
4. **OneAuth Repository**: Modern authentication token cache in `%LOCALAPPDATA%\Microsoft\OneAuth`.
5. **Windows Web Account Manager (WAM)**: TokenBroker cache in `%LOCALAPPDATA%\Microsoft\TokenBroker\Cache`.
6. **Windows Credential Manager**: Generic credentials for `msteams*`, `MicrosoftAccount:target=SSO_POP_*`, and `Windows.Security.Authentication.Web.TokenCache*`.

When tokens desynchronize, users experience:
- Endless "We couldn't sign you in. Please try again" loops.
- Error codes: `AADSTS50020`, `AADSTS50058`, `AADSTS50140`, or `0x80070002`.
- Stale guest tenant profiles that cannot be removed or switched.
- Desktop client displaying a blank or white screen on startup.

`Reset-TeamsIdentityCache.ps1` safely purges these layers and restores Microsoft Teams to a clean, authenticated state.

---

## Key Features & Capabilities

The script features a built-in feature catalog (`-ListFeatures`) detailing each module:

| # | Feature | Parameter | Target / Scope | Description |
| :-: | :--- | :--- | :--- | :--- |
| **1** | **Process Termination** | `-KillTeamsProcesses` | `ms-teams.exe`, `teams.exe`, `msedgewebview2.exe` | Gracefully closes running Teams instances and child WebView2 broker processes to release locks on cache files. |
| **2** | **New Teams Cache Reset** | `-ClearTeamsNewCache` | `%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\` | Purges `LocalCache`, `TempState`, `AC` and triggers native `Reset-AppxPackage` on Windows 10/11. |
| **3** | **Classic Teams Cache Reset** | `-ClearTeamsClassicCache` | `%APPDATA%\Microsoft\Teams\*` | Cleans `blob_storage`, `Cache`, `databases`, `GPUCache`, `IndexedDB`, `Local Storage`, and `tmp`. |
| **4** | **IdentityCache Purge** | `-ClearIdentityCache` | `%LOCALAPPDATA%\Microsoft\IdentityCache` | Clears cached OAuth tokens and tenant identity state files. |
| **5** | **OneAuth Cache Purge** | `-ClearOneAuthCache` | `%LOCALAPPDATA%\Microsoft\OneAuth` | Cleans modern authentication client tokens shared by Office/M365 applications. |
| **6** | **WAM TokenBroker Reset** | `-ClearTokenBrokerCache` | `%LOCALAPPDATA%\Microsoft\TokenBroker\Cache` | Clears Web Account Manager tokens to eliminate stale broker tickets. |
| **7** | **Credential Manager Cleanup** | `-ClearCredentials` | Generic Credentials (`msteams*`, `TokenCache*`) | Enumerates and safely removes cached Teams credentials via `cmdkey` without touching Git or OS credentials. |
| **8** | **Full End-to-End Reset** | `-Reset` | All components (1 through 7 above) | Executes the recommended complete remediation procedure in a single step. |
| **9** | **Automatic Teams Relaunch** | `-RestartTeams` | `ms-teams:` protocol / executable | Automatically launches Microsoft Teams after cleanup completes. |
| **10** | **Simulation / Dry Run** | `-WhatIf` | Non-destructive audit | Analyzes and reports all files, processes, and credentials that would be deleted. |
| **11** | **UAC Auto-Elevation** | `-Elevate` | Elevated Administrator Token | Re-launches the script with elevated privileges via UAC prompt if needed. |
| **12** | **M365 Admin Connectivity** | `-ConnectM365` | Exchange Online / Graph PowerShell | Connects to Microsoft 365 Exchange Online for tenant/mailbox diagnosis (Standard 14). |
| **13** | **Enterprise Audit Logging** | `-LogDirectory` | `C:\Logs` (fallback: `%LOCALAPPDATA%\Logs`) | Records structured timestamps, actions, and status summary tables for Helpdesk ticketing. |

---

## Prerequisites & Security Guidelines

### 1. User Profile Scope
> [!NOTE]
> Microsoft Teams caches, WAM tokens, and Credential Manager entries reside inside the **active user profile** (`%LOCALAPPDATA%`, `%APPDATA%`).
> - Under ordinary circumstances, run this script inside the user's interactive Windows session without elevation.
> - If run under an elevated secondary administrator account, the script alerts the technician that it is executing against the administrator's profile path.

### 2. Language Independence
> [!TIP]
> The script is completely bilingual and language-independent.
> - Administrator verification utilizes the **Well-Known SID** `S-1-5-32-544` (`BuiltInAdministratorsSid`), functioning identically on English ("Administrators") and Polish ("Administratorzy") Windows systems.
> - Credential Manager parsing matches both English (`Target:`) and Polish (`Element docelowy:`) headers.

### 3. Unblocking the Script After Download
When downloaded from GitHub or external sources, Windows marks files with an execution block (Alternate Data Stream `Zone.Identifier`). Always unblock the script prior to execution:

```powershell
Unblock-File -Path .\Reset-TeamsIdentityCache.ps1
```

---

## Parameter Reference

| Parameter | Aliases | Type | Default | Description |
| :--- | :--- | :---: | :---: | :--- |
| `-Reset` | `-Run`, `-Execute`, `-All` | `Switch` | `False` | Executes full end-to-end reset across all components. |
| `-ClearTeamsNewCache` | `-ClearNewTeams`, `-NewTeams` | `Switch` | `False` | Clears New Teams AppX LocalCache, TempState, and resets package. |
| `-ClearTeamsClassicCache` | `-ClearClassicTeams`, `-ClassicTeams` | `Switch` | `False` | Clears legacy Classic Teams subfolders in `%APPDATA%\Microsoft\Teams`. |
| `-ClearIdentityCache` | `-IdentityCache`, `-Identity` | `Switch` | `False` | Purges `%LOCALAPPDATA%\Microsoft\IdentityCache`. |
| `-ClearOneAuthCache` | `-OneAuthCache`, `-OneAuth` | `Switch` | `False` | Purges `%LOCALAPPDATA%\Microsoft\OneAuth`. |
| `-ClearTokenBrokerCache` | `-TokenBrokerCache`, `-TokenBroker`, `-WAM` | `Switch` | `False` | Purges `%LOCALAPPDATA%\Microsoft\TokenBroker\Cache`. |
| `-ClearCredentials` | `-ClearCreds`, `-Credentials` | `Switch` | `False` | Deletes Teams generic credentials from Windows Credential Manager. |
| `-CredentialPattern` | `-CredFilter`, `-Pattern` | `String` | `'teams'` | Target matching pattern for Credential Manager entries. |
| `-KillTeamsProcesses` | `-KillProcesses`, `-StopTeams` | `Switch` | `False` | Terminates `ms-teams.exe`, `teams.exe`, and related WebView2 processes. |
| `-RestartTeams` | `-Relaunch`, `-StartTeams` | `Switch` | `False` | Relaunches Microsoft Teams after maintenance completes. |
| `-Force` | `-f`, `-Quiet`, `-Silent` | `Switch` | `False` | Suppresses interactive user confirmation prompt. |
| `-WhatIf` | `-DryRun`, `-Simulate` | `Switch` | `False` | Performs a dry-run audit without deleting any data or credentials. |
| `-ListFeatures` | `-ShowFeatures`, `-Features` | `Switch` | `False` | Displays the detailed feature catalog and exit. |
| `-LogDirectory` | `-LogPath`, `-LogDir` | `String` | `'C:\Logs'` | Destination directory for execution log files. |
| `-Elevate` | `-RunAsAdmin`, `-Admin` | `Switch` | `False` | Relaunches script with elevated UAC permissions if not already admin. |
| `-ConnectM365` | `-ConnectExchange`, `-M365Admin` | `Switch` | `False` | Initiates Microsoft 365 Exchange Online administrator authentication. |
| `-AdminUserPrincipalName` | `-AdminUPN`, `-AdminEmail` | `String` | *None* | Admin account UPN for non-interactive / explicit Exchange Online login. |
| `-SkipM365Connection` | `-NoM365`, `-SkipExchange` | `Switch` | `False` | Suppresses Microsoft 365 connection checks. |
| `-Help` | `-h` | `Switch` | `False` | Displays the interactive help guide and examples. |

---

## Usage Examples

### 1. Display Help & Usage Reference
Running the script without parameters or with `-Help` displays comprehensive documentation:
```powershell
.\Reset-TeamsIdentityCache.ps1
# or
.\Reset-TeamsIdentityCache.ps1 -Help
```

### 2. List Detailed Features Catalog
Displays the technical details, target paths, and parameters for all features:
```powershell
.\Reset-TeamsIdentityCache.ps1 -ListFeatures
```

### 3. Full Recommended Reset (Interactive Confirmation)
Stops Teams, clears all caches (New Teams, Classic Teams, IdentityCache, OneAuth, WAM TokenBroker), and removes Teams credentials:
```powershell
.\Reset-TeamsIdentityCache.ps1 -Reset
```

### 4. Silent Full Reset with Automatic Teams Restart
Ideal for remote Helpdesk deployment (e.g. via NinjaRMM, Intune, or remote PowerShell):
```powershell
.\Reset-TeamsIdentityCache.ps1 -Reset -RestartTeams -Force
```

### 5. Non-Destructive Simulation (Dry Run)
Audits the workstation and lists exactly how many files, processes, and credentials would be affected:
```powershell
.\Reset-TeamsIdentityCache.ps1 -Reset -WhatIf
```

### 6. Granular Reset: New Teams & WAM Broker Only
Targets specifically the New Teams client and Windows Web Account Manager cache:
```powershell
.\Reset-TeamsIdentityCache.ps1 -ClearTeamsNewCache -ClearTokenBrokerCache -KillTeamsProcesses -Force
```

### 7. Credential Manager Cleanup Only
Removes cached Teams credentials from Windows Credential Manager:
```powershell
.\Reset-TeamsIdentityCache.ps1 -ClearCredentials -Force
```

### 8. Run Full Reset with Exchange Online Administrator Login
Connects to Microsoft 365 Exchange Online as an admin to inspect user mailbox / licensing after resetting the local cache:
```powershell
.\Reset-TeamsIdentityCache.ps1 -Reset -ConnectM365 -AdminUserPrincipalName "admin@contoso.com"
```

---

## Execution Summary & Logging

Every execution generates real-time color-coded console telemetry and appends structured records to a timestamped log file:
- **Default Path**: `C:\Logs\<yyyyMMdd_HHmmss>-Reset-TeamsIdentityCache-<ComputerName>-<UserName>.log`
- **Fallback**: `%LOCALAPPDATA%\Logs\` (if `C:\Logs` cannot be created or accessed)

### Sample Summary Report Output:
```text
================================================================================
                    EXECUTION SUMMARY & AUDIT REPORT
================================================================================
Script       : Reset-TeamsIdentityCache.ps1 v1.0.0
Executed By  : WORKSTATION-01\john.doe
Duration     : 00m 02s
Log File     : C:\Logs\20261019_114846-Reset-TeamsIdentityCache-WORKSTATION-01-john.doe.log
Mode         : LIVE EXECUTION
================================================================================

Component                               Action                     Status    Details
---------                               ------                     ------    -------
Process Termination                     Stop-Process               Success   ms-teams (PID: 21060)
New Teams LocalCache                    Clean Files                Success   3010 root item(s) purged
New Teams AppX                          Reset-AppxPackage          Success   MSTeams_8wekyb3d8bbwe
Classic Teams (blob_storage)            Clean Files                Clean     0 items
Microsoft IdentityCache                 Clean Files                Success   88 root item(s) purged
OneAuth Cache                           Clean Files                Success   51 root item(s) purged
WAM TokenBroker Cache                   Clean Files                Success   57 root item(s) purged
Credential Manager                      cmdkey /delete             Success   msteams_adalsso
Relaunch Teams                          Start-Process              Success   Protocol ms-teams:
================================================================================
```

---

## Troubleshooting & Best Practices

1. **User sees "We ran into an issue" after reset:**
   - Allow 10-15 seconds for Microsoft Teams to rebuild its local web databases upon first launch.
   - Ensure the user enters their corporate Microsoft 365 credentials when prompted.
2. **Locked File Warnings:**
   - If files in `%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache` cannot be deleted, ensure background Outlook or Edge WebView2 processes have been closed using `-KillTeamsProcesses`.
3. **Remote Execution via RMM / PowerShell Remoting:**
   - When running via PsExec or remote sessions, ensure the command executes within the target logged-on user's console token to clean their specific profile.

---

## License

This project is licensed under the [MIT License](LICENSE).
