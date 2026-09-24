# Manage-MailboxUserAccess

A comprehensive PowerShell automation script for Microsoft 365 Exchange Online administrators to manage mailbox access permissions. It allows administrators to **grant** `FullAccess` rights (with optional **AutoMapping** control), **revoke** existing permissions, run silent automated tasks via **Force**, and audit existing permissions.

## Features

- **Grant & Revoke Permissions**: Assign or remove `FullAccess` rights on both `UserMailbox` and `SharedMailbox` objects.
- **AutoMapping Suppression**: Optional `-DisableAutoMapping` switch prevents Outlook Desktop from auto-mounting delegated mailboxes.
- **Confirmation & Automation**: Includes interactive `(Y/N)` confirmation prompts by default, with a `-Force` switch to bypass prompts for CI/CD or scripting.
- **Audit Mode**: Use `-ShowPermissionsOnly` to inspect explicit mailbox permissions without modifying any settings.
- **Identity Pre-Validation**: Validates that both the target mailbox and the specified user exist in the Microsoft 365 tenant before taking action.
- **Session Auto-Detection**: Checks for active Exchange Online sessions and initiates modern administrator login if not connected.

## Prerequisites

- PowerShell 5.1 or PowerShell 7+
- `ExchangeOnlineManagement` module:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

```

* Exchange Online administrative role (e.g., Exchange Administrator or Global Administrator).

## Security & Unblocking

When downloading scripts from GitHub, unblock the file before running:

```powershell
Unblock-File -Path .\Manage-MailboxUserAccess.ps1

```

## Usage Examples

### 1. View Mailbox Permissions Only (Audit Mode)

```powershell
.\Manage-MailboxUserAccess.ps1 -MailboxIdentity "sales@domain.com" -ShowPermissionsOnly

```

### 2. Grant Access with AutoMapping Disabled (Interactive)

```powershell
.\Manage-MailboxUserAccess.ps1 -MailboxIdentity "sales@domain.com" -UserIdentity "admin@domain.com" -DisableAutoMapping

```

### 3. Grant Access Silently (Automated / Non-Interactive)

```powershell
.\Manage-MailboxUserAccess.ps1 -MailboxIdentity "sales@domain.com" -UserIdentity "admin@domain.com" -Force

```

### 4. Revoke Access with Interactive Confirmation

```powershell
.\Manage-MailboxUserAccess.ps1 -MailboxIdentity "sales@domain.com" -UserIdentity "admin@domain.com" -RemoveAccess

```

### 5. Revoke Access Silently

```powershell
.\Manage-MailboxUserAccess.ps1 -MailboxIdentity "sales@domain.com" -UserIdentity "admin@domain.com" -RemoveAccess -Force

```

### 6. Display Built-In Help

```powershell
.\Manage-MailboxUserAccess.ps1 -h

```


## Screenshots & Examples

### Info run
![PowerShell Output](assets/InfoRun.jpg)

### Changing persmission
![HTML Report](assets/Showing_and_changing_permission_to_userinbox.jpg)

---

## Author & Contact

* **Author**: Roman Pindela
* **Email**: [roman.pindela@gmail.com](https://www.google.com/search?q=mailto%3Aroman.pindela%40gmail.com)
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela?utm_source=gemini)
* **Version**: 2.1.0

```

