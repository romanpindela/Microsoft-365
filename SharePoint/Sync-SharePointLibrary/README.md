# Sync-SharePointLibrary

Automated PowerShell utility designed to map and synchronize SharePoint Online document libraries directly to the local Windows 11 / Windows 10 File Explorer using the OneDrive for Business sync client.

Developed for **Elektrimont Sp. z o.o.** internal infrastructure management.

---

## Author & Project Metadata

* **Author**: Roman Pindela
* **Email**: roman.pindela@gmail.com
* **GitHub**: [https://github.com/romanpindela](https://github.com/romanpindela)
* **Organization**: Elektrimont Sp. z o.o.
* **Version**: 1.0.0
* **Target Environment**: Windows 10, Windows 11, Windows Server (Desktop Experience)

---

## Prerequisites & Requirements

### 1. User Sign-in State (OneDrive App)
* The executing user **must be actively logged in** to the local **OneDrive for Business** desktop client.
* **Credentials required**: Corporate Microsoft 365 work email address and domain password (the same identity used for Outlook and Microsoft 365 access).
> **[PL] Ważne:** Użytkownik musi być zalogowany do aplikacji OneDrive na swoim komputerze. Dane do logowania to Twój firmowy adres e-mail i hasło (dokładnie takie jak w Outlooku i portalu Office 365).

### 2. Permissions
* Access to document libraries is bound by role-based access control (RBAC). Only users with granted permissions (assigned directly or via Microsoft 365 Security Groups) will receive file synchronization.

### 3. Execution Policy & Unblock File
When downloading the script from an external repository or GitHub, Windows Defender SmartScreen / PowerShell ExecutionPolicy will flag the script with the alternate data stream (Zone.Identifier). 

Unblock the file prior to execution:
```powershell
Unblock-File -Path .\Sync-SharePointLibrary.ps1