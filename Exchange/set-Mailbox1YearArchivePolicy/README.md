# Exchange Online - Auto-Enable Archive & 1-Year Retention Policy

A robust, idempotent PowerShell automation script to automatically enable Online Archiving (if disabled) and enforce a 1-year (365-day) archive retention policy on any Exchange Online mailbox (User Mailbox or Shared Mailbox)[cite: 5].

## Key Features

- **Automatic Online Archive Activation:** Inspects the current archive status of the target mailbox. If Online Archive is not active, it enables it automatically (`Enable-Mailbox -Archive`).
- **Safe & Non-Destructive:** Uses the `MoveToArchive` retention action. Never purges, permanently deletes, or modifies mailbox permissions.
- **Idempotent Deployment:** Verifies the existence of retention tags and policies before creation to ensure clean repeated runs without errors.
- **Interactive Modern Authentication:** Detects active sessions and automatically triggers interactive sign-in (`Connect-ExchangeOnline`) supporting MFA if unauthenticated[cite: 5].
- **Dependency Management:** Automatically detects and installs the `ExchangeOnlineManagement` module if missing[cite: 5].
- **Sanitized Input:** Guards against malicious input or formatting injection via parameter sanitation[cite: 5].
- **Replication-Aware Exception Handling:** Gracefully handles transient cloud replication delays when invoking `Start-ManagedFolderAssistant`.

---

## Important Security Notice

When downloading PowerShell scripts directly from GitHub, Windows tags the file with a `Zone.Identifier` stream marking it as untrusted, which prevents execution under default policies[cite: 5].

Before running the script, unblock the file in PowerShell[cite: 5]:

```powershell
Unblock-File -Path .\set-Mailbox1YearArchivePolicy.ps1

```

---

## Prerequisites

1. **PowerShell:** Windows PowerShell 5.1 or PowerShell Core (7.x+).


2. **Administrator Role:** Exchange Administrator, Compliance Administrator, or Global Administrator privileges in Microsoft 365.


3. **Execution Policy:** Ensure your execution policy permits script execution:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

```



---

## Usage Examples

### 1. Display Built-in Help

Running without parameters or using `-Help` / `-h` prints full command documentation, syntax, and author info:

```powershell
.\set-Mailbox1YearArchivePolicy.ps1 -Help
# or
.\set-Mailbox1YearArchivePolicy.ps1 -h

```

### 2. Enable Archive and Apply 1-Year Retention Policy

Target a mailbox using its primary SMTP address, UPN, or alias:

```powershell
.\set-Mailbox1YearArchivePolicy.ps1 -Identity "jurnaling@elektrimont.pl"

```

---

## Execution Workflow

1. **Authenticate:** Checks connection to Exchange Online; opens modern sign-in if required.


2. **Validate Target:** Resolves mailbox object.
3. **Enable Archive:** If `ArchiveStatus` is not `Active`, runs `Enable-Mailbox -Archive`.
4. **Create Tag:** Ensures `Default 1 Year Move to Archive` tag exists (Type: `All`, Action: `MoveToArchive`, Retention: `365 days`).
5. **Create Policy:** Ensures `Shared Mailbox 1 Year Archive` policy exists (linking the archive tag and standard `Junk Email` tag).
6. **Assign Policy:** Applies the policy to the mailbox.
7. **Schedule Migration:** Calls `Start-ManagedFolderAssistant` to schedule background data processing.

---

## Author & Contact

* **Author:** Roman Pindela


* **Email:** roman.pindela@gmail.com


* **GitHub:** [https://github.com/romanpindela](https://github.com/romanpindela?utm_source=gemini)

* **Version:** 1.1.0


