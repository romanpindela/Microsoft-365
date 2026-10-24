# Get-M365UserAccessAudit

Enterprise-grade PowerShell access audit and reporting utility for **Microsoft 365** and **Microsoft Entra ID**. Performs a comprehensive, end-to-end security audit of permissions, memberships, and administrative roles across **Entra ID**, **Microsoft Teams**, **SharePoint Online**, and **Exchange Online** for designated users.

---

## 1. Overview & Architecture

`Get-M365UserAccessAudit.ps1` constructs a unified access matrix mapping all direct, nested, and delegated privileges for an identity across Microsoft 365 workloads:

```
                                  [ Target Identity (UPN / ObjectId) ]
                                                   │
        ┌───────────────────┬──────────────────────┼──────────────────────┬───────────────────┐
        ▼                   ▼                      ▼                      ▼                   ▼
  [ Entra ID Groups ]  [ Directory Roles ]    [ MS Teams ]           [ SharePoint ]      [ Exchange Online ]
  - Direct & Nested    - Direct Assignments   - Team Memberships     - M365 Root Sites   - Dist. Lists
  - Security & Unified - Group Inherited      - Channel Access       - Site URLs         - Mailbox Perms:
  - Role-Assignable    - PIM Eligible Sched.    (Private & Shared)   - Site Permissions    * FullAccess
  - Dynamic / Assigned                                                                     * SendAs
  - Cloud / On-Prem AD                                                                     * SendOnBehalf
```

### Key Pillars:
1. **Entra ID Group Membership (Direct vs. Transitive):**
   - Maps Cloud and On-Premises synchronized security groups, Microsoft 365 (Unified) groups, and mail-enabled groups.
   - Resolves direct vs. nested (inherited) membership path and dynamic vs. assigned (static) rules.
   - Flags role-assignable groups (`isAssignableToRole -eq $true`) to prevent privilege escalation blindspots.
2. **Directory Roles & Privileged Identity Management (PIM):**
   - Discovers directory roles assigned directly to user accounts.
   - Maps directory roles inherited via role-assignable security groups.
   - Identifies active and eligible role assignment schedules in Privileged Identity Management (PIM).
3. **Microsoft Teams & Channel Governance:**
   - Detects all Teams associated with user's M365 groups and resolves Team roles (`Owner` vs. `Member`).
   - Audits private and shared channels where the user holds explicit membership or ownership.
4. **SharePoint Online Sites:**
   - Maps associated SharePoint Online site collection root URLs for M365 groups.
   - Optional deep scan (`-IncludeSharePointSiteAudit`) to evaluate explicit and group-inherited site permissions.
5. **Exchange Online Mailbox Delegation:**
   - Maps distribution lists and dynamic distribution groups.
   - Scans mailbox delegation rights (`FullAccess`, `SendAs`, `SendOnBehalf`) assigned directly or inherited via mail-enabled security groups.
6. **Performance & Throttling Mitigation:**
   - Employs pagination (`-All`) for Graph API requests to prevent the 100-record threshold limitation.
   - Uses thread-safe synchronized memory caching (`Hashtables`) for groups, roles, teams, and sites to minimize API roundtrips and protect against HTTP 429 throttling.
7. **Executive Reporting Dashboard:**
   - Generates interactive, modern HTML5 dashboards with real-time client-side search, KPI summary metric cards, workload filter tabs, and sortable columns.
   - Exports UTF-8 with BOM CSV files (ready for Microsoft Excel) and structured JSON files.
   - Implements standardized logging to `C:\Logs\Get-M365UserAccessAudit\` following enterprise standards.

---

## 2. Prerequisites & Permissions

### Minimum Requirements:
* **PowerShell**: **PowerShell 7.2+** recommended (fully compatible with Windows PowerShell 5.1).
* **Required Modules**:
  ```powershell
  # Core Microsoft Graph modules:
  Install-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Teams, Microsoft.Graph.Sites -Scope CurrentUser

  # Exchange Online Management (required when using -IncludeExchangePermissions):
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```

### Least Privilege Permissions:
* **Microsoft Graph API Scopes (Read-Only):**
  - `User.Read.All`: Discover user profiles and account statuses.
  - `Group.Read.All`: Query group metadata, members, and transitive relationships.
  - `Directory.Read.All`: Discover directory objects, owned objects, and organization metadata.
  - `RoleManagement.Read.Directory`: Inspect Entra ID directory role assignments and definitions.
  - `Team.ReadBasic.All` & `Channel.ReadBasic.All`: Enumerate Teams and inspect private/shared channels.
  - `Sites.Read.All`: Resolve SharePoint Online site URLs and permissions.
  - `RoleEligibilitySchedule.Read.Directory` *(Optional, required for `-IncludePIM`)*: Read Entra ID PIM eligible assignments.
* **Exchange Online Role:**
  - `View-Only Recipients` or `Global Reader`.

---

## 3. Parameters

| Parameter | Type | Required | Default | Description |
| :--- | :---: | :---: | :---: | :--- |
| `-UserPrincipalName` | `string[]` | **Yes** | — | One or more User Principal Names (UPN) or ObjectIds of users to audit. Supports pipeline input (`ValueFromPipeline`). Aliases: `-UPN`, `-User`, `-Identity`, `-UserId`. |
| `-IncludeExchangePermissions` | `switch` | No | `$false` | Scans Exchange Online mailboxes for delegated permissions (`FullAccess`, `SendAs`, `SendOnBehalf`) held directly or via groups. |
| `-IncludeSharePointSiteAudit` | `switch` | No | `$false` | Inspects site collection permissions on SharePoint Online sites associated with user groups. |
| `-IncludePIM` | `switch` | No | `$false` | Audits Entra ID Privileged Identity Management (PIM) eligible role assignment schedules. |
| `-ExportFormat` | `string[]` | No | `@('CSV', 'HTML')` | Formats to generate (`ValidateSet('CSV', 'HTML', 'JSON')`). |
| `-OutputPath` | `string` | No | `C:\Logs\Get-M365UserAccessAudit\Reports` | Destination directory for generated reports. |
| `-LogPath` | `string` | No | `C:\Logs\Get-M365UserAccessAudit` | Directory where execution `.log` files are stored. |
| `-TenantId` | `string` | No | `$null` | Specific Entra ID Tenant ID to target during authentication. |
| `-PassThru` | `switch` | No | `$false` | Emits custom `[PSCustomObject]` records to the pipeline for downstream cmdlets or `Out-GridView`. |
| `-Help` / `-h` | `switch` | No | `$false` | Displays built-in help and usage instructions. |

---

## 4. Usage Examples

### Example 1: Standard Comprehensive Single-User Audit
Audits Entra ID groups, directory roles, Teams, and SharePoint sites for a single user:
```powershell
.\Get-M365UserAccessAudit.ps1 -UserPrincipalName "jan.kowalski@contoso.com"
```

### Example 2: Deep Audit with Exchange Delegations and PIM
Includes Exchange Online mailbox permissions and PIM eligible role assignments:
```powershell
.\Get-M365UserAccessAudit.ps1 -UserPrincipalName "anna.nowak@contoso.com" -IncludeExchangePermissions -IncludePIM
```

### Example 3: Multi-User Pipeline Audit with JSON Export
Processes multiple accounts from a string array or file and exports to CSV, HTML, and JSON:
```powershell
"user1@contoso.com", "user2@contoso.com" | .\Get-M365UserAccessAudit.ps1 -ExportFormat CSV,HTML,JSON
```

### Example 4: Pipeline Pass-Through to `Out-GridView`
Emits raw custom objects directly to PowerShell pipeline:
```powershell
.\Get-M365UserAccessAudit.ps1 -UserPrincipalName "admin@contoso.com" -PassThru | Out-GridView
```

### Example 5: Custom Output and Log Directories
```powershell
.\Get-M365UserAccessAudit.ps1 -UserPrincipalName "ceo@contoso.com" `
    -OutputPath "D:\ComplianceReports\2026-Q4" `
    -LogPath "D:\ComplianceLogs"
```

---

## 5. Report Data Schema (15 Standard Attributes)

Each row in the exported CSV, HTML, and JSON reports represents an individual permission, role, or group membership:

| Column Name | Type | Description | Sample Value |
| :--- | :---: | :--- | :--- |
| `AuditTimestamp` | String | Timestamp when the audit record was captured (`yyyy-MM-dd HH:mm:ss`). | `2026-10-24 08:30:15` |
| `UserPrincipalName` | String | User Principal Name (UPN) of the audited identity. | `jan.kowalski@contoso.com` |
| `UserDisplayName` | String | Display Name of the audited identity. | `Jan Kowalski` |
| `AccountEnabled` | Boolean | Account status in Entra ID (`True` / `False`). | `True` |
| `Workload` | String | Service context (`Entra ID`, `Microsoft Teams`, `SharePoint Online`, `Exchange Online`). | `Entra ID` |
| `ResourceType` | String | Type of audited resource (`Security Group`, `M365 Group`, `Directory Role`, `Team`, `SPO Site`, `Mailbox`, etc.). | `Directory Role` |
| `ResourceName` | String | Human-readable name of the group, role, team, site, or mailbox. | `User Administrator` |
| `ResourceId` | String | Entra ID ObjectId or unique identifier of the resource. | `fe930be7-5e62-47db-91af-98c3a49a38b1` |
| `AccessLevel` | String | Effective access or role granted (`Owner`, `Member`, `Directory Role (Active)`, `FullAccess`, `Site Member`). | `Directory Role (Inherited)` |
| `AssignmentMethod` | String | Access assignment vector (`Direct`, `Nested / Group Inherited`, `Dynamic Rule`, `PIM Eligible`). | `Nested / Group Inherited` |
| `GrantedViaGroup` | String | Name of the mediating group if inherited, otherwise `Direct`. | `Sec-IT-Helpdesk-Admins` |
| `GroupMembershipType` | String | Group assignment mechanism (`Assigned (Static)`, `Dynamic`, `N/A`). | `Assigned (Static)` |
| `IsRoleAssignable` | Boolean | Indicates whether the group can be assigned Entra ID directory roles (`True` / `False`). | `True` |
| `SyncSource` | String | Authority of the resource (`Cloud` or `On-Premises AD`). | `Cloud` |
| `ResourceUrl` | String | Hyperlink to the resource (e.g. SharePoint Online site URL) or empty. | `https://contoso.sharepoint.com/sites/IT` |

---

## 6. HTML Executive Dashboard Highlights

The generated `.html` report contains a self-contained web dashboard with:
- **Executive KPI Cards**: Visual metrics for Total Records, Security Groups, M365 Groups, Directory Roles, Teams, SharePoint Sites, and Exchange Mailboxes.
- **Instant Client-Side Search**: Filters table rows instantaneously across all columns as you type without reloading.
- **Workload Filter Pills**: Toggle between All Workloads, Entra ID, Teams, SharePoint, and Exchange Online.
- **Interactive Column Sorting**: Click any header to sort ascending/descending.
- **Visual Badges**: Color-coded badges for high-risk access (Directory Roles, Mailbox FullAccess, Owners, PIM).
- **Direct Resource Links**: Clickable links to SharePoint Online team sites.

---

## 7. Logging Standard (`Write-Log`)

Execution logs are generated in `C:\Logs\Get-M365UserAccessAudit\` (or `-LogPath`):
- **File Naming**: `Get-M365UserAccessAudit_YYYYMMDD_HHMMSS.log`
- **Format**:
  ```text
  [2026-10-24 08:30:00] [INFO   ] [INIT      ] Starting Get-M365UserAccessAudit.ps1 v1.0.0
  [2026-10-24 08:30:02] [SUCCESS] [AUTH      ] Connected to Microsoft Graph. Tenant ID: 72f988bf-86f1-41af-91ab-2d7cd011db47
  [2026-10-24 08:30:05] [INFO   ] [ENTRA_ID  ] Identified: Jan Kowalski (jan.kowalski@contoso.com)
  [2026-10-24 08:30:08] [WARN   ] [ROLES     ] Inherited Directory Role Found: 'Global Reader' via Group 'Sec-Auditors'
  [2026-10-24 08:30:15] [SUCCESS] [REPORT    ] Interactive HTML Report generated: C:\Logs\Get-M365UserAccessAudit\Reports\Get-M365UserAccessAudit_20261024_083015.html
  ```

---

## 8. Security & Unblocking

After downloading or cloning from GitHub, unblock the script before execution:
```powershell
Unblock-File -Path .\Get-M365UserAccessAudit.ps1
```

---

## 9. Author & License

* **Author:** Roman Pindela
* **Email:** [roman.pindela@gmail.com](mailto:roman.pindela@gmail.com)
* **GitHub:** [https://github.com/romanpindela](https://github.com/romanpindela)
* **Version:** 1.0.0
* **Date:** 2026-10-24
* **License:** MIT License
