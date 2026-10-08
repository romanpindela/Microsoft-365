<#
.SYNOPSIS
    Lists and audits licensed Microsoft 365 users via Exchange Online in a grouped tabular format.

.DESCRIPTION
    This script connects to Exchange Online (prompting for administrator authentication if needed),
    retrieves all licensed mailboxes and users, maps their mailbox plans and license assignments to
    friendly names, and presents them in a formatted, grouped tabular layout.
    
    Supports:
    - Displaying help and usage examples when executed without parameters
    - Filtering by partial license or plan name (e.g. -LicenseFilter "Enterprise", "Deskless", "Basic")
    - Simple tabular output vs. comprehensive Detailed mode (-Details) with extra user attributes
    - Exporting audit reports to CSV (-ExportCsv)
    - Pass-through pipeline support (-PassThru)
    - Optional Microsoft Graph integration (-UseGraph) for tenant-wide Entra ID license SKUs

.PARAMETER All
    Switch parameter. Lists all licensed users in the tenant.

.PARAMETER LicenseFilter
    Filters users by matching a partial license name or MailboxPlan (e.g. "Enterprise", "Deskless", "Essentials", "E3").
    Aliases: -License, -Plan.

.PARAMETER Details
    Switch parameter. Displays extended user attributes in additional columns (Department, Title, Office,
    City, Country, Creation Date, Archive Status, Hidden from Address Lists).
    Aliases: -d, -Detailed.

.PARAMETER Search
    Optional string filter for matching DisplayName, UserPrincipalName, or PrimarySmtpAddress.
    Aliases: -FilterUser, -User.

.PARAMETER ExportCsv
    Optional file path to export the collected user data into a UTF-8 CSV report.
    Aliases: -CsvPath, -Export.

.PARAMETER NoGrouping
    Switch parameter. Outputs a flat unified table instead of grouping users by license plan.

.PARAMETER PassThru
    Switch parameter. Returns custom PowerShell objects to the pipeline for downstream commands or Out-GridView.

.PARAMETER AdminUserPrincipalName
    Optional UPN to pre-populate during interactive Connect-ExchangeOnline administrator login.
    Alias: -AdminUPN.

.PARAMETER UseGraph
    Switch parameter. Queries Microsoft Graph (if available) for tenant-wide Entra ID license SKUs and non-mailbox accounts.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Get-M365Users.ps1 -All
    Lists all licensed users grouped by their assigned license/mailbox plan in simple tabular view.

.EXAMPLE
    .\Get-M365Users.ps1 -All -Details
    Lists all licensed users grouped by license with detailed columns (Department, Job Title, Office, Country, etc.).

.EXAMPLE
    .\Get-M365Users.ps1 -LicenseFilter "Enterprise"
    Lists users who hold an Enterprise-tier license plan (e.g. Exchange Online Plan 2 / E3 / E5).

.EXAMPLE
    .\Get-M365Users.ps1 -LicenseFilter "Deskless" -Details
    Lists all Kiosk/Deskless users (F1/F3) with full detailed user attributes.

.EXAMPLE
    .\Get-M365Users.ps1 -Search "smith" -Details
    Searches for users matching "smith" and displays their detailed licensing and profile information.

.EXAMPLE
    .\Get-M365Users.ps1 -All -ExportCsv "C:\Reports\LicensedUsers.csv"
    Audits all licensed users and saves the tabular report to a CSV file.

.EXAMPLE
    .\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView
    Lists all licensed users in a flat table and passes objects to Out-GridView.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName = 'Default')]
    [switch]$All,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('License', 'Plan')]
    [string]$LicenseFilter,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('d', 'Detailed')]
    [switch]$Details,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('FilterUser', 'User')]
    [string]$Search,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('CsvPath', 'Export')]
    [string]$ExportCsv,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$NoGrouping,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$PassThru,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

    [Parameter(ParameterSetName = 'Default')]
    [switch]$UseGraph,

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT: Get-M365Users.ps1",
        "VERSION: 1.0.0",
        "AUTHOR: Roman Pindela",
        "CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Lists and audits licensed Microsoft 365 users via Exchange Online.",
        "    Displays users in a formatted, grouped tabular layout categorized by assigned",
        "    license plan. Supports filtering by partial license name, detailed attribute views,",
        "    CSV export, and pipeline pass-through.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    Requires the 'ExchangeOnlineManagement' PowerShell module.",
        "    Automatically detects active Exchange Online sessions or prompts for administrator sign-in.",
        "    Optionally supports Microsoft Graph (-UseGraph) for tenant-wide Entra ID license SKUs.",
        "",
        "USAGE EXAMPLES:",
        "    # 1. List all licensed users grouped by license in simple tabular view:",
        "    .\Get-M365Users.ps1 -All",
        "",
        "    # 2. List all licensed users with detailed user attributes (Department, Title, etc.):",
        "    .\Get-M365Users.ps1 -All -Details",
        "",
        "    # 3. Filter users by partial license name (e.g., Enterprise, Deskless, Essentials):",
        "    .\Get-M365Users.ps1 -LicenseFilter `"Enterprise`"",
        "",
        "    # 4. Filter by license name and display detailed attributes:",
        "    .\Get-M365Users.ps1 -LicenseFilter `"Deskless`" -Details",
        "",
        "    # 5. Search for a specific user and inspect details:",
        "    .\Get-M365Users.ps1 -Search `"smith`" -Details",
        "",
        "    # 6. Export report to CSV file:",
        "    .\Get-M365Users.ps1 -All -ExportCsv `"C:\Reports\LicensedUsers.csv`"",
        "",
        "    # 7. Pipe custom objects to Out-GridView (flat view):",
        "    .\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView",
        "",
        "PARAMETERS:",
        "    -All                    Switch to list all licensed users in the tenant.",
        "    -LicenseFilter, -License  Filter users by partial license or plan name.",
        "    -Details, -d            Display detailed attributes in additional columns.",
        "    -Search, -User          Filter users by DisplayName, UPN, or email.",
        "    -ExportCsv, -Export     File path to export results to CSV (UTF-8).",
        "    -NoGrouping             Output a single flat table instead of grouped sections.",
        "    -PassThru               Emit custom PSObjects to the pipeline.",
        "    -AdminUserPrincipalName Administrator UPN for Exchange Online login.",
        "    -UseGraph               Query Microsoft Graph for tenant-wide license SKUs.",
        "    -Help, -h               Display this help documentation.",
        "",
        "SECURITY & UNBLOCKING:",
        "    After downloading from GitHub, remember to unblock the script before execution:",
        "    Unblock-File -Path .\Get-M365Users.ps1",
        "================================================================================"
    )
    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# ----------------------------------------------------------------------
# If launched without parameters or with -Help, show help and exit
# ----------------------------------------------------------------------
if ($Help -or $PSBoundParameters.Count -eq 0) {
    Show-ScriptHelp
    exit 0
}

# ----------------------------------------------------------------------
# Input Validation & Sanitization (CleanCode & Security Protection)
# ----------------------------------------------------------------------
$forbiddenCharsRegex = '[\$`;&|><\r\n]'

if (-not [string]::IsNullOrWhiteSpace($LicenseFilter)) {
    $LicenseFilter = $LicenseFilter.Trim()
    if ($LicenseFilter -match $forbiddenCharsRegex) {
        Write-Error "Invalid characters detected in -LicenseFilter. Disallowed characters: $;&|><`$`r`n"
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    $Search = $Search.Trim()
    if ($Search -match $forbiddenCharsRegex) {
        Write-Error "Invalid characters detected in -Search. Disallowed characters: $;&|><`$`r`n"
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
    $AdminUserPrincipalName = $AdminUserPrincipalName.Trim()
    $emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    if ($AdminUserPrincipalName -notmatch $emailPattern) {
        Write-Error "Invalid format for -AdminUserPrincipalName. Provide a valid email address or UPN."
        exit 1
    }
}

if (-not [string]::IsNullOrWhiteSpace($ExportCsv)) {
    $ExportCsv = $ExportCsv.Trim().Trim('"').Trim("'")
    if ($ExportCsv -match '[\$`;&|><\r\n]') {
        Write-Error "Invalid characters detected in -ExportCsv path."
        exit 1
    }
    if ($ExportCsv -notmatch '\.csv$') {
        Write-Error "-ExportCsv must specify a valid file path ending in .csv."
        exit 1
    }
}

# ----------------------------------------------------------------------
# Microsoft 365 Exchange Online Authentication
# ----------------------------------------------------------------------
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Error "ExchangeOnlineManagement module is not installed.`nRun: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
    exit 1
}

try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host "[+] Active Exchange Online session detected." -ForegroundColor Green
} catch {
    Write-Host "[*] No active Exchange Online session detected. Initiating administrator login..." -ForegroundColor Cyan
    try {
        $connectParams = @{
            ShowBanner  = $false
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($AdminUserPrincipalName)) {
            $connectParams['UserPrincipalName'] = $AdminUserPrincipalName
        }
        Connect-ExchangeOnline @connectParams
        Write-Host "[+] Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

# ----------------------------------------------------------------------
# Optional Microsoft Graph Handler (if -UseGraph is requested)
# ----------------------------------------------------------------------
$graphUsersMap = @{}
$graphSkusMap = @{}
$usingGraph = $false

if ($UseGraph) {
    Write-Host "[*] Checking Microsoft Graph module availability..." -ForegroundColor Cyan
    $graphModuleAvailable = (Get-Module -Name Microsoft.Graph.Users -ListAvailable) -and (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)
    
    if (-not $graphModuleAvailable) {
        Write-Warning "Microsoft.Graph module is not installed. Falling back to Exchange Online native plan mapping.`nTo enable Graph: Install-Module Microsoft.Graph.Users, Microsoft.Graph.Authentication -Scope CurrentUser"
    } else {
        try {
            $mgContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $mgContext) {
                Write-Host "[*] Connecting to Microsoft Graph..." -ForegroundColor Cyan
                Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome -ErrorAction Stop
            }
            Write-Host "[+] Microsoft Graph connection established." -ForegroundColor Green

            # Fetch Subscribed SKUs
            $subscribedSkus = @(Get-MgSubscribedSku -All -ErrorAction SilentlyContinue)
            foreach ($sku in $subscribedSkus) {
                $graphSkusMap[$sku.SkuId] = $sku.SkuPartNumber
            }

            # Fetch licensed Graph users
            Write-Host "[*] Querying licensed users via Microsoft Graph..." -ForegroundColor Cyan
            $graphUsers = @(Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable count -All -Property "id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,department,jobTitle,officeLocation,city,country,usageLocation,createdDateTime" -ErrorAction Stop)
            
            foreach ($gu in $graphUsers) {
                $userLicNames = @()
                if ($gu.AssignedLicenses) {
                    foreach ($lic in @($gu.AssignedLicenses)) {
                        if ($graphSkusMap.ContainsKey($lic.SkuId)) {
                            $userLicNames += $graphSkusMap[$lic.SkuId]
                        } else {
                            $userLicNames += $lic.SkuId
                        }
                    }
                }
                $graphUsersMap[$gu.UserPrincipalName.ToLowerInvariant()] = @{
                    GraphUser = $gu
                    Licenses  = ($userLicNames -join ", ")
                }
            }
            $usingGraph = $true
            Write-Host "    Found $($graphUsers.Count) licensed users in Microsoft Graph." -ForegroundColor Gray
        } catch {
            Write-Warning "Failed to query Microsoft Graph: $_. Falling back to Exchange Online native plan mapping."
        }
    }
}

# ----------------------------------------------------------------------
# Helper: Safe Object Property Value Resolver (Set-StrictMode Safe)
# ----------------------------------------------------------------------
function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = '-'
    )
    if ($null -eq $InputObject) { return $DefaultValue }
    
    $prop = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $prop -and $null -ne $prop.Value) {
        $val = [string]$prop.Value
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            return $val
        }
    }
    return $DefaultValue
}

# ----------------------------------------------------------------------
# Helper: Friendly License Plan Name Resolver
# ----------------------------------------------------------------------
function Get-FriendlyLicenseName {
    param(
        [string]$MailboxPlanId,
        [hashtable]$LookupTable
    )

    if ([string]::IsNullOrWhiteSpace($MailboxPlanId)) {
        return "Unassigned / No Mailbox Plan"
    }

    # 1. Exact match in Get-MailboxPlan lookup
    if ($LookupTable.ContainsKey($MailboxPlanId) -and -not [string]::IsNullOrWhiteSpace($LookupTable[$MailboxPlanId])) {
        return $LookupTable[$MailboxPlanId]
    }

    # 2. Extract base plan name (before GUID or hyphenated hash)
    $cleanPlan = ($MailboxPlanId -split '-')[0]

    # Check lookup by clean plan name
    if ($LookupTable.ContainsKey($cleanPlan) -and -not [string]::IsNullOrWhiteSpace($LookupTable[$cleanPlan])) {
        return $LookupTable[$cleanPlan]
    }

    # 3. Known standard Microsoft 365 Exchange Online Mailbox Plan mappings
    switch -Regex ($cleanPlan) {
        "ExchangeOnlineEnterprise"  { return "Exchange Online Plan 2 (Enterprise / E3 / E5)" }
        "ExchangeOnlineDeskless"    { return "Exchange Online Kiosk (Deskless / F1 / F3)" }
        "ExchangeOnlineEssentials"  { return "Exchange Online Plan 1 (Essentials / Business)" }
        "ExchangeOnline"            { return "Exchange Online Plan 1 (Standard / E1)" }
        "ExchangeOnlineDevice"      { return "Exchange Online Device" }
        "ExchangeOnlineArchive"     { return "Exchange Online Archiving" }
        default                     { return $cleanPlan }
    }
}

# ----------------------------------------------------------------------
# Retrieve Mailbox Plans for friendly names
# ----------------------------------------------------------------------
Write-Host "`n[*] Retrieving tenant mailbox plans and licensing profiles..." -ForegroundColor Cyan
$planLookup = @{}
try {
    $allPlans = @(Get-MailboxPlan -ResultSize Unlimited -ErrorAction SilentlyContinue)
    foreach ($plan in $allPlans) {
        $planIdentity = Get-ObjectPropertyValue $plan 'Identity' ''
        $planName = Get-ObjectPropertyValue $plan 'Name' ''
        $planDisplayName = Get-ObjectPropertyValue $plan 'DisplayName' ''
        if (-not [string]::IsNullOrWhiteSpace($planIdentity)) { $planLookup[$planIdentity] = $planDisplayName }
        if (-not [string]::IsNullOrWhiteSpace($planName)) { $planLookup[$planName] = $planDisplayName }
    }
} catch {
    Write-Host "    [i] Unable to query Get-MailboxPlan directly; falling back to heuristic plan mapping." -ForegroundColor Gray
}

# ----------------------------------------------------------------------
# Retrieve Organizational Recipient Details (Department, Title, etc.)
# ----------------------------------------------------------------------
$recipientLookup = @{}
if ($Details -or (-not [string]::IsNullOrWhiteSpace($Search))) {
    Write-Host "[*] Retrieving organizational user details (Department, Title, Location)..." -ForegroundColor Cyan
    try {
        $recipients = @()
        if (Get-Command -Name Get-EXORecipient -ErrorAction SilentlyContinue) {
            $recipients = @(Get-EXORecipient -ResultSize Unlimited -Properties Department, Title, Office, City, CountryOrRegion -ErrorAction Stop)
        } else {
            $recipients = @(Get-Recipient -ResultSize Unlimited -ErrorAction Stop)
        }
        foreach ($r in $recipients) {
            $rUpn = Get-ObjectPropertyValue $r 'UserPrincipalName' ''
            $rEmail = Get-ObjectPropertyValue $r 'PrimarySmtpAddress' ''
            if (-not [string]::IsNullOrWhiteSpace($rUpn)) {
                $recipientLookup[$rUpn.ToLowerInvariant()] = $r
            }
            if (-not [string]::IsNullOrWhiteSpace($rEmail)) {
                $recipientLookup[$rEmail.ToLowerInvariant()] = $r
            }
        }
        Write-Host "    Found $($recipients.Count) recipient profile(s)." -ForegroundColor Gray
    } catch {
        Write-Host "    [i] Unable to query Get-EXORecipient ($($_)). Trying Get-User..." -ForegroundColor Gray
        try {
            $users = @(Get-User -ResultSize Unlimited -ErrorAction Stop)
            foreach ($u in $users) {
                $uUpn = Get-ObjectPropertyValue $u 'UserPrincipalName' ''
                $uEmail = Get-ObjectPropertyValue $u 'WindowsEmailAddress' ''
                if (-not [string]::IsNullOrWhiteSpace($uUpn)) {
                    $recipientLookup[$uUpn.ToLowerInvariant()] = $u
                }
                if (-not [string]::IsNullOrWhiteSpace($uEmail)) {
                    $recipientLookup[$uEmail.ToLowerInvariant()] = $u
                }
            }
        } catch {
            Write-Host "    [i] Recipient organizational details not available." -ForegroundColor Gray
        }
    }
}

# ----------------------------------------------------------------------
# Retrieve Mailboxes & Licensed Users from Exchange Online
# ----------------------------------------------------------------------
Write-Host "[*] Querying Microsoft 365 Exchange Online mailboxes..." -ForegroundColor Cyan

# Valid properties supported by Get-EXOMailbox
$validExoProps = @(
    'MailboxPlan',
    'ArchiveStatus',
    'HiddenFromAddressListsEnabled',
    'WhenMailboxCreated',
    'UsageLocation'
)

$rawMailboxes = @()
try {
    if (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue) {
        $rawMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $validExoProps -ErrorAction Stop)
    } else {
        $rawMailboxes = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
    }
} catch {
    Write-Host "    [!] Get-EXOMailbox failed ($($_)). Retrying with standard Get-Mailbox..." -ForegroundColor Yellow
    $rawMailboxes = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
}

# Filter mailboxes that have an active license assigned
$licensedMailboxes = @($rawMailboxes | Where-Object {
    $sku = Get-ObjectPropertyValue $_ 'SKUAssigned' $null
    $plan = Get-ObjectPropertyValue $_ 'MailboxPlan' ''
    $recType = Get-ObjectPropertyValue $_ 'RecipientTypeDetails' ''

    ($sku -eq $true) -or (-not [string]::IsNullOrWhiteSpace($plan) -and $recType -eq 'UserMailbox')
})

Write-Host "    Found $($licensedMailboxes.Count) licensed mailbox(es) in Exchange Online." -ForegroundColor Gray

# ----------------------------------------------------------------------
# Transform and Normalize Data into Structured Objects
# ----------------------------------------------------------------------
$processedUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($mbx in $licensedMailboxes) {
    $upn = Get-ObjectPropertyValue $mbx 'UserPrincipalName' ''
    if ([string]::IsNullOrWhiteSpace($upn)) {
        $upn = Get-ObjectPropertyValue $mbx 'PrimarySmtpAddress' ''
    }
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }
    $seenUpns.Add($upn) | Out-Null

    $upnLower = $upn.ToLowerInvariant()
    $mailboxPlanRaw = Get-ObjectPropertyValue $mbx 'MailboxPlan' ''
    $friendlyLicense = Get-FriendlyLicenseName -MailboxPlanId $mailboxPlanRaw -LookupTable $planLookup

    # If Graph data is available, enrich or use Graph SKU details
    if ($usingGraph -and $graphUsersMap.ContainsKey($upnLower)) {
        $graphInfo = $graphUsersMap[$upnLower]
        if (-not [string]::IsNullOrWhiteSpace($graphInfo.Licenses)) {
            $friendlyLicense = "$friendlyLicense ($($graphInfo.Licenses))"
        }
    }

    # Retrieve Department, Title, Office, City, CountryOrRegion from recipient lookup or mailbox
    $dept    = '-'
    $title   = '-'
    $office  = '-'
    $city    = '-'
    $country = '-'

    if ($recipientLookup.ContainsKey($upnLower)) {
        $recip   = $recipientLookup[$upnLower]
        $dept    = Get-ObjectPropertyValue $recip 'Department' '-'
        $title   = Get-ObjectPropertyValue $recip 'Title' '-'
        $office  = Get-ObjectPropertyValue $recip 'Office' '-'
        $city    = Get-ObjectPropertyValue $recip 'City' '-'
        $country = Get-ObjectPropertyValue $recip 'CountryOrRegion' '-'
    } else {
        $dept    = Get-ObjectPropertyValue $mbx 'Department' '-'
        $title   = Get-ObjectPropertyValue $mbx 'Title' '-'
        $office  = Get-ObjectPropertyValue $mbx 'Office' '-'
        $city    = Get-ObjectPropertyValue $mbx 'City' '-'
        $country = Get-ObjectPropertyValue $mbx 'CountryOrRegion' '-'
    }

    # Creation date formatting
    $createdRaw = Get-ObjectPropertyValue $mbx 'WhenMailboxCreated' $null
    if ($null -eq $createdRaw -or $createdRaw -eq '-') {
        $createdRaw = Get-ObjectPropertyValue $mbx 'WhenCreated' $null
    }
    $createdDate = '-'
    if ($null -ne $createdRaw -and $createdRaw -is [datetime]) {
        $createdDate = $createdRaw.ToString("yyyy-MM-dd HH:mm")
    } elseif ($null -ne $createdRaw -and $createdRaw -ne '-') {
        $createdDate = [string]$createdRaw
    }

    $archiveVal  = Get-ObjectPropertyValue $mbx 'ArchiveStatus' 'None'
    $hiddenVal   = Get-ObjectPropertyValue $mbx 'HiddenFromAddressListsEnabled' 'False'
    $dispName    = Get-ObjectPropertyValue $mbx 'DisplayName' '-'
    $primarySmtp = Get-ObjectPropertyValue $mbx 'PrimarySmtpAddress' '-'
    $recipType   = Get-ObjectPropertyValue $mbx 'RecipientTypeDetails' 'UserMailbox'
    $usageLoc    = Get-ObjectPropertyValue $mbx 'UsageLocation' '-'

    $cleanPlanName = if (-not [string]::IsNullOrWhiteSpace($mailboxPlanRaw)) {
        ($mailboxPlanRaw -split '-')[0]
    } else {
        '-'
    }

    $userObj = [PSCustomObject]@{
        DisplayName            = $dispName
        UserPrincipalName      = $upn
        PrimarySmtpAddress     = $primarySmtp
        License                = $friendlyLicense
        MailboxPlan            = $cleanPlanName
        RecipientTypeDetails   = $recipType
        Department             = $dept
        Title                  = $title
        Office                 = $office
        City                   = $city
        CountryOrRegion        = $country
        UsageLocation          = $usageLoc
        WhenCreated            = $createdDate
        ArchiveStatus          = $archiveVal
        HiddenFromAddressLists = $hiddenVal
    }

    $processedUsers.Add($userObj)
}

# If Graph is enabled, also add any licensed users who do not have an Exchange mailbox (e.g. Teams-only or PowerBI accounts)
if ($usingGraph) {
    foreach ($entry in $graphUsersMap.GetEnumerator()) {
        $gUpn = $entry.Key
        if (-not $seenUpns.Contains($gUpn)) {
            $gu         = $entry.Value.GraphUser
            $gLic       = $entry.Value.Licenses
            $rawCreated = Get-ObjectPropertyValue $gu 'CreatedDateTime' $null
            $gCreated   = if ($null -ne $rawCreated -and $rawCreated -is [datetime]) { $rawCreated.ToString("yyyy-MM-dd HH:mm") } elseif ($null -ne $rawCreated) { [string]$rawCreated } else { "-" }
            $gDispName = Get-ObjectPropertyValue $gu 'DisplayName' '-'
            $gUpnVal   = Get-ObjectPropertyValue $gu 'UserPrincipalName' '-'
            $gMail     = Get-ObjectPropertyValue $gu 'Mail' '-'
            $gDept     = Get-ObjectPropertyValue $gu 'Department' '-'
            $gTitle    = Get-ObjectPropertyValue $gu 'JobTitle' '-'
            $gOffice   = Get-ObjectPropertyValue $gu 'OfficeLocation' '-'
            $gCity     = Get-ObjectPropertyValue $gu 'City' '-'
            $gCountry  = Get-ObjectPropertyValue $gu 'Country' '-'
            $gUsageLoc = Get-ObjectPropertyValue $gu 'UsageLocation' '-'

            $nonMbxObj = [PSCustomObject]@{
                DisplayName            = $gDispName
                UserPrincipalName      = $gUpnVal
                PrimarySmtpAddress     = $gMail
                License                = if ($gLic) { $gLic } else { "Assigned License (No Mailbox)" }
                MailboxPlan            = "NoExchangeMailbox"
                RecipientTypeDetails   = "EntraUser (No Mailbox)"
                Department             = $gDept
                Title                  = $gTitle
                Office                 = $gOffice
                City                   = $gCity
                CountryOrRegion        = $gCountry
                UsageLocation          = $gUsageLoc
                WhenCreated            = $gCreated
                ArchiveStatus          = "N/A"
                HiddenFromAddressLists = "-"
            }
            $processedUsers.Add($nonMbxObj)
        }
    }
}

# ----------------------------------------------------------------------
# Apply Filters (-LicenseFilter, -Search)
# ----------------------------------------------------------------------
if (-not [string]::IsNullOrWhiteSpace($LicenseFilter)) {
    Write-Host "[*] Applying license filter: '$LicenseFilter'..." -ForegroundColor Cyan
    $processedUsers = [System.Collections.Generic.List[PSCustomObject]]::new(
        @($processedUsers | Where-Object {
            $_.License -like "*$LicenseFilter*" -or
            $_.MailboxPlan -like "*$LicenseFilter*"
        })
    )
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    Write-Host "[*] Applying user search filter: '$Search'..." -ForegroundColor Cyan
    $processedUsers = [System.Collections.Generic.List[PSCustomObject]]::new(
        @($processedUsers | Where-Object {
            $_.DisplayName -like "*$Search*" -or
            $_.UserPrincipalName -like "*$Search*" -or
            $_.PrimarySmtpAddress -like "*$Search*" -or
            $_.Department -like "*$Search*"
        })
    )
}

# ----------------------------------------------------------------------
# Output Results
# ----------------------------------------------------------------------
if ($processedUsers.Count -eq 0) {
    Write-Host "`n[!] No licensed users found matching the specified criteria." -ForegroundColor Yellow
    exit 0
}

Write-Host "`n[+] Found $($processedUsers.Count) licensed user(s) matching criteria.`n" -ForegroundColor Green

# Define Tabular Columns (Simple vs Detailed)
if ($Details) {
    $tableColumns = @(
        @{ Label = 'Display Name';        Expression = { $_.DisplayName };          Width = 24 },
        @{ Label = 'User Principal Name'; Expression = { $_.UserPrincipalName };    Width = 30 },
        @{ Label = 'Department';          Expression = { $_.Department };           Width = 16 },
        @{ Label = 'Job Title';           Expression = { $_.Title };                Width = 20 },
        @{ Label = 'Office / City';       Expression = { if ($_.Office -ne '-' -and $_.City -ne '-') { "$($_.Office) / $($_.City)" } elseif ($_.Office -ne '-') { $_.Office } else { $_.City } }; Width = 18 },
        @{ Label = 'Country';             Expression = { $_.CountryOrRegion };      Width = 10 },
        @{ Label = 'Recipient Type';      Expression = { $_.RecipientTypeDetails }; Width = 16 },
        @{ Label = 'Created Date';        Expression = { $_.WhenCreated };          Width = 17 },
        @{ Label = 'Archive';             Expression = { $_.ArchiveStatus };        Width = 10 }
    )
    if ($NoGrouping) {
        $tableColumns += @{ Label = 'License Plan'; Expression = { $_.License }; Width = 35 }
    }
} else {
    $tableColumns = @(
        @{ Label = 'Display Name';        Expression = { $_.DisplayName };          Width = 25 },
        @{ Label = 'User Principal Name'; Expression = { $_.UserPrincipalName };    Width = 32 },
        @{ Label = 'Primary SMTP';        Expression = { $_.PrimarySmtpAddress };   Width = 32 },
        @{ Label = 'Recipient Type';      Expression = { $_.RecipientTypeDetails }; Width = 16 }
    )
    if ($NoGrouping) {
        $tableColumns += @{ Label = 'Assigned License'; Expression = { $_.License }; Width = 35 }
    }
}

# Render Tabular Display
$groups = $processedUsers | Group-Object -Property License | Sort-Object Count -Descending

if ($NoGrouping) {
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "  ALL LICENSED USERS (Flat View - $($processedUsers.Count) users)" -ForegroundColor White
    Write-Host "================================================================================" -ForegroundColor Cyan
    $processedUsers | Format-Table -Property $tableColumns -AutoSize | Out-String | Write-Host
} else {
    foreach ($grp in $groups) {
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  LICENSE PLAN: $($grp.Name) ($($grp.Count) users)" -ForegroundColor White
        Write-Host "================================================================================" -ForegroundColor Cyan
        
        $grp.Group | Format-Table -Property $tableColumns -AutoSize | Out-String | Write-Host
    }
}

# ----------------------------------------------------------------------
# Executive Summary Breakdown
# ----------------------------------------------------------------------
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "EXECUTIVE LICENSING SUMMARY:" -ForegroundColor White
Write-Host ("  Total Licensed Users Found : {0}" -f $processedUsers.Count) -ForegroundColor Green
Write-Host ("  Total License Plan Groups  : {0}" -f $groups.Count) -ForegroundColor Green
Write-Host "  Breakdown by License Group :" -ForegroundColor White
foreach ($grp in $groups) {
    Write-Host ("    - {0,-55} : {1,4} user(s)" -f $grp.Name, $grp.Count) -ForegroundColor Yellow
}
Write-Host "================================================================================" -ForegroundColor Cyan

# ----------------------------------------------------------------------
# Export to CSV (if requested)
# ----------------------------------------------------------------------
if (-not [string]::IsNullOrWhiteSpace($ExportCsv)) {
    try {
        $exportDir = Split-Path -Path $ExportCsv -Parent
        if (-not [string]::IsNullOrWhiteSpace($exportDir) -and -not (Test-Path -Path $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }
        $processedUsers | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "`n[+] Successfully exported $($processedUsers.Count) user records to CSV:" -ForegroundColor Green
        Write-Host "    $ExportCsv" -ForegroundColor White
    } catch {
        Write-Error "Failed to export data to CSV: $_"
    }
}

# ----------------------------------------------------------------------
# PassThru Pipeline Output (if requested)
# ----------------------------------------------------------------------
if ($PassThru) {
    return $processedUsers
}
