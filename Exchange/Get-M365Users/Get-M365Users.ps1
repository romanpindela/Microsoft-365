<#
.SYNOPSIS
    Lists and audits licensed Microsoft 365 users via Exchange Online and Microsoft Graph in a grouped tabular format.

.DESCRIPTION
    This script connects to Exchange Online (prompting for administrator authentication if needed)
    and Microsoft Graph (to retrieve tenant license inventory, consumed/free license counts, and
    assigned license SKUs). Displays all licensed users in a formatted, grouped tabular layout with:
    - Row numbering (Nr) at the beginning of each list
    - Single consolidated email column and full display name (Imie i nazwisko)
    - Assigned licenses column in every view (both Simple and Detailed modes)
    - Executive summary at the bottom with total user count, consumed licenses, and free licenses by type
    - Filtering by partial license name (-LicenseFilter) and user search (-Search)
    - Detailed mode (-Details) with organizational attributes (Department, Title, Location, Archive)
    - CSV export (-ExportCsv) and pipeline pass-through (-PassThru)

.PARAMETER All
    Switch parameter. Lists all licensed users in the tenant.

.PARAMETER LicenseFilter
    Filters users by matching a partial license name or plan (e.g. "Enterprise", "Deskless", "Business", "E3").
    Aliases: -License, -Plan.

.PARAMETER Details
    Switch parameter. Displays extended user attributes in additional columns (Department, Title, Office/City,
    Country, Creation Date, Archive Status).
    Aliases: -d, -Detailed.

.PARAMETER Search
    Optional string filter for matching DisplayName, Email, or Department.
    Aliases: -FilterUser, -User.

.PARAMETER ExportCsv
    Optional file path to export the collected user data into a UTF-8 CSV report.
    Aliases: -CsvPath, -Export.

.PARAMETER NoGrouping
    Switch parameter. Outputs a flat unified table instead of grouping users by license plan.

.PARAMETER PassThru
    Switch parameter. Returns custom PowerShell objects to the pipeline for downstream commands or Out-GridView.

.PARAMETER SkipGraph
    Switch parameter. Bypasses Microsoft Graph connection and audits solely via Exchange Online.
    Alias: -NoGraph.

.PARAMETER AdminUserPrincipalName
    Optional UPN to pre-populate during interactive Connect-ExchangeOnline administrator login.
    Alias: -AdminUPN.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits.

.PARAMETER h
    Alias for -Help.

.EXAMPLE
    .\Get-M365Users.ps1 -All
    Lists all licensed users with assigned licenses, row numbers, single email column, and tenant license summary.

.EXAMPLE
    .\Get-M365Users.ps1 -All -Details
    Lists all licensed users with full detailed columns (Department, Title, Office, Archive) and license summary.

.EXAMPLE
    .\Get-M365Users.ps1 -LicenseFilter "Business"
    Filters users who hold a Business-tier license plan (e.g. Microsoft 365 Business Premium / Standard).

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
    Version: 1.2.1
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
    [Alias('NoGraph')]
    [switch]$SkipGraph,

    [Parameter(ParameterSetName = 'Default')]
    [Alias('AdminUPN')]
    [string]$AdminUserPrincipalName,

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
        "VERSION: 1.2.1",
        "AUTHOR: Roman Pindela",
        "CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Lists and audits licensed Microsoft 365 users via Exchange Online & Graph.",
        "    Presents users in a structured tabular format with row numbering (Nr),",
        "    single email column, full names, and assigned licenses in every view.",
        "    Displays an executive summary at the bottom with total user count,",
        "    consumed licenses, and free (available) licenses by type.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    Requires 'ExchangeOnlineManagement' and optionally 'Microsoft.Graph.Authentication'.",
        "    Automatically detects active Exchange Online sessions or prompts for sign-in.",
        "",
        "USAGE EXAMPLES:",
        "    # 1. List all licensed users with license inventory summary:",
        "    .\Get-M365Users.ps1 -All",
        "",
        "    # 2. List all licensed users with extended organizational attributes:",
        "    .\Get-M365Users.ps1 -All -Details",
        "",
        "    # 3. Filter users by partial license name (e.g., Enterprise, Business, Kiosk):",
        "    .\Get-M365Users.ps1 -LicenseFilter `"Business`"",
        "",
        "    # 4. Search for a specific user and inspect details:",
        "    .\Get-M365Users.ps1 -Search `"kowalski`" -Details",
        "",
        "    # 5. Export report to CSV file:",
        "    .\Get-M365Users.ps1 -All -ExportCsv `"C:\Reports\LicensedUsers.csv`"",
        "",
        "    # 6. Run strictly via Exchange Online without Microsoft Graph:",
        "    .\Get-M365Users.ps1 -All -SkipGraph",
        "",
        "    # 7. Pipe custom objects to Out-GridView (flat view):",
        "    .\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView",
        "",
        "PARAMETERS:",
        "    -All                    Switch to list all licensed users in the tenant.",
        "    -LicenseFilter, -License  Filter users by partial license or plan name.",
        "    -Details, -d            Display detailed attributes in additional columns.",
        "    -Search, -User          Filter users by DisplayName, Email, or Department.",
        "    -ExportCsv, -Export     File path to export results to CSV (UTF-8).",
        "    -NoGrouping             Output a single flat table instead of grouped sections.",
        "    -PassThru               Emit custom PSObjects to the pipeline.",
        "    -SkipGraph, -NoGraph    Bypass Microsoft Graph connection.",
        "    -AdminUserPrincipalName Administrator UPN for Exchange Online login.",
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
# Friendly M365 SKU Names Dictionary
# ----------------------------------------------------------------------
$knownSkuDictionary = @{
    "ENTERPRISEPACK"            = "Office 365 E3"
    "ENTERPRISEPREMIUM"         = "Office 365 E5"
    "STANDARDPACK"              = "Office 365 E1"
    "SPE_E3"                    = "Microsoft 365 E3"
    "SPE_E5"                    = "Microsoft 365 E5"
    "SPB"                       = "Microsoft 365 Business Premium"
    "O365_BUSINESS_PREMIUM"     = "Microsoft 365 Business Standard"
    "O365_BUSINESS_ESSENTIALS"  = "Microsoft 365 Business Basic"
    "SMB_BUSINESS_PREMIUM"      = "Microsoft 365 Business Premium"
    "SMB_BUSINESS"              = "Microsoft 365 Apps for business"
    "OFFICESUBSCRIPTION"        = "Microsoft 365 Apps for enterprise"
    "EXCHANGEENTERPRISE"        = "Exchange Online Plan 2"
    "EXCHANGESTANDARD"          = "Exchange Online Plan 1"
    "EXCHANGEDESKLESS"          = "Exchange Online Kiosk"
    "TEAMS_EXPLORATORY"         = "Microsoft Teams Exploratory"
    "POWER_BI_STANDARD"         = "Power BI (Free)"
    "POWER_BI_PRO"              = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER" = "Power BI Premium Per User"
    "VISIOCLIENT"               = "Visio Plan 2"
    "PROJECTCLIENT"             = "Project Plan 3"
    "EMS"                       = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                = "Enterprise Mobility + Security E5"
}

# ----------------------------------------------------------------------
# Microsoft Graph Connection & Tenant Subscribed SKUs Inventory
# ----------------------------------------------------------------------
$graphSkusMap = @{}
$graphUsersMap = @{}
$tenantSkuInventory = [System.Collections.Generic.List[PSCustomObject]]::new()
$usingGraph = $false

if (-not $SkipGraph) {
    $graphAuthAvailable = Get-Module -Name Microsoft.Graph.Authentication -ListAvailable
    if ($graphAuthAvailable) {
        try {
            $mgContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $mgContext) {
                Write-Host "[*] Connecting to Microsoft Graph for tenant license inventory (total/free counts)..." -ForegroundColor Cyan
                Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome -ErrorAction Stop
            }
            Write-Host "[+] Microsoft Graph connection established." -ForegroundColor Green
            $usingGraph = $true

            # Fetch Subscribed SKUs via Graph REST
            Write-Host "[*] Retrieving tenant subscription license quotas (SubscribedSkus)..." -ForegroundColor Cyan
            $skusResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -ErrorAction Stop
            
            if ($skusResponse -and $skusResponse.value) {
                foreach ($sku in $skusResponse.value) {
                    $skuId = $sku.skuId
                    $skuPart = $sku.skuPartNumber
                    $friendlyName = if ($knownSkuDictionary.ContainsKey($skuPart)) {
                        $knownSkuDictionary[$skuPart]
                    } else {
                        $skuPart
                    }

                    $graphSkusMap[$skuId] = $friendlyName

                    $prepaid = 0
                    if ($sku.prepaidUnits -and $sku.prepaidUnits.enabled) {
                        $prepaid = [int]$sku.prepaidUnits.enabled
                    }
                    $consumed = [int]$sku.consumedUnits
                    $free = [Math]::Max(0, ($prepaid - $consumed))

                    $skuRow = [PSCustomObject]@{
                        'License Plan / SKU'    = $friendlyName
                        'SKU Part'              = $skuPart
                        'Used (Wykorzystane)'   = $consumed
                        'Free (Wolne)'          = $free
                        'Total (Zakupione)'     = $prepaid
                    }
                    $tenantSkuInventory.Add($skuRow)
                }
                Write-Host "    Found $($tenantSkuInventory.Count) subscription license pool(s)." -ForegroundColor Gray
            }

            # Fetch licensed user details via Graph if Microsoft.Graph.Users is available
            if (Get-Module -Name Microsoft.Graph.Users -ListAvailable) {
                Write-Host "[*] Querying user assigned licenses from Microsoft Graph..." -ForegroundColor Cyan
                $graphUsers = @(Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable count -All -Property "id,displayName,userPrincipalName,mail,accountEnabled,assignedLicenses,department,jobTitle,officeLocation,city,country,usageLocation,createdDateTime" -ErrorAction SilentlyContinue)
                
                foreach ($gu in $graphUsers) {
                    $userLicNames = [System.Collections.Generic.List[string]]::new()
                    if ($gu.AssignedLicenses) {
                        foreach ($lic in @($gu.AssignedLicenses)) {
                            $licId = $lic.SkuId
                            if ($graphSkusMap.ContainsKey($licId)) {
                                $userLicNames.Add($graphSkusMap[$licId])
                            } else {
                                $userLicNames.Add($licId)
                            }
                        }
                    }
                    $graphUsersMap[$gu.UserPrincipalName.ToLowerInvariant()] = @{
                        GraphUser = $gu
                        Licenses  = ($userLicNames -join ", ")
                    }
                }
                Write-Host "    Cached $($graphUsers.Count) user license profile(s) from Graph." -ForegroundColor Gray
            }
        } catch {
            Write-Host "    [i] Microsoft Graph connection skipped or unavailable ($($_)). Proceeding with Exchange Online data." -ForegroundColor Gray
            $usingGraph = $false
        }
    } else {
        Write-Host "    [i] Microsoft.Graph.Authentication module not installed. Proceeding with Exchange Online licensing." -ForegroundColor Gray
    }
}

# ----------------------------------------------------------------------
# Helper: Friendly License Plan Name Resolver (Exchange Online Plans)
# ----------------------------------------------------------------------
function Get-FriendlyLicenseName {
    param(
        [string]$MailboxPlanId,
        [hashtable]$LookupTable
    )

    if ([string]::IsNullOrWhiteSpace($MailboxPlanId)) {
        return "Unassigned / No Mailbox Plan"
    }

    if ($LookupTable.ContainsKey($MailboxPlanId) -and -not [string]::IsNullOrWhiteSpace($LookupTable[$MailboxPlanId])) {
        return $LookupTable[$MailboxPlanId]
    }

    $cleanPlan = ($MailboxPlanId -split '-')[0]

    if ($LookupTable.ContainsKey($cleanPlan) -and -not [string]::IsNullOrWhiteSpace($LookupTable[$cleanPlan])) {
        return $LookupTable[$cleanPlan]
    }

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
    $primarySmtp = Get-ObjectPropertyValue $mbx 'PrimarySmtpAddress' ''
    
    # Consolidate email into a single column
    $email = if (-not [string]::IsNullOrWhiteSpace($primarySmtp) -and $primarySmtp -ne '-') {
        $primarySmtp
    } else {
        $upn
    }
    if ([string]::IsNullOrWhiteSpace($email)) { continue }
    $seenUpns.Add($email) | Out-Null
    if (-not [string]::IsNullOrWhiteSpace($upn)) { $seenUpns.Add($upn) | Out-Null }

    $upnLower = if (-not [string]::IsNullOrWhiteSpace($upn)) { $upn.ToLowerInvariant() } else { $email.ToLowerInvariant() }
    $mailboxPlanRaw = Get-ObjectPropertyValue $mbx 'MailboxPlan' ''
    $friendlyPlan = Get-FriendlyLicenseName -MailboxPlanId $mailboxPlanRaw -LookupTable $planLookup

    # Resolve assigned licenses (from Graph if available, otherwise from Exchange mailbox plan)
    $assignedLicenses = $friendlyPlan
    if ($usingGraph -and $graphUsersMap.ContainsKey($upnLower)) {
        $graphInfo = $graphUsersMap[$upnLower]
        if (-not [string]::IsNullOrWhiteSpace($graphInfo.Licenses)) {
            $assignedLicenses = $graphInfo.Licenses
        }
    }

    # Retrieve Department, Title, Office, City, CountryOrRegion
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
    $recipType   = Get-ObjectPropertyValue $mbx 'RecipientTypeDetails' 'UserMailbox'
    $usageLoc    = Get-ObjectPropertyValue $mbx 'UsageLocation' '-'

    $cleanPlanName = if (-not [string]::IsNullOrWhiteSpace($mailboxPlanRaw)) {
        ($mailboxPlanRaw -split '-')[0]
    } else {
        '-'
    }

    $userObj = [PSCustomObject]@{
        Nr                     = 0
        DisplayName            = $dispName
        Email                  = $email
        UserPrincipalName      = $upn
        License                = $assignedLicenses
        PrimaryLicense         = if (-not [string]::IsNullOrWhiteSpace($assignedLicenses)) { ($assignedLicenses -split ',')[0].Trim() } else { $friendlyPlan }
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

# If Graph is enabled, also add licensed users without an Exchange mailbox (e.g. Teams-only / PowerBI accounts)
if ($usingGraph) {
    foreach ($entry in $graphUsersMap.GetEnumerator()) {
        $gUpn = $entry.Key
        if (-not $seenUpns.Contains($gUpn)) {
            $gu         = $entry.Value.GraphUser
            $gLic       = $entry.Value.Licenses
            $rawCreated = Get-ObjectPropertyValue $gu 'CreatedDateTime' $null
            $gCreated   = if ($null -ne $rawCreated -and $rawCreated -is [datetime]) { $rawCreated.ToString("yyyy-MM-dd HH:mm") } elseif ($null -ne $rawCreated) { [string]$rawCreated } else { "-" }
            $gDispName  = Get-ObjectPropertyValue $gu 'DisplayName' '-'
            $gUpnVal    = Get-ObjectPropertyValue $gu 'UserPrincipalName' '-'
            $gMail      = Get-ObjectPropertyValue $gu 'Mail' '-'
            $gEmail     = if (-not [string]::IsNullOrWhiteSpace($gMail) -and $gMail -ne '-') { $gMail } else { $gUpnVal }
            $gDept      = Get-ObjectPropertyValue $gu 'Department' '-'
            $gTitle     = Get-ObjectPropertyValue $gu 'JobTitle' '-'
            $gOffice    = Get-ObjectPropertyValue $gu 'OfficeLocation' '-'
            $gCity      = Get-ObjectPropertyValue $gu 'City' '-'
            $gCountry   = Get-ObjectPropertyValue $gu 'Country' '-'
            $gUsageLoc  = Get-ObjectPropertyValue $gu 'UsageLocation' '-'

            $licStr = if ($gLic) { $gLic } else { "Assigned License (No Mailbox)" }

            $nonMbxObj = [PSCustomObject]@{
                Nr                     = 0
                DisplayName            = $gDispName
                Email                  = $gEmail
                UserPrincipalName      = $gUpnVal
                License                = $licStr
                PrimaryLicense         = ($licStr -split ',')[0].Trim()
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
            $_.PrimaryLicense -like "*$LicenseFilter*" -or
            $_.MailboxPlan -like "*$LicenseFilter*"
        })
    )
}

if (-not [string]::IsNullOrWhiteSpace($Search)) {
    Write-Host "[*] Applying user search filter: '$Search'..." -ForegroundColor Cyan
    $processedUsers = [System.Collections.Generic.List[PSCustomObject]]::new(
        @($processedUsers | Where-Object {
            $_.DisplayName -like "*$Search*" -or
            $_.Email -like "*$Search*" -or
            $_.UserPrincipalName -like "*$Search*" -or
            $_.Department -like "*$Search*"
        })
    )
}

if ($processedUsers.Count -eq 0) {
    Write-Host "`n[!] No licensed users found matching the specified criteria." -ForegroundColor Yellow
    exit 0
}

# ----------------------------------------------------------------------
# Assign Sequential Row Numbers (Nr)
# ----------------------------------------------------------------------
$userIndex = 1
foreach ($u in $processedUsers) {
    $u.Nr = $userIndex
    $userIndex++
}

Write-Host "`n[+] Found $($processedUsers.Count) licensed user(s) matching criteria.`n" -ForegroundColor Green

# ----------------------------------------------------------------------
# Define Tabular Columns (Simple vs Detailed)
# Single email column, row number (Nr) at start, assigned licenses in all views
# ----------------------------------------------------------------------
if ($Details) {
    $tableColumns = @(
        @{ Label = 'Nr';                Expression = { $_.Nr };                   Width = 4 },
        @{ Label = 'Display Name';      Expression = { $_.DisplayName };          Width = 22 },
        @{ Label = 'Email';             Expression = { $_.Email };                Width = 28 },
        @{ Label = 'Assigned Licenses'; Expression = { $_.License };            Width = 30 },
        @{ Label = 'Department';        Expression = { $_.Department };           Width = 14 },
        @{ Label = 'Job Title';         Expression = { $_.Title };                Width = 18 },
        @{ Label = 'Office / City';     Expression = { if ($_.Office -ne '-' -and $_.City -ne '-') { "$($_.Office) / $($_.City)" } elseif ($_.Office -ne '-') { $_.Office } else { $_.City } }; Width = 16 },
        @{ Label = 'Country';           Expression = { $_.CountryOrRegion };      Width = 8 },
        @{ Label = 'Recipient Type';    Expression = { $_.RecipientTypeDetails }; Width = 14 },
        @{ Label = 'Created Date';      Expression = { $_.WhenCreated };          Width = 16 },
        @{ Label = 'Archive';           Expression = { $_.ArchiveStatus };        Width = 9 }
    )
} else {
    $tableColumns = @(
        @{ Label = 'Nr';                Expression = { $_.Nr };                   Width = 4 },
        @{ Label = 'Display Name';      Expression = { $_.DisplayName };          Width = 25 },
        @{ Label = 'Email';             Expression = { $_.Email };                Width = 32 },
        @{ Label = 'Assigned Licenses'; Expression = { $_.License };            Width = 35 },
        @{ Label = 'Recipient Type';    Expression = { $_.RecipientTypeDetails }; Width = 16 }
    )
}

# ----------------------------------------------------------------------
# Render Tabular Display (Grouped vs Flat)
# ----------------------------------------------------------------------
$groups = $processedUsers | Group-Object -Property PrimaryLicense | Sort-Object Count -Descending

if ($NoGrouping) {
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "  ALL LICENSED USERS (Flat View - $($processedUsers.Count) users)" -ForegroundColor White
    Write-Host "================================================================================" -ForegroundColor Cyan
    $processedUsers | Format-Table -Property $tableColumns -AutoSize | Out-String | Write-Host
} else {
    foreach ($grp in $groups) {
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  LICENSE GROUP: $($grp.Name) ($($grp.Count) users)" -ForegroundColor White
        Write-Host "================================================================================" -ForegroundColor Cyan
        
        $grp.Group | Format-Table -Property $tableColumns -AutoSize | Out-String | Write-Host
    }
}

# ----------------------------------------------------------------------
# Executive Summary Breakdown: Total Users, Free Licenses, Used by Type
# ----------------------------------------------------------------------
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "EXECUTIVE LICENSING SUMMARY" -ForegroundColor White
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ("  Total Licensed Users Found : {0}" -f $processedUsers.Count) -ForegroundColor Green
Write-Host ("  Total License Plan Groups  : {0}" -f $groups.Count) -ForegroundColor Green
Write-Host ""

if ($tenantSkuInventory.Count -gt 0) {
    Write-Host "  TENANT SUBSCRIPTION LICENSE INVENTORY (MICROSOFT 365):" -ForegroundColor White
    
    $inventoryColumns = @(
        @{ Label = 'License Plan / SKU';     Expression = { $_.'License Plan / SKU' };   Width = 35 },
        @{ Label = 'Used (Wykorzystane)';    Expression = { $_.'Used (Wykorzystane)' };  Width = 20; Alignment = 'Right' },
        @{ Label = 'Free (Wolne)';           Expression = { $_.'Free (Wolne)' };         Width = 15; Alignment = 'Right' },
        @{ Label = 'Total (Zakupione)';      Expression = { $_.'Total (Zakupione)' };    Width = 16; Alignment = 'Right' }
    )
    $tenantSkuInventory | Format-Table -Property $inventoryColumns -AutoSize | Out-String | Write-Host

    $totalConsumed = ($tenantSkuInventory | Measure-Object -Property 'Used (Wykorzystane)' -Sum).Sum
    $totalFree     = ($tenantSkuInventory | Measure-Object -Property 'Free (Wolne)' -Sum).Sum
    $totalPurchased= ($tenantSkuInventory | Measure-Object -Property 'Total (Zakupione)' -Sum).Sum

    Write-Host ("  SUBSCRIPTION TOTALS: Used = {0} | Free = {1} | Total = {2}" -f $totalConsumed, $totalFree, $totalPurchased) -ForegroundColor Yellow
} else {
    Write-Host "  EXCHANGE ONLINE LICENSED USERS BY PLAN:" -ForegroundColor White
    
    $localSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($grp in $groups) {
        $localSummary.Add([PSCustomObject]@{
            'License Plan / SKU'          = $grp.Name
            'Used (Wykorzystane)'         = $grp.Count
            'Free (Wolne)'                = '(Requires Graph)'
        })
    }
    
    $localColumns = @(
        @{ Label = 'License Plan / SKU';   Expression = { $_.'License Plan / SKU' };   Width = 45 },
        @{ Label = 'Used (Wykorzystane)';  Expression = { $_.'Used (Wykorzystane)' };  Width = 20; Alignment = 'Right' },
        @{ Label = 'Free (Wolne)';         Expression = { $_.'Free (Wolne)' };         Width = 18; Alignment = 'Right' }
    )
    $localSummary | Format-Table -Property $localColumns -AutoSize | Out-String | Write-Host

    Write-Host "  [i] Notice: Available (Free) license pool counts require Microsoft Graph." -ForegroundColor Gray
    Write-Host "      Run without -SkipGraph with Microsoft.Graph.Authentication to view free quotas." -ForegroundColor Gray
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
