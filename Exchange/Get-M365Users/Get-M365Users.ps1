<#
.SYNOPSIS
    Lists and audits licensed Microsoft 365 users via Exchange Online and Microsoft Graph in a structured, grouped tabular format.

.DESCRIPTION
    This script connects to Exchange Online (prompting for administrator authentication if needed)
    and Microsoft Graph (to retrieve tenant license inventory, consumed/free license quotas, and
    assigned license SKUs). Displays all licensed users in a formatted, grouped tabular layout with:
    - Row numbering (Nr) at the beginning of each list
    - Consolidated single email column and full display name (Imie i nazwisko)
    - Intelligent license separation: Base Commercial Plan vs. Add-ons & Free services
    - Clear user grouping by primary commercial plan (prevents scattering across add-ons)
    - Executive Licensing Summary with paid commercial subscriptions separated from free pools
    - Real-time utilization percentage and capacity status tags ([OK], [LOW], [FULL])
    - Filtering by partial license name (-LicenseFilter), user search (-Search), and summary-only mode (-LicenseSummary)
    - Detailed mode (-Details) with organizational attributes (Department, Title, Location, Creation Date, Archive)
    - CSV export (-ExportCsv) and pipeline pass-through (-PassThru)

.PARAMETER All
    Switch parameter. Lists all licensed users in the tenant.

.PARAMETER LicenseFilter
    Filters users by matching a partial license name or plan (e.g. "Business", "Enterprise", "Standard", "Defender").
    Aliases: -License, -Plan.

.PARAMETER Details
    Switch parameter. Displays extended user attributes in additional columns (Department, Title, Location,
    Recipient Type, Creation Date, Archive Status).
    Aliases: -d, -Detailed.

.PARAMETER Search
    Optional string filter for matching DisplayName, Email, Department, or Title.
    Aliases: -FilterUser, -User.

.PARAMETER LicenseSummary
    Switch parameter. Displays only the Executive Licensing Summary tables (quotas, free/used counts,
    and user distribution) without printing the individual user lists.
    Aliases: -SummaryOnly, -Quota, -Summary.

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
    Lists all licensed users grouped by their base license plan with assigned licenses, row numbers,
    single email column, and tenant license quota summary.

.EXAMPLE
    .\Get-M365Users.ps1 -All -Details
    Lists all licensed users with extended organizational columns (Department, Title, Location, Archive)
    and tenant license quota summary.

.EXAMPLE
    .\Get-M365Users.ps1 -LicenseSummary
    Displays only the executive licensing summary tables (commercial subscriptions, free services,
    and user distribution) without listing individual users.

.EXAMPLE
    .\Get-M365Users.ps1 -LicenseFilter "Business"
    Filters users holding a Business-tier plan (e.g. Microsoft 365 Business Standard / Premium).

.EXAMPLE
    .\Get-M365Users.ps1 -Search "kowalski" -Details
    Searches for users matching "kowalski" and displays detailed organizational information.

.EXAMPLE
    .\Get-M365Users.ps1 -All -ExportCsv "C:\Reports\LicensedUsers.csv"
    Audits all licensed users and saves the comprehensive report to a UTF-8 CSV file.

.EXAMPLE
    .\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView
    Lists all licensed users in a flat table and pipes custom objects to Out-GridView.

.NOTES
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Version: 1.3.0
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
    [Alias('SummaryOnly', 'Quota', 'Summary')]
    [switch]$LicenseSummary,

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
        "VERSION: 1.3.0",
        "AUTHOR: Roman Pindela",
        "CONTACT: roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Lists and audits licensed Microsoft 365 users via Exchange Online & Graph.",
        "    Presents users in a structured tabular format with row numbering (Nr),",
        "    single email column, full names, and assigned licenses in every view.",
        "    Intelligently classifies Base Commercial Plans vs Add-ons to avoid scattering",
        "    users into disjoint groups. Displays an Executive Licensing Summary at the",
        "    bottom with commercial vs free pools, free quotas, and utilization alerts.",
        "",
        "AUTHENTICATION & PREREQUISITES:",
        "    Requires 'ExchangeOnlineManagement' and optionally 'Microsoft.Graph.Authentication'.",
        "    Automatically detects active Exchange Online sessions or prompts for sign-in.",
        "",
        "USAGE EXAMPLES:",
        "    # 1. List all licensed users with executive license summary:",
        "    .\Get-M365Users.ps1 -All",
        "",
        "    # 2. List all licensed users with extended organizational attributes:",
        "    .\Get-M365Users.ps1 -All -Details",
        "",
        "    # 3. Display only the executive license quota summary (no user list):",
        "    .\Get-M365Users.ps1 -LicenseSummary",
        "",
        "    # 4. Filter users by partial license name (e.g., Business, Enterprise, Defender):",
        "    .\Get-M365Users.ps1 -LicenseFilter `"Business`"",
        "",
        "    # 5. Search for a specific user and inspect details:",
        "    .\Get-M365Users.ps1 -Search `"kowalski`" -Details",
        "",
        "    # 6. Export report to CSV file:",
        "    .\Get-M365Users.ps1 -All -ExportCsv `"C:\Reports\LicensedUsers.csv`"",
        "",
        "    # 7. Run strictly via Exchange Online without Microsoft Graph:",
        "    .\Get-M365Users.ps1 -All -SkipGraph",
        "",
        "    # 8. Pipe custom objects to Out-GridView (flat view):",
        "    .\Get-M365Users.ps1 -All -NoGrouping -PassThru | Out-GridView",
        "",
        "PARAMETERS:",
        "    -All                    Switch to list all licensed users in the tenant.",
        "    -LicenseFilter, -License  Filter users by partial license or plan name.",
        "    -Details, -d            Display detailed attributes in additional columns.",
        "    -Search, -User          Filter users by DisplayName, Email, or Department.",
        "    -LicenseSummary, -Quota Display only the executive licensing summary tables.",
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

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName) -and $null -ne $InputObject[$PropertyName]) {
            $val = $InputObject[$PropertyName]
            if ($val -is [string] -and [string]::IsNullOrWhiteSpace($val)) { return $DefaultValue }
            return $val
        }
        return $DefaultValue
    }

    $prop = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $prop -and $null -ne $prop.Value) {
        $val = $prop.Value
        if ($val -is [string] -and [string]::IsNullOrWhiteSpace($val)) { return $DefaultValue }
        return $val
    }
    return $DefaultValue
}

# ----------------------------------------------------------------------
# Helper: Safe Date Formatter (Prevents Console Width Truncation)
# ----------------------------------------------------------------------
function Format-DateValue {
    param([object]$DateValue)
    if ($null -eq $DateValue -or $DateValue -eq '-' -or [string]::IsNullOrWhiteSpace([string]$DateValue)) {
        return '-'
    }
    if ($DateValue -is [datetime]) {
        return $DateValue.ToString("yyyy-MM-dd")
    }
    $parsedDate = [datetime]::MinValue
    if ([datetime]::TryParse([string]$DateValue, [ref]$parsedDate)) {
        return $parsedDate.ToString("yyyy-MM-dd")
    }
    return [string]$DateValue
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
    # Microsoft 365 Commercial Suites
    "O365_BUSINESS_PREMIUM"     = "Microsoft 365 Business Standard"
    "SMB_BUSINESS_PREMIUM"      = "Microsoft 365 Business Premium"
    "SPB"                       = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"  = "Microsoft 365 Business Basic"
    "SMB_BUSINESS"              = "Microsoft 365 Apps for business"
    "OFFICESUBSCRIPTION"        = "Microsoft 365 Apps for enterprise"
    "ENTERPRISEPACK"            = "Office 365 E3"
    "ENTERPRISEPREMIUM"         = "Office 365 E5"
    "STANDARDPACK"              = "Office 365 E1"
    "SPE_E3"                    = "Microsoft 365 E3"
    "SPE_E5"                    = "Microsoft 365 E5"
    "SPE_F1"                    = "Microsoft 365 F1"
    "DESKLESSPACK"              = "Office 365 F3"
    "TEAMS_EXPLORATORY"         = "Microsoft Teams Exploratory"
    "TEAMS_ESSENTIALS"          = "Microsoft Teams Essentials"

    # Exchange Online Standalone Plans
    "EXCHANGEENTERPRISE"        = "Exchange Online Plan 2"
    "EXCHANGESTANDARD"          = "Exchange Online Plan 1"
    "EXCHANGEDESKLESS"          = "Exchange Online Kiosk"
    "EXCHANGEARCHIVE"           = "Exchange Online Archiving"
    "EXCHANGEARCHIVE_ADDON"     = "Exchange Online Archiving"

    # Security & Compliance Add-ons
    "ATP_ENTERPRISE"            = "Defender for Office 365 (Plan 1)"
    "THREAT_INTELLIGENCE"       = "Defender for Office 365 (Plan 2)"
    "CCIBEAGLE"                 = "Microsoft Defender for Cloud Apps"
    "EMS"                       = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                = "Enterprise Mobility + Security E5"
    "AAD_PREMIUM"               = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"            = "Microsoft Entra ID P2"
    "INTUNE_A"                  = "Microsoft Intune Plan 1"

    # Power Platform & Productivity
    "POWER_BI_STANDARD"         = "Power BI (Free)"
    "POWER_BI_PRO"              = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER" = "Power BI Premium Per User"
    "FLOW_FREE"                 = "Power Automate (Free)"
    "POWERAPPS_VIRAL"           = "Power Apps (Free)"
    "VISIOCLIENT"               = "Visio Plan 2"
    "PROJECTCLIENT"             = "Project Plan 3"

    # Other Free / Add-on Pools
    "RIGHTSMANAGEMENT_ADHOC"    = "Azure Rights Management (Free)"
    "WINDOWS_STORE"             = "Windows Store for Business"
    "MCOMEETADV"                = "Teams Audio Conferencing"
    "PHONERF"                   = "Teams Phone Standard"
    "COMMUNICATION_CREDITS"     = "Communication Credits"
}

# ----------------------------------------------------------------------
# Helper: SKU Friendly Name & Rank Resolvers
# ----------------------------------------------------------------------
function Get-FriendlySkuName {
    param([string]$SkuIdentifier)
    if ([string]::IsNullOrWhiteSpace($SkuIdentifier)) { return "-" }
    $clean = $SkuIdentifier.Trim()
    if ($knownSkuDictionary.ContainsKey($clean)) {
        return $knownSkuDictionary[$clean]
    }
    return $clean
}

function Get-SkuRank {
    param([string]$SkuName)
    # Higher rank indicates a primary base commercial subscription plan
    if ($SkuName -match "Business Premium|Business Standard|Business Basic|Apps for|Office 365 E[135]|Microsoft 365 E[35]|Microsoft 365 F[13]|Exchange Online Plan [12]|Exchange Online Kiosk|Teams Exploratory") {
        return 100
    }
    if ($SkuName -match "Defender|Archiving|Entra ID|Intune|Power BI Pro|Power BI Premium|Visio|Project|Security|Mobility") {
        return 50
    }
    if ($SkuName -match "Free|AdHoc|Viral|Windows Store") {
        return 10
    }
    return 30
}

function Get-SkuClassification {
    param(
        [string]$SkuPartNumber,
        [string]$FriendlyName
    )
    if ($SkuPartNumber -in @('FLOW_FREE', 'POWER_BI_STANDARD', 'RIGHTSMANAGEMENT_ADHOC', 'WINDOWS_STORE', 'POWERAPPS_VIRAL') -or
        $FriendlyName -match '\(Free\)' -or $FriendlyName -match 'Free|AdHoc|Viral') {
        return @{
            Category = 'Free / Complimentary'
            Priority = 10
            IsPaid   = $false
        }
    }
    if ($SkuPartNumber -in @('ATP_ENTERPRISE', 'THREAT_INTELLIGENCE', 'CCIBEAGLE', 'EMS', 'EMSPREMIUM', 'AAD_PREMIUM', 'AAD_PREMIUM_P2', 'INTUNE_A', 'POWER_BI_PRO', 'POWER_BI_PREMIUM_PER_USER', 'VISIOCLIENT', 'PROJECTCLIENT', 'EXCHANGEARCHIVE', 'EXCHANGEARCHIVE_ADDON', 'MCOMEETADV', 'PHONERF') -or
        $FriendlyName -match 'Defender|Archiving|Mobility|Entra ID|Intune|Visio|Project|Conferencing') {
        return @{
            Category = 'Paid Commercial Add-on'
            Priority = 50
            IsPaid   = $true
        }
    }
    return @{
        Category = 'Commercial Suite / Core Plan'
        Priority = 100
        IsPaid   = $true
    }
}

function Split-UserLicenses {
    param([string]$AssignedLicensesString)
    if ([string]::IsNullOrWhiteSpace($AssignedLicensesString)) {
        return @{ BaseLicense = "-"; AddOns = "-" }
    }
    $rawParts = @($AssignedLicensesString -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $friendlyParts = @($rawParts | ForEach-Object { Get-FriendlySkuName $_ })
    if ($friendlyParts.Count -eq 0) {
        return @{ BaseLicense = "-"; AddOns = "-" }
    }
    $ranked = @($friendlyParts | ForEach-Object {
        [PSCustomObject]@{
            Name = $_
            Rank = Get-SkuRank $_
        }
    } | Sort-Object Rank -Descending)

    $basePlan = $ranked[0].Name
    $addOnList = @($ranked | Select-Object -Skip 1 | ForEach-Object { $_.Name })
    $addOnStr = if ($addOnList.Count -gt 0) { $addOnList -join ", " } else { "-" }

    return @{
        BaseLicense = $basePlan
        AddOns      = $addOnStr
    }
}

# ----------------------------------------------------------------------
# Microsoft Graph Connection & Tenant Subscribed SKUs Inventory
# ----------------------------------------------------------------------
$graphSkusMap = @{}
$graphUsersMap = @{}
$tenantCommercialSkus = [System.Collections.Generic.List[PSCustomObject]]::new()
$tenantFreeSkus       = [System.Collections.Generic.List[PSCustomObject]]::new()
$usingGraph = $false

if (-not $SkipGraph) {
    $graphAuthAvailable = Get-Module -Name Microsoft.Graph.Authentication -ListAvailable
    if ($graphAuthAvailable) {
        try {
            $mgContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $mgContext) {
                Write-Host "[*] Connecting to Microsoft Graph for tenant license inventory (total/free quotas)..." -ForegroundColor Cyan
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
                    $friendlyName = Get-FriendlySkuName $skuPart

                    $graphSkusMap[$skuId] = $friendlyName

                    $prepaid = 0
                    if ($sku.prepaidUnits -and $sku.prepaidUnits.enabled) {
                        $prepaid = [int]$sku.prepaidUnits.enabled
                    }
                    $consumed = [int]$sku.consumedUnits
                    $free = [Math]::Max(0, ($prepaid - $consumed))

                    $classification = Get-SkuClassification -SkuPartNumber $skuPart -FriendlyName $friendlyName

                    $utilizationPct = if ($prepaid -gt 0) {
                        ([double]$consumed / [double]$prepaid * 100).ToString("N1") + "%"
                    } else {
                        "N/A"
                    }

                    $statusTag = if ($prepaid -eq 0) {
                        "[Unlimited / N/A]"
                    } elseif ($free -eq 0) {
                        "[FULL - 0 FREE]"
                    } elseif ($free -le 2) {
                        "[LOW - $free FREE]"
                    } else {
                        "[OK]"
                    }

                    $rowObj = [PSCustomObject]@{
                        Nr          = 0
                        Plan        = $friendlyName
                        SkuPart     = $skuPart
                        Used        = $consumed
                        Free        = $free
                        Total       = $prepaid
                        Utilization = $utilizationPct
                        Status      = $statusTag
                        Category    = $classification.Category
                        IsPaid      = $classification.IsPaid
                    }

                    if ($classification.IsPaid) {
                        $tenantCommercialSkus.Add($rowObj)
                    } else {
                        $rowObj.Status = "[Available]"
                        $tenantFreeSkus.Add($rowObj)
                    }
                }
                Write-Host "    Found $($tenantCommercialSkus.Count) commercial and $($tenantFreeSkus.Count) free subscription pool(s)." -ForegroundColor Gray
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
        "ExchangeOnlineEnterprise"  { return "Exchange Online Plan 2" }
        "ExchangeOnlineDeskless"    { return "Exchange Online Kiosk" }
        "ExchangeOnlineEssentials"  { return "Exchange Online Plan 1" }
        "ExchangeOnline"            { return "Exchange Online Plan 1" }
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

    # Classify Base License vs Add-ons
    $licClassification = Split-UserLicenses -AssignedLicensesString $assignedLicenses

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

    # Format creation date cleanly (yyyy-MM-dd)
    $createdRaw = Get-ObjectPropertyValue $mbx 'WhenMailboxCreated' $null
    if ($null -eq $createdRaw -or $createdRaw -eq '-') {
        $createdRaw = Get-ObjectPropertyValue $mbx 'WhenCreated' $null
    }
    $createdDate = Format-DateValue $createdRaw

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

    # Compact compound values for detailed display
    $deptTitle = if ($dept -ne '-' -and $title -ne '-') {
        "$dept / $title"
    } elseif ($dept -ne '-') {
        $dept
    } elseif ($title -ne '-') {
        $title
    } else {
        '-'
    }

    $location = if ($city -ne '-' -and $country -ne '-') {
        "$city / $country"
    } elseif ($office -ne '-' -and $country -ne '-') {
        "$office / $country"
    } elseif ($country -ne '-') {
        $country
    } elseif ($city -ne '-') {
        $city
    } elseif ($office -ne '-') {
        $office
    } else {
        '-'
    }

    $userObj = [PSCustomObject]@{
        Nr                     = 0
        DisplayName            = $dispName
        Email                  = $email
        UserPrincipalName      = $upn
        BaseLicense            = $licClassification.BaseLicense
        AddOns                 = $licClassification.AddOns
        License                = $assignedLicenses
        PrimaryLicense         = $licClassification.BaseLicense
        MailboxPlan            = $cleanPlanName
        RecipientTypeDetails   = $recipType
        Department             = $dept
        Title                  = $title
        DeptTitle              = $deptTitle
        Office                 = $office
        City                   = $city
        CountryOrRegion        = $country
        Location               = $location
        UsageLocation          = $usageLoc
        WhenCreated            = $createdDate
        ArchiveStatus          = $archiveVal
        HiddenFromAddressLists = $hiddenVal
    }

    $processedUsers.Add($userObj)
}

# If Graph is enabled, also add licensed users without an Exchange mailbox (e.g. Teams-only / Entra accounts)
if ($usingGraph) {
    foreach ($entry in $graphUsersMap.GetEnumerator()) {
        $gUpn = $entry.Key
        if (-not $seenUpns.Contains($gUpn)) {
            $gu         = $entry.Value.GraphUser
            $gLic       = $entry.Value.Licenses
            $gCreated   = Format-DateValue (Get-ObjectPropertyValue $gu 'CreatedDateTime' $null)
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
            $gLicClassification = Split-UserLicenses -AssignedLicensesString $licStr

            $gDeptTitle = if ($gDept -ne '-' -and $gTitle -ne '-') {
                "$gDept / $gTitle"
            } elseif ($gDept -ne '-') {
                $gDept
            } elseif ($gTitle -ne '-') {
                $gTitle
            } else {
                '-'
            }

            $gLocation = if ($gCity -ne '-' -and $gCountry -ne '-') {
                "$gCity / $gCountry"
            } elseif ($gOffice -ne '-' -and $gCountry -ne '-') {
                "$gOffice / $gCountry"
            } elseif ($gCountry -ne '-') {
                $gCountry
            } elseif ($gCity -ne '-') {
                $gCity
            } else {
                '-'
            }

            $nonMbxObj = [PSCustomObject]@{
                Nr                     = 0
                DisplayName            = $gDispName
                Email                  = $gEmail
                UserPrincipalName      = $gUpnVal
                BaseLicense            = $gLicClassification.BaseLicense
                AddOns                 = $gLicClassification.AddOns
                License                = $licStr
                PrimaryLicense         = $gLicClassification.BaseLicense
                MailboxPlan            = "NoExchangeMailbox"
                RecipientTypeDetails   = "EntraUser (No Mailbox)"
                Department             = $gDept
                Title                  = $gTitle
                DeptTitle              = $gDeptTitle
                Office                 = $gOffice
                City                   = $gCity
                CountryOrRegion        = $gCountry
                Location               = $gLocation
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
            $_.BaseLicense -like "*$LicenseFilter*" -or
            $_.AddOns -like "*$LicenseFilter*" -or
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
            $_.Email -like "*$Search*" -or
            $_.UserPrincipalName -like "*$Search*" -or
            $_.Department -like "*$Search*" -or
            $_.Title -like "*$Search*"
        })
    )
}

if ($processedUsers.Count -eq 0) {
    Write-Host "`n[!] No licensed users found matching the specified criteria." -ForegroundColor Yellow
    exit 0
}

# ----------------------------------------------------------------------
# Grouping & Sequential Numbering
# ----------------------------------------------------------------------
# Order groups by user count descending, then sort users by DisplayName within each group
$groups = @($processedUsers | Group-Object -Property BaseLicense | Sort-Object Count -Descending)

$globalIndex = 1
foreach ($grp in $groups) {
    $sortedMembers = @($grp.Group | Sort-Object -Property DisplayName)
    foreach ($member in $sortedMembers) {
        $member.Nr = $globalIndex
        $globalIndex++
    }
}

Write-Host "`n[+] Found $($processedUsers.Count) licensed user(s) matching criteria.`n" -ForegroundColor Green

# ----------------------------------------------------------------------
# Render User Tables (Grouped vs Flat)
# Skipped if -LicenseSummary is specified
# ----------------------------------------------------------------------
if (-not $LicenseSummary) {
    if ($Details) {
        $tableColumns = @(
            @{ Label = 'Nr';              Expression = { $_.Nr };          Width = 4 },
            @{ Label = 'Display Name';    Expression = { $_.DisplayName }; Width = 22 },
            @{ Label = 'Email';           Expression = { $_.Email };       Width = 32 },
            @{ Label = 'Base License';    Expression = { $_.BaseLicense }; Width = 32 },
            @{ Label = 'Add-on Licenses'; Expression = { $_.AddOns };      Width = 24 },
            @{ Label = 'Dept / Title';    Expression = { $_.DeptTitle };   Width = 20 },
            @{ Label = 'Location';        Expression = { $_.Location };    Width = 15 },
            @{ Label = 'Type';            Expression = { $_.RecipientTypeDetails }; Width = 12 },
            @{ Label = 'Created';         Expression = { $_.WhenCreated }; Width = 11 },
            @{ Label = 'Archive';         Expression = { $_.ArchiveStatus }; Width = 7 }
        )
    } else {
        $tableColumns = @(
            @{ Label = 'Nr';              Expression = { $_.Nr };          Width = 4 },
            @{ Label = 'Display Name';    Expression = { $_.DisplayName }; Width = 25 },
            @{ Label = 'Email';           Expression = { $_.Email };       Width = 34 },
            @{ Label = 'Base License';    Expression = { $_.BaseLicense }; Width = 32 },
            @{ Label = 'Add-on Licenses'; Expression = { $_.AddOns };      Width = 26 },
            @{ Label = 'Department';      Expression = { $_.Department };  Width = 20 }
        )
    }

    $renderWidth = 205

    if ($NoGrouping) {
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "  ALL LICENSED USERS (Flat View - $($processedUsers.Count) users)" -ForegroundColor White
        Write-Host "================================================================================" -ForegroundColor Cyan
        $flatUsers = @($processedUsers | Sort-Object Nr)
        $flatUsers | Format-Table -Property $tableColumns | Out-String -Width $renderWidth | Write-Host
    } else {
        foreach ($grp in $groups) {
            Write-Host "================================================================================" -ForegroundColor Cyan
            Write-Host "  LICENSE GROUP: $($grp.Name) ($($grp.Count) users)" -ForegroundColor White
            Write-Host "================================================================================" -ForegroundColor Cyan
            
            $grpUsers = @($grp.Group | Sort-Object Nr)
            $grpUsers | Format-Table -Property $tableColumns | Out-String -Width $renderWidth | Write-Host
        }
    }
}

# ----------------------------------------------------------------------
# Executive Licensing Summary: Commercial vs Free Quotas & User Breakdown
# ----------------------------------------------------------------------
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "EXECUTIVE LICENSING SUMMARY (PODSUMOWANIE LICENCJI)" -ForegroundColor White
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ("  Total Licensed Users Found : {0}" -f $processedUsers.Count) -ForegroundColor Green
Write-Host ("  Primary Base License Plans : {0}" -f $groups.Count) -ForegroundColor Green
if ($usingGraph) {
    Write-Host ("  Total Subscription Pools   : {0}" -f ($tenantCommercialSkus.Count + $tenantFreeSkus.Count)) -ForegroundColor Green
}
Write-Host ""

if ($usingGraph -and ($tenantCommercialSkus.Count -gt 0 -or $tenantFreeSkus.Count -gt 0)) {
    # 1. Commercial Subscriptions
    if ($tenantCommercialSkus.Count -gt 0) {
        Write-Host "[1] COMMERCIAL SUBSCRIPTIONS (PLATNE LICENCJE I DODATKI):" -ForegroundColor White
        
        $commIdx = 1
        foreach ($c in $tenantCommercialSkus) {
            $c.Nr = $commIdx
            $commIdx++
        }

        $commCols = @(
            @{ Label = 'Nr';                  Expression = { $_.Nr };          Width = 4 },
            @{ Label = 'License / Plan Name'; Expression = { $_.Plan };        Width = 34 },
            @{ Label = 'Used';                Expression = { $_.Used };        Width = 8; Alignment = 'Right' },
            @{ Label = 'Free';                Expression = { $_.Free };        Width = 8; Alignment = 'Right' },
            @{ Label = 'Total';               Expression = { $_.Total };       Width = 8; Alignment = 'Right' },
            @{ Label = 'Util %';              Expression = { $_.Utilization }; Width = 8; Alignment = 'Right' },
            @{ Label = 'Status';              Expression = { $_.Status };      Width = 18 }
        )
        $tenantCommercialSkus | Format-Table -Property $commCols | Out-String -Width 120 | Write-Host

        $totalCommUsed = ($tenantCommercialSkus | Measure-Object -Property 'Used' -Sum).Sum
        $totalCommFree = ($tenantCommercialSkus | Measure-Object -Property 'Free' -Sum).Sum
        $totalCommPrepaid = ($tenantCommercialSkus | Measure-Object -Property 'Total' -Sum).Sum
        $totalCommUtil = if ($totalCommPrepaid -gt 0) {
            ([double]$totalCommUsed / [double]$totalCommPrepaid * 100).ToString("N1") + "%"
        } else {
            "N/A"
        }

        Write-Host ("    COMMERCIAL TOTALS: Used = {0} | Free = {1} | Total = {2} ({3} Utilized)" -f $totalCommUsed, $totalCommFree, $totalCommPrepaid, $totalCommUtil) -ForegroundColor Yellow
        Write-Host ""
    }

    # 2. Complimentary & Free Pools
    if ($tenantFreeSkus.Count -gt 0) {
        Write-Host "[2] COMPLIMENTARY & FREE CLOUD SERVICES (BEZPLATNE USLUGI W CHMURZE):" -ForegroundColor White
        
        $freeIdx = 1
        foreach ($f in $tenantFreeSkus) {
            $f.Nr = $freeIdx
            $freeIdx++
        }

        $freeCols = @(
            @{ Label = 'Nr';                  Expression = { $_.Nr };          Width = 4 },
            @{ Label = 'Service / Pool Name'; Expression = { $_.Plan };        Width = 34 },
            @{ Label = 'Used';                Expression = { $_.Used };        Width = 8; Alignment = 'Right' },
            @{ Label = 'Free';                Expression = { if ($_.Free -gt 0) { ($_.Free).ToString("N0") } else { "-" } }; Width = 10; Alignment = 'Right' },
            @{ Label = 'Total';               Expression = { if ($_.Total -gt 0) { ($_.Total).ToString("N0") } else { "-" } }; Width = 11; Alignment = 'Right' },
            @{ Label = 'Status';              Expression = { $_.Status };      Width = 14 }
        )
        $tenantFreeSkus | Format-Table -Property $freeCols | Out-String -Width 120 | Write-Host
        Write-Host ""
    }

    # 3. User Count Breakdown by Base License Plan
    Write-Host "[3] USER COUNT BY PRIMARY BASE LICENSE (LICZBA UZYTKOWNIKOW WG LICENCJI BAZOWEJ):" -ForegroundColor White
    $userDistribution = [System.Collections.Generic.List[PSCustomObject]]::new()
    $distIdx = 1
    $totalUsersCount = $processedUsers.Count
    foreach ($grp in $groups) {
        $pct = if ($totalUsersCount -gt 0) {
            ([double]$grp.Count / [double]$totalUsersCount * 100).ToString("N1") + "%"
        } else {
            "0.0%"
        }
        $userDistribution.Add([PSCustomObject]@{
            Nr    = $distIdx
            Plan  = $grp.Name
            Users = $grp.Count
            Share = $pct
        })
        $distIdx++
    }

    $distCols = @(
        @{ Label = 'Nr';                        Expression = { $_.Nr };    Width = 4 },
        @{ Label = 'Primary Base License Plan'; Expression = { $_.Plan };  Width = 34 },
        @{ Label = 'User Count';                Expression = { $_.Users }; Width = 11; Alignment = 'Right' },
        @{ Label = 'Share %';                   Expression = { $_.Share }; Width = 9;  Alignment = 'Right' }
    )
    $userDistribution | Format-Table -Property $distCols | Out-String -Width 100 | Write-Host
    Write-Host ("    TOTAL LICENSED USERS: {0} (100.0%)" -f $totalUsersCount) -ForegroundColor Green

} else {
    Write-Host "  EXCHANGE ONLINE LICENSED USERS BY PLAN:" -ForegroundColor White
    
    $localSummary = [System.Collections.Generic.List[PSCustomObject]]::new()
    $localIdx = 1
    foreach ($grp in $groups) {
        $localSummary.Add([PSCustomObject]@{
            Nr                            = $localIdx
            'License Plan / SKU'          = $grp.Name
            'Used (Wykorzystane)'         = $grp.Count
            'Free (Wolne)'                = '(Requires Graph)'
        })
        $localIdx++
    }
    
    $localColumns = @(
        @{ Label = 'Nr';                   Expression = { $_.Nr };                     Width = 4 },
        @{ Label = 'License Plan / SKU';   Expression = { $_.'License Plan / SKU' };   Width = 40 },
        @{ Label = 'Used (Wykorzystane)';  Expression = { $_.'Used (Wykorzystane)' };  Width = 20; Alignment = 'Right' },
        @{ Label = 'Free (Wolne)';         Expression = { $_.'Free (Wolne)' };         Width = 18; Alignment = 'Right' }
    )
    $localSummary | Format-Table -Property $localColumns | Out-String -Width 100 | Write-Host

    Write-Host "  [i] Notice: Available (Free) license pool quotas require Microsoft Graph." -ForegroundColor Gray
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
