#Requires -Version 5.1

<#
.SYNOPSIS
    Manages user licenses in Microsoft Entra ID / Microsoft 365 and generates SKU utilization reports via Microsoft Graph API.

.DESCRIPTION
    Manage-EntraUserLicenses.ps1 is an administrative PowerShell tool built for the Microsoft Graph PowerShell SDK.
    It supports both single-user license operations (inspecting assigned licenses and service plans, assigning new licenses
    with optional service plan exclusions, removing direct licenses) and tenant-wide reporting (listing available SKUs
    and auditing users holding a specific license with quota summaries).

    The script automatically translates friendly Microsoft technical SKU names (SkuPartNumber, e.g., SPE_E3, ENTERPRISEPACK)
    into internal GUIDs (SkuId), verifies available license pool capacity, validates the required UsageLocation attribute,
    detects Group-Based Licensing inheritance, and logs execution details to C:\Logs\.

.PARAMETER UserPrincipalName
    UserPrincipalName (e.g., john.doe@domain.com) or ObjectId of the user in Microsoft Entra ID.
    Aliases: UPN, Identity.

.PARAMETER ShowServicePlans
    Switch parameter (GetUserLicenses mode). Expands the list of all component service plans (ServicePlans)
    for each assigned license along with their provisioning status (ProvisioningStatus).

.PARAMETER Assign
    Switch parameter that activates license assignment mode (AssignLicense).

.PARAMETER Remove
    Switch parameter that activates license removal mode (RemoveLicense).

.PARAMETER Report
    Switch parameter that activates SKU user reporting mode (ReportBySku).

.PARAMETER ListAvailableSkus
    Switch parameter that displays all subscriptions (SKUs) available in the tenant along with unit balances.

.PARAMETER SkuPartNumber
    Standard Microsoft license technical name (e.g., SPE_E3, SPE_E5, ENTERPRISEPACK, EXCHANGESTANDARD, POWER_BI_STANDARD).

.PARAMETER DisabledPlans
    Optional list of service plan names (ServicePlanName, e.g., YAMMER_ENTERPRISE, SWAY) to disable
    when assigning the specified license in -Assign mode.

.PARAMETER ExportCsvPath
    Optional path to a .csv file where the user list generated in -Report mode will be exported.

.PARAMETER TenantId
    Optional Entra ID tenant identifier (GUID or domain name).

.PARAMETER ClientId
    Optional application identifier (App Registration) for non-interactive App-Only authentication.

.PARAMETER CertificateThumbprint
    Optional X.509 certificate thumbprint from Cert:\CurrentUser\My or Cert:\LocalMachine\My (for App-Only authentication).

.PARAMETER Help
    Displays interactive help, mode descriptions, and usage examples, then exits. Alias: -h.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -UserPrincipalName "john.doe@domain.com"
    Displays all licenses assigned to the specified user along with assignment type (Direct / Group).

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -UserPrincipalName "john.doe@domain.com" -ShowServicePlans
    Displays the user's licenses and a detailed list of all component service plans and their statuses.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -UserPrincipalName "john.doe@domain.com" -Assign -SkuPartNumber "SPE_E3" -DisabledPlans @("YAMMER_ENTERPRISE", "SWAY") -WhatIf
    Simulates assigning the Microsoft 365 E3 (SPE_E3) license with Yammer and Sway disabled without making changes in the tenant.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -UserPrincipalName "john.doe@domain.com" -Remove -SkuPartNumber "SPE_E3" -Confirm:$false
    Removes the directly assigned SPE_E3 license from the specified user without an interactive confirmation prompt.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -Report -SkuPartNumber "ENTERPRISEPACK" -ExportCsvPath "C:\Temp\Report_E3.csv"
    Generates a tenant license pool summary for ENTERPRISEPACK, lists all assigned users, and exports the results to CSV.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -ListAvailableSkus
    Displays a table of all available SKU subscriptions in the tenant with enabled, consumed, and available unit counts.

.EXAMPLE
    .\Manage-EntraUserLicenses.ps1 -ListAvailableSkus -TenantId "00000000-0000-0000-0000-000000000000" -ClientId "11111111-1111-1111-1111-111111111111" -CertificateThumbprint "AABBCCDD..."
    Runs the SKU inventory in non-interactive App-Only automation mode using a certificate.

.NOTES
    Author  : Roman Pindela
    Email   : roman.pindela@gmail.com
    GitHub  : https://github.com/romanpindela
    Version : 1.0.0
    Requires: PowerShell 5.1 / 7.0+, Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Identity.DirectoryManagement
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Help')]
param(
    [Parameter(ParameterSetName = 'Help', Mandatory = $false)]
    [Alias('h')]
    [switch]$Help,

    [Parameter(ParameterSetName = 'GetUserLicenses', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'AssignLicense', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'RemoveLicense', Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [Alias('UPN', 'Identity')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'GetUserLicenses', Mandatory = $false)]
    [switch]$ShowServicePlans,

    [Parameter(ParameterSetName = 'AssignLicense', Mandatory = $true)]
    [switch]$Assign,

    [Parameter(ParameterSetName = 'RemoveLicense', Mandatory = $true)]
    [switch]$Remove,

    [Parameter(ParameterSetName = 'ReportBySku', Mandatory = $true)]
    [switch]$Report,

    [Parameter(ParameterSetName = 'ListTenantSkus', Mandatory = $true)]
    [switch]$ListAvailableSkus,

    [Parameter(ParameterSetName = 'AssignLicense', Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveLicense', Mandatory = $true)]
    [Parameter(ParameterSetName = 'ReportBySku', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SkuPartNumber,

    [Parameter(ParameterSetName = 'AssignLicense', Mandatory = $false)]
    [string[]]$DisabledPlans,

    [Parameter(ParameterSetName = 'ReportBySku', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExportCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint
)

#region Configuration & Logging

$script:ScriptMetadata = @{
    Title   = 'Manage-EntraUserLicenses.ps1'
    Version = '1.0.0'
    Author  = 'Roman Pindela'
    Email   = 'roman.pindela@gmail.com'
    GitHub  = 'https://github.com/romanpindela'
}

$script:LogDirectory = 'C:\Logs'
$script:Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:ComputerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { [Environment]::MachineName }
$script:UserName = if ($env:USERNAME) { $env:USERNAME } else { [Environment]::UserName }
$script:LogFilePath = Join-Path -Path $script:LogDirectory -ChildPath ("{0}-Manage-EntraUserLicenses-{1}-{2}.logs" -f $script:Timestamp, $script:ComputerName, $script:UserName)

$script:RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users'
)

function Show-ScriptHelp {
    <#
    .SYNOPSIS
        Displays readable help, operating modes, and usage examples for the script.
    #>
    [CmdletBinding()]
    param()

    Write-Host ''
    Write-Host '================================================================================' -ForegroundColor Cyan
    Write-Host ("  SCRIPT:  {0}" -f $script:ScriptMetadata.Title) -ForegroundColor Cyan
    Write-Host ("  VERSION: {0}" -f $script:ScriptMetadata.Version) -ForegroundColor Cyan
    Write-Host ("  AUTHOR:  {0}" -f $script:ScriptMetadata.Author) -ForegroundColor Cyan
    Write-Host ("  CONTACT: {0} | {1}" -f $script:ScriptMetadata.Email, $script:ScriptMetadata.GitHub) -ForegroundColor Cyan
    Write-Host '================================================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'DESCRIPTION:' -ForegroundColor Yellow
    Write-Host '  Tool for managing Microsoft 365 / Entra ID licenses via Microsoft Graph API.'
    Write-Host '  Supports inspecting user licenses, assigning licenses (with disabled plans),'
    Write-Host '  removing direct licenses, SKU-based user reporting, and tenant SKU inventory.'
    Write-Host '  Execution logs are saved to: C:\Logs\'
    Write-Host ''
    Write-Host 'MODES (ParameterSets):' -ForegroundColor Yellow
    Write-Host '  1. Inspect user licenses (default when -UserPrincipalName is provided):' -ForegroundColor White
    Write-Host '     -UserPrincipalName <UPN/Id> [-ShowServicePlans]' -ForegroundColor Gray
    Write-Host '  2. Assign a license to a user:' -ForegroundColor White
    Write-Host '     -UserPrincipalName <UPN/Id> -Assign -SkuPartNumber <SKU> [-DisabledPlans <Plans>] [-WhatIf]' -ForegroundColor Gray
    Write-Host '  3. Remove a license from a user:' -ForegroundColor White
    Write-Host '     -UserPrincipalName <UPN/Id> -Remove -SkuPartNumber <SKU> [-WhatIf] [-Confirm:$false]' -ForegroundColor Gray
    Write-Host '  4. Report users holding a specific SKU:' -ForegroundColor White
    Write-Host '     -Report -SkuPartNumber <SKU> [-ExportCsvPath <Path.csv>]' -ForegroundColor Gray
    Write-Host '  5. List all available SKUs in the tenant:' -ForegroundColor White
    Write-Host '     -ListAvailableSkus' -ForegroundColor Gray
    Write-Host '  6. Help:' -ForegroundColor White
    Write-Host '     -Help | -h  (or running without parameters)' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'GLOBAL PARAMETERS (optional):' -ForegroundColor Yellow
    Write-Host '  -TenantId <GUID/Domain>  -ClientId <AppId>  -CertificateThumbprint <Thumbprint>' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'USAGE EXAMPLES:' -ForegroundColor Yellow
    Write-Host '  # 1. Inspect user licenses (with expanded service plans):' -ForegroundColor Green
    Write-Host '  .\Manage-EntraUserLicenses.ps1 -UserPrincipalName "john.doe@domain.com" -ShowServicePlans'
    Write-Host ''
    Write-Host '  # 2. List all available licenses (SKUs) in the tenant:' -ForegroundColor Green
    Write-Host '  .\Manage-EntraUserLicenses.ps1 -ListAvailableSkus'
    Write-Host ''
    Write-Host '  # 3. Simulate license assignment with specific plans disabled (-WhatIf):' -ForegroundColor Green
    Write-Host '  .\Manage-EntraUserLicenses.ps1 -UPN "john.doe@domain.com" -Assign -SkuPartNumber "SPE_E3" -DisabledPlans @("YAMMER_ENTERPRISE","SWAY") -WhatIf'
    Write-Host ''
    Write-Host '  # 4. Remove a directly assigned license without confirmation prompt:' -ForegroundColor Green
    Write-Host '  .\Manage-EntraUserLicenses.ps1 -UPN "john.doe@domain.com" -Remove -SkuPartNumber "SPE_E3" -Confirm:$false'
    Write-Host ''
    Write-Host '  # 5. Generate a user report for ENTERPRISEPACK and export to CSV:' -ForegroundColor Green
    Write-Host '  .\Manage-EntraUserLicenses.ps1 -Report -SkuPartNumber "ENTERPRISEPACK" -ExportCsvPath "C:\Temp\Report_E3.csv"'
    Write-Host ''
    Write-Host 'AUTHOR & CONTACT:' -ForegroundColor Yellow
    Write-Host ("  Author : {0}" -f $script:ScriptMetadata.Author) -ForegroundColor Gray
    Write-Host ("  Email  : {0}" -f $script:ScriptMetadata.Email) -ForegroundColor Gray
    Write-Host ("  GitHub : {0}" -f $script:ScriptMetadata.GitHub) -ForegroundColor Gray
    Write-Host '  Full PowerShell help: Get-Help .\Manage-EntraUserLicenses.ps1 -Full' -ForegroundColor DarkGray
    Write-Host ''
}

function Initialize-LogDirectory {
    <#
    .SYNOPSIS
        Creates the C:\Logs\ directory if it does not exist.
    #>
    if (-not (Test-Path -Path $script:LogDirectory)) {
        New-Item -Path $script:LogDirectory -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message with an ISO 8601 timestamp to the console (color-coded) and to the log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formattedLine = "[{0}] [{1}] {2}" -f $now, $Level, $Message

    $color = switch ($Level) {
        'INFO'    { 'Cyan' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
    }

    Write-Host $formattedLine -ForegroundColor $color

    try {
        Add-Content -Path $script:LogFilePath -Value $formattedLine -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host ("[{0}] [WARN] Failed to write entry to log file '{1}': {2}" -f $now, $script:LogFilePath, $_.Exception.Message) -ForegroundColor Yellow
    }
}

#endregion Configuration & Logging

#region Helper Functions

function Get-SkuCatalog {
    <#
    .SYNOPSIS
        Retrieves tenant subscriptions (Get-MgSubscribedSku or Graph REST fallback) and builds case-insensitive lookup hashtables.
    #>
    [CmdletBinding()]
    param()

    Write-Log -Message 'Retrieving available subscriptions (SKUs) from Microsoft Entra ID tenant...' -Level 'INFO'

    if (Get-Module -ListAvailable -Name 'Microsoft.Graph.Identity.DirectoryManagement') {
        $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    }
    else {
        $skusList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $nextUri = 'https://graph.microsoft.com/v1.0/subscribedSkus'

        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
            if ($null -ne $response.value) {
                foreach ($item in $response.value) {
                    $servicePlans = @()
                    if ($null -ne $item.servicePlans) {
                        $servicePlans = @($item.servicePlans | ForEach-Object {
                            [PSCustomObject]@{
                                ServicePlanName    = $_.servicePlanName
                                ServicePlanId      = $_.servicePlanId
                                ProvisioningStatus = $_.provisioningStatus
                            }
                        })
                    }

                    $skusList.Add([PSCustomObject]@{
                        SkuPartNumber    = $item.skuPartNumber
                        SkuId            = $item.skuId
                        ConsumedUnits    = $item.consumedUnits
                        CapabilityStatus = $item.capabilityStatus
                        PrepaidUnits     = [PSCustomObject]@{
                            Enabled   = $item.prepaidUnits.enabled
                            Suspended = $item.prepaidUnits.suspended
                            Warning   = $item.prepaidUnits.warning
                        }
                        ServicePlans     = $servicePlans
                    })
                }
            }
            $nextUri = $response.'@odata.nextLink'
        }

        $skus = $skusList.ToArray()
    }

    $skuMap = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    $skuIdMap = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($sku in $skus) {
        if (-not [string]::IsNullOrWhiteSpace($sku.SkuPartNumber)) {
            $skuMap[$sku.SkuPartNumber] = $sku
        }
        if ($null -ne $sku.SkuId) {
            $skuIdMap[$sku.SkuId.ToString()] = $sku
        }
    }

    return [PSCustomObject]@{
        AllSkus = $skus
        ByPart  = $skuMap
        ById    = $skuIdMap
    }
}

function Invoke-EntraUserLicenseUpdate {
    <#
    .SYNOPSIS
        Updates a user's licenses using Set-MgUserLicense or Invoke-MgGraphRequest fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $false)]
        [hashtable[]]$AddLicenses = @(),

        [Parameter(Mandatory = $false)]
        [string[]]$RemoveLicenses = @()
    )

    if (Get-Module -ListAvailable -Name 'Microsoft.Graph.Users.Actions') {
        Set-MgUserLicense -UserId $UserId -AddLicenses $AddLicenses -RemoveLicenses $RemoveLicenses -ErrorAction Stop | Out-Null
    }
    else {
        $addPayload = @(
            foreach ($lic in $AddLicenses) {
                @{
                    skuId         = $lic.SkuId.ToString()
                    disabledPlans = @($lic.DisabledPlans | ForEach-Object { $_.ToString() })
                }
            }
        )
        $removePayload = @($RemoveLicenses | ForEach-Object { $_.ToString() })

        $body = @{
            addLicenses    = $addPayload
            removeLicenses = $removePayload
        } | ConvertTo-Json -Depth 5

        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/assignLicense"
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    }
}

function Resolve-TargetSku {
    <#
    .SYNOPSIS
        Validates the provided SkuPartNumber against the tenant catalog. Lists available SKUs on error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPartNumber,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Catalog
    )

    if ($Catalog.ByPart.ContainsKey($TargetPartNumber)) {
        return $Catalog.ByPart[$TargetPartNumber]
    }

    Write-Log -Message ("License SkuPartNumber '{0}' was not found in this tenant." -f $TargetPartNumber) -Level 'ERROR'
    Write-Log -Message 'Available SkuPartNumber values in this tenant:' -Level 'INFO'

    $Catalog.AllSkus |
        Sort-Object -Property SkuPartNumber |
        ForEach-Object {
            $enabled = if ($null -ne $_.PrepaidUnits.Enabled) { $_.PrepaidUnits.Enabled } else { 0 }
            $consumed = if ($null -ne $_.ConsumedUnits) { $_.ConsumedUnits } else { 0 }
            $free = $enabled - $consumed
            Write-Host ("  - {0,-30} (SkuId: {1}, Available: {2}/{3})" -f $_.SkuPartNumber, $_.SkuId, $free, $enabled) -ForegroundColor Gray
        }

    throw ("Invalid SkuPartNumber: '{0}'. Execution aborted." -f $TargetPartNumber)
}

function Get-AssignmentDetailsForSku {
    <#
    .SYNOPSIS
        Analyzes a user's LicenseAssignmentStates for a given SkuId and returns assignment type and group IDs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[]]$LicenseAssignmentStates,

        [Parameter(Mandatory = $true)]
        [string]$TargetSkuId
    )

    $matchingStates = @()
    if ($null -ne $LicenseAssignmentStates) {
        $matchingStates = @($LicenseAssignmentStates | Where-Object {
            $null -ne $_.SkuId -and $_.SkuId.ToString() -eq $TargetSkuId
        })
    }

    $hasDirect = $false
    $groupIds = [System.Collections.Generic.List[string]]::new()

    foreach ($state in $matchingStates) {
        if ([string]::IsNullOrWhiteSpace($state.AssignedByGroup)) {
            $hasDirect = $true
        }
        else {
            $groupIds.Add($state.AssignedByGroup.ToString())
        }
    }

    $assignmentType = if ($hasDirect -and $groupIds.Count -gt 0) {
        "Direct + Group ($($groupIds -join ', '))"
    }
    elseif ($groupIds.Count -gt 0) {
        "Group ($($groupIds -join ', '))"
    }
    elseif ($hasDirect) {
        'Direct'
    }
    else {
        'Direct'
    }

    return [PSCustomObject]@{
        HasDirect      = $hasDirect
        GroupIds       = $groupIds.ToArray()
        IsGroupOnly    = (-not $hasDirect -and $groupIds.Count -gt 0)
        AssignmentType = $assignmentType
    }
}

function Get-EntraUserSafe {
    <#
    .SYNOPSIS
        Retrieves a user object from Entra ID along with licensing properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Log -Message ("Looking up user '{0}' in Microsoft Entra ID..." -f $Identity) -Level 'INFO'
    $properties = @(
        'Id',
        'DisplayName',
        'UserPrincipalName',
        'UsageLocation',
        'AccountEnabled',
        'AssignedLicenses',
        'LicenseAssignmentStates'
    )

    try {
        return Get-MgUser -UserId $Identity -Property $properties -ErrorAction Stop
    }
    catch {
        Write-Log -Message ("Failed to retrieve user '{0}': {1}" -f $Identity, $_.Exception.Message) -Level 'ERROR'
        throw
    }
}

#endregion Helper Functions

#region Graph Connection

function Test-RequiredModules {
    <#
    .SYNOPSIS
        Verifies that all required Microsoft.Graph.* modules are installed.
    #>
    [CmdletBinding()]
    param()

    $missingModules = [System.Collections.Generic.List[string]]::new()
    foreach ($moduleName in $script:RequiredModules) {
        $installed = Get-Module -ListAvailable -Name $moduleName
        if (-not $installed) {
            $missingModules.Add($moduleName)
        }
    }

    if ($missingModules.Count -gt 0) {
        Write-Log -Message ("Missing required Microsoft Graph PowerShell SDK modules: {0}" -f ($missingModules -join ', ')) -Level 'ERROR'
        Write-Log -Message 'Install missing modules with: Install-Module Microsoft.Graph -Scope CurrentUser' -Level 'INFO'
        throw 'Missing required Microsoft.Graph.* modules.'
    }
}

function Connect-EntraGraph {
    <#
    .SYNOPSIS
        Establishes or verifies a Microsoft Graph API connection with the required scopes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParameterSetName,

        [string]$TenantIdParam,
        [string]$ClientIdParam,
        [string]$CertThumbprintParam
    )

    $isWriteMode = $ParameterSetName -in @('AssignLicense', 'RemoveLicense')

    $requiredScopes = if ($isWriteMode) {
        @('User.ReadWrite.All', 'Organization.Read.All')
    }
    else {
        @('User.Read.All', 'Organization.Read.All')
    }

    $hasAppOnlyParams = (-not [string]::IsNullOrWhiteSpace($TenantIdParam)) -and
                        (-not [string]::IsNullOrWhiteSpace($ClientIdParam)) -and
                        (-not [string]::IsNullOrWhiteSpace($CertThumbprintParam))

    if ($hasAppOnlyParams) {
        Write-Log -Message ("Connecting to Microsoft Graph in non-interactive App-Only mode (TenantId: {0}, ClientId: {1})..." -f $TenantIdParam, $ClientIdParam) -Level 'INFO'
        Connect-MgGraph -TenantId $TenantIdParam -ClientId $ClientIdParam -CertificateThumbprint $CertThumbprintParam -NoWelcome -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        Write-Log -Message ("Successfully connected to Microsoft Graph (App-Only, Tenant: {0})." -f $ctx.TenantId) -Level 'SUCCESS'
        return
    }

    if ((-not [string]::IsNullOrWhiteSpace($ClientIdParam)) -or (-not [string]::IsNullOrWhiteSpace($CertThumbprintParam))) {
        Write-Log -Message 'To use App-Only authentication, you must provide -TenantId, -ClientId, and -CertificateThumbprint together.' -Level 'ERROR'
        throw 'Incomplete App-Only authentication parameters.'
    }

    $context = Get-MgContext
    $needsReconnect = $false

    if ($null -eq $context) {
        $needsReconnect = $true
    }
    else {
        $currentScopes = @($context.Scopes)
        $hasReadUser = ($currentScopes -contains 'User.Read.All') -or ($currentScopes -contains 'User.ReadWrite.All') -or ($currentScopes -contains 'Directory.Read.All') -or ($currentScopes -contains 'Directory.ReadWrite.All')
        $hasOrgRead = ($currentScopes -contains 'Organization.Read.All') -or ($currentScopes -contains 'Directory.Read.All') -or ($currentScopes -contains 'Directory.ReadWrite.All')
        $hasWriteLicense = (-not $isWriteMode) -or ($currentScopes -contains 'User.ReadWrite.All') -or ($currentScopes -contains 'LicenseAssignment.ReadWrite.All') -or ($currentScopes -contains 'Directory.ReadWrite.All')

        if (-not ($hasReadUser -and $hasOrgRead -and $hasWriteLicense)) {
            Write-Log -Message 'The active Microsoft Graph session lacks sufficient permissions for the selected operation. Re-authentication is required.' -Level 'WARN'
            $needsReconnect = $true
        }
        elseif (-not [string]::IsNullOrWhiteSpace($TenantIdParam) -and $context.TenantId -ne $TenantIdParam) {
            Write-Log -Message ("Active session is connected to a different tenant ({0}). Switching to {1}..." -f $context.TenantId, $TenantIdParam) -Level 'INFO'
            $needsReconnect = $true
        }
    }

    if ($needsReconnect) {
        Write-Log -Message ("Establishing interactive Microsoft Graph connection (Scopes: {0})..." -f ($requiredScopes -join ', ')) -Level 'INFO'
        $connectParams = @{
            Scopes      = $requiredScopes
            NoWelcome   = $true
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantIdParam)) {
            $connectParams['TenantId'] = $TenantIdParam
        }

        Connect-MgGraph @connectParams | Out-Null
        $context = Get-MgContext
        Write-Log -Message ("Connected to Microsoft Graph as '{0}' (Tenant: {1})." -f $context.Account, $context.TenantId) -Level 'SUCCESS'
    }
    else {
        Write-Log -Message ("Using active Microsoft Graph session (Account: {0}, Tenant: {1})." -f $context.Account, $context.TenantId) -Level 'INFO'
    }
}

#endregion Graph Connection

#region Main Execution

if ($Help -or $PSCmdlet.ParameterSetName -eq 'Help') {
    Show-ScriptHelp
    return
}

Initialize-LogDirectory
Write-Log -Message ("Starting Manage-EntraUserLicenses (Mode: {0})" -f $PSCmdlet.ParameterSetName) -Level 'INFO'

try {
    Test-RequiredModules
    Connect-EntraGraph -ParameterSetName $PSCmdlet.ParameterSetName `
                       -TenantIdParam $TenantId `
                       -ClientIdParam $ClientId `
                       -CertThumbprintParam $CertificateThumbprint

    $skuCatalog = if ($PSCmdlet.ParameterSetName -ne 'GetUserLicenses') {
        Get-SkuCatalog
    }
    else {
        $null
    }

    switch ($PSCmdlet.ParameterSetName) {
        'ListTenantSkus' {
            Write-Log -Message 'Generating inventory of available subscriptions (SKUs) in the tenant...' -Level 'INFO'

            $skuResults = foreach ($sku in ($skuCatalog.AllSkus | Sort-Object -Property SkuPartNumber)) {
                $enabled = if ($null -ne $sku.PrepaidUnits.Enabled) { [int]$sku.PrepaidUnits.Enabled } else { 0 }
                $suspended = if ($null -ne $sku.PrepaidUnits.Suspended) { [int]$sku.PrepaidUnits.Suspended } else { 0 }
                $warning = if ($null -ne $sku.PrepaidUnits.Warning) { [int]$sku.PrepaidUnits.Warning } else { 0 }
                $consumed = if ($null -ne $sku.ConsumedUnits) { [int]$sku.ConsumedUnits } else { 0 }
                $available = $enabled - $consumed

                [PSCustomObject]@{
                    SkuPartNumber    = $sku.SkuPartNumber
                    SkuId            = $sku.SkuId
                    Enabled          = $enabled
                    ConsumedUnits    = $consumed
                    AvailableUnits   = $available
                    Suspended        = $suspended
                    Warning          = $warning
                    CapabilityStatus = $sku.CapabilityStatus
                }
            }

            Write-Log -Message ("Found {0} SKU subscriptions in the tenant." -f @($skuResults).Count) -Level 'SUCCESS'
            $skuResults | Format-Table -Property SkuPartNumber, SkuId, Enabled, ConsumedUnits, AvailableUnits, Suspended, Warning, CapabilityStatus -AutoSize
        }

        'GetUserLicenses' {
            $user = Get-EntraUserSafe -Identity $UserPrincipalName
            $licenseDetails = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop)

            $usageLocDisplay = if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) { '<None>' } else { $user.UsageLocation }
            Write-Log -Message ("User: {0} ({1}) | UsageLocation: {2} | AccountEnabled: {3}" -f $user.DisplayName, $user.UserPrincipalName, $usageLocDisplay, $user.AccountEnabled) -Level 'INFO'

            if ($licenseDetails.Count -eq 0) {
                Write-Log -Message 'User has no assigned licenses.' -Level 'INFO'
                break
            }

            $outputList = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($detail in $licenseDetails) {
                $skuIdStr = $detail.SkuId.ToString()
                $assignInfo = Get-AssignmentDetailsForSku -LicenseAssignmentStates $user.LicenseAssignmentStates -TargetSkuId $skuIdStr

                Write-Log -Message ("License: {0} | SkuId: {1} | Assignment: {2}" -f $detail.SkuPartNumber, $skuIdStr, $assignInfo.AssignmentType) -Level 'INFO'

                $servicePlansFormatted = @()
                if ($ShowServicePlans) {
                    $servicePlansFormatted = @(
                        $detail.ServicePlans |
                            Sort-Object -Property ServicePlanName |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    ServicePlanName    = $_.ServicePlanName
                                    ServicePlanId      = $_.ServicePlanId
                                    ProvisioningStatus = $_.ProvisioningStatus
                                }
                            }
                    )

                    $tableText = ($servicePlansFormatted | Format-Table -AutoSize | Out-String).TrimEnd()
                    if (-not [string]::IsNullOrWhiteSpace($tableText)) {
                        Write-Host $tableText -ForegroundColor Gray
                    }
                }

                $outputObj = [PSCustomObject]@{
                    DisplayName       = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    UsageLocation     = $user.UsageLocation
                    AccountEnabled    = $user.AccountEnabled
                    SkuPartNumber     = $detail.SkuPartNumber
                    SkuId             = $detail.SkuId
                    AssignmentType    = $assignInfo.AssignmentType
                    AssignedByGroup   = ($assignInfo.GroupIds -join ', ')
                }

                if ($ShowServicePlans) {
                    $outputObj | Add-Member -NotePropertyName 'ServicePlans' -NotePropertyValue $servicePlansFormatted
                }

                $outputList.Add($outputObj)
            }

            $outputList.ToArray()
        }

        'AssignLicense' {
            $targetSku = Resolve-TargetSku -TargetPartNumber $SkuPartNumber -Catalog $skuCatalog
            $user = Get-EntraUserSafe -Identity $UserPrincipalName

            if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
                Write-Log -Message ("User '{0}' does not have UsageLocation set (e.g., 'PL' or 'US'). Set UsageLocation before assigning a license." -f $user.UserPrincipalName) -Level 'ERROR'
                throw "Missing UsageLocation for user '$($user.UserPrincipalName)'."
            }

            $assignInfo = Get-AssignmentDetailsForSku -LicenseAssignmentStates $user.LicenseAssignmentStates -TargetSkuId $targetSku.SkuId.ToString()
            $hasExistingLicense = @($user.AssignedLicenses | Where-Object { $_.SkuId.ToString() -eq $targetSku.SkuId.ToString() }).Count -gt 0

            if ($assignInfo.HasDirect -or ($hasExistingLicense -and -not $assignInfo.IsGroupOnly)) {
                Write-Log -Message ("User '{0}' already has license '{1}' assigned directly. Skipped." -f $user.UserPrincipalName, $targetSku.SkuPartNumber) -Level 'WARN'
                break
            }

            $enabledUnits = if ($null -ne $targetSku.PrepaidUnits.Enabled) { [int]$targetSku.PrepaidUnits.Enabled } else { 0 }
            $consumedUnits = if ($null -ne $targetSku.ConsumedUnits) { [int]$targetSku.ConsumedUnits } else { 0 }
            $availableUnits = $enabledUnits - $consumedUnits

            if ($availableUnits -le 0) {
                Write-Log -Message ("No available licenses for SKU '{0}' (Enabled: {1}, Consumed: {2}, Available: {3})." -f $targetSku.SkuPartNumber, $enabledUnits, $consumedUnits, $availableUnits) -Level 'ERROR'
                throw "No available licenses for SKU '$($targetSku.SkuPartNumber)'."
            }

            $disabledPlanIds = @()
            if ($DisabledPlans -and $DisabledPlans.Count -gt 0) {
                $planMap = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($plan in $targetSku.ServicePlans) {
                    if (-not [string]::IsNullOrWhiteSpace($plan.ServicePlanName)) {
                        $planMap[$plan.ServicePlanName] = $plan.ServicePlanId
                    }
                }

                $resolvedIds = [System.Collections.Generic.List[guid]]::new()
                foreach ($planName in $DisabledPlans) {
                    if ($planMap.ContainsKey($planName)) {
                        $resolvedIds.Add([guid]$planMap[$planName])
                    }
                    else {
                        Write-Log -Message ("Service plan '{0}' was not found within SKU '{1}'." -f $planName, $targetSku.SkuPartNumber) -Level 'ERROR'
                        Write-Log -Message ("Available service plans in '{0}': {1}" -f $targetSku.SkuPartNumber, (($planMap.Keys | Sort-Object) -join ', ')) -Level 'INFO'
                        throw "Invalid service plan in -DisabledPlans: '$planName'."
                    }
                }
                $disabledPlanIds = $resolvedIds.ToArray()
            }

            $actionDescription = "Assign license $($targetSku.SkuPartNumber)"
            if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, $actionDescription)) {
                $addLicenseObj = @{
                    SkuId         = $targetSku.SkuId
                    DisabledPlans = $disabledPlanIds
                }

                Invoke-EntraUserLicenseUpdate -UserId $user.Id -AddLicenses @($addLicenseObj) -RemoveLicenses @()

                $remainingAfter = $availableUnits - 1
                Write-Log -Message ("Successfully assigned license '{0}' to user '{1}'. Remaining pool: {2}/{3}." -f $targetSku.SkuPartNumber, $user.UserPrincipalName, $remainingAfter, $enabledUnits) -Level 'SUCCESS'

                [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    SkuPartNumber     = $targetSku.SkuPartNumber
                    SkuId             = $targetSku.SkuId
                    Action            = 'Assign'
                    DisabledPlans     = ($DisabledPlans -join ', ')
                    Status            = 'Success'
                }
            }
            else {
                Write-Log -Message ("License assignment '{0}' for '{1}' was skipped (-WhatIf / not confirmed)." -f $targetSku.SkuPartNumber, $user.UserPrincipalName) -Level 'INFO'
            }
        }

        'RemoveLicense' {
            $targetSku = Resolve-TargetSku -TargetPartNumber $SkuPartNumber -Catalog $skuCatalog
            $user = Get-EntraUserSafe -Identity $UserPrincipalName

            $targetSkuIdStr = $targetSku.SkuId.ToString()
            $hasLicense = @($user.AssignedLicenses | Where-Object { $_.SkuId.ToString() -eq $targetSkuIdStr }).Count -gt 0

            if (-not $hasLicense) {
                Write-Log -Message ("User '{0}' does not have license '{1}' assigned. Finished without changes." -f $user.UserPrincipalName, $targetSku.SkuPartNumber) -Level 'WARN'
                break
            }

            $assignInfo = Get-AssignmentDetailsForSku -LicenseAssignmentStates $user.LicenseAssignmentStates -TargetSkuId $targetSkuIdStr
            if ($assignInfo.IsGroupOnly) {
                $groupsText = $assignInfo.GroupIds -join ', '
                Write-Log -Message ("License '{0}' is inherited from a group (Group ID: {1}). Direct removal via API is not possible; remove the user from the licensing group." -f $targetSku.SkuPartNumber, $groupsText) -Level 'ERROR'
                throw "License '$($targetSku.SkuPartNumber)' is assigned exclusively via group ($groupsText)."
            }

            $actionDescription = "Remove license $($targetSku.SkuPartNumber)"
            if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, $actionDescription)) {
                Invoke-EntraUserLicenseUpdate -UserId $user.Id -AddLicenses @() -RemoveLicenses @($targetSku.SkuId.ToString())
                Write-Log -Message ("Successfully removed license '{0}' from user '{1}'." -f $targetSku.SkuPartNumber, $user.UserPrincipalName) -Level 'SUCCESS'

                [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    SkuPartNumber     = $targetSku.SkuPartNumber
                    SkuId             = $targetSku.SkuId
                    Action            = 'Remove'
                    Status            = 'Success'
                }
            }
            else {
                Write-Log -Message ("License removal '{0}' for '{1}' was skipped (-WhatIf / not confirmed)." -f $targetSku.SkuPartNumber, $user.UserPrincipalName) -Level 'INFO'
            }
        }

        'ReportBySku' {
            $targetSku = Resolve-TargetSku -TargetPartNumber $SkuPartNumber -Catalog $skuCatalog
            $targetSkuIdStr = $targetSku.SkuId.ToString()

            $enabledUnits = if ($null -ne $targetSku.PrepaidUnits.Enabled) { [int]$targetSku.PrepaidUnits.Enabled } else { 0 }
            $suspendedUnits = if ($null -ne $targetSku.PrepaidUnits.Suspended) { [int]$targetSku.PrepaidUnits.Suspended } else { 0 }
            $warningUnits = if ($null -ne $targetSku.PrepaidUnits.Warning) { [int]$targetSku.PrepaidUnits.Warning } else { 0 }
            $consumedUnits = if ($null -ne $targetSku.ConsumedUnits) { [int]$targetSku.ConsumedUnits } else { 0 }
            $availableUnits = $enabledUnits - $consumedUnits

            Write-Log -Message '=== SKU License Summary ===' -Level 'INFO'
            Write-Log -Message ("SKU: {0} ({1})" -f $targetSku.SkuPartNumber, $targetSkuIdStr) -Level 'INFO'
            Write-Log -Message ("Enabled: {0} | Suspended: {1} | Warning: {2} | Consumed: {3} | Available: {4}" -f $enabledUnits, $suspendedUnits, $warningUnits, $consumedUnits, $availableUnits) -Level 'INFO'

            Write-Log -Message ("Retrieving users holding license '{0}'..." -f $targetSku.SkuPartNumber) -Level 'INFO'
            $userProperties = @(
                'Id',
                'DisplayName',
                'UserPrincipalName',
                'Department',
                'JobTitle',
                'AccountEnabled',
                'UsageLocation',
                'LicenseAssignmentStates'
            )

            $filterQuery = "assignedLicenses/any(s:s/skuId eq $targetSkuIdStr)"
            $licensedUsers = @(Get-MgUser -Filter $filterQuery -ConsistencyLevel eventual -CountVariable UserCount -All -Property $userProperties -ErrorAction Stop)

            $totalFound = $licensedUsers.Count
            Write-Log -Message ("Found {0} users with license '{1}' assigned." -f $totalFound, $targetSku.SkuPartNumber) -Level 'INFO'

            $reportRows = [System.Collections.Generic.List[PSCustomObject]]::new()
            $index = 0

            foreach ($u in $licensedUsers) {
                $index++
                if ($totalFound -gt 0) {
                    $percentComplete = [math]::Min(100, [math]::Round(($index / $totalFound) * 100))
                    Write-Progress -Activity ("Processing license {0}" -f $targetSku.SkuPartNumber) `
                                   -Status ("User {0} of {1}: {2}" -f $index, $totalFound, $u.UserPrincipalName) `
                                   -PercentComplete $percentComplete
                }

                $assignInfo = Get-AssignmentDetailsForSku -LicenseAssignmentStates $u.LicenseAssignmentStates -TargetSkuId $targetSkuIdStr

                $reportRows.Add([PSCustomObject]@{
                    DisplayName       = $u.DisplayName
                    UserPrincipalName = $u.UserPrincipalName
                    Department        = $u.Department
                    JobTitle          = $u.JobTitle
                    AccountEnabled    = $u.AccountEnabled
                    UsageLocation     = $u.UsageLocation
                    AssignmentType    = $assignInfo.AssignmentType
                })
            }

            Write-Progress -Activity ("Processing license {0}" -f $targetSku.SkuPartNumber) -Completed

            $sortedRows = @($reportRows | Sort-Object -Property DisplayName, UserPrincipalName)

            if (-not [string]::IsNullOrWhiteSpace($ExportCsvPath)) {
                $exportDir = Split-Path -Path $ExportCsvPath -Parent
                if (-not [string]::IsNullOrWhiteSpace($exportDir) -and -not (Test-Path -Path $exportDir)) {
                    New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
                }
                $sortedRows | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                Write-Log -Message ("Exported report ({0} records) to CSV file: {1}" -f $sortedRows.Count, $ExportCsvPath) -Level 'SUCCESS'
            }

            $sortedRows
        }
    }

    Write-Log -Message 'Script execution completed successfully.' -Level 'SUCCESS'
}
catch {
    Write-Log -Message ("Critical error: {0}" -f $_.Exception.Message) -Level 'ERROR'
    if ($_.ScriptStackTrace) {
        Write-Log -Message ("StackTrace: {0}" -f $_.ScriptStackTrace) -Level 'ERROR'
    }
    exit 1
}

#endregion Main Execution
