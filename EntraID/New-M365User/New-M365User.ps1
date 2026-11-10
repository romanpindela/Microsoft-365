<#
.SYNOPSIS
    Automates the provisioning of a new user identity in Microsoft Entra ID / Microsoft 365.

.DESCRIPTION
    New-M365User.ps1 provisions a new user account in Microsoft Entra ID / Microsoft 365,
    configures a secure initial password with mandatory change at next sign-in
    (forceChangePasswordNextSignIn = $true), assigns a specified license plan
    (supporting both friendly SkuPartNumber and GUID SkuId), enforces ISO 3166-1
    country usage location, assigns standard user privileges (no privileged roles),
    sets user profile attributes (GivenName, Surname, Department),
    and optionally adds the user to designated Microsoft 365 / Microsoft Teams groups.

    Running the script without parameters or with -Help / -h displays comprehensive usage
    instructions, parameter details, and copy-paste ready examples.

.PARAMETER UserPrincipalName
    The UserPrincipalName / primary email address of the new user (e.g. john.doe@domain.com).
    Aliases: -Email, -UPN, -User.

.PARAMETER DisplayName
    The full display name of the user (e.g. John Doe). If omitted, it is automatically
    composed from GivenName and Surname, or derived from the prefix of UserPrincipalName.

.PARAMETER GivenName
    The first name of the user (e.g. John). Aliases: -FirstName, -Imie.

.PARAMETER Surname
    The last name / surname of the user (e.g. Doe). Aliases: -LastName, -Nazwisko.

.PARAMETER Department
    The organizational department of the user (e.g. 'IT / Helpdesk', 'Sales'). Aliases: -Dzial.

.PARAMETER MailNickname
    The mail alias / nickname of the user. If omitted, it is automatically derived
    from the portion of the UserPrincipalName preceding the '@' sign.

.PARAMETER UsageLocation
    Two-letter ISO 3166-1 alpha-2 country code required by Microsoft 365 for license assignment.
    Defaults to 'PL'.

.PARAMETER InitialPassword
    The initial account password as a [SecureString]. Passwords in plain text are strictly
    rejected to comply with enterprise credential security standards. Aliases: -Password, -Haslo.

.PARAMETER LicenseSku
    The license SKU to assign. Accepts friendly SkuPartNumber (e.g. 'ENTERPRISEPACK', 'SPE_E3',
    'O365_BUSINESS_PREMIUM') or direct SkuId GUID. Aliases: -License, -Sku, -Licencja.

.PARAMETER TeamGroupNames
    One or more Microsoft 365 / Teams group display names to which the new user should be added.
    Accepts string array or comma-separated list. Aliases: -TeamGroupName, -TeamsGroup, -GrupyTeams, -GrupaTeams.

.PARAMETER LogDirectory
    Path to the directory where operational audit logs will be saved.
    Defaults to 'C:\Logs\Microsoft-365'.

.PARAMETER Help
    Displays custom help, usage instructions, author information, and exits. Aliases: -h.

.PARAMETER WhatIf
    Simulates script execution without making changes in Entra ID or Teams.

.PARAMETER Confirm
    Prompts for confirmation before executing state-changing operations.

.EXAMPLE
    # Complete user provisioning (Email, Name, Department, Teams, License, Password):
    $secPass = "InitialPass123!" | ConvertTo-SecureString -AsPlainText -Force
    .\New-M365User.ps1 -UserPrincipalName "john.doe@domain.com" `
                       -GivenName "John" `
                       -Surname "Doe" `
                       -Department "IT / Helpdesk" `
                       -TeamGroupNames "Technical Support", "General" `
                       -LicenseSku "SPE_E3" `
                       -InitialPassword $secPass

.EXAMPLE
    # Masked console password input (Read-Host):
    $secPass = Read-Host "Enter initial password" -AsSecureString
    .\New-M365User.ps1 -Email "jane.smith@domain.com" `
                       -GivenName "Jane" `
                       -Surname "Smith" `
                       -Department "Sales" `
                       -TeamGroupNames "Global Sales", "Commercial" `
                       -LicenseSku "ENTERPRISEPACK" `
                       -InitialPassword $secPass

.EXAMPLE
    # Display help instructions:
    .\New-M365User.ps1 -Help
    # or simply run without parameters:
    .\New-M365User.ps1

.NOTES
    Script: New-M365User.ps1
    Version: 1.1.2
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Required Modules: Microsoft.Graph.Users, Microsoft.Graph.Authentication
    Optional Modules: Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups
    Required Scopes: User.ReadWrite.All, Directory.ReadWrite.All, Group.ReadWrite.All
#>

[CmdletBinding(DefaultParameterSetName = 'ProvisionUser', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, Position = 0, HelpMessage = "UserPrincipalName / primary email address of the new user.")]
    [Alias('Email', 'UPN', 'User')]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, Position = 1, HelpMessage = "Full display name of the user (DisplayName).")]
    [string]$DisplayName,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "First name of the user (GivenName).")]
    [Alias('FirstName', 'Imie')]
    [string]$GivenName,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Last name / surname of the user (Surname).")]
    [Alias('LastName', 'Nazwisko')]
    [string]$Surname,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Organizational department / division (Department).")]
    [Alias('Dzial')]
    [string]$Department,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Mail alias / nickname (MailNickname). Defaults to UPN prefix.")]
    [string]$MailNickname,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "ISO 3166-1 alpha-2 country code for licensing (defaults to PL).")]
    [ValidatePattern('^[A-Za-z]{2}$')]
    [string]$UsageLocation = 'PL',

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Initial account password as a SecureString.")]
    [Alias('Password', 'Haslo')]
    [securestring]$InitialPassword,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "License plan name (SkuPartNumber) or SkuId GUID.")]
    [Alias('License', 'Sku', 'Licencja')]
    [string]$LicenseSku,

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Microsoft 365 / Teams group display names.")]
    [Alias('TeamGroupName', 'TeamsGroup', 'GrupyTeams', 'GrupaTeams')]
    [string[]]$TeamGroupNames = @(),

    [Parameter(ParameterSetName = 'ProvisionUser', Mandatory = $false, HelpMessage = "Directory path for operational audit logs.")]
    [string]$LogDirectory = 'C:\Logs\Microsoft-365',

    [Parameter(ParameterSetName = 'Help', Mandatory = $false)]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# ENVIRONMENT & MODULE PATH BOOTSTRAP
# ==============================================================================
$localAppData = [Environment]::GetFolderPath('LocalApplicationData')
if (-not [string]::IsNullOrWhiteSpace($localAppData)) {
    $localModulesPath = Join-Path -Path $localAppData -ChildPath "WindowsPowerShell\Modules"
    if (Test-Path -Path $localModulesPath) {
        if ($env:PSModulePath -notlike "*$localModulesPath*") {
            $env:PSModulePath = "$localModulesPath;$env:PSModulePath"
        }
    }
}

# Try importing available Graph modules silently
@('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement', 'Microsoft.Graph.Groups') | ForEach-Object {
    if (-not (Get-Module -Name $_)) {
        try { Import-Module -Name $_ -ErrorAction SilentlyContinue } catch {}
    }
}

# ==============================================================================
# SCRIPT METADATA & TRACKING
# ==============================================================================
$Script:Metadata = @{
    Name        = 'New-M365User'
    Version     = '1.1.2'
    Author      = 'Roman Pindela'
    StartTime   = Get-Date
    LogFileName = "New-M365User_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd')
    HasWarnings = $false
}

$Script:LogFilePath = Join-Path -Path $LogDirectory -ChildPath $Script:Metadata.LogFileName

# ==============================================================================
# HELPER: SAFE OBJECT PROPERTY RESOLVER (Set-StrictMode Safe)
# ==============================================================================
function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$PropertyName,
        [object]$DefaultValue = $null
    )
    if ($null -eq $InputObject) { return $DefaultValue }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName) -and $null -ne $InputObject[$PropertyName]) {
            return $InputObject[$PropertyName]
        }
        return $DefaultValue
    }

    $prop = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $prop -and $null -ne $prop.Value) {
        return $prop.Value
    }
    return $DefaultValue
}

# ==============================================================================
# LOGGING MODULE (Write-M365Log)
# ==============================================================================
function Write-M365Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Message,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $levelTag = "[$Level]".PadRight(9) + " "
    $logEntry = "[$timestamp] $levelTag$Message"

    $consoleColor = switch ($Level) {
        'INFO'    { [ConsoleColor]::Cyan }
        'SUCCESS' { [ConsoleColor]::Green }
        'WARN'    { [ConsoleColor]::Yellow }
        'ERROR'   { [ConsoleColor]::Red }
        'DEBUG'   { [ConsoleColor]::DarkGray }
        Default   { [ConsoleColor]::White }
    }

    Write-Host $logEntry -ForegroundColor $consoleColor

    if ($Level -eq 'WARN') {
        $Script:Metadata.HasWarnings = $true
    }

    try {
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        [System.IO.File]::AppendAllLines($Script:LogFilePath, [string[]]@($logEntry), [System.Text.Encoding]::UTF8)
    } catch {
        Write-Warning "Could not write log entry to '$Script:LogFilePath': $($_.Exception.Message)"
    }
}

# ==============================================================================
# INTERACTIVE HELP MODULE (Show-ScriptHelp)
# ==============================================================================
function Show-ScriptHelp {
    $helpLines = @(
        "================================================================================",
        "SCRIPT  : New-M365User.ps1",
        "VERSION : 1.1.2",
        "AUTHOR  : Roman Pindela",
        "CONTACT : roman.pindela@gmail.com | https://github.com/romanpindela",
        "================================================================================",
        "",
        "DESCRIPTION:",
        "    Automates the secure provisioning of a new user identity in Microsoft",
        "    Entra ID / Microsoft 365. Configures a secure initial password with mandatory",
        "    change at next sign-in, assigns the specified license plan, populates user",
        "    profile attributes (Given Name, Surname, Department), and automatically adds",
        "    the account to designated Microsoft Teams / Microsoft 365 groups.",
        "",
        "PARAMETERS:",
        "    -UserPrincipalName, -Email, -UPN",
        "        Primary email address / login of the new user (e.g. john.doe@domain.com). [REQUIRED]",
        "",
        "    -GivenName, -FirstName",
        "        First name of the user (e.g. John). [OPTIONAL]",
        "",
        "    -Surname, -LastName",
        "        Last name / surname of the user (e.g. Doe). [OPTIONAL]",
        "",
        "    -Department",
        "        Organizational department or unit (e.g. 'IT / Helpdesk', 'Sales'). [OPTIONAL]",
        "",
        "    -DisplayName",
        "        Full display name of the user (e.g. 'John Doe').",
        "        If omitted, automatically assembled from GivenName + Surname or derived",
        "        from the prefix of UserPrincipalName.",
        "",
        "    -TeamGroupNames, -TeamsGroup, -TeamGroupName",
        "        Array of Microsoft 365 / Teams group display names to which the user",
        "        will be added (e.g. 'Technical Support', 'General'). [OPTIONAL]",
        "",
        "    -LicenseSku, -License, -Sku",
        "        License SKU plan name (e.g. 'SPE_E3', 'ENTERPRISEPACK',",
        "        'O365_BUSINESS_PREMIUM') or direct GUID (SkuId). [REQUIRED]",
        "",
        "    -InitialPassword, -Password",
        "        Initial account password provided as a [SecureString]. [REQUIRED]",
        "",
        "    -UsageLocation",
        "        Two-letter ISO 3166-1 alpha-2 country code (default: 'PL').",
        "",
        "    -LogDirectory",
        "        Directory path for operational audit log files (default: C:\Logs\Microsoft-365).",
        "",
        "    -WhatIf",
        "        Simulation switch - displays planned actions without making actual changes.",
        "",
        "    -Help, -h",
        "        Displays this help screen and exits.",
        "",
        "USAGE EXAMPLES:",
        "",
        "  1. Complete user provisioning (Email, First Name, Last Name, Department, Teams, License):",
        "     `$secPass = `"InitialPass123!`" | ConvertTo-SecureString -AsPlainText -Force",
        "     .\New-M365User.ps1 -UserPrincipalName `"john.doe@domain.com`" ``",
        "                        -GivenName `"John`" ``",
        "                        -Surname `"Doe`" ``",
        "                        -Department `"IT / Helpdesk`" ``",
        "                        -TeamGroupNames `"Technical Support`", `"General`" ``",
        "                        -LicenseSku `"SPE_E3`" ``",
        "                        -InitialPassword `$secPass",
        "",
        "  2. Interactive masked password input (Read-Host):",
        "     `$secPass = Read-Host `"Enter initial password`" -AsSecureString",
        "     .\New-M365User.ps1 -Email `"jane.smith@domain.com`" ``",
        "                        -GivenName `"Jane`" ``",
        "                        -Surname `"Smith`" ``",
        "                        -Department `"Sales`" ``",
        "                        -TeamGroupNames `"Global Sales`", `"Commercial`" ``",
        "                        -LicenseSku `"ENTERPRISEPACK`" ``",
        "                        -InitialPassword `$secPass",
        "",
        "  3. Safe dry-run simulation (-WhatIf):",
        "     `$secPass = `"Test1234!`" | ConvertTo-SecureString -AsPlainText -Force",
        "     .\New-M365User.ps1 -UserPrincipalName `"alex.turner@domain.com`" ``",
        "                        -GivenName `"Alex`" ``",
        "                        -Surname `"Turner`" ``",
        "                        -Department `"Finance`" ``",
        "                        -TeamGroupNames `"Finance and Accounting`" ``",
        "                        -LicenseSku `"SPE_E3`" ``",
        "                        -InitialPassword `$secPass ``",
        "                        -WhatIf",
        "",
        "REQUIRED MODULES & GRAPH API PERMISSIONS:",
        "    Modules: Microsoft.Graph.Users, Microsoft.Graph.Authentication",
        "    Scopes : User.ReadWrite.All, Directory.ReadWrite.All, Group.ReadWrite.All",
        "    Sign-in: Connect-MgGraph -Scopes 'User.ReadWrite.All','Directory.ReadWrite.All','Group.ReadWrite.All'",
        "================================================================================"
    )

    Write-Host ($helpLines -join "`n") -ForegroundColor Cyan
}

# Check if help is requested or script is executed without parameters
if ($Help -or $PSBoundParameters.Count -eq 0) {
    Show-ScriptHelp
    exit 0
}

if ($PSBoundParameters.Count -eq 1 -and $PSBoundParameters.ContainsKey('LogDirectory')) {
    Show-ScriptHelp
    exit 0
}

# ==============================================================================
# VALIDATION OF REQUIRED PARAMETERS
# ==============================================================================
$missingParams = @()
if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    $missingParams += "-UserPrincipalName (or alias -Email / -UPN)"
}
if ($null -eq $InitialPassword) {
    $missingParams += "-InitialPassword (or alias -Password as [SecureString])"
}
if ([string]::IsNullOrWhiteSpace($LicenseSku)) {
    $missingParams += "-LicenseSku (or alias -License / -Sku, e.g. 'SPE_E3', 'O365_BUSINESS_PREMIUM')"
}

if ($missingParams.Count -gt 0) {
    Write-Host "`n[ERROR] Missing mandatory parameters for user provisioning:" -ForegroundColor Red
    foreach ($mp in $missingParams) {
        Write-Host "  * $mp" -ForegroundColor Yellow
    }
    Write-Host "`nTo display full help instructions and usage examples, run the script with -Help:" -ForegroundColor White
    Write-Host "    .\New-M365User.ps1 -Help`n" -ForegroundColor Cyan
    exit 1
}

# Auto-compute DisplayName if omitted
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    if (-not [string]::IsNullOrWhiteSpace($GivenName) -and -not [string]::IsNullOrWhiteSpace($Surname)) {
        $DisplayName = "$GivenName $Surname".Trim()
    } elseif (-not [string]::IsNullOrWhiteSpace($GivenName)) {
        $DisplayName = $GivenName.Trim()
    } else {
        $DisplayName = ($UserPrincipalName -split '@')[0]
    }
}

# ==============================================================================
# SECURE PASSWORD HELPER (Zero memory leak)
# ==============================================================================
function Convert-SecureStringToPlainText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [securestring]$SecureString
    )

    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

# ==============================================================================
# RESILIENT GRAPH SKUS & GROUPS RESOLVERS (Works with or without sub-modules)
# ==============================================================================
function Get-TenantSubscribedSkus {
    # 1. Try cmdlet from Microsoft.Graph.Identity.DirectoryManagement
    if (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue) {
        try {
            return @(Get-MgSubscribedSku -All -ErrorAction Stop)
        } catch {
            Write-M365Log "Get-MgSubscribedSku failed ($($_.Exception.Message)), attempting REST fallback..." -Level 'DEBUG'
        }
    }

    # 2. Resilient fallback via Invoke-MgGraphRequest (Microsoft.Graph.Authentication)
    if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
        $allSkus = @()
        while (-not [string]::IsNullOrWhiteSpace($uri)) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            $items = Get-ObjectPropertyValue -InputObject $response -PropertyName 'value' -DefaultValue $response
            if ($null -ne $items) {
                foreach ($item in $items) {
                    $partNumber = Get-ObjectPropertyValue -InputObject $item -PropertyName 'skuPartNumber'
                    if ($null -eq $partNumber) {
                        $partNumber = Get-ObjectPropertyValue -InputObject $item -PropertyName 'SkuPartNumber'
                    }

                    $skuGuid = Get-ObjectPropertyValue -InputObject $item -PropertyName 'skuId'
                    if ($null -eq $skuGuid) {
                        $skuGuid = Get-ObjectPropertyValue -InputObject $item -PropertyName 'SkuId'
                    }

                    $consumed = Get-ObjectPropertyValue -InputObject $item -PropertyName 'consumedUnits'
                    if ($null -eq $consumed) {
                        $consumed = Get-ObjectPropertyValue -InputObject $item -PropertyName 'ConsumedUnits' -DefaultValue 0
                    }

                    $prepObj = Get-ObjectPropertyValue -InputObject $item -PropertyName 'prepaidUnits'
                    if ($null -eq $prepObj) {
                        $prepObj = Get-ObjectPropertyValue -InputObject $item -PropertyName 'PrepaidUnits'
                    }

                    $enabled = 0
                    if ($null -ne $prepObj) {
                        $en = Get-ObjectPropertyValue -InputObject $prepObj -PropertyName 'enabled'
                        if ($null -eq $en) {
                            $en = Get-ObjectPropertyValue -InputObject $prepObj -PropertyName 'Enabled' -DefaultValue 0
                        }
                        $enabled = $en
                    }

                    $allSkus += [pscustomobject]@{
                        SkuPartNumber = $partNumber
                        SkuId         = $skuGuid
                        ConsumedUnits = [int]$consumed
                        PrepaidUnits  = [pscustomobject]@{ Enabled = [int]$enabled }
                    }
                }
            }

            $uri = Get-ObjectPropertyValue -InputObject $response -PropertyName '@odata.nextLink'
        }
        return $allSkus
    }

    throw "Neither 'Get-MgSubscribedSku' nor 'Invoke-MgGraphRequest' is available in current session."
}

function Find-TenantGroup {
    param([string]$DisplayName)

    # 1. Try cmdlet from Microsoft.Graph.Groups
    if (Get-Command Get-MgGroup -ErrorAction SilentlyContinue) {
        try {
            $escaped = $DisplayName.Replace("'", "''")
            return @(Get-MgGroup -Filter "displayName eq '$escaped'" -ErrorAction Stop)
        } catch {
            Write-M365Log "Get-MgGroup failed ($($_.Exception.Message)), attempting REST fallback..." -Level 'DEBUG'
        }
    }

    # 2. Resilient fallback via Invoke-MgGraphRequest
    if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $escaped = [System.Uri]::EscapeDataString($DisplayName.Replace("'", "''"))
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escaped'"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $items = Get-ObjectPropertyValue -InputObject $response -PropertyName 'value' -DefaultValue $response
        $results = @()
        if ($null -ne $items) {
            foreach ($item in $items) {
                $gId = Get-ObjectPropertyValue -InputObject $item -PropertyName 'id'
                if ($null -eq $gId) { $gId = Get-ObjectPropertyValue -InputObject $item -PropertyName 'Id' }
                $gName = Get-ObjectPropertyValue -InputObject $item -PropertyName 'displayName'
                if ($null -eq $gName) { $gName = Get-ObjectPropertyValue -InputObject $item -PropertyName 'DisplayName' }
                $results += [pscustomobject]@{ Id = $gId; DisplayName = $gName }
            }
        }
        return $results
    }

    throw "Neither 'Get-MgGroup' nor 'Invoke-MgGraphRequest' is available in current session."
}

function Add-TenantGroupMember {
    param(
        [string]$GroupId,
        [string]$UserId
    )

    $memberOdataId = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"

    # 1. Try cmdlet from Microsoft.Graph.Groups
    if (Get-Command New-MgGroupMemberByRef -ErrorAction SilentlyContinue) {
        try {
            $byRefCmd = Get-Command New-MgGroupMemberByRef
            if ($byRefCmd.Parameters.ContainsKey('OdataId')) {
                New-MgGroupMemberByRef -GroupId $GroupId -OdataId $memberOdataId -ErrorAction Stop | Out-Null
                return
            } elseif ($byRefCmd.Parameters.ContainsKey('BodyParameter')) {
                $bodyParam = @{ "@odata.id" = $memberOdataId }
                New-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $bodyParam -ErrorAction Stop | Out-Null
                return
            }
        } catch {
            Write-M365Log "New-MgGroupMemberByRef failed ($($_.Exception.Message)), attempting REST fallback..." -Level 'DEBUG'
        }
    }

    # 2. Resilient fallback via Invoke-MgGraphRequest
    if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"
        $body = @{ "@odata.id" = $memberOdataId } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        return
    }

    throw "Neither 'New-MgGroupMemberByRef' nor 'Invoke-MgGraphRequest' is available in current session."
}

# ==============================================================================
# MAIN PROVISIONING LOGIC
# ==============================================================================
Write-M365Log "Initializing user provisioning for: $UserPrincipalName" -Level 'INFO'

# 1. Validation of UPN format
$emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
if ($UserPrincipalName -notmatch $emailRegex) {
    Write-M365Log "Invalid UserPrincipalName format: '$UserPrincipalName'." -Level 'ERROR'
    exit 1
}

# Normalize UsageLocation and MailNickname
$UsageLocation = $UsageLocation.Trim().ToUpper()
if ([string]::IsNullOrWhiteSpace($MailNickname)) {
    $MailNickname = ($UserPrincipalName -split '@')[0]
}

# 2. Verify Microsoft Graph Session and Permissions
try {
    $mgContext = Get-MgContext -ErrorAction Stop
} catch {
    $mgContext = $null
}

if ($null -eq $mgContext -or [string]::IsNullOrWhiteSpace($mgContext.TenantId)) {
    Write-M365Log "No active Microsoft Graph session detected. Please connect using Connect-MgGraph." -Level 'ERROR'
    Write-Host "`nTo establish a session with the required permissions, run:" -ForegroundColor Yellow
    Write-Host "    Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All'`n" -ForegroundColor White
    exit 1
}

Write-M365Log "Active Graph session detected for tenant: $($mgContext.TenantId) (Account: $($mgContext.Account))" -Level 'DEBUG'

# Check required scopes
$requiredScopes = @('User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All')
$activeScopes = @()
$ctxScopes = Get-ObjectPropertyValue -InputObject $mgContext -PropertyName 'Scopes'
if ($null -ne $ctxScopes) {
    $activeScopes = $ctxScopes
}

$missingScopes = @()
foreach ($reqScope in $requiredScopes) {
    if ($activeScopes -notcontains $reqScope) {
        $missingScopes += $reqScope
    }
}

if ($missingScopes.Count -gt 0) {
    Write-M365Log "Active Graph session lacks recommended permissions (missing: $($missingScopes -join ', ')). Some operations may fail with 403 Forbidden." -Level 'WARN'
}

# 3. Validate and resolve License SKU in Tenant
Write-M365Log "Validating availability of license SKU: $LicenseSku..." -Level 'INFO'
$resolvedSkuId = $null
$resolvedSkuPartNumber = $null

try {
    $allSkus = Get-TenantSubscribedSkus
    $targetSku = $allSkus | Where-Object {
        $_.SkuPartNumber -ieq $LicenseSku -or "$($_.SkuId)" -ieq $LicenseSku
    } | Select-Object -First 1

    if ($null -eq $targetSku) {
        Write-M365Log "License SKU '$LicenseSku' was not found in tenant." -Level 'ERROR'
        Write-Host "`nAvailable license SKUs in the current tenant:" -ForegroundColor Yellow
        $allSkus | Select-Object SkuPartNumber, SkuId, ConsumedUnits, @{Name='TotalUnits'; Expression={$_.PrepaidUnits.Enabled}} | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Gray
        exit 1
    }

    $resolvedSkuId = $targetSku.SkuId
    $resolvedSkuPartNumber = $targetSku.SkuPartNumber

    # Check capacity
    $totalUnits = $targetSku.PrepaidUnits.Enabled
    $consumedUnits = $targetSku.ConsumedUnits
    $availableUnits = $totalUnits - $consumedUnits

    if ($availableUnits -le 0) {
        Write-M365Log "All units for license $resolvedSkuPartNumber have been consumed ($consumedUnits/$totalUnits). License assignment may fail." -Level 'WARN'
    } else {
        Write-M365Log "Resolved license: $resolvedSkuPartNumber ($resolvedSkuId). Available units: $availableUnits/$totalUnits." -Level 'DEBUG'
    }
} catch {
    Write-M365Log "Error retrieving tenant licenses: $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

# 4. Check if User already exists in Entra ID
Write-M365Log "Checking UserPrincipalName uniqueness: $UserPrincipalName..." -Level 'DEBUG'
try {
    $existingUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
    if ($null -ne $existingUser) {
        Write-M365Log "User with UserPrincipalName '$UserPrincipalName' already exists in the tenant (UserId: $($existingUser.Id)). Operation aborted." -Level 'ERROR'
        exit 1
    }
} catch {
    Write-M365Log "Error checking account existence: $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

# 5. Create Entra ID User Account
$createdUser = $null
$plainPassword = Convert-SecureStringToPlainText -SecureString $InitialPassword

$userParams = @{
    AccountEnabled    = $true
    DisplayName       = $DisplayName
    UserPrincipalName = $UserPrincipalName
    MailNickname      = $MailNickname
    UsageLocation     = $UsageLocation
    PasswordProfile   = @{
        ForceChangePasswordNextSignIn = $true
        Password                      = $plainPassword
    }
}

if (-not [string]::IsNullOrWhiteSpace($GivenName)) {
    $userParams['GivenName'] = $GivenName
}
if (-not [string]::IsNullOrWhiteSpace($Surname)) {
    $userParams['Surname'] = $Surname
}
if (-not [string]::IsNullOrWhiteSpace($Department)) {
    $userParams['Department'] = $Department
}

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Create Entra ID User Account (Standard User)")) {
        $createdUser = New-MgUser @userParams -ErrorAction Stop
        Write-M365Log "User account successfully created (ID: $($createdUser.Id))" -Level 'SUCCESS'
    } else {
        # WhatIf / Simulation mock
        $createdUser = [pscustomobject]@{
            Id                = '00000000-0000-0000-0000-000000000000'
            DisplayName       = $DisplayName
            UserPrincipalName = $UserPrincipalName
        }
        Write-M365Log "[WHAT-IF] Simulated user account creation: $UserPrincipalName (Mock ID: $($createdUser.Id))" -Level 'INFO'
    }
} catch {
    Write-M365Log "Fatal error while creating user '$UserPrincipalName': $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

# 6. Assign License
Write-M365Log "Assigning license: $resolvedSkuPartNumber..." -Level 'INFO'
$licenseAssigned = $false

try {
    $licenseAdd = @{
        SkuId = [guid]$resolvedSkuId
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Assign License $resolvedSkuPartNumber ($resolvedSkuId)")) {
        Set-MgUserLicense -UserId $createdUser.Id -AddLicenses @($licenseAdd) -RemoveLicenses @() -ErrorAction Stop
        Write-M365Log "License assigned successfully." -Level 'SUCCESS'
        $licenseAssigned = $true
    } else {
        Write-M365Log "[WHAT-IF] Simulated license assignment: $resolvedSkuPartNumber" -Level 'INFO'
        $licenseAssigned = $true
    }
} catch {
    Write-M365Log "Failed to assign license to user account ($($createdUser.Id)): $($_.Exception.Message)" -Level 'WARN'
    $licenseAssigned = $false
}

# 7. Add User to Microsoft 365 / Teams Groups
$groupResults = @()

if ($TeamGroupNames -and $TeamGroupNames.Count -gt 0) {
    foreach ($groupName in $TeamGroupNames) {
        $trimmedGroup = $groupName.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedGroup)) { continue }

        Write-M365Log "Searching for Teams/M365 group: '$trimmedGroup'..." -Level 'DEBUG'
        try {
            $foundGroups = Find-TenantGroup -DisplayName $trimmedGroup

            if ($null -eq $foundGroups -or ($foundGroups | Measure-Object).Count -eq 0) {
                Write-M365Log "Group '$trimmedGroup' was not found in the tenant." -Level 'WARN'
                $groupResults += "[$($trimmedGroup): NOT FOUND]"
                continue
            }

            # Select target group
            $targetGroup = $foundGroups | Select-Object -First 1

            if ($PSCmdlet.ShouldProcess("$trimmedGroup ($($targetGroup.Id))", "Add user $UserPrincipalName as member")) {
                Add-TenantGroupMember -GroupId $targetGroup.Id -UserId $createdUser.Id
                Write-M365Log "Added user to group '$trimmedGroup' (ID: $($targetGroup.Id))." -Level 'SUCCESS'
                $groupResults += "[$($trimmedGroup): ADDED]"
            } else {
                Write-M365Log "[WHAT-IF] Simulated adding user to group: $trimmedGroup" -Level 'INFO'
                $groupResults += "[$($trimmedGroup): SIMULATED]"
            }
        } catch {
            Write-M365Log "Failed to add user to group '$trimmedGroup': $($_.Exception.Message)" -Level 'WARN'
            $groupResults += "[$($trimmedGroup): ERROR]"
        }
    }
}

# 8. Final Status Evaluation
$finalStatus = if ($Script:Metadata.HasWarnings) {
    "SUCCESS (WITH WARNINGS)"
} else {
    "SUCCESS"
}

$finalLogMessage = "Process completed. Final status: $finalStatus."
Write-M365Log $finalLogMessage -Level 'INFO'

# 9. Console Summary Presentation (Section 7 Specification Format)
$groupsDisplayString = if ($groupResults.Count -gt 0) {
    $groupResults -join ', '
} else {
    "None (not specified)"
}

$licenseDisplayString = if ($licenseAssigned) {
    "$resolvedSkuPartNumber (Enforced UsageLocation: $UsageLocation)"
} else {
    "$resolvedSkuPartNumber (LICENSE ASSIGNMENT FAILED)"
}

$departmentDisplay = if (-not [string]::IsNullOrWhiteSpace($Department)) {
    $Department
} else {
    "Not specified"
}

$summaryLines = @(
    "================================================================================",
    "                    USER PROVISIONING SUMMARY",
    "================================================================================",
    ("UserPrincipalName    : " + $UserPrincipalName),
    ("DisplayName          : " + $DisplayName),
    ("UserId (GUID)        : " + $createdUser.Id),
    ("Department           : " + $departmentDisplay),
    ("Role                 : Standard User (no privileged directory roles)"),
    ("License              : " + $licenseDisplayString),
    ("Force Password Change: Yes (at next sign-in)"),
    ("Teams Groups         : " + $groupsDisplayString),
    ("Log File             : " + $Script:LogFilePath),
    ("Final Status         : " + $finalStatus),
    "================================================================================"
)

Write-Host ""
Write-Host ($summaryLines -join "`n") -ForegroundColor Green
Write-Host ""

exit 0
