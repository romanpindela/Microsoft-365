<#
.SYNOPSIS
    Automates the provisioning of a new user identity in Microsoft Entra ID / Microsoft 365.

.DESCRIPTION
    New-M365User.ps1 provisions a new user account in Microsoft Entra ID / Microsoft 365,
    configures a secure initial password with mandatory change at next sign-in
    (forceChangePasswordNextSignIn = $true), assigns a specified license plan
    (supporting both friendly SkuPartNumber and GUID SkuId), enforces ISO 3166-1
    country usage location, assigns standard user privileges (no privileged roles),
    and optionally adds the user to designated Microsoft 365 / Microsoft Teams groups.

    The script provides real-time color-coded console telemetry, structured file
    logging (C:\Logs\Microsoft-365\New-M365User_YYYY-MM-DD.log), simulation mode
    via -WhatIf / -Confirm, and tenant license capacity validation.

.PARAMETER UserPrincipalName
    The UserPrincipalName / primary email address of the new user (e.g. jan.kowalski@domena.pl).
    Mandatory parameter.

.PARAMETER DisplayName
    The full display name of the user (e.g. Jan Kowalski). Mandatory parameter.

.PARAMETER GivenName
    The first name of the user (e.g. Jan). Optional.

.PARAMETER Surname
    The last name / surname of the user (e.g. Kowalski). Optional.

.PARAMETER MailNickname
    The mail alias / nickname of the user. If omitted, it is automatically derived
    from the portion of the UserPrincipalName preceding the '@' sign.

.PARAMETER UsageLocation
    Two-letter ISO 3166-1 alpha-2 country code required by Microsoft 365 for license assignment.
    Defaults to 'PL'.

.PARAMETER InitialPassword
    The initial account password as a [SecureString]. Passwords in plain text are strictly
    rejected to comply with enterprise credential security standards. Mandatory parameter.

.PARAMETER LicenseSku
    The license SKU to assign. Accepts friendly SkuPartNumber (e.g. 'ENTERPRISEPACK', 'SPE_E3',
    'O365_BUSINESS_PREMIUM') or direct SkuId GUID. Mandatory parameter.

.PARAMETER TeamGroupNames
    One or more Microsoft 365 / Teams group display names to which the new user should be added.
    Accepts string array or comma-separated list. Optional.

.PARAMETER LogDirectory
    Path to the directory where operational audit logs will be saved.
    Defaults to 'C:\Logs\Microsoft-365'.

.PARAMETER WhatIf
    Simulates script execution without making changes in Entra ID or Teams.

.PARAMETER Confirm
    Prompts for confirmation before executing state-changing operations.

.EXAMPLE
    $secPass = "P@ssw0rd2026!" | ConvertTo-SecureString -AsPlainText -Force
    .\New-M365User.ps1 -UserPrincipalName "jan.kowalski@domena.pl" -DisplayName "Jan Kowalski" -InitialPassword $secPass -LicenseSku "SPE_E3"
    Provisions a new user with Microsoft 365 E3 license and PL usage location.

.EXAMPLE
    $secPass = Read-Host "Wprowadz haslo poczatkowe" -AsSecureString
    .\New-M365User.ps1 -UserPrincipalName "anna.nowak@domena.pl" `
                       -DisplayName "Anna Nowak" `
                       -GivenName "Anna" `
                       -Surname "Nowak" `
                       -InitialPassword $secPass `
                       -LicenseSku "ENTERPRISEPACK" `
                       -TeamGroupNames "Dzial Handlowy", "Ogolny"
    Provisions a user with Office 365 E3 and automatically adds them to two Teams groups.

.EXAMPLE
    $secPass = "Startowe123!" | ConvertTo-SecureString -AsPlainText -Force
    .\New-M365User.ps1 -UserPrincipalName "tomasz.wisniewski@domena.pl" `
                       -DisplayName "Tomasz Wisniewski" `
                       -InitialPassword $secPass `
                       -LicenseSku "c707743b-110d-45b0-802c-1c83b2162703" `
                       -WhatIf
    Performs a dry-run simulation of user provisioning and license assignment.

.NOTES
    Script: New-M365User.ps1
    Version: 1.0.0
    Author: Roman Pindela
    Email: roman.pindela@gmail.com
    GitHub: https://github.com/romanpindela
    Required Modules: Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Groups
    Required Scopes: User.ReadWrite.All, Directory.ReadWrite.All, Group.ReadWrite.All
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "UserPrincipalName / primary email address of the new user.")]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Full display name of the new user.")]
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName,

    [Parameter(Mandatory = $false, HelpMessage = "First name of the user.")]
    [string]$GivenName,

    [Parameter(Mandatory = $false, HelpMessage = "Last name / surname of the user.")]
    [string]$Surname,

    [Parameter(Mandatory = $false, HelpMessage = "Mail alias / nickname. Defaults to prefix of UPN.")]
    [string]$MailNickname,

    [Parameter(Mandatory = $false, HelpMessage = "Two-letter ISO 3166-1 alpha-2 country code for licensing (e.g. PL).")]
    [ValidatePattern('^[A-Za-z]{2}$')]
    [string]$UsageLocation = 'PL',

    [Parameter(Mandatory = $true, HelpMessage = "Initial password as a SecureString.")]
    [ValidateNotNull()]
    [securestring]$InitialPassword,

    [Parameter(Mandatory = $true, HelpMessage = "License SkuPartNumber (e.g. SPE_E3) or SkuId GUID.")]
    [ValidateNotNullOrEmpty()]
    [string]$LicenseSku,

    [Parameter(Mandatory = $false, HelpMessage = "Microsoft 365 / Teams group display names to add user to.")]
    [string[]]$TeamGroupNames = @(),

    [Parameter(Mandatory = $false, HelpMessage = "Directory path for operational audit logs.")]
    [string]$LogDirectory = 'C:\Logs\Microsoft-365'
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

# ==============================================================================
# SCRIPT METADATA & TRACKING
# ==============================================================================
$Script:Metadata = @{
    Name        = 'New-M365User'
    Version     = '1.0.0'
    Author      = 'Roman Pindela'
    StartTime   = Get-Date
    LogFileName = "New-M365User_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd')
    HasWarnings = $false
}

$Script:LogFilePath = Join-Path -Path $LogDirectory -ChildPath $Script:Metadata.LogFileName

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
    # Format level tag with standard width matching specification:
    # [INFO]    (10 chars), [SUCCESS] (10 chars), [WARN]    (10 chars), [ERROR]   (10 chars)
    $levelTag = "[$Level]".PadRight(9) + " "
    $logEntry = "[$timestamp] $levelTag$Message"

    # Console color mapping
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

    # Append to file
    try {
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        [System.IO.File]::AppendAllLines($Script:LogFilePath, [string[]]@($logEntry), [System.Text.Encoding]::UTF8)
    } catch {
        Write-Warning "Nie mozna zapisac logu do '$Script:LogFilePath': $($_.Exception.Message)"
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
# MAIN PROVISIONING LOGIC
# ==============================================================================
Write-M365Log "Inicjalizacja tworzenia uzytkownika: $UserPrincipalName" -Level 'INFO'

# 1. Validation of UPN format
$emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
if ($UserPrincipalName -notmatch $emailRegex) {
    Write-M365Log "Niepoprawny format UserPrincipalName: '$UserPrincipalName'." -Level 'ERROR'
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
    Write-M365Log "Brak aktywnej sesji Microsoft Graph. Polacz sie uzywajac Connect-MgGraph." -Level 'ERROR'
    Write-Host "`nAby nawiazac sesje z wymaganymi uprawnieniami, wykonaj:" -ForegroundColor Yellow
    Write-Host "    Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All'`n" -ForegroundColor White
    exit 1
}

Write-M365Log "Wykryto aktywna sesje Graph dla tenanta: $($mgContext.TenantId) (Konto: $($mgContext.Account))" -Level 'DEBUG'

# Check required scopes
$requiredScopes = @('User.ReadWrite.All', 'Directory.ReadWrite.All', 'Group.ReadWrite.All')
$activeScopes = @()
if ($mgContext.Scopes) {
    $activeScopes = $mgContext.Scopes
}

$missingScopes = @()
foreach ($reqScope in $requiredScopes) {
    if ($activeScopes -notcontains $reqScope) {
        $missingScopes += $reqScope
    }
}

if ($missingScopes.Count -gt 0) {
    Write-M365Log "Aktywna sesja Graph nie posiada wszystkich zalecanych uprawnien (brak: $($missingScopes -join ', ')). Niektore operacje moga zwrocic 403 Forbidden." -Level 'WARN'
}

# 3. Validate and resolve License SKU in Tenant
Write-M365Log "Weryfikacja dostepnosci pakietu licencyjnego: $LicenseSku..." -Level 'INFO'
$resolvedSkuId = $null
$resolvedSkuPartNumber = $null

try {
    $allSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    $targetSku = $allSkus | Where-Object {
        $_.SkuPartNumber -ieq $LicenseSku -or "$($_.SkuId)" -ieq $LicenseSku
    } | Select-Object -First 1

    if ($null -eq $targetSku) {
        Write-M365Log "Pakiet licencyjny '$LicenseSku' nie zostal odnaleziony w dzierzawie." -Level 'ERROR'
        Write-Host "`nDostepne pakiety licencyjne w biezacym tenancie:" -ForegroundColor Yellow
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
        Write-M365Log "Wszystkie jednostki licencji $resolvedSkuPartNumber zostaly juz wykorzystane ($consumedUnits/$totalUnits). Przypisanie licencji moze zakonczyc sie niepowodzeniem." -Level 'WARN'
    } else {
        Write-M365Log "Rozpoznano licencje: $resolvedSkuPartNumber ($resolvedSkuId). Wolne jednostki: $availableUnits/$totalUnits." -Level 'DEBUG'
    }
} catch {
    Write-M365Log "Blad podczas pobierania licencji dzierzawy: $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

# 4. Check if User already exists in Entra ID
Write-M365Log "Sprawdzanie unikalnosci UserPrincipalName: $UserPrincipalName..." -Level 'DEBUG'
try {
    $existingUser = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
    if ($null -ne $existingUser) {
        Write-M365Log "Uzytkownik o UserPrincipalName '$UserPrincipalName' juz istnieje w dzierzawie (UserId: $($existingUser.Id)). Operacja przerwana." -Level 'ERROR'
        exit 1
    }
} catch {
    Write-M365Log "Blad podczas weryfikacji istnienia konta: $($_.Exception.Message)" -Level 'ERROR'
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

try {
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Utworzenie konta uzytkownika Entra ID (Standard User)")) {
        $createdUser = New-MgUser @userParams -ErrorAction Stop
        Write-M365Log "Konto uzytkownika utworzone (ID: $($createdUser.Id))" -Level 'SUCCESS'
    } else {
        # WhatIf / Simulation mock
        $createdUser = [pscustomobject]@{
            Id                = '00000000-0000-0000-0000-000000000000'
            DisplayName       = $DisplayName
            UserPrincipalName = $UserPrincipalName
        }
        Write-M365Log "[WHAT-IF] Symulacja utworzenia uzytkownika: $UserPrincipalName (ID mock: $($createdUser.Id))" -Level 'INFO'
    }
} catch {
    Write-M365Log "Blad krytyczny podczas tworzenia uzytkownika '$UserPrincipalName': $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}

# 6. Assign License
Write-M365Log "Przypisywanie licencji: $resolvedSkuPartNumber..." -Level 'INFO'
$licenseAssigned = $false

try {
    $licenseAdd = @{
        SkuId = [guid]$resolvedSkuId
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Przypisanie licencji $resolvedSkuPartNumber ($resolvedSkuId)")) {
        Set-MgUserLicense -UserId $createdUser.Id -AddLicenses @($licenseAdd) -RemoveLicenses @() -ErrorAction Stop
        Write-M365Log "Licencja przypisana pomyslnie." -Level 'SUCCESS'
        $licenseAssigned = $true
    } else {
        Write-M365Log "[WHAT-IF] Symulacja przypisania licencji: $resolvedSkuPartNumber" -Level 'INFO'
        $licenseAssigned = $true
    }
} catch {
    Write-M365Log "Nie udalo sie przypisac licencji do konta ($($createdUser.Id)): $($_.Exception.Message)" -Level 'WARN'
    $licenseAssigned = $false
}

# 7. Add User to Microsoft 365 / Teams Groups
$groupResults = @()

if ($TeamGroupNames -and $TeamGroupNames.Count -gt 0) {
    foreach ($groupName in $TeamGroupNames) {
        $trimmedGroup = $groupName.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedGroup)) { continue }

        Write-M365Log "Wyszukiwanie grupy Teams/M365: '$trimmedGroup'..." -Level 'DEBUG'
        try {
            # Filter groups by displayName
            $escapedGroupName = $trimmedGroup.Replace("'", "''")
            $foundGroups = Get-MgGroup -Filter "displayName eq '$escapedGroupName'" -ErrorAction Stop

            if ($null -eq $foundGroups -or ($foundGroups | Measure-Object).Count -eq 0) {
                Write-M365Log "Grupa '$trimmedGroup' nie zostala znaleziona w dzierzawie." -Level 'WARN'
                $groupResults += "[$($trimmedGroup): NIE ZNALEZIONO]"
                continue
            }

            # Select target group
            $targetGroup = $foundGroups | Select-Object -First 1

            if ($PSCmdlet.ShouldProcess("$trimmedGroup ($($targetGroup.Id))", "Dodanie uzytkownika $UserPrincipalName jako czlonka")) {
                $memberOdataId = "https://graph.microsoft.com/v1.0/directoryObjects/$($createdUser.Id)"
                
                # Check cmdlet parameter compatibility across Graph SDK versions
                $byRefCmd = Get-Command New-MgGroupMemberByRef -ErrorAction SilentlyContinue
                if ($null -ne $byRefCmd -and $byRefCmd.Parameters.ContainsKey('OdataId')) {
                    New-MgGroupMemberByRef -GroupId $targetGroup.Id -OdataId $memberOdataId -ErrorAction Stop | Out-Null
                } else {
                    $bodyParam = @{
                        "@odata.id" = $memberOdataId
                    }
                    New-MgGroupMemberByRef -GroupId $targetGroup.Id -BodyParameter $bodyParam -ErrorAction Stop | Out-Null
                }

                Write-M365Log "Dodano uzytkownika do grupy '$trimmedGroup' (ID: $($targetGroup.Id))." -Level 'SUCCESS'
                $groupResults += "[$($trimmedGroup): DODANY]"
            } else {
                Write-M365Log "[WHAT-IF] Symulacja dodania uzytkownika do grupy: $trimmedGroup" -Level 'INFO'
                $groupResults += "[$($trimmedGroup): SYMULACJA]"
            }
        } catch {
            Write-M365Log "Nie udalo sie dodac uzytkownika do grupy '$trimmedGroup': $($_.Exception.Message)" -Level 'WARN'
            $groupResults += "[$($trimmedGroup): BLAD]"
        }
    }
}

# 8. Final Status Evaluation
$finalStatus = if ($Script:Metadata.HasWarnings) {
    "SUKCES (Z OSTRZEZENIAMI)"
} else {
    "SUKCES"
}

$finalLogMessage = "Zakonczenie procesu. Status koncowy: $finalStatus."
Write-M365Log $finalLogMessage -Level 'INFO'

# 9. Console Summary Presentation (Section 7 Specification Format)
$groupsDisplayString = if ($groupResults.Count -gt 0) {
    $groupResults -join ', '
} else {
    "Brak (nie wskazano)"
}

$licenseDisplayString = if ($licenseAssigned) {
    "$resolvedSkuPartNumber (Wymuszone UsageLocation: $UsageLocation)"
} else {
    "$resolvedSkuPartNumber (BLAD PRZYPISANIA)"
}

$summaryLines = @(
    "================================================================================",
    "                    PODSUMOWANIE PROVISIONINGU UZYTKOWNIKA",
    "================================================================================",
    ("UserPrincipalName    : " + $UserPrincipalName),
    ("DisplayName          : " + $DisplayName),
    ("UserId (GUID)        : " + $createdUser.Id),
    ("Rola                 : Standard User (brak rol uprzywilejowanych)"),
    ("Licencja             : " + $licenseDisplayString),
    ("Wymuszenie hasla     : Tak (przy pierwszym logowaniu)"),
    ("Grupy Teams          : " + $groupsDisplayString),
    ("Plik Logu            : " + $Script:LogFilePath),
    ("Status Koncowy       : " + $finalStatus),
    "================================================================================"
)

Write-Host ""
Write-Host ($summaryLines -join "`n") -ForegroundColor Green
Write-Host ""

exit 0
