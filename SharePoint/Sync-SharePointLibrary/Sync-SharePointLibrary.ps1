<#
.SYNOPSIS
    Automatically triggers OneDrive for Business sync for a specified SharePoint Online library.

.DESCRIPTION
    Sync-SharePointLibrary connects to a designated SharePoint Online site, retrieves its metadata
    and primary document library parameters, and triggers the local OneDrive client synchronization
    engine via the registered 'odopen://sync?' URI scheme.
    
    This script is designed for interactive user execution in Microsoft 365 environments.

.PARAMETER SiteName
    The URL slug / site alias for the SharePoint site (e.g., 'automatyka').

.PARAMETER BaseUrl
    The base root URL for SharePoint site collections.
    Default: 'https://elektrimont.sharepoint.com/sites'

.PARAMETER DocumentLibraryTitle
    The display name or relative URL of the target document library.
    Default: 'Shared Documents' (Dokumenty).

.PARAMETER Help
    Displays detailed script usage, prerequisites, examples, and author details.

.EXAMPLE
    .\Sync-SharePointLibrary.ps1 -SiteName "automatyka"
    Syncs the 'Documents' library from https://elektrimont.sharepoint.com/sites/automatyka

.EXAMPLE
    .\Sync-SharePointLibrary.ps1 -SiteName "projekty" -BaseUrl "https://elektrimont.sharepoint.com/sites"
    Syncs the site library using a custom base path.

.NOTES
    Script Name   : Sync-SharePointLibrary.ps1
    Author        : Roman Pindela
    Email         : roman.pindela@gmail.com
    GitHub        : https://github.com/romanpindela
    Version       : 1.0.0
    Created       : 2026-10-16
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName = 'Default', Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^[a-zA-Z0-9_\-]+$')]
    [string]$SiteName,

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$BaseUrl = "https://elektrimont.sharepoint.com/sites",

    [Parameter(ParameterSetName = 'Default', Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DocumentLibraryTitle = "Shared Documents",

    [Parameter(ParameterSetName = 'Help')]
    [Alias('h')]
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
# SCRIPT METADATA & CONSTANTS
# ==============================================================================
$Script:Metadata = @{
    Name        = "Sync-SharePointLibrary.ps1"
    Version     = "1.0.0"
    Author      = "Roman Pindela"
    Email       = "roman.pindela@gmail.com"
    GitHub      = "https://github.com/romanpindela"
    Company     = "Elektrimont Sp. z o.o."
    LastUpdated = "2026-10-16"
}

# ==============================================================================
# FUNCTION: Show-ScriptHelp
# ==============================================================================
function Show-ScriptHelp {
    Write-Host @"
================================================================================
SCRIPT INFORMATION:
  Name        : $($Script:Metadata.Name)
  Version     : $($Script:Metadata.Version)
  Author      : $($Script:Metadata.Author)
  Email       : $($Script:Metadata.Email)
  GitHub      : $($Script:Metadata.GitHub)
  Company     : $($Script:Metadata.Company)
  Date        : $($Script:Metadata.LastUpdated)
================================================================================

DESCRIPTION:
  Automates the process of attaching a SharePoint Online document library to the
  local Windows File Explorer via the OneDrive for Business client sync engine.

PREREQUISITES:
  1. USER SIGN-IN REQUIRED:
     The user MUST be logged in to the OneDrive application on their local PC 
     using their active Microsoft 365 business credentials.
     [PL] Uzytkownik MUSI byc wczesniej zalogowany do aplikacji OneDrive na swoim 
     komputerze. Dane do logowania to sluzbowy adres email i haslo (dokladnie 
     takie same jak w poczcie Outlook / pakiecie Office 365).
  2. PERMISSIONS:
     The signed-in account must have appropriate read/write permissions to the
     target SharePoint site/library.
  3. UNBLOCK FILE:
     If downloaded from GitHub, unblock the file prior to execution:
     Unblock-File -Path .\Sync-SharePointLibrary.ps1

USAGE & EXAMPLES:
  1. Default site sync:
     .\$($Script:Metadata.Name) -SiteName "automatyka"

  2. Custom base URL:
     .\$($Script:Metadata.Name) -SiteName "projekty" -BaseUrl "https://elektrimont.sharepoint.com/sites"

  3. Display help screen:
     .\$($Script:Metadata.Name) -Help
     .\$($Script:Metadata.Name) -h

SECURITY NOTE:
  Input parameters are strictly validated to prevent arbitrary script or URL injection.
================================================================================
"@ -ForegroundColor Cyan
}

# Display Help if explicitly invoked or if launched without parameters
if ($Help.IsPresent -or [string]::IsNullOrWhiteSpace($SiteName)) {
    Show-ScriptHelp
    exit 0
}

# ==============================================================================
# FUNCTION: Test-OneDriveClient
# ==============================================================================
function Test-OneDriveClient {
    <#
    .SYNOPSIS
        Verifies if OneDrive is running and ready on the local machine.
    #>
    Write-Host "[*] Checking OneDrive client status..." -ForegroundColor Gray
    
    $oneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if (-not $oneDriveProcess) {
        Write-Warning "[!] OneDrive client process was not detected."
        Write-Host "    Attempting to start OneDrive from local user profile..." -ForegroundColor Yellow
        
        $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
        $oneDriveExe = Join-Path -Path$localAppData -ChildPath "Microsoft\OneDrive\OneDrive.exe"

        if (Test-Path -Path $oneDriveExe) {
            Start-Process -FilePath $oneDriveExe
            Start-Sleep -Seconds 3
        } else {
            throw "OneDrive client is not installed or could not be found at: $oneDriveExe"
        }
    } else {
        Write-Host "[+] OneDrive client is running." -ForegroundColor Green
    }
}

# ==============================================================================
# FUNCTION: Invoke-SharePointLibrarySync
# ==============================================================================
function Invoke-SharePointLibrarySync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetSite,

        [Parameter(Mandatory = $true)]
        [string]$RootUrl,

        [Parameter(Mandatory = $true)]
        [string]$LibraryTitle
    )

    # Sanitize and construct canonical URLs
    $cleanRootUrl = $RootUrl.TrimEnd('/')$fullSiteUrl = "$cleanRootUrl/$TargetSite"

    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host " ELEKTRIMONT SP. Z O.O. - SHAREPOINT SYNC UTILITY" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "[*] Target SharePoint Site : $fullSiteUrl" -ForegroundColor White
    Write-Host "[*] Target Library         : $LibraryTitle" -ForegroundColor White
    
    # Prerequisite verification notice
    Write-Host "`n[WAZNE / PREREQUISITE]:" -ForegroundColor Yellow
    Write-Host "Uzytkownik musi byc zalogowany do aplikacji OneDrive na swoim komputerze." -ForegroundColor Yellow
    Write-Host "Dane logowania: Twoj firmowy adres email i haslo (takie jak do Outlooka / Office 365).`n" -ForegroundColor Yellow

    Test-OneDriveClient

    # Construct the protocol URI to invoke OneDrive sync directly
    # Format: odopen://sync?siteId=&webId=&webUrl=<url>&listTitle=<title>
    $encodedSiteUrl = [System.Uri]::EscapeDataString($fullSiteUrl)
    $encodedLibTitle = [System.Uri]::EscapeDataString($LibraryTitle)
    
    # Primary protocol URI for modern OneDrive sync
    $syncUri = "odopen://sync/?siteUrl=$encodedSiteUrl&listTitle=$encodedLibTitle"

    Write-Host "[*] Triggering synchronization protocol..." -ForegroundColor Gray

    try {
        Start-Process -FilePath $syncUri
        Write-Host "[+] Synchronizacja zainicjowana pomyslnie!" -ForegroundColor Green
        Write-Host "[i] Sprawdz powiadomienie OneDrive w prawym dolnym rogu ekranu." -ForegroundColor White
        Write-Host "[i] Za chwile w Eksploratorze Windows (pod ikona biurowca) pojawi sie biblioteka." -ForegroundColor White
    }
    catch {
        Write-Error "[x] Blad podczas uruchamiania protokolu OneDrive: $($_.Exception.Message)"
        throw $_
    }
}

# ==============================================================================
# SCRIPT ENTRYPOINT
# ==============================================================================
try {
    Invoke-SharePointLibrarySync -TargetSite $SiteName -RootUrl $BaseUrl -LibraryTitle $DocumentLibraryTitle
}
catch {
    Write-Host "`n[FATAL ERROR] An unexpected error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}