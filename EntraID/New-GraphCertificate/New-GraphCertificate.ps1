<#
.SYNOPSIS
    Generator certyfikatów X.509 dla skryptu Upload-LargeFilesSession.ps1
.DESCRIPTION
    Tworzy samopodpisany certyfikat RSA 2048-bit, instaluje go w wybranym magazynie 
    (CurrentUser / LocalMachine), eksportuje klucz publiczny (.cer) do Entra ID 
    w bezpiecznej lokalizacji oraz wyświetla odciski palca (Thumbprint).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CertificateName = "M365-Graph-OneDriveUploader",

    [Parameter(Mandatory = $false)]
    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$StoreLocation = 'CurrentUser',

    [Parameter(Mandatory = $false)]
    [string]$ExportDirectory = "C:\M365_Certs",

    [Parameter(Mandatory = $false)]
    [int]$ValidityYears = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "         GENERATOR CERTYFIKATÓW X.509 DLA MICROSOFT GRAPH API" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 1. Sprawdzanie i przygotowanie bezpiecznej lokalizacji eksportu
if (-not (Test-Path -Path $ExportDirectory)) {
    New-Item -ItemType Directory -Path $ExportDirectory -Force | Out-Null
    Write-Host "[+] Utworzono bezpieczny katalog eksportu: $ExportDirectory" -ForegroundColor Green
}

# Zabezpieczenie ACL katalogu (Tylko Administratorzy / Autor mają dostęp)
try {
    $acl = Get-Acl $ExportDirectory
    $acl.SetAccessRuleProtection($true, $true) # Wyłączenie dziedziczenia
    Set-Acl -Path $ExportDirectory -AclObject $acl
} catch {
    Write-Host "[!] Ostrzeżenie: Nie udało się w pełni skonfigurować ACL katalogu: $_" -ForegroundColor Yellow
}

# 2. Generowanie certyfikatu
Write-Host "[*] Tworzenie certyfikatu $CertificateName w magazynie $StoreLocation..." -ForegroundColor Cyan

$subjectName = "CN=$CertificateName"
$certStore = "Cert:\$StoreLocation\My"

$cert = New-SelfSignedCertificate -Subject $subjectName `
    -CertStoreLocation $certStore `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($ValidityYears)

$thumbprint = $cert.Thumbprint
Write-Host "[+] Certyfikat pomyślnie wygenerowany!" -ForegroundColor Green
Write-Host "    - Thumbprint: $thumbprint" -ForegroundColor Yellow
Write-Host "    - Magazyn:    $certStore" -ForegroundColor Yellow

# 3. Eksport klucza publicznego (.cer) do Entra ID
$cerFileName = "$CertificateName.cer"
$cerFilePath = Join-Path -Path $ExportDirectory -ChildPath $cerFileName
Export-Certificate -Cert $cert -FilePath $cerFilePath -Type CERT | Out-Null

Write-Host "[+] Wyeksportowano klucz publiczny do Entra ID:" -ForegroundColor Green
Write-Host "    -> $cerFilePath" -ForegroundColor Cyan

# 4. Podsumowanie i instrukcja dalszych kroków
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "PODSUMOWANIE WDROŻENIA:" -ForegroundColor White
Write-Host "1. Wgraj plik .cer do Entra ID -> App registrations -> Certificates & secrets." -ForegroundColor Gray
Write-Host "2. Uruchom skrypt docelowy za pomocą parametru:" -ForegroundColor Gray
Write-Host "   .\Upload-LargeFilesSession.ps1 -CertificateThumbprint '$thumbprint'" -ForegroundColor Yellow
Write-Host "================================================================================" -ForegroundColor Cyan