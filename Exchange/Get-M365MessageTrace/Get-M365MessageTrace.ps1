<#
.SYNOPSIS
    Śledzi przepływ wiadomości (Message Trace) oraz zarządza kwarantanną w Exchange Online.

.DESCRIPTION
    Narzędzie umożliwia:
    1. Śledzenie wiadomości przychodzących i wychodzących (Get-MessageTraceV2).
    2. Wykrywanie wiadomości zablokowanych, oznaczonych jako spam lub zatrzymanych w kwarantannie.
    3. Wyświetlanie unikalnego identyfikatora kwarantanny (Quarantine Identity).
    4. Przywracanie/zwalnianie wiadomości z kwarantanny bezpośrednio do skrzynki odbiorcy.

.PARAMETER User
    Adres e-mail sprawdzanej skrzynki (np. sekretariat@elektrimont.pl).

.PARAMETER Days
    Liczba dni wstecz do przeszukania (od 1 do 10). Domyślnie: 7.

.PARAMETER Direction
    Kierunek przepływu poczty: 'Inbound' lub 'Outbound'. Domyślnie: 'Inbound'.

.PARAMETER OnlyBlockedOrSpam
    Filtruje wyniki wyłącznie do wiadomości zablokowanych, spamu lub kwarantanny.

.PARAMETER ReleaseId
    Identyfikator wiadomości w kwarantannie (Identity z kolumny QuarantineId), którą chcesz zwolnić.

.PARAMETER Help
    Wyświetla pomoc i przykłady użycia.

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -User "sekretariat@elektrimont.pl" -OnlyBlockedOrSpam

.EXAMPLE
    .\Get-M365MessageTrace.ps1 -ReleaseId "c9e782e4-xxxx-xxxx-xxxx-xxxxxxxxxxxx\00000000-0000-0000-0000-000000000000"

.NOTES
    Author:  Roman Pindela
    Email:   roman.pindela@gmail.com
    GitHub:  https://github.com/romanpindela
    Version: 2.0.1
#>

[CmdletBinding(DefaultParameterSetName = "Trace")]
param(
    [Parameter(Mandatory = $false, Position = 0, ParameterSetName = "Trace")]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$User,

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [ValidateRange(1, 10)]
    [int]$Days = 7,

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [ValidateSet("Inbound", "Outbound")]
    [string]$Direction = "Inbound",

    [Parameter(Mandatory = $false, ParameterSetName = "Trace")]
    [switch]$OnlyBlockedOrSpam,

    [Parameter(Mandatory = $true, ParameterSetName = "Release")]
    [string]$ReleaseId,

    [Parameter(Mandatory = $false)]
    [Alias("h")]
    [switch]$Help
)

function Show-ScriptHelp {
    Clear-Host
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host "  M365 Message Trace & Quarantine Manager - v2.0.1" -ForegroundColor Cyan
    Write-Host "  Author: Roman Pindela (roman.pindela@gmail.com) | github.com/romanpindela" -ForegroundColor DarkGray
    Write-Host "=======================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "OPIS:" -ForegroundColor Yellow
    Write-Host "  Skrypt audytuje ruch pocztowy w Exchange Online oraz pozwala zwalniać wiadomości z kwarantanny."
    Write-Host ""
    Write-Host "SKŁADNIA:" -ForegroundColor Yellow
    Write-Host "  Trace wiadomości:" -ForegroundColor Green
    Write-Host "    .\Get-M365MessageTrace.ps1 -User <email> [-Days <1-10>] [-Direction <Inbound|Outbound>] [-OnlyBlockedOrSpam]"
    Write-Host "  Zwolnienie z kwarantanny:" -ForegroundColor Green
    Write-Host "    .\Get-M365MessageTrace.ps1 -ReleaseId <QuarantineIdentity>"
    Write-Host ""
    Write-Host "PARAMETRY:" -ForegroundColor Yellow
    Write-Host "  -User               Docelowa skrzynka pocztowa."
    Write-Host "  -Days               Liczba dni wstecz (1-10). Domyślnie: 7."
    Write-Host "  -Direction          Kierunek: 'Inbound' lub 'Outbound'. Domyślnie: 'Inbound'."
    Write-Host "  -OnlyBlockedOrSpam  Filtruje tylko błędy, spam i kwarantannę."
    Write-Host "  -ReleaseId          Zwalnia z kwarantanny wiadomość o wskazanym Identity."
    Write-Host "  -Help, -h           Wyświetla niniejszą pomoc."
    Write-Host "=======================================================================" -ForegroundColor Cyan
}

if ($Help -or ($PSCmdlet.ParameterSetName -eq "Trace" -and [string]::IsNullOrWhiteSpace($User))) {
    Show-ScriptHelp
    Exit 0
}

# 1. Sprawdzenie połączenia z Exchange Online
Write-Host "[*] Sprawdzanie aktywnej sesji Exchange Online..." -ForegroundColor Cyan
$exoSession = Get-ConnectionInformation | Where-Object { $_.Name -like "*ExchangeOnline*" }

if (-not $exoSession) {
    Write-Host "[!] Brak aktywnej sesji. Rozpoczynanie logowania administratora..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "[+] Pomyślnie nawiązano połączenie z Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Error "[-] Błąd autoryzacji do Exchange Online: $_"
        Exit 1
    }
} else {
    Write-Host "[+] Wykryto istniejącą sesję Exchange Online." -ForegroundColor Green
}

# 2. Obsługa akcji: ZWOLNIENIE WIADOMOŚCI Z KWARANTANNY
if ($PSCmdlet.ParameterSetName -eq "Release") {
    Write-Host "[*] Próba zwolnienia wiadomości z kwarantanny..." -ForegroundColor Cyan
    Write-Host "    Target Identity: $ReleaseId" -ForegroundColor Gray
    try {
        Release-QuarantineMessage -Identity $ReleaseId -ReleaseToAll -Confirm:$false -ErrorAction Stop
        Write-Host "[+] Sukces: Wiadomość została pomyślnie zwolniona i przekazana do skrzynki odbiorcy!" -ForegroundColor Green
    }
    catch {
        Write-Error "[-] Błąd podczas zwalniania wiadomości: $_"
        Write-Host "[!] Upewnij się, czy identyfikator Identity jest poprawny oraz czy wiadomość nie wygasła." -ForegroundColor Yellow
        Exit 1
    }
    Exit 0
}

# 3. Obsługa akcji: ŚLEDZENIE WIADOMOŚCI (MESSAGE TRACE)
$startDate = (Get-Date).AddDays(-$Days)
$endDate = Get-Date

$traceParams = @{
    StartDate = $startDate
    EndDate   = $endDate
}

if ($Direction -eq "Inbound") {
    $traceParams["RecipientAddress"] = $User
    $displayTarget = "SenderAddress"
    $targetHeader = "Sender"
} else {
    $traceParams["SenderAddress"] = $User
    $displayTarget = "RecipientAddress"
    $targetHeader = "Recipient"
}

Write-Host "[*] Pobieranie śladu wiadomości ($Direction) dla '$User' (ostatnie $Days dni)..." -ForegroundColor Cyan

try {
    $results = Get-MessageTraceV2 @traceParams -ErrorAction Stop
}
catch {
    Write-Error "[-] Błąd podczas wykonywania Get-MessageTraceV2: $_"
    Exit 1
}

if (-not $results -or $results.Count -eq 0) {
    Write-Host "[!] Nie znaleziono żadnych wpisów dla $User w wybranym okresie." -ForegroundColor Yellow
    Exit 0
}

# 4. Filtrowanie zdarzeń niepożądanych / spamu
if ($OnlyBlockedOrSpam) {
    Write-Host "[*] Filtrowanie zdarzeń: zablokowane, błędy, spam i kwarantanna..." -ForegroundColor Cyan
    $spamKeywords = @("Quarantined", "Failed", "FilteredAsSpam", "Blocked", "Spam")

    $results = $results | Where-Object {
        $status = $_.Status
        $matched = $false
        foreach ($k in $spamKeywords) {
            if ($status -like "*$k*") {
                $matched = $true
                break
            }
        }
        $matched
    }

    if (-not $results -or $results.Count -eq 0) {
        Write-Host "[+] Czysto: Brak zablokowanych wiadomości ani spamu." -ForegroundColor Green
        Exit 0
    }
}

# 5. Pobieranie metadanych z kwarantanny dla skorelowania wiadomości
$quarantineMap = @{}
if ($Direction -eq "Inbound") {
    Write-Host "[*] Sprawdzanie kwarantanny EOP dla odbiorcy '$User'..." -ForegroundColor Cyan
    try {
        $quarantineItems = Get-QuarantineMessage -RecipientAddress $User -StartReceivedDate $startDate -EndReceivedDate $endDate -PageSize 1000 -ErrorAction SilentlyContinue
        if ($quarantineItems) {
            foreach ($q in $quarantineItems) {
                if ($q.MessageId) {
                    $quarantineMap[$q.MessageId] = $q.Identity
                }
                if ($q.NetworkMessageId) {
                    $quarantineMap[$q.NetworkMessageId] = $q.Identity
                }
            }
            Write-Host "[+] Znaleziono $($quarantineItems.Count) wiadomości w kwarantannie." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "[!] Nie udało się pobrać szczegółów kwarantanny: $_"
    }
}

# 6. Prezentacja wyników
Write-Host "[+] Znaleziono $($results.Count) pasujących zdarzeń:" -ForegroundColor Green

$index = 1
$formattedResults = foreach ($item in $results) {
    $qId = $null
    if ($item.MessageId -and $quarantineMap.ContainsKey($item.MessageId)) {
        $qId = $quarantineMap[$item.MessageId]
    } elseif ($item.NetworkMessageId -and $quarantineMap.ContainsKey($item.NetworkMessageId)) {
        $qId = $quarantineMap[$item.NetworkMessageId]
    }

    [PSCustomObject]@{
        "#"            = $index++
        "Received"     = $item.Received
        $targetHeader  = $item.$displayTarget
        "Subject"      = if ($item.Subject.Length -gt 45) { $item.Subject.Substring(0, 42) + "..." } else { $item.Subject }
        "Status"       = $item.Status
        "InQuarantine" = if ($qId) { "TAK" } elseif ($item.Status -like "*Quarantined*") { "TAK (szukaj)" } else { "NIE" }
        "QuarantineId" = if ($qId) { $qId } else { "-" }
    }
}

# Główna tabela
$formattedResults | Format-Table -Property "#", "Received", $targetHeader, "Subject", "Status", "InQuarantine" -AutoSize

# 7. Wyświetlenie listy wiadomości z kwarantanny
$releasable = $formattedResults | Where-Object { $_.QuarantineId -ne "-" }

if ($releasable) {
    Write-Host "
=======================================================================" -ForegroundColor Yellow
    Write-Host " WIADOMOŚCI MOŻLIWE DO ZWOLNIENIA Z KWARANTANNY" -ForegroundColor Yellow
    Write-Host "=======================================================================" -ForegroundColor Yellow

    foreach ($r in $releasable) {
        Write-Host "[$($r.'#')] Od: $($r.$targetHeader) | Temat: $($r.Subject)" -ForegroundColor White
        Write-Host "    Identity: $($r.QuarantineId)" -ForegroundColor DarkCyan
        Write-Host "    Aby przywrócić wykonaj:" -ForegroundColor Gray
        Write-Host "    .\Get-M365MessageTrace.ps1 -ReleaseId "$($r.QuarantineId)"" -ForegroundColor Green
        Write-Host ""
    }
} else {
    Write-Host "
[i] Brak bezpośrednio zmapowanych obiektów w kwarantannie." -ForegroundColor Gray
    Write-Host "    Jeśli status to 'FilteredAsSpam', wiadomość trafiła bezpośrednio do folderu Wiadomości-śmieci w skrzynce użytkownika." -ForegroundColor Gray
}
