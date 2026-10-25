# Specyfikacja Techniczna: Get-SharePointSitesAndUsers

- **Lokalizacja docelowa repozytorium:** `C:\GitHub\Microsoft-365\SharePoint\`
- **Katalog skryptu:** `C:\GitHub\Microsoft-365\SharePoint\Get-SharePointSitesAndUsers\`
- **Główny plik skryptu:** `Get-SharePointSitesAndUsers.ps1`
- **Wersja:** 1.0.0
- **Środowisko uruchomieniowe:** PowerShell 7.2+ (zalecane) / Windows PowerShell 5.1 (kompatybilność wsteczna)
- **Moduły zależne:** `PnP.PowerShell` (zalecany) lub alternatywnie `Microsoft.Online.SharePoint.PowerShell`

---

## 1. Cel biznesowy i zakres działania

Głównym celem skryptu jest audyt i transparentna inwentaryzacja witryn SharePoint Online w środowisku Microsoft 365 wraz z agregacją ról i uprawnień użytkowników. 

Skrypt umożliwia:
- Pobranie zbiorcze wszystkich witryn w dzierżawie (tenant) (`-All`) lub selektywne pobranie wskazanej witryny (`-Site <Url>`).
- Analizę uprawnień na poziomie każdej witryny (Site Collection Administrators, Ownerzy/Full Control, Członkowie/Write-Edit, Goście/Read/Visitors).
- Zwracanie uporządkowanych obiektów typu `[PSCustomObject]` do potoku PowerShell z możliwością eksportu (np. `Export-Csv`, `Out-GridView`).
- Zapisywanie szczegółowego dziennika zdarzeń (logów) w dedykowanym katalogu `C:\Logs\Get-SharePointSitesAndUsers\`.

---

## 2. Architektura parametrów i zestawy parametrów (ParameterSets)

Skrypt wykorzystuje mechanizm `CmdletBinding()` oraz dwa wykluczające się zestawy parametrów (`ParameterSetName`), co zapobiega błędnemu jednoczesnemu podaniu przełącznika `-All` oraz parametru `-Site`.

```powershell
[CmdletBinding(DefaultParameterSetName = 'AllSites')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'AllSites', Position = 0)]
    [switch]$All,

    [Parameter(Mandatory = $true, ParameterSetName = 'SingleSite', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [Alias('Url', 'SiteUrl')]
    [string]$Site,

    [Parameter(Mandatory = $false)]
    [string]$AdminCenterUrl,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Logs\Get-SharePointSitesAndUsers",

    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalSites
)
```

### Mapowanie parametrów:
- `-All` *(switch)*: Wyzwala pobranie pełnej kolekcji witryn z tenanta.
- `-Site` *(string)*: Przyjmuje bezpośredni URL do witryny (np. `https://contoso.sharepoint.com/sites/Marketing`). Obsługuje aliasy `SiteUrl` i `Url` oraz dane z potoku.
- `-AdminCenterUrl` *(string, opcjonalny)*: URL do centrum administracyjnego SharePoint (np. `https://contoso-admin.sharepoint.com`). W przypadku braku, skrypt powinien wyznaczyć go dynamicznie z konfiguracji sesji lub monitować o połączenie.
- `-LogPath` *(string, opcjonalny)*: Ścieżka do zapisu logów audytowych. Domyślnie `C:\Logs\Get-SharePointSitesAndUsers\`.
- `-IncludePersonalSites` *(switch, opcjonalny)*: Domyślnie wyklucza witryny OneDrive for Business (`/personal/`), chyba że flaga zostanie ustawiona.

---

## 3. Schemat uprawnień i logika mapowania ról

Dla każdej analizowanej witryny, skrypt agreguje tożsamości do 4 ustandaryzowanych kategorii:

| Nazwa pola | Poziom logiczny | Grupy źródłowe / Uprawnienia |
|---|---|---|
| `SiteCollectionAdmins` | Administrator / SCA | Site Collection Administrators (`IsSiteAdmin = $true`). |
| `FullControl` / `Owners` | Pełna kontrola / Właściciel | Domyślna grupa `<Site> Owners`, uprawnienie `Full Control`. |
| `Write` / `Members` | Modyfikacja / Zapis | Domyślna grupa `<Site> Members`, role `Edit`, `Contribute`. |
| `Read` / `Visitors` | Odczyt | Domyślna grupa `<Site> Visitors`, rola `Read`. |

*Uwaga: W przypadku grup Microsoft 365 powiązanych z witryną, członkowie i właściciele grupy są mapowani odpowiednio do Write i Owners.*

---

## 4. Struktura zwracanego obiektu wyjściowego (`PSCustomObject`)

Każdy rekord przekazywany do strumienia wyjściowego posiada ujednolicony format, ułatwiający wyświetlanie w tabeli (`Format-Table`) oraz eksport do CSV:

```powershell
[PSCustomObject]@{
    SiteTitle             = $siteTitle           # Nazwa witryny
    SiteUrl               = $siteUrl             # Bezpośredni adres URL
    StorageUsedMB         = $storageMB           # Zużycie przestrzeni (opcjonalne)
    Template              = $template            # Typ szablonu (np. GROUP#0, STS#3)
    SiteCollectionAdmins  = ($admins -join '; ') # Administratorzy zbioru witryn
    FullControl           = ($full -join '; ')   # Pełna kontrola / Właściciele
    Write                 = ($write -join '; ')  # Uprawnienia zapisu / Członkowie
    Read                  = ($read -join '; ')   # Uprawnienia odczytu / Odwiedzający
    AuditDate             = (Get-Date)           # Timestamp badania
}
```

---

## 5. Standard modułu logowania

Dedykowana funkcja wewnętrzna `Write-AuditLog` realizuje zapis synchroniczny do pliku tekstowego oraz kolorowane wyjście na konsolę:

- **Format pliku logu:** `C:\Logs\Get-SharePointSitesAndUsers\Get-SharePointSitesAndUsers_yyyyMMdd_HHmmss.log`
- **Format linii logu:** `[yyyy-MM-dd HH:mm:ss] [INFO/WARN/ERROR/DEBUG] <Treść zdarzenia>`
- **Obsługa katalogu:** Automatyczna weryfikacja obecności folderu i tworzenie za pomocą `New-Item -ItemType Directory -Force`.

---

## 6. Diagram przepływu procesu

```
+-------------------------------------------------------+
| Inicjalizacja skryptu: Walidacja środowiska & logów   |
+-------------------------------------------------------+
                           |
                           v
+-------------------------------------------------------+
| Sprawdzenie połączenia do SharePoint / PnP Context   |
| (Monit o Connect-PnPOnline / Connect-SPOService)      |
+-------------------------------------------------------+
                           |
          +----------------+----------------+
          |                                 |
   [Parametr: -Site]                [Parametr: -All]
          |                                 |
          v                                 v
+-----------------------+       +-----------------------+
| Pobierz wskazaną      |       | Pobierz pełną listę   |
| witrynę               |       | witryn z tenanta      |
+-----------------------+       +-----------------------+
          |                                 |
          +----------------+----------------+
                           |
                           v
+-------------------------------------------------------+
| Iteracja po witrynach + Write-Progress                |
| - Pobranie Site Collection Admins                     |
| - Pobranie domyślnych grup (Owners, Members, Visitors)|
| - Pobranie dedykowanych przypisań RoleAssignments     |
| - Agregacja tożsamości (UserPrincipalName / Mail / Nazwa)|
+-------------------------------------------------------+
                           |
                           v
+-------------------------------------------------------+
| Emisja PSCustomObject do Pipeline                     |
+-------------------------------------------------------+
                           |
                           v
+-------------------------------------------------------+
| Zakończenie pracy i finalizacja pliku logu            |
+-------------------------------------------------------+
```

---

## 7. Wymagania dotyczące obsługi błędów i odporności (Resilience)

1. **Uprawnienia do pojedynczej witryny:** Błąd braku dostępu lub błąd 401/403 do konkretnej witryny w pętli nie może przerywać całego audytu. Skrypt musi przechwycić błąd w bloku `try/catch`, zarejestrować wpis `[WARN]` lub `[ERROR]` do logu i kontynuować kolejną witrynę.
2. **Pasek postępu (`Write-Progress`):** W trybie `-All` skrypt musi aktualizować status przetwarzania z podaniem aktualnego indeksu, łącznej liczby witryn oraz procentu ukończenia.
3. **Throttling (HTTP 429):** Obsługa zapytań o permissions powinna unikać nadmiarowych wywołań API i poprawnie ponawiać próbę w razie chwilowego ograniczenia przepustowości ze strony Microsoft Graph / SPO.

---

## 8. Kryteria akceptacji

- [ ] Skrypt znajduje się pod ścieżką `C:\GitHub\Microsoft-365\SharePoint\Get-SharePointSitesAndUsers\Get-SharePointSitesAndUsers.ps1`.
- [ ] Wywołanie `Get-SharePointSitesAndUsers -Site "https://tenant.sharepoint.com/sites/Test"` zwraca 1 rekord z poprawnie podzielonymi uprawnieniami (FullControl, Write, Read).
- [ ] Wywołanie `Get-SharePointSitesAndUsers -All` przetwarza cały tenant, ignorując witryny typu OneDrive (`/personal/`), chyba że wymuszono `-IncludePersonalSites`.
- [ ] Wynik można swobodnie przekazać do `Format-Table -AutoSize` lub `Export-Csv -Path export.csv -NoTypeInformation`.
- [ ] Utworzony plik logu zawiera pełny ślad wykonania skryptu.