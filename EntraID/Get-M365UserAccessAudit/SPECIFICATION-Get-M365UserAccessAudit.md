# Specyfikacja Techniczna: Get-M365UserAccessAudit

## 1. Metadane Projektu i Lokalizacja
* **Ścieżka bazowa repozytorium:** `c:\GitHub\Microsoft-365\EntraID\`
* **Katalog projektu (`[NAZWA_KATALOGU]`):** `Get-M365UserAccessAudit`
* **Pełna ścieżka docelowa:** `c:\GitHub\Microsoft-365\EntraID\Get-M365UserAccessAudit\`
* **Główny plik skryptu:** `Get-M365UserAccessAudit.ps1`
* **Plik specyfikacji:** `SPECIFICATION-Get-M365UserAccessAudit.md`
* **Plik dokumentacji:** `README.md`

---

## 2. Cel i Opis Działania Skryptu
Skrypt `Get-M365UserAccessAudit.ps1` służy do kompleksowego, szczegółowego audytu uprawnień oraz członkostw w grupach dla wskazanego użytkownika (lub listy użytkowników) w środowisku Microsoft 365 / Entra ID. 

### Zakres audytu obejmuje:
1. **Entra ID – Grupy (Bezpośrednie i Przechodnie / Transitive):**
   * Grupy zabezpieczeń (*Security Groups* – Cloud oraz On-Premises Synced).
   * Grupy Microsoft 365 (*Unified Groups*).
   * Grupy przypisywalne do ról (*Role-assignable groups*, `isAssignableToRole -eq $true`).
   * Identyfikacja typu członkostwa: `Direct` vs `Nested/Transitive` oraz `Static` vs `Dynamic` (`groupTypes` zawiera `DynamicMembership`).
2. **Entra ID – Role Administracyjne (Directory Roles):**
   * Role przypisane bezpośrednio do użytkownika.
   * Role odziedziczone poprzez członkostwo w grupach zabezpieczeń.
   * (Opcjonalnie via Graph API) Aktywne i kwalifikujące się (*Eligible*) przypisania w ramach **Privileged Identity Management (PIM)**.
3. **Microsoft Teams:**
   * Wykaz zespołów (Teams) powiązanych z grupami M365 użytkownika.
   * Rola w zespole (`Owner` / `Member`) oraz weryfikacja członkostwa w kanałach prywatnych i współdzielonych (*Private / Shared Channels*).
4. **SharePoint Online (w kontekście grup i witryn):**
   * Powiązane adresy URL witryn (`SharePoint Site URL`) dla grup Microsoft 365, do których należy użytkownik.
   * Mapowanie grup zabezpieczeń Entra ID na uprawnienia wewnątrz witryn SharePoint Online (weryfikacja na podstawie grupy powiązanej lub skanowania wybranych witryn/zestawień z Graph API).
5. **Exchange Online:**
   * Grupy dystrybucyjne (*Distribution Lists*), dynamiczne grupy dystrybucyjne oraz grupy zabezpieczeń z włączoną obsługą poczty (*Mail-Enabled Security Groups*).
   * Uprawnienia do skrzynek pocztowych nadane bezpośrednio lub przez grupy zabezpieczeń (`FullAccess`, `SendAs`, `SendOnBehalf`).

---

## 3. Wymagania Techniczne i Architektura
* **Wersja PowerShell:** Zalecany **PowerShell 7.2+** (wsparcie dla równoległości i nowoczesnych modułów) z zachowaniem kompatybilności wstecznej dla **Windows PowerShell 5.1**.
* **Wymagane moduły zewnętrzne:**
  * `Microsoft.Graph.Authentication`
  * `Microsoft.Graph.Users`
  * `Microsoft.Graph.Groups`
  * `Microsoft.Graph.Identity.DirectoryManagement`
  * `Microsoft.Graph.Teams` / `Microsoft.Graph.Sites`
  * `ExchangeOnlineManagement` (opcjonalnie/wymagane dla głębokiego audytu skrzynek i grup dystrybucyjnych EXO).
* **Wymagane uprawnienia (Least Privilege - Read-Only):**
  * **Microsoft Graph Scopes:** `User.Read.All`, `Group.Read.All`, `Directory.Read.All`, `RoleManagement.Read.Directory`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `Sites.Read.All`.
  * **Exchange Online:** Rola *View-Only Recipients* lub *Global Reader*.
* **Elewacja UAC po SID:** `NIE` (skrypt operuje na API chmurowym M365, nie wymaga lokalnych uprawnień administratora stacji roboczej).
* **Obsługa `-WhatIf` (`SupportsShouldProcess`):** `NIE` (skrypt wykonuje wyłącznie operacje odczytu `Get-*` i generowania raportów).
* **Pasek postępu (`Write-Progress`):** `TAK` – wielopoziomowy pasek postępu dla poszczególnych faz audytu (Entra ID -> Role -> Teams -> SharePoint -> Exchange -> Eksport).
* **Moduł logowania:** `TAK` – pełne logowanie przebiegu oraz błędów do katalogu `C:\Logs\Get-M365UserAccessAudit\`.

---

## 4. Parametry Wejściowe (`param()`)

| Parametr | Typ | Wymagany | Domyślnie | Opis |
| :--- | :--- | :---: | :---: | :--- |
| `-UserPrincipalName` | `string[]` | Tak | — | Jeden lub więcej adresów UPN (lub ObjectId) użytkowników podlegających audytowi (obsługa potoku `ValueFromPipeline`). |
| `-IncludeExchangePermissions` | `switch` | Nie | `$false` | Rozszerza audyt o skanowanie uprawnień do skrzynek (`FullAccess`, `SendAs`, `SendOnBehalf`) w Exchange Online (może wydłużyć czas działania). |
| `-IncludeSharePointSiteAudit` | `switch` | Nie | `$false` | Sprawdza powiązane witryny SharePoint Online oraz weryfikuje uprawnienia grup zabezpieczeń w witrynach. |
| `-IncludePIM` | `switch` | Nie | `$false` | Sprawdza przypisania ról *Eligible* w Entra ID Privileged Identity Management (wymaga `RoleEligibilitySchedule.Read.Directory`). |
| `-ExportFormat` | `string[]` | Nie | `@('CSV', 'HTML')` | Format wyjściowy raportu (`ValidateSet('CSV', 'HTML', 'JSON')`). |
| `-OutputPath` | `string` | Nie | `C:\Logs\Get-M365UserAccessAudit\Reports` | Ścieżka zapisu wygenerowanych raportów audytowych. |
| `-LogPath` | `string` | Nie | `C:\Logs\Get-M365UserAccessAudit` | Katalog docelowy dla plików logów `.log`. |
| `-TenantId` | `string` | Nie | `$null` | Identyfikator dzierżawy (Tenant ID), jeśli wymagane jest wymuszenie konkretnego tenanta podczas logowania. |

---

## 5. Architektura Logowania i Obsługi Błędów

### Standard Logowania (`Write-Log`)
* Automatyczne tworzenie katalogu `C:\Logs\Get-M365UserAccessAudit\` przy starcie skryptu (jeśli nie istnieje).
* Nazwa pliku logu: `Get-M365UserAccessAudit_YYYYMMDD_HHMMSS.log`.
* Format wpisu: `[YYYY-MM-DD HH:mm:ss] [LEVEL] [KOMPONENT] Wiadomość` (gdzie `LEVEL` to `INFO`, `WARN`, `ERROR`, `SUCCESS`, `DEBUG`).
* Równoczesny zapis do pliku (`Add-Content -Encoding UTF8`) oraz kolorowane wyjście na konsolę (`Write-Host` / `Write-Verbose`).

### Obsługa Błędów
* Globalne wymuszenie `$ErrorActionPreference = 'Stop'` wewnątrz bloków `try / catch`.
* Weryfikacja połączenia z Microsoft Graph (`Get-MgContext`) oraz Exchange Online (`Get-ConnectionInformation`) przed rozpoczęciem pobierania danych.
* Bezpieczna obsługa brakujących licencji (np. brak P2 przy odpytywaniu o PIM logowany jako `WARN`, bez przerywania całego skryptu).

---

## 6. Przepływ Logiki Skryptu (Workflow)

1. **Inicjalizacja (`begin`):**
   * Utworzenie struktur katalogowych dla logów (`C:\Logs\...`) i raportów.
   * Weryfikacja zainstalowanych modułów (`Microsoft.Graph.*`, `ExchangeOnlineManagement`).
   * Sprawdzenie aktywnej sesji MS Graph i brakujących zakresów uprawnień (`Scopes`). W razie braku – wywołanie `Connect-MgGraph`.
   * Jeśli użyto `-IncludeExchangePermissions`, weryfikacja lub nawiązanie sesji `Connect-ExchangeOnline -ShowBanner:$false`.

2. **Zbieranie Danych dla Użytkownika (`process`):**
   * **Krok 1 (Tożsamość):** Pobranie profilu użytkownika (`Get-MgUser`) – `Id`, `DisplayName`, `UserPrincipalName`, `AccountEnabled`, `Department`, `JobTitle`, `OnPremisesSyncEnabled`.
   * **Krok 2 (Grupy Bezpośrednie vs Przechodnie):**
     * Pobranie grup bezpośrednich (`Get-MgUserMemberOf`).
     * Pobranie wszystkich grup przechodnich (`Get-MgUserTransitiveMemberOf`).
     * Pobranie grup, których użytkownik jest właścicielem (`Get-MgUserOwnedObject`).
     * Porównanie kolekcji w celu oznaczenia ścieżki dostępu: `Direct Member`, `Nested (Transitive) Member`, `Owner`.
   * **Krok 3 (Klasyfikacja Grup i Kontekst Usług):**
     * Dla każdej grupy określenie jej typu:
       * `Entra Security Group` (`SecurityEnabled = true`, `MailEnabled = false`)
       * `Mail-Enabled Security Group` (`SecurityEnabled = true`, `MailEnabled = true`, brak `Unified`)
       * `Microsoft 365 Group` (`GroupTypes` zawiera `Unified`)
       * `Distribution List` (`SecurityEnabled = false`, `MailEnabled = true`)
     * Określenie flagi `IsAssignableToRole` oraz reguły członkostwa (`Dynamic` vs `Assigned`) i źródła (`Cloud` vs `On-Premises AD`).
   * **Krok 4 (Kontekst Entra ID Roles):**
     * Pobranie ról bezpośrednich użytkownika oraz ról przypisanych do grup, w których użytkownik się znajduje (`Get-MgRoleManagementDirectoryRoleAssignment` z filtrem po `PrincipalId` użytkownika oraz `PrincipalId` jego grup).
     * Jeśli włączono `-IncludePIM`, odpytanie `Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance`.
   * **Krok 5 (Kontekst Microsoft Teams):**
     * Identyfikacja grup M365 posiadających aprowizowany zespół Teams (`ResourceProvisioningOptions -contains 'Team'`).
     * Pobranie kanałów prywatnych/współdzielonych, w których użytkownik ma jawne uprawnienia.
   * **Krok 6 (Kontekst SharePoint Online):**
     * Dla każdej grupy M365 pobranie powiązanego adresu witryny SharePoint via Graph API (`Get-MgGroupSite -GroupId <Id> -SiteId 'root'`).
     * Jeśli użyto `-IncludeSharePointSiteAudit`, zestawienie grup zabezpieczeń użytkownika z uprawnieniami na poziomie kolekcji witryn.
   * **Krok 7 (Kontekst Exchange Online):**
     * Zmapowanie grup mailowych i dystrybucyjnych.
     * Jeśli użyto `-IncludeExchangePermissions`, weryfikacja uprawnień delegowanych na skrzynkach (bezpośrednio dla UPN oraz poprzez SID/Tożsamość grup Mail-Enabled Security, do których należy użytkownik).

3. **Agregacja i Generowanie Raportu (`end`):**
   * Zbudowanie zunifikowanej kolekcji obiektów `[PSCustomObject]` prezentującej pełną matrycę uprawnień.
   * Eksport do pliku `.csv` (kodowanie UTF-8 z BOM dla pełnej kompatybilności z MS Excel i polskimi znakami).
   * Wygenerowanie interaktywnego i czytelnego raportu `.html` z podsumowaniem KPI (liczba grup Security, M365, ról Entra ID, zespołów Teams, witryn SPO) oraz tabelą szczegółową.
   * Zamknięcie logu i wyświetlenie podsumowania w konsoli.

---

## 7. Struktura Obiektu Wyjściowego (Data Schema)

Każdy wiersz w wygenerowanym raporcie będzie reprezentował pojedyncze uprawnienie / członkostwo w ustandaryzowanym formacie:

* `AuditTimestamp` : Data i czas wykonania audytu.
* `UserPrincipalName` : UPN badanego użytkownika.
* `UserDisplayName` : Nazwa wyświetlana użytkownika.
* `AccountEnabled` : Status konta (`True`/`False`).
* `Workload` : Kontekst usługi (`Entra ID`, `SharePoint Online`, `Microsoft Teams`, `Exchange Online`).
* `ResourceType` : Typ zasobu (`Security Group`, `M365 Group`, `Distribution Group`, `Directory Role`, `Team`, `SPO Site`, `Mailbox`).
* `ResourceName` : Nazwa grupy, roli, zespołu, witryny lub skrzynki.
* `ResourceId` : `ObjectId` w Entra ID lub unikalny identyfikator zasobu.
* `AccessLevel` : Poziom uprawnień (np. `Member`, `Owner`, `Global Reader`, `FullAccess`, `SendAs`, `Site Member`).
* `AssignmentMethod` : Sposób nadania (`Direct`, `Nested / Group Inherited`, `Dynamic Rule`, `PIM Eligible`).
* `GrantedViaGroup` : Nazwa grupy pośredniczącej (jeśli uprawnienie wynika z członkostwa w grupie zabezpieczeń/M365, w przeciwnym razie `Direct`).
* `GroupMembershipType` : `Assigned (Static)` / `Dynamic` / `N/A`.
* `IsRoleAssignable` : `True` / `False` (czy grupa może mieć przypisywane role Entra ID).
* `SyncSource` : `Cloud` / `On-Premises AD`.
* `ResourceUrl` : Adres URL zasobu (np. adres witryny SharePoint Online) lub puste pole.

---

## 8. Kryteria Akceptacji dla Wdrożenia (Checklist dla Antigravity)
- [ ] Utworzenie katalogu `c:\GitHub\Microsoft-365\EntraID\Get-M365UserAccessAudit\`.
- [ ] Zapisanie niniejszej specyfikacji w pliku `SPECIFICATION-Get-M365UserAccessAudit.md`.
- [ ] Implementacja skryptu `Get-M365UserAccessAudit.ps1` zgodnie z najlepszymi praktykami PowerShell (Comment-Based Help, `CmdletBinding`, silne typowanie parametrów, modularne funkcje pomocnicze).
- [ ] Zastosowanie paginacji (`-All`) przy zapytaniach `Microsoft.Graph` w celu uniknięcia limitu 100 rekordów.
- [ ] Wykorzystanie zapytań wsadowych lub cache'owania słownikowego (`Hashtable`) dla grup i ról, aby zminimalizować liczbę zapytań do Graph API (ochrona przed *throttlingiem* HTTP 429).
- [ ] Przygotowanie dokumentacji użytkownika w pliku `README.md` (zawierającej opis wymaganych uprawnień, przykłady użycia oraz opis kolumn raportu).