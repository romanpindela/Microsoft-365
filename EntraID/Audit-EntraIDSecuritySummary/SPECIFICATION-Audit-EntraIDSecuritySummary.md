# Specyfikacja Techniczna: Audit-EntraIDSecuritySummary

## 1. Metadane i lokalizacja projektu
- **Ścieżka docelowa repozytorium:** `c:\GitHub\Microsoft-365\EntraID\Audit-EntraIDSecuritySummary\`
- **Nazwa pliku skryptu:** `Audit-EntraIDSecuritySummary.ps1`
- **Typ narzędzia:** Audyt bezpieczeństwa / Raportowanie stanu konfiguracji (Read-Only)
- **Kompatybilność:** PowerShell 7.2+ (zalecane) oraz Windows PowerShell 5.1

---

## 2. Cel i zakres działania
Skrypt przeprowadza audyt kluczowych ustawień bezpieczeństwa w dzierżawie (tenancie) Microsoft Entra ID bez wprowadzania jakichkolwiek modyfikacji (tryb tylko do odczytu). Weryfikuje bieżące konfiguracje względem standardów CIS Microsoft 365 Foundations Benchmark oraz dobrych praktyk Microsoftu, a następnie generuje interaktywny, czytelny raport HTML zawierający:
- Wynik sprawdzenia (Status: Zgodny / Niezgodny / Ostrzeżenie).
- Proste wyjaśnienie biznesowe i techniczne ("Dlaczego to ustawienie jest istotne?").
- Konkretne rekomendacje naprawcze (kroki w portalu Entra ID lub polecenia PowerShell).
- Macierz podsumowującą działania według priorytetu akcji (Krytyczny, Wysoki, Średni, Niski).

---

## 3. Zakres weryfikowanych obszarów Entra ID
Skrypt odpytuje środowisko za pośrednictwem modułu **Microsoft.Graph** (lub bezpośrednio przez REST Graph API w sesji aplikacji/użytkownika) w następujących obszarach:

1. **Ustawienia użytkowników i rejestracji aplikacji:**
   - Czy standardowi użytkownicy mogą rejestrować aplikacje firm trzecich (`Users can register applications`).
   - Czy standardowi użytkownicy mogą tworzyć grupy zabezpieczeń lub grupy M365.
   - Ograniczenia dostępu do portalu administracyjnego Entra ID dla użytkowników bez ról (`Restrict access to Microsoft Entra admin center`).
   - Uprawnienia użytkowników do wyrażania zgody na aplikacje (User Consent Settings).

2. **Dostęp gości (Guest Access Settings):**
   - Poziom ograniczeń uprawnień dla kont gości (`GuestUserRoleId`: Restricted vs Default).
   - Uprawnienia do zapraszania gości (czy ograniczone wyłącznie do administratorów).

3. **Autentykacja i metody logowania:**
   - Stan wdrożenia Security Defaults (jeśli brak licencji P1/P2).
   - Rejestracja SSPR (Self-Service Password Reset) oraz dozwolone metody resetu hasła.
   - Rejestracja weryfikacji wieloskładnikowej (MFA) dla kont z rolami uprzywilejowanymi (Global Administrator, Security Administrator, Privileged Role Administrator).

4. **Polityki Dostępów Warunkowych (Conditional Access - jeśli dostępne licencje P1/P2):**
   - Weryfikacja obecności aktywnej polityki wymuszającej MFA dla wszystkich administratorów.
   - Weryfikacja blokady starszych protokołów uwierzytelniania (Legacy Authentication).
   - Weryfikacja wymuszenia MFA dla logowań ryzykownych (jeśli dostępny Entra ID P2).

5. **Role administracyjne i konta awaryjne:**
   - Liczba kont z przypisaną rolą Global Administrator (rekomendacja: min. 2 konta Break-Glass, max. 4 konta).
   - Wykrycie kont Global Administrator bez włączonego MFA.

---

## 4. Architektura parametrów wejściowych

```powershell
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$PSScriptRoot\Reports",

    [Parameter(Mandatory = $false)]
    [switch]$OpenReport,

    [Parameter(Mandatory = $false)]
    [string]$TenantId
)
```

---

## 5. Wymagane uprawnienia (Microsoft Graph Scopes)
Skrypt wymaga wyłącznie uprawnień typu **Read-Only**:
- `Policy.Read.All`
- `Directory.Read.All`
- `RoleManagement.Read.Directory`
- `Policy.Read.ConditionalAccess` (dla analizy Conditional Access)

---

## 6. Struktura raportu HTML (`AuditReport-EntraID-[Data].html`)
Raport generowany jest jako autonomiczny, pojedynczy plik HTML zawierający wbudowane style CSS oraz prosty JavaScript do filtrowania i zwijania sekcji:

1. **Header & Executive Summary:**
   - Data wykonania, ID Tenanta, całkowita liczba sprawdzonych reguł.
   - Wskaźnik bezpieczeństwa (Score Card: % zgodności, liczba podatności wg wag: Critical / High / Medium / Low).
2. **Karty Rekomendacji (Szczegóły):**
   - **Tytuł reguły:** Krótka nazwa ustawienia.
   - **Stan obecny:** Wykryta wartość w dzierżawie.
   - **Rekomendacja:** Pożądana wartość.
   - **Dlaczego to ważne:** Zrozumiałe uzasadnienie ryzyka (np. ochrona przed przejęciem tożsamości, zapobieganie eskalacji uprawnień przez consent phishing).
   - **Kroki naprawcze:** Ścieżka w portalu Entra ID.
3. **Macierz Akcji (Action Summary):**
   - Tabela posortowana wg pilności wdrożenia (Quick Wins vs zmiany architektoniczne).

---

## 7. Wymagania implementacyjne i Clean Code
- **Brak operacji zapisu (`Read-Only`):** Skrypt pod żadnym pozorem nie wykonuje poleceń `Set-*`, `Update-*` ani `New-*` na obiektach Entra ID.
- **Obsługa błędów:** Każde zapytanie Graph API opakowane w blok `try {} catch {}` z eleganckim logowaniem braku licencji lub braku uprawnień (np. pominięcie analizy Conditional Access, gdy tenant nie posiada licencji Entra ID P1/P2).
- **Formatowanie kodu:** Zgodność z regułami PSScriptAnalyzer, jawne deklaracje typów, PascalCase dla funkcji i camelCase dla zmiennych lokalnych.