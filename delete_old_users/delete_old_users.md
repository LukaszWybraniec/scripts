# PowerShell – Usuwanie nieaktywnych użytkowników z systemu Roger

Ten skrypt automatycznie usuwa użytkowników z systemu kontroli dostępu Roger, których konta są nieaktywne lub przeterminowane (na podstawie pola `ActivityUntil`). Działa w oparciu o XML z systemu Roger i loguje wszystkie działania do pliku.

## Co robi skrypt

- Loguje się do systemu Roger jako operator `ADMIN`
- Pobiera pełną listę użytkowników w formacie XML
- Sprawdza pole `ActivityUntil` każdego użytkownika
- Jeśli data zakończenia aktywności jest wcześniejsza niż podana przez użytkownika liczba dni – usuwa konto automatycznie
- Jeśli konto nieaktywne (`Active=False`) i `ActivityUntil` to data domyślna – zapyta użytkownika o decyzję
- Dla każdego usuniętego użytkownika zapisuje szczegóły do pliku logu (imię, nazwisko, numer RCP, grupa, kody kart)

## Wymagania

- Windows z zainstalowanym systemem Roger i dostępem do komponentu COM: `PRMaster.PRMasterAutomation`
- Uprawnienia administratora
- PowerShell 5.1 lub nowszy
- Katalog logów: `D:\Roger_del_logs` (tworzony automatycznie)
- Uprawniony login do systemu Roger (domyślnie `ADMIN` z pustym hasłem)

## Jak uruchomić

1. Otwórz PowerShell jako administrator
2. Uruchom skrypt:
3. Podaj liczbę dni, jako okres graniczny, po upływie którego użytkownik ma zostać usunięty

Skrypt przetworzy listę użytkowników i wykona operacje usunięcia, jeśli zostaną spełnione warunki.

## Logi

Plik logu tworzony jest w katalogu D:\Roger_del_logs, nazwa zawiera datę i godzinę wykonania, np.:
deleted_users_log_20250326_145930.txt
Każdy wpis zawiera:

    Imię i nazwisko

    GUID użytkownika

    Numer pracownika (Evidence)

    Grupa użytkownika

    Komentarze dodatkowe (Custom1–Custom4)

    Kod karty (HEX i DEC)

    Datę usunięcia
    
## Uwagi

    Skrypt jest zgodny ze strukturą danych XML zwracaną przez Roger PR Master

    Użytkownicy, którzy są już oznaczeni jako usunięci (Deleted=True), są pomijani

    System wymaga interakcji tylko w przypadku użytkowników nieaktywnych bez określonej daty zakończenia

    Warto dodać mechanizm archiwizacji logów lub integrację z systemem powiadomień w środowisku produkcyjnym
