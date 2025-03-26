# PowerShell – Rejestracja nowego pracownika do systemów firmy

Skrypt automatyzuje dodawanie nowego pracownika na podstawie danych z pliku Word (obiegówki). Dane są wczytywane, a następnie przekazywane do systemu kontroli dostępu Roger oraz bazy danych MySQL. Skrypt działa z interaktywnym menu w PowerShell.

## Główne funkcje

- Wczytywanie danych pracownika z pliku `.docx`
- Odczyt karty dostępu z czytnika (DEC/HEX/24-bit)
- Przypisanie pracownika do wybranej grupy w systemie Roger
- Dodanie użytkownika do Rogera z kartą dostępu
- Wpisanie numeru karty do pliku Word
- Zapis danych do bazy MySQL (raporty wydajności)

## Wymagania

- System Windows z zainstalowanym MS Word
- Zainstalowany moduł PowerShell: `MySQL`
- Uprawnienia administratora
- Dostęp do katalogu sieciowego z plikami Word:  
  `\\192.168.0.20\common\Karty_obiegowe`
- Połączenie z bazą MySQL na serwerze `192.168.0.21`

## Struktura danych

Plik Word musi zawierać tabelę z danymi pracownika (imię, nazwisko, nr pracownika, stanowisko, firma).  
Dane są odczytywane z konkretnych komórek tabeli.

## Menu główne

Po uruchomieniu skryptu pojawi się menu:

1: Wczytaj kartę obiegową 2: Wczytaj kartę dostępu 3: Wybierz grupę dostępu 4: Dodaj użytkownika do Rogera 5: Dodaj użytkownika do raportów wydajności 9: Wykonaj cały proces (1–4) Q: Wyjście


## Uwagi techniczne

- Skrypt używa COM do komunikacji z MS Word i oprogramowaniem Roger
- Dane są przekazywane jako XML do systemu Roger
- Przed dodaniem użytkownika sprawdzane są duplikaty numeru pracownika i karty
- Po wykonaniu operacji uruchamiane jest wysyłanie konfiguracji do urządzeń

## Bezpieczeństwo

Skrypt wymaga poświadczeń do bazy danych i ma dostęp do lokalnych systemów. Powinien być uruchamiany tylko przez uprawnionych administratorów.

## Autor

Łukasz Wybraniec  
E-mail: wybraniec.lukasz@gmail.com
