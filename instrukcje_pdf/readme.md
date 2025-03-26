# Graficzny interfejs do wyszukiwania i otwierania plików PDF (Monitor ERP)

Ten skrypt uruchamia aplikację GUI (Tkinter), która umożliwia wyszukiwanie i otwieranie plików PDF powiązanych z artykułami w systemie Monitor ERP. Dane są pobierane z bazy danych SQL Anywhere.

## Główne funkcje

- Pobieranie danych z bazy `monitor.db` (Monitor ERP)
- Wyszukiwanie po numerze artykułu
- Wyświetlanie listy znalezionych plików i ich otwieranie
- Kolorowe oznaczenia plików w zależności od ich wieku:
  - **Zielony** – plik starszy niż 14 dni
  - **Czerwony** – plik nowszy niż 14 dni
- Powiadomienie dźwiękowe (syrena) w przypadku wykrycia nowych plików
- Możliwość zmiany lub usunięcia zapisanego hasła do bazy (Keyring)

## Wymagania

- Python 3.x
- Zainstalowane biblioteki:
  - `sqlanydb`
  - `keyring`
  - `tkinter`
  - `Pillow`
  - `pygame`
- Dostęp do lokalnej bazy danych `monitor.db` (SQL Anywhere)
- Pliki graficzne (`g.jpg`) i dźwiękowe (`siren.mp3`) w katalogu roboczym

## Jak używać

1. Przy pierwszym uruchomieniu zostaniesz poproszony o hasło do bazy danych – zostanie ono zapisane lokalnie przez `keyring`.
2. Wprowadź numer artykułu w pole tekstowe i naciśnij Enter lub kliknij „Wyszukaj”.
3. Kliknij przycisk z nazwą pliku, aby go otworzyć.
4. Kliknij „Wyczyść”, aby rozpocząć nowe wyszukiwanie.
5. Kliknij „Usuń hasło”, aby usunąć zapisane hasło i wpisać nowe.

## Logowanie

Wszystkie operacje są logowane do pliku `PDFlog.txt` w katalogu domowym użytkownika.

## Uwagi

- Hasło do bazy jest przechowywane lokalnie przez `keyring` i można je łatwo zmienić.
- Ścieżki do plików PDF są pobierane z kolumn `svag_path` i `fil_namn`.
- W przypadku błędów (brak połączenia, błędna ścieżka, brak pliku) użytkownik otrzyma komunikat.

## Autor

Łukasz Wybraniec  
E-mail: wybraniec.lukasz@gmail.com
