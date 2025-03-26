# Zestaw skryptów PowerShell do analizy i raportowania uprawnień NTFS

Ten projekt zawiera dwa powiązane skrypty PowerShell, które razem tworzą kompletny raport o uprawnieniach NTFS w systemie plików oraz generują gotowy raport Excel z możliwością potwierdzeń i komentarzy.

---

## 1. Skrypt: lista_katalogów_i_uprawnienia.ps1

**Opis:**
Pierwszy skrypt analizuje strukturę folderów na wybranym dysku (domyślnie `W:\`) i zapisuje szczegółowe informacje o użytkownikach oraz grupach z przypisanymi uprawnieniami NTFS do plików `.csv`.

**Funkcje:**
- Pobiera ACL dla każdego katalogu i przypisuje uproszczone poziomy dostępu (np. FullControl, Read/Write).
- Rozpoznaje konta domenowe i członków grup AD.
- Obsługuje wyjątki i błędy (np. brak modułu AD).
- Tworzy jeden plik `.csv` dla każdego analizowanego folderu.

**Parametry:**
- `DrivePath` – lokalizacja katalogów do analizy (domyślnie `W:\`)
- `OutputFolder` – katalog wyjściowy na pliki `.csv` (domyślnie `C:\Uprawnienia`)

---

## 2. Skrypt: sklej.ps1

**Opis:**
Drugi skrypt przekształca wcześniej wygenerowane pliki `.csv` w jeden raport Excel (`.xlsx`), gotowy do przeglądu i zatwierdzeń. Zawiera walidacje danych, formatowanie warunkowe i zabezpieczenia.

**Funkcje:**
- Łączy pliki `.csv` w jeden skoroszyt Excel z wieloma arkuszami.
- Dodaje kolumny "Potwierdzam" i "Zmiana", z walidacją danych (TAK/NIE/ZMIANA i R/RW/RWM).
- Koloruje wiersze warunkowo według statusu.
- Zabezpiecza arkusze hasłem i umożliwia tylko wybrane działania (filtrowanie, sortowanie).
- Tworzy też osobne pliki `.xlsx` dla każdego folderu.

**Parametry:**
- `CsvFolder` – ścieżka do folderu z plikami CSV
- `ExcelFile` – ścieżka docelowa dla zbiorczego pliku `.xlsx`

**Wymaga modułu PowerShell:** [ImportExcel](https://github.com/dfinke/ImportExcel)

---

## Przykładowe użycie

**Krok 1:** Uruchom analizę uprawnień:

```powershell
.\EksportUprawnien.ps1 -DrivePath "W:\dzialy" -OutputFolder "C:\Uprawnienia"
```

**Krok 2:** Wygeneruj raport Excel:

```powershell
.\GenerujRaportExcel.ps1 -CsvFolder "C:\Uprawnienia" -ExcelFile "C:\Uprawnienia\Uprawnienia.xlsx"
```

---

## Autor

Łukasz Wybraniec  
E-mail: wybraniec.lukasz@gmail.com
