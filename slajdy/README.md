# Automatyczne pobieranie slajdów z CMS

Ten skrypt loguje się do wewnętrznego systemu CMS za pomocą CAS, pobiera wybrane prezentacje (slajdy), zapisuje je jako pliki `.zip`, rozpakowuje i kopiuje do folderów sieciowych. Na początku usuwa poprzednie wersje, aby zachować tylko aktualne dane.

## Główne funkcje

- Logowanie do CMS z użyciem CAS (Central Authentication Service)
- Pobranie linków do prezentacji na podstawie identyfikatorów (WNID)
- Pobranie archiwów `.zip` z prezentacjami
- Rozpakowanie tylko właściwych folderów (`tpvision/`)
- Kopiowanie zawartości do katalogów docelowych na serwerach plików
- Usuwanie poprzednich katalogów przed aktualizacją
- Tworzenie logów w pliku `log.txt`

## Wymagania

- Python 3.x
- Pakiety:
  - `requests`
  - `lxml`
- Uprawnienia do zapisów na udziałach sieciowych (SMB):
  - `\\192.168.0.20\common\slajdy\aktualne`
  - `\\192.168.168.4\wspolny\slajdy\aktualne`

## Jak używać

1. Uruchom skrypt z poziomu terminala:

```bash
python aktualizacja_slajdow.py
```

2. Skrypt automatycznie:
   - Usunie wszystkie stare katalogi z prezentacjami
   - Zaloguje się do CMS
   - Pobierze prezentacje o WNID: 45, 92, 1815
   - Rozpakowuje pliki ZIP i umieści je w odpowiednich folderach

## Logi

Wszystkie operacje są zapisywane do pliku:

```
\\192.168.0.20\common\slajdy\log.txt
```

Log zawiera datę, godzinę, etap działania i ewentualne błędy.

## Uwagi

- Skrypt działa na lokalnej sieci wewnętrznej i nie będzie działać poza organizacją
- Pliki ZIP muszą zawierać strukturę `sites/default/files/tpvision/`
- Przed uruchomieniem upewnij się, że inne procesy nie korzystają z katalogów docelowych

## Autor

Łukasz Wybraniec  
E-mail: wybraniec.lukasz@gmail.com
