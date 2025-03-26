param (
    [string]$CsvFolder = 'D:\tisax\uprawnienia',
    [string]$ExcelFile = 'D:\tisax\uprawnienia\Uprawnienia.xlsx'
)

# Funkcja pomocnicza: konwersja numeru kolumny na literę (np. 1 -> A, 27 -> AA)
function Get-ExcelColumnLetter {
    param(
        [int]$ColumnNumber
    )
    $letter = ""
    while ($ColumnNumber -gt 0) {
        $mod = ($ColumnNumber - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $ColumnNumber = [math]::Floor(($ColumnNumber - $mod) / 26)
    }
    return $letter
}

# Sprawdzenie modułu ImportExcel (instalacja, jeśli brak)
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Write-Host 'Moduł ImportExcel nie jest zainstalowany. Instalacja...' -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Usuń istniejący zbiorczy plik Excel, jeśli istnieje
if (Test-Path $ExcelFile) {
    Remove-Item $ExcelFile -Force
}

# Pobranie plików CSV z podanego katalogu (z uwzględnieniem ukrytych katalogów)
$csvFiles = Get-ChildItem -Path $CsvFolder -Filter '*.csv' -Force
# Jeśli chcesz przeszukać również podkatalogi, dodaj -Recurse

if ($csvFiles.Count -eq 0) {
    Write-Host "Brak plików CSV w katalogu $CsvFolder!" -ForegroundColor Red
    exit
}

foreach ($csvFile in $csvFiles) {
    Write-Host "Przetwarzanie: $($csvFile.Name)" -ForegroundColor Cyan

    # Import danych z pliku CSV
    $csvData = Import-Csv -Path $csvFile.FullName

    # Filtruj wiersze:
    # 1. Jeśli wiersz ma puste "Grupa / Użytkownik" i "Imię i Nazwisko"
    #    oraz "Uprawnienia" równe "FullControl" (niewrażliwe na wielkość liter) → pomiń.
    # 2. Jeśli w kolumnie "Grupa / Użytkownik" (po trimowaniu i do małych liter)
    #    znajduje się wpis z listy $excluded.
    $excluded = @(
        "builtin\administrators", 
        "klippan\administrator", 
        "klippan\poladm", 
        "klippan\itgrzegorzg", 
        "klippan\itlukaszw",
        "nt authority\system"
    )
    $csvData = $csvData | Where-Object {
        $group = $_.'Grupa / Użytkownik'
        $name  = $_.'Imię i Nazwisko'
        $perm  = $_.'Uprawnienia'
        
        # Warunek dodatkowy: jeśli Grupa i Imię i Nazwisko są puste, a Uprawnienia = FullControl → pomiń
        if (((-not $group) -or ($group.Trim() -eq "")) -and 
            ((-not $name) -or ($name.Trim() -eq "")) -and 
            ($perm -and $perm.Trim().ToLower() -eq "fullcontrol")) {
            return $false
        }
        
        # Sprawdzenie wartości w "Grupa / Użytkownik"
        if ($group) {
            $user = $group.Trim().ToLower()
            if ($excluded -contains $user) {
                return $false
            }
        }
        return $true
    }

    # Jeśli po filtrze zbiór jest pusty – pomiń ten plik
    if (-not $csvData -or $csvData.Count -eq 0) {
        Write-Host "Plik $($csvFile.Name) nie zawiera danych po filtrze – pomijam..."
        continue
    }

    # Dodaj do każdego wiersza dwie nowe kolumny: 'Potwierdzam' oraz 'Zmiana'
    $csvData = $csvData | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name 'Potwierdzam' -Value '' -Force
        $_ | Add-Member -MemberType NoteProperty -Name 'Zmiana' -Value '' -Force
        $_
    }

    # Ustal nazwę arkusza na podstawie nazwy pliku (bez rozszerzenia, max 31 znaków)
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.Name)
    if ($sheetName.Length -gt 31) {
        $sheetName = $sheetName.Substring(0,31)
    }

    # (Opcjonalnie) Pobierz właściciela katalogu – do komunikatu
    $folderPath = 'D:\tisax\' + $sheetName
    $acl = Get-Acl -Path $folderPath -ErrorAction SilentlyContinue
    $owner = if ($acl) { $acl.Owner } else { 'Nieznany' }

    ## EKSPORT DO ZBIORCZEGO PLIKU EXCEL

    # Eksport danych z CSV – nagłówki w wierszu 1, dane od wiersza 2
    $csvData | Export-Excel -Path $ExcelFile -WorksheetName $sheetName -AutoSize -FreezeTopRow -TableStyle Medium2 -StartRow 1

    # Ustal pozycje kolumn – pobierz nagłówki wygenerowane z obiektów CSV
    $headers = $csvData[0].psobject.Properties.Name
    $potwierdzamIndex = [array]::IndexOf($headers, 'Potwierdzam') + 1
    $colLetter = Get-ExcelColumnLetter $potwierdzamIndex

    $zmianaIndex = [array]::IndexOf($headers, 'Zmiana') + 1
    $zmianaColLetter = Get-ExcelColumnLetter $zmianaIndex

    # Nagłówki są w wierszu 1, dane zaczynają się od wiersza 2:
    $dataStartRow = 2
    $dataEndRow = 1 + $csvData.Count
    if ($dataEndRow -lt $dataStartRow) { $dataEndRow = $dataStartRow }
    $rangePotwierdzam = "${colLetter}${dataStartRow}:${colLetter}${dataEndRow}"
    $rangeZmiana = "${zmianaColLetter}${dataStartRow}:${zmianaColLetter}${dataEndRow}"

    # Zakres całej tabeli (od kolumny A do ostatniej)
    $lastColumn = Get-ExcelColumnLetter($headers.Count)
    $tableRange = "A${dataStartRow}:${lastColumn}${dataEndRow}"

    # Otwórz zbiorczy plik Excel, dodaj walidacje, formatowanie warunkowe oraz zabezpiecz arkusz
    $pkg = Open-ExcelPackage -Path $ExcelFile
    $ws = $pkg.Workbook.Worksheets[$sheetName]
    if ($ws -ne $null) {
        ## Walidacja dla kolumny "Potwierdzam": dozwolone wartości TAK, NIE lub ZMIANA
        $dv = $ws.DataValidations.AddListValidation($rangePotwierdzam)
        $dv.Formula.Values.Clear()
        $dv.Formula.Values.Add("TAK")
        $dv.Formula.Values.Add("NIE")
        $dv.Formula.Values.Add("ZMIANA")
        $dv.ShowErrorMessage = $true
        $dv.ErrorTitle = "Nieprawidłowa wartość"
        $dv.Error = "Wybierz wartość TAK, NIE lub ZMIANA"
        $dv.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::Stop

        ## Walidacja dla kolumny "Zmiana": dozwolone wartości R, RW lub RWM
        $dvZmiana = $ws.DataValidations.AddListValidation($rangeZmiana)
        $dvZmiana.Formula.Values.Clear()
        $dvZmiana.Formula.Values.Add("R")
        $dvZmiana.Formula.Values.Add("RW")
        $dvZmiana.Formula.Values.Add("RWM")
        $dvZmiana.ShowErrorMessage = $true
        $dvZmiana.ErrorTitle = "Nieprawidłowa wartość"
        $dvZmiana.Error = "Wybierz wartość R, RW lub RWM"
        $dvZmiana.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::Stop

        ## Formatowanie warunkowe całych wierszy
        $greenFormula = "=INDIRECT(""" + $colLetter + """ & ROW())=""TAK"""
        $cfGreen = $ws.ConditionalFormatting.AddExpression($tableRange)
        $cfGreen.Formula = $greenFormula
        $cfGreen.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfGreen.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGreen

        $redFormula = "=INDIRECT(""" + $colLetter + """ & ROW())=""NIE"""
        $cfRed = $ws.ConditionalFormatting.AddExpression($tableRange)
        $cfRed.Formula = $redFormula
        $cfRed.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfRed.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightSalmon

        $yellowFormula = "=INDIRECT(""" + $colLetter + """ & ROW())=""ZMIANA"""
        $cfYellow = $ws.ConditionalFormatting.AddExpression($tableRange)
        $cfYellow.Formula = $yellowFormula
        $cfYellow.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfYellow.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::Yellow

        # Zablokuj edycję wszystkich komórek w tabeli...
        $ws.Cells[$tableRange].Style.Locked = $true
        # ... poza komórkami w kolumnach "Potwierdzam" i "Zmiana"
        $ws.Cells[$rangePotwierdzam].Style.Locked = $false
        $ws.Cells[$rangeZmiana].Style.Locked = $false

        # Zabezpiecz arkusz z hasłem "12345", zezwól na autofilter oraz sortowanie
        $ws.Protection.SetPassword("Boras987#^")
        $ws.Protection.AllowAutoFilter = $true
        $ws.Protection.AllowSort = $true
    }
    Close-ExcelPackage $pkg

    ## EKSPORT DO OSOBNEGO PLIKU DLA KAŻDEGO CSV

    # Używamy nazwy arkusza jako nazwy pliku indywidualnego
    $individualFile = Join-Path (Split-Path $ExcelFile -Parent) ("$sheetName.xlsx")
    if (Test-Path $individualFile) { Remove-Item $individualFile -Force }
    $csvData | Export-Excel -Path $individualFile -WorksheetName $sheetName -AutoSize -FreezeTopRow -TableStyle Medium2 -StartRow 1

    $pkgInd = Open-ExcelPackage -Path $individualFile
    $wsInd = $pkgInd.Workbook.Worksheets[$sheetName]
    if ($wsInd -ne $null) {
        # Walidacja dla "Potwierdzam"
        $dvInd = $wsInd.DataValidations.AddListValidation($rangePotwierdzam)
        $dvInd.Formula.Values.Clear()
        $dvInd.Formula.Values.Add("TAK")
        $dvInd.Formula.Values.Add("NIE")
        $dvInd.Formula.Values.Add("ZMIANA")
        $dvInd.ShowErrorMessage = $true
        $dvInd.ErrorTitle = "Nieprawidłowa wartość"
        $dvInd.Error = "Wybierz wartość TAK, NIE lub ZMIANA"
        $dvInd.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::Stop

        # Walidacja dla "Zmiana"
        $dvIndZmiana = $wsInd.DataValidations.AddListValidation($rangeZmiana)
        $dvIndZmiana.Formula.Values.Clear()
        $dvIndZmiana.Formula.Values.Add("R")
        $dvIndZmiana.Formula.Values.Add("RW")
        $dvIndZmiana.Formula.Values.Add("RWM")
        $dvIndZmiana.ShowErrorMessage = $true
        $dvIndZmiana.ErrorTitle = "Nieprawidłowa wartość"
        $dvIndZmiana.Error = "Wybierz wartość R, RW lub RWM"
        $dvIndZmiana.ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::Stop

        # Formatowanie warunkowe – te same formuły co wcześniej
        $cfIndGreen = $wsInd.ConditionalFormatting.AddExpression($tableRange)
        $cfIndGreen.Formula = $greenFormula
        $cfIndGreen.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfIndGreen.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGreen

        $cfIndRed = $wsInd.ConditionalFormatting.AddExpression($tableRange)
        $cfIndRed.Formula = $redFormula
        $cfIndRed.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfIndRed.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightSalmon

        $cfIndYellow = $wsInd.ConditionalFormatting.AddExpression($tableRange)
        $cfIndYellow.Formula = $yellowFormula
        $cfIndYellow.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cfIndYellow.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::Yellow

        # Zablokuj edycję wszystkich komórek w arkuszu...
        $wsInd.Cells[$tableRange].Style.Locked = $true
        # ... poza komórkami w kolumnach "Potwierdzam" i "Zmiana"
        $wsInd.Cells[$rangePotwierdzam].Style.Locked = $false
        $wsInd.Cells[$rangeZmiana].Style.Locked = $false

        # Zabezpiecz arkusz z hasłem "12345", zezwól na autofilter oraz sortowanie
        $wsInd.Protection.SetPassword("17499^")
        $wsInd.Protection.AllowAutoFilter = $true
        $wsInd.Protection.AllowSort = $true
    }
    Close-ExcelPackage $pkgInd
}

Write-Host "Eksport zakończony! Zbiorczy plik: $ExcelFile" -ForegroundColor Green
