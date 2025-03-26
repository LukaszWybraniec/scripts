# Ustawienia logowania – używamy loginu "ADMIN" oraz pustego hasła
$operator = "ADMIN"
$password = ""

# Pobranie od użytkownika liczby dni (okres graniczny)
$daysInput = Read-Host "Podaj liczbę dni (np. 7) jako okres graniczny, po upływie którego użytkownik ma zostać usunięty"
try {
    $days = [int]$daysInput
} catch {
    Write-Error "Podana wartość nie jest poprawną liczbą całkowitą."
    exit
}

# Ustawienie ścieżki katalogu z logami i utworzenie go, jeśli nie istnieje
$logDir = "D:\Roger_del_logs"
if (-not (Test-Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Output "Katalog $logDir został utworzony."
    } catch {
        Write-Error "Nie udało się utworzyć katalogu $logDir: $_"
        exit
    }
}

# Ustawienie nazwy pliku logu z datą wykonania
$dateStr = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "$logDir\deleted_users_log_$dateStr.txt"

try {
    # Utworzenie obiektu COM PR Master API przy użyciu właściwego ProgID
    $prMaster = New-Object -ComObject "PRMaster.PRMasterAutomation"
    if (-not $prMaster) {
        Write-Error "Nie udało się utworzyć obiektu COM PR Master API."
        exit
    }

    # Próba logowania
    $loginResult = $prMaster.Login($operator, $password)
    if (-not $loginResult) {
        Write-Error "Logowanie nie powiodło się. Sprawdź dane logowania."
        exit
    }
    Write-Output "Logowanie powiodło się."

    # Pobranie listy użytkowników w formacie XML
    $usersXml = $prMaster.GetUsersXml()
    if ([string]::IsNullOrEmpty($usersXml)) {
        Write-Output "Brak danych użytkowników."
        exit
    } else {
        Write-Output "Pobrana lista użytkowników (XML):"
        Write-Output $usersXml

        try {
            [xml]$xmlDoc = $usersXml
        } catch {
            Write-Error "Błąd parsowania XML: $_"
            exit
        }
        
        # Przygotowanie słownika kodów kart dostępu z tabeli "Transponder Codes"
        $transponderCodesDict = @{}
        $transponderTable = $xmlDoc.tables.table | Where-Object { $_.name -eq "Transponder Codes" }
        if ($transponderTable) {
            foreach ($tRow in $transponderTable.rows.row) {
                $tUserGUID = [string]$tRow.f[1]
                $tCode = [string]$tRow.f[0]
                if ($tUserGUID -and $tUserGUID -ne "null") {
                    if (-not $transponderCodesDict.ContainsKey($tUserGUID)) {
                        $transponderCodesDict[$tUserGUID] = $tCode
                    }
                }
            }
        }
        
        # Znalezienie tabeli "Users"
        $usersTable = $xmlDoc.tables.table | Where-Object { $_.name -eq "Users" }
        if (-not $usersTable) {
            Write-Error "Nie znaleziono tabeli Users w XML."
            exit
        }
        
        # Obliczenie daty granicznej na podstawie podanej liczby dni
        $cutoffDate = (Get-Date).AddDays(-$days)
        Write-Output "Data graniczna (tydzień temu, lub inny okres): $cutoffDate"
        
        # Funkcja przygotowująca kod karty w obu formatach (szesnastkowym i dziesiętnym):
        function Get-CardCode {
            param (
                [string]$userGUID,
                [hashtable]$dict
            )
            $cardCodeHex = "N/A"
            $cardCodeDec = "N/A"
            if ($dict.ContainsKey($userGUID)) {
                $cardCodeHex = $dict[$userGUID]
                try {
                    $cardCodeDec = [Convert]::ToUInt64($cardCodeHex,16)
                } catch {
                    $cardCodeDec = "Błąd konwersji"
                }
            }
            return @{ Hex = $cardCodeHex; Dec = $cardCodeDec }
        }
        
        # Iteracja przez każdy wiersz (użytkownika)
        foreach ($row in $usersTable.rows.row) {
            # Odczyt wartości pól – indeksy bazują na kolejności zdefiniowanej w XML:
            $userGUID         = [string]$row.f[1]
            $firstName        = [string]$row.f[2]
            $lastName         = [string]$row.f[3]
            $evidence         = [string]$row.f[4]    # nr RCP (T&A ID)
            $groupID          = [string]$row.f[6]    # grupa pracownika
            $official         = [string]$row.f[7]
            $active           = [string]$row.f[8]    # pole Active – "True" lub "False"
            $custom1          = [string]$row.f[11]
            $custom2          = [string]$row.f[12]
            $custom3          = [string]$row.f[13]
            $custom4          = [string]$row.f[14]
            $activityUntilStr = [string]$row.f[17]
            $deletedStr       = [string]$row.f[19]
            
            # Pomijamy użytkowników, którzy już zostali usunięci
            if ($deletedStr -eq "True") {
                Write-Output "Pomijam użytkownika: $firstName $lastName (GUID: $userGUID) - już usunięty."
                continue
            }
            
            # Próba sparsowania daty ActivityUntil, jeśli jest ustawione
            if ($activityUntilStr -and $activityUntilStr -ne "null") {
                try {
                    $activityUntil = [datetime]$activityUntilStr
                    Write-Output "Użytkownik: $firstName $lastName, GUID: $userGUID, ActivityUntil: $activityUntil"
                } catch {
                    Write-Warning "Nie udało się przekonwertować daty ActivityUntil: $activityUntilStr dla użytkownika $firstName $lastName (GUID: $userGUID)"
                    $activityUntil = $null
                }
            } else {
                Write-Output "Użytkownik $firstName $lastName (GUID: $userGUID) nie posiada ustawionej daty ActivityUntil."
                $activityUntil = $null
            }
            
            # Logika decydująca o usunięciu:
            if ($activityUntil -and $activityUntil -lt $cutoffDate) {
                # Warunek 1: ActivityUntil ustawione i wcześniejsze niż data graniczna – usuń automatycznie.
                Write-Output "Usuwanie użytkownika: $firstName $lastName (GUID: $userGUID) – ActivityUntil: $activityUntil jest wcześniejsze niż data graniczna."
                $deleteResult = $prMaster.DeleteUser($userGUID)
                if ([string]::IsNullOrEmpty($deleteResult) -or $deleteResult -eq 0) {
                    Write-Output "Użytkownik usunięty pomyślnie."
                    $cardCode = Get-CardCode -userGUID $userGUID -dict $transponderCodesDict
                    $cardCodeHex = $cardCode.Hex
                    $cardCodeDec = $cardCode.Dec
                    # Zapis do logu
                    $deletionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $logMsg = @"
------------------------------------------------------------
Data usunięcia       : $deletionDate
Imię i nazwisko     : $firstName $lastName
Evidence (T&A ID)   : $evidence
Grupa               : $groupID
Komentarz:
   Custom1         : $custom1
   Custom2         : $custom2
   Custom3         : $custom3
   Custom4         : $custom4
Kod karty           : $cardCodeHex (hex) / $cardCodeDec (dec)
------------------------------------------------------------
"@
                    Add-Content -Path $logFile -Value $logMsg
                } else {
                    Write-Warning "Błąd przy usuwaniu użytkownika $userGUID. HRESULT: $deleteResult"
                }
            }
            elseif ((($activityUntil -eq $null) -or ($activityUntil -eq [datetime]"2099-12-31T23:59:59")) -and ($active -eq "False")) {
                # Warunek 2: brak ActivityUntil LUB ActivityUntil równe "12/31/2099 23:59:59" i konto nieaktywne – zapytaj operatora.
                Write-Output "Konto nie jest aktywne (Active: $active) lub ActivityUntil wskazuje wartość domyślną dla użytkownika: $firstName $lastName (GUID: $userGUID)."
                $userInput = Read-Host "Czy chcesz usunąć tego użytkownika? (T/N)"
                if ($userInput -eq "T" -or $userInput -eq "t") {
                    Write-Output "Usuwanie użytkownika: $firstName $lastName (GUID: $userGUID) – konto nieaktywne lub ActivityUntil równe 12/31/2099 23:59:59."
                    $deleteResult = $prMaster.DeleteUser($userGUID)
                    if ([string]::IsNullOrEmpty($deleteResult) -or $deleteResult -eq 0) {
                        Write-Output "Użytkownik usunięty pomyślnie."
                        $cardCode = Get-CardCode -userGUID $userGUID -dict $transponderCodesDict
                        $cardCodeHex = $cardCode.Hex
                        $cardCodeDec = $cardCode.Dec
                        $deletionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $logMsg = @"
------------------------------------------------------------
Data usunięcia       : $deletionDate
Imię i nazwisko     : $firstName $lastName
Evidence (T&A ID)   : $evidence
Grupa               : $groupID
Komentarz:
   Custom1         : $custom1
   Custom2         : $custom2
   Custom3         : $custom3
   Custom4         : $custom4
Kod karty           : $cardCodeHex (hex) / $cardCodeDec (dec)
------------------------------------------------------------
"@
                        Add-Content -Path $logFile -Value $logMsg
                    } else {
                        Write-Warning "Błąd przy usuwaniu użytkownika $userGUID. HRESULT: $deleteResult"
                    }
                }
                else {
                    Write-Output "Pominięto usunięcie użytkownika $firstName $lastName (GUID: $userGUID)."
                }
            }
            else {
                Write-Output "Użytkownik $firstName $lastName (GUID: $userGUID) nie spełnia kryterium usunięcia."
            }
        }
    }
} catch {
    Write-Error "Wystąpił błąd: $_"
}
