param (
    [string]$DrivePath = "W:\",
    [string]$OutputFolder = "C:\Uprawnienia"
)

# Ustawienie preferencji debugowania – aby zobaczyć komunikaty debug, ustaw na "Continue"
$DebugPreference = "Continue"

# Sprawdzenie, czy moduł Active Directory jest dostępny
if (-Not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    Write-Host "Moduł Active Directory nie jest dostępny. Sprawdzenie członków grup domenowych nie będzie działać!" -ForegroundColor Red
}

# Tworzenie folderu na pliki CSV (jeśli nie istnieje)
if (-Not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}

# Funkcja mapująca uprawnienia NTFS do uproszczonych kategorii
function Get-PermissionLevel {
    param ([string]$rights)
    
    # Usuń "Synchronize"
    $filteredRights = $rights -replace "Synchronize", "" -replace ",\s+", ","
    
    if ($filteredRights -match "FullControl") {
        return "FullControl"
    } elseif ($filteredRights -match "Modify") {
        return "Read/Write/Modify"
    } elseif ($filteredRights -match "Write") {
        return "Read/Write"
    } elseif ($filteredRights -match "Read") {
        return "Read"
    } else {
        return "Unknown"
    }
}

# Nowa funkcja pobierająca dane użytkownika – wyciąga DisplayName oraz atrybut cn.
function Get-UserInfo {
    param ([string]$SamAccountName)
    try {
        $user = Get-ADUser -Filter {SamAccountName -eq $SamAccountName} -Properties DisplayName, GivenName, cn -ErrorAction Stop
        Write-Debug "Pobrano dane użytkownika dla $SamAccountName. DisplayName: '$($user.DisplayName)', GivenName: '$($user.GivenName)', cn: '$($user.cn)'"
        $displayName = if ([string]::IsNullOrEmpty($user.DisplayName)) {
            if (-not [string]::IsNullOrEmpty($user.GivenName)) { $user.GivenName } else { $SamAccountName }
        } else {
            $user.DisplayName
        }
        return [PSCustomObject]@{ DisplayName = $displayName; CN = $user.cn }
    } catch {
        Write-Debug "Błąd przy pobieraniu użytkownika ${SamAccountName}: $_"
        return [PSCustomObject]@{ DisplayName = $SamAccountName; CN = $SamAccountName }
    }
}

# Funkcja pobierająca członków grupy – wynik budujemy jako ArrayList.
function Get-GroupMembers {
    param ([string]$GroupName)
    try {
        if ($GroupName -match "^(.*)\\(.*)$") {
            $GroupName = $matches[2]
        }
        $members = @(Get-ADGroupMember -Identity $GroupName -ErrorAction Stop)
        Write-Debug "Grupa '$GroupName' zwróciła $($members.Count) członków."
        $result = New-Object System.Collections.ArrayList
        foreach ($member in $members) {
            $userInfo = Get-UserInfo $member.SamAccountName
            $null = $result.Add($userInfo)
        }
        Write-Debug "Wynik Get-GroupMembers: $($result.Count) elementów."
        return $result
    } catch {
        Write-Debug "Błąd przy pobieraniu członków grupy '$GroupName': $_"
        return $null
    }
}

# Funkcja przetwarzająca pojedynczy katalog – dodaje rekordy do CSV od razu po znalezieniu członków.
function Process-Folder {
    param (
        [System.IO.DirectoryInfo]$folder
    )
    
    $FolderPath = $folder.FullName
    Write-Host "Przetwarzanie katalogu: $FolderPath" -ForegroundColor Cyan

    $acl = Get-Acl -Path $FolderPath
    $results = @()

    foreach ($ace in $acl.Access) {
        $identity = $ace.IdentityReference.Value
        $permissions = Get-PermissionLevel ($ace.FileSystemRights -join ", ")
        Write-Debug "Przetwarzam wpis ACL: $identity z uprawnieniami: $permissions"

        if ($identity -match "^KLIPPAN\\") {
            $groupMembers = Get-GroupMembers -GroupName $identity
            $groupMembersArray = @($groupMembers)
            Write-Debug "Dla grupy $identity pobrano $($groupMembersArray.Count) członków."
            if ($groupMembersArray.Count -gt 0) {
                foreach ($gm in $groupMembersArray) {
                    $results += [PSCustomObject]@{
                        "Grupa / Użytkownik" = $gm.DisplayName
                        "Imię i Nazwisko"    = $gm.DisplayName
                        "Nazwa"                 = $gm.CN
                        "Uprawnienia"        = $permissions
                    }
                }
            } else {
                Write-Debug "Grupa $identity została pominięta, ponieważ lista członków jest pusta."
            }
        }
        else {
            $userInfo = Get-UserInfo ($identity.Split("\")[-1])
            $results += [PSCustomObject]@{
                "Grupa / Użytkownik" = $identity
                "Imię i Nazwisko"    = $userInfo.DisplayName
                "Nazwa"                 = $userInfo.CN
                "Uprawnienia"        = $permissions
            }
        }
    }

    $csvFile = Join-Path -Path $OutputFolder -ChildPath ("Uprawnienia_" + $folder.Name + ".csv")
    $results | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Dane zapisane do: $csvFile" -ForegroundColor Green
}

# Pobranie listy katalogów (w tym ukrytych) z zadanego dysku
$folders = Get-ChildItem -Path $DrivePath -Directory -Force
if ($folders.Count -eq 0) {
    Write-Host "Brak katalogów na dysku $DrivePath!" -ForegroundColor Red
    exit
}

foreach ($folder in $folders) {
    if ($folder.Name -eq "Etikett Bartender") {
        Write-Host "Pomijam katalog: $($folder.FullName)" -ForegroundColor Yellow
        continue
    }
    if (($folder.Name -eq "Monitor_pliki") -or ($folder.Name -eq "Products")) {
        $subFolders = Get-ChildItem -Path $folder.FullName -Directory -Force
        if ($subFolders.Count -eq 0) {
            Write-Host "Katalog $($folder.FullName) nie zawiera podkatalogów." -ForegroundColor Yellow
        } else {
            foreach ($subFolder in $subFolders) {
                Process-Folder -folder $subFolder
            }
        }
    } else {
        Process-Folder -folder $folder
    }
}

Write-Host "Eksport zakończony! Pliki CSV znajdują się w: $OutputFolder" -ForegroundColor Green
