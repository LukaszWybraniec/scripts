

function GetData
{

    param ($data)
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = "\\192.168.0.20\common\Karty_obiegowe"
    Filter = 'Documents (*.docx)|*.docx'
    }
    $null = $FileBrowser.ShowDialog()


    #write-host $FileBrowser.FileName
    

    $word = New-Object -ComObject Word.application

    $document = $word.Documents.Open($FileBrowser.FileName)
    $data.imieNazwisko = $document.Tables[1].Cell(1,2).range.text
    $data.imie = $data.imieNazwisko.split(" ")[0]
    $data.nazwisko = $data.imieNazwisko.split(" ")[1]
    $data.nrPracownika = $document.Tables[1].Cell(1,4).range.text
    $data.stanowisko = $document.Tables[1].Cell(2,2).range.text
    $data.firma = $document.Tables[1].Cell(7,2).range.text
    $data.file = $FileBrowser.FileName;


    #$data.imieNazwisko = $data.imieNazwisko.Replace(" ","")
    $data.imie = $data.imie  -replace "\W"
    $data.nazwisko = $data.nazwisko -replace "\W"
    $data.nrPracownika = $data.nrPracownika -replace "\W"
    $data.stanowisko = $data.stanowisko -replace "\W"

    if ($data.firma -match "andrenplast"){$data.nrPracownika = "4"+$data.nrPracownika}

    


    $document.close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    Remove-Variable word
    ##pause

    return $data


}

function addUser
{


    param ($data,$id,$rogerData)

    #[int]$idDec = Read-Host -Prompt "Wczytaj kartę"
    

    #$idHex = "{0:X}" -f $idDec

    $imieNorm = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding(1251).GetBytes($data.imie))
    $nazwiskoNorm = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding(1251).GetBytes($data.nazwisko))

    $guid = "-1"

    $xml = "

<tables>
    <table name='Users'>
        <fields>
            <field name='FirstName' type='String' />
            <field name='LastName' type='String' />
            <field name='GroupID' type='Integer' />
            <field name='Evidence' type='Integer' />
        </fields>
        <rows>
            <row>
                <f>"+$imieNorm+"</f>
                <f>"+$nazwiskoNorm+"</f>
                <f>"+$rogerData.grupaRogerID+"</f>
                <f>"+$data.nrPracownika+"</f>
            </row>
        </rows>
        </table>
            <table name='Transponder Codes'>
                <fields>
                    <field name='Code' type='String' />
                </fields>
            <rows>
                <row>
                    <f>"+$id.hex+"</f>
                </row>
            </rows>
    </table>
</tables>
"

    #write-host $xml

    $roger = New-Object -ComObject PRMaster.PRMasterAutomation
    $roger.Login("ADMIN","")

    [bool]$alreadyExistsNrPracownika = $false
    [xml]$usersXml = $roger.GetUsersXml()

    $usersArry = ($usersXml.tables.table | Where-Object name -eq 'Users').rows.row

    foreach ($user in $usersArry){if ($user.f[4] -eq $data.nrPracownika){$alreadyExistsNrPracownika = $true}}

    if ($alreadyExistsNrPracownika){
       Write-Host "PODANY NR PRACOWNIKA JUŻ ISTNIEJE. UŻYTKOWNIK NIE ZOSTAŁ DODANY." -ForegroundColor red
        pause

    }

    [bool]$alreadyExistsTransponderCodes = $false


    $TransponderCodesArry = ($usersXml.tables.table | Where-Object name -eq 'Transponder Codes').rows.row

    foreach ($TransponderCodes in $TransponderCodesArry){if ($TransponderCodes.f[0] -match $id.hex){$alreadyExistsTransponderCodes = $true}}

    if ($alreadyExistsTransponderCodes){
        Write-Host "PODANA KARTA DOSTĘPU JUŻ ISTNIEJE. UŻYTKOWNIK NIE ZOSTAŁ DODANY." -ForegroundColor red
        pause

    }

    if (!($alreadyExistsNrPracownika) -and !($alreadyExistsTransponderCodes)){
        $guid = $roger.AddNewUser(48,$xml,1)
        $roger.SendUserConfigToSystem($guid,0)
        #$rogerData.guid = $guid ##############DO POPRAWY##############
    }


    
    $roger.Logout()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($roger)
    Remove-Variable roger
    return $rogerData
    
    #pause
    
}

function AddIdToDocx{
    param ($data,$id)
    

    $word = New-Object -ComObject Word.application
    [string]$t = $id.dec
    $document = $word.Documents.Open($data.file)
    $document.Tables[2].Cell(6,2).range.text = "$t"

    ###pause
    $document.close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    Remove-Variable word
    

}

function ReadID{
    param ($id)
    

    #write-host "************"$id

    #[int]$idDec = Read-Host -Prompt "Wczytaj kartę"
    #$idHex = "{0:X}" -f $idDec
    $idDec = "-1"
    Write-Host "CZYTNIK MUSI BYĆ W TRYBIE 13 no. in D" -ForegroundColor red
    [int64]$idDec = Read-Host -Prompt "Wczytaj kartę"
    $idHex = "{0:X}" -f $idDec

    $id.dec = $idDec
    $id.hex = $idHex
    ###pause

    $bin = [convert]::ToString($id.dec,2)

    $binNew = $bin.Substring($bin.Length -24, 24)
    $id.decNew = [convert]::ToInt32($binNew,2)
    write-host $id.decNew
    #pause
    
    return $id
}

function AddPerf{

    #konieczny moduł MySQl

    param ($data, $cred, $id)
    import-module -name MySQL

    $numer = $data.nrPracownika
    $imie_nazwisko = $data.imie+" "+$data.nazwisko
    $numer_karty = $id.decNew

    Write-Host $numer
    Write-Host $imie_nazwisko
    Write-Host $id.decNew
    ##pause

    $Connection = Connect-MySqlServer -Credential $cred -ComputerName "192.168.0.21" -Database "raporty"

    Invoke-MySqlQuery  -Query "INSERT INTO pracownicy (numer,numer_karty,imie_nazwisko) VALUES('$numer','$numer_karty','$imie_nazwisko')"

    Disconnect-MySqlServer -Connection $Connection
    ##pause



}

function Show-Data{
    param ($data, $id, $rogerData)

    #Clear-Host
    Write-Host "================================================================================================"
    Write-Host "Imie: "$data.imie
    Write-Host "Nazwisko: "$data.nazwisko
    Write-Host "Numer pracownika: "$data.nrPracownika
    Write-Host "Stanowisko: "$data.stanowisko
    Write-Host "Firma: "$data.firma
    Write-Host "Grupa dostępu: "$rogerData.grupaRogerTxt
    Write-Host "ID grupy dostępu: "$rogerData.grupaRogerID
    Write-Host "Plik obiegówki: "$data.file
    Write-Host "Nr karty DEC: "$id.dec
    Write-Host "Nr karty HEX: "$id.hex
    Write-Host "Nr karty 24bit: "$id.decNew
    Write-Host "GUID z Rogera: "$rogerData.guid
    Write-Host ""
    Write-Host ""

}


function Show-Menu {
    #param ($data)
    

    
    Write-Host "1: Wczytaj kartę obiegową."
    Write-Host "2: Wczytaj kartę dostępu."
    Write-Host "3: Wybierz grupę dostępu do obszarów fabryki."
    Write-Host "4: Dodaj użytkownika do Rogera."
    Write-Host "5: Dodaj użytkownika do raportów wydajności"
    Write-Host "9: Wykonaj kroki 1, 2, 3 oraz 4."
    Write-Host "Q: Wciśnij Q aby wyjść."
}

function Grupa-Roger{
    param ($data,$rogerData,$id)

    Show-Data -data $data -id $id -rogerData $rogerData

    
    $rogerData.grupaRogerID = "250"
    $rogerData.grupaRogerTxt = "Grupa bez Dostępu"
    $roger = New-Object -ComObject PRMaster.PRMasterAutomation
    $roger.Login("ADMIN","")
    $groupName = "ERROR"
    #$grp = ""

    [xml]$groupsXml = $roger.GetGroupsXml()
    $groupsArray = ($groupsXml.tables.table | Where-Object name -eq 'Groups').rows.row

    $roger.Logout()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($roger)
    Remove-Variable roger
    
    Write-Host "Wybierz grupę dostępu:"
    foreach ($group in $groupsArray){Write-Host $group.f[0]$group.f[1]}
    $selection = Read-Host "Please make a selection"

    foreach ($group in $groupsArray){
        if ($group.f[0] -eq $selection){
            $groupName = $group.f[1]
        
        }
    }

    $rogerData.grupaRogerID = $selection
    $rogerData.grupaRogerTxt = $groupName
    
    Write-Host $rogerData.grupaRogerID
    Write-Host $rogerData.grupaRogerTxt
    ##pause
    return $rogerData

}


$data = "" | Select-Object -Property imieNazwisko,imie,nazwisko,nrPracownika,stanowisko,firma,file
$id = "" | Select-Object -Property dec,hex,decNew
$rogerData = "" | Select-Object -Property grupaRogerTxt,grupaRogerID,guid
Clear-Host
do
 {
    
    
    Show-Data -data $data -id $id -rogerData $rogerData
    
    Show-Menu -data $data
    $selection = Read-Host "Please make a selection"
    switch ($selection)
    {
    '1' {
    $data = "" | Select-Object -Property imieNazwisko,imie,nazwisko,nrPracownika,stanowisko,firma,file
    $id = "" | Select-Object -Property dec,hex,decNew
    $rogerData = "" | Select-Object -Property grupaRogerTxt,grupaRogerID,guid
    $data = GetData -data $data
    
    }
    '2' {
    $id = "" | Select-Object -Property dec,hex,decNew
    $id = ReadID -id $id

    }
    '3' {
    $rogerData = "" | Select-Object -Property grupaRogerTxt,grupaRogerID,guid
    $rogerData = Grupa-Roger -roger $rogerData -data $data -id $id

    }
    '4' {
    $rogerData = AddUser -data $data -id $id -rogerData $rogerData
    AddIdToDocx -data $data -id $id
    }
    '5' {
    if (!$cred) {$cred = Get-Credential -UserName root -Message "Enter password"}
    AddPerf -data $data -cred $cred -id $id
    }

    '9' {
    #1
    $data = "" | Select-Object -Property imieNazwisko,imie,nazwisko,nrPracownika,stanowisko,firma,file
    $id = "" | Select-Object -Property dec,hex,decNew
    $rogerData = "" | Select-Object -Property grupaRogerTxt,grupaRogerID,guid
    $data = GetData -data $data
    #2
    $id = "" | Select-Object -Property dec,hex,decNew
    $id = ReadID -id $id
    #3
    $rogerData = "" | Select-Object -Property grupaRogerTxt,grupaRogerID,guid
    $rogerData = Grupa-Roger -roger $rogerData -data $data -id $id
    #4
    $rogerData = AddUser -data $data -id $id -rogerData $rogerData
    AddIdToDocx -data $data -id $id

    }
    }
    ###pause
 }
 until ($selection -eq 'q')
Clear-Variable data -ErrorAction SilentlyContinue
#$cred.Password.Clear()
Clear-Variable cred -ErrorAction SilentlyContinue
Remove-Variable cred -ErrorAction SilentlyContinue

