
#Datei und Pfad angeben
$Path = "D:\Uebungs1.xlsx"
$Sheet = "Sheet1"

#öffnen der Excel
$objExcel = New-Object -ComObject Excel.Application
$WorkBook = $objExcel.Workbooks.Open($Path)
$WorkSheet = $WorkBook.sheets.item($Sheet)

#Erstellen des Datensatzobjekts und DatenListe
$DatensatzListe = New-Object -TypeName "System.Collections.ArrayList"
$Datensatz = [PSCustomObject]@{}
$headerRow = 1   #Zeile mit Attributnamen
$row = 2 #erste Datenzeile
$column = 1 #Startspalte
$rowrange = 6 #Letzte Zeile
$columnrange = 4 #Letzte Spalte

#befüllen der Datensatzliste
while($row -lt $rowrange)
{   
    $Datensatz = [PSCustomObject]@{}   
    while($column -lt $columnrange)
    {
        $header = $WorkSheet.Cells.Item($headerRow,$column).Text
        $Value = $WorkSheet.Cells.Item($row,$column).Text
        $Datensatz | Add-Member -MemberType NoteProperty -Name "$header" -Value "$value"
        $column++
    }
    $DatensatzListe.Add($Datensatz)
    Remove-Variable -Name Datensatz
    $row++
    $column = 1
}

Stop-Process -Name EXCEL
#END
