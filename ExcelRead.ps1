
#D:\Uebungs.xlxs

#$objExcel.WorkBooks | Select-Object -Property name, path, WriteReservedBy
#$objExcel.WorkBooks | Get-Member
#$WorkBook | Get-Member -Name *sheet*





<#
$MyArray = @()

$Datensatz = [PSCustomObject]@{
    ID     = '0'
    Note = '0'
    Assignee    = '0'
}

$ObjektListe = New-Object -TypeName "System.Collections.ArrayList"
#>

#öffnen der Excel
$objExcel = New-Object -ComObject Excel.Application
$WorkBook = $objExcel.Workbooks.Open("D:\Uebungs1.xlsx")
$WorkSheet = $WorkBook.sheets.item("Sheet1")
#Erstellen des Datensatzes
$DatensatzListe = New-Object -TypeName "System.Collections.ArrayList"
$Datensatz = [PSCustomObject]@{}
$header = 1   
$row = 2
$column = 1
$rowrange = 6
$columnrange = 4

while($row -lt $rowrange)
{   
    $Datensatz = [PSCustomObject]@{}   
    while($column -lt $columnrange)
    {
        #$WorkSheet.Cells.Item($row,$column).Text
        #Write-Output "reihe $row spalte $column"
    
        $header = $WorkSheet.Cells.Item($header,$column).Text
        $Value = $WorkSheet.Cells.Item($row,$column).Text
        #$ObjektListe.Add($cell)
        $Datensatz | Add-Member -MemberType NoteProperty -Name "$header" -Value "$value"

        $column++
    }

    $DatensatzListe.Add($Datensatz)
    Remove-Variable -Name Datensatz
    $row++
    $column = 1
}










#erstellen eine Datensatzliste
$row = 2
$column = 1
$ObjektListe = New-Object -TypeName "System.Collections.ArrayList"

while($row -lt $rowrange)
{

    while($column -lt $columnrange)
    {
        $cell = $WorkSheet.Cells.Item($row,$column).Text
        
        $column++
    }
    $column = 1
    $row++
}