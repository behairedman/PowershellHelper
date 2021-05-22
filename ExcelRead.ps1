#Datei und Pfad angeben - static
$folder = "D:\"

$file = Get-ChildItem -Path "D:\" | Where-Object name -Like *xlsx
$Path = $file.fullname



foreach($p in $Path)
{
        #öffnen der Excel
        $objExcel = New-Object -ComObject Excel.Application
        $WorkBook = $objExcel.Workbooks.Open($p)
        $Sheet = $WorkBook.Worksheets | where index -EQ 1
        #$Sheet = $Sheet.name
        $WorkSheet = $WorkBook.sheets.item($Sheet)

        #Erstellen des Datensatzobjekts und DatenListe
        $DatensatzListe = New-Object -TypeName "System.Collections.ArrayList"
        $Datensatz = [PSCustomObject]@{}
        $headerRow = 1   #Zeile mit Attributnamen
        $row = 2 #erste Datenzeile
        $column = 1 #Startspalte
        $rowrange = 10 #Letzte Zeile + 1
        $columnrange = 4 #Letzte Spalte + 1

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

        Write-Output "$p wurde gschrieben"
        sleep -Seconds 10
        Stop-Process -Name EXCEL
        $CSVPath = "$($folder)\Output.csv"

        foreach($_ in $DatensatzListe)
        {
            Export-Csv -InputObject $_ -Path $CSVPath -Append -NoTypeInformation
        }
        
}

#END
