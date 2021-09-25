Connect-PnPOnline -Url "<<SPO site URL>>" -UseWebLogin
$CSVPath = "C:\Temp\ListData_Libexport.csv" # Output CSV
$Lib = "Archive Repository" # library name
$LibItems = Get-PnPListItem -List $Lib -PageSize 5000
$LibDataCollection = @()
$Counter = 0

$LibItems | ForEach-Object {
        $ListItem  = Get-PnPProperty -ClientObject $_ -Property FieldValuesAsText
        $ListRow = New-Object PSObject 
        $Counter++
        Get-PnPField -List $Lib | ForEach-Object {
            $ListRow | Add-Member -MemberType NoteProperty -name $_.InternalName -Value $ListItem[$_.InternalName]
            }
        Write-Progress -PercentComplete ($Counter/$($LibItems.Count)*100) -Activity "Exporting List Items..." -Status "Exporting Item $Counter of $($LibItems.Count)"
        $LibDataCollection += $ListRow
}

$LibDataCollection | Export-CSV $CSVPath -NoTypeInformation