Clear-Host
$path = '\\10.129.58.122\interfaces'
#$path = 'C:\xcom'
#$folder = "Fail"
$folder = "FAIL"

#$filepath = "C:\Temp\"
$filepath = "F:\XCOM\Monitoreo\Reporte\"
$date1 = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$file = $filepath + "Reporte_Rutas.csv"
New-Item $filepath -type directory -force
$NomClient = "ATH"
$Global:FileCount = 0

#////////////////////////////////////

#$pathFails = Get-ChildItem $path -recurse -exclude '*.zip' -Directory -ErrorAction Continue | ?{ $_.PSIsContainer } | Where-Object {($_.Name -match $folder) -and ($_.FullName -notlike '\\10.129.58.122\interfaces\0000\*')} | Select-Object FullName | Export-Csv -Path $file
Get-ChildItem $path -recurse -exclude '*.zip' -Directory -ErrorAction Continue | ?{ $_.PSIsContainer } | Where-Object {($_.Name -match $folder) -and ($_.FullName -notlike '\\10.129.58.122\interfaces\0000\*')} | Select-Object FullName | Export-Csv -Path $file