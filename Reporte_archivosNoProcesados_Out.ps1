Clear-Host
$vig = 15 ### Debe fijar la vigencia de los 
#$age = (Get-Date).AddMinutes(-2)
$path = '\\10.129.58.122\interfaces'
#$path = 'C:\xcom'
#$folder = "Fail"
$folder = "OUT"

#$filepath = "C:\Temp\"
$filepath = "F:\XCOM\Monitoreo\Reporte\"
$date1 = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$file = $filepath + "Reporte_Archivos_" + $date1 + ".html"
New-Item $filepath -type directory -force
$NomClient = "ATH"
$Global:FileCount = 0

#////////////////////////////////////



"<html>" | Out-File $file -append
"<head>" | Out-File $file -append
"<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" | Out-File $file -append
'<title>.: REPORTE ARCHIVOS FALLIDOS :. </title>' | Out-File $file -append
'<STYLE TYPE="text/css">' | Out-File $file -append
"<!--" | Out-File $file -append
"td {" | Out-File $file -append
"font-family: Tahoma;" | Out-File $file -append
"font-size: 11px;" | Out-File $file -append
"border-top: 1px solid #999999;" | Out-File $file -append
"border-right: 1px solid #999999;" | Out-File $file -append
"border-bottom: 1px solid #999999;" | Out-File $file -append
"border-left: 1px solid #999999;" | Out-File $file -append
"padding-top: 0px;" | Out-File $file -append
"padding-right: 0px;" | Out-File $file -append
"padding-bottom: 0px;" | Out-File $file -append
"padding-left: 0px;" | Out-File $file -append
"border: 1px solid black;" | Out-File $file -append
"}" | Out-File $file -append
"body {" | Out-File $file -append
"margin-left: 5px;" | Out-File $file -append
"margin-top: 5px;" | Out-File $file -append
"margin-right: 0px;" | Out-File $file -append
"margin-bottom: 10px;" | Out-File $file -append
"" | Out-File $file -append
"table {" | Out-File $file -append
"border: thin solid #000000;" | Out-File $file -append
"}" | Out-File $file -append
"-->" | Out-File $file -append
"</style>" | Out-File $file -append
"</head>" | Out-File $file -append
"<body>" | Out-File $file -append
"<table width='100%'>" | Out-File $file -append
"<tr bgcolor='#CCCCCC'>" | Out-File $file -append
"<td colspan='7' height='25' align='center'>" | Out-File $file -append
"<font face='tahoma' color='#000000' size='4'><strong> .: REPORTE PLATAFORMA $NomClient - $date1 :. </strong></font>" | Out-File $file -append
"</td>" | Out-File $file -append
"</tr>" | Out-File $file -append
"</table>" | Out-File $file -append
#///////////////////////

"<br><table width='100%'>" | Out-File $file -append
"<tr bgcolor='#CCCCCC'>" | Out-File $file -append
"<td colspan='7' height='25' align='center'>" | Out-File $file -append
"<font face='tahoma' color='#000000' size='4'><strong> .: REPORTE DE ARCHIVOS SUPERIORES 15 MINUTOS XCOM $NomClient :. </strong></font>" | Out-File $file -append
"<tr bgcolor=#CCCCCC>" | Out-File $file -append
"<td align='center'><strong> DIRECTORIO </strong></td>" | Out-File $file -append
"<td align='center'><strong> NOMBRE ARCHIVO </strong></td>" | Out-File $file -append
"<td align='center'><strong> FECHA CREACION </strong></td>" | Out-File $file -append
## "<td align='center'><strong> TIEMPO EN COLA </strong></td>" | Out-File $file -append
"</tr>" | Out-File $file -append
"</td>" | Out-File $file -append
"</tr>" | Out-File $file -append


#$pathFails = Get-ChildItem $path -recurse -Directory | Where-Object {($_.Name -match $folder) -and ($_.FullName -notlike '\\10.129.58.122\interfaces\0000\*')} | Select-Object FullName
#$pathFails = Get-ChildItem $path -recurse -Directory | Where-Object {$_.Name -match $folder} | Select-Object FullName
## $date = "{0:yyyy_MM_dd-HH_mm}" -f (get-date)
$pathFails = import-Csv -path 'F:\XCOM\Monitoreo\Reporte\Reporte_Rutas_Out.csv'

$date = Get-Date

foreach ($pathFail in $pathFails){
    $Files = $null
    Write-Host "Conectividad"
    $pf = $pathFail.FullName
    $pf
    $Files = Get-ChildItem -path $pf -Recurse | Where-Object {$_.LastWriteTime -lt (Get-Date).AddMinutes(-$vig)}
    $Files.count
    if ($Files.count -ge 1) {
        $Global:FileCount = $Global:FileCount + $Files.count
        foreach ($fileF in $Files){
            $name = $fileF.Name
            $Ctime = $fileF.CreationTime
            $pathFail1 = $pathFail.FullName
            "<td align='center'> $pathFail1 </td>" | Out-File $file -append
            "<td align='center'> $name </td>" | Out-File $file -append
            "<td align='center'> $Ctime </td></tr>" | Out-File $file -append
        }
    }
}

$Global:FileCount
$header = Get-Content $file -Raw
if ($Global:FileCount -ge 1) {
    #Send-MailMessage -SmtpServer correo.ath.net -Port 25 -To edwin.giraldo@ibm.com, jglugo@co.ibm.com, amorales@co.ibm.com -From MonXCOM@ath.com.co -Subject “[Urgente]-Reporte de archivos no procesados en ATH Fecha: $date” -BodyAsHtml -Body $header -Attachments $file	
    Write-Host "Correo enviado" -ForegroundColor Green
} else {
    Remove-Item -Path $file -Force
    Write-Host "Archivo Elminado" -ForegroundColor Red
}
