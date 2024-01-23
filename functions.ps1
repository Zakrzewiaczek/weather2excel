function map($value, $fromLow, $fromHigh, $toLow, $toHigh) {
    return ($value - $fromLow) * ($toHigh - $toLow) / ($fromHigh - $fromLow) + $toLow
}

function Check-Time {
    $hour = (Get-Date).Hour
    while($hour -ne 5 -or $hour -ne 9 -or $hour -ne 13 -or $hour -ne 19 -or $hour -ne 21) {
        Start-Sleep -Second 10
    }
}

function make_worksheet {
    param (
        [string]$path,
        [string]$worksheet_name
    )

$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

$worksheet.Name = $worksheet_name

$range = $worksheet.Range("A1:A6")
$range.Merge()

#Wyśrodkowanie tekstu
$range = $worksheet.Range("A1:R100")
$range.HorizontalAlignment = -4108
$range.VerticalAlignment = -4108

#Szerokości rzędów i kolumn
$worksheet.Rows.Item(1).RowHeight = 27
$worksheet.Columns.Item("B:R").ColumnWidth = 16

$worksheet.Columns.Item("I:K").ColumnWidth = 22
$worksheet.Columns.Item("E").ColumnWidth = 22

$range = $worksheet.Range("2:6")
$range.RowHeight = 30

#Tekst
$worksheet.Cells.Item(1, 1).Orientation = 90
$worksheet.Cells.Item(1, 1).Value2 = Get-Date -Format "dd MMMM yyyy  r."

$worksheet.Cells.Item(1, 2).Value2 = "Time"
$worksheet.Cells.Item(1, 3).Value2 = "Weather"
$worksheet.Cells.Item(1, 4).Value2 = "Temperature"
$worksheet.Cells.Item(1, 5).Value2 = "Feels like "
$worksheet.Cells.Item(1, 6).Value2 = "Humidity"
$worksheet.Cells.Item(1, 7).Value2 = "Pressure"
$worksheet.Cells.Item(1, 8).Value2 = "Wind speed"
$worksheet.Cells.Item(1, 9).Value2 = "Wind direction"
$worksheet.Cells.Item(1, 10).Value2 = "Clouds type"
$worksheet.Cells.Item(1, 11).Value2 = "Cloud cover"
$worksheet.Cells.Item(1, 12).Value2 = "UV Level"
$worksheet.Cells.Item(1, 13).Value2 = "Gases: CO2"
$worksheet.Cells.Item(1, 14).Value2 = "Gases: NO2"
$worksheet.Cells.Item(1, 15).Value2 = "Gases: O3"
$worksheet.Cells.Item(1, 16).Value2 = "Gases: SO2"
$worksheet.Cells.Item(1, 17).Value2 = "Gases: PM2.5"
$worksheet.Cells.Item(1, 18).Value2 = "Gases: PM10"

$border = $worksheet.Range("A1:R6").Borders
$border.LineStyle = 1

#Wpisywanie wartości do tabeli
$worksheet.Cells.Item(2, 2).Value2 = "5:00"
$worksheet.Cells.Item(3, 2).Value2 = "9:00"
$worksheet.Cells.Item(4, 2).Value2 = "13:00"
$worksheet.Cells.Item(5, 2).Value2 = "17:00"
$worksheet.Cells.Item(6, 2).Value2 = "21:00"

$workbook.SaveAs($Path)

$workbook.Save()
$workbook.Close()
$excel.Quit()

Remove-Variable excel
}

function open_and_write {
    param (
        [string]$path,
        [int]$row,
        [string]$worksheet_name,
        [string]$APIKey,
        [string]$City
    )

$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($path)
$worksheet = $workbook.Worksheets.Item($worksheet_name)

$URL = "https://api.weatherapi.com/v1/current.json?key=$APIKey&q=$City&aqi=yes"

Write-Host "Downloading data"
$WeatherData = Invoke-RestMethod -Uri $URL -Method Get

#wpisywanie danych z weatherapi.com

Write-Host "Entering data"

$web = "https:" + ($WeatherData.current.condition.icon)
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($web, "$home\icon.jpg")

$picture = $worksheet.Shapes.AddPicture(("$home\icon.jpg"), $false, $true, $range.Left + 164, $range.Top - 3 + ($row * 30), 30, 30)

$worksheet.Cells.Item($row + 1, 4).Value2 = $WeatherData.current.temp_c.ToString() + "°C"
$worksheet.Cells.Item($row + 1, 5).Value2 = $WeatherData.current.feelslike_c.ToString() + "°C"

#kolory temp.

if ($WeatherData.current.temp_c -lt 0) { $worksheet.Cells.Item($row + 1, 4).Interior.Color = 15917529 } #D9E1F2
if ($WeatherData.current.temp_c -gt 0) { $worksheet.Cells.Item($row + 1, 4).Interior.Color = 14277119 } #FFD9D9

if ($WeatherData.current.feelslike_c -lt 0) { $worksheet.Cells.Item($row + 1, 5).Interior.Color = 15917529 } #D9E1F2
if ($WeatherData.current.feelslike_c -gt 0) { $worksheet.Cells.Item($row + 1, 5).Interior.Color = 14277119 } #FFD9D9

############

$worksheet.Cells.Item($row + 1, 6).Value2 = $WeatherData.current.humidity.ToString() + "%"

#kolor hum.

$r = [int](map $WeatherData.current.humidity 0 100 255 65)

$worksheet.Cells.Item($row + 1, 6).Interior.Color = $r + ($r * 256) + 16711680

if ($WeatherData.current.humidity -gt 50) { 
   $worksheet.Cells.Item($row + 1, 6).font.ColorIndex = 2; 
   $worksheet.Cells.Item($row + 1, 6).font.bold = $True 
}

############

$worksheet.Cells.Item($row + 1, 7).Value2 = $WeatherData.current.pressure_mb.ToString() + " hPa"
$worksheet.Cells.Item($row + 1, 8).Value2 = $WeatherData.current.wind_kph.ToString() + " km/h"
$worksheet.Cells.Item($row + 1, 9).Value2 = $WeatherData.current.wind_degree.ToString() + "° (" + $WeatherData.current.wind_dir.ToString() + ")"
$worksheet.Cells.Item($row + 1, 10).Value2 = "Not supported (work in progress)"
$worksheet.Cells.Item($row + 1, 11).Value2 = $WeatherData.current.cloud.ToString() + "%"
$worksheet.Cells.Item($row + 1, 12).Value = $WeatherData.current.uv
$worksheet.Cells.Item($row + 1, 13).Value = $WeatherData.current.air_quality.co
$worksheet.Cells.Item($row + 1, 14).Value = $WeatherData.current.air_quality.no2
$worksheet.Cells.Item($row + 1, 15).Value = $WeatherData.current.air_quality.o3
$worksheet.Cells.Item($row + 1, 16).Value = $WeatherData.current.air_quality.so2
$worksheet.Cells.Item($row + 1, 17).Value = $WeatherData.current.air_quality.pm2_5
$worksheet.Cells.Item($row + 1, 18).Value = $WeatherData.current.air_quality.pm10

del "$home\icon.jpg"

$workbook.Save()
$workbook.Close()
$excel.Quit()

Remove-Variable excel
}

function new_page {
param (
        [string]$path,
        [string]$worksheet_name
    )

$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($path)
$worksheet = $workbook.Worksheets.Add()
$worksheet.Name = $worksheet_name

$workbook.Save()
$workbook.Close()
$excel.Quit()

Remove-Variable excel
}


function wait_until_hour {

param (
        [int]$wait_until_hour,
        [bool]$waiting_until
    )

if($waiting_until) {
    while ((Get-Date).Hour -ne $wait_until_hour) {
        Start-Sleep -Seconds 10
    }
} else {
    while ((Get-Date).Hour -eq $wait_until_hour) {
        Start-Sleep -Seconds 10
    }
}

}

function go_to_folder {

param (
        [string]$folder_name
    )

if (-not (Test-Path -Path $folder_name -PathType Container)) {
    New-Item -ItemType Directory -Name $folder_name | Out-Null
}

Set-Location -Path $folder_name
}

function enter_data {

param (
        [int]$row,
        [string]$worksheet_name
    )

if (-not (Test-Path (Join-Path (Get-Location) ((Get-Date).Day.ToString() + ".xlsx")))) {
    Write-Host "Creating file"
    make_worksheet -path (Join-Path (Get-Location) ((Get-Date).Day.ToString() + ".xlsx")) -worksheet_name $worksheet_name
}

open_and_write -path (Join-Path (Get-Location) ((Get-Date).Day.ToString() + ".xlsx")) -worksheet_name $worksheet_name -row $row -APIKey "7aeb5673292d463786164127241401" -City "Nasielsk"

}
