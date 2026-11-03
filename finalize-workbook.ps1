$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-com.ps1')
$csvExcel = [CsvExcelRunningObject]::Get()
$csvBook = $csvExcel.Workbooks.Item('ConvertCsvToTableZh-build.xlsm')
$outputPath = Join-Path $PSScriptRoot 'addin\ConvertCsvToTableZh.xlam'
$csvExcel.EnableEvents = $false
$csvExcel.DisplayAlerts = $false
$csvBook.Worksheets.Item('CsvOpener_Parameters').Range('B1').Value2 = $true
$csvBook.Title = 'CSV 转换（中文版）'
$csvBook.Subject = '中文界面；UTF-8 / GBK CSV 自动转换'
$csvBook.Author = 'furyutei; Chinese localization'
$csvBook.Save()
$csvBook.IsAddin = $true
$csvBook.SaveAs($outputPath, 55)
$csvBook.Close($false)
$csvExcel.Quit()
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($csvExcel)
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
@{state='built';output=$outputPath;sha256=(Get-FileHash -LiteralPath $outputPath).Hash} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'build\build-status.json') -Encoding utf8
Write-Output ('BUILT: ' + $outputPath)
