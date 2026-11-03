$ErrorActionPreference = 'Stop'
$repoPath = $PSScriptRoot
$buildPath = Join-Path $repoPath 'build'
$sourcePath = Join-Path $buildPath 'ConvertCsvToTableZh-build.xlsm'
$outputPath = Join-Path $repoPath 'addin\ConvertCsvToTableZh.xlam'
$requestPath = Join-Path $buildPath 'finalize.request'
$statusPath = Join-Path $buildPath 'build-status.json'
$csvExcel = $null
$csvBook = $null
try {
    Copy-Item -LiteralPath (Join-Path $repoPath 'src\ConvertCsvToTable.xlsm') -Destination $sourcePath
    $csvExcel = New-Object -ComObject Excel.Application
    $csvExcel.AutomationSecurity = 3
    $csvExcel.EnableEvents = $false
    $csvExcel.DisplayAlerts = $false
    $csvBook = $csvExcel.Workbooks.Open($sourcePath)
    $csvExcel.Visible = $true
    @{ state='editing'; workbook=$sourcePath; hwnd=$csvExcel.Hwnd } | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding utf8
    Write-Output 'BUILD_READY: workbook is open for localized VBA source edits.'
    $buildDeadline = (Get-Date).AddMinutes(35)
    while (-not (Test-Path -LiteralPath $requestPath)) {
        if ((Get-Date) -gt $buildDeadline) { throw 'Build edit timeout.' }
        Start-Sleep -Milliseconds 500
    }
    $csvBook.Worksheets.Item('CsvOpener_Parameters').Range('B1').Value2 = $true
    $csvBook.Title = 'CSV 转换（中文版）'
    $csvBook.Subject = '中文界面；UTF-8 / GBK CSV 自动转换'
    $csvBook.Author = 'furyutei; Chinese localization'
    $csvBook.Save()
    $csvBook.IsAddin = $true
    $csvBook.SaveAs($outputPath, 55)
    $csvBook.Close($false)
    $csvBook = $null
    @{ state='built'; output=$outputPath; sha256=(Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash } | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding utf8
    Write-Output ('BUILT: ' + $outputPath)
} catch {
    @{ state='failed'; error=$_.Exception.Message } | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding utf8
    throw
} finally {
    if ($null -ne $csvBook) { $csvBook.Close($false) }
    if ($null -ne $csvExcel) {
        $csvExcel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($csvExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
