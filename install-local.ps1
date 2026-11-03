$ErrorActionPreference = 'Stop'
$sourcePath = Join-Path $PSScriptRoot 'addin\ConvertCsvToTableZh.xlam'
$installDirectory = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\Addins'
$installPath = Join-Path $installDirectory 'ConvertCsvToTableZh.xlam'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw '请先保存工作并关闭 Excel，再安装。' }
if (-not (Test-Path -LiteralPath $installDirectory)) { [void](New-Item -ItemType Directory -Path $installDirectory) }
if (Test-Path -LiteralPath $installPath) {
    $backupPath = Join-Path $PSScriptRoot ('build\ConvertCsvToTableZh-before-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlam')
    Copy-Item -LiteralPath $installPath -Destination $backupPath
}
Copy-Item -LiteralPath $sourcePath -Destination $installPath -Force
$csvExcel = $null
try {
    $csvExcel = New-Object -ComObject Excel.Application
    $csvExcel.Visible = $false
    $csvExcel.DisplayAlerts = $false
    $blank = $csvExcel.Workbooks.Add()
    $addon = $csvExcel.AddIns.Add($installPath, $false)
    $addon.Installed = $true
    if (-not $addon.Installed) { throw 'Excel 未启用加载项。' }
    $label = $csvExcel.Run("'ConvertCsvToTableZh.xlam'!ZhText", '4E2D 6587')
    if ($label -ne '中文') { throw '安装后的中文功能检查失败。' }
    $blank.Close($false)
    Write-Output ('INSTALLED: ' + $installPath)
} finally {
    if ($null -ne $csvExcel) {
        $csvExcel.DisplayAlerts = $false
        $csvExcel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($csvExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
@{installed=$true;path=$installPath;autoConversion=$true;sha256=(Get-FileHash -LiteralPath $installPath).Hash} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'build\installation.json') -Encoding utf8
