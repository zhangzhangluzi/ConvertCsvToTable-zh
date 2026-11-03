$ErrorActionPreference = 'Stop'
$addonPath = Join-Path $PSScriptRoot 'addin\ConvertCsvToTableZh.xlam'
$cases = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'build\fixtures\expected.json') -Raw | ConvertFrom-Json
$csvExcel = $null
$addonBook = $null
$results = [Collections.Generic.List[object]]::new()
try {
    $csvExcel = New-Object -ComObject Excel.Application
    $csvExcel.Visible = $false
    $csvExcel.DisplayAlerts = $false
    $addonBook = $csvExcel.Workbooks.Open($addonPath)
    $label = $csvExcel.Run("'ConvertCsvToTableZh.xlam'!ZhText", '4E2D 6587')
    if ($label -ne '中文') { throw 'Chinese labels are not working.' }
    foreach ($case in $cases) {
        $before = (Get-FileHash -LiteralPath $case.path).Hash
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $book = $csvExcel.Workbooks.Open([string]$case.path)
        try {
            $converted = $null
            foreach ($sheet in $book.Worksheets) {
                if ($sheet.Name.StartsWith('#')) { $converted = $sheet; break }
            }
            if ($null -eq $converted) { throw ('Automatic conversion did not create a table: ' + $case.name) }
            if ($converted.ListObjects.Count -ne 1) { throw ('Expected one formatted table: ' + $case.name) }
            $data = $converted.ListObjects.Item(1).Range.Value2
            if ($data.GetLength(0) -ne 4 -or $data.GetLength(1) -ne 5) { throw ('Unexpected dimensions: ' + $case.name) }
            for ($row=1; $row -le 4; $row++) {
                for ($col=1; $col -le 5; $col++) {
                    $expected = [string]$case.rows[$row-1][$col-1]
                    $actual = $data[$row,$col]
                    if ($col -eq 5 -and $row -gt 1) {
                        if ([Math]::Abs([double]$actual - [double]$expected) -gt 0.00000001) { throw "Amount mismatch in $($case.name) at $row,$col" }
                    } elseif ([string]$actual -cne $expected) {
                        throw "Text mismatch in $($case.name) at $row,$col : [$actual] vs [$expected]"
                    }
                }
            }
        } finally {
            $csvExcel.DisplayAlerts = $false
            $book.Close($false)
        }
        $after = (Get-FileHash -LiteralPath $case.path).Hash
        if ($before -ne $after) { throw 'Original CSV bytes were modified.' }
        $timer.Stop()
        $results.Add(@{name=$case.name;passed=$true;seconds=[Math]::Round($timer.Elapsed.TotalSeconds,2);originalUnchanged=$true})
        Write-Output ('PASS: ' + $case.name)
    }
    @{passed=$true;excelVersion=$csvExcel.Version;excelBuild=$csvExcel.Build;chineseLabels=$label;cases=$results} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'build\verification.json') -Encoding utf8
} finally {
    if ($null -ne $addonBook) { $addonBook.Close($false) }
    if ($null -ne $csvExcel) {
        $csvExcel.DisplayAlerts = $false
        $csvExcel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($csvExcel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
