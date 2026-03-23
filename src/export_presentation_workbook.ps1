$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$tablesDir = Join-Path $root "exports\tables"
$outputPath = Join-Path $tablesDir "presentation_tables.xlsx"

$sheets = @(
    @{ Name = "Recommended Mix"; Csv = "recommended_mix.csv" },
    @{ Name = "Cluster Summary"; Csv = "cluster_summary.csv" },
    @{ Name = "Representative Bottles"; Csv = "cluster_representatives.csv" }
)

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt $sheets.Count) {
        $null = $workbook.Worksheets.Add()
    }

    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $sheetSpec = $sheets[$i]
        $worksheet = $workbook.Worksheets.Item($i + 1)
        $worksheet.Name = $sheetSpec.Name

        $csvPath = Join-Path $tablesDir $sheetSpec.Csv
        $rows = Import-Csv -Path $csvPath
        if (-not $rows) {
            continue
        }

        $headers = @($rows[0].PSObject.Properties.Name)
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $worksheet.Cells.Item(1, $col + 1) = $headers[$col]
        }

        for ($row = 0; $row -lt $rows.Count; $row++) {
            for ($col = 0; $col -lt $headers.Count; $col++) {
                $value = $rows[$row].($headers[$col])
                $worksheet.Cells.Item($row + 2, $col + 1) = $value
            }
        }

        $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $headers.Count))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15

        $usedRange = $worksheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null
        $worksheet.Application.ActiveWindow.SplitRow = 1
        $worksheet.Application.ActiveWindow.FreezePanes = $true
    }

    while ($workbook.Worksheets.Count -gt $sheets.Count) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $workbook.SaveAs($outputPath, 51)
    Write-Output "Wrote $outputPath"
}
finally {
    if ($workbook) {
        $workbook.Close($true) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit() | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
