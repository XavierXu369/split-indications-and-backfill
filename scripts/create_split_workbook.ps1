[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$InputPath,
    [Parameter(Mandatory = $true)] [string]$OutputPath,
    [Parameter(Mandatory = $true)] [string]$SourceSheet,
    [string]$IdColumn = '序号',
    [string]$NameColumn = '药品',
    [string]$IndicationColumn = '获批适应症',
    [string]$SplitPattern = '[;；]',
    [string]$SplitSheetName = '拆分结果',
    [string]$UpdatedSheetName = '更新版拆分底稿'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ExcelColumn {
    param([Parameter(Mandatory = $true)][int]$ColumnNumber)
    $result = ''
    while ($ColumnNumber -gt 0) {
        $ColumnNumber--
        $result = [char](65 + ($ColumnNumber % 26)) + $result
        $ColumnNumber = [math]::Floor($ColumnNumber / 26)
    }
    return $result
}

function Get-HeaderIndex {
    param(
        [Parameter(Mandatory = $true)][string[]]$Headers,
        [Parameter(Mandatory = $true)][string]$HeaderName
    )
    $index = [array]::IndexOf($Headers, $HeaderName)
    if ($index -lt 0) { throw "Source sheet does not contain required field: $HeaderName" }
    return $index + 1
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input workbook was not found: $InputPath" }
if ([IO.Path]::GetFullPath($InputPath) -eq [IO.Path]::GetFullPath($OutputPath)) { throw 'OutputPath must be different from InputPath.' }
if (Test-Path -LiteralPath $OutputPath) { throw "Output workbook already exists and will not be overwritten: $OutputPath" }
if ([IO.Path]::GetExtension($InputPath) -notin '.xlsx', '.xlsm') { throw 'Input workbook must be .xlsx or .xlsm.' }
if ([IO.Path]::GetDirectoryName($OutputPath) -and -not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($OutputPath)))) { throw "Output directory does not exist: $([IO.Path]::GetDirectoryName($OutputPath))" }
if ($SourceSheet -eq $SplitSheetName -or $SourceSheet -eq $UpdatedSheetName) { throw 'SourceSheet must use a different name from the two output sheets.' }

$excel = $null
$sourceWorkbook = $null
$workbook = $null
$saved = $false
$createdOutput = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # Validate the original workbook in read-only mode before creating any output.
    $sourceWorkbook = $excel.Workbooks.Open($InputPath, $null, $true)
    try {
        $source = $sourceWorkbook.Worksheets.Item($SourceSheet)
    }
    catch {
        throw "Source sheet was not found: $SourceSheet"
    }
    $sourceRows = $source.UsedRange.Rows.Count
    $sourceColumns = $source.UsedRange.Columns.Count
    if ($sourceRows -lt 2) { throw 'Source sheet has no data rows.' }

    $headers = for ($column = 1; $column -le $sourceColumns; $column++) {
        ([string]$source.Cells.Item(1, $column).Text).Trim()
    }
    $idIndex = Get-HeaderIndex -Headers $headers -HeaderName $IdColumn
    $nameIndex = Get-HeaderIndex -Headers $headers -HeaderName $NameColumn
    $indicationIndex = Get-HeaderIndex -Headers $headers -HeaderName $IndicationColumn

    $keys = @{}
    for ($row = 2; $row -le $sourceRows; $row++) {
        $key = ([string]$source.Cells.Item($row, $idIndex).Text).Trim()
        $entityName = ([string]$source.Cells.Item($row, $nameIndex).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($key)) { throw "Blank key found at source row $row." }
        if ($keys.ContainsKey($key)) { throw "Duplicate key found: $key" }
        if ([string]::IsNullOrWhiteSpace($entityName)) { throw "Blank entity name found at source row $row." }
        $keys[$key] = $true
    }
    $sourceWorkbook.Close($false)
    $sourceWorkbook = $null

    Copy-Item -LiteralPath $InputPath -Destination $OutputPath
    $createdOutput = $true
    $workbook = $excel.Workbooks.Open($OutputPath)
    $source = $workbook.Worksheets.Item($SourceSheet)

    # Keep only the complete source sheet. Helper sheets are replaced by standard output sheets.
    for ($index = $workbook.Worksheets.Count; $index -ge 1; $index--) {
        $sheet = $workbook.Worksheets.Item($index)
        if ($sheet.Name -ne $SourceSheet) { $sheet.Delete() }
    }

    $splitSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $splitSheet.Name = $SplitSheetName
    $updatedSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $updatedSheet.Name = $UpdatedSheetName

    $splitSheet.Cells.Item(1, 1).Value2 = $IdColumn
    $splitSheet.Cells.Item(1, 2).Value2 = $NameColumn
    $splitSheet.Cells.Item(1, 3).Value2 = '适应症拆分结果'
    $updatedSheet.Cells.Item(1, 1).Value2 = $IdColumn
    $updatedSheet.Cells.Item(1, 2).Value2 = $NameColumn
    $updatedSheet.Cells.Item(1, 3).Value2 = '适应症拆分结果'

    $sourceToOutput = @{}
    $outputColumn = 4
    for ($sourceColumn = 1; $sourceColumn -le $sourceColumns; $sourceColumn++) {
        if ($sourceColumn -eq $idIndex -or $sourceColumn -eq $nameIndex) { continue }
        $updatedSheet.Cells.Item(1, $outputColumn).Value2 = $source.Cells.Item(1, $sourceColumn).Value2
        $sourceToOutput[$sourceColumn] = $outputColumn
        $outputColumn++
    }

    $outputRow = 2
    $splitSourceCount = 0
    $blankEntities = @()
    for ($sourceRow = 2; $sourceRow -le $sourceRows; $sourceRow++) {
        $key = ([string]$source.Cells.Item($sourceRow, $idIndex).Text).Trim()
        $entityName = ([string]$source.Cells.Item($sourceRow, $nameIndex).Text).Trim()
        $indication = ([string]$source.Cells.Item($sourceRow, $indicationIndex).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($indication)) {
            $parts = @('')
            $blankEntities += "$key-$entityName"
        }
        else {
            $parts = @($indication -split $SplitPattern | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($parts.Count -eq 0) { $parts = @('') }
            elseif ($parts.Count -gt 1) { $splitSourceCount++ }
        }

        foreach ($part in $parts) {
            $splitSheet.Cells.Item($outputRow, 1).Value2 = $key
            $splitSheet.Cells.Item($outputRow, 2).Value2 = $entityName
            $splitSheet.Cells.Item($outputRow, 3).Value2 = $part
            $updatedSheet.Cells.Item($outputRow, 1).Value2 = $key
            $updatedSheet.Cells.Item($outputRow, 2).Value2 = $entityName
            $updatedSheet.Cells.Item($outputRow, 3).Value2 = $part
            $outputRow++
        }
    }
    $lastOutputRow = $outputRow - 1
    $escapedSourceSheet = $SourceSheet.Replace("'", "''")
    $sourceKeyLetter = Get-ExcelColumn -ColumnNumber $idIndex

    foreach ($sourceColumn in $sourceToOutput.Keys) {
        $targetColumn = $sourceToOutput[$sourceColumn]
        $sourceReturnLetter = Get-ExcelColumn -ColumnNumber $sourceColumn
        $formula = ('=IFERROR(XLOOKUP($A2,''{0}''!${1}$2:${1}${2},''{0}''!${3}$2:${3}${2},""),"")' -f $escapedSourceSheet, $sourceKeyLetter, $sourceRows, $sourceReturnLetter)
        $firstFormulaCell = $updatedSheet.Cells.Item(2, $targetColumn)
        $firstFormulaCell.Formula2 = $formula
        $updatedSheet.Range($firstFormulaCell, $updatedSheet.Cells.Item($lastOutputRow, $targetColumn)).FillDown()
        $updatedSheet.Range($updatedSheet.Cells.Item(2, $targetColumn), $updatedSheet.Cells.Item($lastOutputRow, $targetColumn)).NumberFormat = $source.Cells.Item(2, $sourceColumn).NumberFormat
    }

    foreach ($sheet in @($splitSheet, $updatedSheet)) {
        $usedRange = $sheet.UsedRange
        $headerRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $usedRange.Columns.Count))
        $headerRange.Font.Bold = $true
        $headerRange.Font.Color = 16777215
        $headerRange.Interior.Color = 11895295
        $usedRange.WrapText = $true
        $usedRange.VerticalAlignment = -4160
        $null = $usedRange.AutoFilter()
        $sheet.Columns.Item(1).ColumnWidth = 9
        $sheet.Columns.Item(2).ColumnWidth = 28
        $sheet.Columns.Item(3).ColumnWidth = 24
        for ($column = 4; $column -le $usedRange.Columns.Count; $column++) { $sheet.Columns.Item($column).ColumnWidth = 16 }
    }

    $excel.Calculation = -4105 # xlCalculationAutomatic
    $excel.CalculateFull()

    $formulaRange = $updatedSheet.Range($updatedSheet.Cells.Item(2, 4), $updatedSheet.Cells.Item($lastOutputRow, $sourceColumns + 1))
    if ($formulaRange.HasFormula -ne $true) { throw 'Formula coverage check failed in 更新版拆分底稿.' }
    if ([string]::IsNullOrWhiteSpace([string]$updatedSheet.Cells.Item(2, 4).Formula2)) { throw 'Formula creation check failed in 更新版拆分底稿.' }

    $workbook.Save()
    $saved = $true

    [PSCustomObject]@{
        OutputPath = $OutputPath
        SourceSheet = $SourceSheet
        SourceRecords = $sourceRows - 1
        SourceColumns = $sourceColumns
        SplitRecords = $lastOutputRow - 1
        SplitSourceRecords = $splitSourceCount
        BlankIndicationEntities = $blankEntities -join '；'
        UpdatedColumns = $sourceColumns + 1
        FormulaCount = ($lastOutputRow - 1) * ($sourceColumns - 2)
    }
}
catch {
    if ($createdOutput -and -not $saved -and (Test-Path -LiteralPath $OutputPath)) { Remove-Item -LiteralPath $OutputPath -Force }
    throw
}
finally {
    if ($sourceWorkbook) { $sourceWorkbook.Close($false) }
    if ($workbook) { $workbook.Close($saved) }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
