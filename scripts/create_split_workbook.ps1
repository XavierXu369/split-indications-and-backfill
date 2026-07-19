[CmdletBinding()]
param(
    [ValidateSet('Inspect', 'Preview', 'Generate')]
    [string]$Mode = 'Inspect',
    [Parameter(Mandatory = $true)] [string]$InputPath,
    [string]$OutputPath,
    [Parameter(Mandatory = $true)] [string]$SourceSheet,
    [string]$IdColumn = '序号',
    [string]$NameColumn = '药品',
    [string]$IndicationColumn = '获批适应症',
    [string]$SplitPattern = '[;；]',
    [string]$SplitSheetName = '拆分结果',
    [string]$UpdatedSheetName = '更新版拆分底稿',
    [ValidateRange(1, 10)] [int]$PreviewCount = 2,
    [string]$ExpectedInputSha256,
    [string]$ExpectedSourceSignature,
    [switch]$CompleteSourceConfirmed
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

function Get-CanonicalValue {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    $typeName = $Value.GetType().FullName
    if ($Value -is [System.IFormattable]) {
        $text = $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture)
    }
    else {
        $text = [string]$Value
    }
    return "$typeName|$text"
}

function Get-FormulaText {
    param([Parameter(Mandatory = $true)]$Cell)
    try { return [string]$Cell.Formula2 }
    catch { return [string]$Cell.Formula }
}

function Get-WorksheetBounds {
    param([Parameter(Mandatory = $true)]$Sheet)
    $used = $Sheet.UsedRange
    $surface = [PSCustomObject]@{
        StartRow = [int]$used.Row
        StartColumn = [int]$used.Column
        Rows = [int]$used.Rows.Count
        Columns = [int]$used.Columns.Count
    }
    $missing = [Type]::Missing
    $lastRowCell = $Sheet.Cells.Find('*', $missing, -4123, $missing, 1, 2, $false, $missing, $false)
    $lastColumnCell = $Sheet.Cells.Find('*', $missing, -4123, $missing, 2, 2, $false, $missing, $false)
    if ($null -eq $lastRowCell -or $null -eq $lastColumnCell) {
        return [PSCustomObject]@{ Surface = $surface; LastRow = 0; LastColumn = 0 }
    }
    return [PSCustomObject]@{
        Surface = $surface
        LastRow = [int]$lastRowCell.Row
        LastColumn = [int]$lastColumnCell.Column
    }
}

function Get-WorksheetSignature {
    param(
        [Parameter(Mandatory = $true)]$Sheet,
        [Parameter(Mandatory = $true)][int]$LastRow,
        [Parameter(Mandatory = $true)][int]$LastColumn
    )
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append("$($Sheet.Name)|$LastRow|$LastColumn`n")
    for ($row = 1; $row -le $LastRow; $row++) {
        for ($column = 1; $column -le $LastColumn; $column++) {
            $cell = $Sheet.Cells.Item($row, $column)
            $valueToken = Get-CanonicalValue -Value $cell.Value2
            $formula = Get-FormulaText -Cell $cell
            $format = [string]$cell.NumberFormat
            [void]$builder.Append("$row|$column|$valueToken|$formula|$format`n")
        }
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($builder.ToString())
        $hash = $sha.ComputeHash($bytes)
        return (($hash | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $sha.Dispose()
    }
}

function Get-XlookupFormula {
    param(
        [Parameter(Mandatory = $true)][string]$EscapedSourceSheet,
        [Parameter(Mandatory = $true)][string]$SourceKeyLetter,
        [Parameter(Mandatory = $true)][string]$SourceReturnLetter,
        [Parameter(Mandatory = $true)][int]$SourceLastRow
    )
    return ('=IFERROR(LET(_v,XLOOKUP($A2,''{0}''!${1}$2:${1}${2},''{0}''!${3}$2:${3}${2},""),IF(_v="","",_v)),"")' -f $EscapedSourceSheet, $SourceKeyLetter, $SourceLastRow, $SourceReturnLetter)
}

function Join-ReportItems {
    param([object[]]$Items)
    if ($null -eq $Items -or $Items.Count -eq 0) { return '无' }
    return ($Items -join '；')
}

function Invoke-Preflight {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$InputSha256
    )
    $blockers = [Collections.Generic.List[string]]::new()
    $warnings = [Collections.Generic.List[string]]::new()
    $blankEntities = [Collections.Generic.List[string]]::new()
    $emptySegmentEntities = [Collections.Generic.List[string]]::new()
    $duplicateItemEntities = [Collections.Generic.List[string]]::new()
    $delimiterOnlyEntities = [Collections.Generic.List[string]]::new()
    $previewRows = [Collections.Generic.List[object]]::new()
    $dependencySheets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    try { $source = $Workbook.Worksheets.Item($SourceSheet) }
    catch { throw "Source sheet was not found: $SourceSheet" }

    $sheetNames = @()
    $hiddenSheets = @()
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $sheet = $Workbook.Worksheets.Item($index)
        $sheetNames += [string]$sheet.Name
        if ([int]$sheet.Visible -ne -1) { $hiddenSheets += [string]$sheet.Name }
    }
    $extraSheets = @($sheetNames | Where-Object { $_ -ne $SourceSheet })
    if ($extraSheets.Count -gt 0) { $warnings.Add("标准输出将删除非来源 Sheet：$($extraSheets -join '、')") }
    if ($hiddenSheets.Count -gt 0) { $warnings.Add("工作簿包含隐藏 Sheet：$($hiddenSheets -join '、')") }
    if ($Workbook.Names.Count -gt 0) { $warnings.Add("工作簿包含 $($Workbook.Names.Count) 个定义名称，生成前已纳入依赖风险提示") }

    $bounds = Get-WorksheetBounds -Sheet $source
    $sourceRows = $bounds.LastRow
    $sourceColumns = $bounds.LastColumn
    if ($sourceRows -lt 2) { $blockers.Add('来源 Sheet 没有数据行。') }
    if ($sourceColumns -lt 1) { $blockers.Add('来源 Sheet 没有有效字段。') }

    $headers = @()
    $headerCounts = @{}
    if ($sourceColumns -gt 0) {
        for ($column = 1; $column -le $sourceColumns; $column++) {
            $header = ([string]$source.Cells.Item(1, $column).Value2).Trim()
            $headers += $header
            if ([string]::IsNullOrWhiteSpace($header)) {
                $blockers.Add("第 $column 列表头为空。")
            }
            elseif ($headerCounts.ContainsKey($header)) {
                $headerCounts[$header]++
            }
            else {
                $headerCounts[$header] = 1
            }
        }
    }
    $duplicateHeaders = @($headerCounts.Keys | Where-Object { $headerCounts[$_] -gt 1 })
    foreach ($header in $duplicateHeaders) { $blockers.Add("表头重复：$header") }

    $idMatches = @($headers | ForEach-Object -Begin { $i = 0 } -Process { $i++; if ($_ -eq $IdColumn) { $i } })
    $nameMatches = @($headers | ForEach-Object -Begin { $i = 0 } -Process { $i++; if ($_ -eq $NameColumn) { $i } })
    $indicationMatches = @($headers | ForEach-Object -Begin { $i = 0 } -Process { $i++; if ($_ -eq $IndicationColumn) { $i } })
    if ($idMatches.Count -ne 1) { $blockers.Add("稳定键字段应恰好出现一次：$IdColumn") }
    if ($nameMatches.Count -ne 1) { $blockers.Add("实体名称字段应恰好出现一次：$NameColumn") }
    if ($indicationMatches.Count -ne 1) { $blockers.Add("适应症字段应恰好出现一次：$IndicationColumn") }
    if (@($headers | Where-Object { $_ -eq '适应症拆分结果' }).Count -gt 0) { $blockers.Add('来源表已包含“适应症拆分结果”字段，会造成输出字段冲突。') }

    $idIndex = if ($idMatches.Count -eq 1) { [int]$idMatches[0] } else { 0 }
    $nameIndex = if ($nameMatches.Count -eq 1) { [int]$nameMatches[0] } else { 0 }
    $indicationIndex = if ($indicationMatches.Count -eq 1) { [int]$indicationMatches[0] } else { 0 }

    $keyTokens = @{}
    $blankKeyCount = 0
    $duplicateKeys = [Collections.Generic.List[string]]::new()
    $blankNameCount = 0
    $blankIndicationCount = 0
    $delimiterRecordCount = 0
    $splitSourceCount = 0
    $expectedSplitRows = 0
    $mergedCellFound = $false

    if ($sourceRows -ge 2 -and $idIndex -gt 0 -and $nameIndex -gt 0 -and $indicationIndex -gt 0) {
        for ($row = 2; $row -le $sourceRows; $row++) {
            $keyCell = $source.Cells.Item($row, $idIndex)
            $nameCell = $source.Cells.Item($row, $nameIndex)
            $indicationCell = $source.Cells.Item($row, $indicationIndex)
            $keyValue = $keyCell.Value2
            $keyDisplay = ([string]$keyCell.Text).Trim()
            $name = ([string]$nameCell.Value2).Trim()
            $indication = ([string]$indicationCell.Value2).Trim()
            $entityLabel = if ([string]::IsNullOrWhiteSpace($keyDisplay)) { "第$row行-$name" } else { "$keyDisplay-$name" }

            if ($null -eq $keyValue -or [string]::IsNullOrWhiteSpace([string]$keyValue)) {
                $blankKeyCount++
            }
            else {
                $keyToken = Get-CanonicalValue -Value $keyValue
                if ($keyTokens.ContainsKey($keyToken)) { $duplicateKeys.Add($keyDisplay) }
                else { $keyTokens[$keyToken] = $row }
            }
            if ([string]::IsNullOrWhiteSpace($name)) { $blankNameCount++ }

            if ([string]::IsNullOrWhiteSpace($indication)) {
                $blankIndicationCount++
                $blankEntities.Add($entityLabel)
                $expectedSplitRows++
                $previewRows.Add([PSCustomObject]@{ Kind = 'Blank'; Key = $keyDisplay; Name = $name; Original = ''; Parts = @('') })
                continue
            }

            $hasDelimiter = [regex]::IsMatch($indication, $SplitPattern)
            if ($hasDelimiter) { $delimiterRecordCount++ }
            $rawParts = @([regex]::Split($indication, $SplitPattern))
            $nonblankParts = @($rawParts | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($rawParts.Count -gt $nonblankParts.Count) { $emptySegmentEntities.Add($entityLabel) }
            if ($nonblankParts.Count -eq 0) {
                $delimiterOnlyEntities.Add($entityLabel)
                continue
            }
            $expectedSplitRows += $nonblankParts.Count
            if ($nonblankParts.Count -gt 1) { $splitSourceCount++ }

            $seenParts = @{}
            $hasDuplicatePart = $false
            foreach ($part in $nonblankParts) {
                if ($seenParts.ContainsKey($part)) { $hasDuplicatePart = $true }
                else { $seenParts[$part] = $true }
            }
            if ($hasDuplicatePart) { $duplicateItemEntities.Add($entityLabel) }
            $kind = if ($nonblankParts.Count -gt 1) { 'Split' } else { 'Single' }
            $previewRows.Add([PSCustomObject]@{ Kind = $kind; Key = $keyDisplay; Name = $name; Original = $indication; Parts = $nonblankParts })
        }
    }

    if ($blankKeyCount -gt 0) { $blockers.Add("发现 $blankKeyCount 个空稳定键。") }
    if ($duplicateKeys.Count -gt 0) { $blockers.Add("发现重复稳定键：$($duplicateKeys -join '、')") }
    if ($blankNameCount -gt 0) { $blockers.Add("发现 $blankNameCount 个空实体名称。") }
    if ($delimiterOnlyEntities.Count -gt 0) { $blockers.Add("非空适应症仅包含分隔符/空白：$($delimiterOnlyEntities -join '、')") }
    if ($blankIndicationCount -gt 0) { $warnings.Add("$blankIndicationCount 条空适应症将各保留一行空拆分结果") }
    if ($emptySegmentEntities.Count -gt 0) { $warnings.Add("以下记录含连续/首尾分隔符，空片段将忽略：$($emptySegmentEntities -join '、')") }
    if ($duplicateItemEntities.Count -gt 0) { $warnings.Add("以下记录含重复适应症词项，将原样保留重复行：$($duplicateItemEntities -join '、')") }

    if ($sourceRows -gt 0 -and $sourceColumns -gt 0) {
        for ($row = 1; $row -le $sourceRows -and -not $mergedCellFound; $row++) {
            for ($column = 1; $column -le $sourceColumns; $column++) {
                if ($source.Cells.Item($row, $column).MergeCells) { $mergedCellFound = $true; break }
            }
        }
    }
    if ($mergedCellFound) { $blockers.Add('有效来源区域包含合并单元格。') }

    if ($extraSheets.Count -gt 0 -and $sourceRows -gt 0 -and $sourceColumns -gt 0) {
        $formulaCells = $null
        try { $formulaCells = $source.Range($source.Cells.Item(1, 1), $source.Cells.Item($sourceRows, $sourceColumns)).SpecialCells(-4123) }
        catch { $formulaCells = $null }
        if ($null -ne $formulaCells) {
            foreach ($cell in $formulaCells.Cells) {
                $formula = Get-FormulaText -Cell $cell
                foreach ($otherSheet in $extraSheets) {
                    $escapedExcelName = $otherSheet.Replace("'", "''")
                    $pattern = "(?i)(?:'" + [regex]::Escape($escapedExcelName) + "'|" + [regex]::Escape($otherSheet) + ")!"
                    if ([regex]::IsMatch($formula, $pattern)) { [void]$dependencySheets.Add($otherSheet) }
                }
            }
        }
    }
    if ($dependencySheets.Count -gt 0) { $blockers.Add("来源 Sheet 的公式依赖将被删除的 Sheet：$(@($dependencySheets) -join '、')") }

    $signature = if ($sourceRows -gt 0 -and $sourceColumns -gt 0) { Get-WorksheetSignature -Sheet $source -LastRow $sourceRows -LastColumn $sourceColumns } else { '' }
    return [PSCustomObject]@{
        InputSha256 = $InputSha256
        SourceSignature = $signature
        SourceRows = $sourceRows
        SourceColumns = $sourceColumns
        SourceRecords = [math]::Max(0, $sourceRows - 1)
        SurfaceStart = "R$($bounds.Surface.StartRow)C$($bounds.Surface.StartColumn)"
        SurfaceRows = $bounds.Surface.Rows
        SurfaceColumns = $bounds.Surface.Columns
        SheetNames = $sheetNames
        ExtraSheets = $extraSheets
        HiddenSheets = $hiddenSheets
        DefinedNameCount = [int]$Workbook.Names.Count
        Headers = $headers
        IdIndex = $idIndex
        NameIndex = $nameIndex
        IndicationIndex = $indicationIndex
        BlankKeyCount = $blankKeyCount
        DuplicateKeys = @($duplicateKeys)
        BlankNameCount = $blankNameCount
        BlankIndicationCount = $blankIndicationCount
        BlankEntities = @($blankEntities)
        DelimiterRecordCount = $delimiterRecordCount
        SplitSourceCount = $splitSourceCount
        EmptySegmentEntities = @($emptySegmentEntities)
        DuplicateItemEntities = @($duplicateItemEntities)
        ExpectedSplitRows = $expectedSplitRows
        DependencySheets = @($dependencySheets)
        PreviewRows = @($previewRows)
        Blockers = @($blockers)
        Warnings = @($warnings)
        CanProceed = ($blockers.Count -eq 0)
    }
}

function Get-PublicInspectionResult {
    param([Parameter(Mandatory = $true)]$Preflight)
    return [PSCustomObject]@{
        Mode = $Mode
        CanProceed = $Preflight.CanProceed
        InputPath = $InputPath
        InputSha256 = $Preflight.InputSha256
        InputBytes = (Get-Item -LiteralPath $InputPath).Length
        InputLastWriteTime = (Get-Item -LiteralPath $InputPath).LastWriteTime
        SourceSheet = $SourceSheet
        SourceSignature = $Preflight.SourceSignature
        SurfaceUsedRange = "$($Preflight.SurfaceStart), $($Preflight.SurfaceRows)×$($Preflight.SurfaceColumns)"
        EffectiveDimensions = "$($Preflight.SourceRows)×$($Preflight.SourceColumns)（含表头）"
        SourceRecords = $Preflight.SourceRecords
        SourceFields = $Preflight.SourceColumns
        WorkbookSheets = Join-ReportItems -Items $Preflight.SheetNames
        ExtraSheets = Join-ReportItems -Items $Preflight.ExtraSheets
        HiddenSheets = Join-ReportItems -Items $Preflight.HiddenSheets
        DefinedNames = $Preflight.DefinedNameCount
        BlankKeys = $Preflight.BlankKeyCount
        DuplicateKeys = Join-ReportItems -Items $Preflight.DuplicateKeys
        BlankEntityNames = $Preflight.BlankNameCount
        BlankIndications = $Preflight.BlankIndicationCount
        BlankIndicationEntities = Join-ReportItems -Items $Preflight.BlankEntities
        DelimiterRecords = $Preflight.DelimiterRecordCount
        ActuallySplitRecords = $Preflight.SplitSourceCount
        ExpectedSplitRows = $Preflight.ExpectedSplitRows
        EmptySegmentWarnings = Join-ReportItems -Items $Preflight.EmptySegmentEntities
        DuplicateItemWarnings = Join-ReportItems -Items $Preflight.DuplicateItemEntities
        SourceSheetDependencies = Join-ReportItems -Items $Preflight.DependencySheets
        Blockers = Join-ReportItems -Items $Preflight.Blockers
        Warnings = Join-ReportItems -Items $Preflight.Warnings
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input workbook was not found: $InputPath" }
$InputPath = [IO.Path]::GetFullPath($InputPath)
$inputExtension = [IO.Path]::GetExtension($InputPath).ToLowerInvariant()
if ($inputExtension -notin '.xlsx', '.xlsm') { throw 'Input workbook must be .xlsx or .xlsm.' }
try { $null = [regex]::new($SplitPattern) }
catch { throw "SplitPattern is not a valid regular expression: $SplitPattern" }
if ($SourceSheet -eq $SplitSheetName -or $SourceSheet -eq $UpdatedSheetName -or $SplitSheetName -eq $UpdatedSheetName) { throw 'Source and output sheet names must be distinct.' }

if ($Mode -eq 'Generate') {
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'OutputPath is required in Generate mode.' }
    if (-not $CompleteSourceConfirmed) { throw 'Generate mode requires -CompleteSourceConfirmed.' }
    if ([string]::IsNullOrWhiteSpace($ExpectedInputSha256) -or [string]::IsNullOrWhiteSpace($ExpectedSourceSignature)) { throw 'Generate mode requires ExpectedInputSha256 and ExpectedSourceSignature from the approved inspection.' }
    $OutputPath = [IO.Path]::GetFullPath($OutputPath)
    if ($InputPath -eq $OutputPath) { throw 'OutputPath must be different from InputPath.' }
    if (Test-Path -LiteralPath $OutputPath) { throw "Output workbook already exists and will not be overwritten: $OutputPath" }
    if ([IO.Path]::GetExtension($OutputPath).ToLowerInvariant() -ne $inputExtension) { throw 'Input and output workbook extensions must match.' }
    $outputDirectory = [IO.Path]::GetDirectoryName($OutputPath)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) { throw "Output directory does not exist: $outputDirectory" }
}
elseif ($Mode -eq 'Preview' -and -not $CompleteSourceConfirmed) {
    throw 'Preview mode requires -CompleteSourceConfirmed.'
}

$inputSha256 = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
$excel = $null
$sourceWorkbook = $null
$workbook = $null
$verifyWorkbook = $null
$createdOutput = $false
$verified = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $sourceWorkbook = $excel.Workbooks.Open($InputPath, $null, $true)
    $preflight = Invoke-Preflight -Workbook $sourceWorkbook -InputSha256 $inputSha256
    if ($ExpectedInputSha256 -and $ExpectedInputSha256.ToLowerInvariant() -ne $preflight.InputSha256) {
        $preflight.Blockers += '输入文件 SHA-256 与已批准版本不一致。'
        $preflight.CanProceed = $false
    }
    if ($ExpectedSourceSignature -and $ExpectedSourceSignature.ToLowerInvariant() -ne $preflight.SourceSignature) {
        $preflight.Blockers += '来源 Sheet 签名与已批准版本不一致。'
        $preflight.CanProceed = $false
    }

    if ($Mode -eq 'Inspect') {
        Get-PublicInspectionResult -Preflight $preflight
        return
    }
    if (-not $preflight.CanProceed) { throw "Preflight failed: $(Join-ReportItems -Items $preflight.Blockers)" }

    if ($Mode -eq 'Preview') {
        $selected = [Collections.Generic.List[object]]::new()
        foreach ($kind in @('Split', 'Blank', 'Single')) {
            foreach ($candidate in @($preflight.PreviewRows | Where-Object { $_.Kind -eq $kind })) {
                if ($selected.Count -ge $PreviewCount) { break }
                $selected.Add($candidate)
            }
            if ($selected.Count -ge $PreviewCount) { break }
        }
        $previewLines = @()
        foreach ($candidate in $selected) {
            $partsText = if ($candidate.Parts.Count -eq 1 -and [string]::IsNullOrWhiteSpace([string]$candidate.Parts[0])) { '<空白，保留一行>' } else { $candidate.Parts -join ' | ' }
            $previewLines += "$($candidate.Key)-$($candidate.Name)：[$($candidate.Original)] → [$partsText]"
        }
        $exampleSourceColumn = @(1..$preflight.SourceColumns | Where-Object { $_ -ne $preflight.IdIndex -and $_ -ne $preflight.NameIndex } | Select-Object -First 1)[0]
        $formulaExample = Get-XlookupFormula -EscapedSourceSheet $SourceSheet.Replace("'", "''") -SourceKeyLetter (Get-ExcelColumn $preflight.IdIndex) -SourceReturnLetter (Get-ExcelColumn $exampleSourceColumn) -SourceLastRow $preflight.SourceRows
        [PSCustomObject]@{
            Mode = 'Preview'
            CanProceed = $true
            InputSha256 = $preflight.InputSha256
            SourceSignature = $preflight.SourceSignature
            ExpectedSplitRows = $preflight.ExpectedSplitRows
            Examples = $previewLines -join "`n"
            FormulaExample = $formulaExample
            Warnings = Join-ReportItems -Items $preflight.Warnings
            NextGate = '等待用户明确批准后，方可运行 Generate。'
        }
        return
    }

    $sourceWorkbook.Close($false)
    $sourceWorkbook = $null
    Copy-Item -LiteralPath $InputPath -Destination $OutputPath
    $createdOutput = $true
    $workbook = $excel.Workbooks.Open($OutputPath)
    $source = $workbook.Worksheets.Item($SourceSheet)

    for ($index = $workbook.Worksheets.Count; $index -ge 1; $index--) {
        $sheet = $workbook.Worksheets.Item($index)
        if ($sheet.Name -ne $SourceSheet) { $sheet.Delete() }
    }
    $splitSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $splitSheet.Name = $SplitSheetName
    $updatedSheet = $workbook.Worksheets.Add([Type]::Missing, $workbook.Worksheets.Item($workbook.Worksheets.Count))
    $updatedSheet.Name = $UpdatedSheetName

    foreach ($sheet in @($splitSheet, $updatedSheet)) {
        $sheet.Cells.Item(1, 1).Value2 = $IdColumn
        $sheet.Cells.Item(1, 2).Value2 = $NameColumn
        $sheet.Cells.Item(1, 3).Value2 = '适应症拆分结果'
    }

    $sourceToOutput = [Collections.Generic.List[object]]::new()
    $targetColumn = 4
    for ($sourceColumn = 1; $sourceColumn -le $preflight.SourceColumns; $sourceColumn++) {
        if ($sourceColumn -eq $preflight.IdIndex -or $sourceColumn -eq $preflight.NameIndex) { continue }
        $updatedSheet.Cells.Item(1, $targetColumn).Value2 = $source.Cells.Item(1, $sourceColumn).Value2
        $sourceToOutput.Add([PSCustomObject]@{ SourceColumn = $sourceColumn; TargetColumn = $targetColumn })
        $targetColumn++
    }

    $outputSourceRows = [Collections.Generic.List[int]]::new()
    $firstOutputBySource = @{}
    $outputRow = 2
    for ($sourceRow = 2; $sourceRow -le $preflight.SourceRows; $sourceRow++) {
        $keyCell = $source.Cells.Item($sourceRow, $preflight.IdIndex)
        $nameCell = $source.Cells.Item($sourceRow, $preflight.NameIndex)
        $keyValue = $keyCell.Value2
        $nameValue = $nameCell.Value2
        $indication = ([string]$source.Cells.Item($sourceRow, $preflight.IndicationIndex).Value2).Trim()
        if ([string]::IsNullOrWhiteSpace($indication)) { $parts = @('') }
        else { $parts = @([regex]::Split($indication, $SplitPattern) | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }
        $firstOutputBySource[$sourceRow] = $outputRow
        foreach ($part in $parts) {
            foreach ($targetSheet in @($splitSheet, $updatedSheet)) {
                $targetSheet.Cells.Item($outputRow, 1).NumberFormat = $keyCell.NumberFormat
                $targetSheet.Cells.Item($outputRow, 2).NumberFormat = $nameCell.NumberFormat
                if ($keyValue -is [string]) { $targetSheet.Cells.Item($outputRow, 1).Value2 = [string]$keyValue }
                else { $targetSheet.Cells.Item($outputRow, 1).Value = $keyValue }
                $targetSheet.Cells.Item($outputRow, 2).Value2 = [string]$nameValue
                $targetSheet.Cells.Item($outputRow, 3).Value2 = [string]$part
            }
            $outputSourceRows.Add($sourceRow)
            $outputRow++
        }
    }
    $lastOutputRow = $outputRow - 1
    if (($lastOutputRow - 1) -ne $preflight.ExpectedSplitRows) { throw 'Generated row count does not equal the approved preflight expectation.' }

    $escapedSourceSheet = $SourceSheet.Replace("'", "''")
    $sourceKeyLetter = Get-ExcelColumn -ColumnNumber $preflight.IdIndex
    foreach ($mapping in $sourceToOutput) {
        $sourceReturnLetter = Get-ExcelColumn -ColumnNumber $mapping.SourceColumn
        $formula = Get-XlookupFormula -EscapedSourceSheet $escapedSourceSheet -SourceKeyLetter $sourceKeyLetter -SourceReturnLetter $sourceReturnLetter -SourceLastRow $preflight.SourceRows
        $firstFormulaCell = $updatedSheet.Cells.Item(2, $mapping.TargetColumn)
        $firstFormulaCell.Formula2 = $formula
        $formulaTargetRange = $updatedSheet.Range($firstFormulaCell, $updatedSheet.Cells.Item($lastOutputRow, $mapping.TargetColumn))
        $null = $formulaTargetRange.FillDown()

        $formats = @{}
        for ($sourceRow = 2; $sourceRow -le $preflight.SourceRows; $sourceRow++) {
            $format = [string]$source.Cells.Item($sourceRow, $mapping.SourceColumn).NumberFormat
            if (-not $formats.ContainsKey($format)) { $formats[$format] = $true }
        }
        if ($formats.Count -eq 1) {
            $formulaTargetRange.NumberFormat = @($formats.Keys)[0]
        }
        else {
            for ($offset = 0; $offset -lt $outputSourceRows.Count; $offset++) {
                $sourceRow = $outputSourceRows[$offset]
                $updatedSheet.Cells.Item($offset + 2, $mapping.TargetColumn).NumberFormat = $source.Cells.Item($sourceRow, $mapping.SourceColumn).NumberFormat
            }
        }
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
        $sheet.Columns.Item(1).ColumnWidth = 12
        $sheet.Columns.Item(2).ColumnWidth = 28
        $sheet.Columns.Item(3).ColumnWidth = 28
        for ($column = 4; $column -le $usedRange.Columns.Count; $column++) { $sheet.Columns.Item($column).ColumnWidth = 16 }
    }

    $verificationTarget = $null
    $blankReturnTarget = $null
    foreach ($mapping in $sourceToOutput) {
        for ($sourceRow = 2; $sourceRow -le $preflight.SourceRows; $sourceRow++) {
            $sourceValue = $source.Cells.Item($sourceRow, $mapping.SourceColumn).Value2
            if ($null -eq $verificationTarget -and $null -ne $sourceValue -and -not [string]::IsNullOrWhiteSpace([string]$sourceValue)) {
                $verificationTarget = [PSCustomObject]@{ SourceRow = $sourceRow; SourceColumn = $mapping.SourceColumn; OutputRow = $firstOutputBySource[$sourceRow]; OutputColumn = $mapping.TargetColumn }
            }
            if ($null -eq $blankReturnTarget -and ($null -eq $sourceValue -or [string]::IsNullOrWhiteSpace([string]$sourceValue))) {
                $blankReturnTarget = [PSCustomObject]@{ SourceRow = $sourceRow; SourceColumn = $mapping.SourceColumn; OutputRow = $firstOutputBySource[$sourceRow]; OutputColumn = $mapping.TargetColumn }
            }
        }
    }
    if ($null -eq $verificationTarget) { throw 'No nonblank source value is available for formula return-value verification.' }

    $excel.Calculation = -4105
    $null = $excel.CalculateFullRebuild()
    $workbook.Save()
    $workbook.Close($true)
    $workbook = $null

    $verifyWorkbook = $excel.Workbooks.Open($OutputPath, $null, $true)
    $actualSheetNames = @()
    for ($index = 1; $index -le $verifyWorkbook.Worksheets.Count; $index++) { $actualSheetNames += [string]$verifyWorkbook.Worksheets.Item($index).Name }
    $expectedSheetNames = @($SourceSheet, $SplitSheetName, $UpdatedSheetName)
    if ($actualSheetNames.Count -ne 3 -or (($actualSheetNames -join '|') -ne ($expectedSheetNames -join '|'))) { throw "Three-sheet verification failed: $($actualSheetNames -join '、')" }

    $verifySource = $verifyWorkbook.Worksheets.Item($SourceSheet)
    $verifySplit = $verifyWorkbook.Worksheets.Item($SplitSheetName)
    $verifyUpdated = $verifyWorkbook.Worksheets.Item($UpdatedSheetName)
    $verifySourceSignature = Get-WorksheetSignature -Sheet $verifySource -LastRow $preflight.SourceRows -LastColumn $preflight.SourceColumns
    if ($verifySourceSignature -ne $preflight.SourceSignature) { throw 'Source-sheet preservation verification failed.' }

    $splitBounds = Get-WorksheetBounds -Sheet $verifySplit
    $updatedBounds = Get-WorksheetBounds -Sheet $verifyUpdated
    if (($splitBounds.LastRow - 1) -ne $preflight.ExpectedSplitRows -or $splitBounds.LastColumn -ne 3) { throw '拆分结果 dimensions verification failed.' }
    if (($updatedBounds.LastRow - 1) -ne $preflight.ExpectedSplitRows -or $updatedBounds.LastColumn -ne ($preflight.SourceColumns + 1)) { throw '更新版拆分底稿 dimensions verification failed.' }

    $formulaRange = $verifyUpdated.Range($verifyUpdated.Cells.Item(2, 4), $verifyUpdated.Cells.Item($updatedBounds.LastRow, $updatedBounds.LastColumn))
    $formulaCells = $null
    try { $formulaCells = $formulaRange.SpecialCells(-4123) }
    catch { $formulaCells = $null }
    $actualFormulaCount = if ($null -eq $formulaCells) { 0 } else { [long]$formulaCells.CountLarge }
    $expectedFormulaCount = [long]$preflight.ExpectedSplitRows * [long]($preflight.SourceColumns - 2)
    if ($actualFormulaCount -ne $expectedFormulaCount -or [long]$formulaRange.Cells.CountLarge -ne $expectedFormulaCount) { throw "Formula coverage verification failed: expected $expectedFormulaCount, found $actualFormulaCount." }

    $formulaErrorCount = 0
    try { $formulaErrorCount = [long]$formulaRange.SpecialCells(-4123, 16).CountLarge }
    catch { $formulaErrorCount = 0 }
    if ($formulaErrorCount -gt 0) { throw "Formula error verification failed: $formulaErrorCount error cells." }

    $expectedReturn = $verifySource.Cells.Item($verificationTarget.SourceRow, $verificationTarget.SourceColumn).Value2
    $actualReturn = $verifyUpdated.Cells.Item($verificationTarget.OutputRow, $verificationTarget.OutputColumn).Value2
    if ((Get-CanonicalValue $expectedReturn) -ne (Get-CanonicalValue $actualReturn)) { throw 'Formula return-value verification failed.' }
    $blankReturnPreserved = '不适用（来源无空白回填字段）'
    if ($null -ne $blankReturnTarget) {
        $blankReturnedValue = $verifyUpdated.Cells.Item($blankReturnTarget.OutputRow, $blankReturnTarget.OutputColumn).Value2
        if ($null -ne $blankReturnedValue -and -not [string]::IsNullOrWhiteSpace([string]$blankReturnedValue)) { throw 'Blank source value was not preserved as blank.' }
        $blankReturnPreserved = '通过'
    }

    $blankSplitRows = 0
    $splitValueAudit = [Collections.Generic.List[string]]::new()
    for ($row = 2; $row -le $splitBounds.LastRow; $row++) {
        $splitValue = $verifySplit.Cells.Item($row, 3).Value2
        $splitValueAudit.Add("$row=$(Get-CanonicalValue $splitValue)")
        if ($null -eq $splitValue -or [string]::IsNullOrWhiteSpace([string]$splitValue)) { $blankSplitRows++ }
    }
    if ($blankSplitRows -ne $preflight.BlankIndicationCount) { throw "Blank-indication retention verification failed: expected $($preflight.BlankIndicationCount), found $blankSplitRows; $($splitValueAudit -join ' | ')" }
    $originalIndicationTarget = @($sourceToOutput | Where-Object { $_.SourceColumn -eq $preflight.IndicationIndex })[0].TargetColumn
    if ([string]$verifyUpdated.Cells.Item(1, $originalIndicationTarget).Value2 -ne $IndicationColumn) { throw 'Original indication field retention verification failed.' }

    $verifyWorkbook.Close($false)
    $verifyWorkbook = $null
    $finalInputSha256 = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($finalInputSha256 -ne $preflight.InputSha256) { throw 'Input workbook changed during generation.' }
    $verified = $true

    [PSCustomObject]@{
        Mode = 'Generate'
        VerificationPassed = $true
        OutputPath = $OutputPath
        InputSha256 = $preflight.InputSha256
        SourceSignature = $preflight.SourceSignature
        SourceSheetPreserved = $true
        SourceRecords = $preflight.SourceRecords
        SourceFields = $preflight.SourceColumns
        SplitRecords = $preflight.ExpectedSplitRows
        AddedRows = $preflight.ExpectedSplitRows - $preflight.SourceRecords
        OutputFields = $preflight.SourceColumns + 1
        DelimiterRecords = $preflight.DelimiterRecordCount
        ActuallySplitRecords = $preflight.SplitSourceCount
        BlankIndicationsRetained = $preflight.BlankIndicationCount
        BlankIndicationEntities = Join-ReportItems -Items $preflight.BlankEntities
        EmptySegmentWarnings = Join-ReportItems -Items $preflight.EmptySegmentEntities
        DuplicateItemWarnings = Join-ReportItems -Items $preflight.DuplicateItemEntities
        FormulaCount = $actualFormulaCount
        FormulaErrors = $formulaErrorCount
        FormulaReturnValueCheck = '通过'
        BlankReturnCheck = $blankReturnPreserved
        SheetCheck = '通过（恰好三个 Sheet）'
        Warnings = Join-ReportItems -Items $preflight.Warnings
    }
}
catch {
    if ($verifyWorkbook) { try { $verifyWorkbook.Close($false) } catch { }; $verifyWorkbook = $null }
    if ($workbook) { try { $workbook.Close($false) } catch { }; $workbook = $null }
    if ($sourceWorkbook) { try { $sourceWorkbook.Close($false) } catch { }; $sourceWorkbook = $null }
    if ($createdOutput -and -not $verified -and (Test-Path -LiteralPath $OutputPath)) {
        try { Remove-Item -LiteralPath $OutputPath -Force } catch { }
    }
    throw
}
finally {
    if ($sourceWorkbook) { try { $sourceWorkbook.Close($false) } catch { } }
    if ($workbook) { try { $workbook.Close($false) } catch { } }
    if ($verifyWorkbook) { try { $verifyWorkbook.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
