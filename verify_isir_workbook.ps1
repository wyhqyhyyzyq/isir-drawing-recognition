param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [int] $ExpectedCharacteristics,

    [Parameter(Mandatory = $true)]
    [string] $MappingPath,

    [Parameter(Mandatory = $true)]
    [string] $ExpectedPartName,

    [Parameter(Mandatory = $true)]
    [string] $ExpectedDrawingNo,

    [Parameter(Mandatory = $true)]
    [string] $ExpectedRevision,

    [Parameter(Mandatory = $true)]
    [string] $ExpectedMaterial,

    [string] $ExpectedSupplier = ""
)

$ErrorActionPreference = "Stop"
$errors = New-Object "System.Collections.Generic.List[string]"
$excel = $null
$workbook = $null

function Add-Error {
    param([string] $Message)
    $errors.Add($Message)
}

function Text-AtMergedTopLeft {
    param($Range)
    return ([string] $Range.MergeArea.Cells.Item(1).Text).Trim()
}

function Assert-ExpectedText {
    param(
        [string] $Label,
        [string] $Actual,
        [string] $Expected
    )
    if ($Expected -ne "" -and $Actual -ne $Expected) {
        Add-Error "$Label mismatch: expected '$Expected', got '$Actual'"
    }
}

try {
    $resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $resolvedMappingPath = (Resolve-Path -LiteralPath $MappingPath).Path
    $mapping = Get-Content -LiteralPath $resolvedMappingPath -Encoding UTF8 -Raw | ConvertFrom-Json
    $mappingRows = @($mapping.rows) + @($mapping.non_balloon_rows)
    if ($mappingRows.Count -ne $ExpectedCharacteristics) {
        Add-Error "mapping contains $($mappingRows.Count) rows, expected $ExpectedCharacteristics"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($resolvedPath, 0, $true)

    $expectedSheets = @(
        "ISIR 初期样品检验报告",
        "全尺寸检验数据",
        "全尺寸检验数据 (2)"
    )
    if ($workbook.Worksheets.Count -ne 3) {
        Add-Error "workbook must contain exactly 3 worksheets"
    }
    for ($index = 1; $index -le [Math]::Min(3, $workbook.Worksheets.Count); $index++) {
        $actualName = [string] $workbook.Worksheets.Item($index).Name
        if ($actualName -ne $expectedSheets[$index - 1]) {
            Add-Error "worksheet $index must be '$($expectedSheets[$index - 1])', got '$actualName'"
        }
    }

    $cover = $workbook.Worksheets.Item($expectedSheets[0])
    $data1 = $workbook.Worksheets.Item($expectedSheets[1])
    $data2 = $workbook.Worksheets.Item($expectedSheets[2])

    if ((Text-AtMergedTopLeft $cover.Range("B1")) -ne "") {
        Add-Error "customer field B1 must be blank"
    }
    Assert-ExpectedText "part name" (Text-AtMergedTopLeft $cover.Range("H8")) $ExpectedPartName
    Assert-ExpectedText "drawing number" (Text-AtMergedTopLeft $cover.Range("N8")) $ExpectedDrawingNo
    Assert-ExpectedText "revision" (Text-AtMergedTopLeft $cover.Range("X6")) $ExpectedRevision
    Assert-ExpectedText "material" (Text-AtMergedTopLeft $cover.Range("O13")) $ExpectedMaterial
    Assert-ExpectedText "supplier" (Text-AtMergedTopLeft $cover.Range("D7")) $ExpectedSupplier
    if ($ExpectedSupplier -eq "" -and (Text-AtMergedTopLeft $cover.Range("D7")) -ne "") {
        Add-Error "supplier D7 must be cleared when no current supplier is provided"
    }

    $formulaErrors = 0
    foreach ($sheet in @($cover, $data1, $data2)) {
        try {
            $errorCells = $sheet.UsedRange.SpecialCells(-4123, 16)
            if ($null -ne $errorCells) {
                $formulaErrors += $errorCells.Count
            }
        }
        catch {
        }

        $sheet.Activate()
        $window = $workbook.Windows.Item(1)
        if ($window.View -ne 1) {
            Add-Error "$($sheet.Name) must use normal view"
        }
        if ($window.Zoom -ne 100) {
            Add-Error "$($sheet.Name) must use 100% zoom"
        }
        if ($window.ScrollRow -ne 1 -or $window.ScrollColumn -ne 1) {
            Add-Error "$($sheet.Name) must open at row 1, column 1"
        }
    }
    if ($formulaErrors -ne 0) {
        Add-Error "workbook contains $formulaErrors formula errors"
    }

    foreach ($sheet in @($data1, $data2)) {
        if ([Math]::Abs($sheet.Columns.Item("B").ColumnWidth - 4.5) -gt 0.01) {
            Add-Error "$($sheet.Name) column B width must be 4.5"
        }
        if ([Math]::Abs($sheet.Rows("6:47").RowHeight - 22) -gt 0.01) {
            Add-Error "$($sheet.Name) rows 6:47 height must be 22"
        }
        foreach ($borderIndex in @(7, 8, 9, 10, 11, 12)) {
            if ($sheet.Range("B3:AF47").Borders.Item($borderIndex).Weight -ne 2) {
                Add-Error "$($sheet.Name) border $borderIndex must use thin weight 2"
            }
        }
    }

    $expectedFirst = [Math]::Min($ExpectedCharacteristics, 21)
    $expectedSecond = [Math]::Max($ExpectedCharacteristics - 21, 0)
    $actualFirst = $excel.WorksheetFunction.CountA($data1.Range("B6:B47"))
    $actualSecond = $excel.WorksheetFunction.CountA($data2.Range("B6:B47"))
    if ($actualFirst -ne $expectedFirst) {
        Add-Error "first data sheet expected $expectedFirst items, got $actualFirst"
    }
    if ($actualSecond -ne $expectedSecond) {
        Add-Error "continuation sheet expected $expectedSecond items, got $actualSecond"
    }
    $expectedImageFirst = @($mappingRows[0..([Math]::Max($expectedFirst - 1, 0))] | Where-Object { $_.spec_render_mode -eq "image_crop" }).Count
    $secondRows = if ($expectedSecond -gt 0) { @($mappingRows[21..($mappingRows.Count - 1)]) } else { @() }
    $expectedImageSecond = @($secondRows | Where-Object { $_.spec_render_mode -eq "image_crop" }).Count
    if ($excel.WorksheetFunction.CountA($data1.Range("B6:D47")) -ne ((3 * $expectedFirst) - $expectedImageFirst)) {
        Add-Error "first data sheet B:D contains residual or missing values"
    }
    if ($excel.WorksheetFunction.CountA($data2.Range("B6:D47")) -ne ((3 * $expectedSecond) - $expectedImageSecond)) {
        Add-Error "continuation sheet B:D contains residual or missing values"
    }
    if ($expectedSecond -eq 0 -and $excel.WorksheetFunction.CountA($data2.Range("B6:AC47")) -ne 0) {
        Add-Error "unused continuation sheet B6:AC47 must be completely blank"
    }

    for ($itemIndex = 0; $itemIndex -lt $mappingRows.Count; $itemIndex++) {
        $sheet = if ($itemIndex -lt 21) { $data1 } else { $data2 }
        $row = 6 + (2 * ($itemIndex % 21))
        $actualNo = Text-AtMergedTopLeft $sheet.Range("B$row")
        $actualType = Text-AtMergedTopLeft $sheet.Range("C$row")
        $actualRequirement = Text-AtMergedTopLeft $sheet.Range("D$row")
        $expectedRow = $mappingRows[$itemIndex]
        if ($actualNo -ne [string] $expectedRow.no) {
            Add-Error "item $($itemIndex + 1) number mismatch: expected '$($expectedRow.no)', got '$actualNo'"
        }
        if ($actualType -ne [string] $expectedRow.table_type) {
            Add-Error "item $($itemIndex + 1) type mismatch: expected '$($expectedRow.table_type)', got '$actualType'"
        }
        if ($expectedRow.spec_render_mode -eq "image_crop") {
            if ($actualRequirement -ne "") {
                Add-Error "item $($itemIndex + 1) GD&T specification text must be blank, got '$actualRequirement'"
            }
            $expectedShapeName = "ISIR规格图_$($expectedRow.no)"
            $specShape = $null
            for ($shapeIndex = 1; $shapeIndex -le $sheet.Shapes.Count; $shapeIndex++) {
                $candidate = $sheet.Shapes.Item($shapeIndex)
                if ([string] $candidate.Name -eq $expectedShapeName) { $specShape = $candidate; break }
            }
            if ($null -eq $specShape) {
                Add-Error "item $($itemIndex + 1) missing specification image '$expectedShapeName'"
            }
            else {
                $specCell = $sheet.Range("D$row").MergeArea
                if ($specShape.Left -lt ($specCell.Left - 0.1) -or $specShape.Top -lt ($specCell.Top - 0.1) -or
                    ($specShape.Left + $specShape.Width) -gt ($specCell.Left + $specCell.Width + 0.1) -or
                    ($specShape.Top + $specShape.Height) -gt ($specCell.Top + $specCell.Height + 0.1)) {
                    Add-Error "item $($itemIndex + 1) specification image is outside D-column merged cell"
                }
            }
        }
        elseif ($actualRequirement -ne [string] $expectedRow.table_requirement) {
            Add-Error "item $($itemIndex + 1) requirement mismatch: expected '$($expectedRow.table_requirement)', got '$actualRequirement'"
        }
    }

    foreach ($sheet in @($data1, $data2)) {
        if ($excel.WorksheetFunction.CountA($sheet.Range("E6:AF47")) -ne 0) {
            Add-Error "$($sheet.Name) measurement or judgment region E6:AF47 must be blank"
        }
    }

    $lowerShapes = @()
    for ($shapeIndex = 1; $shapeIndex -le $cover.Shapes.Count; $shapeIndex++) {
        $shape = $cover.Shapes.Item($shapeIndex)
        if ($shape.Top -ge 250) {
            $lowerShapes += [string] $shape.Name
        }
    }
    if ($lowerShapes.Count -ne 1 -or $lowerShapes[0] -ne "ISIR编号图纸") {
        Add-Error "cover lower area must contain exactly one shape named ISIR编号图纸; got: $($lowerShapes -join ', ')"
    }

    $result = [pscustomobject]@{
        Status = if ($errors.Count -eq 0) { "passed" } else { "failed" }
        Workbook = $resolvedPath
        Mapping = $resolvedMappingPath
        ExpectedCharacteristics = $ExpectedCharacteristics
        FormulaErrors = $formulaErrors
        Errors = $errors
    }
    $result | ConvertTo-Json -Depth 5
    if ($errors.Count -ne 0) {
        exit 1
    }
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void] [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}



