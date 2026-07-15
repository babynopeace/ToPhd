param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$DataPath,

    [string]$UnitDataPath,

    [string]$Title = "ToPhd 导师候选追踪表",

    [string]$QueryDate = (Get-Date -Format "yyyy-MM-dd"),

    [switch]$Template
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Get-ExcelColor {
    param([int]$R, [int]$G, [int]$B)
    return $R + (256 * $G) + (65536 * $B)
}

$headers = @(
    "编号", "候选类型", "机构类型", "研究路线", "学校或科研院所",
    "学院/研究所/实验室", "所在城市", "专业/学科", "导师", "申请类型",
    "研究匹配依据", "硬性条件状态", "招生状态", "官方截止时间", "主要风险",
    "官方来源", "查询日期", "备注"
)

$unitHeaders = @(
    "编号", "机构类型", "报考单位", "研究生招生官网", "最新博士招生简章",
    "博士招生通知入口", "专业目录或学院细则", "当前适用年度", "2027状态", "查询日期", "备注"
)

$rawRows = @()
if (-not $Template) {
    if (-not $DataPath) {
        throw "生成非空追踪表时必须提供 -DataPath。"
    }
    if (-not (Test-Path -LiteralPath $DataPath)) {
        throw "找不到候选数据文件：$DataPath"
    }
    $rawRows = @(Get-Content -Encoding utf8 -LiteralPath $DataPath | ConvertFrom-Csv -Delimiter '|')
}

$rows = @(foreach ($row in $rawRows) {
    $applicationType = if ($row.机构类型 -eq "科研院所") {
        "国科大普通招考学博/联合培养（待2027目录确认）"
    }
    else {
        "普通招考全日制学博（待2027目录确认）"
    }

    $hardRequirement = if ($row.候选类型 -eq "专业扩展") {
        "跨学科接收、外语要求及2027专业目录待核验"
    }
    else {
        "外语要求、专业背景及2027招生规则待核验"
    }

    [PSCustomObject]@{
        "编号" = $row.编号
        "候选类型" = $row.候选类型
        "机构类型" = $row.机构类型
        "研究路线" = $row.研究路线
        "学校或科研院所" = $row.学校或科研院所
        "学院/研究所/实验室" = $row.学院研究所实验室
        "所在城市" = $row.所在城市
        "专业/学科" = $row.专业学科
        "导师" = $row.导师
        "申请类型" = $applicationType
        "研究匹配依据" = $row.研究匹配依据
        "硬性条件状态" = $hardRequirement
        "招生状态" = $row.招生状态
        "官方截止时间" = "待2027简章"
        "主要风险" = $row.主要风险
        "官方来源" = $row.官方来源
        "查询日期" = $QueryDate
        "备注" = $row.备注
    }
})

$unitRows = @()
if ($UnitDataPath) {
    if (-not (Test-Path -LiteralPath $UnitDataPath)) {
        throw "找不到招生单位数据文件：$UnitDataPath"
    }
    $unitRows = @(Get-Content -Encoding utf8 -LiteralPath $UnitDataPath | ConvertFrom-Csv -Delimiter '|')
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($outputFullPath)
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
}
if (Test-Path -LiteralPath $outputFullPath) {
    Remove-Item -LiteralPath $outputFullPath -Force
}

$excel = $null
$book = $null
$overview = $null
$pool = $null
$unitSheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $book = $excel.Workbooks.Add()
    while ($book.Worksheets.Count -lt 3) {
        $book.Worksheets.Add() | Out-Null
    }
    while ($book.Worksheets.Count -gt 3) {
        $book.Worksheets.Item($book.Worksheets.Count).Delete()
    }

    $overview = $book.Worksheets.Item(1)
    $pool = $book.Worksheets.Item(2)
    $unitSheet = $book.Worksheets.Item(3)
    $overview.Name = "总览"
    $pool.Name = "候选池"
    $unitSheet.Name = "招生单位"

    $lastColumn = $headers.Count
    $firstDataRow = 4
    $lastDataRow = if ($rows.Count -gt 0) { $firstDataRow + $rows.Count - 1 } else { $firstDataRow }

    $pool.Cells.Font.Name = "微软雅黑"
    $pool.Cells.Font.Size = 10

    $titleRange = $pool.Range($pool.Cells.Item(1, 1), $pool.Cells.Item(1, $lastColumn))
    $titleRange.Merge()
    $titleRange.Value2 = $Title
    $titleRange.Font.Name = "微软雅黑"
    $titleRange.Font.Size = 18
    $titleRange.Font.Bold = $true
    $titleRange.Font.Color = Get-ExcelColor 255 255 255
    $titleRange.Interior.Color = Get-ExcelColor 31 78 121
    $titleRange.HorizontalAlignment = -4108
    $titleRange.VerticalAlignment = -4108
    $pool.Rows.Item(1).RowHeight = 34

    $directCount = @($rows | Where-Object { $_.候选类型 -eq "专业对口" }).Count
    $extensionCount = @($rows | Where-Object { $_.候选类型 -eq "专业扩展" }).Count
    $instituteCount = @($rows | Where-Object { $_.机构类型 -eq "科研院所" }).Count
    $summaryText = if ($Template) {
        '模板说明：候选类型分为“专业对口”和“专业扩展”；招生信息必须按申请年度继续核验。'
    }
    else {
        "共 $($rows.Count) 条候选｜专业对口 $directCount 条｜专业扩展 $extensionCount 条｜科研院所 $instituteCount 条｜候选池不等于最终可申请名单"
    }
    $summaryRange = $pool.Range($pool.Cells.Item(2, 1), $pool.Cells.Item(2, $lastColumn))
    $summaryRange.Merge()
    $summaryRange.Value2 = $summaryText
    $summaryRange.Font.Name = "微软雅黑"
    $summaryRange.Font.Size = 10
    $summaryRange.Font.Color = Get-ExcelColor 64 64 64
    $summaryRange.Interior.Color = Get-ExcelColor 221 235 247
    $summaryRange.WrapText = $true
    $summaryRange.HorizontalAlignment = -4131
    $summaryRange.VerticalAlignment = -4108
    $pool.Rows.Item(2).RowHeight = 32

    for ($column = 1; $column -le $lastColumn; $column++) {
        $pool.Cells.Item(3, $column).Value2 = $headers[$column - 1]
    }
    $headerRange = $pool.Range($pool.Cells.Item(3, 1), $pool.Cells.Item(3, $lastColumn))
    $headerRange.Font.Name = "微软雅黑"
    $headerRange.Font.Size = 10
    $headerRange.Font.Bold = $true
    $headerRange.Font.Color = Get-ExcelColor 255 255 255
    $headerRange.Interior.Color = Get-ExcelColor 68 114 196
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.WrapText = $true
    $pool.Rows.Item(3).RowHeight = 34

    if ($rows.Count -gt 0) {
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
            for ($columnIndex = 0; $columnIndex -lt $lastColumn; $columnIndex++) {
                $propertyName = $headers[$columnIndex]
                $pool.Cells.Item($firstDataRow + $rowIndex, $columnIndex + 1).Value2 = [string]$rows[$rowIndex].$propertyName
            }
        }
        $dataRange = $pool.Range($pool.Cells.Item($firstDataRow, 1), $pool.Cells.Item($lastDataRow, $lastColumn))
        $dataRange.WrapText = $true
        $dataRange.VerticalAlignment = -4160
        $dataRange.HorizontalAlignment = -4131
        $dataRange.Font.Name = "微软雅黑"
        $dataRange.Font.Size = 9

        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
            $sheetRow = $firstDataRow + $rowIndex
            $rowRange = $pool.Range($pool.Cells.Item($sheetRow, 1), $pool.Cells.Item($sheetRow, $lastColumn))
            if ($rows[$rowIndex].候选类型 -eq "专业对口") {
                $rowRange.Interior.Color = Get-ExcelColor 235 244 252
            }
            else {
                $rowRange.Interior.Color = Get-ExcelColor 255 248 225
            }
            if ($rows[$rowIndex].招生状态 -like "*已满*") {
                $rowRange.Interior.Color = Get-ExcelColor 255 220 220
                $rowRange.Font.Color = Get-ExcelColor 156 0 6
            }
            $pool.Rows.Item($sheetRow).RowHeight = 58

            $sourceCell = $pool.Cells.Item($sheetRow, 16)
            $sourceUrl = [string]$rows[$rowIndex].官方来源
            if ($sourceUrl) {
                $pool.Hyperlinks.Add($sourceCell, $sourceUrl, "", "", "打开官方页面") | Out-Null
                $sourceCell.Font.Color = Get-ExcelColor 5 99 193
                $sourceCell.Font.Underline = $true
            }
        }
    }

    $tableLastRow = if ($rows.Count -gt 0) { $lastDataRow } else { 3 }
    $tableRange = $pool.Range($pool.Cells.Item(3, 1), $pool.Cells.Item($tableLastRow, $lastColumn))
    $tableRange.Borders.LineStyle = 1
    $tableRange.Borders.Weight = 2
    $tableRange.Borders.Color = Get-ExcelColor 191 191 191
    $tableRange.AutoFilter() | Out-Null

    $widths = @(6, 12, 11, 21, 23, 30, 10, 24, 11, 27, 48, 32, 31, 14, 40, 18, 13, 36)
    for ($column = 1; $column -le $lastColumn; $column++) {
        $pool.Columns.Item($column).ColumnWidth = $widths[$column - 1]
    }
    foreach ($column in @(1, 2, 3, 7, 9, 14, 16, 17)) {
        $pool.Columns.Item($column).HorizontalAlignment = -4108
    }

    $pool.Activate()
    $excel.ActiveWindow.SplitRow = 3
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.Zoom = 80
    $excel.ActiveWindow.DisplayGridlines = $false

    $overview.Cells.Font.Name = "微软雅黑"
    $overview.Cells.Font.Size = 11
    $overviewTitle = $overview.Range("A1:F1")
    $overviewTitle.Merge()
    $overviewTitle.Value2 = "ToPhd 候选池总览"
    $overviewTitle.Font.Name = "微软雅黑"
    $overviewTitle.Font.Size = 20
    $overviewTitle.Font.Bold = $true
    $overviewTitle.Font.Color = Get-ExcelColor 255 255 255
    $overviewTitle.Interior.Color = Get-ExcelColor 31 78 121
    $overviewTitle.HorizontalAlignment = -4108
    $overviewTitle.VerticalAlignment = -4108
    $overview.Rows.Item(1).RowHeight = 38

    $metrics = @(
        @("候选总数", $rows.Count),
        @("专业对口", $directCount),
        @("专业扩展", $extensionCount),
        @("985高校", @($rows | Where-Object { $_.机构类型 -eq "985高校" }).Count),
        @("科研院所", $instituteCount),
        @("招生单位", $unitRows.Count)
    )
    $overview.Cells.Item(3, 1).Value2 = "指标"
    $overview.Cells.Item(3, 2).Value2 = "数量"
    $overview.Range("A3:B3").Font.Bold = $true
    $overview.Range("A3:B3").Interior.Color = Get-ExcelColor 68 114 196
    $overview.Range("A3:B3").Font.Color = Get-ExcelColor 255 255 255
    for ($index = 0; $index -lt $metrics.Count; $index++) {
        $overview.Cells.Item(4 + $index, 1).Value2 = $metrics[$index][0]
        $overview.Cells.Item(4 + $index, 2).Value2 = [string]$metrics[$index][1]
    }
    $overview.Range("A3:B9").Borders.LineStyle = 1
    $overview.Range("A3:B9").Borders.Weight = 2

    $overview.Range("D3:F3").Merge()
    $overview.Range("D3:F3").Value2 = "颜色与筛选说明"
    $overview.Range("D3:F3").Font.Bold = $true
    $overview.Range("D3:F3").Interior.Color = Get-ExcelColor 68 114 196
    $overview.Range("D3:F3").Font.Color = Get-ExcelColor 255 255 255
    $overview.Range("D4:F4").Merge()
    $overview.Range("D4:F4").Value2 = "浅蓝：专业对口；浅黄：专业扩展；浅红：已知当年不可申请。"
    $overview.Range("D4:F4").WrapText = $true
    $overview.Range("D5:F5").Merge()
    $overview.Range("D5:F5").Value2 = '打开“候选池”后，可按候选类型、机构类型、城市、研究路线和招生状态直接筛选。'
    $overview.Range("D5:F5").WrapText = $true
    $overview.Range("D6:F8").Merge()
    $overview.Range("D6:F8").Value2 = "证据边界：导师主页或团队页面只用于确认研究方向，不代表2027年一定招生。普通招考资格、招生专业、个人名额和截止时间必须继续以2027年官方简章、专业目录或导师明确回复为准。"
    $overview.Range("D6:F8").WrapText = $true
    $overview.Range("D3:F8").Borders.LineStyle = 1
    $overview.Range("D3:F8").Borders.Weight = 2

    $overview.Cells.Item(11, 1).Value2 = "机构分布"
    $overview.Cells.Item(11, 1).Font.Bold = $true
    $overview.Cells.Item(11, 1).Font.Size = 13
    $overview.Cells.Item(12, 1).Value2 = "学校或科研院所"
    $overview.Cells.Item(12, 2).Value2 = "候选数"
    $overview.Range("A12:B12").Font.Bold = $true
    $overview.Range("A12:B12").Interior.Color = Get-ExcelColor 68 114 196
    $overview.Range("A12:B12").Font.Color = Get-ExcelColor 255 255 255
    $groups = @($rows | Group-Object 学校或科研院所 | Sort-Object -Property @{Expression = 'Count'; Descending = $true}, @{Expression = 'Name'; Descending = $false})
    for ($index = 0; $index -lt $groups.Count; $index++) {
        $overview.Cells.Item(13 + $index, 1).Value2 = $groups[$index].Name
        $overview.Cells.Item(13 + $index, 2).Value2 = [string]$groups[$index].Count
    }
    if ($groups.Count -gt 0) {
        $overview.Range($overview.Cells.Item(12, 1), $overview.Cells.Item(12 + $groups.Count, 2)).Borders.LineStyle = 1
        $overview.Range($overview.Cells.Item(12, 1), $overview.Cells.Item(12 + $groups.Count, 2)).Borders.Weight = 2
    }

    $overview.Columns.Item(1).ColumnWidth = 34
    $overview.Columns.Item(2).ColumnWidth = 12
    $overview.Columns.Item(3).ColumnWidth = 4
    $overview.Columns.Item(4).ColumnWidth = 25
    $overview.Columns.Item(5).ColumnWidth = 25
    $overview.Columns.Item(6).ColumnWidth = 25
    $overview.Rows.Item(4).RowHeight = 28
    $overview.Rows.Item(5).RowHeight = 42
    $overview.Rows.Item(6).RowHeight = 34
    $overview.Rows.Item(7).RowHeight = 34
    $overview.Rows.Item(8).RowHeight = 34
    $overview.Range("A1:F40").VerticalAlignment = -4108

    $unitColumnCount = $unitHeaders.Count
    $unitFirstDataRow = 4
    $unitLastDataRow = if ($unitRows.Count -gt 0) { $unitFirstDataRow + $unitRows.Count - 1 } else { $unitFirstDataRow }

    $unitSheet.Cells.Font.Name = "微软雅黑"
    $unitSheet.Cells.Font.Size = 10

    $unitTitleRange = $unitSheet.Range($unitSheet.Cells.Item(1, 1), $unitSheet.Cells.Item(1, $unitColumnCount))
    $unitTitleRange.Merge()
    $unitTitleRange.Value2 = "报考单位招生信息"
    $unitTitleRange.Font.Name = "微软雅黑"
    $unitTitleRange.Font.Size = 18
    $unitTitleRange.Font.Bold = $true
    $unitTitleRange.Font.Color = Get-ExcelColor 255 255 255
    $unitTitleRange.Interior.Color = Get-ExcelColor 31 78 121
    $unitTitleRange.HorizontalAlignment = -4108
    $unitTitleRange.VerticalAlignment = -4108
    $unitSheet.Rows.Item(1).RowHeight = 34

    $unitSummaryRange = $unitSheet.Range($unitSheet.Cells.Item(2, 1), $unitSheet.Cells.Item(2, $unitColumnCount))
    $unitSummaryRange.Merge()
    $unitSummaryRange.Value2 = if ($Template) {
        "每个报考单位单独记录招生官网、简章、通知和专业目录；目标年度未发布时必须标记历史参考。"
    }
    else {
        "共 $($unitRows.Count) 个报考单位｜官方招生信息与导师研究方向分开核验｜第三方平台只用于发现线索"
    }
    $unitSummaryRange.Font.Name = "微软雅黑"
    $unitSummaryRange.Font.Color = Get-ExcelColor 64 64 64
    $unitSummaryRange.Interior.Color = Get-ExcelColor 221 235 247
    $unitSummaryRange.WrapText = $true
    $unitSummaryRange.VerticalAlignment = -4108
    $unitSheet.Rows.Item(2).RowHeight = 32

    for ($column = 1; $column -le $unitColumnCount; $column++) {
        $unitSheet.Cells.Item(3, $column).Value2 = $unitHeaders[$column - 1]
    }
    $unitHeaderRange = $unitSheet.Range($unitSheet.Cells.Item(3, 1), $unitSheet.Cells.Item(3, $unitColumnCount))
    $unitHeaderRange.Font.Name = "微软雅黑"
    $unitHeaderRange.Font.Bold = $true
    $unitHeaderRange.Font.Color = Get-ExcelColor 255 255 255
    $unitHeaderRange.Interior.Color = Get-ExcelColor 68 114 196
    $unitHeaderRange.HorizontalAlignment = -4108
    $unitHeaderRange.VerticalAlignment = -4108
    $unitHeaderRange.WrapText = $true
    $unitSheet.Rows.Item(3).RowHeight = 34

    if ($unitRows.Count -gt 0) {
        for ($rowIndex = 0; $rowIndex -lt $unitRows.Count; $rowIndex++) {
            $sheetRow = $unitFirstDataRow + $rowIndex
            for ($columnIndex = 0; $columnIndex -lt $unitColumnCount; $columnIndex++) {
                $propertyName = $unitHeaders[$columnIndex]
                $unitSheet.Cells.Item($sheetRow, $columnIndex + 1).Value2 = [string]$unitRows[$rowIndex].$propertyName
            }

            $unitRowRange = $unitSheet.Range($unitSheet.Cells.Item($sheetRow, 1), $unitSheet.Cells.Item($sheetRow, $unitColumnCount))
            $unitRowRange.WrapText = $true
            $unitRowRange.VerticalAlignment = -4160
            $unitStatus = [string]$unitRows[$rowIndex].'2027状态'
            $unitRowRange.Interior.Color = if ($unitStatus -like "*不适用*") {
                Get-ExcelColor 255 220 220
            }
            elseif ($unitStatus -like "已发布*") {
                Get-ExcelColor 226 239 218
            }
            else {
                Get-ExcelColor 255 248 225
            }
            $unitSheet.Rows.Item($sheetRow).RowHeight = 64

            $unitLinks = @(
                @{ Column = 4; Text = "打开招生官网" },
                @{ Column = 5; Text = "打开招生简章" },
                @{ Column = 6; Text = "打开博士通知" },
                @{ Column = 7; Text = "打开目录或细则" }
            )
            foreach ($link in $unitLinks) {
                $linkCell = $unitSheet.Cells.Item($sheetRow, $link.Column)
                $linkProperty = $unitHeaders[$link.Column - 1]
                $linkUrl = [string]$unitRows[$rowIndex].$linkProperty
                if ($linkUrl) {
                    $unitSheet.Hyperlinks.Add($linkCell, $linkUrl, "", "", $link.Text) | Out-Null
                    $linkCell.Font.Color = Get-ExcelColor 5 99 193
                    $linkCell.Font.Underline = $true
                }
            }
        }
    }

    $unitTableLastRow = if ($unitRows.Count -gt 0) { $unitLastDataRow } else { 3 }
    $unitTableRange = $unitSheet.Range($unitSheet.Cells.Item(3, 1), $unitSheet.Cells.Item($unitTableLastRow, $unitColumnCount))
    $unitTableRange.Borders.LineStyle = 1
    $unitTableRange.Borders.Weight = 2
    $unitTableRange.Borders.Color = Get-ExcelColor 191 191 191
    $unitTableRange.AutoFilter() | Out-Null

    $unitWidths = @(7, 12, 25, 18, 20, 20, 23, 14, 46, 13, 42)
    for ($column = 1; $column -le $unitColumnCount; $column++) {
        $unitSheet.Columns.Item($column).ColumnWidth = $unitWidths[$column - 1]
    }
    foreach ($column in @(1, 2, 4, 5, 6, 7, 8, 10)) {
        $unitSheet.Columns.Item($column).HorizontalAlignment = -4108
    }

    $unitSheet.Activate()
    $excel.ActiveWindow.SplitRow = 3
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.Zoom = 82
    $excel.ActiveWindow.DisplayGridlines = $false

    $overview.Activate()
    $excel.ActiveWindow.Zoom = 95
    $excel.ActiveWindow.DisplayGridlines = $false

    $book.SaveAs($outputFullPath, 51)
    $book.Close($true)
    $excel.Quit()
    Write-Output $outputFullPath
}
catch {
    $lineNumber = $_.InvocationInfo.ScriptLineNumber
    $lineText = $_.InvocationInfo.Line
    Write-Error "Excel 追踪表生成失败，脚本第 $lineNumber 行：$lineText`n$($_.Exception.Message)"
    throw
}
finally {
    if ($book) {
        try { $book.Close($false) } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($unitSheet, $overview, $pool, $book, $excel)) {
        if ($comObject) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
