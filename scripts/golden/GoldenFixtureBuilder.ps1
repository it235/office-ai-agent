# Golden fixture builder for E11.
# Fixtures are generated from JSON specs, never committed as binaries.
# Word and Excel use the OpenXml SDK directly; PowerPoint uses the SDK package graph
# with a minimal slide master/layout/theme, then the result is read back by the SDK.

$script:GoldenOpenXmlLoaded = $false

function Initialize-GoldenOpenXml {
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    if ($script:GoldenOpenXmlLoaded) { return }
    Add-Type -Path (Join-Path $RepoRoot "packages\DocumentFormat.OpenXml.Framework.3.1.0\lib\net46\DocumentFormat.OpenXml.Framework.dll")
    Add-Type -Path (Join-Path $RepoRoot "packages\DocumentFormat.OpenXml.3.1.0\lib\net46\DocumentFormat.OpenXml.dll")
    $script:GoldenOpenXmlLoaded = $true
}

function Add-GoldenPartOf {
    param($Parent, [Parameter(Mandatory = $true)][Type]$PartType)

    $method = $Parent.GetType().GetMethods() |
        Where-Object { $_.Name -eq "AddNewPart" -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0 } |
        Select-Object -First 1
    if ($null -eq $method) { throw "AddNewPart generic overload not found on $($Parent.GetType().Name)" }
    return $method.MakeGenericMethod($PartType).Invoke($Parent, @())
}

function Get-GoldenDescendantsOf {
    param($Element, [Parameter(Mandatory = $true)][Type]$ElementType)

    $result = New-Object System.Collections.ArrayList
    foreach ($child in $Element.ChildElements) {
        if ($ElementType.IsInstanceOfType($child)) { [void]$result.Add($child) }
        foreach ($nested in (Get-GoldenDescendantsOf -Element $child -ElementType $ElementType)) {
            [void]$result.Add($nested)
        }
    }
    return ,$result
}

function Get-GoldenText {
    param($Element)
    if ($null -eq $Element) { return "" }
    return [string]$Element.InnerText
}

function Get-GoldenWordStyleId {
    param($Paragraph)
    if ($null -ne $Paragraph.ParagraphProperties -and $null -ne $Paragraph.ParagraphProperties.ParagraphStyleId) {
        return [string]$Paragraph.ParagraphProperties.ParagraphStyleId.Val.Value
    }
    return ""
}

$script:GoldenWordStylesXml = @'
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="120" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>
'@

function New-GoldenWordParagraph {
    param(
        [Parameter(Mandatory = $true)]$Body,
        [string]$Text = "",
        [string]$Style = "Normal"
    )

    $paragraph = New-Object DocumentFormat.OpenXml.Wordprocessing.Paragraph
    if (-not [string]::IsNullOrWhiteSpace($Style)) {
        $properties = New-Object DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties
        $styleId = New-Object DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId
        $styleId.Val = $Style
        [void]$properties.Append($styleId)
        [void]$paragraph.Append($properties)
    }
    $run = New-Object DocumentFormat.OpenXml.Wordprocessing.Run
    $textNode = New-Object DocumentFormat.OpenXml.Wordprocessing.Text
    $textNode.Text = [string]$Text
    [void]$run.Append($textNode)
    [void]$paragraph.Append($run)
    [void]$Body.Append($paragraph)
}

function New-GoldenWordTable {
    param($Body, $Rows)

    $table = New-Object DocumentFormat.OpenXml.Wordprocessing.Table
    $grid = New-Object DocumentFormat.OpenXml.Wordprocessing.TableGrid
    $columnCount = 0
    foreach ($row in @($Rows)) {
        if (@($row).Count -gt $columnCount) { $columnCount = @($row).Count }
    }
    for ($i = 0; $i -lt $columnCount; $i++) {
        $column = New-Object DocumentFormat.OpenXml.Wordprocessing.GridColumn
        $column.Width = "2400"
        [void]$grid.Append($column)
    }
    [void]$table.Append($grid)
    foreach ($row in @($Rows)) {
        $tableRow = New-Object DocumentFormat.OpenXml.Wordprocessing.TableRow
        foreach ($value in @($row)) {
            $cell = New-Object DocumentFormat.OpenXml.Wordprocessing.TableCell
            $paragraph = New-Object DocumentFormat.OpenXml.Wordprocessing.Paragraph
            $run = New-Object DocumentFormat.OpenXml.Wordprocessing.Run
            $textNode = New-Object DocumentFormat.OpenXml.Wordprocessing.Text
            $textNode.Text = [string]$value
            [void]$run.Append($textNode)
            [void]$paragraph.Append($run)
            [void]$cell.Append($paragraph)
            [void]$tableRow.Append($cell)
        }
        [void]$table.Append($tableRow)
    }
    [void]$Body.Append($table)
}

function New-GoldenWordFile {
    param($Spec, [Parameter(Mandatory = $true)][string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create(
        $Path, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document)
    try {
        $mainPart = $document.AddMainDocumentPart()
        $mainPart.Document = New-Object DocumentFormat.OpenXml.Wordprocessing.Document
        $body = New-Object DocumentFormat.OpenXml.Wordprocessing.Body
        [void]$mainPart.Document.Append($body)

        $stylesPart = Add-GoldenPartOf -Parent $mainPart -PartType ([DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart])
        $stylesPart.Styles = New-Object DocumentFormat.OpenXml.Wordprocessing.Styles
        $stylesPart.Styles.InnerXml = $script:GoldenWordStylesXml

        foreach ($block in @($Spec.blocks)) {
            switch ([string]$block.type) {
                "paragraph" { New-GoldenWordParagraph -Body $body -Text ([string]$block.text) -Style ([string]$block.style) }
                "table" { New-GoldenWordTable -Body $body -Rows $block.rows }
                default { throw "Fixture $($Spec.id): unknown word block type '$($block.type)'" }
            }
        }
    }
    finally {
        $document.Dispose()
    }
}

function Get-GoldenColumnName {
    param([int]$Index)
    $name = ""
    $value = $Index
    while ($value -gt 0) {
        $remainder = ($value - 1) % 26
        $name = ([char](65 + $remainder)).ToString() + $name
        $value = [int](($value - $remainder) / 26)
    }
    return $name
}

function New-GoldenExcelCell {
    param([string]$Reference, $Value)

    $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
    $cell.CellReference = $Reference
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) {
        $text = [string]$Value
        if ($text.StartsWith("=")) {
            $formula = New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormula
            $formula.Text = $text.Substring(1)
            [void]$cell.Append($formula)
            return ,$cell
        }
        $cell.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
        $inline = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
        $textNode = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
        $textNode.Text = $text
        [void]$inline.Append($textNode)
        [void]$cell.Append($inline)
        return ,$cell
    }
    $number = [Convert]::ToDouble($Value, [System.Globalization.CultureInfo]::InvariantCulture)
    $cellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue
    $cellValue.Text = $number.ToString("0.################", [System.Globalization.CultureInfo]::InvariantCulture)
    [void]$cell.Append($cellValue)
    return ,$cell
}

function New-GoldenExcelFile {
    param($Spec, [Parameter(Mandatory = $true)][string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create(
        $Path, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    try {
        $workbookPart = $document.AddWorkbookPart()
        $workbookPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
        $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
        [void]$workbookPart.Workbook.Append($sheets)

        $sheetId = 1
        foreach ($sheetSpec in @($Spec.sheets)) {
            $worksheetPart = Add-GoldenPartOf -Parent $workbookPart -PartType ([DocumentFormat.OpenXml.Packaging.WorksheetPart])
            $worksheetPart.Worksheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet
            $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
            [void]$worksheetPart.Worksheet.Append($sheetData)

            $rowIndex = 1
            foreach ($rowValues in @($sheetSpec.rows)) {
                $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
                $row.RowIndex = [uint32]$rowIndex
                $columnIndex = 1
                foreach ($value in @($rowValues)) {
                    $reference = (Get-GoldenColumnName -Index $columnIndex) + $rowIndex.ToString()
                    try {
                        $cell = New-GoldenExcelCell -Reference $reference -Value $value
                    }
                    catch {
                        throw "Fixture $($Spec.id) sheet '$($sheetSpec.name)' cell ${reference}: $($_.Exception.Message)"
                    }
                    if ($null -ne $cell) { [void]$row.Append($cell) }
                    $columnIndex++
                }
                [void]$sheetData.Append($row)
                $rowIndex++
            }

            $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
            $sheet.Id = $workbookPart.GetIdOfPart($worksheetPart)
            $sheet.SheetId = [uint32]$sheetId
            $sheet.Name = [string]$sheetSpec.name
            [void]$sheets.Append($sheet)
            $sheetId++
        }
    }
    finally {
        $document.Dispose()
    }
}

function New-GoldenPptShapeXml {
    param(
        [int]$ShapeId,
        [string]$Name,
        [string[]]$Lines,
        [int]$X,
        [int]$Y,
        [int]$Cx,
        [int]$Cy,
        [int]$FontSize,
        [bool]$Bold,
        [bool]$Bullet
    )

    $paragraphs = New-Object System.Text.StringBuilder
    foreach ($line in @($Lines)) {
        $escaped = [System.Security.SecurityElement]::Escape([string]$line)
        $runProperties = '<a:rPr lang="zh-CN" sz="' + $FontSize + '"' + $(if ($Bold) { ' b="1"' } else { '' }) + "/>"
        $paragraphProperties = if ($Bullet) {
            '<a:pPr marL="342900" indent="-342900"><a:buChar char="&#8226;"/></a:pPr>'
        } else {
            '<a:pPr algn="l"/>'
        }
        [void]$paragraphs.Append('<a:p>' + $paragraphProperties + '<a:r>' + $runProperties + '<a:t>' + $escaped + '</a:t></a:r></a:p>')
    }

    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$ShapeId" name="$Name"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$X" y="$Y"/><a:ext cx="$Cx" cy="$Cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>$($paragraphs.ToString())</p:txBody>
</p:sp>
"@
}

function New-GoldenPptFile {
    param($Spec, [Parameter(Mandatory = $true)][string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Create(
        $Path, [DocumentFormat.OpenXml.PresentationDocumentType]::Presentation)
    try {
        $presentationPart = $document.AddPresentationPart()
        $presentationPart.Presentation = New-Object DocumentFormat.OpenXml.Presentation.Presentation

        $masterPart = Add-GoldenPartOf -Parent $presentationPart -PartType ([DocumentFormat.OpenXml.Packaging.SlideMasterPart])
        $masterPart.SlideMaster = New-Object DocumentFormat.OpenXml.Presentation.SlideMaster
        $masterPart.SlideMaster.CommonSlideData = New-Object DocumentFormat.OpenXml.Presentation.CommonSlideData
        $masterPart.SlideMaster.CommonSlideData.ShapeTree = New-Object DocumentFormat.OpenXml.Presentation.ShapeTree
        $masterPart.SlideMaster.CommonSlideData.ShapeTree.InnerXml = '<p:nvGrpSpPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        $masterPart.SlideMaster.ColorMap = New-Object DocumentFormat.OpenXml.Presentation.ColorMap
        foreach ($pair in @(@("bg1","lt1"),@("tx1","dk1"),@("bg2","lt2"),@("tx2","dk2"),@("accent1","accent1"),@("accent2","accent2"),@("accent3","accent3"),@("accent4","accent4"),@("accent5","accent5"),@("accent6","accent6"),@("hlink","hlink"),@("folHlink","folHlink"))) {
            $masterPart.SlideMaster.ColorMap.SetAttribute([DocumentFormat.OpenXml.OpenXmlAttribute]::new($pair[0], "", $pair[1]))
        }

        $themePart = Add-GoldenPartOf -Parent $masterPart -PartType ([DocumentFormat.OpenXml.Packaging.ThemePart])
        $themePart.Theme = New-Object DocumentFormat.OpenXml.Drawing.Theme
        $themePart.Theme.InnerXml = $script:GoldenThemeXml

        $layoutPart = Add-GoldenPartOf -Parent $masterPart -PartType ([DocumentFormat.OpenXml.Packaging.SlideLayoutPart])
        $layoutPart.SlideLayout = New-Object DocumentFormat.OpenXml.Presentation.SlideLayout
        $layoutPart.SlideLayout.CommonSlideData = New-Object DocumentFormat.OpenXml.Presentation.CommonSlideData
        $layoutPart.SlideLayout.CommonSlideData.ShapeTree = New-Object DocumentFormat.OpenXml.Presentation.ShapeTree
        $layoutPart.SlideLayout.CommonSlideData.ShapeTree.InnerXml = '<p:nvGrpSpPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        $layoutPart.SlideLayout.CommonSlideData.Name = "Blank"
        $layoutPart.SlideLayout.Type = [DocumentFormat.OpenXml.Presentation.SlideLayoutValues]::Blank
        $layoutPart.SlideLayout.Preserve = $true
        $layoutPart.SlideLayout.ColorMapOverride = New-Object DocumentFormat.OpenXml.Presentation.ColorMapOverride
        $layoutPart.SlideLayout.ColorMapOverride.MasterColorMapping = New-Object DocumentFormat.OpenXml.Drawing.MasterColorMapping

        $masterPart.SlideMaster.SlideLayoutIdList = New-Object DocumentFormat.OpenXml.Presentation.SlideLayoutIdList
        $layoutId = New-Object DocumentFormat.OpenXml.Presentation.SlideLayoutId
        $layoutId.Id = [uint32]2147483649
        $layoutId.RelationshipId = $masterPart.GetIdOfPart($layoutPart)
        [void]$masterPart.SlideMaster.SlideLayoutIdList.Append($layoutId)

        $presentationPart.Presentation.SlideMasterIdList = New-Object DocumentFormat.OpenXml.Presentation.SlideMasterIdList
        $masterId = New-Object DocumentFormat.OpenXml.Presentation.SlideMasterId
        $masterId.Id = [uint32]2147483648
        $masterId.RelationshipId = $presentationPart.GetIdOfPart($masterPart)
        [void]$presentationPart.Presentation.SlideMasterIdList.Append($masterId)

        $presentationPart.Presentation.SlideIdList = New-Object DocumentFormat.OpenXml.Presentation.SlideIdList
        $presentationPart.Presentation.SlideSize = New-Object DocumentFormat.OpenXml.Presentation.SlideSize
        $presentationPart.Presentation.SlideSize.Cx = [int32]12192000
        $presentationPart.Presentation.SlideSize.Cy = [int32]6858000
        $presentationPart.Presentation.SlideSize.Type = [DocumentFormat.OpenXml.Presentation.SlideSizeValues]::Screen16x9
        $presentationPart.Presentation.NotesSize = New-Object DocumentFormat.OpenXml.Presentation.NotesSize
        $presentationPart.Presentation.NotesSize.Cx = [int32]6858000
        $presentationPart.Presentation.NotesSize.Cy = [int32]9144000

        $slideNumber = 1
        foreach ($slideSpec in @($Spec.slides)) {
            $slidePart = Add-GoldenPartOf -Parent $presentationPart -PartType ([DocumentFormat.OpenXml.Packaging.SlidePart])
            $slidePart.Slide = New-Object DocumentFormat.OpenXml.Presentation.Slide
            [void]$slidePart.AddPart($layoutPart)
            $slidePart.Slide.CommonSlideData = New-Object DocumentFormat.OpenXml.Presentation.CommonSlideData

            $titleXml = New-GoldenPptShapeXml -ShapeId 2 -Name "Title" -Lines @([string]$slideSpec.title) `
                -X 838200 -Y 365125 -Cx 10515600 -Cy 1325563 -FontSize 3200 -Bold $true -Bullet $false
            $bodyLines = @($slideSpec.bullets | ForEach-Object { [string]$_ })
            $bodyXml = ""
            if ($bodyLines.Count -gt 0) {
                $bodyXml = New-GoldenPptShapeXml -ShapeId 3 -Name "Body" -Lines $bodyLines `
                    -X 838200 -Y 1825625 -Cx 10515600 -Cy 4351338 -FontSize 2000 -Bold $false -Bullet $true
            }
            $slidePart.Slide.CommonSlideData.InnerXml = '<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + $titleXml + $bodyXml + '</p:spTree>'
            $slidePart.Slide.ColorMapOverride = New-Object DocumentFormat.OpenXml.Presentation.ColorMapOverride
            $slidePart.Slide.ColorMapOverride.MasterColorMapping = New-Object DocumentFormat.OpenXml.Drawing.MasterColorMapping

            $slideId = New-Object DocumentFormat.OpenXml.Presentation.SlideId
            $slideId.Id = [uint32](255 + $slideNumber)
            $slideId.RelationshipId = $presentationPart.GetIdOfPart($slidePart)
            [void]$presentationPart.Presentation.SlideIdList.Append($slideId)
            $slideNumber++
        }
    }
    finally {
        $document.Dispose()
    }
}

$script:GoldenThemeXml = @'
<a:themeElements xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
<a:extraClrSchemeLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
'@

function New-GoldenFixtureFile {
    param($Spec, [Parameter(Mandatory = $true)][string]$Path)

    switch ([string]$Spec.kind) {
        "word" { New-GoldenWordFile -Spec $Spec -Path $Path }
        "excel" { New-GoldenExcelFile -Spec $Spec -Path $Path }
        "ppt" { New-GoldenPptFile -Spec $Spec -Path $Path }
        default { throw "Fixture $($Spec.id): unknown kind '$($Spec.kind)'" }
    }
}

function Get-GoldenExtension {
    param([string]$Kind)
    switch ($Kind) {
        "word" { return ".docx" }
        "excel" { return ".xlsx" }
        "ppt" { return ".pptx" }
        default { throw "Unknown fixture kind '$Kind'" }
    }
}

function Get-GoldenWordModel {
    param([string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($Path, $false)
    try {
        $paragraphs = New-Object System.Collections.ArrayList
        $tables = New-Object System.Collections.ArrayList
        foreach ($block in $document.MainDocumentPart.Document.Body.ChildElements) {
            if ($block -is [DocumentFormat.OpenXml.Wordprocessing.Paragraph]) {
                [void]$paragraphs.Add([pscustomobject]@{
                    Style = (Get-GoldenWordStyleId -Paragraph $block)
                    Text = [string]$block.InnerText
                })
            }
            elseif ($block -is [DocumentFormat.OpenXml.Wordprocessing.Table]) {
                $rows = New-Object System.Collections.ArrayList
                foreach ($row in (Get-GoldenDescendantsOf -Element $block -ElementType ([DocumentFormat.OpenXml.Wordprocessing.TableRow]))) {
                    $cells = New-Object System.Collections.ArrayList
                    foreach ($cell in (Get-GoldenDescendantsOf -Element $row -ElementType ([DocumentFormat.OpenXml.Wordprocessing.TableCell]))) {
                        [void]$cells.Add([string]$cell.InnerText)
                    }
                    [void]$rows.Add(@($cells))
                }
                [void]$tables.Add(@($rows))
            }
        }
        return [pscustomobject]@{ Kind = "word"; Paragraphs = @($paragraphs); Tables = @($tables) }
    }
    finally {
        $document.Dispose()
    }
}

function Get-GoldenExcelModel {
    param([string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Path, $false)
    try {
        $sheets = New-Object System.Collections.ArrayList
        $sheetElements = Get-GoldenDescendantsOf -Element $document.WorkbookPart.Workbook -ElementType ([DocumentFormat.OpenXml.Spreadsheet.Sheet])
        foreach ($sheetElement in $sheetElements) {
            $relationshipId = [string]$sheetElement.Id.Value
            $worksheetPart = $null
            foreach ($candidate in $document.WorkbookPart.WorksheetParts) {
                if ($document.WorkbookPart.GetIdOfPart($candidate) -eq $relationshipId) { $worksheetPart = $candidate; break }
            }
            if ($null -eq $worksheetPart) { throw "Fixture ${Path}: sheet '$($sheetElement.Name)' has no worksheet part" }
            $cells = New-Object System.Collections.ArrayList
            foreach ($cell in (Get-GoldenDescendantsOf -Element $worksheetPart.Worksheet -ElementType ([DocumentFormat.OpenXml.Spreadsheet.Cell]))) {
                $kind = "text"
                $value = ""
                $formula = ""
                if ($null -ne $cell.DataType -and $cell.DataType.Value -eq [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString) {
                    $value = [string]$cell.InlineString.Text.Text
                }
                elseif ($null -ne $cell.CellValue) {
                    $kind = "number"
                    $value = [string]$cell.CellValue.Text
                }
                if ($null -ne $cell.CellFormula) {
                    $kind = "formula"
                    $formula = [string]$cell.CellFormula.Text
                }
                [void]$cells.Add([pscustomobject]@{
                    Ref = [string]$cell.CellReference.Value
                    Kind = $kind
                    Value = $value
                    Formula = $formula
                })
            }
            [void]$sheets.Add([pscustomobject]@{ Name = [string]$sheetElement.Name.Value; Cells = @($cells) })
        }
        return [pscustomobject]@{ Kind = "excel"; Sheets = @($sheets) }
    }
    finally {
        $document.Dispose()
    }
}

function Get-GoldenPptModel {
    param([string]$Path)

    $document = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Open($Path, $false)
    try {
        $slides = New-Object System.Collections.ArrayList
        $slideIndex = 0
        foreach ($slideId in $document.PresentationPart.Presentation.SlideIdList.ChildElements) {
            $slideIndex++
            $slidePart = $document.PresentationPart.GetPartById($slideId.RelationshipId.Value)
            $shapes = New-Object System.Collections.ArrayList
            foreach ($shape in $slidePart.Slide.CommonSlideData.ShapeTree.ChildElements) {
                if ($shape.LocalName -ne "sp") { continue }
                $drawingProperties = Get-GoldenDescendantsOf -Element $shape -ElementType ([DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties])
                $shapeName = if ($drawingProperties.Count -gt 0) { [string]$drawingProperties[0].Name.Value } else { "" }
                $paragraphs = New-Object System.Collections.ArrayList
                foreach ($paragraph in (Get-GoldenDescendantsOf -Element $shape -ElementType ([DocumentFormat.OpenXml.Drawing.Paragraph]))) {
                    $texts = Get-GoldenDescendantsOf -Element $paragraph -ElementType ([DocumentFormat.OpenXml.Drawing.Text])
                    $line = ""
                    foreach ($textNode in $texts) { $line += [string]$textNode.Text }
                    [void]$paragraphs.Add($line)
                }
                [void]$shapes.Add([pscustomobject]@{ Name = $shapeName; Paragraphs = @($paragraphs) })
            }
            [void]$slides.Add([pscustomobject]@{ Index = $slideIndex; Shapes = @($shapes) })
        }
        return [pscustomobject]@{ Kind = "ppt"; Slides = @($slides) }
    }
    finally {
        $document.Dispose()
    }
}

function Get-GoldenFixtureModel {
    param([Parameter(Mandatory = $true)][string]$Path)

    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        ".docx" { return Get-GoldenWordModel -Path $Path }
        ".xlsx" { return Get-GoldenExcelModel -Path $Path }
        ".pptx" { return Get-GoldenPptModel -Path $Path }
        default { throw "Not a golden fixture file: $Path" }
    }
}

function Get-GoldenFixtureDigest {
    param([Parameter(Mandatory = $true)][string]$Path)

    $model = Get-GoldenFixtureModel -Path $Path
    $lines = New-Object System.Collections.ArrayList
    if ($model.Kind -eq "word") {
        [void]$lines.Add("K|word")
        [void]$lines.Add("PC|$(@($model.Paragraphs).Count)")
        [void]$lines.Add("TC|$(@($model.Tables).Count)")
        foreach ($paragraph in $model.Paragraphs) { [void]$lines.Add("P|$($paragraph.Style)|$($paragraph.Text)") }
        foreach ($table in $model.Tables) {
            foreach ($row in $table) { [void]$lines.Add("T|" + (@($row) -join "~")) }
        }
    }
    elseif ($model.Kind -eq "excel") {
        [void]$lines.Add("K|excel")
        [void]$lines.Add("SC|$(@($model.Sheets).Count)")
        foreach ($sheet in $model.Sheets) {
            [void]$lines.Add("S|$($sheet.Name)")
            foreach ($cell in $sheet.Cells) { [void]$lines.Add("C|$($cell.Ref)|$($cell.Kind)|$($cell.Value)|$($cell.Formula)") }
        }
    }
    else {
        [void]$lines.Add("K|ppt")
        [void]$lines.Add("SC|$(@($model.Slides).Count)")
        foreach ($slide in $model.Slides) {
            [void]$lines.Add("SL|$($slide.Index)")
            foreach ($shape in $slide.Shapes) { [void]$lines.Add("SP|$($shape.Name)|" + (@($shape.Paragraphs) -join "~")) }
        }
    }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($lines -join "`n"))
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function New-GoldenFixtures {
    param(
        [Parameter(Mandatory = $true)][string]$SpecDir,
        [Parameter(Mandatory = $true)][string]$OutputDir
    )

    $specFiles = @(Get-ChildItem -LiteralPath $SpecDir -Filter *.json -File -Recurse | Sort-Object FullName)
    if ($specFiles.Count -eq 0) { throw "No golden fixture specs under $SpecDir" }

    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    $fixtures = New-Object System.Collections.ArrayList
    $seenIds = @{}
    foreach ($specFile in $specFiles) {
        $spec = Get-Content -LiteralPath $specFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        $id = [string]$spec.id
        if ([string]::IsNullOrWhiteSpace($id)) { throw "Fixture spec $($specFile.Name) has no id" }
        if ($seenIds.ContainsKey($id)) { throw "Duplicate fixture id '$id'" }
        $seenIds[$id] = $true
        if ([string]::IsNullOrWhiteSpace([string]$spec.dataPolicy) -or [string]$spec.dataPolicy -ne "fictional") {
            throw "Fixture '$id' must declare dataPolicy=fictional; golden fixtures may not carry customer data"
        }

        $extension = Get-GoldenExtension -Kind ([string]$spec.kind)
        $fileName = $id.Replace("/", "-") + $extension
        $path = Join-Path $OutputDir $fileName
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
        New-GoldenFixtureFile -Spec $spec -Path $path
        if (-not (Test-Path -LiteralPath $path)) { throw "Fixture '$id' was not generated" }

        $model = Get-GoldenFixtureModel -Path $path
        $digest = Get-GoldenFixtureDigest -Path $path
        [void]$fixtures.Add([pscustomobject]@{
            id = $id
            kind = [string]$spec.kind
            appType = [string]$spec.appType
            file = $fileName
            contentSha256 = $digest
            bytes = (Get-Item -LiteralPath $path).Length
        })
        Write-Verbose "Generated $id -> $fileName ($digest)"
    }

    $manifest = [pscustomobject]@{
        schemaVersion = 1
        generator = "scripts/New-GoldenFixtures.ps1"
        fixtures = @($fixtures | Sort-Object id)
    }
    $manifestPath = Join-Path $OutputDir "manifest.json"
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    return $manifest
}
