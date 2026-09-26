<#
.SYNOPSIS
    Module for creating and managing Adaptive Cards for Microsoft Teams.

.DESCRIPTION
    This module provides functions to generate JSON payloads for Adaptive Cards to create Fact Sets, Accordions, Tables, Code Blocks, Icons, Charts and more.
    These payloads can be used to post messages to Microsoft Teams channels via Workflow webhooks.

.AUTHOR
    EW

.COPYRIGHT
    No

.LICENSE
    MIT

.VERSION
    0.0.7

.NOTES
    - Requires PowerShell 5.1 or later.
    - For more information, see the Adaptive Cards documentation at https://adaptivecards.io and https://adaptivecards.microsoft.com/
    - Text formatting: https://adaptivecards.microsoft.com/?topic=text-formatting    
    - Icons: https://adaptivecards.microsoft.com/?topic=icon-catalog
    
.EXAMPLE
    $cardContent = @(
        New-TextBlock -Text "Hello, Teams!"
    )
    New-AdaptiveCard -BodyContent $cardContent | ConvertTo-Json -Depth 20
    
#>

function New-AdaptiveCard {
    param (
        [Parameter(Mandatory = $true)]
        [array]$BodyContent
    )

    $adaptiveCard = [pscustomobject]@{
        type = 'AdaptiveCard'
        body = @()
        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
        version = '1.4'        
    }

    foreach ($item in $BodyContent) {
        $adaptiveCard.body += $item
    }    
    return $adaptiveCard
}

function New-TextBlock {
    param (
        [Parameter(Mandatory = $false)] 
        [bool]$isSubtle,

        [Parameter(Mandatory = $false)]
        [bool]$separator,
        
        [Parameter(Mandatory = $false)]
        [int]$maxLines,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'Small', 'Medium', 'Large', 'ExtraLarge', IgnoreCase = $true)]
        [string]$size,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'lighter', 'bolder', IgnoreCase = $false)]
        [string]$weight,

        [Parameter(Mandatory = $false)]
        [bool]$wrap,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'Dark', 'Light', 'Accent', 'Good', 'Warning', 'Attention', IgnoreCase = $true)]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'monospace', IgnoreCase = $false)]
        [string]$fontType,

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'None', 'ExtraSmall', 'Small', 'Medium', 'Large', 'ExtraLarge', 'Padding', IgnoreCase = $true)]
        [string]$spacing,

        [Parameter(Mandatory = $true)]
        [string]$text
    )
    begin {
        $textBlock = [pscustomobject]@{
            type                = 'TextBlock'
            text                = $text
        }
    }
    process {
        if ($isSubtle)              { $textBlock | Add-Member -NotePropertyName 'isSubtle' -NotePropertyValue $isSubtle }
        if ($separator)             { $textBlock | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($maxLines)              { $textBlock | Add-Member -NotePropertyName 'maxLines' -NotePropertyValue $maxLines }
        if ($size)                  { $textBlock | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
        if ($weight)                { $textBlock | Add-Member -NotePropertyName 'weight' -NotePropertyValue $weight }
        if ($wrap)                  { $textBlock | Add-Member -NotePropertyName 'wrap' -NotePropertyValue $wrap }
        if ($color)                 { $textBlock | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($fonttype)              { $textBlock | Add-Member -NotePropertyName 'fonttype' -NotePropertyValue $fonttype }
        if ($horizontalAlignment)   { $textBlock | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($spacing)               { $textBlock | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
    }
    end {
        return $textBlock
    }
}

function New-CodeBlock {
    param (
        [Parameter(Mandatory = $true)]
        [string]$codeSnippet,

        [Parameter(Mandatory = $true)][ValidateSet(
            'Bash',
            'C',
            'C++',
            'C#',
            'CSS',
            'DOS',
            'Go',
            'GraphQL',
            'HTML',
            'Java',
            'JavaScript',
            'JSON',
            'Perl',
            'PHP',
            'PlainText',
            'PowerShell',
            'Python',
            'SQL',
            'TypeScript',
            'Visual Basic',
            'Verilog',
            'VHDL',
            'XML',
            IgnoreCase = $false
        )]
        [string]$language,

        [Parameter(Mandatory = $false)]
        [int]$startLineNumber
    )
    begin {
        $codeBlock = [pscustomobject]@{
            type        = 'CodeBlock'
            codeSnippet = $codeSnippet
            language    = $language
        }
    }
    process {
        if ($startLineNumber) {
            $codeBlock | Add-Member -NotePropertyName 'startLineNumber' -NotePropertyValue $startLineNumber
        }
    }
    end {
        return $codeBlock
    }
}

function New-Table {
    param (        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject]$object,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')]
        [string]$highlightValueMatch,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')][ValidateSet('dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$highlightValueStyle,

        [Parameter(Mandatory = $false)]
        [bool]$firstRowAsHeader = $true,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$headerRowStyle,
        
        [Parameter(Mandatory = $false)]
        [bool]$showGridLines = $true,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$gridStyle = 'default',

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalCellContentAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('top', 'center', 'bottom', IgnoreCase = $false)]
        [string]$verticalCellContentAlignment
    )
    begin {        
        $table = [pscustomobject]@{
            type                = 'Table'
            gridStyle           = $gridStyle
            firstRowAsHeader    = $firstRowAsHeader
            showGridLines       = $showGridLines
            columns             = @()
            rows                = @()
        }

        # Add optional attributes if provided
        if ($horizontalCellContentAlignment) {
            $table.horizontalCellContentAlignment = $horizontalCellContentAlignment
        }
        if ($verticalCellContentAlignment) {
            $table.verticalCellContentAlignment = $verticalCellContentAlignment
        }

        $columns = @()        
        $isHighlighting = $highlightValueMatch -and $highlightValueStyle
    }
    process {
        if ($columns.Count -eq 0) {

            # Get the noteproperties from the first object
            $columns = $Object.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name

            # Add correct number of columns (one per property)
            foreach ($column in $columns) {
                $table.columns += @{
                    width = 1 # 'auto' worked fine before, but now results in unaligned tables in teams channels, soo.. 1.
                }
            }

            # Add the header row
            $headerRow = @{
                type  = 'TableRow'
                cells = @()
            }

            # Add optional attributes if provided
            if ($headerRowStyle) {
                $headerRow.style = $headerRowStyle
            }
        
            # Add the header row
            foreach ($column in $columns) {
                $headerRow.cells += @{
                    type  = 'TableCell'
                    items = @(
                        @{
                            type = 'TextBlock'
                            text = $column # Noteproperty names
                        }
                    )
                }
            }

            $table.rows += $headerRow
        }

        # Process each object and add a row to the table
        $row = @{
            type  = 'TableRow'
            cells = @()
        }

        foreach ($column in $columns) {
            $textValue = [string]$Object.$column
            $textBlock = @{
                type = 'TextBlock'
                text = $textValue
            }

            if ($isHighlighting -and $textValue -match $highlightValueMatch) {
                $textBlock.color = $highlightValueStyle
            }

            $row.cells += @{
                type  = 'TableCell'
                items = @($textBlock)
            }
        }

        $table.rows += $row
    }
    end {        
        return $table
    }
}

function New-Image {
    param (
        [Parameter(Mandatory = $true)]
        [string]$url,

        [Parameter(Mandatory = $false)]
        [string]$altText = 'image', #Yes, this is mandatory in spec, but getting errors from not providing alt-texts is not amusing, let's be honest.

        [Parameter(Mandatory = $false)]
        [string]$backgroundColor,

        [Parameter(Mandatory = $false)][ValidateScript({$_ -ceq 'auto' -or $_ -ceq 'stretch' -or $_ -match "^\d+px$"}, ErrorMessage = "Height must be either lowercase 'auto', 'stretch', or a number followed by 'px'.")]
        [string]$height = "auto",

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('auto', 'stretch', 'small', 'medium', 'large', IgnoreCase = $false)]
        [string]$size,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'person', IgnoreCase = $false)]
        [string]$style,

        [Parameter(Mandatory = $false)][ValidatePattern("^\d+(px)?$", ErrorMessage = "Width must be an integer, optionally followed by 'px' to specify this unit.")]
        [string]$width,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'none', 'small', 'medium', 'large', 'extraLarge', 'padding', IgnoreCase = $false)]
        [string]$spacing = 'default',

        [Parameter(Mandatory = $false)]
        [bool]$separator
    )
    begin {
        $imageObject = [PSCustomObject]@{
            type = 'Image'
            url  = $url
            altText  = $altText
        }
    }
    process {
        if ($backgroundColor)       { $imageObject | Add-Member -NotePropertyName 'backgroundColor' -NotePropertyValue $vackgroundColor }
        if ($height -ne 'auto')     { $imageObject | Add-Member -NotePropertyName 'height' -NotePropertyValue $height }
        if ($horizontalAlignment)   { $imageObject | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }    
        if ($size)                  { $imageObject | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
        if ($style)                 { $imageObject | Add-Member -NotePropertyName 'style' -NotePropertyValue $style }
        if ($width)                 { $imageObject | Add-Member -NotePropertyName 'width' -NotePropertyValue $width }
        if ($spacing -ne 'default') { $imageObject | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($separator)             { $imageObject | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
    }
    end {
        return $imageObject
    }
}

function New-Icon {
    param (
        [Parameter(Mandatory = $true)]
        [string]$name,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'Dark', 'Light', 'Accent', 'Good', 'Warning', 'Attention', IgnoreCase = $false)]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet('Standard', 'xxSmall', 'xSmall', 'Small', 'Medium', 'Large', 'xLarge', 'xxlarge', IgnoreCase = $false)]
        [string]$size,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'none', 'small', 'medium', 'large', 'extraLarge', 'xxLarge', IgnoreCase = $false)]
        [string]$spacing,

        [Parameter(Mandatory = $false)][ValidateSet('Left', 'Center', 'Right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)]
        [bool]$separator,

        [Parameter(Mandatory = $false)]
        [bool]$isVisible = $true
    )
    begin {
        $icon = [pscustomobject]@{
            type = 'Icon'
            name = $name
        }
    }
    process {
        if ($color)                  { $icon | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($size)                   { $icon | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
        if ($spacing)                { $icon | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($horizontalAlignment)    { $icon | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($separator)              { $icon | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if (-not $isVisible)         { $icon | Add-Member -NotePropertyName 'isVisible' -NotePropertyValue $isVisible }
    }
    end {
        return $icon
    }
}

function New-ColumnSet {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Columns,

        [Parameter(Mandatory = $false)]
        [bool]$separator,

        [Parameter(Mandatory = $false)][ValidateSet('Default', 'None', 'ExtraSmall', 'Small', 'Medium', 'Large', 'ExtraLarge', 'Padding', IgnoreCase = $false)]
        [string]$spacing,

        [Parameter(Mandatory = $false)][ValidateSet('Left', 'Center', 'Right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)]
        [bool]$bleed
    )
    begin {
        $columnSet = [pscustomobject]@{
            type    = 'ColumnSet'
            columns = @()
        }
    }
    process {
        foreach ($column in $Columns) {
            $columnSet.columns += $column
        }

        if ($separator)           { $columnSet | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($spacing)             { $columnSet | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($horizontalAlignment) { $columnSet | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($bleed)               { $columnSet | Add-Member -NotePropertyName 'bleed' -NotePropertyValue $bleed }
    }
    end {
        return $columnSet
    }
}

function New-FactSet {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Facts
    )
    begin {
        $factSet = [pscustomobject]@{
            type  = 'FactSet'
            facts = @()
        }
    }
    process {
        foreach ($fact in $Facts) {
            $factSet.facts += [pscustomobject]@{
                title = $fact.title
                value = $fact.value
            }
        }
    }
    end {
        return $factSet
    }
}

function New-ProgressBar {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 100)]
        [int]$value,

        [Parameter(Mandatory = $false)]
        [int]$max = 100,

        [Parameter(Mandatory = $false)]
        [string]$label,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'small', 'medium', 'large', IgnoreCase = $false)]
        [string]$size
    )
    begin {
        $progressBar = [pscustomobject]@{
            type  = 'ProgressBar'
            value = $value
            max   = $max
        }
    }
    process {
        if ($label) { $progressBar | Add-Member -NotePropertyName 'label' -NotePropertyValue $label }
        if ($color) { $progressBar | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($size)  { $progressBar | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
    }
    end {
        return $progressBar
    }
}

function New-AccordionPage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$headerTitle,

        [Parameter(Mandatory = $true)]
        [array]$items,

        [Parameter(Mandatory = $false)]
        [bool]$expanded,

        [Parameter(Mandatory = $false)]
        [string]$headerIconName,

        [Parameter(Mandatory = $false)]
        [bool]$showBorder,

        [Parameter(Mandatory = $false)]
        [bool]$roundedCorners,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'emphasis', 'accent', 'good', 'attention', 'warning', IgnoreCase = $true)]
        [string]$style
    )
    begin {
        $accordionPage = [pscustomobject]@{
            type        = 'AccordionPage'
            headerTitle = $headerTitle
            items       = @()
        }
    }
    process {
        foreach ($item in $items) {
            $accordionPage.items += $item
        }

        if ($expanded)       { $accordionPage | Add-Member -NotePropertyName 'expanded' -NotePropertyValue $expanded }
        if ($headerIconName) { $accordionPage | Add-Member -NotePropertyName 'headerIconName' -NotePropertyValue $headerIconName }
        if ($showBorder)     { $accordionPage | Add-Member -NotePropertyName 'showBorder' -NotePropertyValue $showBorder }
        if ($roundedCorners) { $accordionPage | Add-Member -NotePropertyName 'roundedCorners' -NotePropertyValue $roundedCorners }
        if ($style)          { $accordionPage | Add-Member -NotePropertyName 'style' -NotePropertyValue $style }
    }
    end {
        return $accordionPage
    }
}

function New-Accordion {
    param (
        [Parameter(Mandatory = $true)]
        [array]$pages,

        [Parameter(Mandatory = $false)]
        [bool]$allowCollapseAllPages = $true,

        [Parameter(Mandatory = $false)]
        [bool]$allowMultipleExpandedPages = $false
    )
    begin {
        $accordion = [pscustomobject]@{
            type                      = 'Accordion'
            pages                     = @()
            allowCollapseAllPages     = $allowCollapseAllPages
            allowMultipleExpandedPages = $allowMultipleExpandedPages
        }
    }
    process {
        foreach ($page in $pages) {
            $accordion.pages += $page
        }
    }
    end {
        return $accordion
    }
}

function New-VerticalBarChart {
    param (
        [Parameter(Mandatory = $true)]
        [array]$data,

        [Parameter(Mandatory = $false)]
        [string]$title,

        [Parameter(Mandatory = $false)]
        [string]$xAxisTitle,

        [Parameter(Mandatory = $false)]
        [string]$yAxisTitle,

        [Parameter(Mandatory = $false)][ValidateSet(
            'good',
            'warning',
            'attention',
            'neutral',
            'categoricalRed',
            'categoricalPurple',
            'categoricalLavender',
            'categoricalBlue',
            'categoricalLightBlue',
            'categoricalTeal',
            'categoricalGreen',
            'categoricalLime',
            'categoricalMarigold',
            'sequential1',
            'sequential2',
            'sequential3',
            'sequential4',
            'sequential5',
            'sequential6',
            'sequential7',
            'sequential8',
            'divergingBlue',
            'divergingLightBlue',
            'divergingCyan',
            'divergingTeal',
            'divergingYellow',
            'divergingPeach',
            'divergingLightRed',
            'divergingRed',
            'divergingMaroon',
            'divergingGray',
            'sequentialRed1',
            'sequentialRed2',
            'sequentialRed3',
            'sequentialRed4',
            'sequentialRed5',
            'sequentialRed6',
            'sequentialRed7',
            'sequentialRed8',
            'sequentialGreen1',
            'sequentialGreen2',
            'sequentialGreen3',
            'sequentialGreen4',
            'sequentialGreen5',
            'sequentialGreen6',
            'sequentialGreen7',
            'sequentialGreen8',
            'sequentialYellow1',
            'sequentialYellow2',
            'sequentialYellow3',
            'sequentialYellow4',
            'sequentialYellow5',
            'sequentialYellow6',
            'sequentialYellow7',
            'sequentialYellow8',
            IgnoreCase = $false
        )]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet(
            'categorical',
            'sequential',
            'sequentialred',
            'sequentialgreen',
            'sequentialyellow',
            'diverging',
            IgnoreCase = $false
        )]
        [string]$colorSet,

        [Parameter(Mandatory = $false)]
        [bool]$showBarValues,

        [Parameter(Mandatory = $false)]
        [bool]$showLegend,

        [Parameter(Mandatory = $false)]
        [bool]$showTitle,

        [Parameter(Mandatory = $false)]
        [double]$yMin,

        [Parameter(Mandatory = $false)]
        [double]$yMax,

        [Parameter(Mandatory = $false)][ValidateSet('auto', 'stretch', IgnoreCase = $false)]
        [string]$height,

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)]
        [string]$id,

        [Parameter(Mandatory = $false)]
        [bool]$isSortKey,

        [Parameter(Mandatory = $false)]
        [bool]$isVisible,

        [Parameter(Mandatory = $false)]
        [string]$key,

        [Parameter(Mandatory = $false)]
        [string]$lang,

        [Parameter(Mandatory = $false)][ValidatePattern("^\d+px$", ErrorMessage = "maxWidth must be a number followed by 'px'.")]
        [string]$maxWidth,

        [Parameter(Mandatory = $false)]
        [hashtable]$requires,

        [Parameter(Mandatory = $false)]
        [bool]$separator,

        [Parameter(Mandatory = $false)][ValidateSet(
            'none',
            'extraSmall',
            'small',
            'default',
            'medium',
            'large',
            'extraLarge',
            'padding',
            IgnoreCase = $false
        )]
        [string]$spacing,

        [Parameter(Mandatory = $false)][ValidateSet(
            'VeryNarrow',
            'Narrow',
            'Standard',
            'Wide',
            'atLeast:VeryNarrow',
            'atMost:VeryNarrow',
            'atLeast:Narrow',
            'atMost:Narrow',
            'atLeast:Standard',
            'atMost:Standard',
            'atLeast:Wide',
            'atMost:Wide',
            IgnoreCase = $false
        )]
        [string]$targetWidth
    )
    begin {
        $verticalBarChart = [pscustomobject]@{
            type = 'Chart.VerticalBar'
            data = @()
        }
    }
    process {
        foreach ($dataPoint in $data) {
            $point = [pscustomobject]@{
                x = $dataPoint.x
                y = $dataPoint.y
            }

            if ($dataPoint.PSObject.Properties['color']) {
                $point | Add-Member -NotePropertyName 'color' -NotePropertyValue $dataPoint.color
            }

            $verticalBarChart.data += $point
        }

        if ($color)                { $verticalBarChart | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($colorSet)             { $verticalBarChart | Add-Member -NotePropertyName 'colorSet' -NotePropertyValue $colorSet }
        if ($title)                { $verticalBarChart | Add-Member -NotePropertyName 'title' -NotePropertyValue $title }
        if ($xAxisTitle)           { $verticalBarChart | Add-Member -NotePropertyName 'xAxisTitle' -NotePropertyValue $xAxisTitle }
        if ($yAxisTitle)           { $verticalBarChart | Add-Member -NotePropertyName 'yAxisTitle' -NotePropertyValue $yAxisTitle }
        if ($showBarValues)        { $verticalBarChart | Add-Member -NotePropertyName 'showBarValues' -NotePropertyValue $showBarValues }
        if ($showLegend)           { $verticalBarChart | Add-Member -NotePropertyName 'showLegend' -NotePropertyValue $showLegend }
        if ($showTitle)            { $verticalBarChart | Add-Member -NotePropertyName 'showTitle' -NotePropertyValue $showTitle }
        if ($PSBoundParameters.ContainsKey('yMin')) { $verticalBarChart | Add-Member -NotePropertyName 'yMin' -NotePropertyValue $yMin }
        if ($PSBoundParameters.ContainsKey('yMax')) { $verticalBarChart | Add-Member -NotePropertyName 'yMax' -NotePropertyValue $yMax }
        if ($height)               { $verticalBarChart | Add-Member -NotePropertyName 'height' -NotePropertyValue $height }
        if ($horizontalAlignment)  { $verticalBarChart | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($id)                   { $verticalBarChart | Add-Member -NotePropertyName 'id' -NotePropertyValue $id }
        if ($isSortKey)            { $verticalBarChart | Add-Member -NotePropertyName 'isSortKey' -NotePropertyValue $isSortKey }
        if ($PSBoundParameters.ContainsKey('isVisible')) { $verticalBarChart | Add-Member -NotePropertyName 'isVisible' -NotePropertyValue $isVisible }
        if ($key)                  { $verticalBarChart | Add-Member -NotePropertyName 'key' -NotePropertyValue $key }
        if ($lang)                 { $verticalBarChart | Add-Member -NotePropertyName 'lang' -NotePropertyValue $lang }
        if ($maxWidth)             { $verticalBarChart | Add-Member -NotePropertyName 'maxWidth' -NotePropertyValue $maxWidth }
        if ($requires)             { $verticalBarChart | Add-Member -NotePropertyName 'requires' -NotePropertyValue $requires }
        if ($separator)            { $verticalBarChart | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($spacing)              { $verticalBarChart | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($targetWidth)          { $verticalBarChart | Add-Member -NotePropertyName 'targetWidth' -NotePropertyValue $targetWidth }
    }
    end {
        return $verticalBarChart
    }
}

function New-GaugeChart {
    param (
        [Parameter(Mandatory = $true)]
        [double]$value,

        [Parameter(Mandatory = $true)]
        [array]$segments,

        [Parameter(Mandatory = $false)]
        [double]$min = 0,

        [Parameter(Mandatory = $false)]
        [double]$max = 100,

        [Parameter(Mandatory = $false)]
        [string]$title,

        [Parameter(Mandatory = $false)][ValidateSet('fraction', 'percentage', IgnoreCase = $false)]
        [string]$valueFormat,

        [Parameter(Mandatory = $false)][ValidateSet('categorical', 'sequential', 'sequentialRed', 'sequentialGreen', 'sequentialYellow', 'diverging', IgnoreCase = $false)]
        [string]$colorSet,

        [Parameter(Mandatory = $false)]
        [bool]$showTitle,

        [Parameter(Mandatory = $false)]
        [bool]$showLegend,

        [Parameter(Mandatory = $false)]
        [bool]$showMinMax,

        [Parameter(Mandatory = $false)]
        [bool]$showNeedle,

        [Parameter(Mandatory = $false)]
        [bool]$showOutlines,

        [Parameter(Mandatory = $false)]
        [string]$id,

        [Parameter(Mandatory = $false)][ValidateSet('auto', 'stretch', IgnoreCase = $false)]
        [string]$height,

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)]
        [bool]$isSortKey,

        [Parameter(Mandatory = $false)]
        [bool]$isVisible,

        [Parameter(Mandatory = $false)]
        [string]$key,

        [Parameter(Mandatory = $false)]
        [string]$lang,

        [Parameter(Mandatory = $false)][ValidatePattern("^\d+px$", ErrorMessage = "maxWidth must be a number followed by 'px'.")]
        [string]$maxWidth,

        [Parameter(Mandatory = $false)]
        [hashtable]$requires,

        [Parameter(Mandatory = $false)]
        [bool]$separator,

        [Parameter(Mandatory = $false)][ValidateSet('none', 'extraSmall', 'small', 'default', 'medium', 'large', 'extraLarge', 'padding', IgnoreCase = $false)]
        [string]$spacing,

        [Parameter(Mandatory = $false)][ValidateSet('VeryNarrow', 'Narrow', 'Standard', 'Wide', 'atLeast:VeryNarrow', 'atMost:VeryNarrow', 'atLeast:Narrow', 'atMost:Narrow', 'atLeast:Standard', 'atMost:Standard', 'atLeast:Wide', 'atMost:Wide', IgnoreCase = $false)]
        [string]$targetWidth
    )
    begin {
        $gauge = [pscustomobject]@{
            type     = 'Chart.Gauge'
            value    = $value
            segments = @()
        }
    }
    process {
        foreach ($segment in $segments) {
            $gaugeSegment = [pscustomobject]@{
                legend = $segment.legend
                size   = $segment.size
            }

            if ($segment.ContainsKey('color')) {
                $gaugeSegment | Add-Member -NotePropertyName 'color' -NotePropertyValue $segment.color
            }

            $gauge.segments += $gaugeSegment
        }

        if ($min -ne 0)        { $gauge | Add-Member -NotePropertyName 'min' -NotePropertyValue $min }
        if ($max -ne 100)       { $gauge | Add-Member -NotePropertyName 'max' -NotePropertyValue $max }
        if ($title)             { $gauge | Add-Member -NotePropertyName 'title' -NotePropertyValue $title }
        if ($valueFormat)       { $gauge | Add-Member -NotePropertyName 'valueFormat' -NotePropertyValue $valueFormat }
        if ($colorSet)          { $gauge | Add-Member -NotePropertyName 'colorSet' -NotePropertyValue $colorSet }
        if ($showTitle)         { $gauge | Add-Member -NotePropertyName 'showTitle' -NotePropertyValue $showTitle }
        if ($showLegend)        { $gauge | Add-Member -NotePropertyName 'showLegend' -NotePropertyValue $showLegend }
        if ($showMinMax)        { $gauge | Add-Member -NotePropertyName 'showMinMax' -NotePropertyValue $showMinMax }
        if ($showNeedle)        { $gauge | Add-Member -NotePropertyName 'showNeedle' -NotePropertyValue $showNeedle }
        if ($showOutlines)      { $gauge | Add-Member -NotePropertyName 'showOutlines' -NotePropertyValue $showOutlines }
        if ($id)                { $gauge | Add-Member -NotePropertyName 'id' -NotePropertyValue $id }
        if ($height)            { $gauge | Add-Member -NotePropertyName 'height' -NotePropertyValue $height }
        if ($horizontalAlignment) { $gauge | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($isSortKey)         { $gauge | Add-Member -NotePropertyName 'isSortKey' -NotePropertyValue $isSortKey }
        if ($PSBoundParameters.ContainsKey('isVisible')) { $gauge | Add-Member -NotePropertyName 'isVisible' -NotePropertyValue $isVisible }
        if ($key)               { $gauge | Add-Member -NotePropertyName 'key' -NotePropertyValue $key }
        if ($lang)              { $gauge | Add-Member -NotePropertyName 'lang' -NotePropertyValue $lang }
        if ($maxWidth)          { $gauge | Add-Member -NotePropertyName 'maxWidth' -NotePropertyValue $maxWidth }
        if ($requires)          { $gauge | Add-Member -NotePropertyName 'requires' -NotePropertyValue $requires }
        if ($separator)         { $gauge | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($spacing)           { $gauge | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($targetWidth)       { $gauge | Add-Member -NotePropertyName 'targetWidth' -NotePropertyValue $targetWidth }
    }
    end {
        return $gauge
    }
}

function New-HorizontalBarChart {
    param (
        [Parameter(Mandatory = $true)]
        [array]$data,

        [Parameter(Mandatory = $false)]
        [string]$title,

        [Parameter(Mandatory = $false)]
        [string]$xAxisTitle,

        [Parameter(Mandatory = $false)]
        [string]$yAxisTitle,

        [Parameter(Mandatory = $false)][ValidateSet('AbsoluteWithAxis', 'AbsoluteNoAxis', 'PartToWhole', IgnoreCase = $false)]
        [string]$displayMode,

        [Parameter(Mandatory = $false)][ValidateSet(
            'good',
            'warning',
            'attention',
            'neutral',
            'categoricalRed',
            'categoricalPurple',
            'categoricalLavender',
            'categoricalBlue',
            'categoricalLightBlue',
            'categoricalTeal',
            'categoricalGreen',
            'categoricalLime',
            'categoricalMarigold',
            'sequential1',
            'sequential2',
            'sequential3',
            'sequential4',
            'sequential5',
            'sequential6',
            'sequential7',
            'sequential8',
            'divergingBlue',
            'divergingLightBlue',
            'divergingCyan',
            'divergingTeal',
            'divergingYellow',
            'divergingPeach',
            'divergingLightRed',
            'divergingRed',
            'divergingMaroon',
            'divergingGray',
            IgnoreCase = $false
        )]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet(
            'categorical',
            'sequential',
            'sequentialRed',
            'sequentialGreen',
            'sequentialYellow',
            'diverging',
            IgnoreCase = $false
        )]
        [string]$colorSet,

        [Parameter(Mandatory = $false)]
        [bool]$showBarValues,

        [Parameter(Mandatory = $false)][ValidateSet('auto', 'stretch', IgnoreCase = $false)]
        [string]$height,

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)]
        [string]$id,

        [Parameter(Mandatory = $false)]
        [bool]$isVisible,

        [Parameter(Mandatory = $false)]
        [string]$lang,

        [Parameter(Mandatory = $false)]
        [hashtable]$requires,

        [Parameter(Mandatory = $false)]
        [bool]$separator,

        [Parameter(Mandatory = $false)][ValidateSet('none', 'small', 'default', 'medium', 'large', 'extraLarge', 'padding', IgnoreCase = $false)]
        [string]$spacing,

        [Parameter(Mandatory = $false)][ValidateSet(
            'VeryNarrow',
            'Narrow',
            'Standard',
            'Wide',
            'atLeast:VeryNarrow',
            'atMost:VeryNarrow',
            'atLeast:Narrow',
            'atMost:Narrow',
            'atLeast:Standard',
            'atMost:Standard',
            'atLeast:Wide',
            'atMost:Wide',
            IgnoreCase = $false
        )]
        [string]$targetWidth
    )
    begin {
        $horizontalBarChart = [pscustomobject]@{
            type = 'Chart.HorizontalBar'
            data = @()
        }
    }
    process {
        foreach ($dataPoint in $data) {
            $point = [pscustomobject]@{
                x = $dataPoint.x
                y = $dataPoint.y
            }

            if ($dataPoint.ContainsKey('color')) {
                $point | Add-Member -NotePropertyName 'color' -NotePropertyValue $dataPoint.color
            }

            $horizontalBarChart.data += $point
        }

        if ($title)                { $horizontalBarChart | Add-Member -NotePropertyName 'title' -NotePropertyValue $title }
        if ($xAxisTitle)           { $horizontalBarChart | Add-Member -NotePropertyName 'xAxisTitle' -NotePropertyValue $xAxisTitle }
        if ($yAxisTitle)           { $horizontalBarChart | Add-Member -NotePropertyName 'yAxisTitle' -NotePropertyValue $yAxisTitle }
        if ($displayMode)          { $horizontalBarChart | Add-Member -NotePropertyName 'displayMode' -NotePropertyValue $displayMode }
        if ($color)                { $horizontalBarChart | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($colorSet)             { $horizontalBarChart | Add-Member -NotePropertyName 'colorSet' -NotePropertyValue $colorSet }
        if ($showBarValues)        { $horizontalBarChart | Add-Member -NotePropertyName 'showBarValues' -NotePropertyValue $showBarValues }
        if ($height)               { $horizontalBarChart | Add-Member -NotePropertyName 'height' -NotePropertyValue $height }
        if ($horizontalAlignment)  { $horizontalBarChart | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($id)                   { $horizontalBarChart | Add-Member -NotePropertyName 'id' -NotePropertyValue $id }
        if ($PSBoundParameters.ContainsKey('isVisible')) { $horizontalBarChart | Add-Member -NotePropertyName 'isVisible' -NotePropertyValue $isVisible }
        if ($lang)                 { $horizontalBarChart | Add-Member -NotePropertyName 'lang' -NotePropertyValue $lang }
        if ($requires)             { $horizontalBarChart | Add-Member -NotePropertyName 'requires' -NotePropertyValue $requires }
        if ($separator)            { $horizontalBarChart | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($spacing)              { $horizontalBarChart | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($targetWidth)          { $horizontalBarChart | Add-Member -NotePropertyName 'targetWidth' -NotePropertyValue $targetWidth }
    }
    end {
        return $horizontalBarChart
    }
}

function Send-JsonToTeamsWebhook {
    param (
        [Parameter(Mandatory = $true)]
        [string]$webhookURI,

        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [pscustomobject]$adaptiveCard,

        [Parameter(Mandatory = $false)]
        [switch]$fullWidth,

        [Parameter(Mandatory = $false)]
        [switch]$onlyConvertToJson
    )

    $attachment = [pscustomobject]@{
        contentType = 'application/vnd.microsoft.card.adaptive'
        contentUrl = $null
        content = $adaptiveCard
    }

    $message = [pscustomobject]@{
        type = 'message'
        attachments = @()
    }
    $message.attachments += $attachment

    if ($fullWidth) {
        $msteamsProperty = @{
            width = 'Full'
        }
        $message.attachments[0].content | Add-Member -MemberType NoteProperty -Name msteams -Value $msteamsProperty
    }

    $json = ($message | ConvertTo-Json -Depth 20) -replace '\\\\', '\' #-replace "\\", '&#92;'
    
    if ($onlyConvertToJson) {
        Write-Output $json
        Break
    }

    $parameters = @{
        "URI"         = $webhookURI
        "Method"      = 'POST'
        "Body"        = $json
        "ContentType" = 'application/json; charset=UTF-8'
        "ErrorAction" = 'Stop'
    }
    try {
        Invoke-RestMethod @parameters
    }
    catch {
        Write-Error "Failed to send request: $($_.Exception.Message)"
    }
}
