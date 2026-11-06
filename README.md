![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)
![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSAdaptiveCard)

# PSAdaptiveCard
The PSAdaptiveCard PowerShell module designed to create JSON-formatted Adaptive Cards. Adaptive Cards are a way to present content in Microsoft Teams. This module includes functions to construct various elements of an Adaptive Card, and a function for posting these messages to Teams channels using Workflow webhooks.

---

## Install
`Install-Module -Name PSAdaptiveCard`

---

## Getting started with webhook workflows in Teams
To get started with worksflows and generate a webhook URL, look [here](https://support.microsoft.com/en-us/office/post-a-workflow-when-a-webhook-request-is-received-in-microsoft-teams-8ae491c7-0394-4861-ba59-055e33f75498#:~:text=An%20Incoming%20webhook%20lets%20external,a%20webhook%20request%20is%20received.&text=next%20to%20the%20channel%20or,for%2C%20and%20then%20select%20Workflows.).

---

## A minimal complete example of use
```PowerShell
#Use functions to add a header-text and a table to a card.
$cardContent = @(
    New-TextBlock -Size extraLarge -Weight bolder -Text 'Services'
    Get-Service | Select-Object Name, DisplayName, Status -First 5 | New-Table -HighlightValueMatch 'Stopped' -HighlightValueStyle 'attention'
)

#We must provide a webhook-URL. This is a partial example:
$WebhookURI = 'https://prod-140.westeurope.logic.azure.com:443/workflows/[REDACTED]/triggers/manual...'

#Put the elements inside an Adaptive Card, then wrap the card as an attachment to a message, convert it to JSON, and POST it to webhook URL:
New-AdaptiveCard -BodyContent $cardContent | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```
This results in the following message to your Teams-channel:
<img width="926" height="428" alt="image" src="https://github.com/user-attachments/assets/ee52cfdb-770d-43c7-a036-ee9351a2e8ae" />

---

## Functions

### New-AdaptiveCard

Creates a new Adaptive Card object using the supplied body content.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `BodyContent` | Array | Yes | Array of Adaptive Card elements to include in the card body. |

**Example**

```PwSh
$cardContent = @(
    New-TextBlock -Text "Hello, Teams!"
    New-TextBlock -Text "Welcome to the channel."
)

$card = New-AdaptiveCard -BodyContent $cardContent
$card | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="927" height="216" alt="image" src="https://github.com/user-attachments/assets/0a7eb0b7-d459-4786-86a5-d60dde6113cf" />

---

### New-TextBlock

Creates a TextBlock element with optional formatting and layout properties.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Text` | String | Yes | Text content to display. |
| `isSubtle` | Boolean | No | Displays the text using a subtle style. |
| `separator` | Boolean | No | Adds a separator above the element. |
| `maxLines` | Integer | No | Maximum number of lines to display. |
| `size` | String | No | Text size: `Default`, `Small`, `Medium`, `Large`, `ExtraLarge`. |
| `weight` | String | No | Font weight: `default`, `lighter`, `bolder`. |
| `wrap` | Boolean | No | Enables text wrapping. |
| `color` | String | No | Text color: `Default`, `Dark`, `Light`, `Accent`, `Good`, `Warning`, `Attention`. |
| `fontType` | String | No | Font type: `default`, `monospace`. |
| `horizontalAlignment` | String | No | Text alignment: `left`, `center`, `right`. |
| `spacing` | String | No | Spacing: `Default`, `None`, `ExtraSmall`, `Small`, `Medium`, `Large`, `ExtraLarge`, `Padding`. |

**Example**

```PwSh
$cardContent = @(
    New-TextBlock `
    -Text "System Status: Operational" `
    -size "Large" `
    -weight "bolder" `
    -color "Good" `
    -wrap $true
)
New-AdaptiveCard -BodyContent $cardContent | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```
<img width="927" height="189" alt="image" src="https://github.com/user-attachments/assets/e4e6b747-434b-4b3e-a967-e0993325ddba" />

---

### New-CodeBlock

Creates a code block with syntax highlighting and optional starting line number.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `codeSnippet` | String | Yes | Code to display. |
| `language` | String | Yes | Programming language used for syntax highlighting. |
| `startLineNumber` | Integer | No | Starting line number displayed in the code block. |

**Supported languages**

`Bash`, `C`, `C++`, `C#`, `CSS`, `DOS`, `Go`, `GraphQL`, `HTML`, `Java`, `JavaScript`, `JSON`, `Perl`, `PHP`, `PlainText`, `PowerShell`, `Python`, `SQL`, `TypeScript`, `Visual Basic`, `Verilog`, `VHDL`, `XML`.

**Example**

```PwSh
$code = @"
Get-Service |
Where-Object Status -eq 'Running'
"@

$codeBlock = New-CodeBlock -codeSnippet $code -language "PowerShell" -startLineNumber 1

New-AdaptiveCard -BodyContent $codeBlock | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```
<img width="925" height="219" alt="image" src="https://github.com/user-attachments/assets/d89744fb-3074-4530-8c54-f1546a5d945e" />

---

### New-Table

Creates an Adaptive Card Table from one or more PowerShell objects.

The function automatically detects the object's NoteProperty properties and uses them as table columns. Each input object becomes a row.

Supports optional highlighting of cell values based on a regular expression match.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `object` | PSObject | Yes | Object used to generate a table row. Accepts pipeline input. |
| `highlightValueMatch` | String | No | Regular expression used to identify values that should be highlighted. |
| `highlightValueStyle` | String | No | Highlight color: `dark`, `light`, `accent`, `good`, `warning`, `attention`. |
| `firstRowAsHeader` | Boolean | No | Determines whether the first row is treated as a header. Default: `$true`. |
| `headerRowStyle` | String | No | Header style: `default`, `dark`, `light`, `accent`, `good`, `warning`, `attention`. |
| `showGridLines` | Boolean | No | Determines whether grid lines are displayed. Default: `$true`. |
| `gridStyle` | String | No | Grid style: `default`, `dark`, `light`, `accent`, `good`, `warning`, `attention`. Default: `default`. |
| `horizontalCellContentAlignment` | String | No | Horizontal alignment: `left`, `center`, `right`. |
| `verticalCellContentAlignment` | String | No | Vertical alignment: `top`, `center`, `bottom`. |

**Example 1**

```PwSh
$data = @(
    [PSCustomObject]@{
        Name = "Server01"
        Status = "Online"
        CPU = "25%"
    }
    [PSCustomObject]@{
        Name = "Server02"
        Status = "Offline"
        CPU = "0%"
    }
)
$table = $data | New-Table
New-AdaptiveCard -BodyContent $table | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="924" height="276" alt="image" src="https://github.com/user-attachments/assets/76fbcebc-7283-4534-9ce8-3ab499fc23e3" />

**Example 2**

```PwSh
$table = Get-Service W32Time, Spooler | Select-Object Name, DisplayName, Status | New-Table
New-AdaptiveCard -BodyContent $table | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="920" height="269" alt="image" src="https://github.com/user-attachments/assets/02895525-7095-49b4-9e5d-139e56c3805f" />

---

### New-Image

Creates an Image element with optional size, alignment, and display properties.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `url` | String | Yes | Image URL. |
| `altText` | String | No | Alternative text for the image. Default: `image`. |
| `backgroundColor` | String | No | Background color of the image element. |
| `height` | String | No | Image height: `auto`, `stretch`, or a pixel value such as `200px`. Default: `auto`. |
| `horizontalAlignment` | String | No | Alignment: `left`, `center`, `right`. |
| `size` | String | No | Image size: `auto`, `stretch`, `small`, `medium`, `large`. |
| `style` | String | No | Image style: `default`, `person`. |
| `width` | String | No | Image width in pixels, either as an integer or with the `px` suffix. |
| `spacing` | String | No | Spacing: `default`, `none`, `small`, `medium`, `large`, `extraLarge`, `padding`. Default: `default`. |
| `separator` | Boolean | No | Adds a separator above the image. |

**Example**

```PwSh
$img = New-Image -url "https://adaptivecards.microsoft.com/images/painting3.jpeg" -altText "Painting" -horizontalAlignment "center"
New-AdaptiveCard -BodyContent $img | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="920" height="592" alt="image" src="https://github.com/user-attachments/assets/868e30f9-109c-435c-ae86-6330a50b978a" />

---

### New-Icon

Creates an Icon element.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `name` | String | Yes | Icon name from the Adaptive Cards icon catalog. |
| `color` | String | No | Icon color: `Default`, `Dark`, `Light`, `Accent`, `Good`, `Warning`, `Attention`. |
| `size` | String | No | Icon size: `Standard`, `xxSmall`, `xSmall`, `Small`, `Medium`, `Large`, `xLarge`, `xxlarge`. |
| `spacing` | String | No | Spacing: `Default`, `none`, `small`, `medium`, `large`, `extraLarge`, `xxLarge`. |
| `horizontalAlignment` | String | No | Alignment: `Left`, `Center`, `Right`. |
| `separator` | Boolean | No | Adds a separator above the icon. |
| `isVisible` | Boolean | No | Controls icon visibility. Default: `$true`. |

**Example**

```PwSh
$icon = New-Icon -name "CheckmarkCircle" -color "Good" -size "Large" -horizontalAlignment "Center"
New-AdaptiveCard -BodyContent $icon | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="923" height="222" alt="image" src="https://github.com/user-attachments/assets/b309f123-7644-4b32-a06d-ba11bfbcb99c" />

---

### New-ColumnSet

Creates a ColumnSet element containing one or more columns.

Columns can contain other Adaptive Card elements, allowing more complex layouts.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Columns` | Array | Yes | Array of column objects. |
| `separator` | Boolean | No | Adds a separator above the ColumnSet. |
| `spacing` | String | No | Spacing: `Default`, `None`, `ExtraSmall`, `Small`, `Medium`, `Large`, `ExtraLarge`, `Padding`. |
| `horizontalAlignment` | String | No | Alignment: `Left`, `Center`, `Right`. |
| `bleed` | Boolean | No | Allows the ColumnSet to bleed into surrounding container padding. |

**Example**

```PwSh
$columns = @(
@{
type = "Column"
width = "stretch"
items = @(
New-TextBlock -Text "Server 01" -weight "bolder"
New-TextBlock -Text "Online" -color "Good"
)
}
@{
type = "Column"
width = "stretch"
items = @(
New-TextBlock -Text "Server 02" -weight "bolder"
New-TextBlock -Text "Offline" -color "Attention"
)
}
)

New-ColumnSet -Columns $columns
```

---

### New-FactSet

Creates a FactSet element containing title/value pairs.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Facts` | Array | Yes | Array of objects containing `title` and `value` properties. Accepts pipeline input. |

**Example 1**

```PwSh
$facts = @(
    [PSCustomObject]@{
        title = "Hostname"
        value = "Server01"
    }
    [PSCustomObject]@{
        title = "Operating System"
        value = "Windows Server 2022"
    }
    [PSCustomObject]@{
        title = "Status"
        value = "Online"
    }
)
New-AdaptiveCard -BodyContent ($facts | New-FactSet) | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="926" height="239" alt="image" src="https://github.com/user-attachments/assets/0aac760c-cb1c-4610-ab9d-d354def469c9" />


**Example 2**

```PwSH
$Properties = @(    
    @{Name = 'title'; expression = {$_.Name}}
    @{Name = 'value'; expression = {$_.DisplayName}}    
)
$factSet = Get-Service | Select-Object $Properties -First 2 | New-FactSet
New-AdaptiveCard -BodyContent $factSet | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth

```

<img width="923" height="212" alt="image" src="https://github.com/user-attachments/assets/dca1eb46-1f3e-4e9d-b045-d2bc87f69df1" />


---

### New-ProgressBar

Creates a ProgressBar element.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `value` | Integer | Yes | Current progress value. Valid range: 0–100. |
| `max` | Integer | No | Maximum progress value. Default: `100`. |
| `label` | String | No | Label displayed alongside the progress bar. |
| `color` | String | No | Color: `default`, `accent`, `good`, `warning`, `attention`. |
| `size` | String | No | Size: `default`, `small`, `medium`, `large`. |

**Example**

```PwSh
$pb = New-ProgressBar -value 75 -max 100 -label "Deployment Progress" -color "good" -size "large"
New-AdaptiveCard -BodyContent $pb | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="926" height="176" alt="image" src="https://github.com/user-attachments/assets/457ec273-fb05-4fd9-8f7f-d5cd1f72cfc9" />


---

### New-AccordionPage

Creates an individual AccordionPage element.

An AccordionPage consists of a header and an array of Adaptive Card elements displayed when the page is expanded.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `headerTitle` | String | Yes | Title displayed in the accordion header. |
| `items` | Array | Yes | Array of Adaptive Card elements displayed inside the page. |
| `expanded` | Boolean | No | Determines whether the page is initially expanded. |
| `headerIconName` | String | No | Icon displayed in the header. |
| `showBorder` | Boolean | No | Controls whether the page displays a border. |
| `roundedCorners` | Boolean | No | Enables rounded corners. |
| `style` | String | No | Page style: `default`, `emphasis`, `accent`, `good`, `attention`, `warning`. |

**Example**

```PwSh
$accordionPage = New-AccordionPage -headerTitle "Server Information" -items @(
    New-TextBlock -Text "Hostname: Server01"
    New-TextBlock -Text "Status: Online" -color "Good"
) -expanded $true -showBorder $true -roundedCorners $true -style "emphasis"

$accordion = New-Accordion -pages $accordionPage -allowCollapseAllPages $true -allowMultipleExpandedPages $true

New-AdaptiveCard -BodyContent $accordion | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="922" height="295" alt="image" src="https://github.com/user-attachments/assets/c0d785ff-bb0d-47f0-b88f-9e82fb59a007" />


---

### New-Accordion

Creates an Accordion element containing one or more AccordionPage elements.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `pages` | Array | Yes | Array of AccordionPage objects. |
| `allowCollapseAllPages` | Boolean | No | Allows all pages to be collapsed. Default: `$true`. |
| `allowMultipleExpandedPages` | Boolean | No | Allows multiple pages to remain expanded simultaneously. Default: `$false`. |

**Example**

```PwSh
# See previous example (New-AccordionPage).
```

---

### New-VerticalBarChart

Creates a vertical bar chart using the `Chart.VerticalBar` Adaptive Card element.

Each data point requires an `x` and `y` property. An optional `color` property can be supplied for individual data points.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `data` | Array | Yes | Array of data points containing `x` and `y` properties. |
| `title` | String | No | Chart title. |
| `xAxisTitle` | String | No | X-axis title. |
| `yAxisTitle` | String | No | Y-axis title. |
| `color` | String | No | Chart color or palette value. |
| `colorSet` | String | No | Color set: `categorical`, `sequential`, `sequentialred`, `sequentialgreen`, `sequentialyellow`, `diverging`. |
| `showBarValues` | Boolean | No | Displays values on the bars. |
| `showLegend` | Boolean | No | Displays the chart legend. |
| `showTitle` | Boolean | No | Displays the chart title. |
| `yMin` | Double | No | Minimum Y-axis value. |
| `yMax` | Double | No | Maximum Y-axis value. |
| `height` | String | No | Height: `auto`, `stretch`. |
| `horizontalAlignment` | String | No | Alignment: `left`, `center`, `right`. |
| `id` | String | No | Element identifier. |
| `isSortKey` | Boolean | No | Determines whether the element is used as a sorting key. |
| `isVisible` | Boolean | No | Controls element visibility. |
| `key` | String | No | Element key. |
| `lang` | String | No | Language identifier. |
| `maxWidth` | String | No | Maximum width in pixels, e.g. `600px`. |
| `requires` | Hashtable | No | Required feature declarations. |
| `separator` | Boolean | No | Adds a separator above the chart. |
| `spacing` | String | No | Spacing: `none`, `extraSmall`, `small`, `default`, `medium`, `large`, `extraLarge`, `padding`. |
| `targetWidth` | String | No | Target width: `VeryNarrow`, `Narrow`, `Standard`, `Wide`, `atLeast:VeryNarrow`, `atMost:VeryNarrow`, `atLeast:Narrow`, `atMost:Narrow`, `atLeast:Standard`, `atMost:Standard`, `atLeast:Wide`, `atMost:Wide`. |

**Color values**

`good`, `warning`, `attention`, `neutral`, `categoricalRed`, `categoricalPurple`, `categoricalLavender`, `categoricalBlue`, `categoricalLightBlue`, `categoricalTeal`, `categoricalGreen`, `categoricalLime`, `categoricalMarigold`, `sequential1`–`sequential8`, `divergingBlue`, `divergingLightBlue`, `divergingCyan`, `divergingTeal`, `divergingYellow`, `divergingPeach`, `divergingLightRed`, `divergingRed`, `divergingMaroon`, `divergingGray`, `sequentialRed1`–`sequentialRed8`, `sequentialGreen1`–`sequentialGreen8`, `sequentialYellow1`–`sequentialYellow8`.

**Example**

```PwSh
$data = @(
    @{ x = "January"; y = 120 }
    @{ x = "February"; y = 180 }
    @{ x = "March"; y = 150 }
    @{ x = "April"; y = 220 }
)

$verticalBarChart = New-VerticalBarChart `
-data $data `
-title "Monthly Sales" `
-xAxisTitle "Month" `
-yAxisTitle "Sales" `
-showTitle $true `
-showBarValues $true `
-showLegend $true `
-colorSet "categorical"

New-AdaptiveCard -BodyContent $verticalBarChart | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="923" height="532" alt="image" src="https://github.com/user-attachments/assets/a681d876-c0d1-4e70-8420-21938548bede" />


---

### New-GaugeChart

Creates a gauge chart using the `Chart.Gauge` Adaptive Card element.

The gauge consists of a current value and one or more segments. Each segment defines a legend, size, and optional color.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `value` | Double | Yes | Current gauge value. |
| `segments` | Array | Yes | Array of segments containing `legend` and `size` properties. Optional `color` property supported. |
| `min` | Double | No | Minimum gauge value. Default: `0`. |
| `max` | Double | No | Maximum gauge value. Default: `100`. |
| `title` | String | No | Gauge title. |
| `valueFormat` | String | No | Value format: `fraction`, `percentage`. |
| `colorSet` | String | No | Color set: `categorical`, `sequential`, `sequentialRed`, `sequentialGreen`, `sequentialYellow`, `diverging`. |
| `showTitle` | Boolean | No | Displays the gauge title. |
| `showLegend` | Boolean | No | Displays the gauge legend. |
| `showMinMax` | Boolean | No | Displays minimum and maximum values. |
| `showNeedle` | Boolean | No | Displays the gauge needle. |
| `showOutlines` | Boolean | No | Displays segment outlines. |
| `id` | String | No | Element identifier. |
| `height` | String | No | Height: `auto`, `stretch`. |
| `horizontalAlignment` | String | No | Alignment: `left`, `center`, `right`. |
| `isSortKey` | Boolean | No | Determines whether the element is used as a sorting key. |
| `isVisible` | Boolean | No | Controls element visibility. |
| `key` | String | No | Element key. |
| `lang` | String | No | Language identifier. |
| `maxWidth` | String | No | Maximum width in pixels, e.g. `600px`. |
| `requires` | Hashtable | No | Required feature declarations. |
| `separator` | Boolean | No | Adds a separator above the gauge. |
| `spacing` | String | No | Spacing: `none`, `extraSmall`, `small`, `default`, `medium`, `large`, `extraLarge`, `padding`. |
| `targetWidth` | String | No | Target width: `VeryNarrow`, `Narrow`, `Standard`, `Wide`, `atLeast:VeryNarrow`, `atMost:VeryNarrow`, `atLeast:Narrow`, `atMost:Narrow`, `atLeast:Standard`, `atMost:Standard`, `atLeast:Wide`, `atMost:Wide`. |

**Example**

```PwSh
$segments = @(
    @{
        legend = "Completed"
        size = 75
        color = "good"
    }
    @{
        legend = "Remaining"
        size = 25
        color = "neutral"
    }
)

$gaugeChart = New-GaugeChart `
-value 75 `
-segments $segments `
-min 0 `
-max 100 `
-title "Deployment Status" `
-valueFormat "percentage" `
-showTitle $true `
-showLegend $true `
-showMinMax $true `
-showNeedle $true

New-AdaptiveCard -BodyContent $gaugeChart | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="923" height="350" alt="image" src="https://github.com/user-attachments/assets/282719f9-5c87-4017-97f2-4687614771a9" />


---

### New-HorizontalBarChart

Creates a horizontal bar chart using the `Chart.HorizontalBar` Adaptive Card element.

Each data point requires an `x` and `y` property. An optional `color` property can be supplied for individual data points.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `data` | Array | Yes | Array of data points containing `x` and `y` properties. |
| `title` | String | No | Chart title. |
| `xAxisTitle` | String | No | X-axis title. |
| `yAxisTitle` | String | No | Y-axis title. |
| `displayMode` | String | No | Display mode: `AbsoluteWithAxis`, `AbsoluteNoAxis`, `PartToWhole`. |
| `color` | String | No | Chart color or palette value. |
| `colorSet` | String | No | Color set: `categorical`, `sequential`, `sequentialRed`, `sequentialGreen`, `sequentialYellow`, `diverging`. |
| `showBarValues` | Boolean | No | Displays values on the bars. |
| `height` | String | No | Height: `auto`, `stretch`. |
| `horizontalAlignment` | String | No | Alignment: `left`, `center`, `right`. |
| `id` | String | No | Element identifier. |
| `isVisible` | Boolean | No | Controls element visibility. |
| `lang` | String | No | Language identifier. |
| `requires` | Hashtable | No | Required feature declarations. |
| `separator` | Boolean | No | Adds a separator above the chart. |
| `spacing` | String | No | Spacing: `none`, `small`, `default`, `medium`, `large`, `extraLarge`, `padding`. |
| `targetWidth` | String | No | Target width: `VeryNarrow`, `Narrow`, `Standard`, `Wide`, `atLeast:VeryNarrow`, `atMost:VeryNarrow`, `atLeast:Narrow`, `atMost:Narrow`, `atLeast:Standard`, `atMost:Standard`, `atLeast:Wide`, `atMost:Wide`. |

**Color values**

`good`, `warning`, `attention`, `neutral`, `categoricalRed`, `categoricalPurple`, `categoricalLavender`, `categoricalBlue`, `categoricalLightBlue`, `categoricalTeal`, `categoricalGreen`, `categoricalLime`, `categoricalMarigold`, `sequential1`–`sequential8`, `divergingBlue`, `divergingLightBlue`, `divergingCyan`, `divergingTeal`, `divergingYellow`, `divergingPeach`, `divergingLightRed`, `divergingRed`, `divergingMaroon`, `divergingGray`.

**Example**

```PwSh
$data = @(
    @{ x = "CPU"; y = 65 }
    @{ x = "Memory"; y = 80 }
    @{ x = "Disk"; y = 45 }
    @{ x = "Network"; y = 30 }
)

$horizontalBarChart = New-HorizontalBarChart `
-data $data `
-title "Resource Utilization" `
-xAxisTitle "Utilization (%)" `
-yAxisTitle "Resource" `
-displayMode "AbsoluteWithAxis" `
-showBarValues $true `
-colorSet "categorical"

New-AdaptiveCard -BodyContent $horizontalBarChart | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

<img width="923" height="381" alt="image" src="https://github.com/user-attachments/assets/aac0a3f6-4968-41c1-9b74-f2336d14d84c" />


---

### Send-JsonToTeamsWebhook

Converts an Adaptive Card into a Teams-compatible message payload and sends it to a Microsoft Teams Workflow webhook.

The function can also return the generated JSON without sending it.

**Parameters**

| Parameter | Type | Required | Description |
|---|---|---|---|
| `webhookURI` | String | Yes | Microsoft Teams Workflow webhook URL. |
| `adaptiveCard` | PSCustomObject | Yes | Adaptive Card object to send. Accepts pipeline input. |
| `fullWidth` | Switch | No | Adds the Teams-specific property to request full-width card rendering. |
| `onlyConvertToJson` | Switch | No | Returns the generated JSON without sending a webhook request. |

**Example – Send an Adaptive Card to Teams**

```PwSh
$cardContent = @(
New-TextBlock `
-Text "Deployment completed successfully" `
-size "Large" `
-weight "bolder" `
-color "Good"

New-TextBlock `
-Text "All servers have been updated."
)

$card = New-AdaptiveCard -BodyContent $cardContent

Send-JsonToTeamsWebhook `
-webhookURI "https://example.webhook.office.com/..." `
-adaptiveCard $card `
-fullWidth
```

**Example – Generate JSON without sending (for debugging)**

````PwSh
$card = New-AdaptiveCard -BodyContent @(
New-TextBlock -Text "Hello, Teams!"
)

$json = Send-JsonToTeamsWebhook `
-webhookURI "https://example.webhook.office.com/..." `
-adaptiveCard $card `
-onlyConvertToJson

$json
````

**Example – Pipeline usage**

```PwSh
$card = New-AdaptiveCard -BodyContent @(
New-TextBlock -Text "Pipeline example"
)

$card | Send-JsonToTeamsWebhook `
-webhookURI "https://example.webhook.office.com/..."
```

---

## Complete Example

The following example demonstrates how multiple functions can be combined to create a complete Adaptive Card and send it to Microsoft Teams.

```PwSh
# --------------------------------------------------
# Get services on local computer and do some math
# --------------------------------------------------
$services = Get-Service | Select-Object Name, Status

$totalServices = $services.Count
$runningServices = @($services | Where-Object { $_.Status -eq 'Running' })
$stoppedServices = @($services | Where-Object { $_.Status -ne 'Running' })

$runningCount = $runningServices.Count
$stoppedCount = $stoppedServices.Count

$runningPercentage = if ($totalServices -gt 0) {
    [math]::Round(($runningCount / $totalServices) * 100, 1)
} else { 0 }

$stoppedPercentage = if ($totalServices -gt 0) {
    [math]::Round(($stoppedCount / $totalServices) * 100, 1)
} else { 0 }


# --------------------------------------------------
# Create horizontal bar chart
# --------------------------------------------------
$chartData = @(
    @{
        x = "Running ($runningCount | $runningPercentage%)"
        y = $runningCount
        color = "good"
    },
    @{
        x = "Not Running ($stoppedCount | $stoppedPercentage%)"
        y = $stoppedCount
        color = "attention"
    }
)
$chart = New-HorizontalBarChart -data $chartData -title "Windows Services Status" -xAxisTitle "Number of Services" -yAxisTitle "Status" -displayMode "AbsoluteWithAxis" -showBarValues $true


# --------------------------------------------------
# Create accordion with details on stopped services
# --------------------------------------------------
$first50 = $stoppedServices | Select-Object Name, Status -First 40 | New-Table -firstRowAsHeader:$true
$detailsPage = New-AccordionPage -headerTitle "Stopped services" -items @($first50) -roundedCorners $true -style warning
$detailsAccordion = New-Accordion -pages @($detailsPage)


# --------------------------------------------------
# Create Adaptive Card and send it
# --------------------------------------------------
$cardContent = @(
    New-TextBlock -text "Windows Services Overview" -size "Large" -weight "bolder" -color "Accent"
    New-TextBlock -text "Total services: $totalServices" -weight "bolder" -spacing "Medium"
    $chart
    $detailsAccordion
)

$adaptiveCard = New-AdaptiveCard -BodyContent $cardContent
Send-JsonToTeamsWebhook -webhookURI $webhookURI -adaptiveCard $adaptiveCard -fullWidth
```

<img width="924" height="623" alt="image" src="https://github.com/user-attachments/assets/cf895b85-0962-4cd6-866c-703147d745f7" />

---

## References
- Adaptive Cards: https://adaptivecards.io
- Microsoft Adaptive Cards: https://adaptivecards.microsoft.com/
- Text formatting: https://adaptivecards.microsoft.com/?topic=text-formatting
- Icon catalog: https://adaptivecards.microsoft.com/?topic=icon-catalog
