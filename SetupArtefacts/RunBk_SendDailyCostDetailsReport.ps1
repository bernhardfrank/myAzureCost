param (
    [parameter(Mandatory = $false,
        HelpMessage = "Enter a en-us formatted date e.g. '12/30/2019'")]
    [String]$myDate
)
$ErrorActionPreference = "Stop"

#region Consumption date
try {
    $ConsumptionDate = [dateTime]::Parse($myDate)
}
catch {
    $ConsumptionDate = [dateTime]::Today.AddDays(-1)      #default to yesterday
}
Write-Output "Get consumption of $($ConsumptionDate.ToString("dd'/'MM'/'yyyy"))"
#endregion

#region Automation variables & culture
[string]$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"
[string]$AcsEndpoint = Get-AutomationVariable -Name "AcsEndpoint"
[string]$SenderAddress = Get-AutomationVariable -Name "SenderAddress"
[string]$RecipientEmail = Get-AutomationVariable -Name "RecipientEmail"
[string]$CultureInfo = Get-AutomationVariable -Name "CultureInfo"
[string]$AzureCostStorageAccountName = Get-AutomationVariable -Name "AzureCostStorageAccountName"

try {
    $destculture = [CultureInfo]::new($CultureInfo)
    Write-Output "Attachment/number culture: $CultureInfo"
}
catch [System.Globalization.CultureNotFoundException] {
    Write-Output "Culture '$CultureInfo' not recognized, falling back to en-US."
    $destculture = [CultureInfo]::new("en-US")
}
#endregion

#region Helper functions

function Get-MiToken {
    # Gets a token for the System-assigned Managed Identity via the
    # Azure Automation sandbox identity endpoint.
    param([Parameter(Mandatory = $true)][string]$Resource)
    if (-not $env:IDENTITY_ENDPOINT -or -not $env:IDENTITY_HEADER) {
        throw "IDENTITY_ENDPOINT not available. Is this runbook running in Azure Automation with a managed identity enabled?"
    }
    $uri = "{0}?resource={1}" -f $env:IDENTITY_ENDPOINT, [uri]::EscapeDataString($Resource)
    $headers = @{
        "X-IDENTITY-HEADER" = $env:IDENTITY_HEADER
        "Metadata"          = "True"
    }
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -UseBasicParsing
    return $response.access_token
}

function Invoke-TableRequest {
    # Thin wrapper around the Azure Table REST API, authenticated with the managed identity.
    # Used instead of the AzTable module, since AzTable's legacy SDK doesn't support AAD tokens.
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Body
    )
    $headers = @{
        "Authorization" = "Bearer $(Get-MiToken -Resource 'https://storage.azure.com/')"
        "Accept"        = "application/json;odata=nometadata"
        "x-ms-version"  = "2020-12-06"
        "x-ms-date"     = [DateTime]::UtcNow.ToString("R")
    }
    if ($Body) {
        Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -Body $Body -ContentType "application/json" -UseBasicParsing
    }
    else {
        Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -UseBasicParsing
    }
}

function Set-CostTableEntity {
    # PUT = insert-or-replace: creates the row if missing, overwrites it if present.
    # TotalCost is stored with InvariantCulture ("." decimal separator) because it gets parsed
    # back later via [decimal]$_.TotalCost, and that cast always uses InvariantCulture regardless
    # of the process's current culture. Storing it with "{0:N7}" -f $TotalCost (culture-dependent)
    # would silently corrupt the value on any host whose culture uses "," as the decimal separator:
    # e.g. 15.234 -> stored as "15,2340000" -> read back (comma treated as a thousands separator,
    # not a decimal point) as 152340000.
    param($TableEndpoint, $TableName, $PartitionKey, $RowKey, $TotalCost, $Year)
    $uri = "{0}/{1}(PartitionKey='{2}',RowKey='{3}')" -f $TableEndpoint.TrimEnd('/'), $TableName, $PartitionKey, $RowKey
    $entity = @{ TotalCost = ([decimal]$TotalCost).ToString("F7", [CultureInfo]::InvariantCulture); Year = $Year } | ConvertTo-Json
    Invoke-TableRequest -Method Put -Uri $uri -Body $entity
}

function Get-CostTableEntity {
    param($TableEndpoint, $TableName, $PartitionKey, $RowKey)
    $uri = "{0}/{1}(PartitionKey='{2}',RowKey='{3}')" -f $TableEndpoint.TrimEnd('/'), $TableName, $PartitionKey, $RowKey
    try {
        Invoke-TableRequest -Method Get -Uri $uri
    }
    catch {
        if ([int]$_.Exception.Response.StatusCode -eq 404) { return $null }
        throw
    }
}

function Get-CostTableEntities {
    # Queries every row in a partition via the table's collection endpoint + an OData $filter.
    # Used to sum month-to-date cost from the same daily rows Set-CostTableEntity already writes,
    # instead of a second Cost Management API call.
    param($TableEndpoint, $TableName, [string]$Filter)
    $uri = "{0}/{1}()?`$filter={2}" -f $TableEndpoint.TrimEnd('/'), $TableName, [uri]::EscapeDataString($Filter)
    @((Invoke-TableRequest -Method Get -Uri $uri).value)
}

function Get-MonthToDateCost {
    # Sums TotalCost for every day in $AsOfDate's month/year up to and including $AsOfDate.
    # Call this AFTER Set-CostTableEntity has written $AsOfDate's own row.
    param($TableEndpoint, $TableName, [datetime]$AsOfDate)
    $filter = "PartitionKey eq '$($AsOfDate.ToString('MMMM', [CultureInfo]::InvariantCulture))' and Year eq $($AsOfDate.Year)"
    $rows = Get-CostTableEntities -TableEndpoint $TableEndpoint -TableName $TableName -Filter $filter
    $sum = $rows | Where-Object { [int]$_.RowKey -le $AsOfDate.Day } |
        ForEach-Object { [decimal]$_.TotalCost } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    if ($null -eq $sum) { $sum = 0 }
    return $sum
}

function New-ChartJsFunction {
    # Marks a raw JS function body to be spliced verbatim into the chart config by
    # ConvertTo-ChartJsConfig, instead of being serialized as an inert JSON string.
    param([Parameter(Mandatory)][string]$Code)
    "@@JSFUNC:$([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Code)))@@"
}

function ConvertTo-ChartJsConfig {
    # Serializes a chart config to a JS object-literal string, splicing in any New-ChartJsFunction
    # placeholders as real (unquoted) function code. QuickChart only executes datalabels formatter
    # functions when the whole "chart" field of the request is submitted as JS source rather than a
    # parsed JSON object - sending the hashtable as a nested JSON object (as before) leaves a
    # formatter string inert and the plugin just prints the raw numeric value. Verified against the
    # live quickchart.io API: chart-as-JS-string renders "70%"/"30%"; chart-as-JSON-object renders
    # "70"/"30".
    param([Parameter(Mandatory)]$Config)
    $json = $Config | ConvertTo-Json -Depth 10 -Compress
    [regex]::Replace($json, '"@@JSFUNC:([A-Za-z0-9+/=]+)@@"', {
            param($m) [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($m.Groups[1].Value))
        })
}

function Get-QuickChartPng {
    # Renders a Chart.js config to a PNG via quickchart.io. Needed because the PowerShell 7.2
    # Automation sandbox runs on Linux, so GDI+/System.Windows.Forms charting isn't available.
    param(
        [Parameter(Mandatory)]$ChartConfig,
        [int]$Width = 320,
        [int]$Height = 240
    )
    $requestBody = @{
        version         = "2"
        backgroundColor = "white"
        width           = $Width
        height          = $Height
        format          = "png"
        chart           = ConvertTo-ChartJsConfig -Config $ChartConfig
    } | ConvertTo-Json -Depth 10 -Compress

    # Invoke-WebRequest (not Invoke-RestMethod) is required here: Invoke-RestMethod can mangle
    # binary responses by coercing them into a string, corrupting the PNG bytes.
    $response = Invoke-WebRequest -Uri "https://quickchart.io/chart" -Method Post -Body $requestBody -ContentType "application/json" -UseBasicParsing
    return $response.Content
}

function Get-ChartPalette {
    # Office/Excel-like categorical palette. "other" always renders in neutral gray regardless of
    # its position, so it reads as "miscellaneous" rather than just another category.
    param([Parameter(Mandatory)][string[]]$Names)
    $palette = @("#0078D4", "#ED7D31", "#C00000", "#70AD47", "#5B2C87", "#268785", "#FFC000")
    $colors = @()
    $next = 0
    foreach ($name in $Names) {
        if ($name -eq "other") {
            $colors += "#A6A6A6"
        }
        else {
            $colors += $palette[$next % $palette.Count]
            $next++
        }
    }
    $colors
}

function New-PercentageDoughnutChartConfig {
    # Builds a doughnut chart driven by each item's share of the total (Percentage), not its
    # absolute currency amount, so the chart reads as "share of spend" rather than raw numbers.
    # Data labels are anchored just outside the ring (dark, bold) instead of inside the colored
    # slice, so they stay legible regardless of slice color or how thin the slice is. The legend
    # is off: Name + Percentage are already in the table below, and (Chart.js v2's legend option
    # lives at options.legend, not options.plugins.legend) keeping both risked them colliding at
    # the top of the ring - dropping the redundant legend removes that collision outright.
    param([Parameter(Mandatory)]$Result)
    @{
        type    = "doughnut"
        data    = @{
            labels   = @($Result | ForEach-Object { "$($_.Name) ($($_.Percentage)%)" })
            datasets = @(@{
                    data            = @($Result.Percentage)
                    backgroundColor = @(Get-ChartPalette -Names $Result.Name)
                })
        }
        options = @{
            legend  = @{ display = $false }
            layout  = @{ padding = @{ top = 32; bottom = 24; left = 32; right = 32 } }
            plugins = @{
                datalabels = @{
                    color     = "#222222"
                    anchor    = "end"
                    align     = "end"
                    offset    = 8
                    font      = @{ weight = "bold"; size = 12 }
                    # Value shown here is a percentage, not a cost - append "%" so it can never be
                    # misread as a currency amount.
                    formatter = New-ChartJsFunction 'function(value) { return value + "%"; }'
                }
            }
        }
    }
}

function New-HorizontalCostShareBarChartConfig {
    # One horizontal bar per item, sized by its actual cost - the x-axis reads in real currency -
    # but the axis is capped at the day's total cost (100% of spend) instead of being auto-scaled
    # to the largest single item. Chart.js's default auto-max reflects only the largest bar, which
    # visually exaggerates it and makes smaller shares hard to read; capping at the true total
    # keeps every bar's length proportional to its real share of total spend. Each bar is still
    # labeled with its percentage share (via a parallel "percentages" array + a datalabels
    # formatter), so the printed number is never mistaken for a currency amount.
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][string]$Currency, [Parameter(Mandatory)][decimal]$TotalCost)
    # The axis max must cover every bar's own value, not just $TotalCost rounded to 2 decimals -
    # on a low-usage subscription $TotalCost can itself be sub-cent, which would round to 0.00 and
    # hand Chart.js a degenerate [0, 0] range; it then falls back to an unrelated auto-scaled range
    # (observed: a negative one), which makes no sense for a cost axis. Rounding to 4 digits and
    # taking the larger of $TotalCost and the actual breakdown sum, with a tiny positive floor,
    # guarantees a valid, strictly-positive [0, max] range that fits every bar.
    $resultSum = ($Result.Sum | Measure-Object -Sum).Sum
    if (-not $resultSum) { $resultSum = 0 }
    $axisMax = [Math]::Round([double]([Math]::Max([double]$TotalCost, [double]$resultSum)), 4)
    if ($axisMax -le 0) { $axisMax = 0.0001 }
    @{
        type    = "horizontalBar"
        data    = @{
            labels   = @($Result.Name)
            datasets = @(@{
                    data            = @($Result.Sum | ForEach-Object { [Math]::Round([decimal]$_, 4) })
                    percentages     = @($Result.Percentage)
                    backgroundColor = @(Get-ChartPalette -Names $Result.Name)
                })
        }
        options = @{
            legend  = @{ display = $false }
            layout  = @{ padding = @{ top = 10; bottom = 10; left = 10; right = 44 } }
            scales  = @{
                xAxes = @(@{ ticks = @{ min = 0; max = $axisMax }; scaleLabel = @{ display = $true; labelString = $Currency } })
                yAxes = @(@{ gridLines = @{ display = $false } })
            }
            plugins = @{
                datalabels = @{
                    color     = "#333333"
                    anchor    = "end"
                    align     = "end"
                    font      = @{ size = 11; weight = "bold" }
                    formatter = New-ChartJsFunction 'function(value, context) { return context.dataset.percentages[context.dataIndex] + "%"; }'
                }
            }
        }
    }
}

function Get-CostBreakdown {
    # Groups the usage rows by $GroupProperty and returns Name/Count/Sum/Percentage, collapsing
    # every group at or below $OtherThresholdPercent into a single "other" row.
    param(
        [Parameter(Mandatory)]$Csv,
        [Parameter(Mandatory)][string]$GroupProperty,
        [Parameter(Mandatory)][decimal]$TotalCost,
        [decimal]$OtherThresholdPercent = 3
    )
    $groups = $Csv | Group-Object -Property $GroupProperty | ForEach-Object {
        $sum = ($_.Group | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum
        [PSCustomObject]@{
            Name       = "$($_.Name)"
            Count      = $_.Group.Count
            Sum        = $sum
            Percentage = [Math]::Round([decimal](($sum * 100) / $TotalCost), 2)
        }
    }

    $result = @($groups | Where-Object Percentage -gt $OtherThresholdPercent | Sort-Object Percentage -Descending)

    $otherGroups = @($groups | Where-Object Percentage -le $OtherThresholdPercent)
    if ($otherGroups.Count -gt 0) {
        $otherSum = ($otherGroups | Measure-Object Sum -Sum).Sum
        $result += [PSCustomObject]@{
            Name       = "other"
            Count      = ($otherGroups | Measure-Object Count -Sum).Sum
            Sum        = $otherSum
            Percentage = [Math]::Round([decimal](($otherSum * 100) / $TotalCost), 2)
        }
    }
    return $result
}

function Format-Money {
    param([Parameter(Mandatory)][decimal]$Value, [Parameter(Mandatory)]$Culture, [Parameter(Mandatory)][string]$Currency, [int]$DecimalDigits = 2)
    "{0} {1}" -f $Value.ToString("N$DecimalDigits", $Culture), $Currency
}

function New-ShareBarHtml {
    # A small colored "data bar" (Excel style) for a 0-100 percentage. Built from a 1-row,
    # 2-cell table with background colors rather than CSS width/height, since that's the
    # reliable way to render a bar in Outlook's Word rendering engine.
    param([Parameter(Mandatory)][double]$Percentage)
    $pct = [Math]::Round([Math]::Max(0, [Math]::Min(100, $Percentage)), 1)
    $fillWidth = [Math]::Max(1, [Math]::Round($pct))
    @"
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;border-collapse:collapse;" class="share-bar"><tr>
<td width="$fillWidth%" height="6" style="background:#0078D4;font-size:1px;line-height:6px;">&nbsp;</td>
<td height="6" style="background:#E8E8E8;font-size:1px;line-height:6px;">&nbsp;</td>
</tr></table>
<div style="font-size:12px;color:#666666;margin-top:1px;" class="share-pct">$($pct.ToString("N1", [CultureInfo]::InvariantCulture)) %</div>
"@
}

function New-ShareTableHtml {
    # 3-column table (Label / Amount / Share bar), blue header, zebra-striped rows. Used for the
    # cost-history table and each category/resource-group/region breakdown table.
    param(
        [Parameter(Mandatory)]$Rows,   # objects with .Label, .Amount (preformatted string), .Percentage (double)
        [Parameter(Mandatory)][string]$FirstColumnHeader,
        [string]$AmountHeader = "Cost",
        [string]$ShareHeader = "Share"
    )
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;border-collapse:collapse;">')
    [void]$sb.Append("<tr><td style=""background:#0078D4;color:#ffffff;font-weight:bold;padding:4px 6px;font-size:14px;"">$FirstColumnHeader</td><td align=""right"" style=""background:#0078D4;color:#ffffff;font-weight:bold;padding:4px 6px;font-size:14px;"">$AmountHeader</td><td width=""35%"" style=""background:#0078D4;color:#ffffff;font-weight:bold;padding:4px 6px;font-size:14px;"">$ShareHeader</td></tr>")
    $i = 0
    foreach ($row in $Rows) {
        $bg = if ($i % 2 -eq 0) { "#FFFFFF" } else { "#F3F3F3" }
        [void]$sb.Append("<tr><td style=""background:$bg;color:#222222;border-bottom:1px solid #DDDDDD;padding:4px 6px;font-size:14px;"">$([System.Net.WebUtility]::HtmlEncode($row.Label))</td>")
        [void]$sb.Append("<td align=""right"" style=""background:$bg;color:#222222;border-bottom:1px solid #DDDDDD;padding:4px 6px;font-size:14px;white-space:nowrap;"">$([System.Net.WebUtility]::HtmlEncode($row.Amount))</td>")
        [void]$sb.Append("<td style=""background:$bg;border-bottom:1px solid #DDDDDD;padding:4px 6px;"">$(New-ShareBarHtml -Percentage $row.Percentage)</td></tr>")
        $i++
    }
    [void]$sb.Append('</table>')
    $sb.ToString()
}

function New-PlainTableHtml {
    # Generic N-column table (blue header, zebra-striped rows), no share bar - used for Top 10 Consumers.
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][string[]]$Alignments,
        [Parameter(Mandatory)]$Rows,   # array of arrays: each inner array is one row's cell values, same order as $Headers
        [string[]]$ColumnClasses       # optional, same length as $Headers - e.g. "hide-mobile" to drop a column on narrow screens
    )
    if (-not $ColumnClasses) { $ColumnClasses = @('') * $Headers.Count }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;border-collapse:collapse;">')
    [void]$sb.Append('<tr>')
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        [void]$sb.Append("<td align=""$($Alignments[$c])"" class=""$($ColumnClasses[$c])"" style=""background:#0078D4;color:#ffffff;font-weight:bold;padding:4px 6px;font-size:14px;"">$($Headers[$c])</td>")
    }
    [void]$sb.Append('</tr>')
    $i = 0
    foreach ($rowCells in $Rows) {
        $bg = if ($i % 2 -eq 0) { "#FFFFFF" } else { "#F3F3F3" }
        [void]$sb.Append('<tr>')
        for ($c = 0; $c -lt $rowCells.Count; $c++) {
            [void]$sb.Append("<td align=""$($Alignments[$c])"" class=""$($ColumnClasses[$c])"" style=""background:$bg;color:#222222;border-bottom:1px solid #DDDDDD;padding:4px 6px;font-size:14px;"">$([System.Net.WebUtility]::HtmlEncode([string]$rowCells[$c]))</td>")
        }
        [void]$sb.Append('</tr>')
        $i++
    }
    [void]$sb.Append('</table>')
    $sb.ToString()
}

function New-ChartTableSection {
    # Heading, table on the left, chart on the right - a genuine HTML table (table-layout:fixed +
    # explicit width= attributes on the <td>s), because Outlook desktop's Word rendering engine
    # ignores <style>/media queries entirely: whatever default markup we emit IS what it shows, so
    # the two-column layout has to be the base structure rather than something CSS switches on.
    # Mobile clients (which do honour media queries) collapse this back to one column - chart above
    # table, matching the original reference design - via the ".stack-*" rules under
    # "@media screen and (max-width:600px)".
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$ContentId,
        [Parameter(Mandatory)][string]$TableHtml,
        [string]$ImageAlt = $Title
    )
    @"
<div class="report-section">
<h3 class="section-title" style="margin:20px 0 8px 0;font-size:15px;color:#000000;font-weight:bold;">$Title</h3>
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;table-layout:fixed;border-collapse:collapse;" class="stack-table"><tr>
<td width="48%" valign="top" class="stack-cell stack-cell-table" style="width:48%;padding:0 10px 0 0;">$TableHtml</td>
<td width="50%" valign="top" class="stack-cell stack-cell-chart" style="width:50%;padding:0;">
<img src="cid:$ContentId" alt="$ImageAlt" width="448" class="report-chart" style="max-width:448px;width:100%;height:auto;display:block;margin:0 auto;" />
</td>
</tr></table>
</div>
"@
}

function New-TableOnlySection {
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$TableHtml)
    @"
<div class="report-section">
<h3 class="section-title" style="margin:20px 0 8px 0;font-size:15px;color:#000000;font-weight:bold;">$Title</h3>
$TableHtml
</div>
"@
}

#endregion

#region Request the cost details report from Microsoft.CostManagement
$bearer_token = Get-MiToken -Resource "https://management.azure.com/"
Write-Output "[OK] ARM token retrieved for Microsoft.CostManagement API."

$api = "2026-08-01"
$uri = "https://management.azure.com/subscriptions/{0}/providers/Microsoft.CostManagement/generateCostDetailsReport?api-version=$api" -f $SubscriptionId
$headers = @{
    "Authorization" = "Bearer {0}" -f $bearer_token
}
$body = @{
    metric     = "ActualCost"
    timePeriod = @{ end = $ConsumptionDate.ToString("yyyy-MM-dd"); start = $ConsumptionDate.ToString("yyyy-MM-dd") }   #will be just one day, but can be a range of days
}
$json = $body | ConvertTo-Json -Depth 10

$request = $null
$maxRetries = 3
$retryCount = 0

do {
    try {
        Write-Output "Requesting Microsoft.CostManagement to generateCostDetailsReport: $($ConsumptionDate.ToString('yyyy-MM-dd'))"
        $request = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $json -UseBasicParsing
        $retryCount = $maxRetries  # Exit loop on success
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Output "No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)"
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

if (!($request.RawContent -match "Location: (\S*)")) {
    Write-Output "No valid response after $maxRetries attempts. Exiting script."
    exit 1
}
$locationuri = $Matches[1]
#endregion

#region Poll for the generated report and download the CSV
$reportName = "$($ConsumptionDate.ToString('yyyy-MM-dd'))_report.csv"
$retryCount = 0
$bloburl = $null

do {
    try {
        $jsonresponse = Invoke-WebRequest -Uri $locationuri -Method Get -Headers $headers -UseBasicParsing
        $bloburl = ($jsonresponse.Content | ConvertFrom-Json).manifest.blobs.bloblink
        if ($bloburl) {
            Write-Output "Blob URL retrieved successfully."
            $retryCount = $maxRetries  # Exit loop on success
        }
        else {
            Write-Output "Blob URL not found in the response. Retrying..."
            throw "Blob URL not found"
        }
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Output "No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)"
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

Invoke-WebRequest -Uri $bloburl -Method Get -UseBasicParsing -OutFile "$Env:temp\$reportName"
$csv = Import-Csv "$Env:temp\$reportName" -Encoding UTF8
#endregion

<# Example row schema (for reference):
invoiceId                    :
previousInvoiceId            :
billingAccountId              : e35a..........5f0
billingAccountName            : Bernhard
billingProfileId               : DD...........PGB
billingProfileName             : Bernhard Frank
invoiceSectionId                : 400a..........d60
invoiceSectionName              : Bernhard Frank
resellerName                    :
resellerMpnId                   :
costCenter                      :
billingPeriodEndDate            :
billingPeriodStartDate          :
servicePeriodEndDate            : 10/01/2026
servicePeriodStartDate          : 09/01/2026
date                             : 09/10/2026
serviceFamily                    : Compute
productOrderId                   : 9e8e7ee3-d886-4f18-d3b5-410387f5924d
productOrderName                 : Azure plan
consumedService                  : microsoft.azurestackhci
meterId                          : 79440372-a360-5b70-914f-f9e5adfcdf0c
meterName                        : Standard Trial Fee
meterCategory                    : Azure Local
meterSubCategory                 : Azure Local
meterRegion                      : Global
ProductId                        : DZH318Z0MVPD000L
ProductName                      : Azure Local - Standard
SubscriptionId                   : 80c67.........a2413
subscriptionName                 : AzurePayGo
publisherType                    : Microsoft
publisherId                      :
publisherName                    : Microsoft
resourceGroupName                : rg-azlocal
ResourceId                       : /subscriptions/80c673c.......13/resourcegroups/rg-azlocal/providers/microsoft.azurestackhci/clusters/hcimx
resourceLocation                 : westeurope
location                          : EU West
effectivePrice                   : 0
quantity                         : 32
unitOfMeasure                     : 1/Day
chargeType                        : Usage
billingCurrency                   : EUR
pricingCurrency                   : USD
costInBillingCurrency             : 0
costInPricingCurrency             : 0
costInUsd                         : 0
paygCostInBillingCurrency         : 0
paygCostInUsd                     : 0
exchangeRatePricingToBilling      : 0.858663918942126052
exchangeRateDate                  : 09/01/2026
isAzureCreditEligible             : True
serviceInfo1                      :
serviceInfo2                      :
additionalInfo                    :
tags                              :
PayGPrice                         : 0
frequency                         : UsageBased
term                               :
reservationId                     :
reservationName                   :
pricingModel                      : OnDemand
unitPrice                         : 0
costAllocationRuleName            :
benefitId                         :
benefitName                       :
provider                          : Azure
#>

#region Transform & export the culture-formatted CSV attachment
$reportNameCulture = "$($ConsumptionDate.ToString('yyyy-MM-dd'))_$($CultureInfo)_report.csv"
$transformedUsagePath = "$Env:temp\$reportNameCulture"

$csv | Select-Object `
    @{N = 'date'; E = { [System.DateTime]::Parse($_.date, [CultureInfo]::new("en-us")).ToString("d", $destculture) } }, `
    serviceFamily, consumedService, meterName, meterCategory, meterSubCategory, meterRegion, ProductName, resourceGroupName, `
    @{N = 'ResourceName'; E = { ($_.ResourceId -split '/')[-1] } }, `
    @{N = 'quantity'; E = { ([decimal]$_.quantity).ToString($destculture) } }, `
    @{N = 'paygCostInBillingCurrency'; E = { ([decimal]$_.paygCostInBillingCurrency).ToString($destculture) } }, `
    billingCurrency, unitOfMeasure, `
    @{N = 'unitPrice'; E = { ([decimal]$_.unitPrice).ToString($destculture) } }, `
    @{N = 'exchangeRatePricingToBilling'; E = { ([decimal]$_.exchangeRatePricingToBilling).ToString($destculture) } }, `
    meterId, tags |
Export-Csv $transformedUsagePath -Encoding UTF8 -Delimiter ';' -NoTypeInformation

$totalCost = ($csv | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum
Write-Output "Total cost for $($ConsumptionDate.ToString('yyyy-MM-dd')): $totalCost"

$currency = $csv | Select-Object -First 1 -ExpandProperty billingCurrency
if ([string]::IsNullOrWhiteSpace($currency)) { $currency = "EUR" }
#endregion

#region Cost history table (Azure Storage Table)
$tableName = "myazurecosttable"
$tableEndpoint = "https://$AzureCostStorageAccountName.table.core.windows.net"
Write-Output "Storage account: $AzureCostStorageAccountName"

Write-Output "[INFO] Writing today's cost to table..."
Set-CostTableEntity -TableEndpoint $tableEndpoint -TableName $tableName `
    -PartitionKey $ConsumptionDate.ToString('MMMM', [CultureInfo]::InvariantCulture) -RowKey $ConsumptionDate.ToString('dd', [CultureInfo]::InvariantCulture) `
    -TotalCost $totalCost -Year $ConsumptionDate.Year

$monthToDateCost = Get-MonthToDateCost -TableEndpoint $tableEndpoint -TableName $tableName -AsOfDate $ConsumptionDate
Write-Output "Month-to-date cost through $($ConsumptionDate.ToString('yyyy-MM-dd')): $monthToDateCost"

$last7Days = @()
for ($date = $ConsumptionDate.AddDays(-6); $date -lt $ConsumptionDate; $date = $date.AddDays(1)) {
    $row = Get-CostTableEntity -TableEndpoint $tableEndpoint -TableName $tableName `
        -PartitionKey $date.ToString('MMMM', [CultureInfo]::InvariantCulture) -RowKey $date.ToString('dd', [CultureInfo]::InvariantCulture)
    if ($row) {
        $row | Add-Member -NotePropertyName Date -NotePropertyValue $date -Force
        $last7Days += $row
    }
}
# Today's row is built directly from the $totalCost already calculated above, instead of being
# read back from the table - so the chart/table always include today even if that read-after-write
# were ever to miss (e.g. a transient error on the extra round trip), rather than silently omitting it.
$last7Days += [PSCustomObject]@{
    PartitionKey = $ConsumptionDate.ToString('MMMM', [CultureInfo]::InvariantCulture)
    RowKey       = $ConsumptionDate.ToString('dd', [CultureInfo]::InvariantCulture)
    TotalCost    = ([decimal]$totalCost).ToString("F7", [CultureInfo]::InvariantCulture)
    Year         = $ConsumptionDate.Year
    Date         = $ConsumptionDate
}
$last7Days | Format-Table Date, PartitionKey, RowKey, Year, TotalCost
#endregion

#region Cost breakdowns (category / resource group / region)
$costPerCatResult = Get-CostBreakdown -Csv $csv -GroupProperty 'meterCategory' -TotalCost $totalCost
$costsPerRGResult = Get-CostBreakdown -Csv $csv -GroupProperty 'resourceGroupName' -TotalCost $totalCost
$costsPerRegionResult = Get-CostBreakdown -Csv $csv -GroupProperty 'location' -TotalCost $totalCost

Write-Output "========================"
Write-Output "Total costs per category"
$costPerCatResult | Format-Table -AutoSize

Write-Output "========================"
Write-Output "Costs per resource group"
$costsPerRGResult | Format-Table -AutoSize

Write-Output "========================"
Write-Output "Costs per region"
$costsPerRegionResult | Format-Table -AutoSize

Write-Output "========================"
Write-Output "Top 10 consumers (paygCostInBillingCurrency)"
$csv | Sort-Object 'paygCostInBillingCurrency' -Descending | Select-Object -First 10 |
    Format-Table @{N = 'ResourceName'; E = { ($_.ResourceId -split '/')[-1] } }, 'paygCostInBillingCurrency', MeterName, meterCategory -AutoSize

Write-Output "========================"
Write-Output "Top 3 consumers per category"
$csv | Group-Object -Property meterCategory | ForEach-Object { $_.Group | Sort-Object 'paygCostInBillingCurrency' -Descending | Select-Object -First 3 } |
    Format-Table @{N = 'ResourceName'; E = { ($_.ResourceId -split '/')[-1] } }, 'paygCostInBillingCurrency', meterCategory -AutoSize
#endregion

#region Generate charts
Write-Output "[INFO] Generating charts..."

$costHistoryChartConfig = @{
    type    = "bar"
    data    = @{
        labels   = @($last7Days | ForEach-Object { $_.Date.ToString("ddd dd.MM.", $destculture) })
        datasets = @(@{
                label           = "Total Cost"
                # Rounded to 4 digits, not 2: on a low-usage subscription, daily costs can be
                # sub-cent, and rounding to 2 decimals collapses them to a "0" bar/label that looks
                # like the day is missing from the chart entirely.
                data            = @($last7Days | ForEach-Object { [Math]::Round([decimal]$_.TotalCost, 4) })
                backgroundColor = "#0078D4"
            })
    }
    options = @{
        legend  = @{ display = $false }
        layout  = @{ padding = @{ top = 24; bottom = 6; left = 6; right = 16 } }
        scales  = @{
            xAxes = @(@{ gridLines = @{ display = $false } })
            yAxes = @(@{ gridLines = @{ color = "#E6E6E6" }; ticks = @{ beginAtZero = $true }; scaleLabel = @{ display = $true; labelString = $currency } })
        }
        plugins = @{
            datalabels = @{
                color  = "#333333"
                anchor = "end"
                align  = "end"
                font   = @{ size = 11 }
            }
        }
    }
}

$costHistoryChartBytes = Get-QuickChartPng -ChartConfig $costHistoryChartConfig -Width 640 -Height 240
$costsPerCatChartBytes = Get-QuickChartPng -ChartConfig (New-PercentageDoughnutChartConfig -Result $costPerCatResult) -Width 480 -Height 320
$rgChartHeight = [Math]::Max(160, 50 + (50 * $costsPerRGResult.Count))
$costsPerRGChartBytes = Get-QuickChartPng -ChartConfig (New-HorizontalCostShareBarChartConfig -Result $costsPerRGResult -Currency $currency -TotalCost $totalCost) -Width 460 -Height $rgChartHeight
$costsPerRegionChartBytes = Get-QuickChartPng -ChartConfig (New-PercentageDoughnutChartConfig -Result $costsPerRegionResult) -Width 480 -Height 320
#endregion

#region Build table HTML fragments
$last7Sum = $last7Days | ForEach-Object { [decimal]$_.TotalCost } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
if (-not $last7Sum) { $last7Sum = 1 }   # guard against a divide-by-zero if history is empty

$historyRows = $last7Days | ForEach-Object {
    [PSCustomObject]@{
        Label      = $_.Date.ToString("ddd dd.MM.", $destculture)
        Amount     = Format-Money -Value ([decimal]$_.TotalCost) -Culture $destculture -Currency $currency -DecimalDigits 4
        Percentage = [double]([decimal]$_.TotalCost * 100 / $last7Sum)
    }
}
$historyTableHtml = New-ShareTableHtml -Rows $historyRows -FirstColumnHeader "Day" -AmountHeader "Cost" -ShareHeader "Share of 7 days"

$catTableHtml = New-ShareTableHtml -FirstColumnHeader "Category" -Rows ($costPerCatResult | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Name; Amount = Format-Money -Value $_.Sum -Culture $destculture -Currency $currency; Percentage = [double]$_.Percentage }
    })

$rgTableHtml = New-ShareTableHtml -FirstColumnHeader "Resource Group" -Rows ($costsPerRGResult | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Name; Amount = Format-Money -Value $_.Sum -Culture $destculture -Currency $currency; Percentage = [double]$_.Percentage }
    })

$regionTableHtml = New-ShareTableHtml -FirstColumnHeader "Region" -Rows ($costsPerRegionResult | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Name; Amount = Format-Money -Value $_.Sum -Culture $destculture -Currency $currency; Percentage = [double]$_.Percentage }
    })

$top10Rows = $csv | Sort-Object 'paygCostInBillingCurrency' -Descending | Select-Object -First 10 | ForEach-Object {
    , @(
        ($_.ResourceId -split '/')[-1],
        (Format-Money -Value ([decimal]$_.paygCostInBillingCurrency) -Culture $destculture -Currency $currency -DecimalDigits 4),
        $_.MeterName,
        $_.meterCategory
    )
}
$top10TableHtml = New-PlainTableHtml -Headers @("Resource", "Cost", "Meter", "Category") -Alignments @("left", "right", "left", "left") -Rows $top10Rows -ColumnClasses @("", "", "", "hide-mobile")
#endregion

#region Build the HTML email body
$kpiHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:16px;" class="kpi-table"><tr>
<td width="50%" style="width:50%;background:#0078D4;padding:12px;text-align:center;" class="kpi-tile">
  <div style="font-size:11px;color:#ffffff;" class="kpi-label">Cost Yesterday</div>
  <div style="font-size:20px;font-weight:bold;color:#ffffff;" class="kpi-value">$(Format-Money -Value $totalCost -Culture $destculture -Currency $currency)</div>
</td>
<td width="50%" style="width:50%;background:#004E8C;padding:12px;text-align:center;" class="kpi-tile">
  <div style="font-size:11px;color:#ffffff;" class="kpi-label">Month to Date</div>
  <div style="font-size:20px;font-weight:bold;color:#ffffff;" class="kpi-value">$(Format-Money -Value $monthToDateCost -Culture $destculture -Currency $currency)</div>
</td>
</tr></table>
"@

$htmlBody = @"
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<!-- "Mobile background is black, even the text" is mobile mail apps' automatic dark-mode recoloring:
     it inverted the background but couldn't repaint text color it can't see, because most of this
     email's text relies on inherited color from <body> or on the <style> block below - which
     Outlook's mobile app in particular is known to ignore for this purpose - rather than an
     explicit per-element color. The fix below is two-part: these meta tags ("only light" is the
     correct token order - "light only" is invalid and silently ignored) opt out where honored
     (Apple/iOS Mail, Yahoo), and every heading/table cell in the body further down now also
     carries its color as an explicit inline style, so Outlook's engine has something concrete to
     pair with the background it flips instead of leaving inherited/default black text behind. -->
<meta name="color-scheme" content="only light" />
<meta name="supported-color-schemes" content="light" />
<title>Azure Cost Report</title>
<style>
    body { margin:0; padding:0; background:#FFFFFF; font-family:'Segoe UI', Arial, sans-serif; color:#222222; }
    /* Gmail's app doesn't honour the color-scheme meta tags above; instead it recolors elements
       itself and tags them with data-ogsc/data-ogsb, which lets this rule claw the colors back. */
    [data-ogsc] body, [data-ogsc] .email-container, [data-ogsc] .report-box, [data-ogsb] { background:#FFFFFF !important; color:#222222 !important; }
    /* Widened from 680px so the .stack-table chart column has enough room to actually render at
       its larger size - bumping the <img>'s own max-width alone would do nothing once it's capped
       by a column that's still only, say, 46% of a 680px box. */
    .email-container { width:100%; max-width:960px; margin:0 auto; padding:16px; box-sizing:border-box; background:#FFFFFF; }
    .report-box { border:1px solid #E0E0E0; padding:15px; background:#FFFFFF; }
    h2.report-title { margin:0 0 4px 0; font-size:22px; color:#0078D4; font-weight:bold; }
    p.report-subtitle { margin:0 0 14px 0; font-size:12px; color:#666666; }
    h3.section-title { margin:20px 0 8px 0; font-size:15px; color:#000000; font-weight:bold; }
    p.report-footer { margin-top:18px; margin-bottom:0; font-size:11px; color:#999999; }
    /* The two-column table/chart layout above (.stack-table) is the default markup so Outlook
       desktop's Word engine - which ignores <style>/media queries entirely - always renders it as
       two columns. Everything below is what CAN respond to media queries: real browsers, mobile
       mail apps (Gmail, Apple Mail, Outlook mobile), and print. Selectors need !important because
       the email body itself is all inline styles, which normally win over a stylesheet rule but
       not over an !important one. */

    /* Phones/narrow viewports: collapse the two-column tables back to one column (chart above
       table, matching the original reference design), and drop the Top 10 table's Category column
       so it doesn't get squeezed unreadably narrow. (The background is plain white everywhere now -
       not just here - because many mobile mail apps strip <style> blocks, or even the <body> tag
       itself, so a media query targeting "body" can silently never fire; white is instead baked in
       as the unconditional default above, plus inline on <body>/.email-container as a second layer
       of safety for clients that also ignore the <style> block entirely.) */
    @media screen and (max-width:600px) {
        /* .stack-table's row must become a flex container (not block) for "order" below to have
           any effect - so "tr" is deliberately left out of this block-ifying rule; a selector that
           also matched "tr" here would be more specific than ".stack-table tr" below and, since
           both are !important, would win on specificity and silently defeat the flex/order rule. */
        .stack-table, .stack-table > tbody { display:block !important; width:100% !important; }
        .stack-table tr { display:flex !important; flex-direction:column !important; width:100% !important; }
        .stack-cell { display:block !important; width:100% !important; padding:0 !important; }
        .stack-cell-chart { order:1 !important; margin:0 0 10px 0 !important; }
        .stack-cell-table { order:2 !important; }
        .report-chart { max-width:448px !important; }
        .hide-mobile { display:none !important; }
    }

    /* Compact everything down to fit a single A4 page when printed (e.g. Ctrl+P from a browser). */
    @media print {
        @page { size: A4; margin: 10mm; }
        body, .email-container, .report-box { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        .email-container { max-width:100% !important; width:100% !important; padding:0 !important; }
        .report-box { padding:6px !important; border:none !important; }
        h2.report-title { font-size:15px !important; margin:0 0 2px 0 !important; }
        p.report-subtitle { font-size:7px !important; margin:0 0 4px 0 !important; }
        h3.section-title { font-size:9px !important; margin:4px 0 2px 0 !important; }
        p.report-footer { font-size:6px !important; margin-top:3px !important; }
        .kpi-table { margin-bottom:4px !important; }
        .kpi-tile { padding:3px !important; }
        .kpi-label { font-size:7px !important; }
        .kpi-value { font-size:11px !important; }
        .report-section { page-break-inside: avoid; }
        .stack-cell-table { width:62% !important; }
        .stack-cell-chart { width:34% !important; }
        .report-chart { max-width:204px !important; margin:2px auto 3px auto !important; }
        table td { font-size:8px !important; padding:1.5px 4px !important; }
        /* These two carry explicit inline font-size/height, so the generic "table td" rule above
           can't reach them - inherited font-size never overrides an element's own inline value. */
        .share-bar td { height:3px !important; line-height:3px !important; }
        .share-pct { font-size:7px !important; margin-top:0 !important; line-height:8px !important; }
    }
</style>
</head>
<body bgcolor="#FFFFFF" style="margin:0;padding:0;background:#FFFFFF;font-family:'Segoe UI', Arial, sans-serif;color:#222222;">
<div class="email-container" style="width:100%;max-width:960px;margin:0 auto;padding:16px;box-sizing:border-box;background:#FFFFFF;">
  <div class="report-box" style="border:1px solid #E0E0E0;padding:15px;background:#FFFFFF;">
    <h2 class="report-title" style="margin:0 0 4px 0;font-size:22px;color:#0078D4;font-weight:bold;">Azure Cost Report</h2>
    <p class="report-subtitle" style="margin:0 0 14px 0;font-size:12px;color:#666666;">As of $($ConsumptionDate.ToString("D", $destculture)) &middot; Subscription $SubscriptionId</p>
    $kpiHtml
    $(New-ChartTableSection -Title "Cost History (last 7 days)" -ContentId "costHistoryChart" -TableHtml $historyTableHtml)
    $(New-ChartTableSection -Title "Costs per Category" -ContentId "costsPerCatChart" -TableHtml $catTableHtml)
    $(New-TableOnlySection -Title "Top 10 Consumers" -TableHtml $top10TableHtml)
    $(New-ChartTableSection -Title "Costs per Resource Group" -ContentId "costsPerRGChart" -TableHtml $rgTableHtml)
    $(New-ChartTableSection -Title "Costs per Region" -ContentId "costsPerRegionChart" -TableHtml $regionTableHtml)
    <p class="report-footer" style="margin-top:18px;margin-bottom:0;font-size:11px;color:#999999;">Generated automatically &middot; $($ConsumptionDate.ToString('yyyy-MM-dd'))<br />For a single-page, print-friendly version, open the attached "AzureCostReport_Print.html" in a browser and print from there - Outlook's own print view does not apply the compact layout.</p>
  </div>
</div>
</body>
</html>
"@

# A standalone copy for printing: Outlook's own Print command uses Word's rendering engine, which
# doesn't reliably apply the @media print rules above, so the compact single-page layout only shows
# up if the report is opened in an actual browser and printed from there. Outlook can't resolve
# cid: image references outside of the email itself, so this copy embeds the charts as data: URIs
# (same technique Preview-DailyCostReport.ps1 uses) to make it a fully self-contained file.
$printHtmlBody = $htmlBody `
    -replace 'cid:costHistoryChart', "data:image/png;base64,$([Convert]::ToBase64String($costHistoryChartBytes))" `
    -replace 'cid:costsPerCatChart', "data:image/png;base64,$([Convert]::ToBase64String($costsPerCatChartBytes))" `
    -replace 'cid:costsPerRGChart', "data:image/png;base64,$([Convert]::ToBase64String($costsPerRGChartBytes))" `
    -replace 'cid:costsPerRegionChart', "data:image/png;base64,$([Convert]::ToBase64String($costsPerRegionChartBytes))"
#endregion

#region Send email via Azure Communication Services
Write-Output "[INFO] Get token for communication service and send email..."
$acsToken = Get-MiToken -Resource "https://communication.azure.com/"

$csvBytes = [byte[]](0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::UTF8.GetBytes((Get-Content $transformedUsagePath -Raw))

$mailBody = @{
    senderAddress = $SenderAddress
    recipients    = @{
        to = @(@{ address = $RecipientEmail })
    }
    content       = @{
        subject = "Azure Cost Report $($ConsumptionDate.ToString('yyyy-MM-dd')): Yesterday $(Format-Money -Value $totalCost -Culture $destculture -Currency $currency) | MTD $(Format-Money -Value $monthToDateCost -Culture $destculture -Currency $currency)"
        html    = $htmlBody
    }
    attachments   = @(
        @{
            name            = $reportNameCulture
            contentType     = "text/csv"
            contentInBase64 = [Convert]::ToBase64String($csvBytes)
        },
        @{
            name            = "costHistoryChart.png"
            contentType     = "image/png"
            contentInBase64 = [Convert]::ToBase64String($costHistoryChartBytes)
            contentId       = "costHistoryChart"
        },
        @{
            name            = "costsPerCatChart.png"
            contentType     = "image/png"
            contentInBase64 = [Convert]::ToBase64String($costsPerCatChartBytes)
            contentId       = "costsPerCatChart"
        },
        @{
            name            = "costsPerRGChart.png"
            contentType     = "image/png"
            contentInBase64 = [Convert]::ToBase64String($costsPerRGChartBytes)
            contentId       = "costsPerRGChart"
        },
        @{
            name            = "costsPerRegionChart.png"
            contentType     = "image/png"
            contentInBase64 = [Convert]::ToBase64String($costsPerRegionChartBytes)
            contentId       = "costsPerRegionChart"
        },
        @{
            name            = "AzureCostReport_Print.html"
            contentType     = "text/html"
            contentInBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($printHtmlBody))
        }
    )
}

$sendUri = "{0}/emails:send?api-version=2025-09-01" -f $AcsEndpoint.TrimEnd("/")
$sendHeaders = @{
    "Authorization" = "Bearer {0}" -f $acsToken
    "Content-Type"  = "application/json"
}

$sendResult = Invoke-RestMethod -Uri $sendUri -Method Post -Headers $sendHeaders `
    -Body ($mailBody | ConvertTo-Json -Depth 10) -UseBasicParsing

Write-Output "[OK] Email queued. Operation-Id: $($sendResult.id), Status: $($sendResult.status)"
#endregion
