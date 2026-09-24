param (
    [parameter(Mandatory = $false,
        HelpMessage = "Enter a en-us formatted date e.g. '12/30/2019'")]
    [String]$myDate
)

try {
    $ConsumptionDate = [dateTime]::Parse($myDate)
}
catch {
    $ConsumptionDate = [dateTime]::Today.AddDays(-1)      #default to yesterday
}
Write-Output "Get consumption of $($ConsumptionDate.ToString("dd'/'MM'/'yyyy"))"

# Get Azure Automation Variables (SubscriptionId, ACS Endpoint, Sender/Recipient, Culture, LookbackDays)
[string]$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"
[string]$AcsEndpoint = Get-AutomationVariable -Name "AcsEndpoint"
[string]$SenderAddress = Get-AutomationVariable -Name "SenderAddress"
[string]$RecipientEmail = Get-AutomationVariable -Name "RecipientEmail"
[string]$CultureInfo = Get-AutomationVariable -Name "CultureInfo" 
[string]$AzureCostStorageAccountID = Get-AutomationVariable -Name "AzureCostStorageAccountID" 

try
{
    "...attachments destination culture: $CultureInfo"
    $destculture = [CultureInfo]::new("$CultureInfo")
}
catch [System.Globalization.CultureNotFoundException]
{
    "$CultureInfo did not work ... using en-US instead."
    $destculture = [CultureInfo]::new("en-US")
}

function Get-MiToken {
    # Holt ein Token der System-assigned Managed Identity ueber den
    # Automation-Sandbox-Identity-Endpunkt.
    param([Parameter(Mandatory = $true)][string]$Resource)
    if (-not $env:IDENTITY_ENDPOINT -or -not $env:IDENTITY_HEADER) {
        throw "IDENTITY_ENDPOINT nicht verfuegbar. Laeuft das Runbook in Azure Automation mit aktivierter Managed Identity?"
    }
    $uri = "{0}?resource={1}" -f $env:IDENTITY_ENDPOINT, [uri]::EscapeDataString($Resource)
    $headers = @{
        "X-IDENTITY-HEADER" = $env:IDENTITY_HEADER
        "Metadata"          = "True"
    }
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -UseBasicParsing
    return $response.access_token
}

$bearer_token = Get-MiToken -Resource "https://management.azure.com/"
Write-Output "[OK] ARM-Token retrieved successfully for Microsoft.CostManagement API."

#$ConsumptionDate = [System.DateTime]::Parse("2026-09-10", [CultureInfo]::new("en-us"))

$api = "2026-08-01" #"2025-03-01"
$SubscriptionId = "80c673c5-92dc-4815-8087-df3020fa2413"    
$uri = "https://management.azure.com/subscriptions/{0}/providers/Microsoft.CostManagement/generateCostDetailsReport?api-version=$api" -f $SubscriptionId
$headers = @{
    "Authorization" = "Bearer {0}" -f $bearer_token
}
$Body = @{
    metric     = "ActualCost"
    timePeriod = @{ end = $ConsumptionDate.ToString("yyyy-MM-dd"); start = $ConsumptionDate.ToString("yyyy-MM-dd") }   #will be just one day, but can be a range of days
}
$json = $Body | ConvertTo-Json -Depth 10

$request = $null
$maxRetries = 3
$retryCount = 0

do {
    try {
        Write-Output ("Requesting Microsoft.CostManagement to generateCostDetailsReport: $($ConsumptionDate.ToString('yyyy-MM-dd'))")
        $request = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $json -UseBasicParsing
        $retryCount = $maxRetries  # Exit loop on success
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
             Write-Output ("No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)")
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

if (!($request.RawContent -match "Location: (\S*)" )) {
     Write-Output ("No valid response after $maxRetries attempts. Exiting script.")
    exit 1
}

## match found, proceed with the next steps
$locationuri = $Matches[1]
$reportName = "$($ConsumptionDate.ToString('yyyy-MM-dd'))_report.csv"
$jsonresponse = $null
$retryCount = 0
$bloburl = $null

do {
    try {
        $jsonresponse = Invoke-WebRequest -Uri $locationuri -Method Get -Headers $headers -UseBasicParsing
        $bloburl = ($jsonresponse.Content | ConvertFrom-Json ).manifest.blobs.bloblink
        if ($bloburl) {
             Write-Output ("Blob URL retrieved successfully: $bloburl")
        } else {
             Write-Output ("Blob URL not found in the response. Retrying...")
            throw "Blob URL not found"
        }
        $retryCount = $maxRetries  # Exit loop on success
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
             Write-Output ("No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)")
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

invoke-webrequest -Uri $bloburl -Method Get -UseBasicParsing -OutFile "$Env:temp\$reportName"
$csv = import-csv "$Env:temp\$reportName" -Encoding UTF8
$reportNameCulture = "$($ConsumptionDate.ToString('yyyy-MM-dd'))_$($CultureInfo)_report.csv"
$transformedUsagePath = "$Env:temp\$reportNameCulture"
# invoiceId,previousInvoiceId,billingAccountId,billingAccountName,billingProfileId,billingProfileName,invoiceSectionId,invoiceSectionName,resellerName,resellerMpnId,costCenter,billingPeriodEndDate,billingPeriodStartDate,servicePeriodEndDate,servicePeriodStartDate,date,serviceFamily,productOrderId,productOrderName,consumedService,meterId,meterName,meterCategory,meterSubCategory,meterRegion,ProductId,ProductName,SubscriptionId,subscriptionName,publisherType,publisherId,publisherName,resourceGroupName,ResourceId,resourceLocation,location,effectiv
<#
invoiceId                    : 
previousInvoiceId            : 
billingAccountId             : e35a..........5f0
billingAccountName           : Bernhard
billingProfileId             : DD...........PGB
billingProfileName           : Bernhard Frank
invoiceSectionId             : 400a..........d60
invoiceSectionName           : Bernhard Frank
resellerName                 : 
resellerMpnId                : 
costCenter                   : 
billingPeriodEndDate         : 
billingPeriodStartDate       : 
servicePeriodEndDate         : 10/01/2026
servicePeriodStartDate       : 09/01/2026
date                         : 09/10/2026
serviceFamily                : Compute
productOrderId               : 9e8e7ee3-d886-4f18-d3b5-410387f5924d
productOrderName             : Azure plan
consumedService              : microsoft.azurestackhci
meterId                      : 79440372-a360-5b70-914f-f9e5adfcdf0c
meterName                    : Standard Trial Fee
meterCategory                : Azure Local
meterSubCategory             : Azure Local
meterRegion                  : Global
ProductId                    : DZH318Z0MVPD000L
ProductName                  : Azure Local - Standard
SubscriptionId               : 80c67.........a2413
subscriptionName             : AzurePayGo
publisherType                : Microsoft
publisherId                  : 
publisherName                : Microsoft
resourceGroupName            : rg-azlocal
ResourceId                   : /subscriptions/80c673c.......13/resourcegroups/rg-azlocal/providers/microsoft.azurestackhci/clusters/hcimx
resourceLocation             : westeurope
location                     : EU West
effectivePrice               : 0
quantity                     : 32
unitOfMeasure                : 1/Day
chargeType                   : Usage
billingCurrency              : EUR
pricingCurrency              : USD
costInBillingCurrency        : 0
costInPricingCurrency        : 0
costInUsd                    : 0
paygCostInBillingCurrency    : 0
paygCostInUsd                : 0
exchangeRatePricingToBilling : 0.858663918942126052
exchangeRateDate             : 09/01/2026
isAzureCreditEligible        : True
serviceInfo1                 : 
serviceInfo2                 : 
additionalInfo               : 
tags                         : 
PayGPrice                    : 0
frequency                    : UsageBased
term                         : 
reservationId                : 
reservationName              : 
pricingModel                 : OnDemand
unitPrice                    : 0
costAllocationRuleName       : 
benefitId                    : 
benefitName                  : 
provider                     : Azure
#>

$csv | Select-object @{N = 'date'; E = { "{0}" -f [System.DateTime]::Parse($_.date, [CultureInfo]::new("en-us")).ToString("d", $destculture) } }, serviceFamily, consumedService, meterName, meterCategory, meterSubCategory, meterRegion, ProductName, resourceGroupName, @{N = 'ResourceName'; E = { ($_.ResourceId.Split('/')) | select -Last 1 } }, @{N = 'quantity'; E = { $([decimal]$_.quantity).ToString($destculture) } }, @{N = 'paygCostInBillingCurrency'; E = { $([decimal]$_.paygCostInBillingCurrency).ToString($destculture) } }, billingCurrency, unitOfMeasure, @{N = 'unitPrice'; E = { $([decimal]$_.unitPrice).ToString($destculture) } }, @{N = 'exchangeRatePricingToBilling'; E = { $([decimal]$_.exchangeRatePricingToBilling).ToString($destculture) } }, meterId, tags | Export-Csv "$Env:temp\$reportNameCulture" -Encoding UTF8 -Delimiter ';' -NoTypeInformation

$totalCost = $($csv | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum
$storageAccount = $AzureCostStorageAccountID | Split-Path -Leaf

$storageAccount

$resourceGroupName = $AzureCostStorageAccountID.Split('/')[3]

$resourceGroupName

$tablename = "myazurecosttable"
#region get history data from table
$sa = Get-AzStorageAccount -Name $storageAccount -ResourceGroupName $resourceGroupName        
$ctx = $sa.Context
$cloudTable = (Get-AzStorageTable -Name $tableName -Context $ctx).CloudTable

$cloudTable
Write-Output "[INFO] Get token for communication service and send email..."

#update or new
try {
    $entry = Get-AzTableRow -Table $cloudTable -PartitionKey $ConsumptionDate.ToString('MMMM') -rowKey "$($ConsumptionDate.ToString('dd'))"
    $entry.TotalCost = "{0:N7}" -f $totalCost
    $entry.Year = $ConsumptionDate.Year
    $entry | Update-AzTableRow -table $cloudTable
}
catch {
    Add-AzTableRow -table $cloudTable -partitionKey $ConsumptionDate.ToString('MMMM') -rowKey "$($ConsumptionDate.ToString('dd'))" -property @{"TotalCost" = $("{0:N7}" -f $totalCost); "Year" = $ConsumptionDate.Year }
}

Get-AzTableRow -Table $cloudTable -PartitionKey $ConsumptionDate.ToString('MMMM') -rowKey "$($ConsumptionDate.ToString('dd'))"
#Get last 7 days
$last7Days = @()
for ($date = $ConsumptionDate.AddDays(-6); $date -le $ConsumptionDate; $date += [System.timespan]::new(1, 0, 0, 0)) { 
    $last7Days += Get-AzTableRow -Table $cloudTable -PartitionKey $date.ToString('MMMM') -rowKey "$($date.ToString('dd'))"
}

$last7Days | ft RowKey, PartitionKey, Year, TotalCost


#region total costs per category
$costPerCat = $csv | Group-Object -Property meterCategory | % { $Sum = ($_.Group | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum; $myobj = [PSCustomObject]@{Name = "$($_.Name)"; Count = $($_.Group.Count); Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }; $myobj }
$costPerCatResult = @()
$costPerCatResult += ($costPerCat | Where-Object Percentage -GT 3 | Sort-Object Percentage -Descending)#.GetEnumerator()

$Sum = (($costPerCat | Where-Object Percentage -le 3) | Measure-Object Sum -Sum).Sum
$costPerCatResult += [PSCustomObject]@{Name = "other"; Count = (($costPerCat | Where-Object Percentage -le 3) | Measure-Object Count -Sum).Sum; Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }
"========================"
"Total costs per category"
$costPerCatResult | Select-Object Name, Count, @{N = 'Sum'; E = { "{0:N2}" -f $_.Sum } }, @{N = 'Percentage'; E = { "{0:N2}%" -f $_.Percentage } } | ft -AutoSize
#endregion 

#region Top 10 consumers 
"========================"
"'paygCostInBillingCurrency'"
$csv | Sort-Object 'paygCostInBillingCurrency' -Descending | Select-Object -First 10 | ft @{N = 'ResourceName'; E = { ($_.ResourceId.Split('/')) | select -Last 1 } }, 'paygCostInBillingCurrency', MeterName, meterCategory | ft -AutoSize
#endregion 

#region Costs per RG
$costsPerRG = $csv | Group-Object -Property resourceGroupName | % { $Sum = ($_.Group | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum; $myobj = [PSCustomObject]@{Name = "$($_.Name)"; Count = "$($_.Count)"; Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }; $myobj }
$costsPerRGResult = @()
$costsPerRGResult += $costsPerRG | Where-Object Percentage -GT 3 | Sort-Object Percentage -Descending

$Sum = (($costsPerRG | Where-Object Percentage -le 3) | Measure-Object Sum -Sum).Sum
$costsPerRGResult += [PSCustomObject]@{Name = "other"; Count = (($costsPerRG | Where-Object Percentage -le 3) | Measure-Object Count -Sum).Sum; Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }
"========================"
"Costs per RG"
$costsPerRGResult | ft -AutoSize
#endregion  

#region Costs per Region
$costsPerRegion = $csv | Group-Object -Property location | % { $Sum = ($_.Group | Measure-Object 'paygCostInBillingCurrency' -Sum).Sum; $myobj = [PSCustomObject]@{Name = "$($_.Name)"; Count = "$($_.Count)"; Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }; $myobj }
$costsPerRegionResult = @()
$costsPerRegionResult += $costsPerRegion | Where-Object Percentage -GT 3 | Sort-Object Percentage -Descending

$Sum = (($costsPerRegion | Where-Object Percentage -le 3) | Measure-Object Sum -Sum).Sum
$costsPerRegionResult += [PSCustomObject]@{Name = "other"; Count = (($costsPerRegion | Where-Object Percentage -le 3) | Measure-Object Count -Sum).Sum; Sum = $Sum; Percentage = [Math]::Round([decimal]((100 * $Sum) / $totalCost), 2) }
"========================"
"Costs per Region"
$costsPerRegionResult | ft -AutoSize
#endregion

#region Top 3 consumers per category
$top3ConsumersPerCat = $csv | Group-Object -Property meterCategory | % { $_.Group | sort-object 'paygCostInBillingCurrency' -Descending | Select-Object -First 3 }
"========================"
"Top 3 consumers per category"
$top3ConsumersPerCat | ft @{N = 'ResourceName'; E = { ($_.ResourceId.Split('/')) | select -Last 1 } }, 'paygCostInBillingCurrency', meterCategory -AutoSize
#endregion 



$htmlBody = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Your Azure Daily Usage Email</title>
</head>
<body>
<p><h2>Hello,</h2></p>
<p>This is your daily usage report of <b>$($ConsumptionDate.ToString("d",$destculture))</b> for subscription: <b>$($AzureContext.Subscription)</b>.</p>
<p>(cultureinfo: <b>$CultureInfo.</b>)</p>
<p>hope you'll find it useful.</p>
"@
$htmlBody += "<p><h3>Costs History:</h3>"
$htmlBody += "<table style=""width:auto; height: auto;""><tr><td><img src='cid:costHistoryChart'></td><td>"
$htmlBody += $($last7Days | Select-Object @{N = 'Day'; E = { "{0}" -f $_.RowKey } }, @{N = 'Month'; E = { "{0}" -f $_.PartitionKey } }, Year, TotalCost | ConvertTo-Html -Property Day, Month, Year, TotalCost -Fragment)
$htmlBody += "</td></tr></table></p>"
$htmlBody += "<p><h3>Costs Per Category:</h3>"
$htmlBody += "<table style=""width:auto; height: auto;""><tr><td><img src='cid:costsPerCatChart'></td><td>"
$htmlBody += $($costPerCatResult | Select-Object Name, Count, @{N = 'Sum'; E = { "{0:N2}" -f $_.Sum } }, @{N = 'Percentage'; E = { "{0:N2}%" -f $_.Percentage } } | ConvertTo-Html -Property Name, Count, Sum, Percentage -Fragment)
$htmlBody += "</td></tr></table></p>"
$htmlBody += "<p><h3>Top 10 Consumers:</h3>"
$htmlBody += $($csv | Sort-Object 'paygCostInBillingCurrency' -Descending | Select-Object -First 10 | ConvertTo-Html -Property @{L = 'ResourceName'; E = { $($_.ResourceName -replace "(.{20})(.*)", '$1...') } }, @{L = 'paygCostInBillingCurrency'; E = { $("{0:N2}" -f $($_.'paygCostInBillingCurrency')) } }, MeterName, MeterCategory -Fragment)
$htmlBody += "</p>"
$htmlBody += "<p><h3>Costs per RG:</h3>"
$htmlBody += "<table style=""width:auto; height: auto;""><tr><td><img src='cid:costsPerRGChart'></td><td>"
$htmlBody += $($costsPerRGResult | Select-Object Name, Count, @{N = 'Sum'; E = { "{0:N2}" -f $_.Sum } }, @{N = 'Percentage'; E = { "{0:N2}%" -f $_.Percentage } } | ConvertTo-Html -Property Name, Count, Sum, Percentage -Fragment)
$htmlBody += "</td></tr></table></p>"
$htmlBody += "<p><h3>Costs Per Region:</h3>"
$htmlBody += "<table style=""width:auto; height: auto;""><tr><td><img src='cid:costsPerRegionChart'></td><td>"
$htmlBody += $($costsPerRegionResult | Select-Object Name, Count, @{N = 'Sum'; E = { "{0:N2}" -f $_.Sum } }, @{N = 'Percentage'; E = { "{0:N2}%" -f $_.Percentage } } | ConvertTo-Html -Property Name, Count, Sum, Percentage -Fragment)
$htmlBody += "</td></tr></table></p>"
$htmlBody += "</body></html>"


# You will need to authenticate before running this command.
# Learn more at https://docs.microsoft.com/rest/api/communication/authentication

# =========================================================================
# 6. VERSAND VIA AZURE COMMUNICATION SERVICES
# =========================================================================
$ErrorActionPreference = "Stop"

Write-Output "[INFO] Get token for communication service and send email..."
$acsToken = Get-MiToken -Resource "https://communication.azure.com/"

$csvBytes = [byte[]](0xEF, 0xBB, 0xBF) + [System.Text.Encoding]::UTF8.GetBytes($(get-content $transformedUsagePath | Out-String))

$mailBody = @{
    senderAddress = $SenderAddress
    recipients    = @{
        to = @(@{ address = $RecipientEmail })
    }
    content       = @{
        subject= "Your Daily Azure Usage Report for $($ConsumptionDate.ToString("d",$destculture))"
        html = $htmlBody
    }
    attachments   = @(
        @{
            name            = $reportNameCulture
            contentType     = "text/csv"
            contentInBase64 = [Convert]::ToBase64String($csvBytes)
        }
    )
}

$mailBody

$sendUri = "{0}/emails:send?api-version=2023-03-31" -f $AcsEndpoint.TrimEnd("/")

$sendUri

$sendHeaders = @{
    "Authorization" = "Bearer {0}" -f $acsToken
    "Content-Type"  = "application/json"
}

$sendHeaders

$sendResult = Invoke-RestMethod -Uri $sendUri -Method Post -Headers $sendHeaders `
    -Body ($mailBody | ConvertTo-Json -Depth 10) -UseBasicParsing

Write-Output ("[OK]  Email queued. Operation-Id: {0}, Status: {1}" -f $sendResult.id, $sendResult.status)
