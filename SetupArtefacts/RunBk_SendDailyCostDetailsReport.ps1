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


# Prevent inheriting any existing AzContext
Disable-AzContextAutosave -Scope Process
# Connect using the system-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity).Context
# Set subscription context
Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext
# Example: List all resource groups

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
$Matches.Clear()

do {
    try {
        Write-Host "Requesting Microsoft.CostManagement to generateCostDetailsReport: $ConsumptionDate.ToString("yyyy-MM-dd")"
        $request = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $json -UseBasicParsing
        $retryCount = $maxRetries  # Exit loop on success
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Host "No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)"
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

if (!($request.RawContent -match "Location: (\S*)" )) {
    Write-Host "No valid response after $maxRetries attempts. Exiting script."
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
            Write-Host "Blob URL retrieved successfully: $bloburl"
        } else {
            Write-Host "Blob URL not found in the response. Retrying..."
            throw "Blob URL not found"
        }
        $retryCount = $maxRetries  # Exit loop on success
    }
    catch {
        $retryCount++
        if ($retryCount -lt $maxRetries) {
            Write-Host "No valid response. Retrying in 30 seconds... (Attempt $retryCount/$maxRetries)"
            Start-Sleep -Seconds 30
        }
    }
} while ($retryCount -lt $maxRetries)

invoke-webrequest -Uri $bloburl -Method Get -UseBasicParsing -OutFile "$Env:temp\$reportName"
$csv = import-csv "$Env:temp\$reportName" -Encoding UTF8
#$CultureName = "de-DE"
#$destculture = [System.Globalization.CultureInfo]::new($CultureName)
$reportNameCulture = "$($ConsumptionDate.ToString('yyyy-MM-dd'))_$($CultureName)_report.csv"
$transformedUsagePath = "$Env:temp\$reportNameCulture"
$csv | Select-object @{N = 'date'; E = { "{0}" -f [System.DateTime]::Parse($_.date, [CultureInfo]::new("en-us")).ToString("d", $destculture) } }, serviceFamily, consumedService, meterName, meterCategory, meterSubCategory, meterRegion, ProductName, resourceGroupName, @{N = 'ResourceName'; E = { ($_.ResourceId.Split('/')) | select -Last 1 } }, @{N = 'quantity'; E = { $([decimal]$_.quantity).ToString($destculture) } }, @{N = 'paygCostInBillingCurrency'; E = { $([decimal]$_.paygCostInBillingCurrency).ToString($destculture) } }, billingCurrency, unitOfMeasure, @{N = 'unitPrice'; E = { $([decimal]$_.unitPrice).ToString($destculture) } }, @{N = 'exchangeRatePricingToBilling'; E = { $([decimal]$_.exchangeRatePricingToBilling).ToString($destculture) } }, meterId, tags | Export-Csv "$Env:temp\$reportNameCulture" -Encoding UTF8 -Delimiter ';' -NoTypeInformation


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
$htmlBody += "</body></html>"


# You will need to authenticate before running this command.
# Learn more at https://docs.microsoft.com/rest/api/communication/authentication

# =========================================================================
# 6. VERSAND VIA AZURE COMMUNICATION SERVICES
# =========================================================================
$ErrorActionPreference = "Stop"

Write-Output "[INFO] Hole Token fuer ACS und versende E-Mail..."
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

Write-Output ("[OK]   E-Mail in Versandwarteschlange. Operation-Id: {0}, Status: {1}" -f $sendResult.id, $sendResult.status)
