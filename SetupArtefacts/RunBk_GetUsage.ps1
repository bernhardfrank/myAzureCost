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
[string]$LookbackDays = Get-AutomationVariable -Name "LookbackDays"

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



$UsageAggregations = @()
$ErrorActionPreference = "SilentlyContinue"
$UsageAggregates = $null
do {
    if ($UsageAggregates.ContinuationToken) {
        "continue"
        $UsageAggregates = Get-UsageAggregates -ContinuationToken $($UsageAggregates.ContinuationToken) -ShowDetail $true -Verbose -ReportedStartTime $ConsumptionDate -ReportedEndTime $ConsumptionDate.addHours(25) -AggregationGranularity Hourly
    }
    else {
        "first data"
        $UsageAggregates = Get-UsageAggregates -ShowDetail $true -Verbose -ReportedStartTime $ConsumptionDate -ReportedEndTime $ConsumptionDate.addHours(25) -AggregationGranularity Hourly
    }

    foreach ($item in $UsageAggregates.UsageAggregations) {
        $UsageAggregations += $item
    }
}
while ($UsageAggregates.ContinuationToken)

$UsageToExport = $UsageAggregations | % { $_.Properties | select-object UsageStartTime, UsageEndTime, MeterCategory, MeterSubCategory, MeterName, @{N = 'InstanceName'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.resourceUri.Split('/') | select -Last 1 } }, @{N = 'RG'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.resourceUri.Split('/')[4] } }, @{N = 'Location'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.location } }, @{N = 'Quantity'; E = { $_.Quantity } }, Unit, MeterId, @{N = 'Tags'; E = { ($_.InstanceData | ConvertFrom-Json).'Microsoft.Resources'.tags } } } | where { ($(get-Date $_.UsageStartTime) -ge $(Get-date $ConsumptionDate.ToShortDateString()) -and ($(get-Date $_.UsageStartTime) -lt $(Get-date $ConsumptionDate.AddDays(1).ToShortDateString()))) } 
# sum up quantities of instances with same MeterID,date and rg 
$data = $UsageToExport | Group-Object InstanceName, RG, MeterID
$result = @()
$result += foreach ($item in $data) {
    $item.Group | Select-Object -Unique @{N = 'UsageStartTime'; E = { $($ConsumptionDate.ToString("d")) } }, @{N = 'UsageEndTime'; E = { $($ConsumptionDate.AddDays(1).ToString("d")) } }, MeterCategory, MeterSubCategory, MeterName, InstanceName, RG, Location, @{Name = 'Quantity'; Expression = { (($item.Group) | Measure-Object -Property Quantity -sum).Sum } }, Unit, MeterId, Tags
}
$result

$reportName = "AzureUsage$($ConsumptionDate.ToString("yyyy-MM-dd"))Consumption.csv"
$exportPath = "$Env:temp\$reportName"
$result | Export-Csv "$exportPath" -Encoding UTF8 -Delimiter ';' -NoTypeInformation

get-content $exportPath | Out-String | Write-Output

$transformedUsagePath = "$Env:temp\$($ConsumptionDate.ToString("yyyyMMdd"))ConsumptionCulture.csv"
$result | Select-object @{N = 'UsageStartTime'; E = { "{0}" -f [System.DateTime]::Parse($_.UsageStartTime,[CultureInfo]::new("en-us")).ToString("d",$destculture) } }, @{N = 'UsageEndTime'; E = { "{0}" -f [System.DateTime]::Parse($_.UsageEndTime,[CultureInfo]::new("en-us")).ToString("d",$destculture) } }, MeterCategory, MeterSubCategory, MeterName, InstanceName, RG, Location, @{N = 'Quantity'; E = { $([decimal]$_.Quantity).ToString($destculture) } }, Unit, MeterId, Tags | Export-Csv "$transformedUsagePath" -Encoding UTF8 -Delimiter ';' -NoTypeInformation

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
        to = @(@{ address = "bernhard@bernhardfrank.cloud" })
    }
    content       = @{
        subject= "Your Daily Azure Usage Report for $($ConsumptionDate.ToString("d",$destculture))"
        html = $htmlBody
    }
    attachments   = @(
        @{
            name            = $reportName
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