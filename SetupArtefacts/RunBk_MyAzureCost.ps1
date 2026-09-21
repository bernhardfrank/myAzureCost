<#
.SYNOPSIS
    myAzureCost 2.0 Runbook - taeglicher Azure-Kostenreport per E-Mail.
.DESCRIPTION
    Laeuft in Azure Automation mit System-assigned Managed Identity.
    Benoetigt KEINE PowerShell-Module (reines REST):
      - Token via Automation-Identity-Endpunkt (IDENTITY_ENDPOINT)
      - Kosten via Cost Management Query API (api-version 2025-03-01)
      - Versand via Azure Communication Services Email (api-version 2023-03-31)
    Throttling-Schutz: nur 3 API-Abfragen gesamt, 20 s Pause dazwischen,
    Retry-After-Header wird bei HTTP 429 beruecksichtigt.
    Report-Inhalt:
      - Gesamtkosten Vortag + Monat bis heute
      - Historie der letzten N Tage (Balken)
      - Kosten je Resource Group, je Service, je Region (Vortag)
      - CSV-Anhang: Kosten je Ressource (Vortag)
    Erforderliche Rollen der Managed Identity:
      - Cost Management Reader auf der Subscription
      - Contributor auf der ACS-Ressource (Mailversand per Entra-Token)
.NOTES
    =========================================================================
    Author      : Alexander Ortha
    Company     : Alexander Ortha IT Solutions
    Contact     : https://ortha-itsolutions.de/
    Created     : 16.09.2026
    Version     : 1.2.2
    Copyright   : (c) 2026 Alexander Ortha IT Solutions. All rights reserved.
    -------------------------------------------------------------------------
    Code created by Alexander Ortha.
    Development supported through AI-tools.
    -------------------------------------------------------------------------
    1.1.0: Throttling-Fix - 3 statt 6 Abfragen, Retry-After beachtet,
           max. 2 Gruppierungen je Abfrage (API-Limit), 20 s Pausen.
    1.1.1: UsageDate-Konvertierung tolerant (Zahl yyyyMMdd ODER ISO-Datum);
           unbekanntes Format wird mit Rohwert gemeldet.
    1.2.0: 429-Handling korrigiert - Cost-Management-spezifische
           Retry-After-Header werden gelesen (Maximum), ClientType-Header
           gesetzt (eigenes Rate-Limit-Kontingent), Quota-Header werden
           bei 429 als [DIAG] ins Job-Log geschrieben, Backoff verlaengert.
    1.2.1: Spaltennamen (Kosten/Datum) werden dynamisch aufgeloest statt
           fest angenommen; jede Abfrage loggt ihre Spalten als [DIAG];
           fehlende Spalten werden mit Spaltenliste gemeldet.
    1.2.2: Kritischer Fix - Write-Output innerhalb von Invoke-CostQuery
           verschmutzte den Rueckgabewert (Ursache der UsageDate-Fehler
           seit 1.1.0). Funktionsintern nur noch Warning-Stream; Erfolgs-
           Logging uebernimmt Write-QueryLog auf oberster Ebene.
    =========================================================================
.EXAMPLE
    Wird per Zeitplan mit Parametern gestartet (siehe Setup-MyAzureCost.ps1).
#>

# Get Azure Automation Variables (SubscriptionId, ACS Endpoint, Sender/Recipient, Culture, LookbackDays)
[string]$SubscriptionId = Get-AutomationVariable -Name "SubscriptionId"
[string]$AcsEndpoint = Get-AutomationVariable -Name "AcsEndpoint"
[string]$SenderAddress = Get-AutomationVariable -Name "SenderAddress"
[string]$RecipientEmail = Get-AutomationVariable -Name "RecipientEmail"
[string]$CultureInfo = Get-AutomationVariable -Name "CultureInfo" 
[string]$LookbackDays = Get-AutomationVariable -Name "LookbackDays"


$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Culture  = [System.Globalization.CultureInfo]::GetCultureInfo($CultureInfo)
$Lookback = [int]$LookbackDays
if ($Lookback -lt 2)  { $Lookback = 2 }
if ($Lookback -gt 30) { $Lookback = 30 }

$QueryPauseSeconds = 20   # Pause zwischen den Cost-Management-Abfragen

# =========================================================================
# HILFSFUNKTIONEN
# =========================================================================
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

function Invoke-CostQuery {
    # Fragt die Cost Management Query API ab und liefert Objekte zurueck.
    # Throttling (HTTP 429): Retry-After-Header wird beruecksichtigt,
    # bis zu 6 Versuche mit ansteigenden Wartezeiten.
    param(
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $true)][hashtable]$Body,
        [Parameter(Mandatory = $true)][string]$Label
    )
    $uri = "https://management.azure.com/subscriptions/{0}/providers/Microsoft.CostManagement/query?api-version=2025-03-01" -f $SubscriptionId
    $headers = @{
        "Authorization" = "Bearer {0}" -f $Token
        "Content-Type"  = "application/json"
        "ClientType"    = "myAzureCost-AOIT"
    }
    $json = $Body | ConvertTo-Json -Depth 10
    $maxAttempts = 6
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $json -UseBasicParsing
            # Ergebnis in Objekte umwandeln (columns + rows)
            $columns = @($response.properties.columns | ForEach-Object { $_.name })
            $items = @()
            foreach ($row in $response.properties.rows) {
                $obj = [ordered]@{}
                for ($i = 0; $i -lt $columns.Count; $i++) {
                    $obj[$columns[$i]] = $row[$i]
                }
                $items += [pscustomobject]$obj
            }
            # WICHTIG: Innerhalb dieser Funktion darf NICHT mit Write-Output
            # geloggt werden - das wuerde in den Rueckgabewert einfliessen.
            # Erfolgs-Logging uebernimmt der Aufrufer (Write-QueryLog).
            return ,$items
        }
        catch {
            $statusCode = $null
            $retryAfter = 0
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                try {
                    # Die Cost-Management-API meldet die Wartezeit NICHT im
                    # Standard-Header 'Retry-After', sondern in eigenen Headern.
                    # Es wird das Maximum aller vorhandenen Werte verwendet.
                    $respHeaders = $_.Exception.Response.Headers
                    $retryHeaderNames = @(
                        "Retry-After",
                        "x-ms-ratelimit-microsoft.consumption-retry-after",
                        "x-ms-ratelimit-microsoft.costmanagement-entity-retry-after",
                        "x-ms-ratelimit-microsoft.costmanagement-clienttype-retry-after",
                        "x-ms-ratelimit-microsoft.costmanagement-tenant-retry-after",
                        "x-ms-ratelimit-microsoft.costmanagement-client-retry-after",
                        "x-ms-ratelimit-microsoft.costmanagement-qpu-retry-after"
                    )
                    foreach ($hn in $retryHeaderNames) {
                        $hv = $respHeaders[$hn]
                        if ($hv) {
                            $hvInt = 0
                            if ([int]::TryParse([string]$hv, [ref]$hvInt) -and $hvInt -gt $retryAfter) {
                                $retryAfter = $hvInt
                            }
                        }
                    }
                    # Diagnose: alle Ratelimit-Header ins Job-Log schreiben
                    # (Warning-Stream, um die Rueckgabe nicht zu verschmutzen).
                    if ($statusCode -eq 429) {
                        foreach ($hn in $respHeaders.AllKeys) {
                            if ($hn -like "x-ms-ratelimit*") {
                                Write-Warning ("[DIAG] {0} = {1}" -f $hn, $respHeaders[$hn])
                            }
                        }
                    }
                }
                catch {
                    Write-Warning ("[WARN] Header-Auswertung fehlgeschlagen: {0}" -f $_.Exception.Message)
                }
            }
            if ($statusCode -eq 429 -and $attempt -lt $maxAttempts) {
                $wait = 60 * $attempt
                if (($retryAfter + 5) -gt $wait) { $wait = $retryAfter + 5 }
                if ($wait -gt 600) { $wait = 600 }
                Write-Warning ("Abfrage '{0}': Throttling (429). Warte {1} s (Versuch {2}/{3}, gemeldete Wartezeit: {4} s)..." -f $Label, $wait, $attempt, $maxAttempts, $retryAfter)
                Start-Sleep -Seconds $wait
            }
            else {
                throw
            }
        }
    }
}

function Format-Money {
    param([double]$Value, [string]$Currency)
    return ("{0} {1}" -f $Value.ToString("N2", $Culture), $Currency)
}

function Get-RgFromResourceId {
    # Extrahiert den Resource-Group-Namen aus einer Azure-Ressourcen-ID.
    param([string]$ResourceId)
    if ($ResourceId -match "(?i)/resourcegroups/([^/]+)/") {
        return $Matches[1]
    }
    return "(ohne Resource Group)"
}

function Convert-UsageDate {
    # Konvertiert den UsageDate-Wert der Cost Management API in [datetime].
    # Die API liefert je nach Version/Scope eine Zahl (yyyyMMdd) ODER einen
    # ISO-Datumsstring. Beide Formate werden akzeptiert; bei unbekanntem
    # Format wird der Rohwert in der Fehlermeldung ausgegeben.
    param([object]$Value)
    $s = ([string]$Value).Trim()
    if ($s -match "^\d{8}$") {
        return [datetime]::ParseExact($s, "yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date
    }
    throw ("UsageDate-Wert '{0}' hat ein unerwartetes Format (weder yyyyMMdd noch ISO-Datum)." -f $s)
}

function Resolve-ColumnName {
    # Ermittelt den tatsaechlichen Spaltennamen aus einem Ergebnisobjekt:
    # erst exakte Kandidaten, dann Namensmuster. Liefert $null, wenn nichts
    # passt - der Aufrufer entscheidet ueber die Fehlermeldung.
    param(
        [object]$Sample,
        [string[]]$Preferred,
        [string]$Pattern
    )
    if ($null -eq $Sample) { return $null }
    $propNames = @($Sample.PSObject.Properties | ForEach-Object { $_.Name })
    foreach ($cand in $Preferred) {
        foreach ($p in $propNames) {
            if ($p -ieq $cand) { return $p }
        }
    }
    foreach ($p in $propNames) {
        if ($p -like $Pattern) { return $p }
    }
    return $null
}

function Group-CostRows {
    # Aggregiert Kosten-Objekte nach einem Schluessel (Compat-sicher).
    param(
        [array]$Rows,
        [scriptblock]$KeySelector,
        [string]$CostProperty
    )
    $map = @{}
    foreach ($r in $Rows) {
        $key = & $KeySelector $r
        if ([string]::IsNullOrWhiteSpace($key)) { $key = "(unbekannt)" }
        if (-not $map.ContainsKey($key)) { $map[$key] = 0.0 }
        $map[$key] = $map[$key] + [double]$r.$CostProperty
    }
    $result = @()
    foreach ($key in $map.Keys) {
        $result += [pscustomobject]@{ Key = $key; Cost = $map[$key] }
    }
    return $result
}

function ConvertTo-HtmlTable {
    # Baut eine HTML-Tabelle mit Balkenanzeige (Anteil an Gesamtsumme).
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$KeyHeader,
        [Parameter(Mandatory = $true)][array]$Rows,     # Objekte mit .Key und .Cost
        [Parameter(Mandatory = $true)][string]$Currency,
        [Parameter(Mandatory = $false)][switch]$KeepOrder
    )
    $total = 0.0
    foreach ($r in $Rows) { $total += [double]$r.Cost }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(("<h3 style='margin:18px 0 6px 0;font-family:Segoe UI,Arial,sans-serif;'>{0}</h3>" -f $Title))
    [void]$sb.Append("<table style='border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>")
    [void]$sb.Append(("<tr style='background:#0078d4;color:#fff;'><th style='padding:6px;text-align:left;'>{0}</th><th style='padding:6px;text-align:right;'>Kosten</th><th style='padding:6px;text-align:left;width:35%;'>Anteil</th></tr>" -f $KeyHeader))
    if ($KeepOrder) {
        $sorted = $Rows
    } else {
        $sorted = $Rows | Sort-Object { [double]$_.Cost } -Descending
    }
    $rowIndex = 0
    foreach ($r in $sorted) {
        $pct = 0
        if ($total -gt 0) { $pct = [math]::Round(([double]$r.Cost / $total) * 100, 0) }
        $bg = "#ffffff"
        if (($rowIndex % 2) -eq 1) { $bg = "#f3f3f3" }
        $keyText = [System.Net.WebUtility]::HtmlEncode([string]$r.Key)
        [void]$sb.Append(("<tr style='background:{0};'>" -f $bg))
        [void]$sb.Append(("<td style='padding:5px;border-bottom:1px solid #ddd;'>{0}</td>" -f $keyText))
        [void]$sb.Append(("<td style='padding:5px;border-bottom:1px solid #ddd;text-align:right;white-space:nowrap;'>{0}</td>" -f (Format-Money -Value ([double]$r.Cost) -Currency $Currency)))
        [void]$sb.Append(("<td style='padding:5px;border-bottom:1px solid #ddd;'><div style='background:#0078d4;height:12px;width:{0}%;min-width:2px;'></div></td>" -f $pct))
        [void]$sb.Append("</tr>")
        $rowIndex++
    }
    [void]$sb.Append("</table>")
    return $sb.ToString()
}

# =========================================================================
# 1. TOKEN + ZEITRAEUME
# =========================================================================
Write-Output "[INFO] Hole Token fuer Cost Management API..."
$armToken = Get-MiToken -Resource "https://management.azure.com/"
Write-Output "[OK]   ARM-Token erhalten."

$todayUtc      = (Get-Date).ToUniversalTime().Date
$yesterday     = $todayUtc.AddDays(-1)
$monthStart    = New-Object System.DateTime($todayUtc.Year, $todayUtc.Month, 1)
$lookbackStart = $yesterday.AddDays(-1 * ($Lookback - 1))

# Ein Zeitraum fuer Historie UND Monat-bis-heute (eine Abfrage statt zwei):
$dailyStart = $lookbackStart
if ($monthStart -lt $dailyStart) { $dailyStart = $monthStart }

$fromYesterday = $yesterday.ToString("yyyy-MM-ddT00:00:00+00:00")
$toYesterday   = $yesterday.ToString("yyyy-MM-ddT23:59:59+00:00")

$aggregation = @{ totalCost = @{ name = "Cost"; function = "Sum" } }

# =========================================================================
# 2. ABFRAGEN (nur 3 gesamt - Throttling-Schutz)
# =========================================================================
function Write-QueryLog {
    # Loggt Zeilenanzahl und Spaltennamen einer Abfrage. Wird ausschliesslich
    # auf oberster Ebene aufgerufen, damit die Ausgabe im Job-Log landet und
    # keine Funktions-Rueckgabe verschmutzt.
    param([string]$Label, [array]$Items)
    $list = @($Items)
    Write-Output ("[OK]   Abfrage '{0}' erfolgreich ({1} Zeilen)." -f $Label, $list.Count)
    if ($list.Count -gt 0) {
        $cols = @($list[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ", "
        Write-Output ("[DIAG] Abfrage '{0}': Spalten = {1}" -f $Label, $cols)
    }
}

Write-Output ("[INFO] Frage Kosten fuer {0} ab (3 Abfragen, {1} s Pause)..." -f $yesterday.ToString("yyyy-MM-dd"), $QueryPauseSeconds)

# --- Abfrage 1: Vortag je Ressource + Service (max. 2 Gruppierungen) ---
# Resource Group wird lokal aus der ResourceId extrahiert.
$byRes = Invoke-CostQuery -Token $armToken -Label "Ressourcen" -Body @{
    type      = "ActualCost"
    timeframe = "Custom"
    timePeriod = @{ from = $fromYesterday; to = $toYesterday }
    dataset   = @{
        granularity = "None"
        aggregation = $aggregation
        grouping    = @(
            @{ type = "Dimension"; name = "ResourceId" },
            @{ type = "Dimension"; name = "ServiceName" }
        )
    }
}
Write-QueryLog -Label "Ressourcen" -Items $byRes
Start-Sleep -Seconds $QueryPauseSeconds

# --- Abfrage 2: Vortag je Region ---
$byLoc = Invoke-CostQuery -Token $armToken -Label "Regionen" -Body @{
    type      = "ActualCost"
    timeframe = "Custom"
    timePeriod = @{ from = $fromYesterday; to = $toYesterday }
    dataset   = @{
        granularity = "None"
        aggregation = $aggregation
        grouping    = @(@{ type = "Dimension"; name = "ResourceLocation" })
    }
}
Write-QueryLog -Label "Regionen" -Items $byLoc
Start-Sleep -Seconds $QueryPauseSeconds

# --- Abfrage 3: Taegliche Kosten (deckt Historie UND Monat-bis-heute ab) ---
$daily = Invoke-CostQuery -Token $armToken -Label "Tagesverlauf" -Body @{
    type      = "ActualCost"
    timeframe = "Custom"
    timePeriod = @{
        from = $dailyStart.ToString("yyyy-MM-ddT00:00:00+00:00")
        to   = $yesterday.ToString("yyyy-MM-ddT23:59:59+00:00")
    }
    dataset   = @{
        granularity = "Daily"
        aggregation = $aggregation
    }
}
Write-QueryLog -Label "Tagesverlauf" -Items $daily
Write-Output "[OK]   Alle Kostenabfragen abgeschlossen."

# =========================================================================
# 3. AUFBEREITEN
# =========================================================================
$byRes = @($byRes)
$byLoc = @($byLoc)
$daily = @($daily)

# Spaltennamen dynamisch aufloesen - die tatsaechlichen Namen stehen als
# [DIAG]-Zeilen im Job-Log. Bekannte Kandidaten zuerst, dann Namensmuster.
$costCands = @("Cost", "PreTaxCost", "totalCost", "CostUSD")
$dateCands = @("UsageDate", "UsageDateTime", "BillingDate", "Date")

$resCostCol = $null
if ($byRes.Count -gt 0) {
    $resCostCol = Resolve-ColumnName -Sample $byRes[0] -Preferred $costCands -Pattern "*Cost*"
    if (-not $resCostCol) {
        throw ("Keine Kosten-Spalte in Abfrage 'Ressourcen' gefunden. Vorhandene Spalten: {0}" -f (@($byRes[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ", "))
    }
}
$locCostCol = $null
if ($byLoc.Count -gt 0) {
    $locCostCol = Resolve-ColumnName -Sample $byLoc[0] -Preferred $costCands -Pattern "*Cost*"
}
$dailyCostCol = $null
$dailyDateCol = $null
if ($daily.Count -gt 0) {
    $dailyCostCol = Resolve-ColumnName -Sample $daily[0] -Preferred $costCands -Pattern "*Cost*"
    $dailyDateCol = Resolve-ColumnName -Sample $daily[0] -Preferred $dateCands -Pattern "*Date*"
    if (-not $dailyCostCol -or -not $dailyDateCol) {
        throw ("Kosten- oder Datums-Spalte in Abfrage 'Tagesverlauf' nicht gefunden. Vorhandene Spalten: {0}" -f (@($daily[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ", "))
    }
}

$currency = "EUR"
$firstWithCurrency = @($byRes) + @($daily) | Where-Object { $_.PSObject.Properties.Name -contains "Currency" } | Select-Object -First 1
if ($firstWithCurrency) { $currency = $firstWithCurrency.Currency }

# Vortag gesamt
$totalYesterday = 0.0
foreach ($r in $byRes) { $totalYesterday += [double]$r.$resCostCol }

# Monat bis heute: Summe der Tageswerte ab Monatsanfang
# (am 1. eines Monats gibt es noch keine Tageswerte im neuen Monat -> 0)
$totalMtd = 0.0
foreach ($d in $daily) {
    $dDate = Convert-UsageDate -Value $d.$dailyDateCol
    if ($dDate -ge $monthStart) { $totalMtd += [double]$d.$dailyCostCol }
}

# Aggregationen aus Abfrage 1 (lokal, ohne weitere API-Aufrufe)
$rgRows  = @()
$svcRows = @()
if ($byRes.Count -gt 0) {
    $rgRows  = @(Group-CostRows -Rows $byRes -CostProperty $resCostCol -KeySelector { param($r) Get-RgFromResourceId -ResourceId ([string]$r.ResourceId) })
    $svcRows = @(Group-CostRows -Rows $byRes -CostProperty $resCostCol -KeySelector { param($r) [string]$r.ServiceName })
}
$locRows = @()
if ($byLoc.Count -gt 0 -and $locCostCol) {
    $locRows = @($byLoc | ForEach-Object { [pscustomobject]@{ Key = $_.ResourceLocation; Cost = [double]$_.$locCostCol } })
}

# Historie: nur die letzten N Tage anzeigen
$histRows = @()
if ($daily.Count -gt 0) {
    foreach ($h in ($daily | Sort-Object { Convert-UsageDate -Value $_.$dailyDateCol })) {
        $d = Convert-UsageDate -Value $h.$dailyDateCol
        if ($d -ge $lookbackStart) {
            $histRows += [pscustomobject]@{ Key = $d.ToString("ddd dd.MM.", $Culture); Cost = [double]$h.$dailyCostCol }
        }
    }
}

# =========================================================================
# 4. HTML-REPORT
# =========================================================================
$reportDate = $yesterday.ToString("dddd, dd. MMMM yyyy", $Culture)
$html = New-Object System.Text.StringBuilder
[void]$html.Append("<html><body style='margin:0;padding:16px;background:#fafafa;'>")
[void]$html.Append("<div style='max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #e0e0e0;padding:20px;font-family:Segoe UI,Arial,sans-serif;'>")
[void]$html.Append(("<h2 style='margin:0 0 4px 0;color:#0078d4;'>Azure Kostenreport</h2>"))
[void]$html.Append(("<p style='margin:0 0 14px 0;color:#666;font-size:13px;'>Zeitraum: {0} &middot; Subscription: {1}</p>" -f $reportDate, $SubscriptionId))
[void]$html.Append("<table style='width:100%;border-collapse:collapse;'><tr>")
[void]$html.Append(("<td style='background:#0078d4;color:#fff;padding:14px;text-align:center;width:50%;'><div style='font-size:12px;'>Kosten Vortag</div><div style='font-size:22px;font-weight:bold;'>{0}</div></td>" -f (Format-Money -Value $totalYesterday -Currency $currency)))
[void]$html.Append(("<td style='background:#004e8c;color:#fff;padding:14px;text-align:center;width:50%;'><div style='font-size:12px;'>Monat bis heute</div><div style='font-size:22px;font-weight:bold;'>{0}</div></td>" -f (Format-Money -Value $totalMtd -Currency $currency)))
[void]$html.Append("</tr></table>")

if ($histRows.Count -gt 0) {
    [void]$html.Append((ConvertTo-HtmlTable -Title ("Verlauf letzte {0} Tage" -f $Lookback) -KeyHeader "Tag" -Rows $histRows -Currency $currency -KeepOrder))
}
if ($rgRows.Count -gt 0) {
    [void]$html.Append((ConvertTo-HtmlTable -Title "Kosten je Resource Group (Vortag)" -KeyHeader "Resource Group" -Rows $rgRows -Currency $currency))
}
if ($svcRows.Count -gt 0) {
    [void]$html.Append((ConvertTo-HtmlTable -Title "Kosten je Service (Vortag)" -KeyHeader "Service" -Rows $svcRows -Currency $currency))
}
if ($locRows.Count -gt 0) {
    [void]$html.Append((ConvertTo-HtmlTable -Title "Kosten je Region (Vortag)" -KeyHeader "Region" -Rows $locRows -Currency $currency))
}
if ($rgRows.Count -eq 0 -and $histRows.Count -eq 0) {
    [void]$html.Append("<p style='color:#a00;'>Keine Kostendaten gefunden. Kostendaten des Vortags stehen erst mit einigen Stunden Verzug bereit - ggf. Zeitplan spaeter ansetzen.</p>")
}
[void]$html.Append("<p style='margin-top:18px;color:#999;font-size:11px;'>Automatisch erstellt durch myAzureCost 2.0 &middot; Alexander Ortha IT Solutions</p>")
[void]$html.Append("</div></body></html>")

# =========================================================================
# 5. CSV-ANHANG (Kosten je Ressource, Vortag)
# =========================================================================
$csvSb = New-Object System.Text.StringBuilder
[void]$csvSb.AppendLine("Ressource;ResourceGroup;Service;Kosten;Waehrung")
foreach ($r in ($byRes | Sort-Object { [double]$_.$resCostCol } -Descending)) {
    $resName = [string]$r.ResourceId
    if ($resName -match "/") { $resName = $resName.Split("/")[-1] }
    $rgName = Get-RgFromResourceId -ResourceId ([string]$r.ResourceId)
    $line = "{0};{1};{2};{3};{4}" -f $resName, $rgName, $r.ServiceName, ([double]$r.$resCostCol).ToString("N4", $Culture), $currency
    [void]$csvSb.AppendLine($line)
}
# UTF-8 mit BOM, damit Excel Umlaute korrekt anzeigt
$bom      = [byte[]](0xEF, 0xBB, 0xBF)
$csvBytes = $bom + [System.Text.Encoding]::UTF8.GetBytes($csvSb.ToString())
$csvB64   = [Convert]::ToBase64String($csvBytes)
$csvName  = "AzureCost_{0}.csv" -f $yesterday.ToString("yyyy-MM-dd")

# =========================================================================
# 6. VERSAND VIA AZURE COMMUNICATION SERVICES
# =========================================================================
Write-Output "[INFO] Hole Token fuer ACS und versende E-Mail..."
$acsToken = Get-MiToken -Resource "https://communication.azure.com/"

$mailBody = @{
    senderAddress = $SenderAddress
    recipients    = @{
        to = @(@{ address = $RecipientEmail })
    }
    content       = @{
        subject = ("Azure Kostenreport {0}: {1}" -f $yesterday.ToString("dd.MM.yyyy"), (Format-Money -Value $totalYesterday -Currency $currency))
        html    = $html.ToString()
    }
    attachments   = @(
        @{
            name            = $csvName
            contentType     = "text/csv"
            contentInBase64 = $csvB64
        }
    )
}

$sendUri = "{0}/emails:send?api-version=2023-03-31" -f $AcsEndpoint.TrimEnd("/")
$sendHeaders = @{
    "Authorization" = "Bearer {0}" -f $acsToken
    "Content-Type"  = "application/json"
}
$sendResult = Invoke-RestMethod -Uri $sendUri -Method Post -Headers $sendHeaders `
    -Body ($mailBody | ConvertTo-Json -Depth 10) -UseBasicParsing

Write-Output ("[OK]   E-Mail in Versandwarteschlange. Operation-Id: {0}, Status: {1}" -f $sendResult.id, $sendResult.status)
Write-Output ("[OK]   Report fuer {0} an {1} gesendet. Gesamt: {2}" -f $yesterday.ToString("yyyy-MM-dd"), $RecipientEmail, (Format-Money -Value $totalYesterday -Currency $currency))
