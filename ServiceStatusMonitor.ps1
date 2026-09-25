<#
.SYNOPSIS
  ServiceStatusMonitor - vendor/cloud service status watchdog and morning digest.

.DESCRIPTION
  Polls public status feeds (Statuspage JSON, RSS, Google incidents JSON) and
  probes live HTTPS endpoints, then reports:

    -Mode Watchdog  (default) Silent unless something transitions to DOWN
                    (or recovers). Sends an alert email via Outlook COM.
                    Re-alerts every $ReAlertHours while still down.
    -Mode Digest    Always sends the full morning status report email.
    -Mode Test      Console output only. No email, no state changes.

  Every run also rewrites StatusDashboard.html next to this script.

  Requires: Windows PowerShell 5.1, classic Outlook (COM), interactive user
  session. Do NOT run as SYSTEM - Outlook COM is unsupported there.

  Severity model:
    DOWN    = hard outage (alert-worthy, wakes you up via email)
    WARN    = degraded / recent incident activity (digest + dashboard only)
    UNKNOWN = the check itself failed (never alerts; often local network)
    OK      = all clear
#>
param(
    [ValidateSet('Watchdog','Digest','Test')]
    [string]$Mode = 'Watchdog',

    # Alert/digest recipient. Defaults to $env:SSM_EMAIL_TO, else the Outlook profile's own address.
    [string]$EmailTo = $env:SSM_EMAIL_TO,

    # Your tenant's SharePoint host for the live M365 probe, e.g. contoso.sharepoint.com.
    # Leave empty to skip the SharePoint probe.
    [string]$SharePointHost = $env:SSM_SHAREPOINT_HOST,

    [string]$Title = 'Service Status'
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------- settings --
$Root           = $PSScriptRoot
$StateFile      = Join-Path $Root 'state.json'
$LogFile        = Join-Path $Root 'ServiceStatusMonitor.log'
$DashboardFile  = Join-Path $Root 'StatusDashboard.html'
$ReAlertHours   = 4
$HttpTimeoutSec = 20

# Service checklist. Types:
#   statuspage - Atlassian Statuspage: polls <Url>/api/v2/summary.json
#   rss        - RSS feed; items newer than WindowHours => RssSeverity
#                (optional TitleFilter regex keeps only matching items)
#   googlejson - Google incidents.json; active = incident with no "end"
#   http       - live HTTPS probes; any response < 500 counts as reachable
$Services = @(
    @{ Name = 'Microsoft 365 (login, Outlook, Teams, SharePoint)'
       Type = 'http'
       Urls = @(
           'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
           'https://outlook.office365.com',
           'https://teams.microsoft.com'
       ) + $(if ($SharePointHost) { @("https://$SharePointHost") } else { @() })
       Link = 'https://status.cloud.microsoft/' }

    @{ Name = 'Azure'
       Type = 'rss'
       Url  = 'https://azure.status.microsoft/en-us/status/feed/'
       WindowHours = 24
       RssSeverity = 'DOWN'
       Link = 'https://azure.status.microsoft/en-us/status' }

    @{ Name = 'AWS'
       Type = 'rss'
       Url  = 'https://status.aws.amazon.com/rss/all.rss'
       WindowHours = 6
       RssSeverity = 'WARN'
       Link = 'https://health.aws.amazon.com/health/status' }

    @{ Name = 'Cloudflare'
       Type = 'statuspage'
       Url  = 'https://www.cloudflarestatus.com'
       Link = 'https://www.cloudflarestatus.com' }

    @{ Name = 'Google Cloud'
       Type = 'googlejson'
       Url  = 'https://status.cloud.google.com/incidents.json'
       Link = 'https://status.cloud.google.com/' }

    @{ Name = 'Google Workspace'
       Type = 'googlejson'
       Url  = 'https://www.google.com/appsstatus/dashboard/incidents.json'
       Link = 'https://www.google.com/appsstatus/dashboard/' }

    @{ Name = 'GitHub'
       Type = 'statuspage'
       Url  = 'https://www.githubstatus.com'
       Link = 'https://www.githubstatus.com' }

    @{ Name = 'Zoom'
       Type = 'statuspage'
       Url  = 'https://www.zoomstatus.com'
       Link = 'https://www.zoomstatus.com' }

    @{ Name = 'Atlassian (Jira / Confluence)'
       Type = 'statuspage'
       Url  = 'https://status.atlassian.com'
       Link = 'https://status.atlassian.com' }

    # More examples - any Atlassian Statuspage vendor is one entry:
    #   @{ Name = 'Dropbox'; Type = 'statuspage'; Url = 'https://status.dropbox.com'; Link = 'https://status.dropbox.com' }
    # RSS feeds can be filtered to the product you care about:
    #   @{ Name = 'DigitalOcean Spaces'; Type = 'rss'; Url = 'https://status.digitalocean.com/history.rss'; WindowHours = 24
    #      RssSeverity = 'WARN'; TitleFilter = 'Spaces'; Link = 'https://status.digitalocean.com' }
)

# ----------------------------------------------------------------- logging --
function Write-Log {
    param([string]$Message)
    $line = ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Mode, $Message)
    Add-Content -Path $LogFile -Value $line -Encoding ASCII
    if ($Mode -eq 'Test') { Write-Host $line }
}

function Limit-LogSize {
    try {
        if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt 1MB) {
            # Parentheses required: same-file Get-Content|Set-Content pipeline fails otherwise
            (Get-Content $LogFile -Tail 2000) | Set-Content $LogFile -Encoding ASCII
        }
    } catch { }
}

# ------------------------------------------------------------------ checks --
function Test-Statuspage {
    param($Svc)
    $r = Invoke-RestMethod -Uri ($Svc.Url + '/api/v2/summary.json') -TimeoutSec $HttpTimeoutSec
    $status = switch ($r.status.indicator) {
        'none'     { 'OK' }
        'minor'    { 'WARN' }
        'major'    { 'DOWN' }
        'critical' { 'DOWN' }
        default    { 'WARN' }
    }
    $detail = [string]$r.status.description
    $names = @($r.incidents | Where-Object { $_ } | ForEach-Object { $_.name })
    if ($names.Count -gt 0) {
        $detail = $detail + ' | Incidents: ' + (($names | Select-Object -First 3) -join '; ')
    }
    return @{ Status = $status; Detail = $detail }
}

function ConvertFrom-RssDate {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($t, [ref]$dto)) { return $dto.LocalDateTime }
    # RFC822 offsets like "+0530" need a colon for .NET
    $t2 = $t -replace '([+-]\d{2})(\d{2})$', '$1:$2'
    if ([DateTimeOffset]::TryParse($t2, [ref]$dto)) { return $dto.LocalDateTime }
    # US timezone abbreviations (AWS feed uses PDT/PST) are not parseable by .NET
    $tzMap = @{ 'PDT' = '-07:00'; 'PST' = '-08:00'; 'MDT' = '-06:00'; 'MST' = '-07:00'
                'CDT' = '-05:00'; 'CST' = '-06:00'; 'EDT' = '-04:00'; 'EST' = '-05:00'
                'UT'  = '+00:00' }
    foreach ($abbr in $tzMap.Keys) {
        if ($t -match ('\s' + $abbr + '$')) {
            $t3 = $t -replace ('\s' + $abbr + '$'), (' ' + $tzMap[$abbr])
            if ([DateTimeOffset]::TryParse($t3, [ref]$dto)) { return $dto.LocalDateTime }
        }
    }
    return $null
}

function Test-RssFeed {
    param($Svc)
    $resp = Invoke-WebRequest -Uri $Svc.Url -TimeoutSec $HttpTimeoutSec -UseBasicParsing
    [xml]$xml = $resp.Content
    $items = @($xml.rss.channel.item | Where-Object { $_ })
    if ($Svc['TitleFilter']) {
        $items = @($items | Where-Object { [string]$_.title -match $Svc.TitleFilter })
    }
    $cutoff = (Get-Date).AddHours(-1 * $Svc.WindowHours)
    $recent = @()
    foreach ($it in $items) {
        $when = ConvertFrom-RssDate ([string]$it.pubDate)
        if ($null -ne $when -and $when -gt $cutoff) {
            $recent += @{ Title = [string]$it.title; When = $when }
        }
    }
    if ($recent.Count -eq 0) {
        return @{ Status = 'OK'; Detail = ('No incidents in feed (last {0}h)' -f $Svc.WindowHours) }
    }
    $latest = $recent | Sort-Object { $_.When } -Descending | Select-Object -First 1
    $detail = ('{0} feed item(s) in last {1}h. Latest: {2}' -f $recent.Count, $Svc.WindowHours, $latest.Title)
    return @{ Status = $Svc.RssSeverity; Detail = $detail }
}

function Test-GoogleIncidents {
    param($Svc)
    $incidents = Invoke-RestMethod -Uri $Svc.Url -TimeoutSec 30
    $active = @($incidents | Where-Object { -not $_.end })
    if ($active.Count -eq 0) {
        return @{ Status = 'OK'; Detail = 'No active incidents' }
    }
    $high = @($active | Where-Object { [string]$_.severity -eq 'high' })
    $status = 'WARN'
    if ($high.Count -gt 0) { $status = 'DOWN' }
    $desc = [string]($active[0].external_desc) -replace '\s+', ' '
    if ($desc.Length -gt 140) { $desc = $desc.Substring(0, 140) + '...' }
    return @{ Status = $status; Detail = ('{0} active incident(s). {1}' -f $active.Count, $desc) }
}

function Test-HttpEndpoints {
    param($Svc)
    $failed = @()
    foreach ($u in $Svc.Urls) {
        $up = $false
        try {
            $resp = Invoke-WebRequest -Uri $u -TimeoutSec $HttpTimeoutSec -UseBasicParsing -MaximumRedirection 5
            if ([int]$resp.StatusCode -lt 500) { $up = $true }
        } catch [System.Net.WebException] {
            $r = $_.Exception.Response
            if ($null -ne $r) {
                try { if ([int]$r.StatusCode -lt 500) { $up = $true } } catch { }
            }
        } catch { }
        if (-not $up) { $failed += ([Uri]$u).Host }
    }
    $total = @($Svc.Urls).Count
    if ($failed.Count -eq 0) {
        return @{ Status = 'OK'; Detail = ('All {0} endpoints reachable' -f $total) }
    }
    if ($failed.Count -eq $total) {
        return @{ Status = 'DOWN'; Detail = ('All endpoints unreachable: ' + ($failed -join ', ')) }
    }
    return @{ Status = 'WARN'; Detail = ('Unreachable: ' + ($failed -join ', ')) }
}

# ------------------------------------------------------------------- state --
function Get-MonitorState {
    if (-not (Test-Path $StateFile)) { return @{} }
    try {
        $json = Get-Content $StateFile -Raw | ConvertFrom-Json
        $h = @{}
        foreach ($p in $json.PSObject.Properties) {
            $h[$p.Name] = @{
                Status    = [string]$p.Value.Status
                Since     = [string]$p.Value.Since
                LastAlert = [string]$p.Value.LastAlert
            }
        }
        return $h
    } catch {
        Write-Log ('WARNING: could not read state file, starting fresh. ' + $_.Exception.Message)
        return @{}
    }
}

function Save-MonitorState {
    param($State)
    $State | ConvertTo-Json -Depth 4 | Set-Content -Path $StateFile -Encoding ASCII
}

# -------------------------------------------------------------------- html --
function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        'OK'      { return '#2e7d32' }
        'WARN'    { return '#e09b00' }
        'DOWN'    { return '#c62828' }
        default   { return '#757575' }
    }
}

function New-StatusHtml {
    param($Results, [string]$Heading, [string]$Intro)
    $enc = [System.Net.WebUtility]
    $rank = @{ 'DOWN' = 0; 'UNKNOWN' = 1; 'WARN' = 2; 'OK' = 3 }
    $sorted = $Results | Sort-Object { $rank[$_.Status] }, { $_.Name }

    $rows = ''
    foreach ($r in $sorted) {
        $color = Get-StatusColor $r.Status
        $name = $enc::HtmlEncode($r.Name)
        if ($r.Link) {
            $name = '<a href="' + $r.Link + '" style="color:#1a4b8b;text-decoration:none;">' + $name + '</a>'
        }
        $rows += '<tr>' +
            '<td style="padding:7px 10px;border-bottom:1px solid #e3e3e3;white-space:nowrap;">' +
            '<span style="display:inline-block;padding:2px 10px;border-radius:10px;background:' + $color +
            ';color:#ffffff;font-weight:bold;font-size:12px;">' + $r.Status + '</span></td>' +
            '<td style="padding:7px 10px;border-bottom:1px solid #e3e3e3;font-weight:bold;">' + $name + '</td>' +
            '<td style="padding:7px 10px;border-bottom:1px solid #e3e3e3;color:#444444;">' +
            $enc::HtmlEncode($r.Detail) + '</td></tr>'
    }

    $introHtml = ''
    if ($Intro) { $introHtml = '<p style="font-size:14px;color:#333333;">' + $Intro + '</p>' }

    $html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:860px;">' +
        '<h2 style="color:#1a4b8b;margin-bottom:4px;">' + $enc::HtmlEncode($Heading) + '</h2>' +
        '<p style="color:#888888;font-size:12px;margin-top:0;">Generated ' +
        (Get-Date -Format 'dddd, MMMM d, yyyy h:mm tt') + ' on ' + $env:COMPUTERNAME + ' (mode: ' + $Mode + ')</p>' +
        $introHtml +
        '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
        '<tr style="background:#f0f4fa;">' +
        '<th style="text-align:left;padding:7px 10px;">Status</th>' +
        '<th style="text-align:left;padding:7px 10px;">Service</th>' +
        '<th style="text-align:left;padding:7px 10px;">Detail</th></tr>' +
        $rows + '</table>' +
        '<p style="color:#888888;font-size:11px;margin-top:14px;">' +
        'DOWN = hard outage. WARN = degraded or recent incident activity. UNKNOWN = check failed (often local network). ' +
        'Tenant-specific Microsoft 365 incidents arrive separately via admin center service health email notifications.</p>' +
        '</div>'
    return $html
}

function Write-Dashboard {
    param($Results)
    $body = New-StatusHtml -Results $Results -Heading "$Title Dashboard" -Intro ''
    $page = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
        '<meta http-equiv="refresh" content="600">' +
        "<title>$Title</title></head>" +
        '<body style="background:#fafafa;padding:20px;">' + $body + '</body></html>'
    Set-Content -Path $DashboardFile -Value $page -Encoding UTF8
}

# ------------------------------------------------------------------- email --
function Get-OutlookApp {
    # Attach to a running Outlook first; hidden COM-launched instances trap dialogs,
    # so if we must start it, start it visibly via App Paths and poll.
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { }
    $exe = $null
    try {
        $exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction Stop).'(default)'
    } catch { }
    if ($exe) { Start-Process -FilePath $exe } else { Start-Process -FilePath 'outlook.exe' }
    Write-Log 'Outlook was not running; started it visibly and waiting for COM...'
    for ($i = 0; $i -lt 12; $i++) {
        Start-Sleep -Seconds 5
        try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { }
    }
    throw 'Outlook did not become available via COM within 60 seconds.'
}

function Send-StatusEmail {
    param([string]$Subject, [string]$HtmlBody)
    $ol = $null; $mail = $null
    try {
        $ol = Get-OutlookApp
        $to = $EmailTo
        if (-not $to) { $to = $ol.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress }
        if (-not $to) { $to = $ol.Session.CurrentUser.Address }
        $mail = $ol.CreateItem(0)
        $mail.To = $to
        $mail.Subject = $Subject
        $mail.HTMLBody = $HtmlBody
        $mail.Send()
        Write-Log ('Email sent: ' + $Subject)
    } finally {
        if ($null -ne $mail) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
        if ($null -ne $ol)   { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol) }
    }
}

# -------------------------------------------------------------------- main --
Limit-LogSize
Write-Log '=== Run start ==='

$results = @()
foreach ($svc in $Services) {
    $res = $null
    try {
        $res = switch ($svc.Type) {
            'statuspage' { Test-Statuspage       $svc }
            'rss'        { Test-RssFeed          $svc }
            'googlejson' { Test-GoogleIncidents  $svc }
            'http'       { Test-HttpEndpoints    $svc }
        }
    } catch {
        $res = @{ Status = 'UNKNOWN'; Detail = ('Check failed: ' + $_.Exception.Message) }
    }
    $results += @{ Name = $svc.Name; Status = $res.Status; Detail = $res.Detail; Link = $svc['Link'] }
    Write-Log ('{0}: {1} - {2}' -f $svc.Name, $res.Status, $res.Detail)
}

Write-Dashboard -Results $results

$downCount    = @($results | Where-Object { $_.Status -eq 'DOWN' }).Count
$warnCount    = @($results | Where-Object { $_.Status -eq 'WARN' }).Count
$unknownCount = @($results | Where-Object { $_.Status -eq 'UNKNOWN' }).Count

if ($Mode -eq 'Test') {
    $results | ForEach-Object { [pscustomobject]$_ } |
        Sort-Object @{ Expression = { @{ 'DOWN' = 0; 'UNKNOWN' = 1; 'WARN' = 2; 'OK' = 3 }[$_.Status] } } |
        Format-Table Status, Name, Detail -AutoSize -Wrap
    Write-Host ('Dashboard written to: ' + $DashboardFile)
    Write-Log '=== Run end (test, no email/state) ==='
    return
}

if ($Mode -eq 'Digest') {
    $issues = $downCount + $warnCount
    if ($issues -eq 0) {
        $subject = ('[Service Status] Morning digest - all clear ({0} services OK)' -f @($results).Count)
        $intro = 'All monitored services look healthy. Have a good one.'
    } else {
        $subject = ('[Service Status] Morning digest - {0} DOWN, {1} degraded' -f $downCount, $warnCount)
        $intro = 'Heads up: some services are reporting problems. Details below, worst first.'
    }
    if ($unknownCount -gt 0) {
        $intro += (' ({0} check(s) could not run - see UNKNOWN rows.)' -f $unknownCount)
    }
    $html = New-StatusHtml -Results $results -Heading 'Morning Service Status Digest' -Intro $intro
    Send-StatusEmail -Subject $subject -HtmlBody $html
    Write-Log '=== Run end (digest) ==='
    return
}

# Watchdog mode: compare against previous state, alert on transitions only.
$state = Get-MonitorState
$now = Get-Date
$newAlerts = @()
$stillDown = @()
$recovered = @()

foreach ($r in $results) {
    $prevStatus = 'OK'
    $prev = $null
    if ($state.ContainsKey($r.Name)) { $prev = $state[$r.Name]; $prevStatus = $prev.Status }

    if ($r.Status -eq 'DOWN') {
        if ($prevStatus -ne 'DOWN') {
            $newAlerts += $r
            $state[$r.Name] = @{
                Status    = 'DOWN'
                Since     = $now.ToString('o')
                LastAlert = $now.ToString('o')
            }
        } else {
            $lastAlert = [DateTime]::Parse($prev.LastAlert)
            if (($now - $lastAlert).TotalHours -ge $ReAlertHours) {
                $since = [DateTime]::Parse($prev.Since)
                $r.Detail = ('STILL DOWN since {0}. {1}' -f $since.ToString('MM/dd h:mm tt'), $r.Detail)
                $stillDown += $r
                $prev.LastAlert = $now.ToString('o')
                $state[$r.Name] = $prev
            }
        }
    } else {
        if ($prevStatus -eq 'DOWN') {
            $sinceTxt = ''
            try {
                $since = [DateTime]::Parse($prev.Since)
                $sinceTxt = (' (was down {0:N1} hours)' -f ($now - $since).TotalHours)
            } catch { }
            $r.Detail = ('RECOVERED' + $sinceTxt + '. ' + $r.Detail)
            $recovered += $r
        }
        if ($state.ContainsKey($r.Name) -and $state[$r.Name].Status -ne $r.Status) {
            $state[$r.Name] = @{ Status = $r.Status; Since = $now.ToString('o'); LastAlert = '' }
        } elseif (-not $state.ContainsKey($r.Name)) {
            $state[$r.Name] = @{ Status = $r.Status; Since = $now.ToString('o'); LastAlert = '' }
        }
    }
}

Save-MonitorState -State $state

$toNotify = @($newAlerts) + @($stillDown) + @($recovered)
if ($toNotify.Count -gt 0) {
    $parts = @()
    if ($newAlerts.Count -gt 0) { $parts += ('DOWN: ' + (($newAlerts   | ForEach-Object { $_.Name }) -join ', ')) }
    if ($stillDown.Count -gt 0) { $parts += ('STILL DOWN: ' + (($stillDown | ForEach-Object { $_.Name }) -join ', ')) }
    if ($recovered.Count -gt 0) { $parts += ('RECOVERED: ' + (($recovered | ForEach-Object { $_.Name }) -join ', ')) }
    $subject = '[Service Status] ALERT - ' + ($parts -join ' | ')
    $intro = 'The overnight watchdog detected a change. Full current picture below, worst first.'
    $html = New-StatusHtml -Results $results -Heading 'Service Status Alert' -Intro $intro
    Send-StatusEmail -Subject $subject -HtmlBody $html
} else {
    Write-Log ('No transitions. DOWN={0} WARN={1} UNKNOWN={2}' -f $downCount, $warnCount, $unknownCount)
}
Write-Log '=== Run end (watchdog) ==='
