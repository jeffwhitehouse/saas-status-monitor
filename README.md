# saas-status-monitor

Personal early-warning system for cloud and vendor outages, for a Windows workstation. Two
scheduled tasks:

| Task | Schedule | What it does |
|---|---|---|
| `ServiceStatusWatchdog` | every 30 min, 24/7 | Silent unless a service transitions to **DOWN** (or recovers), then emails you immediately. Re-alerts every 4 h while still down. |
| `ServiceStatusMorningDigest` | daily 6:30 AM | Always emails the full status table, worst-first. |

Every run also rewrites `StatusDashboard.html`, handy as a desktop shortcut.

## What's monitored

- **Microsoft 365**: live HTTPS probes of login, Outlook, Teams and (optionally) your tenant's
  SharePoint host. There is no public M365 status API without Graph consent; tenant-specific
  incident detail comes from the M365 admin center's service-health email notifications.
- **Azure**: official RSS. Any item in the feed = DOWN (Microsoft only posts significant incidents there).
- **AWS**: all-services RSS; items in the last 6 h = WARN (digest only, the feed is chatty).
- **Cloudflare, GitHub, Zoom, Atlassian**: Statuspage JSON (`/api/v2/summary.json`);
  minor = WARN, major/critical = DOWN.
- **Google Cloud / Google Workspace**: incidents.json; active high-severity = DOWN, else WARN.
- RSS feeds can be filtered to one product with `TitleFilter` (example in the script).

### Optional vendors

A built-in catalog of 35 more Statuspage vendors can be switched on by name, no editing required:
files and collaboration (Dropbox, Box, ShareFile, Notion, Miro, Figma...), work management (Asana,
monday.com, ClickUp, Airtable), identity (1Password, Duo, JumpCloud), device management and remote support
(Jamf, Kandji, NinjaOne, Kaseya/Datto, TeamViewer), phones and messaging (Dialpad, GoTo, Twilio, HubSpot),
cloud and developer platforms, and finance tools (QuickBooks, Xero, Shopify).

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ServiceStatusMonitor.ps1 -Mode Catalog      # list names
[Environment]::SetEnvironmentVariable('SSM_ADD_SERVICES', 'Dropbox,Duo,Jamf', 'User')          # tasks pick it up
powershell -NoProfile -ExecutionPolicy Bypass -File ServiceStatusMonitor.ps1 -Mode Test -Add Box  # one-off try
```

For anything else, edit the `$Services` array. Any Atlassian Statuspage vendor is one line.
Some vendors publish no machine-readable status feed at all; those can't be watched this way.

## Severity model

- **DOWN**: hard outage. The only thing that triggers an alert email.
- **WARN**: degraded / recent incident activity. Digest and dashboard only.
- **UNKNOWN**: the check itself failed (often local network). Never alerts.

## Setup

```powershell
# optional: recipient (defaults to your Outlook profile's address) and SharePoint probe host
[Environment]::SetEnvironmentVariable('SSM_EMAIL_TO', 'you@example.com', 'User')
[Environment]::SetEnvironmentVariable('SSM_SHAREPOINT_HOST', 'contoso.sharepoint.com', 'User')

.\Install-ServiceStatusTasks.ps1
```

Manual runs:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ServiceStatusMonitor.ps1 -Mode Test     # console only
powershell -NoProfile -ExecutionPolicy Bypass -File ServiceStatusMonitor.ps1 -Mode Digest   # sends the email
```

## Files

- `ServiceStatusMonitor.ps1`: everything. Service list, re-alert interval and timeouts at the top.
- `Run-ServiceStatusMonitor.vbs`: hidden-launch wrapper (a bare powershell.exe task action flashes a console).
- `Install-ServiceStatusTasks.ps1`: registers both tasks; re-run after changing schedules.
- `state.json` (generated): watchdog memory. Delete to reset.
- `ServiceStatusMonitor.log` (generated): rolling log, self-trims at 1 MB.

## Design constraints and gotchas

- **Windows PowerShell 5.1.** Keep `.ps1` files pure ASCII with a UTF-8 BOM (5.1 reads BOM-less
  files as ANSI). `Marshal.GetActiveObject` does not exist in PowerShell 7.
- **Email is Outlook COM**, so zero stored credentials. That requires an interactive session:
  tasks run as you, "run only when logged on". Don't move this to SYSTEM or an RMM without
  replacing the email mechanism.
- If Outlook isn't running, the script starts it **visibly**; hidden COM-launched Outlook traps dialogs.
- AWS RSS pubDates use `PDT`/`PST`, which .NET can't parse; `ConvertFrom-RssDate` maps them.
- Some vendors host Statuspage on an unexpected hostname rather than `status.<vendor>.com`; confirm
  `<Url>/api/v2/summary.json` returns JSON before adding one.
- Monitoring stops if the machine sleeps; pair it with a never-sleep guard.

## Uninstall

```powershell
Unregister-ScheduledTask -TaskName ServiceStatusWatchdog, ServiceStatusMorningDigest -Confirm:$false
```

## License

MIT - see [LICENSE](LICENSE).
