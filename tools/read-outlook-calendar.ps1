param([ValidateRange(1, 90)][int]$Days = 35)

$ErrorActionPreference = "Stop"
$outlook = $null
$namespace = $null
$calendar = $null
$items = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    $start = Get-Date
    $end = $start.AddDays($Days)
    $restriction = "[Start] >= '" + $start.ToString("g") + "' AND [Start] < '" + $end.ToString("g") + "'"
    $events = @()

    foreach ($item in $items.Restrict($restriction)) {
        # Only meetings accepted by the current user (3) or organized by them
        # (1). Status 5 is cancelled and 7 is received-and-cancelled.
        if (
            $item.Class -ne 26 -or
            $item.MeetingStatus -in @(5, 7) -or
            $item.ResponseStatus -notin @(1, 3)
        ) { continue }
        $joinUrl = $null
        $text = [string]$item.Body
        if ($text -match 'https://teams\.microsoft\.com/l/meetup-join/[^\s<>"'']+') { $joinUrl = $Matches[0] }
        $events += [pscustomobject]@{
            id = "local-" + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(([string]$item.EntryID + "|" + $item.Start.ToString("o"))))
            subject = if ([string]::IsNullOrWhiteSpace([string]$item.Subject)) { "Reunião sem título" } else { [string]$item.Subject }
            start = $item.Start.ToString("yyyy-MM-ddTHH:mm:ss")
            end = $item.End.ToString("yyyy-MM-ddTHH:mm:ss")
            joinUrl = $joinUrl
        }
    }

    $events | ConvertTo-Json -Depth 4 -Compress
}
finally {
    foreach ($comObject in @($items, $calendar, $namespace, $outlook)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) }
    }
}
