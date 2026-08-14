$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$appUrl = if ($env:WORK_ORGANIZER_URL) {
  $env:WORK_ORGANIZER_URL
} else {
  "http://localhost:3000"
}
$bridgeUrl = "http://127.0.0.1:47831/health"
$logDirectory = Join-Path $env:LOCALAPPDATA "Work Organizer\Logs"

function Test-App {
  try {
    $uri = [Uri]$appUrl
    $testUrl = if ($uri.Host -in @("localhost", "127.0.0.1", "::1")) {
      "http://127.0.0.1:{0}/__local/health" -f $uri.Port
    } else {
      $appUrl
    }
    $response = Invoke-WebRequest -Uri $testUrl -UseBasicParsing -TimeoutSec 2
    return $response.StatusCode -ge 200 -and $response.StatusCode -lt 500
  } catch {
    return $false
  }
}

function Test-Bridge {
  try {
    $response = Invoke-WebRequest -Uri $bridgeUrl -UseBasicParsing -TimeoutSec 2
    return $response.StatusCode -eq 200
  } catch {
    # Older bridge versions do not expose /health. Accept an existing listener
    # on the dedicated loopback port so upgrading does not create a conflict.
    $client = New-Object System.Net.Sockets.TcpClient
    try {
      $client.Connect("127.0.0.1", 47831)
      return $client.Connected
    } catch {
      return $false
    } finally {
      $client.Dispose()
    }
  }
}

function Show-LaunchError([string]$message) {
  Add-Type -AssemblyName PresentationFramework
  [System.Windows.MessageBox]::Show(
    $message,
    "Work Organizer",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Error
  ) | Out-Null
}

try {
  $appHost = ([Uri]$appUrl).Host
  $isLocalApp = $appHost -in @("localhost", "127.0.0.1", "::1")
  if ($isLocalApp -and -not (Test-App)) {
    New-Item -ItemType Directory -Force -Path $logDirectory | Out-Null
    Start-Process -FilePath "node.exe" `
      -ArgumentList @((Join-Path $projectRoot "tools\local-app-server.mjs")) `
      -WorkingDirectory $projectRoot `
      -WindowStyle Hidden

    $appReady = $false
    for ($attempt = 0; $attempt -lt 90; $attempt++) {
      Start-Sleep -Milliseconds 500
      if (Test-App) {
        $appReady = $true
        break
      }
    }
    if (-not $appReady) {
      throw "Não foi possível iniciar o Work Organizer local."
    }
  }

  if (-not (Test-Bridge)) {
    $node = (Get-Command node.exe -ErrorAction Stop).Source
    New-Item -ItemType Directory -Force -Path $logDirectory | Out-Null
    Start-Process -FilePath $node `
      -ArgumentList @((Join-Path $projectRoot "tools\outlook-bridge.mjs")) `
      -WorkingDirectory $projectRoot `
      -WindowStyle Hidden

    $bridgeReady = $false
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
      Start-Sleep -Milliseconds 250
      if (Test-Bridge) {
        $bridgeReady = $true
        break
      }
    }
    if (-not $bridgeReady) {
      throw "Não foi possível iniciar a ponte do Outlook."
    }
  }

  $browserCandidates = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  $browser = $browserCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
  if ($browser) {
    Start-Process -FilePath $browser -ArgumentList @("--app=$appUrl", "--start-maximized")
  } else {
    Start-Process $appUrl
  }
} catch {
  Show-LaunchError $_.Exception.Message
  exit 1
}
