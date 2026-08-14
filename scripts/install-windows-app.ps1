$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$launcher = Join-Path $PSScriptRoot "launch-windows-app.ps1"
$desktop = [Environment]::GetFolderPath("Desktop")
$startMenu = Join-Path ([Environment]::GetFolderPath("Programs")) "Work Organizer.lnk"
$desktopShortcut = Join-Path $desktop "Work Organizer.lnk"
$powerShell = Join-Path $PSHOME "powershell.exe"
$browserCandidates = @(
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
$icon = $browserCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
$shell = New-Object -ComObject WScript.Shell

foreach ($shortcutPath in @($desktopShortcut, $startMenu)) {
  $shortcut = $shell.CreateShortcut($shortcutPath)
  $shortcut.TargetPath = $powerShell
  $shortcut.Arguments = "-NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$launcher`""
  $shortcut.WorkingDirectory = $projectRoot
  $shortcut.Description = "Work Organizer"
  if ($icon) {
    $shortcut.IconLocation = "$icon,0"
  }
  $shortcut.Save()
}

Write-Host "Work Organizer instalado no Ambiente de Trabalho e no menu Iniciar."
Write-Host "Podes agora abrir a aplicação pelo novo ícone."
