$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

# Keep the local production server alive for the scheduled task session.
& node.exe tools\local-app-server.mjs
