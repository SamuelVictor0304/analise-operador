param(
    [string]$RepoPath = (Resolve-Path "$PSScriptRoot\..").Path,
    [string]$TaskName = "AnaliseOperadoresAutoCommitResultados"
)

$ErrorActionPreference = "Stop"

$watchScript = Join-Path $RepoPath "scripts\watch-resultados-git.ps1"
if (-not (Test-Path -LiteralPath $watchScript)) {
    throw "Script nao encontrado: $watchScript"
}

$quotedScript = '"' + $watchScript + '"'
$quotedRepo = '"' + $RepoPath + '"'
$arguments = "-NoProfile -ExecutionPolicy Bypass -File $quotedScript -RepoPath $quotedRepo"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arguments -WorkingDirectory $RepoPath
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Monitora a base de resultados e faz commit/push automatico quando ela for alterada." -Force | Out-Null

Write-Host "Tarefa instalada: $TaskName"
Write-Host "Para iniciar agora:"
Write-Host "Start-ScheduledTask -TaskName `"$TaskName`""
