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

try {
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Monitora a base de resultados e faz commit/push automatico quando ela for alterada." -Force | Out-Null

    Write-Host "Tarefa instalada: $TaskName"
    Write-Host "Para iniciar agora:"
    Write-Host "Start-ScheduledTask -TaskName `"$TaskName`""
}
catch {
    $startupDir = [Environment]::GetFolderPath("Startup")
    $startupCmd = Join-Path $startupDir "$TaskName.cmd"
    $cmd = "@echo off`r`nstart `"`" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$watchScript`" -RepoPath `"$RepoPath`"`r`n"
    Set-Content -LiteralPath $startupCmd -Value $cmd -Encoding ASCII

    Write-Host "Nao foi possivel instalar no Agendador de Tarefas: $($_.Exception.Message)"
    Write-Host "Fallback instalado na inicializacao do Windows:"
    Write-Host $startupCmd
    Write-Host "Para iniciar agora:"
    Write-Host "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$watchScript`" -RepoPath `"$RepoPath`""
}
