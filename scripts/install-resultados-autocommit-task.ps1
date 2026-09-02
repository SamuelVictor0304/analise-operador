param(
    [string]$RepoPath = (Resolve-Path "$PSScriptRoot\..").Path,
    [string]$TaskName = "AnaliseOperadoresAutoCommitResultados",
    [int]$IntervalSeconds = 90
)

$ErrorActionPreference = "Stop"

$watchScript = Join-Path $RepoPath "scripts\watch-resultados-git.ps1"
if (-not (Test-Path -LiteralPath $watchScript)) {
    throw "Script nao encontrado: $watchScript"
}

$quotedScript = '"' + $watchScript + '"'
$quotedRepo = '"' + $RepoPath + '"'
$logDir = Join-Path $RepoPath "logs"
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
$logFile = Join-Path $logDir "resultados-autocommit.log"
$quotedLog = '"' + $logFile + '"'
$arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $quotedScript -RepoPath $quotedRepo -LogFile $quotedLog -RunOnce -QuietWhenClean"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arguments -WorkingDirectory $RepoPath
$periodicTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Seconds $IntervalSeconds)
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)

$startupDir = [Environment]::GetFolderPath("Startup")
$startupCmd = Join-Path $startupDir "$TaskName.cmd"

try {
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger @($periodicTrigger, $logonTrigger) -Settings $settings -Description "Verifica periodicamente a base de resultados e faz commit/push automatico quando ela for alterada." -Force | Out-Null

    if (Test-Path -LiteralPath $startupCmd) {
        Remove-Item -LiteralPath $startupCmd -Force
    }

    Start-ScheduledTask -TaskName $TaskName

    Write-Host "Tarefa instalada: $TaskName"
    Write-Host "Execucao: a cada $IntervalSeconds segundos, no logon e com 3 tentativas em caso de falha."
    Write-Host "A primeira verificacao foi iniciada agora."
}
catch {
    $pollSeconds = $IntervalSeconds
    $cmd = "@echo off`r`nstart `"`" /min powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File `"$watchScript`" -RepoPath `"$RepoPath`" -PollSeconds $pollSeconds -QuietWhenClean`r`n"
    Set-Content -LiteralPath $startupCmd -Value $cmd -Encoding ASCII

    $backgroundArguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watchScript`" -RepoPath `"$RepoPath`" -PollSeconds $pollSeconds -QuietWhenClean"
    Start-Process `
        -FilePath "powershell.exe" `
        -ArgumentList $backgroundArguments `
        -WorkingDirectory $RepoPath `
        -WindowStyle Hidden

    Write-Host "Nao foi possivel instalar no Agendador de Tarefas: $($_.Exception.Message)"
    Write-Host "Fallback instalado na inicializacao do Windows:"
    Write-Host $startupCmd
    Write-Host "Monitor por polling iniciado agora em segundo plano."
}
