param(
    [string]$RepoPath = (Resolve-Path "$PSScriptRoot\..").Path,
    [string]$FileName = "NOVA BASE RESULTADOS 2026.xlsm",
    [string]$Remote = "origin",
    [string]$Branch = "main",
    [int]$DebounceSeconds = 20
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$stamp] $Message"
}

function Wait-FileStable {
    param(
        [string]$Path,
        [int]$StableChecks = 3,
        [int]$DelaySeconds = 5
    )

    $lastLength = -1
    $lastWrite = $null
    $stable = 0

    while ($stable -lt $StableChecks) {
        if (-not (Test-Path -LiteralPath $Path)) {
            Start-Sleep -Seconds $DelaySeconds
            continue
        }

        $item = Get-Item -LiteralPath $Path
        $sameState = ($item.Length -eq $lastLength) -and ($item.LastWriteTimeUtc -eq $lastWrite)
        $canRead = $false

        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            $stream.Close()
            $canRead = $true
        }
        catch {
            $canRead = $false
        }

        if ($sameState -and $canRead) {
            $stable += 1
        }
        else {
            $stable = 0
            $lastLength = $item.Length
            $lastWrite = $item.LastWriteTimeUtc
        }

        Start-Sleep -Seconds $DelaySeconds
    }
}

function Invoke-Git {
    param([string[]]$Args)
    $output = & git @Args 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "git $($Args -join ' ') falhou:`n$output"
    }
    return $output
}

function Commit-Resultados {
    Set-Location -LiteralPath $RepoPath
    $targetPath = Join-Path $RepoPath $FileName

    Wait-FileStable -Path $targetPath

    $status = & git status --porcelain -- $FileName
    if ([string]::IsNullOrWhiteSpace($status)) {
        Write-Log "Nenhuma alteracao pendente em '$FileName'."
        return
    }

    Write-Log "Alteracao detectada em '$FileName'. Criando commit."
    Invoke-Git -Args @("add", "--", $FileName) | Out-Null

    $staged = & git diff --cached --name-only -- $FileName
    if ([string]::IsNullOrWhiteSpace($staged)) {
        Write-Log "Arquivo nao gerou diferenca staged."
        return
    }

    $commitStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Invoke-Git -Args @("commit", "-m", "Atualiza base de resultados ($commitStamp)") | Out-Null

    try {
        Invoke-Git -Args @("push", $Remote, $Branch) | Out-Null
        Write-Log "Commit enviado para $Remote/$Branch."
    }
    catch {
        Write-Log "Push falhou. Tentando rebase antes de reenviar."
        Invoke-Git -Args @("pull", "--rebase", $Remote, $Branch) | Out-Null
        Invoke-Git -Args @("push", $Remote, $Branch) | Out-Null
        Write-Log "Commit enviado para $Remote/$Branch apos rebase."
    }
}

Set-Location -LiteralPath $RepoPath
$watchPath = Join-Path $RepoPath $FileName
$watchDir = Split-Path -Parent $watchPath
$watchFile = Split-Path -Leaf $watchPath

Write-Log "Monitorando '$watchPath'. Pressione Ctrl+C para parar."

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchDir
$watcher.Filter = $watchFile
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$events = @("Changed", "Created", "Renamed")
foreach ($eventName in $events) {
    Register-ObjectEvent -InputObject $watcher -EventName $eventName -SourceIdentifier "ResultadosGit$eventName" | Out-Null
}

try {
    while ($true) {
        $event = Wait-Event -Timeout 5
        if ($null -eq $event) {
            continue
        }

        Remove-Event -EventIdentifier $event.EventIdentifier
        Start-Sleep -Seconds $DebounceSeconds

        Get-Event | Where-Object { $_.SourceIdentifier -like "ResultadosGit*" } | Remove-Event

        try {
            Commit-Resultados
        }
        catch {
            Write-Log "Erro na automacao: $($_.Exception.Message)"
        }
    }
}
finally {
    foreach ($eventName in $events) {
        Unregister-Event -SourceIdentifier "ResultadosGit$eventName" -ErrorAction SilentlyContinue
    }
    $watcher.Dispose()
}
