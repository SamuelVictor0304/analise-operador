param(
    [string]$RepoPath = (Resolve-Path "$PSScriptRoot\..").Path,
    [string]$FileName = "NOVA BASE RESULTADOS 2026.xlsm",
    [string]$Remote = "origin",
    [string]$Branch = "main",
    [int]$PollSeconds = 90,
    [string]$LogFile = "",
    [switch]$RunOnce,
    [switch]$QuietWhenClean
)

$ErrorActionPreference = "Stop"

# Evita crescimento desnecessario do log nas verificacoes em que nada mudou.
if (-not $PSBoundParameters.ContainsKey("QuietWhenClean")) {
    $QuietWhenClean = $true
}

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$stamp] $Message"
    Write-Host $line
    if (-not [string]::IsNullOrWhiteSpace($LogFile)) {
        Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    }
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
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
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
    param([string[]]$GitArgs)
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        $output = & git @GitArgs 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }
    if ($exitCode -ne 0) {
        throw "git $($GitArgs -join ' ') falhou:`n$output"
    }
    return $output
}

function Commit-Resultados {
    Set-Location -LiteralPath $RepoPath
    $targetPath = Join-Path $RepoPath $FileName

    $currentBranch = (& git branch --show-current).Trim()
    if ($LASTEXITCODE -ne 0 -or $currentBranch -ne $Branch) {
        throw "Branch atual '$currentBranch' difere da branch configurada '$Branch'."
    }

    Wait-FileStable -Path $targetPath

    $status = & git status --porcelain -- $FileName
    if ([string]::IsNullOrWhiteSpace($status)) {
        if (-not $QuietWhenClean) {
            Write-Log "Nenhuma alteracao pendente em '$FileName'."
        }
        return
    }

    Write-Log "Alteracao detectada em '$FileName'. Criando commit."
    Invoke-Git -GitArgs @("add", "--", $FileName) | Out-Null

    $staged = & git diff --cached --name-only -- $FileName
    if ([string]::IsNullOrWhiteSpace($staged)) {
        Write-Log "Arquivo nao gerou diferenca staged."
        return
    }

    $commitStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # --only garante que outros arquivos eventualmente staged nunca entrem no autocommit.
    Invoke-Git -GitArgs @("commit", "--only", "-m", "Atualiza base de resultados ($commitStamp)", "--", $FileName) | Out-Null

    try {
        Invoke-Git -GitArgs @("push", $Remote, $Branch) | Out-Null
        Write-Log "Commit enviado para $Remote/$Branch."
    }
    catch {
        Write-Log "Push falhou. Tentando rebase antes de reenviar."
        Invoke-Git -GitArgs @("pull", "--rebase", "--autostash", $Remote, $Branch) | Out-Null
        Invoke-Git -GitArgs @("push", $Remote, $Branch) | Out-Null
        Write-Log "Commit enviado para $Remote/$Branch apos rebase."
    }
}

Set-Location -LiteralPath $RepoPath
$defaultLogDir = Join-Path $RepoPath "logs"
if ([string]::IsNullOrWhiteSpace($LogFile)) {
    if (-not (Test-Path -LiteralPath $defaultLogDir)) {
        New-Item -ItemType Directory -Path $defaultLogDir | Out-Null
    }
    $LogFile = Join-Path $defaultLogDir "resultados-autocommit.log"
}

$watchPath = Join-Path $RepoPath $FileName

if ($RunOnce) {
    try {
        Commit-Resultados
        exit 0
    }
    catch {
        Write-Log "Erro na automacao: $($_.Exception.Message)"
        exit 1
    }
}

$mutexCreated = $false
$mutex = [System.Threading.Mutex]::new($true, "Local\AnaliseOperadoresAutoCommitResultados", [ref]$mutexCreated)
if (-not $mutexCreated) {
    Write-Log "Monitor de autocommit ja esta em execucao. Encerrando instancia duplicada."
    $mutex.Dispose()
    exit 0
}

Write-Log "Monitoramento ativo por verificacao a cada $PollSeconds segundos: '$watchPath'."

try {
    while ($true) {
        try {
            Commit-Resultados
        }
        catch {
            Write-Log "Erro na automacao: $($_.Exception.Message)"
        }
        Start-Sleep -Seconds $PollSeconds
    }
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
