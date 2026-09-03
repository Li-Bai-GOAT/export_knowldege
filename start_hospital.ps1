param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'export_knowldege\config.json'),
    [switch]$Headless,
    [switch]$CreateBge,
    [switch]$UploadBge,
    [switch]$ExportOnly,
    [switch]$CreateOnly,
    [switch]$UploadOnly,
    [switch]$BgeOnly,
    [switch]$PauseOnExit
)

$ErrorActionPreference = 'Stop'

$repoRoot = $PSScriptRoot
$scriptRoot = Join-Path $repoRoot 'export_knowldege'
$pythonPath = Join-Path $repoRoot '.venv\Scripts\python.exe'
$entryScript = Join-Path $scriptRoot 'knowledge_query.py'
$resolvedConfig = [System.IO.Path]::GetFullPath($ConfigPath)

function Stop-Run([string]$Message) {
    Write-Host "[ERROR] $Message" -ForegroundColor Red
    if ($PauseOnExit) { Read-Host 'Press Enter to exit' | Out-Null }
    exit 1
}

try {
    Write-Host '=== Heren knowledge export: one hospital / one account ===' -ForegroundColor Cyan
    Write-Host "Repository: $repoRoot"
    Write-Host "Config:     $resolvedConfig"

    if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
        Stop-Run "Virtual environment not found: $pythonPath. Run: python -m venv .venv"
    }
    if (-not (Test-Path -LiteralPath $entryScript -PathType Leaf)) {
        Stop-Run "Entry script not found: $entryScript"
    }
    if (-not (Test-Path -LiteralPath $resolvedConfig -PathType Leaf)) {
        Stop-Run "config.json not found. Copy config.example.json to config.json and fill in one hospital account."
    }

    Push-Location $scriptRoot
    try {
        Write-Host '[1/3] Checking Python dependencies...' -ForegroundColor Yellow
        & $pythonPath -c 'import playwright, requests, openpyxl'
        if ($LASTEXITCODE -ne 0) {
            Write-Host '[INFO] Installing dependencies...' -ForegroundColor Yellow
            & $pythonPath -m pip install --upgrade pip
            & $pythonPath -m pip install 'playwright>=1.40' 'requests>=2.31' 'openpyxl>=3.1'
            if ($LASTEXITCODE -ne 0) { Stop-Run 'Dependency installation failed.' }
        }

        Write-Host '[2/3] Starting the complete hospital workflow...' -ForegroundColor Yellow
        $arguments = @('.\knowledge_query.py', '--config', $resolvedConfig)
        if ($Headless) { $arguments += '--headless' }
        if ($CreateBge) { $arguments += '--create-bge' }
        if ($UploadBge) { $arguments += '--upload-bge' }
        if ($ExportOnly) { $arguments += '--export-only' }
        if ($CreateOnly) { $arguments += '--create-only' }
        if ($UploadOnly) { $arguments += '--upload-only' }
        if ($BgeOnly) { $arguments += '--bge-only' }

        Write-Host "Command: $pythonPath $($arguments -join ' ')" -ForegroundColor DarkGray
        & $pythonPath @arguments
        $exitCode = $LASTEXITCODE

        Write-Host '[3/3] Workflow finished.' -ForegroundColor Yellow
        if ($exitCode -eq 0) {
            Write-Host 'SUCCESS: check output, logs, and _metadata snapshots.' -ForegroundColor Green
        } else {
            Write-Host "FAILED: knowledge_query.py returned exit code $exitCode. Check the log for details." -ForegroundColor Red
        }
    }
    finally {
        Pop-Location
    }

    if ($PauseOnExit) { Read-Host 'Press Enter to exit' | Out-Null }
    exit $exitCode
}
catch {
    Stop-Run $_.Exception.Message
}
