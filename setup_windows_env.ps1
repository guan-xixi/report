
param(
    [string]$ProjectRoot = $PSScriptRoot,
    [string]$PythonCmd = "python"
)

$ErrorActionPreference = "Stop"

Write-Host "[1/6] Project root: $ProjectRoot"
Set-Location $ProjectRoot

Write-Host "[2/6] Create .venv"

$Python = Get-Command $PythonCmd -ErrorAction SilentlyContinue

if (-not $Python) {
    throw "找不到 Python。请先安装 Python，并配置 PATH。"
}

& $Python.Source --version
& $Python.Source -m venv .venv

$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    throw "虚拟环境创建失败。"
}

Write-Host "[3/6] Upgrade pip"
& $VenvPython -m pip install --upgrade pip

Write-Host "[4/6] Install base dependencies"

$Requirements = Join-Path $ProjectRoot "requirements.windows.txt"

if (-not (Test-Path $Requirements)) {
    throw "找不到 requirements.windows.txt"
}

& $VenvPython -m pip install -r $Requirements

Write-Host "[5/6] Verify imports"
& $VenvPython -c "import openpyxl, yaml; print('base deps ok')"

Write-Host "[6/6] Optional ACP check"

try {
    & $VenvPython -c "import acp; print('acp ok')"
    if ($LASTEXITCODE -ne 0) {
        throw "ACP import failed"
    }
}
catch {
    Write-Warning "ACP 导入失败，需要确认 iCode 使用的 ACP SDK。"
}

Write-Host ""
Write-Host "Environment setup completed."
Write-Host "Activate: .\.venv\Scripts\Activate.ps1"
Write-Host "Run: .\.venv\Scripts\python.exe .\bench_fill_full_exe.py"

