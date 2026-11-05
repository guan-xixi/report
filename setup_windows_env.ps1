param(
  [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
  [string]$PythonCmd = "py -3"
)

$ErrorActionPreference = "Stop"

Write-Host "[1/6] Project root: $ProjectRoot"
Set-Location $ProjectRoot

Write-Host "[2/6] Create .venv"
& cmd /c "$PythonCmd -m venv .venv"

Write-Host "[3/6] Upgrade pip"
& .\.venv\Scripts\python.exe -m pip install --upgrade pip

Write-Host "[4/6] Install base dependencies"
& .\.venv\Scripts\python.exe -m pip install -r .\co-review_bench\requirements.windows.txt

Write-Host "[5/6] Verify imports"
& .\.venv\Scripts\python.exe -c "import openpyxl, yaml; print('base deps ok')"

Write-Host "[6/6] Optional ACP check"
try {
  & .\.venv\Scripts\python.exe -c "import acp; print('acp ok')"
} catch {
  Write-Warning "acp import failed in this venv. If your iCode runtime bundles ACP, you can still run with that Python. Otherwise install ACP package used by your iCode distribution."
}

Write-Host "Done. Activate with: .\.venv\Scripts\Activate.ps1"
Write-Host "Then run: python .\co-review_bench\bench_fill_full_exe.py"
