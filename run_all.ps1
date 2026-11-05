param(
  [string]$RootDir = "D:\co-review-work",
  [string]$PythonCmd = "py -3",
  [int]$RunTimeoutSec = 600,
  [switch]$SkipPatchGeneration,
  [switch]$PatchBackup
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Assert-PathExists {
  param(
    [Parameter(Mandatory = $true)][string]$PathValue,
    [Parameter(Mandatory = $true)][string]$Label,
    [switch]$File
  )

  if ($File) {
    if (-not (Test-Path -LiteralPath $PathValue -PathType Leaf)) {
      throw "[$Label] file not found: $PathValue"
    }
  } else {
    if (-not (Test-Path -LiteralPath $PathValue -PathType Container)) {
      throw "[$Label] directory not found: $PathValue"
    }
  }
}

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent $ScriptDir

$IcodeExe = Join-Path $RootDir "iCode\icode.exe"
$BenchXlsx = Join-Path $RootDir "iCode\bench_info.xlsx"
$PkgDir = Join-Path $RootDir "co-review-pkgs"
$WorkBase = Join-Path $RootDir "work"
$AssetRoot = Join-Path $RootDir "co-review\co-review"
$RuntimeSrc = Join-Path $RootDir "co-review\runtime-src\scripts"

Write-Host "[1/8] Validate required paths"
Assert-PathExists -PathValue $RootDir -Label "RootDir"
Assert-PathExists -PathValue $IcodeExe -Label "CO_REVIEW_ICODE_EXE" -File
Assert-PathExists -PathValue $BenchXlsx -Label "CO_REVIEW_BENCH_XLSX" -File
Assert-PathExists -PathValue $PkgDir -Label "CO_REVIEW_PKG_DIR"
Assert-PathExists -PathValue $AssetRoot -Label "CO_REVIEW_ASSET_ROOT"
Assert-PathExists -PathValue $RuntimeSrc -Label "CO_REVIEW_RUNTIME_SRC"
Assert-PathExists -PathValue $ScriptDir -Label "co-review_bench"

if (-not (Test-Path -LiteralPath $WorkBase -PathType Container)) {
  New-Item -ItemType Directory -Path $WorkBase | Out-Null
}

$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SetupScript = Join-Path $ScriptDir "setup_windows_env.ps1"

Write-Host "[2/8] Ensure Python venv"
if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
  Write-Host "  .venv missing, running setup_windows_env.ps1"
  & powershell -ExecutionPolicy Bypass -File $SetupScript -ProjectRoot $ProjectRoot -PythonCmd $PythonCmd
}
Assert-PathExists -PathValue $VenvPython -Label "VenvPython" -File

Write-Host "[3/8] Export runtime environment"
$env:CO_REVIEW_ICODE_EXE = $IcodeExe
$env:CO_REVIEW_ASSET_ROOT = $AssetRoot
$env:CO_REVIEW_RUNTIME_SRC = $RuntimeSrc
$env:CO_REVIEW_BENCH_XLSX = $BenchXlsx
$env:CO_REVIEW_PKG_DIR = $PkgDir
$env:CO_REVIEW_WORK_BASE = $WorkBase
$env:CO_REVIEW_RUN_TIMEOUT_SEC = "$RunTimeoutSec"

Write-Host "[4/8] Show effective configuration"
Write-Host "  CO_REVIEW_ICODE_EXE=$($env:CO_REVIEW_ICODE_EXE)"
Write-Host "  CO_REVIEW_ASSET_ROOT=$($env:CO_REVIEW_ASSET_ROOT)"
Write-Host "  CO_REVIEW_RUNTIME_SRC=$($env:CO_REVIEW_RUNTIME_SRC)"
Write-Host "  CO_REVIEW_BENCH_XLSX=$($env:CO_REVIEW_BENCH_XLSX)"
Write-Host "  CO_REVIEW_PKG_DIR=$($env:CO_REVIEW_PKG_DIR)"
Write-Host "  CO_REVIEW_WORK_BASE=$($env:CO_REVIEW_WORK_BASE)"
Write-Host "  CO_REVIEW_RUN_TIMEOUT_SEC=$($env:CO_REVIEW_RUN_TIMEOUT_SEC)"

Write-Host "[5/8] Optional patch generation"
if (-not $SkipPatchGeneration) {
  $PatchScript = Join-Path $ScriptDir "generate_review_patches.py"
  Assert-PathExists -PathValue $PatchScript -Label "generate_review_patches.py" -File

  $patchArgs = @(
    $PatchScript,
    "--bench-script", (Join-Path $ScriptDir "bench_fill_full_exe.py"),
    "--bench-xlsx", $BenchXlsx,
    "--pkg-dir", $PkgDir
  )
  if ($PatchBackup) {
    $patchArgs += "--backup"
  }

  & $VenvPython @patchArgs
} else {
  Write-Host "  Skipped patch generation by request"
}

Write-Host "[6/8] Run bench_fill_full_exe.py"
$BenchScript = Join-Path $ScriptDir "bench_fill_full_exe.py"
Assert-PathExists -PathValue $BenchScript -Label "bench_fill_full_exe.py" -File
& $VenvPython $BenchScript

Write-Host "[7/8] Quick artifact check"
$BenchLog = Join-Path $WorkBase "bench_runs.jsonl"
if (Test-Path -LiteralPath $BenchLog -PathType Leaf) {
  Write-Host "  Found: $BenchLog"
} else {
  Write-Warning "bench_runs.jsonl not found under $WorkBase"
}

Write-Host "[8/8] Done"
Write-Host "  Workflow finished. Inspect output in: $WorkBase"
