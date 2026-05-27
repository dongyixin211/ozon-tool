param(
    [string]$ReleaseDir = ""
)

$ErrorActionPreference = "Stop"

$ExeBuild = $PSScriptRoot
$ToolRoot = Split-Path -Parent $ExeBuild
$DistDir = Join-Path $ExeBuild "dist"
$SpecFile = Join-Path $ExeBuild "OzonTool.spec"

if ($ReleaseDir) {
    $ReleaseDir = $ReleaseDir.Trim().Trim('"').TrimEnd('\')
    $ReleaseRoot = (Resolve-Path -LiteralPath $ReleaseDir).Path
} else {
    $ReleaseRoot = Join-Path (Split-Path -Parent $ToolRoot) "ozon-plsj"
}

function Assert-LastExitCode {
    param([string]$Step)
    if ($LASTEXITCODE -ne 0) {
        throw "$Step failed (exit code $LASTEXITCODE)."
    }
}

function Find-Python310Plus {
    $candidates = @()

    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        foreach ($tag in @("3.12", "3.11", "3.10")) {
            try {
                $exe = & py -$tag -c "import sys; print(sys.executable)" 2>$null
                if ($exe -and (Test-Path $exe)) {
                    $candidates += $exe.Trim()
                }
            } catch {
            }
        }
    }

    $fixedPaths = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe"
    )
    $candidates += $fixedPaths

    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) {
        $candidates += $cmd.Source
    }

    # Codex 内置 Python 放最后：其 pip 镜像常装不上 PyInstaller
    $candidates += "C:\Users\23393\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (-not $candidate -or -not (Test-Path $candidate)) {
            continue
        }
        try {
            $major = [int](& $candidate -c "import sys; print(sys.version_info.major)" 2>$null)
            $minor = [int](& $candidate -c "import sys; print(sys.version_info.minor)" 2>$null)
            if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 10)) {
                return $candidate
            }
        } catch {
        }
    }

    return $null
}

function Install-PythonPackages {
    param([string]$PythonExe)

    $packages = @("pip", "pyinstaller", "openpyxl", "pillow")
    $indexes = @(
        "https://pypi.org/simple",
        "https://pypi.tuna.tsinghua.edu.cn/simple"
    )

    foreach ($index in $indexes) {
        Write-Host "pip install (index: $index) ..."
        & $PythonExe -m pip install --upgrade @packages -i $index
        if ($LASTEXITCODE -eq 0) {
            return
        }
        Write-Host "WARN: pip install failed on $index" -ForegroundColor Yellow
    }

    throw "pip install failed on all indexes. Try: $PythonExe -m pip install pyinstaller -i https://pypi.org/simple"
}

function Test-PyInstallerReady {
    param([string]$PythonExe)
    & $PythonExe -c "import PyInstaller; print(PyInstaller.__version__)"
    Assert-LastExitCode "PyInstaller import check"
}

$PythonExe = Find-Python310Plus
if (-not $PythonExe) {
    throw @"
Python 3.10+ not found.
Install from https://www.python.org/downloads/ (check "Add to PATH"), then run this script again.
"@
}

Write-Host "Python:" $PythonExe
Write-Host "Release:" $ReleaseRoot
Write-Host ""

$distExeBefore = Join-Path $DistDir "OzonTool.exe"
$beforeTime = $null
if (Test-Path $distExeBefore) {
    $beforeTime = (Get-Item $distExeBefore).LastWriteTime
}

Install-PythonPackages -PythonExe $PythonExe
Test-PyInstallerReady -PythonExe $PythonExe

Push-Location $ExeBuild
try {
    Write-Host "PyInstaller build ..."
    & $PythonExe -m PyInstaller --noconfirm --clean $SpecFile
    Assert-LastExitCode "PyInstaller"
} finally {
    Pop-Location
}

$BuiltExe = Join-Path $DistDir "OzonTool.exe"
if (-not (Test-Path $BuiltExe)) {
    throw "Build failed, missing file: $BuiltExe"
}

$afterTime = (Get-Item $BuiltExe).LastWriteTime
if ($beforeTime -and $afterTime -le $beforeTime) {
    throw @"
Build output was not updated (still $($afterTime.ToString('yyyy-MM-dd HH:mm:ss'))).
PyInstaller may have failed silently. Delete dist\OzonTool.exe and rebuild.
"@
}

$Stamp = Get-Date -Format "yyyyMMdd_HHmm"
$ReleaseName = "OzonTool_$Stamp.exe"
New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null
$ReleasePath = Join-Path $ReleaseRoot $ReleaseName
Copy-Item -Path $BuiltExe -Destination $ReleasePath -Force

Write-Host ""
Write-Host "Build OK:" $BuiltExe
Write-Host "Built at:" $afterTime
Write-Host "Release:" $ReleasePath
