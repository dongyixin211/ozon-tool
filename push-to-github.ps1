# Create dongyixin211/ozon-tool on GitHub and push. Run once: gh auth login

$ErrorActionPreference = "Stop"
$gh = "C:\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path $gh)) {
    $gh = "gh"
}

Set-Location $PSScriptRoot

& $gh auth status 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "GitHub login required. Opening browser..."
    & $gh auth login -h github.com -p https -w --skip-ssh-key
}

$remoteUrl = "https://github.com/dongyixin211/ozon-tool.git"
$hasOrigin = git remote get-url origin 2>$null
if (-not $hasOrigin) {
    git remote add origin $remoteUrl
} else {
    git remote set-url origin $remoteUrl
}

$repoExists = $false
try {
    & $gh repo view dongyixin211/ozon-tool 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $repoExists = $true }
} catch {
    $repoExists = $false
}
if (-not $repoExists) {
    Write-Host "Creating repo dongyixin211/ozon-tool ..."
    & $gh repo create ozon-tool --public --description "Ozon listing and image tool" --source=. --remote=origin --push
} else {
    git branch -M main 2>$null
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        git push -u origin master
    }
}

Write-Host "Done: https://github.com/dongyixin211/ozon-tool"
