param(
    [switch]$SkipDependencyInstall
)

$ErrorActionPreference = 'Stop'
$sourceRoot = $PSScriptRoot
$projectRoot = Join-Path $sourceRoot 'officejs'
$runtimeRoot = Join-Path (Join-Path $env:USERPROFILE 'AIForExcel') 'officejs\runtime'
$nodePath = Join-Path $runtimeRoot 'node.exe'
$npmPath = Join-Path $runtimeRoot 'npm.cmd'

function Get-NodeRuntime {
    if (Test-Path -LiteralPath $nodePath) {
        $version = (& $nodePath --version).Trim()
        if ($version -match '^v24\.') { return }
        throw "Existing local Node.js runtime is $version; Node.js 24 is required."
    }

    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    Write-Host 'Downloading the official Node.js 24 Windows runtime...'
    $releases = Invoke-RestMethod -Uri 'https://nodejs.org/dist/index.json'
    $release = $releases | Where-Object {
        $_.version -match '^v24\.' -and $_.files -contains 'win-x64-zip'
    } | Select-Object -First 1
    if (-not $release) { throw 'Could not find a Node.js 24 Windows runtime.' }

    $archive = Join-Path $env:TEMP ("aiforexcel-node-" + $release.version + '.zip')
    $folder = Join-Path $env:TEMP ("aiforexcel-node-" + [guid]::NewGuid().ToString('N'))
    try {
        $archiveName = "node-{0}-win-x64.zip" -f $release.version
        $baseUrl = "https://nodejs.org/dist/{0}/" -f $release.version
        Invoke-WebRequest -Uri ($baseUrl + $archiveName) -OutFile $archive
        $checksums = (Invoke-WebRequest -Uri ($baseUrl + 'SHASUMS256.txt')).Content
        $expected = (($checksums -split "`n") | Where-Object { $_ -match "^([0-9a-f]{64})\s+\*?$([regex]::Escape($archiveName))\s*$" } | ForEach-Object { $Matches[1] } | Select-Object -First 1)
        if (-not $expected -or (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expected.ToLowerInvariant()) { throw 'Downloaded Node.js archive failed SHA-256 verification.' }
        Expand-Archive -LiteralPath $archive -DestinationPath $folder -Force
        $extracted = Get-ChildItem -LiteralPath $folder -Directory | Select-Object -First 1
        if (-not $extracted -or -not (Test-Path (Join-Path $extracted.FullName 'node.exe'))) { throw 'Downloaded Node.js archive is invalid.' }
        Copy-Item -Path (Join-Path $extracted.FullName '*') -Destination $runtimeRoot -Recurse -Force
    } finally {
        Remove-Item -LiteralPath $archive -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Get-NodeRuntime
if (-not (Test-Path -LiteralPath $npmPath)) { throw 'The downloaded Node.js runtime does not contain npm.' }
if (-not $SkipDependencyInstall) {
    Write-Host 'Installing project dependencies...'
    Push-Location $projectRoot
    try { & $npmPath ci --registry=https://registry.npmjs.org/; if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed (exit $LASTEXITCODE)." } }
    finally { Pop-Location }
}

Write-Host 'Building and registering AI for Excel...'
& (Join-Path $sourceRoot 'manage-local.ps1') -Action install -NodePath $nodePath -OutsideApp
if ($LASTEXITCODE -ne 0) { throw "AI for Excel installation failed (exit $LASTEXITCODE)." }
Write-Host 'Installation completed. Open Excel and add AI for Excel from the shared folder if it is not already visible.'
