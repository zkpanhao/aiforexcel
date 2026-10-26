param(
    [ValidateSet('install','start','status','rollback','catalog','verify')][string]$Action = 'status',
    [string]$NodePath,
    [switch]$OutsideApp,
    [string]$LogPath,
    [string]$ResultPath,
    [switch]$WindowsContext
)
$ErrorActionPreference = 'Stop'
if (-not $NodePath) {
    $installed = Join-Path $env:USERPROFILE 'AIForExcel\officejs\runtime\node.exe'
    if (Test-Path -LiteralPath $installed) { $NodePath = $installed }
    else { $NodePath = (Get-Command node.exe -ErrorAction Stop).Source }
}
$version = & $NodePath --version
if ($LASTEXITCODE -ne 0 -or $version -notmatch '^v24\.') { throw 'Node.js 24 is required. Supply -NodePath with a verified Node.js 24 executable.' }
if ($OutsideApp) {
    # ShellExecute/RunAs descendants can inherit the desktop app's virtual registry.
    # WMI creates an ordinary Windows process; verify the result from that context.
    $invocationId = [guid]::NewGuid().ToString('N')
    $log = Join-Path $PSScriptRoot ('officejs\work\local-' + $Action + '-' + $invocationId + '.log')
    $resultFile = Join-Path $PSScriptRoot ('officejs\work\local-' + $Action + '-' + $invocationId + '.json')
    New-Item -ItemType Directory -Path (Split-Path $log) -Force | Out-Null
    $systemPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments = '"' + $systemPowerShell + '" -NoProfile -File "' + $PSCommandPath + '" -Action ' + $Action + ' -NodePath "' + $NodePath + '" -LogPath "' + $log + '" -ResultPath "' + $resultFile + '" -WindowsContext'
    $created = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine=$arguments;CurrentDirectory=$PSScriptRoot}
    if ($created.ReturnValue -ne 0) { throw "Windows process creation failed ($($created.ReturnValue))." }
    $deadline = [DateTime]::UtcNow.AddMinutes(5)
    while (-not (Test-Path -LiteralPath $resultFile)) {
        if ([DateTime]::UtcNow -gt $deadline) { throw "Configuration is still pending. Inspect $log before retrying." }
        Start-Sleep -Milliseconds 250
    }
    $result = Get-Content -LiteralPath $resultFile -Raw | ConvertFrom-Json
    Get-Content -LiteralPath $log -ErrorAction SilentlyContinue
    if ($result.ExitCode -ne 0) { throw "Windows-context $Action failed (exit $($result.ExitCode)). See $log" }
    return
}
if ($WindowsContext -and $Action -eq 'install') {
    $administrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $administrator) {
        # Elevate only after leaving the app context, so registry writes remain real.
        $arguments = '-NoProfile -File "' + $PSCommandPath + '" -Action ' + $Action + ' -NodePath "' + $NodePath + '" -LogPath "' + $LogPath + '" -ResultPath "' + $ResultPath + '" -WindowsContext'
        try {
            $systemPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $elevated = Start-Process -FilePath $systemPowerShell -Verb RunAs -WindowStyle Hidden -ArgumentList $arguments -PassThru
            $elevated.WaitForExit()
        } catch {
            [pscustomobject]@{ExitCode=1;Error=$_.Exception.Message;Action=$Action} | ConvertTo-Json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
        }
        return
    }
}
if ($LogPath) { Start-Transcript -Path $LogPath -Force | Out-Null }
$operationExitCode = 0
Push-Location (Join-Path $PSScriptRoot 'officejs')
try {
    if ($Action -eq 'verify') { & $NodePath (Join-Path $PSScriptRoot 'officejs\scripts\check-native.cjs') }
    else { & $NodePath (Join-Path $PSScriptRoot 'officejs\scripts\install-local.cjs') $Action }
    if ($LASTEXITCODE -ne 0) { throw "Local $Action failed (exit $LASTEXITCODE)." }
} catch {
    $operationExitCode = 1
    Write-Error $_ -ErrorAction Continue
} finally {
    Pop-Location
    if ($LogPath) { Stop-Transcript | Out-Null }
    if ($ResultPath) {
        $temporaryResult = $ResultPath + '.tmp'
        [pscustomobject]@{ExitCode=$operationExitCode;User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Action=$Action} | ConvertTo-Json | Set-Content -LiteralPath $temporaryResult -Encoding UTF8
        Move-Item -LiteralPath $temporaryResult -Destination $ResultPath
    }
}
if ($operationExitCode -ne 0) { throw "Local $Action failed. See the operation log." }
