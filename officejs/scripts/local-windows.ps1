param(
    [Parameter(Mandatory=$true)][string]$InstallRoot,
    [Parameter(Mandatory=$true)][ValidateSet('Migrate','Register','Catalog','Share')][string]$Action,
    [string]$SourceRoot
)
$ErrorActionPreference = 'Stop'
trap {
    if (Test-Path -LiteralPath (Join-Path $InstallRoot 'logs')) {
        ($_ | Out-String) | Add-Content -LiteralPath (Join-Path $InstallRoot 'logs\catalog-error.log')
    }
    Write-Error $_ -ErrorAction Continue
    exit 1
}
$expectedRoot = [IO.Path]::GetFullPath((Join-Path $env:USERPROFILE 'AIForExcel\officejs'))
if ([IO.Path]::GetFullPath($InstallRoot) -ne $expectedRoot) { throw 'Unexpected installation directory.' }
$manifestId = 'e6e1a25b-a2dc-4c80-9177-6c1e9e0e9d61'
if ($Action -eq 'Migrate') {
    $listeners = @(Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue)
    foreach ($ownerId in @($listeners.OwningProcess | Select-Object -Unique)) {
        $owner = Get-CimInstance Win32_Process -Filter "ProcessId=$ownerId"
        if (-not $owner -or -not $owner.ExecutablePath -or -not $owner.CommandLine) { throw "Cannot verify the owner of port 3000 ($ownerId). Run installation with administrator confirmation; no process was stopped." }
        $installedNode = Join-Path $InstallRoot 'runtime\node.exe'
        $isInstalled = $owner.ExecutablePath -eq $installedNode -and $owner.CommandLine.Contains((Join-Path $InstallRoot 'local-server.cjs'))
        $legacyRoot = Join-Path $env:LOCALAPPDATA 'AIForExcel\officejs'
        $isLegacy = $owner.ExecutablePath -eq (Join-Path $legacyRoot 'runtime\node.exe') -and $owner.CommandLine.Contains((Join-Path $legacyRoot 'local-server.cjs'))
        $isDevelopment = $SourceRoot -and $owner.Name -eq 'node.exe' -and $owner.CommandLine.Contains((Join-Path $SourceRoot 'node_modules\webpack-cli\bin\cli.js')) -and $owner.CommandLine.Contains('serve')
        if (-not ($isInstalled -or $isDevelopment -or $isLegacy)) { throw "Port 3000 belongs to an unrelated process ($ownerId); no process was stopped." }
        Stop-Process -Id $ownerId -ErrorAction Stop
    }
    return
}
if ($Action -eq 'Register') {
    $developer = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
    New-Item -Path $developer -Force | Out-Null
    New-ItemProperty -Path $developer -Name $manifestId -Value (Join-Path $InstallRoot 'catalog\manifest.xml') -PropertyType String -Force | Out-Null
    $launcher = Join-Path $InstallRoot 'start-hidden.vbs'
    $exe = (Join-Path $InstallRoot 'runtime\node.exe').Replace('"','""')
    $server = (Join-Path $InstallRoot 'local-server.cjs').Replace('"','""')
    $escapedRoot = $InstallRoot.Replace('"','""')
    $vbs = 'CreateObject("WScript.Shell").Run """' + $exe + '"" ""' + $server + '"" ""' + $escapedRoot + '""", 0, False'
    Set-Content -LiteralPath $launcher -Value $vbs -Encoding Unicode
    $run = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    New-Item -Path $run -Force | Out-Null
    New-ItemProperty -Path $run -Name 'AIForExcel' -Value ('wscript.exe "' + $launcher + '"') -PropertyType String -Force | Out-Null
    Write-Output 'Local manifest registered; hidden host will start at user sign-in.'
    return
}
# A read-only share for the current Windows user, containing only public manifest XML.
# Creating the SMB share requires an administrator; do not alter firewall or Office policies.
$catalog = Join-Path $InstallRoot 'catalog'
$shareName = 'AIForExcel-' + $env:USERNAME
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$share = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if ($share -and $share.Path -ne $catalog) { throw 'An unrelated share already uses the intended name.' }
if (-not $share) {
    $administrator = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $administrator) {
        # Elevate only share creation. Trust registration remains in the original user's HKCU.
        $elevatedArgs = '-NoProfile -File "' + $PSCommandPath + '" -InstallRoot "' + $InstallRoot + '" -Action Share'
        $elevated = Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -ArgumentList $elevatedArgs -Wait -PassThru
        if ($elevated.ExitCode -ne 0) { throw 'Administrator share creation failed or was cancelled.' }
        $share = Get-SmbShare -Name $shareName -ErrorAction Stop
        if ($share.Path -ne $catalog) { throw 'Administrator created an unexpected share.' }
    } else {
        New-SmbShare -Name $shareName -Path $catalog -ReadAccess $identity -FolderEnumerationMode AccessBased | Out-Null
    }
}
if ($Action -eq 'Share') { return }
$catalogId = '{B3596401-2405-4E0D-8DA6-A1F0EC41E830}'
$registry = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\' + $catalogId
New-Item -Path $registry -Force | Out-Null
New-ItemProperty -Path $registry -Name Id -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registry -Name Url -Value ('\\localhost\' + $shareName) -PropertyType String -Force | Out-Null
New-ItemProperty -Path $registry -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
Write-Output ('Catalog configured: \\localhost\' + $shareName + '. Restart Excel after saving your work, then add from Shared Folder.')
