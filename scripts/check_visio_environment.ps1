# check_visio_environment.ps1
# Check Visio environment: auto-detect sandbox/interactive user, check COM or proxy

param(
    [switch]$TryCom,
    [switch]$KeepVisioOpen
)

$ErrorActionPreference = "Stop"

$currentUser = [Environment]::UserName
$isSandbox = $currentUser -like "*Sandbox*" -or $currentUser -like "*Offline*" -or $currentUser -eq "CodexSandboxOffline"

$scriptDir = Split-Path -Parent $PSCommandPath
$skillDir = Split-Path -Parent $scriptDir

$candidateProxyDirs = @(
    "D:\codex\projects\draw\scratch\visio_proxy"
)
$candidateProxyDirs += $(Join-Path $env:TEMP "visio_proxy")
$candidateProxyDirs += $(Join-Path $skillDir "scratch\visio_proxy")

function Find-VisioExe {
    $cmd = Get-Command visio -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $roots = @("C:/Program Files", "C:/Program Files (x86)")
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $hit = Get-ChildItem -LiteralPath $root -Recurse -Filter VISIO.EXE -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName
        if ($hit) { return $hit }
    }
    return $null
}

function Find-Tool($names) {
    foreach ($name in $names) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    return $null
}

function Test-ProxyAlive($proxyDir) {
    $lockFile = Join-Path $proxyDir "server.lock"
    $statusFile = Join-Path $proxyDir "server.status"
    if (-not (Test-Path -LiteralPath $lockFile)) { return $null }
    $status = $null
    if (Test-Path -LiteralPath $statusFile) {
        try { $status = Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    }
    return [pscustomobject]@{
        proxyDir = $proxyDir
        lockExists = $true
        status = $status
    }
}

$result = [ordered]@{
    isSandbox = $isSandbox
    currentUser = $currentUser
    visioExe = Find-VisioExe
    soffice = Find-Tool @("soffice", "libreoffice")
    pdftoppm = Find-Tool @("pdftoppm")
    comStartup = "未测试"
    proxyAvailable = $null
    activeVisio = @()
}

$result.activeVisio = @(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
        id = $_.Id
        sessionId = $_.SessionId
        title = $_.MainWindowTitle
        startTime = $_.StartTime
    }
})

foreach ($d in $candidateProxyDirs) {
    if (Test-Path -LiteralPath $d) {
        $proxyInfo = Test-ProxyAlive $d
        if ($proxyInfo) {
            $result.proxyAvailable = $proxyInfo
            break
        }
    }
}

if ($TryCom -and -not $isSandbox) {
    $visio = $null
    try {
        for ($i = 1; $i -le 6; $i++) {
            try {
                $visio = New-Object -ComObject Visio.Application
                Start-Sleep -Milliseconds 800
                $visio.Visible = $true
                $visio.AlertResponse = 7
                $result.comStartup = "成功"
                break
            } catch {
                if ($i -eq 6) { throw }
                Start-Sleep -Milliseconds (700 * $i)
            }
        }
    } catch {
        $result.comStartup = "失败：$($_.Exception.Message)"
    } finally {
        if ($visio -ne $null -and -not $KeepVisioOpen) {
            try { $visio.Quit() } catch { }
        }
    }
} elseif ($TryCom -and $isSandbox) {
    $result.comStartup = "跳过（沙箱用户，不支持直连 COM）"
}

$result | ConvertTo-Json -Depth 5
