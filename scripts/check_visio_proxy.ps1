# check_visio_proxy.ps1
# 检查 Visio COM 代理状态和环境

$ErrorActionPreference = "Stop"

$proxyDir = "D:\codex\projects\draw\scratch\visio_proxy"
$lockFile = Join-Path $proxyDir "server.lock"
$statusFile = Join-Path $proxyDir "server.status"

$result = [ordered]@{
    proxyRunning = $false
    proxyPid = $null
    proxyStatus = $null
    visioExe = $null
    visioProcesses = @()
    sandboxUser = [Environment]::UserName
    interactiveUser = $null
}

# 检查 Visio 可执行文件
$cmd = Get-Command visio -ErrorAction SilentlyContinue
if ($cmd) { $result.visioExe = $cmd.Source } else {
    $roots = @("C:/Program Files", "C:/Program Files (x86)")
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        $hit = Get-ChildItem -LiteralPath $root -Recurse -Filter VISIO.EXE -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName
        if ($hit) { $result.visioExe = $hit; break }
    }
}

# 检查代理状态
if (Test-Path -LiteralPath $lockFile) {
    $result.proxyRunning = $true
    $lockContent = Get-Content -LiteralPath $lockFile -Raw
    if ($lockContent -match "pid=(\d+)") { $result.proxyPid = [int]$Matches[1] }
}

if (Test-Path -LiteralPath $statusFile) {
    $result.proxyStatus = Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8 | ConvertFrom-Json
}

# 检查当前 Visio 进程
$result.visioProcesses = @(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object {
    [ordered]@{
        id = $_.Id
        sessionId = $_.SessionId
        startTime = $_.StartTime
    }
})

$result | ConvertTo-Json -Depth 5
