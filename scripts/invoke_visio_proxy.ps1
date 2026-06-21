# invoke_visio_proxy.ps1
# 客户端脚本 - 在 Codex 沙箱用户中运行
# 通过文件型 IPC 向 Visio COM 代理发送命令并等待结果

param(
    [Parameter(Mandatory = $true)]
    [string]$Action,

    [string]$ArgsJson = "{}",

    [string]$ProxyDir = "D:\codex\projects\draw\scratch\visio_proxy",

    [int]$TimeoutSeconds = 120,

    [int]$PollIntervalMs = 200
)

$ErrorActionPreference = "Stop"

$cmdDir = Join-Path $ProxyDir "commands"
$rspDir = Join-Path $ProxyDir "responses"
$lockFile = Join-Path $ProxyDir "server.lock"
$statusFile = Join-Path $ProxyDir "server.status"

# 检查代理是否存活
if (-not (Test-Path -LiteralPath $lockFile)) {
    throw "Visio COM 代理未运行！请先让 ASUS 用户运行 start_visio_proxy.bat"
}

# 生成唯一命令 ID
$cmdId = [Guid]::NewGuid().ToString("N")
$cmdFile = Join-Path $cmdDir "$cmdId.json"
$rspFile = Join-Path $rspDir "$cmdId.json"

# 解析参数
$argsObj = if ($ArgsJson -ne "{}") { $ArgsJson | ConvertFrom-Json } else { @{} }

# 写入命令
$cmd = @{
    action = $Action
    args = $argsObj
}
$cmd | ConvertTo-Json -Depth 10 | Set-Content -Path $cmdFile -Encoding UTF8

# 等待响应
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $rspFile) {
        $result = Get-Content -LiteralPath $rspFile -Raw -Encoding UTF8 | ConvertFrom-Json
        Remove-Item -LiteralPath $rspFile -Force -ErrorAction SilentlyContinue
        if (-not $result.success) {
            throw "Visio 命令失败: $($result.error)"
        }
        return $result
    }
    Start-Sleep -Milliseconds $PollIntervalMs
}

throw "Visio 命令超时 ($TimeoutSeconds 秒): $Action"
