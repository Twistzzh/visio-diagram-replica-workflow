# create_visio_via_proxy.ps1
# 通过 Visio COM 代理执行 JSON 图形计划（替代 create_visio_from_plan.ps1）
# 用法与 create_visio_from_plan.ps1 相同，但通过文件型 IPC 跨会话通信

param(
    [Parameter(Mandatory = $true)]
    [string]$PlanPath,

    [Parameter(Mandatory = $true)]
    [string]$OutVsdx,

    [string]$OutEmf = "",
    [switch]$Visible,
    [switch]$KeepVisioOpen
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $PlanPath)) {
    throw "找不到 JSON 图形计划文件：$PlanPath"
}

$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$outDir = Split-Path -Parent $OutVsdx
if ($outDir) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

$proxyDir = "D:\codex\projects\draw\scratch\visio_proxy"
$clientScript = "D:\codex\projects\draw\scripts\invoke_visio_proxy.ps1"

function Invoke-Proxy($action, $argsObj) {
    $argsJson = $argsObj | ConvertTo-Json -Depth 10 -Compress
    $result = & $clientScript -Action $action -ArgsJson $argsJson -ProxyDir $proxyDir
    if (-not $result.success) { throw "代理命令失败 ($action): $($result.error)" }
    return $result
}

# 初始化
Invoke-Proxy "new_document" @{ name = if ($plan.page.name) { [string]$plan.page.name } else { "Visio 复刻图" } }

$scale = if ($plan.page.scalePxPerInch) { [double]$plan.page.scalePxPerInch } else { 100.0 }
$pageH = [double]($plan.page.heightPx)

Invoke-Proxy "set_page" @{
    widthPx = [double]$plan.page.widthPx
    heightPx = $pageH
    scalePxPerInch = $scale
}

# 绘制图形
foreach ($shape in @($plan.shapes)) {
    $type = "$($shape.type)".ToLowerInvariant()
    $style = if ($shape.style) { $shape.style } else { [pscustomobject]@{} }
    $text = if ($null -ne $shape.text) { [string]$shape.text } else { "" }

    $drawArgs = @{
        type = $type
        pageHeightPx = $pageH
        scalePxPerInch = $scale
    }

    if ($type -in @("rect", "oval", "text")) {
        $drawArgs.x1 = [double]$shape.x1
        $drawArgs.y1 = [double]$shape.y1
        $drawArgs.x2 = [double]$shape.x2
        $drawArgs.y2 = [double]$shape.y2
        $drawArgs.style = $style
        $drawArgs.text = $text
    } elseif ($type -eq "line") {
        $drawArgs.x1 = [double]$shape.x1
        $drawArgs.y1 = [double]$shape.y1
        $drawArgs.x2 = [double]$shape.x2
        $drawArgs.y2 = [double]$shape.y2
        $drawArgs.style = $style
        if ($shape.arrow) { $drawArgs.arrow = "$($shape.arrow)" }
    } elseif ($type -eq "polyline") {
        $drawArgs.points = @($shape.points)
        $drawArgs.style = $style
        if ($shape.arrow) { $drawArgs.arrow = "$($shape.arrow)" }
    } else {
        Write-Warning "跳过不支持的图形类型: $type"
        continue
    }

    Invoke-Proxy "draw_shape" $drawArgs
}

# 保存
Invoke-Proxy "save_as" @{ path = $OutVsdx }

# 导出
if ($OutEmf) {
    Invoke-Proxy "export_page" @{ path = $OutEmf }
}

# 如果不保持打开，关闭文档
if (-not $KeepVisioOpen) {
    Invoke-Proxy "close_document" @{}
}

[pscustomobject]@{
    vsdx = $OutVsdx
    emf = $OutEmf
    pages = 1
} | ConvertTo-Json -Depth 3
