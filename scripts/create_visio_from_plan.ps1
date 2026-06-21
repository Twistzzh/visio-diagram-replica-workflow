# create_visio_from_plan.ps1
# Generate .vsdx from JSON plan
# Auto-route: non-sandbox -> direct COM; sandbox -> Visio COM proxy

param(
    [Parameter(Mandatory = $true)]
    [string]$PlanPath,

    [Parameter(Mandatory = $true)]
    [string]$OutVsdx,

    [Alias("OutPreview")]
    [string]$OutEmf = "",
    [switch]$Visible,
    [switch]$KeepVisioOpen,
    [switch]$PreferDirect
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $PlanPath)) {
    throw "找不到 JSON 图形计划文件：$PlanPath"
}

$currentUser = [Environment]::UserName
$isSandbox = $currentUser -like "*Sandbox*" -or $currentUser -like "*Offline*" -or $currentUser -eq "CodexSandboxOffline"

$scriptDir = Split-Path -Parent $PSCommandPath

if ($isSandbox -and -not $PreferDirect) {
    Write-Host "检测到沙箱用户 ($currentUser)，通过 Visio COM 代理执行..."

    $proxyClient = Join-Path $scriptDir "create_visio_via_proxy.ps1"
    if (-not (Test-Path -LiteralPath $proxyClient)) {
        throw "沙箱模式需要代理脚本，但找不到: $proxyClient"
    }

    $proxyParams = @{
        PlanPath = $PlanPath
        OutVsdx = $OutVsdx
    }
    if ($OutEmf) { $proxyParams.OutEmf = $OutEmf }
    if ($KeepVisioOpen) { $proxyParams.KeepVisioOpen = $true }

    & $proxyClient @proxyParams
    exit $LASTEXITCODE
}

# ---- Non-sandbox: direct COM ----
if ($isSandbox -and $PreferDirect) {
    Write-Host "检测到沙箱用户 ($currentUser)，但已指定 -PreferDirect，尝试直接通过 Visio COM 执行..."
}

$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$outDir = Split-Path -Parent $OutVsdx
if ($outDir) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
if ($OutEmf -and (Split-Path -Parent $OutEmf)) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutEmf) | Out-Null
}

$script:page = $null
$script:pageHeightPx = [double]($plan.page.heightPx)
$script:scale = if ($plan.page.scalePxPerInch) { [double]$plan.page.scalePxPerInch } else { 100.0 }

function X([double]$px) { $px / $script:scale }
function Y([double]$py) { ($script:pageHeightPx - $py) / $script:scale }
function Cell($shape, [string]$cell, [string]$formula) { try { $shape.CellsU($cell).FormulaU = $formula } catch { } }

function Style-Shape($shape, $style) {
    $fill = if ($style.fill) { $style.fill } else { "RGB(255,255,255)" }
    $line = if ($style.line) { $style.line } else { "RGB(35,35,35)" }
    $weight = if ($style.weight) { [double]$style.weight } else { 1.0 }
    $dash = [bool]$style.dash
    $noFill = [bool]$style.noFill
    $noLine = [bool]$style.noLine

    if ($noFill) { Cell $shape "FillPattern" "0" } else { Cell $shape "FillPattern" "1"; Cell $shape "FillForegnd" $fill }
    if ($noLine) { Cell $shape "LinePattern" "0" } else {
        Cell $shape "LinePattern" $(if ($dash) { "2" } else { "1" })
        Cell $shape "LineColor" $line
        Cell $shape "LineWeight" "$weight pt"
    }
}

function Style-Text($shape, $style) {
    $font = if ($style.fontSize) { [double]$style.fontSize } else { 8.0 }
    $color = if ($style.textColor) { $style.textColor } else { "RGB(40,40,40)" }
    $bold = [bool]$style.bold
    $align = if ($null -ne $style.align) { [int]$style.align } else { 1 }
    Cell $shape "Char.Font" "FONT(""Arial"")"
    Cell $shape "Char.Size" "$font pt"
    Cell $shape "Char.Color" $color
    Cell $shape "Char.Style" $(if ($bold) { "1" } else { "0" })
    Cell $shape "Para.HorzAlign" "$align"
    Cell $shape "VerticalAlign" "1"
    Cell $shape "TxtWidth" "Width*0.94"
}

function Draw-Shape($item) {
    $type = "$($item.type)".ToLowerInvariant()
    $style = if ($item.style) { $item.style } else { [pscustomobject]@{} }
    $text = if ($null -ne $item.text) { [string]$item.text } else { "" }

    if ($type -eq "rect" -or $type -eq "text") {
        $shape = $script:page.DrawRectangle((X $item.x1), (Y $item.y2), (X $item.x2), (Y $item.y1))
        if ($type -eq "text") {
            $style | Add-Member -NotePropertyName noFill -NotePropertyValue $true -Force
            $style | Add-Member -NotePropertyName noLine -NotePropertyValue $true -Force
        }
        Style-Shape $shape $style
        if ($text -ne "") { $shape.Text = $text; Style-Text $shape $style }
        return
    }

    if ($type -eq "oval") {
        $shape = $script:page.DrawOval((X $item.x1), (Y $item.y2), (X $item.x2), (Y $item.y1))
        Style-Shape $shape $style
        if ($text -ne "") { $shape.Text = $text; Style-Text $shape $style }
        return
    }

    if ($type -eq "line") {
        $shape = $script:page.DrawLine((X $item.x1), (Y $item.y1), (X $item.x2), (Y $item.y2))
        $style | Add-Member -NotePropertyName noFill -NotePropertyValue $true -Force
        Style-Shape $shape $style
        Set-Arrow $shape $item.arrow
        return
    }

    if ($type -eq "polyline") {
        $points = @($item.points)
        for ($i = 0; $i -lt ($points.Count - 1); $i++) {
            $p1 = $points[$i]
            $p2 = $points[$i + 1]
            $shape = $script:page.DrawLine((X $p1[0]), (Y $p1[1]), (X $p2[0]), (Y $p2[1]))
            $style | Add-Member -NotePropertyName noFill -NotePropertyValue $true -Force
            Style-Shape $shape $style
            if ($i -eq ($points.Count - 2)) { Set-Arrow $shape $item.arrow }
        }
        return
    }

    throw "不支持的图形类型：$type"
}

function Set-Arrow($shape, $arrow) {
    $value = if ($arrow) { "$arrow" } else { "end" }
    if ($value -eq "end" -or $value -eq "both") { Cell $shape "EndArrow" "13" }
    if ($value -eq "begin" -or $value -eq "both") { Cell $shape "BeginArrow" "13" }
}

function Invoke-ComRetry([scriptblock]$Action, [int]$Attempts = 10, [int]$BaseDelayMs = 500) {
    for ($i = 1; $i -le $Attempts; $i++) {
        try { return & $Action } catch {
            if ($i -eq $Attempts) { throw }
            Start-Sleep -Milliseconds ($BaseDelayMs * $i)
        }
    }
}

function New-VisioApplication {
    for ($i = 1; $i -le 5; $i++) {
        $app = $null
        try {
            $app = New-Object -ComObject Visio.Application
            Start-Sleep -Seconds 4
            return $app
        } catch {
            if ($app -ne $null) { try { $app.Quit() } catch { } }
            if ($i -eq 5) {
                throw "无法通过 COM 自动化启动 Visio。请关闭空白 Visio 窗口，等待几秒后重试。原始错误：$($_.Exception.Message)"
            }
            Start-Sleep -Seconds (2 * $i)
        }
    }
}

$script:visio = $null
$doc = $null
try {
    $script:visio = New-VisioApplication
    if ($Visible) { Invoke-ComRetry { $script:visio.Visible = $true } | Out-Null }
    try { Invoke-ComRetry { $script:visio.AlertResponse = 7 } | Out-Null } catch { }

    $doc = Invoke-ComRetry { $script:visio.Documents.Add("") }

    $script:page = $script:visio.ActivePage
    $script:page.Name = if ($plan.page.name) { [string]$plan.page.name } else { "Visio 复刻图" }
    $script:page.PageSheet.CellsU("PageWidth").FormulaU = "$([double]$plan.page.widthPx / $script:scale) in"
    $script:page.PageSheet.CellsU("PageHeight").FormulaU = "$([double]$plan.page.heightPx / $script:scale) in"

    foreach ($shape in @($plan.shapes)) {
        Draw-Shape $shape
    }

    if ($plan.referenceImage) {
        $refPath = [string]$plan.referenceImage
        if (Test-Path -LiteralPath $refPath) {
            $refPage = $doc.Pages.Add()
            $refPage.Name = "原图参考"
            $refPage.PageSheet.CellsU("PageWidth").FormulaU = "$([double]$plan.page.widthPx / $script:scale) in"
            $refPage.PageSheet.CellsU("PageHeight").FormulaU = "$([double]$plan.page.heightPx / $script:scale) in"
            $pic = $refPage.Import($refPath)
            $pic.CellsU("PinX").FormulaU = "$([double]$plan.page.widthPx / $script:scale / 2) in"
            $pic.CellsU("PinY").FormulaU = "$([double]$plan.page.heightPx / $script:scale / 2) in"
            $pic.CellsU("Width").FormulaU = "$([double]$plan.page.widthPx / $script:scale) in"
            $pic.CellsU("Height").FormulaU = "$([double]$plan.page.heightPx / $script:scale) in"
        }
    }

    $doc.SaveAs($OutVsdx)
    if ($OutEmf) { $script:page.Export($OutEmf) }
    [pscustomobject]@{
        vsdx = $OutVsdx
        emf = $OutEmf
        pages = $doc.Pages.Count
    } | ConvertTo-Json -Depth 3
}
finally {
    if ($doc -ne $null) { try { $doc.Saved = $true } catch { } }
    if ($script:visio -ne $null -and -not $KeepVisioOpen) { try { $script:visio.Quit() } catch { } }
}
