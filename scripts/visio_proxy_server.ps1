# visio_proxy_server.ps1
# 在 ASUS 交互用户会话中运行，保持 Visio COM 常驻
# 通过文件型 IPC 接收 Codex 沙箱用户的命令
#
# 支持的 action:
#   new_document       - 创建新文档
#   draw_shape         - 绘制图形 (type: rect/oval/line/polyline/text)
#   set_page           - 设置页面尺寸
#   save_as            - 保存为 .vsdx
#   export_page        - 导出为 .emf
#   close_document     - 关闭当前文档
#   get_active_page    - 获取当前页面信息
#   ping               - 检查代理是否存活

param(
    [string]$ProxyDir = "D:\codex\projects\draw\scratch\visio_proxy",
    [int]$PollIntervalMs = 500,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$cmdDir = Join-Path $ProxyDir "commands"
$rspDir = Join-Path $ProxyDir "responses"
$lockFile = Join-Path $ProxyDir "server.lock"
$statusFile = Join-Path $ProxyDir "server.status"

# ---- Helper functions ----
function Write-Status($state, $msg) {
    $status = @{ state = $state; message = $msg; pid = $pid; timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
    $status | ConvertTo-Json -Depth 3 | Set-Content -Path $statusFile -Encoding UTF8
}

function Cell($shape, [string]$cell, [string]$formula) {
    try { $shape.CellsU($cell).FormulaU = $formula } catch { }
}

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
    Cell $shape "Char.Size" "$font pt"
    Cell $shape "Char.Color" $color
    Cell $shape "Char.Style" $(if ($bold) { "1" } else { "0" })
    Cell $shape "Para.HorzAlign" "$align"
    Cell $shape "VerticalAlign" "1"
    Cell $shape "TxtWidth" "Width*0.94"
}

function Set-Arrow($shape, $arrow) {
    $value = if ($arrow) { "$arrow" } else { "end" }
    if ($value -eq "end" -or $value -eq "both") { Cell $shape "EndArrow" "13" }
    if ($value -eq "begin" -or $value -eq "both") { Cell $shape "BeginArrow" "13" }
}

# ---- Command execution ----
function Execute-Command($cmd) {
    $action = $cmd.action.ToLowerInvariant()
    $args = if ($cmd.args) { $cmd.args } else { @{} }

    switch ($action) {
        "ping" {
            return @{ success = $true; pid = $pid; visioAlive = ($null -ne $script:visio) }
        }

        "new_document" {
            if ($script:doc -ne $null) { try { $script:doc.Saved = $true; $script:doc.Close() } catch { } }
            $script:doc = $script:visio.Documents.Add("")
            $script:page = $script:visio.ActivePage
            if ($args.name) { $script:page.Name = [string]$args.name }
            return @{ success = $true; pageName = $script:page.Name }
        }

        "set_page" {
            $scale = if ($args.scalePxPerInch) { [double]$args.scalePxPerInch } else { 100.0 }
            $wIn = [double]$args.widthPx / $scale
            $hIn = [double]$args.heightPx / $scale
            $script:page.PageSheet.CellsU("PageWidth").FormulaU = "$wIn in"
            $script:page.PageSheet.CellsU("PageHeight").FormulaU = "$hIn in"
            return @{ success = $true; widthIn = $wIn; heightIn = $hIn }
        }

        "draw_shape" {
            $type = "$($args.type)".ToLowerInvariant()
            $style = if ($args.style) { $args.style } else { @{} }
            $text = if ($null -ne $args.text) { [string]$args.text } else { "" }
            $scale = if ($args.scalePxPerInch) { [double]$args.scalePxPerInch } else { 100.0 }
            $pageH = if ($args.pageHeightPx) { [double]$args.pageHeightPx } else { 800.0 }

            function X([double]$px) { $px / $scale }
            function Y([double]$py) { ($pageH - $py) / $scale }

            $shape = $null
            switch ($type) {
                "rect" {
                    $shape = $script:page.DrawRectangle((X $args.x1), (Y $args.y2), (X $args.x2), (Y $args.y1))
                    Style-Shape $shape $style
                    if ($text -ne "") { $shape.Text = $text; Style-Text $shape $style }
                }
                "oval" {
                    $shape = $script:page.DrawOval((X $args.x1), (Y $args.y2), (X $args.x2), (Y $args.y1))
                    Style-Shape $shape $style
                    if ($text -ne "") { $shape.Text = $text; Style-Text $shape $style }
                }
                "line" {
                    $shape = $script:page.DrawLine((X $args.x1), (Y $args.y1), (X $args.x2), (Y $args.y2))
                    $style.noFill = $true
                    Style-Shape $shape $style
                    Set-Arrow $shape $args.arrow
                }
                "text" {
                    $shape = $script:page.DrawRectangle((X $args.x1), (Y $args.y2), (X $args.x2), (Y $args.y1))
                    $style.noFill = $true
                    $style.noLine = $true
                    Style-Shape $shape $style
                    if ($text -ne "") { $shape.Text = $text; Style-Text $shape $style }
                }
                "polyline" {
                    $points = @($args.points)
                    for ($i = 0; $i -lt ($points.Count - 1); $i++) {
                        $p1 = $points[$i]
                        $p2 = $points[$i + 1]
                        $shape = $script:page.DrawLine((X $p1[0]), (Y $p1[1]), (X $p2[0]), (Y $p2[1]))
                        $style.noFill = $true
                        Style-Shape $shape $style
                        if ($i -eq ($points.Count - 2)) { Set-Arrow $shape $args.arrow }
                    }
                }
                default {
                    throw "不支持的图形类型: $type"
                }
            }
            return @{ success = $true; type = $type }
        }

        "save_as" {
            $path = [string]$args.path
            $dir = Split-Path -Parent $path
            if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
            $script:doc.SaveAs($path)
            return @{ success = $true; path = $path }
        }

        "export_page" {
            $path = [string]$args.path
            $dir = Split-Path -Parent $path
            if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
            $script:page.Export($path)
            return @{ success = $true; path = $path }
        }

        "close_document" {
            if ($script:doc -ne $null) {
                $script:doc.Saved = $true
                $script:doc.Close()
                $script:doc = $null
                $script:page = $null
            }
            return @{ success = $true }
        }

        "get_active_page" {
            if ($script:page -eq $null) { return @{ success = $false; error = "没有活动页面" } }
            $w = $script:page.PageSheet.CellsU("PageWidth").ResultIU
            $h = $script:page.PageSheet.CellsU("PageHeight").ResultIU
            return @{
                success = $true
                name = $script:page.Name
                widthIn = $w
                heightIn = $h
                shapeCount = $script:page.Shapes.Count
            }
        }

        default {
            throw "未知操作: $action"
        }
    }
}

# ---- Main ----
Write-Status "starting" "正在启动 Visio COM 代理..."

$script:visio = $null
$script:doc = $null
$script:page = $null

try {
    for ($i = 1; $i -le 10; $i++) {
        try {
            $script:visio = New-Object -ComObject Visio.Application
            Start-Sleep -Seconds 3
            Write-Host "Visio COM 启动成功"
            break
        } catch {
            if ($i -eq 10) { throw "无法启动 Visio COM: $($_.Exception.Message)" }
            Write-Host "尝试 $i 失败，等待重试..."
            Start-Sleep -Seconds (2 * $i)
        }
    }

    if ($Visible) { $script:visio.Visible = $true }
    try { $script:visio.AlertResponse = 7 } catch { }

    "pid=$pid" | Set-Content -Path $lockFile -Encoding ASCII
    Write-Status "running" "Visio COM 代理运行中 (PID: $pid)"

    Write-Host "Visio COM 代理已就绪，轮询命令目录: $cmdDir"
    Write-Host "按 Ctrl+C 停止代理"

    while ($true) {
        $cmdFiles = Get-ChildItem -Path $cmdDir -Filter "*.json" -ErrorAction SilentlyContinue
        foreach ($cmdFile in $cmdFiles) {
            $cmdId = [System.IO.Path]::GetFileNameWithoutExtension($cmdFile.Name)
            $rspFile = Join-Path $rspDir "$cmdId.json"

            try {
                $cmd = Get-Content -LiteralPath $cmdFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                Write-Host "执行命令: $($cmd.action) (ID: $cmdId)"
                $result = Execute-Command $cmd
                $result | ConvertTo-Json -Depth 5 | Set-Content -Path $rspFile -Encoding UTF8
                Write-Host "命令完成: $cmdId"
            } catch {
                $errorResult = @{ success = $false; error = $_.Exception.Message }
                $errorResult | ConvertTo-Json -Depth 3 | Set-Content -Path $rspFile -Encoding UTF8
                Write-Host "命令失败: $cmdId - $($_.Exception.Message)"
            } finally {
                Remove-Item -LiteralPath $cmdFile.FullName -Force -ErrorAction SilentlyContinue
            }
        }
        Start-Sleep -Milliseconds $PollIntervalMs
    }
} finally {
    Write-Status "stopped" "代理已停止"
    Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue
    if ($script:doc -ne $null) { try { $script:doc.Saved = $true; $script:doc.Close() } catch { } }
    if ($script:visio -ne $null) { try { $script:visio.Quit() } catch { } }
}
